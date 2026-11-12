//! Windows Excel COM plumbing shared by the conversion backend and the
//! open-check tool. Adapted from the repository's existing isolated
//! `benchmarks/xls-to-xlsx-com` runner (same Excel configuration so the
//! baseline stays comparable).

use anyhow::{anyhow, bail, ensure, Context, Result};
use std::{
    collections::BTreeSet,
    ffi::c_void,
    mem::ManuallyDrop,
    sync::mpsc,
    thread,
    time::Instant,
};
use windows::{
    core::{BSTR, GUID, PCWSTR},
    Win32::{
        Foundation::{CloseHandle, HANDLE, HWND, VARIANT_BOOL, WAIT_OBJECT_0},
        System::{
            Com::{
                CLSIDFromProgID, CoCreateInstance, CoInitializeEx, CoUninitialize, IDispatch,
                CLSCTX_LOCAL_SERVER, COINIT_APARTMENTTHREADED, DISPATCH_FLAGS, DISPATCH_METHOD,
                DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPPARAMS, EXCEPINFO,
            },
            Ole::DISPID_PROPERTYPUT,
            ProcessStatus::{K32EnumProcesses, K32GetProcessMemoryInfo, PROCESS_MEMORY_COUNTERS},
            Threading::{
                GetProcessTimes, OpenProcess, WaitForSingleObject,
                PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_SYNCHRONIZE,
            },
            Variant::{
                VariantClear, VARIANT, VT_BOOL, VT_BSTR, VT_DISPATCH, VT_ERROR, VT_I2, VT_I4,
                VT_I8, VT_R8, VT_UI4,
            },
        },
        UI::WindowsAndMessaging::{
            DispatchMessageW, GetWindowThreadProcessId, PeekMessageW, TranslateMessage, MSG,
            PM_REMOVE, WM_QUIT,
        },
    },
};

pub const XL_CALCULATION_MANUAL: i32 = -4135;
pub const XL_OPEN_XML_WORKBOOK: i32 = 51;
pub const XL_PDF_TYPE: i32 = 0;

/// Run a closure on a dedicated STA thread and wait for its result.
pub fn sta_call<T, F>(work: F) -> Result<T>
where
    T: Send + 'static,
    F: FnOnce() -> Result<T> + Send + 'static,
{
    let (send, receive) = mpsc::sync_channel(1);
    let thread = thread::Builder::new()
        .name("xls2x-com-sta".to_owned())
        .spawn(move || {
            let result = work();
            let _ = send.send(result);
        })
        .context("could not create COM STA worker thread")?;
    let result = receive.recv().context("COM STA worker did not return")?;
    thread
        .join()
        .map_err(|_| anyhow!("COM STA worker panicked"))?;
    result
}

pub fn pump_messages() {
    let mut message = MSG::default();
    unsafe {
        while PeekMessageW(&mut message, None, 0, 0, PM_REMOVE).as_bool() {
            if message.message == WM_QUIT {
                break;
            }
            let _ = TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }
}

pub struct Apartment;
impl Apartment {
    pub fn initialize() -> Result<Self> {
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok()? };
        Ok(Self)
    }
}
impl Drop for Apartment {
    fn drop(&mut self) {
        unsafe { CoUninitialize() };
    }
}

/// One owned Excel application instance with strict PID ownership checks.
pub struct ComApp {
    object: Option<IDispatch>,
    pub process_id: u32,
    owned: bool,
    process_handle: Option<HANDLE>,
}

impl ComApp {
    /// Create `Excel.Application`, prove it is a fresh instance we own, then
    /// apply the same quiet-configuration the production baseline uses.
    /// `configure_calc` mirrors the baseline's manual-calculation setup, which
    /// requires a transient blank workbook.
    pub fn create_and_configure(prog_id: &str, configure_calc: bool) -> Result<Self> {
        let before = process_ids()?;
        let wide = prog_id.encode_utf16().chain([0]).collect::<Vec<_>>();
        let class = unsafe { CLSIDFromProgID(PCWSTR(wide.as_ptr())) }
            .with_context(|| format!("resolve {prog_id}"))?;
        let object: IDispatch = unsafe { CoCreateInstance(&class, None, CLSCTX_LOCAL_SERVER) }
            .with_context(|| format!("create {prog_id}"))?;
        let mut app = Self {
            object: Some(object),
            process_id: 0,
            owned: false,
            process_handle: None,
        };

        let hwnd = app.property("Hwnd")?.integer()? as u32 as usize as *mut c_void;
        let mut pid = 0u32;
        unsafe {
            GetWindowThreadProcessId(HWND(hwnd), Some(&mut pid));
        }
        let books = app.property("Workbooks")?.dispatch()?;
        let count = property(&books, "Count")?.integer()?;
        ensure!(
            pid != 0 && !before.contains(&pid) && count == 0,
            "Excel.Application did not create a new empty owned instance (PID {pid}, workbook count {count}); refusing to touch it"
        );
        app.process_handle = Some(
            unsafe { OpenProcess(PROCESS_SYNCHRONIZE, false, pid) }
                .with_context(|| format!("open synchronization handle for Excel PID {pid}"))?,
        );
        app.process_id = pid;
        app.owned = true;
        if let Err(error) = app.configure(configure_calc) {
            let _ = app.shutdown();
            return Err(error).context("configure owned Excel instance");
        }
        Ok(app)
    }

    fn object(&self) -> Result<&IDispatch> {
        self.object
            .as_ref()
            .context("Excel COM object was released")
    }

    /// Cloned reference to the application dispatch (for fixture generation).
    pub fn object_dispatch(&self) -> Result<IDispatch> {
        Ok(self.object()?.clone())
    }

    pub fn property(&self, name: &str) -> Result<AutoVariant> {
        property(self.object()?, name)
    }

    fn configure(&mut self, configure_calc: bool) -> Result<()> {
        let object = self.object()?.clone();
        put(&object, "Visible", AutoVariant::boolean(false))?;
        put(&object, "DisplayAlerts", AutoVariant::boolean(false))?;
        put(&object, "EnableEvents", AutoVariant::boolean(false))?;
        put(&object, "ScreenUpdating", AutoVariant::boolean(false))?;
        if !configure_calc {
            return Ok(());
        }
        let workbooks = property(&object, "Workbooks")?.dispatch()?;
        let blank = method(&workbooks, "Add", vec![])?.dispatch()?;
        let calculation_result = (|| -> Result<()> {
            put(&object, "Calculation", AutoVariant::int(XL_CALCULATION_MANUAL))?;
            put(&object, "CalculateBeforeSave", AutoVariant::boolean(false))?;
            let calculation = property(&object, "Calculation")?.integer()? as i32;
            ensure!(
                calculation == XL_CALCULATION_MANUAL,
                "Excel did not enter manual calculation mode"
            );
            Ok(())
        })();
        let close_result = method(&blank, "Close", vec![AutoVariant::boolean(false)])
            .map(|_| ())
            .context("close calculation-configuration workbook");
        calculation_result?;
        close_result?;
        Ok(())
    }

    /// Open a workbook read-only; returns the workbook dispatch.
    pub fn open_workbook(&self, path: &str) -> Result<IDispatch> {
        let books = self.property("Workbooks")?.dispatch()?;
        method(
            &books,
            "Open",
            vec![
                AutoVariant::text(path),
                AutoVariant::int(0),
                AutoVariant::boolean(true),
            ],
        )
        .with_context(|| format!("Excel Workbooks.Open failed for {path}"))?
        .dispatch()
    }

    /// Quit Excel and wait (bounded) for the owned PID to exit.
    pub fn shutdown(&mut self) -> Result<bool> {
        let quit_result = if let Some(object) = self.object.take() {
            let result = if self.owned {
                method(&object, "Quit", vec![])
                    .map(|_| ())
                    .context("Excel.Application.Quit failed")
            } else {
                Ok(())
            };
            drop(object);
            result
        } else {
            Ok(())
        };
        let exit_result = if let Some(handle) = self.process_handle.take() {
            let wait = unsafe { WaitForSingleObject(handle, 30000) };
            let _ = unsafe { CloseHandle(handle) };
            ensure!(
                wait == WAIT_OBJECT_0,
                "owned Excel PID {} did not exit after Quit (wait={wait:?}); it was not force-terminated",
                self.process_id
            );
            Ok(true)
        } else {
            Ok(false)
        };
        quit_result?;
        exit_result
    }

    /// Peak working set (bytes) and kernel+user CPU time (ms) of the owned
    /// Excel process. Call before shutdown for a meaningful peak.
    pub fn memory(&self) -> Result<(u64, f64)> {
        memory_and_cpu_of_pid(self.process_id)
    }
}

impl Drop for ComApp {
    fn drop(&mut self) {
        if let Some(object) = self.object.take() {
            if self.owned {
                let _ = method(&object, "Quit", vec![]);
            }
        }
        if let Some(handle) = self.process_handle.take() {
            let _ = unsafe { CloseHandle(handle) };
        }
    }
}

/// An open workbook handle that closes itself on drop.
pub struct ComWorkbook {
    object: Option<IDispatch>,
    open: bool,
}

impl ComWorkbook {
    pub fn from_dispatch(object: IDispatch) -> Self {
        Self {
            object: Some(object),
            open: true,
        }
    }

    pub fn save_as_xlsx(&self, path: &str) -> Result<()> {
        method(
            self.object()?,
            "SaveAs",
            vec![AutoVariant::text(path), AutoVariant::int(XL_OPEN_XML_WORKBOOK)],
        )
        .map(|_| ())
        .context("Excel Workbook.SaveAs failed")
    }

    pub fn export_pdf(&self, path: &str) -> Result<()> {
        method(
            self.object()?,
            "ExportAsFixedFormat",
            vec![
                AutoVariant::int(XL_PDF_TYPE),
                AutoVariant::text(path),
                AutoVariant::int(0),
                AutoVariant::boolean(false),
                AutoVariant::boolean(false),
            ],
        )
        .map(|_| ())
        .context("Excel ExportAsFixedFormat failed")
    }

    pub fn sheet_count(&self) -> Result<usize> {
        let sheets = property(self.object()?, "Sheets")?.dispatch()?;
        Ok(property(&sheets, "Count")?.integer()? as usize)
    }

    /// `Worksheet.PageSetup.Pages.Count` summed over sheets (print pages).
    pub fn total_print_pages(&self) -> Result<u64> {
        let sheets = property(self.object()?, "Sheets")?.dispatch()?;
        let count = property(&sheets, "Count")?.integer()?;
        let mut total = 0u64;
        for index in 1..=count {
            let sheet = method(&sheets, "Item", vec![AutoVariant::int(index as i32)])?.dispatch()?;
            let pages = (|| -> Result<i64> {
                let setup = property(&sheet, "PageSetup")?.dispatch()?;
                let collection = property(&setup, "Pages")?.dispatch()?;
                property(&collection, "Count")?.integer()
            })()
            .unwrap_or(0);
            total += pages.max(0) as u64;
        }
        Ok(total)
    }

    pub fn close(&mut self) -> Result<()> {
        if let Some(object) = self.object.take() {
            let result = method(&object, "Close", vec![AutoVariant::boolean(false)])
                .map(|_| ())
                .context("Excel Workbook.Close failed");
            drop(object);
            self.open = false;
            result
        } else {
            Ok(())
        }
    }
}

impl ComWorkbook {
    fn object(&self) -> Result<&IDispatch> {
        self.object.as_ref().context("workbook COM object released")
    }
}

impl Drop for ComWorkbook {
    fn drop(&mut self) {
        if self.open {
            if let Some(object) = self.object.take() {
                let _ = method(&object, "Close", vec![AutoVariant::boolean(false)]);
            }
        }
    }
}

pub fn memory_and_cpu_of_pid(pid: u32) -> Result<(u64, f64)> {
    let handle = unsafe { OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid) }
        .with_context(|| format!("open query handle for PID {pid}"))?;
    let mut peak = 0u64;
    unsafe {
        let mut counters = PROCESS_MEMORY_COUNTERS::default();
        counters.cb = std::mem::size_of::<PROCESS_MEMORY_COUNTERS>() as u32;
        if K32GetProcessMemoryInfo(handle, &mut counters, counters.cb).as_bool() {
            peak = counters.PeakWorkingSetSize as u64;
        }
    }
    let mut cpu_ms = 0.0f64;
    unsafe {
        let mut creation = Default::default();
        let mut exit = Default::default();
        let mut kernel = Default::default();
        let mut user = Default::default();
        if GetProcessTimes(
            handle,
            &mut creation,
            &mut exit,
            &mut kernel,
            &mut user,
        )
        .is_ok()
        {
            cpu_ms = (filetime_u64(&user) + filetime_u64(&kernel)) as f64 / 10_000.0;
        }
    }
    let _ = unsafe { CloseHandle(handle) };
    Ok((peak, cpu_ms))
}

fn filetime_u64(time: &windows::Win32::Foundation::FILETIME) -> u64 {
    ((time.dwHighDateTime as u64) << 32) | time.dwLowDateTime as u64
}

fn process_ids() -> Result<BTreeSet<u32>> {
    let mut capacity = 4096usize;
    loop {
        let mut ids = vec![0u32; capacity];
        let mut bytes = 0u32;
        let ok = unsafe {
            K32EnumProcesses(ids.as_mut_ptr(), (ids.len() * 4) as u32, &mut bytes).as_bool()
        };
        ensure!(ok, "EnumProcesses failed");
        if (bytes as usize) < ids.len() * 4 {
            ids.truncate(bytes as usize / 4);
            return Ok(ids.into_iter().collect());
        }
        capacity *= 2;
        ensure!(
            capacity <= 1_048_576,
            "process list is too large to confirm Excel ownership"
        );
    }
}

pub struct AutoVariant(VARIANT);
impl AutoVariant {
    pub fn empty() -> Self {
        Self(VARIANT::default())
    }
    /// The canonical "optional parameter not supplied" VARIANT:
    /// VT_ERROR with DISP_E_PARAMNOTFOUND. Excel rejects VT_EMPTY for omitted
    /// optional arguments ("发生意外").
    pub fn missing() -> Self {
        const DISP_E_PARAMNOTFOUND: i32 = 0x8002_0004u32 as i32;
        let mut result = Self::empty();
        unsafe {
            let raw = &mut *result.0.Anonymous.Anonymous;
            raw.vt = VT_ERROR;
            raw.Anonymous.scode = DISP_E_PARAMNOTFOUND;
        }
        result
    }
    pub fn text(value: &str) -> Self {
        let mut result = Self::empty();
        unsafe {
            let raw = &mut *result.0.Anonymous.Anonymous;
            raw.vt = VT_BSTR;
            raw.Anonymous.bstrVal = ManuallyDrop::new(BSTR::from(value));
        }
        result
    }
    pub fn int(value: i32) -> Self {
        let mut result = Self::empty();
        unsafe {
            let raw = &mut *result.0.Anonymous.Anonymous;
            raw.vt = VT_I4;
            raw.Anonymous.lVal = value;
        }
        result
    }
    pub fn boolean(value: bool) -> Self {
        let mut result = Self::empty();
        unsafe {
            let raw = &mut *result.0.Anonymous.Anonymous;
            raw.vt = VT_BOOL;
            raw.Anonymous.boolVal = VARIANT_BOOL(if value { -1 } else { 0 });
        }
        result
    }
    pub fn float(value: f64) -> Self {
        let mut result = Self::empty();
        unsafe {
            let raw = &mut *result.0.Anonymous.Anonymous;
            raw.vt = VT_R8;
            raw.Anonymous.dblVal = value;
        }
        result
    }
    pub fn from_dispatch(value: IDispatch) -> Result<Self> {
        let mut result = Self::empty();
        unsafe {
            let raw = &mut *result.0.Anonymous.Anonymous;
            raw.vt = VT_DISPATCH;
            raw.Anonymous.pdispVal = ManuallyDrop::new(Some(value));
        }
        Ok(result)
    }
    pub fn integer(&self) -> Result<i64> {
        unsafe {
            let raw = &self.0.Anonymous.Anonymous;
            match raw.vt {
                VT_I2 => Ok(raw.Anonymous.iVal as i64),
                VT_I4 => Ok(raw.Anonymous.lVal as i64),
                VT_I8 => Ok(raw.Anonymous.llVal),
                VT_UI4 => Ok(raw.Anonymous.ulVal as i64),
                other => bail!("COM returned non-integer variant type {other:?}"),
            }
        }
    }
    /// Numeric extraction accepting both integer and floating VARIANT types.
    pub fn as_f64(&self) -> Result<f64> {
        unsafe {
            let raw = &self.0.Anonymous.Anonymous;
            match raw.vt {
                VT_I2 => Ok(raw.Anonymous.iVal as f64),
                VT_I4 => Ok(raw.Anonymous.lVal as f64),
                VT_I8 => Ok(raw.Anonymous.llVal as f64),
                VT_UI4 => Ok(raw.Anonymous.ulVal as f64),
                VT_R8 => Ok(raw.Anonymous.dblVal),
                other => bail!("COM returned non-numeric variant type {other:?}"),
            }
        }
    }
    pub fn dispatch(&self) -> Result<IDispatch> {
        unsafe {
            let raw = &self.0.Anonymous.Anonymous;
            ensure!(
                raw.vt == VT_DISPATCH,
                "COM returned non-dispatch variant type {:?}",
                raw.vt
            );
            (*raw.Anonymous.pdispVal)
                .as_ref()
                .cloned()
                .context("COM returned a null dispatch")
        }
    }
}
impl Drop for AutoVariant {
    fn drop(&mut self) {
        unsafe {
            let _ = VariantClear(&mut self.0);
        }
    }
}

pub fn invoke(
    object: &IDispatch,
    name: &str,
    flags: DISPATCH_FLAGS,
    mut args: Vec<AutoVariant>,
) -> Result<AutoVariant> {
    let wide = name.encode_utf16().chain([0]).collect::<Vec<_>>();
    let name_ptr = PCWSTR(wide.as_ptr());
    let mut id = 0;
    let iid = GUID::zeroed();
    unsafe { object.GetIDsOfNames(&iid, &name_ptr, 1, 0x400, &mut id) }
        .with_context(|| format!("lookup COM member {name}"))?;
    args.reverse();
    let mut property_put = DISPID_PROPERTYPUT;
    let named = flags == DISPATCH_PROPERTYPUT;
    let params = DISPPARAMS {
        rgvarg: args.as_mut_ptr().cast(),
        rgdispidNamedArgs: if named {
            &mut property_put
        } else {
            std::ptr::null_mut()
        },
        cArgs: args.len() as u32,
        cNamedArgs: u32::from(named),
    };
    let mut result = AutoVariant::empty();
    let mut exception = EXCEPINFO::default();
    let mut arg_error = 0u32;
    let status = unsafe {
        object.Invoke(
            id,
            &iid,
            0x400,
            flags,
            &params,
            Some(&mut result.0),
            Some(&mut exception),
            Some(&mut arg_error),
        )
    };
    if status.is_err() {
        if let Some(fill) = exception.pfnDeferredFillIn {
            let _ = unsafe { fill(&mut exception) };
        }
    }
    let detail = exception.bstrDescription.to_string();
    unsafe {
        ManuallyDrop::drop(&mut exception.bstrSource);
        ManuallyDrop::drop(&mut exception.bstrDescription);
        ManuallyDrop::drop(&mut exception.bstrHelpFile);
    }
    status.with_context(|| {
        if detail.trim().is_empty() {
            format!("COM invocation {name} failed")
        } else {
            format!("COM invocation {name} failed: {}", detail.trim())
        }
    })?;
    Ok(result)
}

pub fn property(object: &IDispatch, name: &str) -> Result<AutoVariant> {
    invoke(object, name, DISPATCH_PROPERTYGET, vec![])
}

/// Parameterized property-get / method with a single argument (e.g. `Range`,
/// `Item`, `Cells`).
pub fn invoke_arg(object: &IDispatch, name: &str, arg: AutoVariant) -> Result<AutoVariant> {
    invoke(object, name, DISPATCH_PROPERTYGET, vec![arg])
        .with_context(|| format!("parameterized get {name}"))
}

pub fn method(object: &IDispatch, name: &str, args: Vec<AutoVariant>) -> Result<AutoVariant> {
    // Excel's collection `Item` and other default members are only reachable
    // with DISPATCH_METHOD|DISPATCH_PROPERTYGET combined; DISPATCH_METHOD alone
    // returns DISP_E_MEMBERNOTFOUND for them.
    let flags = DISPATCH_METHOD | DISPATCH_PROPERTYGET;
    invoke(object, name, flags, args)
}

pub fn put(object: &IDispatch, name: &str, value: AutoVariant) -> Result<()> {
    invoke(object, name, DISPATCH_PROPERTYPUT, vec![value]).map(|_| ())
}

#[allow(dead_code)]
fn elapsed_marker(start: Instant) -> f64 {
    start.elapsed().as_secs_f64() * 1000.0
}
