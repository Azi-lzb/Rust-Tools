//! Direct Rust IDispatch calls for the isolated Excel COM benchmark.

use crate::{FileTask, FileTiming, Session};
use anyhow::{anyhow, bail, ensure, Context, Result};
use std::{
    collections::BTreeSet, ffi::c_void, mem::ManuallyDrop, sync::mpsc, thread, time::Instant,
};
use windows::{
    core::{BSTR, GUID, PCWSTR},
    Win32::{
        Foundation::{CloseHandle, HWND, VARIANT_BOOL, WAIT_OBJECT_0},
        System::{
            Com::{
                CLSIDFromProgID, CoCreateInstance, CoInitializeEx, CoUninitialize, IDispatch,
                CLSCTX_LOCAL_SERVER, COINIT_APARTMENTTHREADED, DISPATCH_FLAGS, DISPATCH_METHOD,
                DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPPARAMS, EXCEPINFO,
            },
            Ole::DISPID_PROPERTYPUT,
            ProcessStatus::K32EnumProcesses,
            Threading::{
                GetCurrentThreadId, OpenProcess, WaitForSingleObject, PROCESS_SYNCHRONIZE,
            },
            Variant::{
                VariantClear, VARIANT, VT_BOOL, VT_BSTR, VT_DISPATCH, VT_I2, VT_I4, VT_I8, VT_UI4,
            },
        },
        UI::WindowsAndMessaging::{
            DispatchMessageW, PeekMessageW, TranslateMessage, MSG, PM_REMOVE, WM_QUIT,
        },
    },
};

const XL_CALCULATION_MANUAL: i32 = -4135;
const XL_OPEN_XML_WORKBOOK: i32 = 51;

pub fn run_session(files: Vec<FileTask>) -> Result<Session> {
    let (send, receive) = mpsc::sync_channel(1);
    let thread = thread::Builder::new()
        .name("xls-to-xlsx-benchmark-sta".to_owned())
        .spawn(move || {
            let result = run_on_sta(files);
            let _ = send.send(result);
        })
        .context("could not create native COM STA worker thread")?;
    let result = receive
        .recv()
        .context("native COM STA worker did not return a result")?;
    thread
        .join()
        .map_err(|_| anyhow!("native COM STA worker panicked"))?;
    result
}

fn run_on_sta(files: Vec<FileTask>) -> Result<Session> {
    let sta_thread_id = unsafe { GetCurrentThreadId() };
    let phase = Instant::now();
    let apartment = Apartment::initialize()?;
    let sta_init_ms = elapsed_ms(phase);

    let phase = Instant::now();
    let mut app = match Application::create_and_configure("Excel.Application") {
        Ok(app) => app,
        Err(error) => return Err(error).context("create/configure native Excel COM instance"),
    };
    let excel_create_config_ms = elapsed_ms(phase);
    let excel_pid = app.process_id;
    let worker_pid = std::process::id();

    let mut file_results = Vec::with_capacity(files.len());
    for (index, file) in files.into_iter().enumerate() {
        file_results.push(app.convert_one(file, index));
        pump_messages();
    }

    let quit_start = Instant::now();
    let excel_exit_confirmed = app.shutdown()?;
    let excel_quit_ms = elapsed_ms(quit_start);
    drop(app);

    let sta_release_start = Instant::now();
    drop(apartment);
    let sta_release_ms = elapsed_ms(sta_release_start);

    Ok(Session {
        worker_pid,
        sta_thread_id: serde_json::json!(sta_thread_id),
        excel_pid: Some(excel_pid),
        excel_instance_count: 1,
        excel_exit_confirmed: Some(excel_exit_confirmed),
        timings: [
            ("sta_init_ms".to_owned(), sta_init_ms),
            ("excel_create_config_ms".to_owned(), excel_create_config_ms),
            ("excel_quit_ms".to_owned(), excel_quit_ms),
            ("sta_release_ms".to_owned(), sta_release_ms),
        ]
        .into_iter()
        .collect(),
        files: file_results,
        fatal_error: None,
        worker_start_ms: 0.0,
        worker_exit_ms: 0.0,
        worker_exit_code: Some(0),
    })
}

fn elapsed_ms(start: Instant) -> f64 {
    start.elapsed().as_secs_f64() * 1000.0
}

fn pump_messages() {
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

struct Apartment;
impl Apartment {
    fn initialize() -> Result<Self> {
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok()? };
        Ok(Self)
    }
}
impl Drop for Apartment {
    fn drop(&mut self) {
        unsafe { CoUninitialize() };
    }
}

struct Application {
    object: Option<IDispatch>,
    process_id: u32,
    owned: bool,
    process_handle: Option<windows::Win32::Foundation::HANDLE>,
}

impl Application {
    fn create_and_configure(prog_id: &str) -> Result<Self> {
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
            windows::Win32::UI::WindowsAndMessaging::GetWindowThreadProcessId(
                HWND(hwnd),
                Some(&mut pid),
            );
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
        if let Err(error) = app.configure() {
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

    fn property(&self, name: &str) -> Result<AutoVariant> {
        property(self.object()?, name)
    }

    fn configure(&mut self) -> Result<()> {
        let object = self.object()?.clone();
        put(&object, "Visible", AutoVariant::boolean(false))?;
        put(&object, "DisplayAlerts", AutoVariant::boolean(false))?;
        put(&object, "EnableEvents", AutoVariant::boolean(false))?;
        put(&object, "ScreenUpdating", AutoVariant::boolean(false))?;

        let workbooks = property(&object, "Workbooks")?.dispatch()?;
        let blank = method(&workbooks, "Add", vec![])?.dispatch()?;
        let calculation_result = (|| -> Result<()> {
            put(
                &object,
                "Calculation",
                AutoVariant::int(XL_CALCULATION_MANUAL),
            )?;
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

    fn convert_one(&self, file: FileTask, index: usize) -> FileTiming {
        let started = Instant::now();
        let mut result = FileTiming {
            source: file.input.to_string_lossy().into_owned(),
            target: file.output.to_string_lossy().into_owned(),
            index,
            status: "failed".to_owned(),
            open_ms: None,
            save_as_ms: None,
            close_ms: None,
            total_ms: 0.0,
            error: None,
            close_error: None,
        };
        let open_start = Instant::now();
        let workbook_object = match (|| -> Result<IDispatch> {
            let books = self.property("Workbooks")?.dispatch()?;
            method(
                &books,
                "Open",
                vec![
                    AutoVariant::text(&file.input.to_string_lossy()),
                    AutoVariant::int(0),
                    AutoVariant::boolean(true),
                ],
            )
            .with_context(|| format!("Excel Workbooks.Open failed for {}", file.input.display()))?
            .dispatch()
        })() {
            Ok(workbook) => workbook,
            Err(error) => {
                result.error = Some(format!("{error:#}"));
                result.total_ms = elapsed_ms(started);
                return result;
            }
        };
        result.open_ms = Some(elapsed_ms(open_start));
        let mut workbook = Workbook {
            object: Some(workbook_object),
            open: true,
        };

        let save_start = Instant::now();
        let save_result = method(
            workbook.object.as_ref().expect("open workbook"),
            "SaveAs",
            vec![
                AutoVariant::text(&file.output.to_string_lossy()),
                AutoVariant::int(XL_OPEN_XML_WORKBOOK),
            ],
        )
        .map(|_| ())
        .context("Excel Workbook.SaveAs failed");
        result.save_as_ms = Some(elapsed_ms(save_start));

        let close_start = Instant::now();
        match workbook.close() {
            Ok(()) => result.close_ms = Some(elapsed_ms(close_start)),
            Err(error) => result.close_error = Some(format!("{error:#}")),
        }
        if let Err(error) = save_result {
            result.error = Some(format!("{error:#}"));
        } else if result.close_error.is_none() {
            result.status = "succeeded".to_owned();
        }
        result.total_ms = elapsed_ms(started);
        result
    }

    fn shutdown(&mut self) -> Result<bool> {
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
            ensure!(wait == WAIT_OBJECT_0, "owned Excel PID {} did not exit after Quit (wait={wait:?}); it was not force-terminated", self.process_id);
            Ok(true)
        } else {
            Ok(false)
        };
        quit_result?;
        exit_result
    }
}

impl Drop for Application {
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

struct Workbook {
    object: Option<IDispatch>,
    open: bool,
}
impl Workbook {
    fn close(&mut self) -> Result<()> {
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
impl Drop for Workbook {
    fn drop(&mut self) {
        if self.open {
            if let Some(object) = self.object.take() {
                let _ = method(&object, "Close", vec![AutoVariant::boolean(false)]);
            }
        }
    }
}

fn process_ids() -> Result<BTreeSet<u32>> {
    let mut capacity = 4096usize;
    loop {
        let mut ids = vec![0u32; capacity];
        let mut bytes = 0u32;
        unsafe { K32EnumProcesses(ids.as_mut_ptr(), (ids.len() * 4) as u32, &mut bytes).ok()? };
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

struct AutoVariant(VARIANT);
impl AutoVariant {
    fn empty() -> Self {
        Self(VARIANT::default())
    }
    fn text(value: &str) -> Self {
        let mut result = Self::empty();
        unsafe {
            let raw = &mut *result.0.Anonymous.Anonymous;
            raw.vt = VT_BSTR;
            raw.Anonymous.bstrVal = ManuallyDrop::new(BSTR::from(value));
        }
        result
    }
    fn int(value: i32) -> Self {
        let mut result = Self::empty();
        unsafe {
            let raw = &mut *result.0.Anonymous.Anonymous;
            raw.vt = VT_I4;
            raw.Anonymous.lVal = value;
        }
        result
    }
    fn boolean(value: bool) -> Self {
        let mut result = Self::empty();
        unsafe {
            let raw = &mut *result.0.Anonymous.Anonymous;
            raw.vt = VT_BOOL;
            raw.Anonymous.boolVal = VARIANT_BOOL(if value { -1 } else { 0 });
        }
        result
    }
    fn integer(&self) -> Result<i64> {
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
    fn dispatch(&self) -> Result<IDispatch> {
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

fn invoke(
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
fn property(object: &IDispatch, name: &str) -> Result<AutoVariant> {
    invoke(object, name, DISPATCH_PROPERTYGET, vec![])
}
fn method(object: &IDispatch, name: &str, args: Vec<AutoVariant>) -> Result<AutoVariant> {
    invoke(object, name, DISPATCH_METHOD, args)
}
fn put(object: &IDispatch, name: &str, value: AutoVariant) -> Result<()> {
    invoke(object, name, DISPATCH_PROPERTYPUT, vec![value]).map(|_| ())
}
