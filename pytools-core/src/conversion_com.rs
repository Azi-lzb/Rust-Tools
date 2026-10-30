//! Windows spreadsheet format conversion through native Rust COM automation.
//!
//! COM objects remain on one dedicated STA thread. A batch reuses only the
//! Excel/WPS instances it created, while every input handed to this module is
//! already a disposable copy. Word conversion remains on its existing path.

use anyhow::{anyhow, bail, ensure, Context, Result};
use std::{
    collections::BTreeSet,
    ffi::c_void,
    mem::ManuallyDrop,
    path::{Path, PathBuf},
    sync::mpsc::{self, Receiver, RecvTimeoutError, Sender},
    thread::{self, JoinHandle},
    time::Duration,
};
use windows::{
    core::{BSTR, GUID, PCWSTR},
    Win32::{
        Foundation::{HWND, VARIANT_BOOL},
        System::{
            Com::{
                CLSIDFromProgID, CoCreateInstance, CoInitializeEx, CoUninitialize, IDispatch,
                CLSCTX_LOCAL_SERVER, COINIT_APARTMENTTHREADED, DISPATCH_FLAGS, DISPATCH_METHOD,
                DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPPARAMS, EXCEPINFO,
            },
            Ole::DISPID_PROPERTYPUT,
            ProcessStatus::K32EnumProcesses,
            Variant::{
                VariantClear, VARIANT, VT_BOOL, VT_BSTR, VT_DISPATCH, VT_I2, VT_I4, VT_I8, VT_UI4,
            },
        },
        UI::WindowsAndMessaging::{
            DispatchMessageW, GetWindowThreadProcessId, PeekMessageW, TranslateMessage, MSG,
            PM_REMOVE, WM_QUIT,
        },
    },
};

enum Request {
    Convert {
        input: PathBuf,
        output: PathBuf,
        source_ext: String,
        reply: Sender<std::result::Result<SpreadsheetConversion, String>>,
    },
    Stop,
}

pub(crate) struct SpreadsheetConversion {
    pub source_format: Option<i32>,
    pub application: &'static str,
    pub process_id: u32,
}

/// One session belongs to one conversion request or interactive batch. It
/// never transfers COM interfaces between threads.
pub(crate) struct ComWorker {
    requests: Sender<Request>,
    thread: Option<JoinHandle<()>>,
}

impl ComWorker {
    pub(crate) fn start() -> Result<Self> {
        let (requests, receiver) = mpsc::channel();
        let (ready, initialized) = mpsc::sync_channel(1);
        let thread = thread::Builder::new()
            .name("pytools-spreadsheet-com".to_owned())
            .spawn(move || {
                let mut context = match ComContext::new() {
                    Ok(context) => {
                        let _ = ready.send(Ok(()));
                        context
                    }
                    Err(error) => {
                        let _ = ready.send(Err(format!("{error:#}")));
                        return;
                    }
                };
                worker_loop(&mut context, receiver);
            })
            .context("启动 Excel/WPS COM 工作线程失败")?;
        match initialized
            .recv()
            .context("Excel/WPS COM 工作线程未完成初始化")?
        {
            Ok(()) => Ok(Self {
                requests,
                thread: Some(thread),
            }),
            Err(message) => {
                let _ = thread.join();
                bail!("初始化 Excel/WPS COM 失败：{message}");
            }
        }
    }

    pub(crate) fn convert(
        &mut self,
        input: &Path,
        output: &Path,
        source_ext: &str,
    ) -> Result<SpreadsheetConversion> {
        let (reply, response) = mpsc::channel();
        self.requests
            .send(Request::Convert {
                input: input.to_owned(),
                output: output.to_owned(),
                source_ext: source_ext.to_owned(),
                reply,
            })
            .context("Excel/WPS COM 工作线程不可用")?;
        response
            .recv()
            .context("Excel/WPS COM 工作线程未返回结果")?
            .map_err(|message| anyhow!("{message}"))
    }
}

impl Drop for ComWorker {
    fn drop(&mut self) {
        let _ = self.requests.send(Request::Stop);
        if let Some(thread) = self.thread.take() {
            if thread.join().is_err() {
                eprintln!(
                    "{}",
                    crate::run_logging::terminal_line(
                        crate::run_logging::Level::Warn,
                        "Excel/WPS COM 工作线程异常退出"
                    )
                );
            }
        }
    }
}

fn worker_loop(context: &mut ComContext, receiver: Receiver<Request>) {
    loop {
        if !pump_messages() {
            break;
        }
        match receiver.recv_timeout(Duration::from_millis(20)) {
            Ok(Request::Convert {
                input,
                output,
                source_ext,
                reply,
            }) => {
                let result = context
                    .convert(&input, &output, &source_ext)
                    .map_err(|error| format!("{error:#}"));
                let _ = reply.send(result);
            }
            Ok(Request::Stop) | Err(RecvTimeoutError::Disconnected) => break,
            Err(RecvTimeoutError::Timeout) => {}
        }
    }
}

fn pump_messages() -> bool {
    let mut message = MSG::default();
    unsafe {
        while PeekMessageW(&mut message, None, 0, 0, PM_REMOVE).as_bool() {
            if message.message == WM_QUIT {
                return false;
            }
            let _ = TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }
    true
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

#[derive(Clone, Copy, Eq, PartialEq)]
enum AppKind {
    Excel,
    Wps,
}

struct ComContext {
    // Applications are dropped before the apartment is uninitialized.
    excel: Option<Application>,
    wps: Option<Application>,
    ordinary_choice: Option<AppKind>,
    _apartment: Apartment,
}

impl ComContext {
    fn new() -> Result<Self> {
        Ok(Self {
            excel: None,
            wps: None,
            ordinary_choice: None,
            _apartment: Apartment::initialize()?,
        })
    }

    fn ensure_wps(&mut self) -> Result<()> {
        if self.wps.is_none() {
            self.wps =
                Some(Application::create_first(&["KET.Application"])?.context(
                    "No supported Office/WPS COM engine is available: WPS KET.Application",
                )?);
        }
        Ok(())
    }

    fn choose_ordinary(&mut self) -> Result<AppKind> {
        if let Some(kind) = self.ordinary_choice {
            return Ok(kind);
        }
        if self.excel.is_none() {
            self.excel = Application::create_first(&["Excel.Application"])?;
        }
        let kind = if self.excel.is_some() {
            AppKind::Excel
        } else {
            self.ensure_wps()?;
            AppKind::Wps
        };
        self.ordinary_choice = Some(kind);
        Ok(kind)
    }

    fn convert(
        &mut self,
        input: &Path,
        output: &Path,
        source_ext: &str,
    ) -> Result<SpreadsheetConversion> {
        let kind = if source_ext == "et" {
            self.ensure_wps()?;
            AppKind::Wps
        } else {
            self.choose_ordinary()?
        };
        let app = match kind {
            AppKind::Excel => self.excel.as_mut().expect("Excel selected"),
            AppKind::Wps => self.wps.as_mut().expect("WPS selected"),
        };
        let result =
            app.convert(input, output, source_ext)
                .map(|source_format| SpreadsheetConversion {
                    source_format,
                    application: app.prog_id,
                    process_id: app.process_id,
                });
        if result.is_err() {
            // A failed Open/SaveAs can leave an application in an unknown state.
            // Close only our own instance and let the next file start cleanly.
            match kind {
                AppKind::Excel => self.excel = None,
                AppKind::Wps => self.wps = None,
            }
            if self.ordinary_choice == Some(kind) {
                self.ordinary_choice = None;
            }
        }
        result
    }
}

impl Drop for ComContext {
    fn drop(&mut self) {
        drop(self.excel.take());
        drop(self.wps.take());
    }
}

struct Application {
    object: IDispatch,
    prog_id: &'static str,
    process_id: u32,
    owned: bool,
}

impl Application {
    /// An unavailable registration is eligible for the next configured
    /// application. Once an object exists, ownership or setup errors are
    /// returned explicitly rather than silently switching engines.
    fn create_first(names: &[&'static str]) -> Result<Option<Self>> {
        for &prog_id in names {
            let before = process_ids()?;
            let wide = prog_id.encode_utf16().chain([0]).collect::<Vec<_>>();
            let class = match unsafe { CLSIDFromProgID(PCWSTR(wide.as_ptr())) } {
                Ok(class) => class,
                Err(_) => continue,
            };
            let object: IDispatch =
                match unsafe { CoCreateInstance(&class, None, CLSCTX_LOCAL_SERVER) } {
                    Ok(object) => object,
                    Err(_) => continue,
                };
            let mut app = Self {
                object,
                prog_id,
                process_id: 0,
                owned: false,
            };
            let hwnd = property(&app.object, "Hwnd")?.integer()? as u32 as usize as *mut c_void;
            let mut pid = 0u32;
            unsafe {
                GetWindowThreadProcessId(HWND(hwnd), Some(&mut pid));
            }
            let books = property(&app.object, "Workbooks")?.dispatch()?;
            let count = property(&books, "Count")?.integer()?;
            ensure!(
                pid != 0 && !before.contains(&pid) && count == 0,
                "{prog_id} 未创建可确认的独立空实例；为保护已打开的工作簿，停止转换"
            );
            app.process_id = pid;
            app.owned = true;
            put(&app.object, "Visible", AutoVariant::boolean(false))?;
            put(&app.object, "DisplayAlerts", AutoVariant::boolean(false))?;
            return Ok(Some(app));
        }
        Ok(None)
    }

    fn convert(&mut self, input: &Path, output: &Path, source_ext: &str) -> Result<Option<i32>> {
        let workbooks = property(&self.object, "Workbooks")?.dispatch()?;
        let object = method(
            &workbooks,
            "Open",
            vec![
                AutoVariant::text(&input.to_string_lossy()),
                AutoVariant::int(0),
                AutoVariant::boolean(true),
            ],
        )
        .with_context(|| format!("{} 打开源工作簿失败", self.prog_id))?
        .dispatch()?;
        let mut book = Workbook { object, open: true };
        let source_format = if source_ext == "et" {
            let format = property(&book.object, "FileFormat")?.integer()? as i32;
            ensure!(
                format == 65521,
                "Input has .et extension but KET opened FileFormat {format}; expected native ET FileFormat 65521"
            );
            Some(format)
        } else {
            None
        };
        method(
            &book.object,
            "SaveAs",
            vec![
                AutoVariant::text(&output.to_string_lossy()),
                AutoVariant::int(51),
            ],
        )
        .with_context(|| format!("{} 写入 XLSX 失败", self.prog_id))?;
        book.close()?;
        Ok(source_format)
    }
}

impl Drop for Application {
    fn drop(&mut self) {
        if self.owned {
            if let Err(error) = method(&self.object, "Quit", vec![]) {
                eprintln!(
                    "{}",
                    crate::run_logging::terminal_line(
                        crate::run_logging::Level::Warn,
                        &format!("无法退出本次启动的 {} 实例：{error:#}", self.prog_id)
                    )
                );
            }
        }
    }
}

struct Workbook {
    object: IDispatch,
    open: bool,
}
impl Workbook {
    fn close(&mut self) -> Result<()> {
        method(&self.object, "Close", vec![AutoVariant::boolean(false)])
            .context("关闭临时源工作簿失败")?;
        self.open = false;
        Ok(())
    }
}
impl Drop for Workbook {
    fn drop(&mut self) {
        if self.open {
            let _ = method(&self.object, "Close", vec![AutoVariant::boolean(false)]);
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
            "进程列表过大，无法安全确认 COM 实例归属"
        );
    }
}

// Automation variants own their BSTR/interface references. The wrapper clears
// every input and output exactly once, including on an Invoke error.
#[repr(transparent)]
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
                other => bail!("COM 返回了非整数类型：{other:?}"),
            }
        }
    }
    fn dispatch(&self) -> Result<IDispatch> {
        unsafe {
            let raw = &self.0.Anonymous.Anonymous;
            ensure!(raw.vt == VT_DISPATCH, "COM 返回了非对象类型：{:?}", raw.vt);
            (*raw.Anonymous.pdispVal)
                .as_ref()
                .cloned()
                .context("COM 返回了空对象")
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
        .with_context(|| format!("查找 COM 操作 {name} 失败"))?;
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
            format!("COM 操作 {name} 失败")
        } else {
            format!("COM 操作 {name} 失败：{}", detail.trim())
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

#[cfg(test)]
mod tests {
    use super::*;
    use excel_reader::read_workbook;
    use std::time::{SystemTime, UNIX_EPOCH};

    struct TestDirectory(PathBuf);
    impl TestDirectory {
        fn new() -> Result<Self> {
            let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
            let path = std::env::temp_dir()
                .join(format!("rust-tools-wps-com-{}-{nonce}", std::process::id()));
            std::fs::create_dir(&path)?;
            Ok(Self(path))
        }
    }
    impl Drop for TestDirectory {
        fn drop(&mut self) {
            let parent = std::env::temp_dir().canonicalize();
            let target = self.0.canonicalize();
            if let (Ok(parent), Ok(target)) = (parent, target) {
                let is_ours = target.parent() == Some(parent.as_path())
                    && target.file_name().is_some_and(|name| {
                        name.to_string_lossy().starts_with("rust-tools-wps-com-")
                    });
                if is_ours {
                    let _ = std::fs::remove_dir_all(target);
                }
            }
        }
    }

    #[test]
    #[ignore = "requires local WPS KET COM; operates only on disposable synthetic copies"]
    fn wps_converts_two_xls_files_and_reuses_its_owned_instance() -> Result<()> {
        let root = TestDirectory::new()?;
        let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/timeline/wide_summary_elapsed_duration.xls");
        let fixture_before = std::fs::read(&fixture)?;
        let first = root.0.join("private-one.xls");
        let second = root.0.join("private-two.xls");
        std::fs::copy(&fixture, &first)?;
        std::fs::copy(&fixture, &second)?;
        let mut context = ComContext::new()?;
        context.ensure_wps()?;
        context.ordinary_choice = Some(AppKind::Wps);
        let first_out = root.0.join("one.xlsx");
        let second_out = root.0.join("two.xlsx");
        let first_result = context.convert(&first, &first_out, "xls")?;
        let second_result = context.convert(&second, &second_out, "xls")?;
        assert_eq!(first_result.application, "KET.Application");
        assert_eq!(first_result.process_id, second_result.process_id);
        let original = read_workbook(&fixture)?;
        let converted = read_workbook(&first_out)?;
        assert_eq!(
            original
                .sheets
                .iter()
                .map(|sheet| &sheet.name)
                .collect::<Vec<_>>(),
            converted
                .sheets
                .iter()
                .map(|sheet| &sheet.name)
                .collect::<Vec<_>>()
        );
        assert!(converted
            .sheets
            .iter()
            .any(|sheet| !sheet.values.is_empty()));
        assert_eq!(std::fs::read(&fixture)?, fixture_before);
        Ok(())
    }
}
