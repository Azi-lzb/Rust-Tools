//! Windows format conversion through native Rust COM automation.
//!
//! COM objects remain on one dedicated STA thread. A batch reuses only the
//! Office/WPS instances it created, while every input handed to this module is
//! already a disposable copy. Spreadsheets route through Excel/WPS/KET, Word
//! documents through Word/WPS Writer - both on the same worker thread and the
//! same batch-scoped instance ownership rules.

use anyhow::{anyhow, bail, ensure, Context, Result};
use std::{
    collections::{BTreeSet, HashMap},
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
            ProcessStatus::{K32EnumProcessModulesEx, K32EnumProcesses, K32GetModuleBaseNameW},
            Threading::{
                OpenProcess, TerminateProcess, PROCESS_QUERY_LIMITED_INFORMATION,
                PROCESS_TERMINATE,
            },
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
    ConvertWord {
        input: PathBuf,
        output: PathBuf,
        reply: Sender<std::result::Result<WordConversion, String>>,
    },
    ExtractWordText {
        input: PathBuf,
        reply: Sender<std::result::Result<String, String>>,
    },
    Stop,
}

pub(crate) struct SpreadsheetConversion {
    pub source_format: Option<i32>,
    pub application: &'static str,
    pub process_id: u32,
}

pub(crate) struct WordConversion {
    pub application: &'static str,
    /// Some for owned instances; None when connected to a shared host.
    pub process_id: Option<u32>,
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
            .name("pytools-com".to_owned())
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

    /// Convert one disposable Word document copy to .docx through the
    /// batch-owned Word/WPS Writer instance.
    pub(crate) fn convert_word(
        &mut self,
        input: &Path,
        output: &Path,
    ) -> Result<WordConversion> {
        let (reply, response) = mpsc::channel();
        self.requests
            .send(Request::ConvertWord {
                input: input.to_owned(),
                output: output.to_owned(),
                reply,
            })
            .context("Word COM 工作线程不可用")?;
        response
            .recv()
            .context("Word COM 工作线程未返回结果")?
            .map_err(|message| anyhow!("{message}"))
    }

    /// Extract the plain text of one disposable document copy for fidelity
    /// validation, through the same batch-owned Word/WPS Writer instance.
    pub(crate) fn extract_word_text(&mut self, input: &Path) -> Result<String> {
        let (reply, response) = mpsc::channel();
        self.requests
            .send(Request::ExtractWordText {
                input: input.to_owned(),
                reply,
            })
            .context("Word COM 工作线程不可用")?;
        response
            .recv()
            .context("Word COM 工作线程未返回结果")?
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
            Ok(Request::ConvertWord { input, output, reply }) => {
                let result = context
                    .convert_word(&input, &output)
                    .map_err(|error| format!("{error:#}"));
                let _ = reply.send(result);
            }
            Ok(Request::ExtractWordText { input, reply }) => {
                let result = context
                    .extract_word_text(&input)
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
    word: Option<Application>,
    ordinary_choice: Option<AppKind>,
    _apartment: Apartment,
}

impl ComContext {
    fn new() -> Result<Self> {
        Ok(Self {
            excel: None,
            wps: None,
            word: None,
            ordinary_choice: None,
            _apartment: Apartment::initialize()?,
        })
    }

    fn ensure_word(&mut self) -> Result<()> {
        if self.word.is_none() {
            self.word = Some(
                Application::create_first(
                    &[
                        "Word.Application",
                        "kwps.Application",
                        "KWPS.Application",
                        "wps.Application",
                    ],
                    AppFamily::Word,
                )?
                .with_context(|| {
                    "No supported Office/WPS COM engine is available: Word.Application / kwps.Application / KWPS.Application / wps.Application"
                })?,
            );
        }
        Ok(())
    }

    fn convert_word(
        &mut self,
        input: &Path,
        output: &Path,
    ) -> Result<WordConversion> {
        self.ensure_word()?;
        let app = self.word.as_mut().expect("Word/WPS Writer selected");
        let result = app
            .convert_word(input, output)
            .map(|()| WordConversion {
                application: app.prog_id,
                process_id: app.owned.then_some(app.process_id),
            });
        if result.is_err() {
            // A failed Open/SaveAs can leave the writer in an unknown state.
            // Drop only our own instance so the next file starts cleanly.
            self.word = None;
        }
        result
    }

    fn extract_word_text(&mut self, input: &Path) -> Result<String> {
        self.ensure_word()?;
        let app = self.word.as_mut().expect("Word/WPS Writer selected");
        let result = app.extract_word_text(input);
        if result.is_err() {
            self.word = None;
        }
        result
    }

    fn ensure_wps(&mut self) -> Result<()> {
        if self.wps.is_none() {
            self.wps = Some(
                Application::create_first(&["KET.Application"], AppFamily::Spreadsheet)?
                .context(
                    "No supported Office/WPS COM engine is available: WPS KET.Application",
                )?,
            );
        }
        Ok(())
    }

    fn choose_ordinary(&mut self) -> Result<AppKind> {
        if let Some(kind) = self.ordinary_choice {
            return Ok(kind);
        }
        if self.excel.is_none() {
            self.excel = Application::create_first(&["Excel.Application"], AppFamily::Spreadsheet)?;
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
        drop(self.word.take());
    }
}

struct Application {
    object: IDispatch,
    prog_id: &'static str,
    process_id: u32,
    /// True when this creation spawned a verifiably new process; only such
    /// instances are unconditionally Quit on drop.
    owned: bool,
    /// True when COM connected to an already-running shared host (no new
    /// process, no ownership). Such instances are quit on drop only while
    /// they still hold zero open documents.
    connected_shared: bool,
}

/// Host-application family: the DisplayAlerts value type and the document
/// collection name differ between Excel/WPS Spreadsheets and Word/WPS Writer.
#[derive(Clone, Copy, Debug)]
enum AppFamily {
    Spreadsheet,
    Word,
}

impl AppFamily {
    fn collection(self) -> &'static str {
        match self {
            Self::Spreadsheet => "Workbooks",
            Self::Word => "Documents",
        }
    }
}

impl Application {
    /// An unavailable registration is eligible for the next configured
    /// application. Once an object exists, ownership or setup errors are
    /// returned explicitly rather than silently switching engines.
    fn create_first(names: &[&'static str], family: AppFamily) -> Result<Option<Self>> {
        for &prog_id in names {
            let before_ids = process_ids()?;
            let before_snapshot = process_snapshot()?;
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
                connected_shared: false,
            };
            // Preferred: the application reports its own main window handle.
            let window_pid = property(&app.object, "Hwnd")
                .ok()
                .and_then(|hwnd| hwnd.integer().ok())
                .map(|value| value as u32)
                .filter(|pid| *pid != 0);
            let process_id = match window_pid {
                Some(pid) if !before_ids.contains(&pid) => Some(pid),
                _ => {
                    // Headless instances may not expose a main window: fall
                    // back to the process that appeared alongside creation.
                    let family_images: &[&str] = match family {
                        AppFamily::Spreadsheet => &["excel.exe", "et.exe", "wps.exe"],
                        AppFamily::Word => &["winword.exe", "wps.exe"],
                    };
                    let snapshot = process_snapshot()?;
                    newly_owned_process(&before_ids, &snapshot, family_images)
                }
            };
            let collection_dispatch =
                property(&app.object, family.collection())?.dispatch()?;
            let count = property(&collection_dispatch, "Count")?.integer()?;
            ensure!(
                count == 0,
                "{prog_id} 连接到的宿主已打开文档；为保护用户数据，停止转换"
            );
            match process_id {
                Some(pid) => {
                    ensure!(
                        !before_ids.contains(&pid),
                        "{prog_id} 连接到了已存在的进程实例；为保护已打开的文档，停止转换"
                    );
                    app.process_id = pid;
                    app.owned = true;
                }
                None => {
                    // Shared multi-use host (e.g. WPS keeps one process for
                    // every window): no new process, no ownership. Safe to
                    // use while it stays document-free; never Quit blindly.
                    app.connected_shared = true;
                }
            }
            put(&app.object, "Visible", AutoVariant::boolean(false))?;
            match family {
                AppFamily::Spreadsheet => {
                    put(&app.object, "DisplayAlerts", AutoVariant::boolean(false))?;
                }
                AppFamily::Word => {
                    // Word expects the WdAlertLevel enum (wdAlertsNone = 0),
                    // not a Boolean.
                    put(&app.object, "DisplayAlerts", AutoVariant::int(0))?;
                }
            }
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

    /// Open a disposable source document read-only and save it as .docx.
    /// Older WPS Writer builds lack SaveAs2, so SaveAs is the explicit
    /// fallback - the candidate is validated afterwards either way.
    fn convert_word(&mut self, input: &Path, output: &Path) -> Result<()> {
        let documents = property(&self.object, "Documents")?.dispatch()?;
        let object = method(
            &documents,
            "Open",
            vec![
                AutoVariant::text(&input.to_string_lossy()),
                AutoVariant::boolean(false),
                AutoVariant::boolean(true),
            ],
        )
        .with_context(|| format!("{} 打开源文档失败", self.prog_id))?
        .dispatch()?;
        let mut document = WordDocument { object, open: true };
        let save_as2 = method(
            &document.object,
            "SaveAs2",
            vec![
                AutoVariant::text(&output.to_string_lossy()),
                AutoVariant::int(16),
            ],
        );
        if save_as2.is_err() {
            method(
                &document.object,
                "SaveAs",
                vec![
                    AutoVariant::text(&output.to_string_lossy()),
                    AutoVariant::int(16),
                ],
            )
            .with_context(|| format!("{} 写入 DOCX 失败", self.prog_id))?;
        }
        document.close()?;
        Ok(())
    }

    /// Plain text of a disposable document copy for fidelity validation.
    fn extract_word_text(&mut self, input: &Path) -> Result<String> {
        let documents = property(&self.object, "Documents")?.dispatch()?;
        let object = method(
            &documents,
            "Open",
            vec![
                AutoVariant::text(&input.to_string_lossy()),
                AutoVariant::boolean(false),
                AutoVariant::boolean(true),
            ],
        )
        .with_context(|| format!("{} 打开源文档失败", self.prog_id))?
        .dispatch()?;
        let mut document = WordDocument { object, open: true };
        let content = property(&document.object, "Content")?.dispatch()?;
        let value = property(&content, "Text")?;
        let text = value.string()?;
        document.close()?;
        Ok(text)
    }
}

struct WordDocument {
    object: IDispatch,
    open: bool,
}
impl WordDocument {
    fn close(&mut self) -> Result<()> {
        method(&self.object, "Close", vec![AutoVariant::boolean(false)])
            .context("关闭临时源文档失败")?;
        self.open = false;
        Ok(())
    }
}
impl Drop for WordDocument {
    fn drop(&mut self) {
        if self.open {
            let _ = method(&self.object, "Close", vec![AutoVariant::boolean(false)]);
        }
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
            // Word Quit may linger while marshalled references are released.
            // Wait briefly, then terminate the exact process this instance
            // created (never a process it merely connected to).
            if self.process_id != 0 {
                for _ in 0..12 {
                    thread::sleep(Duration::from_millis(250));
                    match process_ids() {
                        Ok(ids) if !ids.contains(&self.process_id) => return,
                        Ok(_) => {}
                        Err(_) => return,
                    }
                }
                unsafe {
                    if let Ok(handle) = OpenProcess(
                        PROCESS_TERMINATE | PROCESS_QUERY_LIMITED_INFORMATION,
                        false,
                        self.process_id,
                    ) {
                        let _ = TerminateProcess(handle, 0);
                        let _ = windows::Win32::Foundation::CloseHandle(handle);
                    }
                }
                eprintln!(
                    "{}",
                    crate::run_logging::terminal_line(
                        crate::run_logging::Level::Warn,
                        &format!("{} 实例未响应 Quit，已按 PID 终止本次启动的进程", self.prog_id)
                    )
                );
            }
            return;
        }
        if self.connected_shared {
            // Quit a shared host only while it still holds no documents;
            // anything the user opened meanwhile must survive.
            let documents_empty = property(&self.object, family_collection(&self.prog_id))
                .ok()
                .and_then(|collection| collection.dispatch().ok())
                .and_then(|collection| property(&collection, "Count").ok())
                .and_then(|count| count.integer().ok())
                .map(|count| count == 0);
            if documents_empty == Some(true) {
                let _ = method(&self.object, "Quit", vec![]);
            }
        }
    }
}

/// The host document collection for a ProgID, resolved without the family
/// enum for drop-time checks.
fn family_collection(prog_id: &str) -> &'static str {
    match prog_id {
        "Word.Application" | "kwps.Application" | "KWPS.Application" | "wps.Application" => {
            "Documents"
        }
        _ => "Workbooks",
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

/// PID -> lowercase executable image name for every process we can name.
/// Used to prove which process a freshly created COM instance belongs to.
fn process_snapshot() -> Result<HashMap<u32, String>> {
    let mut names = HashMap::new();
    for pid in process_ids()? {
        let Ok(handle) = (unsafe {
            windows::Win32::System::Threading::OpenProcess(
                windows::Win32::System::Threading::PROCESS_QUERY_LIMITED_INFORMATION,
                false,
                pid,
            )
        }) else {
            continue;
        };
        let mut buffer = [0u16; 512];
        let written = unsafe { K32GetModuleBaseNameW(handle, None, &mut buffer) };
        unsafe {
            let _ = windows::Win32::Foundation::CloseHandle(handle);
        }
        if written > 0 && (written as usize) < buffer.len() {
            let image = String::from_utf16_lossy(&buffer[..written as usize]);
            names.insert(pid, image.to_ascii_lowercase());
        }
    }
    Ok(names)
}

/// The PID of the process this COM creation spawned: the single new process
/// whose image name matches the expected host application. Returns None when
/// the evidence is ambiguous rather than guessing.
fn newly_owned_process(
    before: &BTreeSet<u32>,
    snapshots: &HashMap<u32, String>,
    image_names: &[&str],
) -> Option<u32> {
    let candidates: Vec<u32> = snapshots
        .iter()
        .filter(|(pid, _)| !before.contains(*pid))
        .filter(|(_, name)| image_names.iter().any(|image| name == image))
        .map(|(pid, _)| *pid)
        .collect();
    if candidates.len() == 1 {
        Some(candidates[0])
    } else {
        None
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
    fn string(&self) -> Result<String> {
        unsafe {
            let raw = &self.0.Anonymous.Anonymous;
            ensure!(raw.vt == VT_BSTR, "COM 返回了非文本类型：{:?}", raw.vt);
            let bstr = raw.Anonymous.bstrVal.as_ref();
            let slice: &[u16] = bstr;
            Ok(String::from_utf16_lossy(slice))
        }
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

    /// Seed one synthetic .doc file through the local Word/WPS Writer COM so
    /// the conversion test operates on a disposable, self-created document.
    fn create_synthetic_doc(path: &Path, sentence: &str) -> Result<()> {
        let script = format!(
            "$ErrorActionPreference='Stop';
$app=$null;
foreach ($name in @('Word.Application','kwps.Application','KWPS.Application','wps.Application')) {{ if ($null -eq $app) {{ try {{ $app=New-Object -ComObject $name }} catch {{}} }} }}
if ($null -eq $app) {{ throw 'No Word/WPS Writer COM engine available' }}
try {{ $app.Visible=$false; $doc=$app.Documents.Add(); $doc.Content.Text='{1}'; $doc.SaveAs2('{0}',0); $doc.Close($false) }} finally {{ $app.Quit() }}
",
            path.to_string_lossy().replace("'", "''"),
            sentence.replace('\'', "''"),
        );
        let script_path = path.with_extension("make-doc.ps1");
        std::fs::write(&script_path, script)?;
        let output = std::process::Command::new("powershell.exe")
            .args(["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File"])
            .arg(&script_path)
            .output()?;
        let _ = std::fs::remove_file(&script_path);
        anyhow::ensure!(
            output.status.success() && path.is_file(),
            "create synthetic .doc failed: {}",
            String::from_utf8_lossy(&output.stderr)
        );
        Ok(())
    }

    #[test]
    #[ignore = "requires local Word/WPS Writer COM; operates only on disposable synthetic documents"]
    fn word_com_converts_two_docs_and_reuses_its_owned_instance() -> Result<()> {
        let root = TestDirectory::new()?;
        let sentence = "pytools word com session reuse";
        let first_doc = root.0.join("one.doc");
        let second_doc = root.0.join("two.doc");
        create_synthetic_doc(&first_doc, sentence)?;
        create_synthetic_doc(&second_doc, sentence)?;
        let first_before = std::fs::read(&first_doc)?;
        let second_before = std::fs::read(&second_doc)?;

        let mut context = ComContext::new()?;
        context.ensure_word()?;
        let first_out = root.0.join("one.docx");
        let second_out = root.0.join("two.docx");
        let first = context.convert_word(&first_doc, &first_out)?;
        let second = context.convert_word(&second_doc, &second_out)?;
        // Session-reuse evidence: one owned Writer instance served both files.
        assert_eq!(first.process_id, second.process_id);
        assert_eq!(first.application, second.application);

        for output in [&first_out, &second_out] {
            let bytes = std::fs::read(output)?;
            anyhow::ensure!(
                bytes.starts_with(b"PK"),
                "converted DOCX is not a valid OOXML ZIP package"
            );
        }
        // Fidelity validation contract: same extraction path, same text.
        let source_text = context.extract_word_text(&first_doc)?;
        let candidate_text = context.extract_word_text(&first_out)?;
        let normalize = |text: &str| text.replace("\r\n", "\n").trim_end_matches('\n').to_owned();
        assert_eq!(normalize(&source_text), normalize(&candidate_text));
        // The caller-owned sources stay untouched.
        assert_eq!(std::fs::read(&first_doc)?, first_before);
        assert_eq!(std::fs::read(&second_doc)?, second_before);
        Ok(())
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
