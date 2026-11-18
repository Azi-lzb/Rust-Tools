//! Optional external engine adapter. No menu or automatic routing integration.
#[cfg(test)]
mod archive_tests;
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use sha2::{Digest, Sha256};
use std::fs::{self, File, OpenOptions};
use std::io::{BufRead, BufReader, Read, Write};
use std::path::{Path, PathBuf};
use std::process::{Child, ChildStdin, Command, Stdio};
use std::sync::mpsc::{self, Receiver};
use std::thread;
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};

#[derive(Debug)]
pub struct EngineError(pub String);
impl std::fmt::Display for EngineError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(&self.0)
    }
}
impl std::error::Error for EngineError {}
type Result<T> = std::result::Result<T, EngineError>;
fn err(e: impl std::fmt::Display) -> EngineError {
    EngineError(e.to_string())
}

#[derive(Clone, Debug, Serialize)]
#[serde(tag = "type", content = "value", rename_all = "snake_case")]
pub enum CellValue {
    Text(String),
    Number(f64),
    Boolean(bool),
    Blank,
}

#[derive(Clone, Debug, Default, Serialize)]
pub struct CellStyle {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size: Option<u16>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_rgb: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub number_format: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap_text: Option<bool>,
}

#[derive(Clone, Debug, Serialize)]
#[serde(tag = "op", rename_all = "snake_case")]
pub enum WorkbookOperation {
    /// A serialized, validated ExcelEditPlan primitive. Business rules stay in Rust.
    EditPlanOperation { operation: Value },
    ReplaceHeaderText {
        sheet: String,
        search: String,
        replacement: String,
    },
    SetValue {
        sheet: String,
        cell: String,
        value: CellValue,
    },
    SetFormula {
        sheet: String,
        cell: String,
        formula: String,
    },
    SetStyle {
        sheet: String,
        cell: String,
        style: CellStyle,
    },
}

pub struct WorkbookRequest {
    pub run_id: String,
    pub input: PathBuf,
    pub candidate: PathBuf,
}
pub struct ModifyRequest {
    pub workbook: WorkbookRequest,
    pub operations: Vec<WorkbookOperation>,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct EngineTimings {
    pub startup_ms: f64,
    pub load_ms: f64,
    pub transform_ms: f64,
    pub save_ms: f64,
    #[serde(default)]
    pub preserve_defaults_ms: f64,
    #[serde(default)]
    pub raw_sdk_archive_ms: Option<f64>,
    pub worker_ms: f64,
    pub end_to_end_ms: f64,
}
#[derive(Debug, Serialize)]
pub struct CandidateWorkbook {
    pub path: PathBuf,
    pub evaluation_output: bool,
    /// None means an older/direct Worker did not report a Save-stage archive.
    pub sdk_save_archive: Option<SdkSaveArchive>,
    pub input_trace: Option<PathBuf>,
    pub source_active_index: u64,
    pub output_active_index: u64,
    pub source_sheets: Vec<String>,
    pub output_sheets: Vec<String>,
    pub worker_pid: u64,
    pub worker_attempts: u64,
    pub draining: bool,
    pub timings: EngineTimings,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct SdkSaveArchive {
    pub path: PathBuf,
    pub sha256: String,
    pub manifest_path: PathBuf,
    pub stage: String,
}
pub trait CalcEngine {
    fn calculate(&mut self, request: &WorkbookRequest) -> Result<CandidateWorkbook>;
}
pub trait ModifyEngine {
    fn modify(&mut self, request: &ModifyRequest) -> Result<CandidateWorkbook>;
}
pub trait WorkbookConverterEngine {
    fn convert(&mut self, request: &WorkbookRequest) -> Result<CandidateWorkbook>;
}
#[derive(Debug, Serialize)]
pub struct CandidateDocument {
    pub path: PathBuf,
    pub source_paragraph_hashes: Vec<String>,
    pub source_auxiliary_paragraph_hashes: Vec<String>,
    pub source_image_hashes: Vec<String>,
    pub source_sections: u64,
    pub source_tables: u64,
    pub source_shapes: u64,
    pub worker_pid: u64,
    pub timings: EngineTimings,
}
pub trait DocumentConverterEngine {
    fn convert_document(&mut self, request: &WorkbookRequest) -> Result<CandidateDocument>;
}

pub struct WorkerConfig {
    pub executable: PathBuf,
    pub arguments: Vec<String>,
    pub diagnostics_dir: PathBuf,
    pub startup_timeout: Duration,
    pub request_timeout: Duration,
}
struct Session {
    child: Child,
    stdin: Option<ChildStdin>,
    rx: Receiver<Result<Value>>,
    reader: Option<thread::JoinHandle<()>>,
    writer: Option<thread::JoinHandle<()>>,
}
struct OperationSidecar(Option<PathBuf>);
impl Drop for OperationSidecar {
    fn drop(&mut self) {
        if let Some(path) = &self.0 {
            let _ = fs::remove_file(path);
        }
    }
}
impl Session {
    fn receive(&self, timeout: Duration) -> Result<Value> {
        self.rx
            .recv_timeout(timeout)
            .map_err(|e| err(format!("Worker response unavailable: {e}")))?
    }
    fn send(&mut self, request: &Value, timeout: Duration) -> Result<Value> {
        let data = serde_json::to_vec(request).map_err(err)?;
        if data.len() > 1024 * 1024 {
            return Err(err("Request exceeds 1 MiB"));
        }
        let start = Instant::now();
        let mut stdin = self
            .stdin
            .take()
            .ok_or_else(|| err("Worker stdin closed"))?;
        let (tx, rx) = mpsc::channel();
        self.writer = Some(thread::spawn(move || {
            let result = stdin
                .write_all(&data)
                .and_then(|_| stdin.write_all(b"\n"))
                .and_then(|_| stdin.flush());
            let _ = tx.send((stdin, result));
        }));
        let (stdin, written) = rx
            .recv_timeout(timeout)
            .map_err(|e| err(format!("Worker write timed out/disconnected: {e}")))?;
        self.stdin = Some(stdin);
        if let Some(writer) = self.writer.take() {
            let _ = writer.join();
        }
        written.map_err(err)?;
        self.receive(timeout.saturating_sub(start.elapsed()))
    }
}
impl Drop for Session {
    fn drop(&mut self) {
        self.stdin.take();
        let deadline = Instant::now() + Duration::from_secs(1);
        loop {
            match self.child.try_wait() {
                Ok(Some(_)) => break,
                _ if Instant::now() >= deadline => {
                    let _ = self.child.kill();
                    let _ = self.child.wait();
                    break;
                }
                _ => thread::sleep(Duration::from_millis(10)),
            }
        }
        if let Some(reader) = self.reader.take() {
            let _ = reader.join();
        }
        if let Some(writer) = self.writer.take() {
            let _ = writer.join();
        }
    }
}

pub struct AsposeWorkerBackend {
    config: WorkerConfig,
    session: Option<Session>,
    sequence: u64,
    generation: u64,
    loads: u64,
    engine_name: &'static str,
}
impl AsposeWorkerBackend {
    pub fn new(config: WorkerConfig) -> Self {
        Self {
            config,
            session: None,
            sequence: 0,
            generation: 0,
            loads: 0,
            engine_name: "Aspose.Cells",
        }
    }
    fn start(&mut self) -> Result<f64> {
        if self.session.is_some() {
            return Ok(0.0);
        }
        let start = Instant::now();
        fs::create_dir_all(&self.config.diagnostics_dir).map_err(err)?;
        self.generation += 1;
        let stamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map_err(err)?
            .as_nanos();
        let log = OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(
                self.config
                    .diagnostics_dir
                    .join(format!("worker-{stamp}-{}.jsonl", self.generation)),
            )
            .map_err(err)?;
        let mut command = Command::new(&self.config.executable);
        command.args(&self.config.arguments);
        if self.engine_name == "Aspose.Cells" {
            let raw_dir = self.config.diagnostics_dir.join("raw-sdk-save");
            fs::create_dir_all(&raw_dir).map_err(err)?;
            command.env(
                "RUSTTOOLS_ASPOSE_RAW_ARCHIVE_DIR",
                fs::canonicalize(raw_dir).map_err(err)?,
            );
        }
        #[cfg(windows)]
        {
            use std::os::windows::process::CommandExt;
            command.creation_flags(0x08000000);
        }
        let mut child = command
            .stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .stderr(Stdio::from(log))
            .spawn()
            .map_err(err)?;
        let stdin = child.stdin.take();
        let stdout = child
            .stdout
            .take()
            .ok_or_else(|| err("Missing worker stdout"))?;
        let (tx, rx) = mpsc::channel();
        let reader = thread::spawn(move || {
            let mut reader = BufReader::new(stdout);
            loop {
                let mut buf = Vec::new();
                let result = reader
                    .by_ref()
                    .take(2 * 1024 * 1024 + 1)
                    .read_until(b'\n', &mut buf);
                let value = match result {
                    Ok(0) => {
                        let _ = tx.send(Err(err("Worker exited before response")));
                        break;
                    }
                    Ok(_) if buf.len() > 2 * 1024 * 1024 => {
                        let _ = tx.send(Err(err("Oversized worker response")));
                        break;
                    }
                    Ok(_) => serde_json::from_slice(&buf).map_err(err),
                    Err(e) => Err(err(e)),
                };
                if tx.send(value).is_err() {
                    break;
                }
            }
        });
        let session = Session {
            child,
            stdin,
            rx,
            reader: Some(reader),
            writer: None,
        };
        let ready = session.receive(self.config.startup_timeout)?;
        if ready["ready"] != true
            || ready["protocol_version"] != 1
            || ready["engine"] != self.engine_name
            || ready["pid"].as_u64() != Some(session.child.id() as u64)
        {
            return Err(err("Invalid worker READY handshake"));
        }
        self.session = Some(session);
        self.loads = 0;
        Ok(start.elapsed().as_secs_f64() * 1000.0)
    }
    pub fn health_check(&mut self) -> Result<u64> {
        self.start()?;
        self.sequence += 1;
        let id = format!("ping-{}", self.sequence);
        let q = json!({"protocol_version":1,"request_id":id,"run_id":"health","method":"Ping"});
        let result = self
            .session
            .as_mut()
            .unwrap()
            .send(&q, self.config.request_timeout);
        match result {
            Ok(v) if v["ok"] == true && v["request_id"] == id => {
                v["pid"].as_u64().ok_or_else(|| err("Missing worker pid"))
            }
            _ => {
                self.session.take();
                Err(err("Worker health check failed"))
            }
        }
    }
    pub fn shutdown(&mut self) {
        self.session.take();
    }
    pub fn generations(&self) -> u64 {
        self.generation
    }
    fn execute(
        &mut self,
        req: &WorkbookRequest,
        method: &str,
        operations: &[WorkbookOperation],
    ) -> Result<CandidateWorkbook> {
        let start = Instant::now();
        let input = fs::canonicalize(&req.input).map_err(err)?;
        if req.run_id.is_empty() || !req.candidate.is_absolute() {
            return Err(err("run_id and absolute candidate required"));
        }
        match fs::symlink_metadata(&req.candidate) {
            Ok(_) => return Err(err("Candidate already exists; refusing overwrite")),
            Err(e) if e.kind() == std::io::ErrorKind::NotFound => {}
            Err(e) => return Err(err(e)),
        }
        let before = fingerprint(&input)?;
        let operation_bytes = serde_json::to_vec(operations).map_err(err)?;
        if operation_bytes.len() > 64 * 1024 * 1024 {
            return Err(err("Operations exceed the 64 MiB transaction limit"));
        }
        let mut operation_sidecar = OperationSidecar(None);
        if operation_bytes.len() > 512 * 1024 {
            fs::create_dir_all(&self.config.diagnostics_dir).map_err(err)?;
            let stamp = SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .map_err(err)?
                .as_nanos();
            let path = fs::canonicalize(&self.config.diagnostics_dir)
                .map_err(err)?
                .join(format!("operations-{}-{stamp}.json", std::process::id()));
            let mut file = OpenOptions::new()
                .write(true)
                .create_new(true)
                .open(&path)
                .map_err(err)?;
            operation_sidecar.0 = Some(path);
            let written = file
                .write_all(&operation_bytes)
                .and_then(|()| file.sync_all());
            drop(file);
            written.map_err(err)?;
        }
        let mut import_paths = std::collections::BTreeSet::new();
        for op in operations {
            if let WorkbookOperation::EditPlanOperation { operation: p } = op {
                if p["operation"] == "import_range" {
                    let path = p["source_workbook"]
                        .as_str()
                        .ok_or_else(|| err("Import source path missing"))?;
                    import_paths.insert(fs::canonicalize(path).map_err(err)?);
                }
            }
        }
        let import_fingerprints: Vec<_> = import_paths
            .iter()
            .map(|p| Ok((p, fingerprint(p)?)))
            .collect::<Result<_>>()?;
        let loads = 1 + import_paths.len() as u64;
        if loads > 99 {
            return Err(err("One transaction exceeds workbook load budget"));
        }
        if self.loads + loads > 99 {
            self.session.take();
        }
        let startup_ms = self.start()?;
        self.sequence += 1;
        let id = format!("task-{}", self.sequence);
        let mut q = json!({"protocol_version":1,"request_id":id,"run_id":req.run_id,"method":method,
                       "input":input,"candidate":req.candidate,"operations":operations});
        if let Some(path) = &operation_sidecar.0 {
            q["operations"] = json!([]);
            q["operations_file"] = json!(path);
            let hash = fingerprint(path)?
                .iter()
                .map(|byte| format!("{byte:02X}"))
                .collect::<String>();
            q["operations_sha256"] = json!(hash);
        }
        let expected_pid = self.session.as_ref().unwrap().child.id() as u64;
        // Audit metadata is independent of business validation and never feeds
        // worksheet removal decisions. This scope is deliberately not advertised
        // as the complete set of logical business inputs.
        let input_trace = match write_input_trace(
            &self.config.diagnostics_dir,
            &req.run_id,
            &id,
            method,
            std::iter::once(input.as_path()).chain(import_paths.iter().map(PathBuf::as_path)),
        ) {
            Ok(path) => Some(path),
            Err(error) => {
                eprintln!("Aspose 输入追溯记录写入失败：{error}");
                None
            }
        };
        let reply = self
            .session
            .as_mut()
            .unwrap()
            .send(&q, self.config.request_timeout);
        let result = match reply {
            Ok(v) if v["request_id"] == id => v,
            _ => {
                self.session.take();
                return Err(err(
                    "Worker response lost/invalid; task NOT automatically retried",
                ));
            }
        };
        if result["draining"] == true {
            self.session.take();
        }
        if fingerprint(&input)? != before {
            return Err(err("Source changed; refusing candidate"));
        }
        for (path, before) in import_fingerprints {
            if fingerprint(path)? != before {
                return Err(err("Import source changed; refusing candidate"));
            }
        }
        self.loads = result["attempts"].as_u64().unwrap_or(99);
        if result["ok"] != true {
            return Err(err(format!("{}: {}", result["code"], result["message"])));
        }
        let attempts = result["attempts"]
            .as_u64()
            .ok_or_else(|| err("Missing attempt count"))?;
        if result["pid"].as_u64() != Some(expected_pid)
            || !(1..=99).contains(&attempts)
            || (result["draining"] == true) != (attempts == 99)
        {
            self.session.take();
            return Err(err("Worker identity or lifecycle counter mismatch"));
        }
        if result["candidate"].as_str() != req.candidate.to_str() {
            return Err(err("Candidate path mismatch"));
        }
        let evaluation_output = result["evaluation_output"]
            .as_bool()
            .ok_or_else(|| err("Missing evaluation marker"))?;
        let sdk_save_archive = verify_sdk_archive(
            &result["sdk_save_archive"],
            &self.config.diagnostics_dir,
            &req.run_id,
            &id,
        )?;
        let source_sheets: Vec<String> =
            serde_json::from_value(result["source_sheets"].clone()).map_err(err)?;
        if input
            .extension()
            .is_some_and(|e| e.eq_ignore_ascii_case("xlsx"))
            && inspect_candidate(&input)?.0 != source_sheets
        {
            return Err(err("Worker source worksheet metadata differs from input"));
        }
        let output_sheets: Vec<String> =
            serde_json::from_value(result["output_sheets"].clone()).map_err(err)?;
        let (actual_sheets, active) = inspect_candidate(&req.candidate)?;
        let mut expected_sheets = source_sheets.clone();
        for op in operations {
            if let WorkbookOperation::EditPlanOperation { operation: p } = op {
                let text = |key: &str| p[key].as_str().unwrap_or_default().to_owned();
                match p["operation"].as_str() {
                    Some("rename_sheet") => {
                        let position = expected_sheets
                            .iter()
                            .position(|s| *s == text("old_name"))
                            .ok_or_else(|| err("Rename source absent"))?;
                        expected_sheets[position] = text("new_name");
                    }
                    Some("delete_sheet") => expected_sheets.retain(|s| *s != text("name")),
                    Some("create_sheet") => expected_sheets.push(text("name")),
                    Some("create_sheet_at") => {
                        let index = p["index"]
                            .as_u64()
                            .ok_or_else(|| err("Invalid sheet index"))?
                            as usize;
                        if index > expected_sheets.len() {
                            return Err(err("Sheet index out of bounds"));
                        }
                        expected_sheets.insert(index, text("name"));
                    }
                    Some("move_sheet") => {
                        let position = expected_sheets
                            .iter()
                            .position(|s| *s == text("sheet"))
                            .ok_or_else(|| err("Move source absent"))?;
                        let name = expected_sheets.remove(position);
                        let index = p["new_index"]
                            .as_u64()
                            .ok_or_else(|| err("Invalid move index"))?
                            as usize;
                        if index > expected_sheets.len() {
                            return Err(err("Move index out of bounds"));
                        }
                        expected_sheets.insert(index, name);
                    }
                    _ => {}
                }
            }
        }
        if actual_sheets != output_sheets
            || active >= actual_sheets.len() as u64
            || active != result["output_active_index"].as_u64().unwrap_or(u64::MAX)
        {
            return Err(err("Candidate workbook metadata verification failed"));
        }
        verify_business_roster(&actual_sheets, &expected_sheets)?;
        Ok(CandidateWorkbook {
            path: req.candidate.clone(),
            evaluation_output,
            sdk_save_archive,
            input_trace,
            source_active_index: result["source_active_index"]
                .as_u64()
                .ok_or_else(|| err("Missing source active index"))?,
            output_active_index: active,
            source_sheets,
            output_sheets,
            worker_pid: result["pid"].as_u64().ok_or_else(|| err("Missing pid"))?,
            worker_attempts: result["attempts"]
                .as_u64()
                .ok_or_else(|| err("Missing attempt count"))?,
            draining: result["draining"] == true,
            timings: EngineTimings {
                startup_ms,
                load_ms: result["load_ms"]
                    .as_f64()
                    .ok_or_else(|| err("Missing load timing"))?,
                transform_ms: result["transform_ms"]
                    .as_f64()
                    .ok_or_else(|| err("Missing transform timing"))?,
                save_ms: result["save_ms"]
                    .as_f64()
                    .ok_or_else(|| err("Missing save timing"))?,
                preserve_defaults_ms: result["preserve_defaults_ms"].as_f64().unwrap_or(0.0),
                raw_sdk_archive_ms: result["raw_sdk_archive_ms"].as_f64(),
                worker_ms: result["processing_ms"]
                    .as_f64()
                    .ok_or_else(|| err("Missing timing"))?,
                end_to_end_ms: start.elapsed().as_secs_f64() * 1000.0,
            },
        })
    }
}
/// Shares bounded JSONL transport/lifecycle with Cells, using a separate DLL
/// and protocol engine identity. No document process calls enter business code.
pub struct AsposeDocumentBackend(AsposeWorkerBackend);
impl AsposeDocumentBackend {
    pub fn new(config: WorkerConfig) -> Self {
        let mut host = AsposeWorkerBackend::new(config);
        host.engine_name = "Aspose.Words";
        Self(host)
    }
}
impl DocumentConverterEngine for AsposeDocumentBackend {
    fn convert_document(&mut self, request: &WorkbookRequest) -> Result<CandidateDocument> {
        let host = &mut self.0;
        let clock = Instant::now();
        let input = fs::canonicalize(&request.input).map_err(err)?;
        if request.run_id.is_empty()
            || !request.candidate.is_absolute()
            || request.candidate.exists()
        {
            return Err(err("New absolute candidate and run_id required"));
        }
        let hash = fingerprint(&input)?;
        let startup_ms = host.start()?;
        host.sequence += 1;
        let id = format!("document-{}", host.sequence);
        let expected_pid = host.session.as_ref().unwrap().child.id() as u64;
        let q = json!({"protocol_version":1,"request_id":id,"run_id":request.run_id,"method":"Convert","input":input,"candidate":request.candidate});
        let response = host
            .session
            .as_mut()
            .unwrap()
            .send(&q, host.config.request_timeout);
        let result = match response {
            Ok(v) if v["request_id"] == id && v["pid"].as_u64() == Some(expected_pid) => v,
            _ => {
                host.session.take();
                return Err(err(
                    "Document worker response lost/invalid; task NOT automatically retried",
                ));
            }
        };
        if result["draining"] == true {
            host.session.take();
        }
        if fingerprint(&input)? != hash {
            return Err(err("Document source changed"));
        }
        if result["ok"] != true {
            return Err(err(format!("{}: {}", result["code"], result["message"])));
        }
        let attempts = result["attempts"]
            .as_u64()
            .ok_or_else(|| err("Missing document lifecycle counter"))?;
        if !(1..=99).contains(&attempts)
            || (result["draining"] == true) != (attempts == 99)
            || result["candidate"].as_str() != request.candidate.to_str()
            || result["evaluation_output"] != true
        {
            host.session.take();
            return Err(err("Document candidate metadata mismatch"));
        }
        let mut zip =
            zip::ZipArchive::new(File::open(&request.candidate).map_err(err)?).map_err(err)?;
        let mut xml = String::new();
        zip.by_name("word/document.xml")
            .map_err(err)?
            .read_to_string(&mut xml)
            .map_err(err)?;
        let document = roxmltree::Document::parse(&xml).map_err(err)?;
        if document.root_element().tag_name().name() != "document" {
            return Err(err("Invalid DOCX document root"));
        }
        let number = |key: &str| {
            result[key]
                .as_f64()
                .ok_or_else(|| err(format!("Missing document timing {key}")))
        };
        Ok(CandidateDocument {
            path: request.candidate.clone(),
            source_paragraph_hashes: serde_json::from_value(
                result["source_paragraph_hashes"].clone(),
            )
            .map_err(err)?,
            source_auxiliary_paragraph_hashes: serde_json::from_value(
                result["source_auxiliary_paragraph_hashes"].clone(),
            )
            .map_err(err)?,
            source_image_hashes: serde_json::from_value(result["source_image_hashes"].clone())
                .map_err(err)?,
            source_sections: result["source_sections"]
                .as_u64()
                .ok_or_else(|| err("Missing section count"))?,
            source_tables: result["source_tables"]
                .as_u64()
                .ok_or_else(|| err("Missing table count"))?,
            source_shapes: result["source_shapes"]
                .as_u64()
                .ok_or_else(|| err("Missing shape count"))?,
            worker_pid: expected_pid,
            timings: EngineTimings {
                startup_ms,
                load_ms: number("load_ms")?,
                transform_ms: 0.0,
                save_ms: number("save_ms")?,
                preserve_defaults_ms: 0.0,
                raw_sdk_archive_ms: None,
                worker_ms: number("processing_ms")?,
                end_to_end_ms: clock.elapsed().as_secs_f64() * 1000.0,
            },
        })
    }
}
impl CalcEngine for AsposeWorkerBackend {
    fn calculate(&mut self, req: &WorkbookRequest) -> Result<CandidateWorkbook> {
        self.execute(req, "Calculate", &[])
    }
}
impl WorkbookConverterEngine for AsposeWorkerBackend {
    fn convert(&mut self, req: &WorkbookRequest) -> Result<CandidateWorkbook> {
        self.execute(req, "Convert", &[])
    }
}
impl ModifyEngine for AsposeWorkerBackend {
    fn modify(&mut self, req: &ModifyRequest) -> Result<CandidateWorkbook> {
        if req.operations.is_empty() {
            return Err(err("No modification operations"));
        }
        for op in &req.operations {
            if let WorkbookOperation::SetValue {
                value: CellValue::Number(n),
                ..
            } = op
            {
                if !n.is_finite() {
                    return Err(err("Nonfinite cell number"));
                }
            }
        }
        self.execute(&req.workbook, "Modify", &req.operations)
    }
}

fn verify_business_roster(actual: &[String], expected: &[String]) -> Result<()> {
    if expected.is_empty() || actual != expected {
        return Err(err("Candidate must contain exactly the planned business worksheets"));
    }
    Ok(())
}

fn fingerprint(path: &Path) -> Result<Vec<u8>> {
    Ok(Sha256::digest(fs::read(path).map_err(err)?).to_vec())
}

fn hex_fingerprint(path: &Path) -> Result<String> {
    Ok(fingerprint(path)?
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect())
}

fn trace_input(path: &Path) -> Value {
    let before = hex_fingerprint(path).ok();
    let names = (|| -> Result<Vec<String>> {
        let mut zip = zip::ZipArchive::new(File::open(path).map_err(err)?).map_err(err)?;
        let mut part = zip.by_name("xl/workbook.xml").map_err(err)?;
        if part.size() > 64 * 1024 * 1024 {
            return Err(err("workbook metadata exceeds audit limit"));
        }
        let mut xml = String::new();
        part.read_to_string(&mut xml).map_err(err)?;
        let document = roxmltree::Document::parse(&xml).map_err(err)?;
        let names: Vec<String> = document
            .descendants()
            .filter(|n| {
                n.has_tag_name((
                    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                    "sheet",
                ))
            })
            .map(|n| {
                n.attribute("name")
                    .ok_or_else(|| err("worksheet name absent"))
                    .map(str::to_owned)
            })
            .collect::<Result<_>>()?;
        if names.is_empty() {
            return Err(err("worksheet inventory absent"));
        }
        Ok(names)
    })();
    let stable = before.is_some() && before == hex_fingerprint(path).ok();
    match names {
        Ok(names) if stable => {
            json!({"path":path,"sha256":before,"sheets":names,"sheets_status":"known", "metadata_source":"OOXML workbook part before dispatch"})
        }
        _ => {
            json!({"path":path,"sha256":if stable { before } else { None },"sheets":null,"sheets_status":"unknown", "unknown_reason":if stable { "not independently readable as OOXML; XLS inventory not implemented in this adapter" } else { "input unreadable or changed during metadata capture" }})
        }
    }
}

fn write_input_trace<'a>(
    directory: &Path,
    run_id: &str,
    request_id: &str,
    method: &str,
    inputs: impl Iterator<Item = &'a Path>,
) -> Result<PathBuf> {
    let root = directory.join("input-traces");
    fs::create_dir_all(&root).map_err(err)?;
    let root = fs::canonicalize(root).map_err(err)?;
    let stamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map_err(err)?
        .as_nanos();
    let path = root.join(format!("inputs-{}-{stamp}.json", std::process::id()));
    let inventory: Vec<_> = inputs.map(trace_input).collect();
    let mut file = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&path)
        .map_err(err)?;
    serde_json::to_writer_pretty(&mut file, &json!({
        "schema_version":1, "event_name":"aspose_worker_inputs_captured", "run_id":run_id,
        "request_id":request_id,"method":method,"engine":"aspose_worker",
        "input_scope":"worker_request_primary_and_imports", "logical_business_inputs_complete":false,
        "unknown_business_inputs_reason":"upstream business inputs serialized into operations are not visible to this adapter",
        "inputs":inventory
    })).map_err(err)?;
    file.flush().and_then(|()| file.sync_all()).map_err(err)?;
    Ok(path)
}

fn verify_sdk_archive(
    value: &Value,
    diagnostics: &Path,
    run_id: &str,
    request_id: &str,
) -> Result<Option<SdkSaveArchive>> {
    if value.is_null() {
        return Ok(None);
    } // Older Workers do not have this capability.
    let archive: SdkSaveArchive = serde_json::from_value(value.clone()).map_err(err)?;
    let expected_root = fs::canonicalize(diagnostics.join("raw-sdk-save")).map_err(err)?;
    for path in [&archive.path, &archive.manifest_path] {
        if !path.is_absolute()
            || fs::canonicalize(path.parent().ok_or_else(|| err("archive parent absent"))?)
                .map_err(err)?
                != expected_root
        {
            return Err(err("SDK archive outside configured evidence directory"));
        }
        let metadata = fs::symlink_metadata(path).map_err(err)?;
        if !metadata.is_file() || metadata.file_type().is_symlink() {
            return Err(err("SDK archive is not a regular file"));
        }
    }
    let hash = hex_fingerprint(&archive.path)?;
    if archive.stage != "sdk_save_before_worker_preservation"
        || !hash.eq_ignore_ascii_case(&archive.sha256)
    {
        return Err(err("SDK archive stage or hash mismatch"));
    }
    let manifest: Value =
        serde_json::from_slice(&fs::read(&archive.manifest_path).map_err(err)?).map_err(err)?;
    if manifest["stage"] != archive.stage
        || manifest["run_id"] != run_id
        || manifest["request_id"] != request_id
        || manifest["archive_sha256"]
            .as_str()
            .is_none_or(|s| !hash.eq_ignore_ascii_case(s))
    {
        return Err(err("SDK archive manifest identity mismatch"));
    }
    Ok(Some(archive))
}
fn inspect_candidate(path: &Path) -> Result<(Vec<String>, u64)> {
    let mut zip = zip::ZipArchive::new(File::open(path).map_err(err)?).map_err(err)?;
    for required in [
        "[Content_Types].xml",
        "_rels/.rels",
        "xl/_rels/workbook.xml.rels",
    ] {
        let mut text = String::new();
        let mut part = zip.by_name(required).map_err(err)?;
        if part.size() > 8 * 1024 * 1024 {
            return Err(err("Package metadata too large"));
        }
        part.read_to_string(&mut text).map_err(err)?;
        roxmltree::Document::parse(&text).map_err(err)?;
    }
    let mut xml = String::new();
    let mut entry = zip.by_name("xl/workbook.xml").map_err(err)?;
    if entry.size() > 8 * 1024 * 1024 {
        return Err(err("Workbook metadata too large"));
    }
    entry.read_to_string(&mut xml).map_err(err)?;
    let doc = roxmltree::Document::parse(&xml).map_err(err)?;
    let names: Vec<String> = doc
        .descendants()
        .filter(|n| n.has_tag_name("sheet"))
        .filter_map(|n| n.attribute("name").map(str::to_owned))
        .collect();
    let active = doc
        .descendants()
        .find(|n| n.has_tag_name("workbookView"))
        .and_then(|n| n.attribute("activeTab"))
        .unwrap_or("0")
        .parse()
        .map_err(err)?;
    if names.is_empty() || active >= names.len() as u64 {
        return Err(err("Invalid candidate sheets/activeTab"));
    }
    drop(entry);
    let mut rel_xml = String::new();
    zip.by_name("xl/_rels/workbook.xml.rels")
        .map_err(err)?
        .read_to_string(&mut rel_xml)
        .map_err(err)?;
    let rels = roxmltree::Document::parse(&rel_xml).map_err(err)?;
    for sheet in doc.descendants().filter(|n| n.has_tag_name("sheet")) {
        let id = sheet
            .attribute((
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                "id",
            ))
            .ok_or_else(|| err("Sheet missing relationship id"))?;
        let rel = rels
            .descendants()
            .find(|n| n.has_tag_name("Relationship") && n.attribute("Id") == Some(id))
            .ok_or_else(|| err("Sheet relationship missing"))?;
        if rel.attribute("TargetMode") == Some("External") {
            return Err(err("External worksheet relationship"));
        }
        let target = rel
            .attribute("Target")
            .ok_or_else(|| err("Relationship target missing"))?;
        if target.contains("..") || target.contains('\\') {
            return Err(err("Unsafe worksheet target"));
        }
        let part_name = if target.starts_with('/') {
            target.trim_start_matches('/').to_owned()
        } else {
            format!("xl/{target}")
        };
        let mut part = zip.by_name(&part_name).map_err(err)?;
        if part.size() > 64 * 1024 * 1024 {
            return Err(err("Worksheet XML over validated size limit"));
        }
        let mut text = String::new();
        part.read_to_string(&mut text).map_err(err)?;
        let worksheet = roxmltree::Document::parse(&text).map_err(err)?;
        if worksheet.root_element().tag_name().name() != "worksheet" {
            return Err(err("Invalid worksheet XML"));
        }
    }
    Ok((names, active))
}
