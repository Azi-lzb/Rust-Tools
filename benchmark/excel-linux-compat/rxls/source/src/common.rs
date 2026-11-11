//! Shared types and utilities for the XLS→XLSX conversion feasibility bench.

use anyhow::{Context, Result};
use calamine::Reader as _;
use serde::Serialize;
use sha2::{Digest, Sha256};
use std::io::Cursor;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};
use std::time::{Instant, SystemTime, UNIX_EPOCH};

pub const TOOL_VERSION: &str = env!("CARGO_PKG_VERSION");

/// Conversion engine under test.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum Engine {
    /// A: Excel COM, one instance per CLI run (batch reuse).
    ComBatch,
    /// A': Excel COM, one Excel instance per file.
    ComPerFile,
    /// B: rXLS direct conversion.
    Rxls,
    /// C-values: calamine + rust_xlsxwriter, values only.
    CalamineValues,
    /// C-formulas: calamine + rust_xlsxwriter, values + formulas (no styles).
    CalamineFormulas,
}

impl Engine {
    pub fn id(self) -> &'static str {
        match self {
            Engine::ComBatch => "com-batch",
            Engine::ComPerFile => "com-perfile",
            Engine::Rxls => "rxls",
            Engine::CalamineValues => "calamine-values",
            Engine::CalamineFormulas => "calamine-formulas",
        }
    }

    pub fn parse(s: &str) -> Option<Self> {
        match s {
            "com-batch" | "excel_com" | "com" => Some(Engine::ComBatch),
            "com-perfile" => Some(Engine::ComPerFile),
            "rxls" => Some(Engine::Rxls),
            "calamine-values" | "calamine" => Some(Engine::CalamineValues),
            "calamine-formulas" => Some(Engine::CalamineFormulas),
            _ => None,
        }
    }

    /// True when the engine drives Excel COM (Windows-only, external process).
    pub fn is_com(self) -> bool {
        matches!(self, Engine::ComBatch | Engine::ComPerFile)
    }
}

/// Output validation depth applied before an output is committed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Validate {
    None,
    /// ZIP magic + trailing-size sanity only (used inside timed benchmarks).
    Quick,
    /// Quick + full re-parse by rxls and calamine (independent readers).
    Full,
}

impl Validate {
    pub fn parse(s: &str) -> Option<Self> {
        match s {
            "none" => Some(Validate::None),
            "quick" => Some(Validate::Quick),
            "full" => Some(Validate::Full),
            _ => None,
        }
    }
}

/// Per-file stage timings (milliseconds). Fields stay `None` when a stage does
/// not apply to the backend (e.g. COM has no parse/serialize split).
#[derive(Debug, Default, Serialize)]
pub struct FileTiming {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub read_ms: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub parse_ms: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub serialize_ms: Option<f64>,
    /// Output re-parse validation; deliberately outside `total_ms`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub validate_ms: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub write_ms: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub com_open_ms: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub com_save_as_ms: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub com_close_ms: Option<f64>,
    /// End-to-end wall time for this file in this backend's accounting.
    pub total_ms: f64,
}

/// Per-file conversion result.
#[derive(Debug, Default, Serialize)]
pub struct FileOutcome {
    pub input: String,
    pub output: String,
    /// `ok` | `failed` | `skipped`
    pub status: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error: Option<String>,
    /// rXLS only: `to_xlsx_checked` rejected the workbook and the sanitized
    /// `to_xlsx` path produced the committed bytes.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub checked_error: Option<String>,
    pub timing: FileTiming,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub input_bytes: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub output_bytes: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub input_sha256: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub output_sha256: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rxls_diag: Option<serde_json::Value>,
}

impl FileOutcome {
    pub fn new(input: &Path, output: &Path) -> Self {
        Self {
            input: input.to_string_lossy().into_owned(),
            output: output.to_string_lossy().into_owned(),
            status: "failed".to_owned(),
            ..Default::default()
        }
    }

    pub fn hash_input(&mut self, bytes: &[u8]) {
        self.input_bytes = Some(bytes.len() as u64);
        self.input_sha256 = Some(sha256_hex(bytes));
    }

    pub fn hash_output(&mut self, bytes: &[u8]) {
        self.output_bytes = Some(bytes.len() as u64);
        self.output_sha256 = Some(sha256_hex(bytes));
    }
}

/// Excel instance metadata (COM backends only).
#[derive(Debug, Default, Serialize)]
pub struct ComMeta {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub excel_pid: Option<u32>,
    pub excel_instances: usize,
    pub sta_init_ms: f64,
    pub excel_create_config_ms: f64,
    pub excel_quit_ms: f64,
    pub sta_release_ms: f64,
    pub excel_exit_confirmed: bool,
    /// Peak working set (bytes) observed on the owned Excel process.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub excel_peak_ws_bytes: Option<u64>,
    /// Kernel+user CPU time (ms) consumed by the owned Excel process(es).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub excel_cpu_ms: Option<f64>,
}

/// Worker process resource usage (in-process backends).
#[derive(Debug, Default, Serialize)]
pub struct MemMeta {
    pub worker_peak_ws_bytes: u64,
    pub worker_cpu_ms: f64,
}

/// One conversion run (a single CLI `convert` or one benchmark round).
#[derive(Debug, Serialize)]
pub struct BatchSummary {
    pub tool_version: String,
    pub engine: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub mode: Option<String>,
    pub started_at_epoch_ms: u64,
    pub finished_at_epoch_ms: u64,
    pub ok_count: usize,
    pub failed_count: usize,
    pub skipped_count: usize,
    /// Wall time of the conversion loop itself (ms).
    pub total_ms: f64,
    /// Process/session setup time before the first conversion (ms).
    pub init_ms: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub com: Option<ComMeta>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub memory: Option<MemMeta>,
    pub input_total_bytes: u64,
    pub output_total_bytes: u64,
    pub files: Vec<FileOutcome>,
}

impl BatchSummary {
    pub fn new(engine: Engine, mode: Option<String>) -> Self {
        Self {
            tool_version: TOOL_VERSION.to_owned(),
            engine: engine.id().to_owned(),
            mode,
            started_at_epoch_ms: now_epoch_ms(),
            finished_at_epoch_ms: 0,
            ok_count: 0,
            failed_count: 0,
            skipped_count: 0,
            total_ms: 0.0,
            init_ms: 0.0,
            com: None,
            memory: None,
            input_total_bytes: 0,
            output_total_bytes: 0,
            files: Vec::new(),
        }
    }

    pub fn finalize(&mut self) {
        self.finished_at_epoch_ms = now_epoch_ms();
        for f in &self.files {
            match f.status.as_str() {
                "ok" => self.ok_count += 1,
                "skipped" => self.skipped_count += 1,
                _ => self.failed_count += 1,
            }
            self.input_total_bytes += f.input_bytes.unwrap_or(0);
            self.output_total_bytes += f.output_bytes.unwrap_or(0);
        }
    }
}

pub fn now_epoch_ms() -> u64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_millis() as u64)
        .unwrap_or(0)
}

pub fn elapsed_ms(start: Instant) -> f64 {
    start.elapsed().as_secs_f64() * 1000.0
}

pub fn sha256_hex(data: &[u8]) -> String {
    let mut hasher = Sha256::new();
    hasher.update(data);
    let digest = hasher.finalize();
    let mut out = String::with_capacity(digest.len() * 2);
    for byte in digest {
        out.push_str(&format!("{byte:02x}"));
    }
    out
}

/// Expand the CLI inputs into a sorted, de-duplicated list of `.xls` files.
/// Directories are walked (optionally recursive); non-`.xls` files are ignored
/// with a stderr note.
pub fn list_xls_files(inputs: &[PathBuf], recursive: bool) -> Result<Vec<PathBuf>> {
    let mut out: Vec<PathBuf> = Vec::new();
    let mut noted: Vec<PathBuf> = Vec::new();
    fn push_file(path: &Path, out: &mut Vec<PathBuf>, noted: &mut Vec<PathBuf>) {
        let is_xls = path
            .extension()
            .map(|e| e.eq_ignore_ascii_case("xls"))
            .unwrap_or(false);
        if is_xls {
            out.push(path.to_path_buf());
        } else {
            noted.push(path.to_path_buf());
        }
    }
    fn walk_dir(dir: &Path, recursive: bool, out: &mut Vec<PathBuf>, noted: &mut Vec<PathBuf>) {
        let mut entries: Vec<PathBuf> = match std::fs::read_dir(dir) {
            Ok(rd) => rd
                .filter_map(|e| e.ok())
                .map(|e| e.path())
                .collect::<Vec<_>>(),
            Err(_) => return,
        };
        entries.sort();
        for entry in entries {
            if entry.is_dir() {
                if recursive {
                    walk_dir(&entry, recursive, out, noted);
                }
            } else if entry.is_file() {
                push_file(&entry, out, noted);
            }
        }
    }
    for input in inputs {
        if input.is_dir() {
            walk_dir(input, recursive, &mut out, &mut noted);
        } else if input.is_file() {
            push_file(input, &mut out, &mut noted);
        } else {
            anyhow::bail!("input does not exist: {}", input.display());
        }
    }
    for skipped in &noted {
        eprintln!("note: ignored non-.xls input: {}", skipped.display());
    }
    out.sort();
    out.dedup();
    Ok(out)
}

static TMP_COUNTER: AtomicU64 = AtomicU64::new(0);

/// Write converted bytes to a temp file in `out_dir`, validate, then atomically
/// rename to `<stem>.xlsx`. Never overwrites an existing output.
///
/// Returns `(final_path, write_ms, validate_ms)`.
pub fn write_and_commit(
    out_dir: &Path,
    stem: &str,
    bytes: &[u8],
    level: Validate,
) -> Result<(PathBuf, f64, f64), String> {
    let final_path = out_dir.join(format!("{stem}.xlsx"));
    if final_path.exists() {
        return Err(format!("output_exists: {}", final_path.display()));
    }
    let unique = TMP_COUNTER.fetch_add(1, Ordering::Relaxed);
    let tmp = out_dir.join(format!(
        ".{stem}.xlsx.tmp-{}-{unique}",
        std::process::id()
    ));
    let write_start = Instant::now();
    if let Err(e) = std::fs::write(&tmp, bytes) {
        return Err(format!("write temp failed: {e}"));
    }
    let write_ms = elapsed_ms(write_start);
    let validate_start = Instant::now();
    let check = validate_xlsx_bytes(bytes, level);
    let validate_ms = elapsed_ms(validate_start);
    if let Err(reason) = check {
        let _ = std::fs::remove_file(&tmp);
        return Err(format!("validation failed: {reason}"));
    }
    match std::fs::rename(&tmp, &final_path) {
        Ok(()) => Ok((final_path, write_ms, validate_ms)),
        Err(e) => {
            let _ = std::fs::remove_file(&tmp);
            Err(format!(
                "commit rename failed ({e}); temp removed, no partial output left"
            ))
        }
    }
}

/// Structural validation of candidate `.xlsx` bytes.
pub fn validate_xlsx_bytes(bytes: &[u8], level: Validate) -> Result<(), String> {
    if level == Validate::None {
        return Ok(());
    }
    if bytes.len() < 4 || &bytes[..4] != b"PK\x03\x04" {
        return Err(format!(
            "not a ZIP/XLSX container ({} bytes, bad magic)",
            bytes.len()
        ));
    }
    if level == Validate::Quick {
        return Ok(());
    }
    // Independent re-parse #1: rxls.
    rxls::Workbook::open(bytes).map_err(|e| format!("rxls re-parse failed: {e}"))?;
    // Independent re-parse #2: calamine (owned buffer keeps the reader 'static).
    let cursor = Cursor::new(bytes.to_vec());
    calamine::Xlsx::<Cursor<Vec<u8>>>::new(cursor)
        .map_err(|e| format!("calamine re-parse failed: {e:?}"))?;
    Ok(())
}

/// Fresh temp output directory under a parent, removed on drop.
pub struct TempDir(pub PathBuf);

impl TempDir {
    pub fn create(parent: &Path, prefix: &str) -> Result<Self> {
        std::fs::create_dir_all(parent).context("create parent dir")?;
        for i in 0..10_000u32 {
            let candidate = parent.join(format!(
                "{prefix}-{}-{i}",
                std::process::id()
            ));
            if !candidate.exists() {
                std::fs::create_dir_all(&candidate).context("create temp dir")?;
                return Ok(TempDir(candidate));
            }
        }
        anyhow::bail!("could not create temp dir under {}", parent.display());
    }
}

impl Drop for TempDir {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.0);
    }
}
