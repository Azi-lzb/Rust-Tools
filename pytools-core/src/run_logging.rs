//! Task-level logging primitives shared by interactive business operations.
//!
//! A file's final state is separate from shared output publication.  The
//! tracker never infers success from a successful read or from a log message.
use chrono::Local;
use serde::Serialize;
use std::{
    fs::{self, File, OpenOptions},
    io::{self, Read, Seek, SeekFrom, Write},
    path::{Path, PathBuf},
    sync::atomic::{AtomicU64, Ordering},
    time::{Instant, SystemTime, UNIX_EPOCH},
};

static RUN_SEQUENCE: AtomicU64 = AtomicU64::new(0);

/// Extract Windows or Unix names even when a log is viewed on another OS.
pub fn portable_file_name(path: &Path) -> String {
    let full = path.as_os_str().to_string_lossy();
    full.rsplit(['/', '\\']).next().unwrap_or(&full).to_owned()
}

/// Return full paths only for the explicitly selected developer log view.
pub fn path_for_log_view(path: &Path, developer_view: bool) -> String {
    if developer_view {
        path.to_string_lossy().into_owned()
    } else {
        portable_file_name(path)
    }
}

/// Keep user-view diagnostic references resolvable from the configured log
/// directory without persisting the machine-specific absolute path.
pub fn diagnostic_path_for_log_view(path: &Path, developer_view: bool) -> String {
    if developer_view {
        path.to_string_lossy().into_owned()
    } else {
        Path::new("run-events")
            .join(portable_file_name(path))
            .to_string_lossy()
            .into_owned()
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
#[serde(rename_all = "SCREAMING_SNAKE_CASE")]
pub enum Level {
    Debug,
    Info,
    Warn,
    Error,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
#[serde(rename_all = "SCREAMING_SNAKE_CASE")]
pub enum TaskStatus {
    Success,
    PartialSuccess,
    Failed,
    Cancelled,
    NoOp,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
#[serde(rename_all = "snake_case")]
pub enum FileStatus {
    Pending,
    InProgress,
    Succeeded,
    Failed,
    Skipped,
    Cancelled,
}

#[derive(Clone, Debug, Serialize)]
pub struct LogEvent {
    pub schema_version: u8,
    pub timestamp: String,
    pub level: Level,
    pub event_name: String,
    pub run_id: String,
    pub feature_id: String,
    pub stage: String,
    pub status: String,
    pub message: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub file_index: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub file_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub file_path: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sheet_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cell_address: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub config_row: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_code: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub engine: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub elapsed_ms: Option<u128>,
    /// Developer-only diagnostic context (for example an anyhow error chain).
    /// Business data must never be placed here.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub details: Option<serde_json::Value>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize)]
pub struct InputStats {
    pub selected: usize,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub discovered: Option<usize>,
    pub attempted: usize,
    pub succeeded: usize,
    pub failed: usize,
    pub skipped: usize,
    pub cancelled: usize,
    pub pending: usize,
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize)]
pub struct SheetStats {
    pub discovered: usize,
    pub attempted: usize,
    pub succeeded: usize,
    pub failed: usize,
    pub skipped: usize,
}

#[derive(Clone, Debug, Default, Serialize)]
pub struct RecordStats {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rows_processed: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rows_written: Option<usize>,
}

#[derive(Clone, Debug, Default, Serialize)]
pub struct OutputStats {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub workbooks_saved: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sheets_saved: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub saved: Option<bool>,
}

#[derive(Clone, Debug, Default, Serialize)]
pub struct PerformanceStats {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub read_elapsed_ms: Option<u128>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub write_elapsed_ms: Option<u128>,
}

#[derive(Clone, Debug, Serialize)]
pub struct ErrorSummary {
    pub stage: String,
    pub code: String,
    pub message: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub file_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sheet_name: Option<String>,
}

#[derive(Clone, Debug, Serialize)]
pub struct InputResult {
    pub index: usize,
    pub file_name: String,
    pub status: FileStatus,
    pub attempted: bool,
}

#[derive(Clone, Debug, Serialize)]
pub struct RunSummary {
    pub schema_version: u8,
    pub run_id: String,
    pub feature_id: String,
    pub feature_name: String,
    pub started_at: String,
    pub ended_at: String,
    pub total_elapsed_ms: u128,
    pub overall_status: TaskStatus,
    pub input_stats: InputStats,
    /// False when an entry point cannot yet account for every selected input.
    pub input_stats_available: bool,
    pub inputs: Vec<InputResult>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sheet_stats: Option<SheetStats>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub record_stats: Option<RecordStats>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub output_stats: Option<OutputStats>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub performance_stats: Option<PerformanceStats>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_summary: Option<ErrorSummary>,
    pub output_paths: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub diagnostic_log_path: Option<String>,
}

/// Each selected input is identified by its position, including repeated paths.
pub struct RunTracker {
    pub run_id: String,
    pub feature_id: String,
    pub feature_name: String,
    pub started_at: String,
    started: Instant,
    inputs: Vec<InputResult>,
    discovered: Option<usize>,
}

impl RunTracker {
    pub fn new(feature_id: &str, feature_name: &str, files: &[PathBuf]) -> Self {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default()
            .as_nanos();
        let sequence = RUN_SEQUENCE.fetch_add(1, Ordering::Relaxed);
        let run_id = format!(
            "pytools-{}-{}-{nonce}-{sequence}",
            feature_id.to_ascii_lowercase(),
            std::process::id()
        );
        Self::with_id(feature_id, feature_name, files, run_id)
    }

    /// Reuse a pre-existing run identifier while migrating a legacy entry point.
    pub fn with_id(
        feature_id: &str,
        feature_name: &str,
        files: &[PathBuf],
        run_id: String,
    ) -> Self {
        Self {
            run_id,
            feature_id: feature_id.to_owned(),
            feature_name: feature_name.to_owned(),
            started_at: Local::now().to_rfc3339(),
            started: Instant::now(),
            discovered: None,
            inputs: files
                .iter()
                .enumerate()
                .map(|(index, path)| InputResult {
                    index: index + 1,
                    file_name: portable_file_name(path),
                    status: FileStatus::Pending,
                    attempted: false,
                })
                .collect(),
        }
    }

    pub fn set_discovered(&mut self, count: usize) {
        self.discovered = Some(count);
    }

    /// Adopt an existing task's monotonic elapsed duration while adapting a
    /// legacy report produced by another module.
    pub fn set_existing_timing(&mut self, started_at: String, elapsed_ms: u128) {
        self.started_at = started_at;
        let now = Instant::now();
        let millis = u64::try_from(elapsed_ms).unwrap_or(u64::MAX);
        self.started = now.checked_sub(std::time::Duration::from_millis(millis)).unwrap_or(now);
    }

    /// Install a discovered file list once without resetting the task clock.
    pub fn set_inputs(&mut self, files: &[PathBuf]) -> Result<(), &'static str> {
        if !self.inputs.is_empty() { return Err("inputs already initialized"); }
        self.inputs = files.iter().enumerate().map(|(index, path)| InputResult {
            index: index + 1,
            file_name: portable_file_name(path),
            status: FileStatus::Pending,
            attempted: false,
        }).collect();
        Ok(())
    }

    pub fn state(&self, index: usize) -> Option<FileStatus> {
        index
            .checked_sub(1)
            .and_then(|n| self.inputs.get(n))
            .map(|item| item.status)
    }

    pub fn input_result(&self, index: usize) -> Option<&InputResult> {
        index.checked_sub(1).and_then(|n| self.inputs.get(n))
    }

    pub fn skip_unfinished(&mut self) {
        for item in &mut self.inputs {
            if matches!(item.status, FileStatus::Pending | FileStatus::InProgress) {
                item.status = FileStatus::Skipped;
            }
        }
    }

    /// Reject duplicate finalization or contradictory transitions.
    pub fn transition(&mut self, index: usize, next: FileStatus) -> Result<(), &'static str> {
        let item = index
            .checked_sub(1)
            .and_then(|n| self.inputs.get_mut(n))
            .ok_or("invalid input index")?;
        match (item.status, next) {
            (FileStatus::Pending, FileStatus::InProgress | FileStatus::Failed) => {
                item.attempted = true
            }
            (FileStatus::Pending, FileStatus::Skipped | FileStatus::Cancelled) => {}
            (
                FileStatus::InProgress,
                FileStatus::Succeeded
                | FileStatus::Failed
                | FileStatus::Skipped
                | FileStatus::Cancelled,
            ) => {}
            _ => return Err("invalid input state transition"),
        }
        item.status = next;
        Ok(())
    }

    pub fn input_stats(&self) -> InputStats {
        let mut stats = InputStats {
            selected: self.inputs.len(),
            discovered: self.discovered,
            ..InputStats::default()
        };
        for file in &self.inputs {
            stats.attempted += usize::from(file.attempted);
            match file.status {
                FileStatus::Succeeded => stats.succeeded += 1,
                FileStatus::Failed => stats.failed += 1,
                FileStatus::Skipped => stats.skipped += 1,
                FileStatus::Cancelled => stats.cancelled += 1,
                FileStatus::Pending | FileStatus::InProgress => stats.pending += 1,
            }
        }
        stats
    }

    pub fn event(
        &self,
        level: Level,
        name: &str,
        stage: &str,
        status: &str,
        message: impl Into<String>,
    ) -> LogEvent {
        LogEvent {
            schema_version: 1,
            timestamp: Local::now().to_rfc3339(),
            level,
            event_name: name.to_owned(),
            run_id: self.run_id.clone(),
            feature_id: self.feature_id.clone(),
            stage: stage.to_owned(),
            status: status.to_owned(),
            message: message.into(),
            file_index: None,
            file_name: None,
            file_path: None,
            sheet_name: None,
            cell_address: None,
            config_row: None,
            error_code: None,
            engine: None,
            elapsed_ms: None,
            details: None,
        }
    }

    /// `critical_failure` represents shared output or configuration failure.
    pub fn finish(
        &self,
        critical_failure: bool,
        cancelled: bool,
        no_op: bool,
        output_paths: Vec<String>,
        diagnostic_log_path: Option<String>,
    ) -> RunSummary {
        let input_stats = self.input_stats();
        let overall_status = if critical_failure {
            TaskStatus::Failed
        } else if cancelled || input_stats.cancelled > 0 {
            TaskStatus::Cancelled
        } else if input_stats.failed > 0 && input_stats.succeeded > 0 {
            TaskStatus::PartialSuccess
        } else if input_stats.failed > 0 || input_stats.pending > 0 {
            TaskStatus::Failed
        } else if no_op {
            TaskStatus::NoOp
        } else {
            TaskStatus::Success
        };
        RunSummary {
            schema_version: 1,
            run_id: self.run_id.clone(),
            feature_id: self.feature_id.clone(),
            feature_name: self.feature_name.clone(),
            started_at: self.started_at.clone(),
            ended_at: Local::now().to_rfc3339(),
            total_elapsed_ms: self.started.elapsed().as_millis(),
            overall_status,
            input_stats,
            input_stats_available: true,
            inputs: self.inputs.clone(),
            sheet_stats: None,
            record_stats: None,
            output_stats: None,
            performance_stats: None,
            error_summary: None,
            output_paths,
            diagnostic_log_path,
        }
    }
}

/// Append each event immediately, retaining earlier diagnostics if a task fails.
pub struct EventWriter {
    path: PathBuf,
    file: File,
}
impl EventWriter {
    pub fn create(dir: &Path, feature_id: &str, run_id: &str) -> io::Result<Self> {
        if feature_id.is_empty()
            || run_id.is_empty()
            || !feature_id
                .bytes()
                .all(|b| b.is_ascii_alphanumeric() || b == b'-' || b == b'_')
            || !run_id
                .bytes()
                .all(|b| b.is_ascii_alphanumeric() || b == b'-' || b == b'_')
        {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "invalid run log identifier",
            ));
        }
        fs::create_dir_all(dir)?;
        let path = dir.join(format!("{feature_id}_{run_id}.jsonl"));
        let file = OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&path)?;
        Ok(Self { path, file })
    }
    pub fn path(&self) -> &Path {
        &self.path
    }
    pub fn write(&mut self, event: &LogEvent) -> io::Result<()> {
        serde_json::to_writer(&mut self.file, event).map_err(io::Error::other)?;
        self.file.write_all(b"\n")
    }
}

/// Retention is limited to complete, self-identifying JSONL runs.  In-progress
/// and unrecognized files are never deleted; no directories are traversed.
pub fn retain_completed_events(
    dir: &Path,
    days: Option<u64>,
    max_bytes: Option<u64>,
) -> io::Result<usize> {
    if days.is_none() && max_bytes.is_none() {
        return Ok(0);
    }
    let metadata = match fs::symlink_metadata(dir) {
        Ok(metadata) if metadata.is_dir() && !metadata.file_type().is_symlink() => metadata,
        Ok(_) => {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "diagnostic directory is not a real directory",
            ))
        }
        Err(error) if error.kind() == io::ErrorKind::NotFound => return Ok(0),
        Err(error) => return Err(error),
    };
    let _ = metadata;
    let mut files = Vec::new();
    for entry in fs::read_dir(dir)? {
        let entry = entry?;
        let kind = entry.file_type()?;
        if !kind.is_file() || kind.is_symlink() {
            continue;
        }
        let path = entry.path();
        if path.extension().and_then(|value| value.to_str()) != Some("jsonl") {
            continue;
        }
        let mut file = File::open(&path)?;
        let metadata = file.metadata()?;
        let length = metadata.len();
        file.seek(SeekFrom::Start(length.saturating_sub(65536)))?;
        let mut tail = Vec::new();
        file.read_to_end(&mut tail)?;
        let Some(line) = tail
            .split(|byte| *byte == b'\n')
            .rev()
            .find(|line| !line.is_empty())
        else {
            continue;
        };
        let Ok(event) = serde_json::from_slice::<serde_json::Value>(line) else {
            continue;
        };
        let (Some(feature), Some(run_id)) =
            (event["feature_id"].as_str(), event["run_id"].as_str())
        else {
            continue;
        };
        if event["schema_version"] != 1
            || event["event_name"] != "run_finished"
            || path.file_name().and_then(|name| name.to_str())
                != Some(format!("{feature}_{run_id}.jsonl").as_str())
        {
            continue;
        }
        files.push((path, metadata.modified()?, length));
    }
    files.sort_by_key(|(_, modified, _)| *modified);
    let now = SystemTime::now();
    let mut total = files
        .iter()
        .fold(0u64, |sum, (_, _, size)| sum.saturating_add(*size));
    let mut removed = 0;
    for (path, modified, size) in files {
        let expired = days
            .and_then(|value| value.checked_mul(86_400))
            .is_some_and(|seconds| {
                now.duration_since(modified)
                    .is_ok_and(|age| age.as_secs() > seconds)
            });
        let over_limit = max_bytes.is_some_and(|limit| total > limit);
        if expired || over_limit {
            // Recheck type; never follow a symbolic link selected during scan.
            let current = fs::symlink_metadata(&path)?;
            if current.is_file() && !current.file_type().is_symlink() {
                fs::remove_file(path)?;
                total = total.saturating_sub(size);
                removed += 1;
            }
        }
    }
    Ok(removed)
}

pub fn terminal_line(level: Level, message: &str) -> String {
    let label = match level {
        Level::Debug => "DEBUG",
        Level::Info => "INFO",
        Level::Warn => "WARN",
        Level::Error => "ERROR",
    };
    format!(
        "[{}] [{label}] {}",
        Local::now().format("%Y-%m-%d %H:%M:%S"),
        message
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    fn inputs() -> Vec<PathBuf> {
        ["机构甲.xlsx", "机构乙.xlsx", "机构丙.xlsx"]
            .into_iter()
            .map(PathBuf::from)
            .collect()
    }

    #[test]
    fn event_identity_and_dates_are_correlated() {
        let first = RunTracker::new("1-2", "使用区域汇总", &inputs());
        let second = RunTracker::new("1-2", "使用区域汇总", &inputs());
        assert_ne!(first.run_id, second.run_id);
        let event = first.event(Level::Info, "run_started", "config", "started", "开始");
        assert_eq!(event.run_id, first.run_id);
        assert_eq!(event.feature_id, "1-2");
        chrono::DateTime::parse_from_rfc3339(&event.timestamp).unwrap();
        let summary = first.finish(false, true, false, vec![], None);
        chrono::DateTime::parse_from_rfc3339(&summary.started_at).unwrap();
        chrono::DateTime::parse_from_rfc3339(&summary.ended_at).unwrap();
        assert_eq!(summary.run_id, event.run_id);
    }

    #[test]
    fn states_are_exclusive_and_reading_is_not_success() {
        let mut run = RunTracker::new("1-2", "汇总", &inputs());
        run.transition(1, FileStatus::InProgress).unwrap();
        assert_eq!(run.input_stats().succeeded, 0);
        run.transition(1, FileStatus::Succeeded).unwrap();
        run.transition(2, FileStatus::InProgress).unwrap();
        run.transition(2, FileStatus::Failed).unwrap();
        run.transition(3, FileStatus::Skipped).unwrap();
        assert!(run.transition(2, FileStatus::Succeeded).is_err());
        let stats = run.input_stats();
        assert_eq!(
            (
                stats.selected,
                stats.attempted,
                stats.succeeded,
                stats.failed,
                stats.skipped,
                stats.pending
            ),
            (3, 2, 1, 1, 1, 0)
        );
        assert_eq!(
            run.finish(false, false, false, vec![], None).overall_status,
            TaskStatus::PartialSuccess
        );

        let mut immediate_failure = RunTracker::new("1-2", "汇总", &inputs()[..1]);
        immediate_failure.transition(1, FileStatus::Failed).unwrap();
        assert_eq!(immediate_failure.input_stats().attempted, 1);
        assert_eq!(immediate_failure.input_stats().failed, 1);
    }

    #[test]
    fn shared_output_failure_overrides_successful_input() {
        let mut run = RunTracker::new("1-2", "汇总", &inputs()[..1]);
        run.transition(1, FileStatus::InProgress).unwrap();
        run.transition(1, FileStatus::Succeeded).unwrap();
        let summary = run.finish(true, false, false, vec![], None);
        assert_eq!(summary.input_stats.succeeded, 1);
        assert_eq!(summary.overall_status, TaskStatus::Failed);
    }

    #[test]
    fn output_stats_do_not_turn_unknown_values_into_zero_or_false() {
        let value = serde_json::to_value(OutputStats::default()).unwrap();
        assert_eq!(value, serde_json::json!({}));

        let failed_output = OutputStats {
            workbooks_saved: Some(0),
            sheets_saved: None,
            saved: Some(false),
        };
        let value = serde_json::to_value(failed_output).unwrap();
        assert_eq!(value["workbooks_saved"], 0);
        assert_eq!(value["saved"], false);
        assert!(value.get("sheets_saved").is_none());
    }

    #[test]
    fn user_diagnostic_path_is_relative_and_developer_path_keeps_location() {
        let path = Path::new(r"C:\private\logs\run-events\PT-1-1_example.jsonl");
        assert_eq!(
            diagnostic_path_for_log_view(path, false),
            Path::new("run-events")
                .join("PT-1-1_example.jsonl")
                .to_string_lossy()
        );
        assert_eq!(diagnostic_path_for_log_view(path, true), path.to_string_lossy());
    }

    #[test]
    fn no_op_cancellation_and_pending_are_distinct() {
        let mut run = RunTracker::new("1-2", "汇总", &inputs());
        run.transition(1, FileStatus::Skipped).unwrap();
        run.transition(2, FileStatus::Cancelled).unwrap();
        assert_eq!(
            run.finish(false, true, false, vec![], None).overall_status,
            TaskStatus::Cancelled
        );
        assert_eq!(run.input_stats().pending, 1);
        let mut empty = RunTracker::new("1-2", "汇总", &inputs());
        for index in 1..=3 {
            empty.transition(index, FileStatus::Skipped).unwrap();
        }
        assert_eq!(
            empty
                .finish(false, false, true, vec![], None)
                .overall_status,
            TaskStatus::NoOp
        );
    }

    #[test]
    fn jsonl_preserves_unicode_and_optional_location() -> io::Result<()> {
        let root = std::env::temp_dir().join(format!(
            "pytools-log-test-{}-{}",
            std::process::id(),
            RUN_SEQUENCE.fetch_add(1, Ordering::Relaxed)
        ));
        let run = RunTracker::new("1-2", "汇总", &inputs());
        let mut writer = EventWriter::create(&root, &run.feature_id, &run.run_id)?;
        let mut event = run.event(
            Level::Error,
            "sheet_failed",
            "read_sheet",
            "failed",
            "无法读取工作表",
        );
        event.file_name = Some("机构乙.xlsx".into());
        event.sheet_name = Some("贷款表".into());
        event.cell_address = Some("E42".into());
        event.error_code = Some("SHEET_READ_FAILED".into());
        event.details = Some(serde_json::json!({"error_chain": "reader failed: invalid workbook"}));
        writer.write(&event)?;
        let path = writer.path().to_path_buf();
        drop(writer);
        let content = fs::read_to_string(&path)?;
        let decoded: serde_json::Value =
            serde_json::from_str(content.trim()).map_err(io::Error::other)?;
        assert_eq!(decoded["sheet_name"], "贷款表");
        assert_eq!(decoded["cell_address"], "E42");
        assert_eq!(decoded["details"]["error_chain"], "reader failed: invalid workbook");
        assert!(decoded.get("file_path").is_none());
        fs::remove_file(path)?;
        fs::remove_dir(root)?;
        Ok(())
    }

    #[test]
    fn invalid_log_path_fails_without_changing_inputs() {
        let run = RunTracker::new("1-2", "汇总", &inputs());
        assert!(EventWriter::create(Path::new("."), "../bad", &run.run_id).is_err());
        assert_eq!(run.input_stats().selected, 3);
        assert_eq!(run.input_stats().pending, 3);
    }

    #[test]
    fn windows_extended_paths_do_not_leak_into_file_names() {
        let path = PathBuf::from(r"\\?\C:\业务\机构甲-长期中文文件名称.xlsx");
        let run = RunTracker::new("1-2", "汇总", &[path]);
        let summary = run.finish(false, true, false, vec![], None);
        assert_eq!(summary.inputs[0].file_name, "机构甲-长期中文文件名称.xlsx");
    }

    #[test]
    fn retention_preserves_unfinished_and_unrelated_files() -> io::Result<()> {
        let root = std::env::temp_dir().join(format!(
            "pytools-retention-test-{}-{}",
            std::process::id(),
            RUN_SEQUENCE.fetch_add(1, Ordering::Relaxed)
        ));
        let finished = RunTracker::new("1-2", "汇总", &[]);
        let active = RunTracker::new("1-2", "汇总", &[]);
        let mut done_writer = EventWriter::create(&root, &finished.feature_id, &finished.run_id)?;
        done_writer.write(&finished.event(
            Level::Info,
            "run_finished",
            "run",
            "success",
            "结束",
        ))?;
        let done_path = done_writer.path().to_owned();
        drop(done_writer);
        let mut active_writer = EventWriter::create(&root, &active.feature_id, &active.run_id)?;
        active_writer.write(&active.event(Level::Info, "run_started", "run", "started", "开始"))?;
        let active_path = active_writer.path().to_owned();
        let unrelated = root.join("business.xlsx");
        fs::write(&unrelated, b"business data")?;
        assert_eq!(retain_completed_events(&root, None, Some(1))?, 1);
        assert!(!done_path.exists());
        assert!(active_path.exists());
        assert_eq!(fs::read(&unrelated)?, b"business data");
        drop(active_writer);
        fs::remove_file(active_path)?;
        fs::remove_file(unrelated)?;
        fs::remove_dir(root)
    }
}
