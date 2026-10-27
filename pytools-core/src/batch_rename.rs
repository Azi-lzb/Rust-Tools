//! Rust implementation of PyCode's configuration-driven batch file rename (3-3).

use anyhow::{bail, Context, Result};
use chrono::Local;
use excel_reader::{read_workbook, uses_1904_date_system, SheetDump};
use serde::Serialize;
use serde_json::Value;
use std::io;
use std::path::{Path, PathBuf};
use std::time::Instant;

const RENAME_SHEET: &str = "重命名配置";

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct BatchRenameRow {
    pub source_path: String,
    pub original_name: String,
    pub new_name: String,
    pub status: String,
    pub reason: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct BatchRenameResult {
    pub ok: usize,
    pub skip: usize,
    pub rows: Vec<BatchRenameRow>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct BatchRenameProgressEvent {
    pub sequence: usize,
    pub timestamp: String,
    pub feature_id: String,
    pub phase: String,
    pub status: String,
    pub completed_files: usize,
    pub total_files: usize,
    pub renamed_files: usize,
    pub skipped_files: usize,
    pub stage_elapsed_ms: u128,
    pub total_elapsed_ms: u128,
    pub components: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub actual_engine: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_type: Option<String>,
    pub user_message: String,
}

#[derive(Debug, Default)]
struct RenameMaps {
    short_to_full: Vec<(String, String)>,
    full_to_code: Vec<(String, String)>,
    settings: Vec<(String, String)>,
}

/// Rename paths using the original 3-3 configuration contract.
///
/// The configuration workbook is read through the shared Calamine adapter. Each
/// input path is renamed within its parent directory; directories are supported
/// because the source operation accepts both files and folders. Existing names
/// are never intentionally selected as destinations: collisions receive `_1`,
/// `_2`, ... before the original extension.
pub fn run_batch_rename(
    config_path: impl AsRef<Path>,
    paths: &[PathBuf],
) -> Result<BatchRenameResult> {
    run_batch_rename_with_progress(config_path, paths, |_| {})
}

pub fn run_batch_rename_with_progress<F>(
    config_path: impl AsRef<Path>,
    paths: &[PathBuf],
    mut on_event: F,
) -> Result<BatchRenameResult>
where
    F: FnMut(&BatchRenameProgressEvent),
{
    let task_started = Instant::now();
    let configuration_started = Instant::now();
    let mut sequence = 0usize;
    emit_progress(
        &mut on_event,
        &mut sequence,
        &task_started,
        &configuration_started,
        "configuration_read",
        "started",
        0,
        paths.len(),
        0,
        0,
        None,
        None,
        "正在读取批量重命名配置".to_owned(),
    );

    let config_path = config_path.as_ref();
    let configuration = (|| -> Result<(RenameMaps, String, String)> {
        let workbook = read_workbook(config_path)
            .with_context(|| format!("read rename configuration {}", config_path.display()))?;
        let sheet = workbook
            .sheets
            .iter()
            .find(|sheet| sheet.name == RENAME_SHEET)
            .with_context(|| format!("configuration workbook has no {RENAME_SHEET:?} sheet"))?;
        let date_1904 = uses_1904_date_system(config_path)?;
        let maps = load_rename_maps(sheet, date_1904)?;

        let date_text = setting(&maps.settings, "数据日期");
        let report_name = setting(&maps.settings, "报表名称");
        if maps.short_to_full.is_empty()
            || maps.full_to_code.is_empty()
            || date_text.is_empty()
            || report_name.is_empty()
        {
            bail!("重命名配置 缺少必要配置：简称/全称、全称(代码映射)/代码、数据日期、报表名称。");
        }
        Ok((maps, date_text, report_name))
    })();
    let (maps, date_text, report_name) = match configuration {
        Ok(configuration) => configuration,
        Err(error) => {
            emit_progress(
                &mut on_event,
                &mut sequence,
                &task_started,
                &configuration_started,
                "configuration_read",
                "failed",
                0,
                paths.len(),
                0,
                0,
                Some("Calamine"),
                Some("configuration_error"),
                "批量重命名配置读取失败，本次任务已停止".to_owned(),
            );
            return Err(error);
        }
    };
    emit_progress(
        &mut on_event,
        &mut sequence,
        &task_started,
        &configuration_started,
        "configuration_read",
        "completed",
        0,
        paths.len(),
        0,
        0,
        Some("Calamine"),
        None,
        "批量重命名配置读取完成".to_owned(),
    );

    let processing_started = Instant::now();
    let mut result = BatchRenameResult {
        ok: 0,
        skip: 0,
        rows: Vec::with_capacity(paths.len()),
    };
    emit_progress(
        &mut on_event,
        &mut sequence,
        &task_started,
        &processing_started,
        "file_processing",
        "started",
        0,
        paths.len(),
        result.ok,
        result.skip,
        None,
        None,
        format!("开始处理 {} 个文件或目录", paths.len()),
    );
    for (index, path) in paths.iter().enumerate() {
        let original_name = path
            .file_name()
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_default();
        let source_path = path.to_string_lossy().into_owned();
        let row = rename_one(
            path,
            &original_name,
            &source_path,
            &maps,
            &date_text,
            &report_name,
        );
        if row.status == "renamed" {
            result.ok += 1;
        } else {
            result.skip += 1;
        }
        result.rows.push(row);
        emit_progress(
            &mut on_event,
            &mut sequence,
            &task_started,
            &processing_started,
            "file_processing",
            "progress",
            index + 1,
            paths.len(),
            result.ok,
            result.skip,
            Some("RustFilesystem"),
            None,
            format!(
                "已处理 {}/{} 个文件或目录：改名 {} 个，跳过 {} 个",
                index + 1,
                paths.len(),
                result.ok,
                result.skip
            ),
        );
    }

    emit_progress(
        &mut on_event,
        &mut sequence,
        &task_started,
        &processing_started,
        "file_processing",
        "completed",
        paths.len(),
        paths.len(),
        result.ok,
        result.skip,
        (!paths.is_empty()).then_some("RustFilesystem"),
        None,
        format!(
            "文件处理完成：改名 {} 个，跳过 {} 个",
            result.ok, result.skip
        ),
    );
    let task_finished = Instant::now();
    emit_progress(
        &mut on_event,
        &mut sequence,
        &task_started,
        &task_finished,
        "task",
        "completed",
        paths.len(),
        paths.len(),
        result.ok,
        result.skip,
        None,
        None,
        format!(
            "批量重命名完成：改名 {} 个，跳过 {} 个",
            result.ok, result.skip
        ),
    );
    Ok(result)
}

fn emit_progress<F>(
    on_event: &mut F,
    sequence: &mut usize,
    task_started: &Instant,
    phase_started: &Instant,
    phase: &str,
    status: &str,
    completed_files: usize,
    total_files: usize,
    renamed_files: usize,
    skipped_files: usize,
    actual_engine: Option<&str>,
    error_type: Option<&str>,
    user_message: String,
) where
    F: FnMut(&BatchRenameProgressEvent),
{
    *sequence += 1;
    on_event(&BatchRenameProgressEvent {
        sequence: *sequence,
        timestamp: Local::now().to_rfc3339(),
        feature_id: "PT-4-1".to_owned(),
        phase: phase.to_owned(),
        status: status.to_owned(),
        completed_files,
        total_files,
        renamed_files,
        skipped_files,
        stage_elapsed_ms: phase_started.elapsed().as_millis(),
        total_elapsed_ms: task_started.elapsed().as_millis(),
        components: vec!["Rust".to_owned()],
        actual_engine: actual_engine.map(str::to_owned),
        error_type: error_type.map(str::to_owned),
        user_message,
    });
}

fn rename_one(
    path: &Path,
    original_name: &str,
    source_path: &str,
    maps: &RenameMaps,
    date_text: &str,
    report_name: &str,
) -> BatchRenameRow {
    let skip = |reason: String| BatchRenameRow {
        source_path: source_path.to_owned(),
        original_name: original_name.to_owned(),
        new_name: String::new(),
        status: "skipped".to_owned(),
        reason,
    };

    if !path.exists() {
        return skip("not_exists".to_owned());
    }

    let lowered_name = original_name.to_lowercase();
    let Some((_, full_name)) = maps
        .short_to_full
        .iter()
        .find(|(short_name, _)| lowered_name.contains(&short_name.to_lowercase()))
    else {
        return skip("short_name_not_matched".to_owned());
    };
    let Some((_, code)) = maps.full_to_code.iter().find(|(full, _)| full == full_name) else {
        return skip(format!("full_name_no_code full={full_name}"));
    };

    let suffix = path
        .extension()
        .map(|extension| format!(".{}", extension.to_string_lossy()))
        .unwrap_or_default();
    let desired_name = format!("{date_text} {code} {full_name} {report_name}{suffix}");
    let desired_path = path.with_file_name(&desired_name);

    if same_resolved_path(&desired_path, path) {
        return skip("name_same".to_owned());
    }

    match rename_to_available_path(path, &desired_path) {
        Ok(target_path) => BatchRenameRow {
            source_path: source_path.to_owned(),
            original_name: original_name.to_owned(),
            new_name: target_path
                .file_name()
                .map(|name| name.to_string_lossy().into_owned())
                .unwrap_or_default(),
            status: "renamed".to_owned(),
            reason: String::new(),
        },
        Err(error) => skip(format!("rename_failed err={error}")),
    }
}

fn rename_to_available_path(source: &Path, desired_path: &Path) -> io::Result<PathBuf> {
    let parent = desired_path.parent().unwrap_or_else(|| Path::new("."));
    let stem = desired_path
        .file_stem()
        .map(|value| value.to_string_lossy().into_owned())
        .unwrap_or_default();
    let extension = desired_path
        .extension()
        .map(|value| format!(".{}", value.to_string_lossy()))
        .unwrap_or_default();
    for suffix_index in 0usize.. {
        let candidate = if suffix_index == 0 {
            desired_path.to_path_buf()
        } else {
            parent.join(format!("{stem}_{suffix_index}{extension}"))
        };
        match rename_without_replacing(source, &candidate) {
            Ok(()) => return Ok(candidate),
            Err(error) if error.kind() == io::ErrorKind::AlreadyExists => continue,
            Err(error) => return Err(error),
        }
    }
    unreachable!("usize collision suffix range exhausted")
}

#[cfg(windows)]
fn rename_without_replacing(source: &Path, destination: &Path) -> io::Result<()> {
    use std::os::windows::ffi::OsStrExt;
    use windows_sys::Win32::Storage::FileSystem::MoveFileW;

    let source = source
        .as_os_str()
        .encode_wide()
        .chain(std::iter::once(0))
        .collect::<Vec<_>>();
    let destination = destination
        .as_os_str()
        .encode_wide()
        .chain(std::iter::once(0))
        .collect::<Vec<_>>();
    let moved = unsafe { MoveFileW(source.as_ptr(), destination.as_ptr()) };
    if moved == 0 {
        Err(io::Error::last_os_error())
    } else {
        Ok(())
    }
}

#[cfg(target_os = "linux")]
fn rename_without_replacing(source: &Path, destination: &Path) -> io::Result<()> {
    use std::ffi::CString;
    use std::os::raw::c_int;
    use std::os::unix::ffi::OsStrExt;

    const AT_FDCWD: c_int = -100;
    const RENAME_NOREPLACE: u32 = 1;

    let source = CString::new(source.as_os_str().as_bytes())
        .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "source path contains NUL"))?;
    let destination = CString::new(destination.as_os_str().as_bytes()).map_err(|_| {
        io::Error::new(io::ErrorKind::InvalidInput, "destination path contains NUL")
    })?;
    let result = unsafe {
        renameat2(
            AT_FDCWD,
            source.as_ptr(),
            AT_FDCWD,
            destination.as_ptr(),
            RENAME_NOREPLACE,
        )
    };
    if result == 0 {
        Ok(())
    } else {
        Err(io::Error::last_os_error())
    }
}

#[cfg(not(any(windows, target_os = "linux")))]
fn rename_without_replacing(_source: &Path, _destination: &Path) -> io::Result<()> {
    Err(io::Error::new(
        io::ErrorKind::Unsupported,
        "atomic no-replace rename is unavailable on this platform",
    ))
}

#[cfg(target_os = "linux")]
unsafe extern "C" {
    fn renameat2(
        olddirfd: std::os::raw::c_int,
        oldpath: *const std::os::raw::c_char,
        newdirfd: std::os::raw::c_int,
        newpath: *const std::os::raw::c_char,
        flags: u32,
    ) -> std::os::raw::c_int;
}

fn same_resolved_path(left: &Path, right: &Path) -> bool {
    match (std::fs::canonicalize(left), std::fs::canonicalize(right)) {
        (Ok(left), Ok(right)) => left == right,
        _ => left == right,
    }
}

fn load_rename_maps(sheet: &SheetDump, date_1904: bool) -> Result<RenameMaps> {
    let mut columns = std::collections::BTreeMap::<String, usize>::new();
    for (address, value) in &sheet.values {
        let Some((column, row)) = parse_cell_address(address) else {
            continue;
        };
        if row != 1 {
            continue;
        }
        let header = cell_raw_text(Some(value), date_1904);
        if !header.is_empty() {
            // pandas keeps the first duplicate header's meaning and disambiguates
            // later duplicate labels; all required labels are looked up by name.
            columns.entry(header).or_insert(column);
        }
    }

    let max_row = sheet
        .values
        .keys()
        .filter_map(|address| parse_cell_address(address).map(|(_, row)| row))
        .max()
        .unwrap_or(1);
    let mut maps = RenameMaps::default();
    for row in 2..=max_row {
        let get = |header: &str| {
            columns
                .get(header)
                .and_then(|column| sheet.values.get(&cell_address(*column, row)))
                .map(|value| cell_text(Some(value), date_1904))
                .unwrap_or_default()
                .trim()
                .to_owned()
        };
        let short = get("简称");
        let full = get("全称");
        let code = get("代码");
        let code_full = get("全称(代码映射)");
        let key = get("键");
        let value = columns
            .get("值")
            .and_then(|column| sheet.values.get(&cell_address(*column, row)))
            .map(|value| setting_cell_text(&key, value, date_1904))
            .unwrap_or_default()
            .trim()
            .to_owned();
        if !short.is_empty() && !full.is_empty() {
            ordered_set(&mut maps.short_to_full, short, full);
        }
        if !code_full.is_empty() && !code.is_empty() {
            ordered_set(&mut maps.full_to_code, code_full, code);
        }
        if !key.is_empty() && !value.is_empty() {
            ordered_set(&mut maps.settings, key, value);
        }
    }
    Ok(maps)
}

fn ordered_set(entries: &mut Vec<(String, String)>, key: String, value: String) {
    if let Some((_, current)) = entries.iter_mut().find(|(existing, _)| existing == &key) {
        *current = value;
    } else {
        entries.push((key, value));
    }
}

fn setting(settings: &[(String, String)], key: &str) -> String {
    settings
        .iter()
        .find(|(name, _)| name == key)
        .map(|(_, value)| value.clone())
        .unwrap_or_default()
}

fn cell_text(value: Option<&Value>, date_1904: bool) -> String {
    cell_raw_text(value, date_1904).trim().to_owned()
}

fn setting_cell_text(key: &str, value: &Value, date_1904: bool) -> String {
    let text = cell_text(Some(value), date_1904);
    if key != "数据日期" {
        return text;
    }

    match value.get("type").and_then(Value::as_str) {
        // A date setting is a calendar date. Keep Excel's time component out
        // of generated filenames, where ':' is invalid on Windows.
        Some("datetime_serial") => text.split_whitespace().next().unwrap_or(&text).to_owned(),
        Some("datetime_iso") | Some("string") => normalize_filename_date(&text).unwrap_or(text),
        _ => text,
    }
}

fn normalize_filename_date(text: &str) -> Option<String> {
    let bytes = text.as_bytes();
    let has_time = bytes
        .get(10)
        .is_some_and(|separator| matches!(*separator, b'T' | b' '));
    let date = if has_time { text.get(..10)? } else { text };
    let bytes = date.as_bytes();
    if bytes.len() != 10
        || !bytes[..4].iter().all(u8::is_ascii_digit)
        || !bytes[5..7].iter().all(u8::is_ascii_digit)
        || !bytes[8..10].iter().all(u8::is_ascii_digit)
        || !matches!(bytes[4], b'-' | b'/')
        || bytes[7] != bytes[4]
    {
        return None;
    }
    Some(format!("{}-{}-{}", &date[..4], &date[5..7], &date[8..10]))
}

fn cell_raw_text(value: Option<&Value>, date_1904: bool) -> String {
    let Some(value) = value else {
        return String::new();
    };
    let Some(kind) = value.get("type").and_then(Value::as_str) else {
        return value.as_str().unwrap_or_default().to_owned();
    };
    let Some(data) = value.get("value") else {
        return String::new();
    };
    match kind {
        "string" | "datetime_iso" | "duration_iso" | "error" => {
            data.as_str().unwrap_or_default().to_owned()
        }
        "boolean" => data
            .as_bool()
            .map(|value| if value { "True" } else { "False" }.to_owned())
            .unwrap_or_default(),
        "integer" => data
            .as_i64()
            .map(|number| number.to_string())
            .unwrap_or_default(),
        "datetime_serial" => data
            .as_f64()
            .and_then(|serial| {
                const DAYS_1904_TO_1900: f64 = 1462.0;
                let serial = if date_1904 && serial >= 1.0 {
                    serial + DAYS_1904_TO_1900
                } else {
                    serial
                };
                crate::summary_ops::python_datetime_serial_text(serial)
            })
            .unwrap_or_default(),
        "number" => data
            .as_f64()
            .filter(|number| number.is_finite())
            .map(|number| {
                if number.fract() == 0.0 {
                    format!("{number:.1}")
                } else {
                    number.to_string()
                }
            })
            .unwrap_or_default(),
        _ => String::new(),
    }
}

fn parse_cell_address(address: &str) -> Option<(usize, usize)> {
    let split = address.bytes().take_while(u8::is_ascii_alphabetic).count();
    if split == 0 || split == address.len() {
        return None;
    }
    let mut column = 0usize;
    for byte in address[..split].bytes() {
        column = column.checked_mul(26)? + usize::from(byte.to_ascii_uppercase() - b'A' + 1);
    }
    let row = address[split..].parse::<usize>().ok()?;
    (row > 0).then_some((column, row))
}

fn cell_address(mut column: usize, row: usize) -> String {
    let mut letters = String::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        letters.insert(0, char::from(b'A' + remainder as u8));
        column = (column - 1) / 26;
    }
    format!("{letters}{row}")
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;
    use std::time::{SystemTime, UNIX_EPOCH};

    struct TestDirectory(PathBuf);

    impl TestDirectory {
        fn new() -> Self {
            let nonce = SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .expect("clock must be after Unix epoch")
                .as_nanos();
            let path = std::env::temp_dir().join(format!(
                "pytools-batch-rename-noreplace-{}-{nonce}",
                std::process::id()
            ));
            fs::create_dir_all(&path).expect("create isolated test directory");
            Self(path)
        }
    }

    impl Drop for TestDirectory {
        fn drop(&mut self) {
            let _ = fs::remove_dir_all(&self.0);
        }
    }

    #[test]
    fn no_replace_rename_keeps_an_existing_destination_unchanged() {
        let directory = TestDirectory::new();
        let source = directory.0.join("source.xlsx");
        let destination = directory.0.join("destination.xlsx");
        fs::write(&source, b"source bytes").expect("write source");
        fs::write(&destination, b"existing destination").expect("write destination");

        let error = rename_without_replacing(&source, &destination)
            .expect_err("existing destination must not be replaced");

        assert_eq!(error.kind(), io::ErrorKind::AlreadyExists);
        assert_eq!(fs::read(&source).unwrap(), b"source bytes");
        assert_eq!(fs::read(&destination).unwrap(), b"existing destination");
    }

    #[test]
    fn ordered_mapping_updates_value_without_changing_match_priority() {
        let mut entries = Vec::new();
        ordered_set(&mut entries, "建".to_owned(), "第一版".to_owned());
        ordered_set(&mut entries, "建行".to_owned(), "第二项".to_owned());
        ordered_set(&mut entries, "建".to_owned(), "后续覆盖".to_owned());

        assert_eq!(
            entries,
            [
                ("建".to_owned(), "后续覆盖".to_owned()),
                ("建行".to_owned(), "第二项".to_owned())
            ]
        );
        assert_eq!(
            entries
                .iter()
                .find(|(short, _)| "客户_建行.xlsx"
                    .to_lowercase()
                    .contains(&short.to_lowercase()))
                .map(|(_, full)| full.as_str()),
            Some("后续覆盖")
        );
    }

    #[test]
    fn cell_coordinates_round_trip_for_multiletter_columns() {
        for column in [1, 26, 27, 52, 53, 702, 703] {
            let address = cell_address(column, 12);
            assert_eq!(parse_cell_address(&address), Some((column, 12)));
        }
    }

    #[test]
    fn date_serial_config_values_match_python_datetime_text_for_both_epochs() {
        let date_1900 = serde_json::json!({"type": "datetime_serial", "value": 46053.0});
        let date_1904 = serde_json::json!({"type": "datetime_serial", "value": 44591.0});
        let ordinary_number = serde_json::json!({"type": "number", "value": 46053.0});

        assert_eq!(
            cell_raw_text(Some(&date_1900), false),
            "2026-01-31 00:00:00"
        );
        assert_eq!(cell_raw_text(Some(&date_1904), true), "2026-01-31 00:00:00");
        assert_eq!(cell_raw_text(Some(&ordinary_number), false), "46053.0");
    }

    #[test]
    fn date_setting_uses_filename_safe_date_for_excel_date_values() {
        let date_1900 = serde_json::json!({"type": "datetime_serial", "value": 46053.0});
        let date_1904 = serde_json::json!({"type": "datetime_serial", "value": 44591.0});
        let date_iso = serde_json::json!({"type": "datetime_iso", "value": "2026-01-31T00:00:00"});
        let date_text = serde_json::json!({"type": "string", "value": "2026-01-31"});
        let timestamp_text = serde_json::json!({"type": "string", "value": "2026-01-31 00:00:00"});
        let slash_date = serde_json::json!({"type": "string", "value": "2026/01/31"});
        let unrelated_text = serde_json::json!({"type": "string", "value": "月报"});

        assert_eq!(
            setting_cell_text("数据日期", &date_1900, false),
            "2026-01-31"
        );
        assert_eq!(
            setting_cell_text("数据日期", &date_1904, true),
            "2026-01-31"
        );
        assert_eq!(
            setting_cell_text("数据日期", &date_iso, false),
            "2026-01-31"
        );
        assert_eq!(
            setting_cell_text("数据日期", &date_text, false),
            "2026-01-31"
        );
        assert_eq!(
            setting_cell_text("数据日期", &timestamp_text, false),
            "2026-01-31"
        );
        assert_eq!(
            setting_cell_text("数据日期", &slash_date, false),
            "2026-01-31"
        );
        assert_eq!(
            setting_cell_text("报表名称", &unrelated_text, false),
            "月报"
        );
        assert_eq!(
            setting_cell_text("报表名称", &date_1900, false),
            "2026-01-31 00:00:00"
        );
    }
}

/// Compatibility alias for the original Python 3-3 operation identifier.
#[deprecated(note = "use `run_batch_rename`")]
pub fn run_batch_rename_33(
    config_path: impl AsRef<Path>,
    paths: &[PathBuf],
) -> Result<BatchRenameResult> {
    run_batch_rename(config_path, paths)
}
