//! Configuration, filename, and operation helpers for pytools file-sort prefixes.
//!
//! Operation 4-6/4-7 read their configured content cells with the shared Excel
//! reader and rename files without saving the workbooks. Operations 4-9/4-10
//! handle file-sort prefixes and use the shared Excel reader/writer. No operation
//! in this module invokes COM or Office.

use anyhow::{anyhow, bail, Context, Result};
use chrono::Local;
use excel_reader::{read_workbook, SheetDump};
use excel_writer::Workbook;
use serde::Serialize;
use std::collections::BTreeMap;
use std::fs::{self, OpenOptions};
use std::io;
use std::path::{Path, PathBuf};
use std::time::{Instant, SystemTime, UNIX_EPOCH};

const CONFIG_SHEET: &str = "文件排序配置";
const REPORT_SHEET: &str = "文件排序前缀结果";
const REPORT_HEADERS: [&str; 8] = [
    "源文件路径",
    "原文件名",
    "新文件名",
    "命中配置行",
    "排序前缀",
    "匹配关键词",
    "状态",
    "备注",
];
const REMOVE_REPORT_SHEET: &str = "取消文件排序前缀结果";
const REMOVE_REPORT_HEADERS: [&str; 5] = ["源文件路径", "原文件名", "新文件名", "状态", "备注"];

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FileSortConfigRow {
    pub row_no: usize,
    pub enabled: String,
    pub sort_prefix: String,
    pub keywords: String,
    pub remark: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct FileSortRule {
    pub row_no: usize,
    pub prefix: String,
    pub keywords: Vec<String>,
    pub keywords_text: String,
    pub remark: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct FileSortPrefixRow {
    pub source_path: String,
    pub original_name: String,
    pub new_name: String,
    pub config_row: Option<usize>,
    pub sort_prefix: String,
    pub matched_keywords: String,
    pub status: String,
    pub remark: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct FileSortPrefixRunReport {
    pub output: PathBuf,
    pub rules: usize,
    pub files: usize,
    pub renamed: usize,
    pub unmatched: usize,
    pub skipped: usize,
    pub failed: usize,
    pub rows: Vec<FileSortPrefixRow>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct FileSortPrefixRemovalRow {
    pub source_path: String,
    pub original_name: String,
    pub new_name: String,
    pub status: String,
    pub remark: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct FileSortPrefixRemovalRunReport {
    pub output: PathBuf,
    pub rules: usize,
    pub files: usize,
    pub renamed: usize,
    pub skipped: usize,
    pub failed: usize,
    pub rows: Vec<FileSortPrefixRemovalRow>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct FileSortProgressEvent {
    pub sequence: usize,
    pub feature_id: String,
    pub phase: String,
    pub status: String,
    pub completed_files: usize,
    pub total_files: usize,
    pub elapsed_ms: u128,
    pub components: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub actual_engine: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_type: Option<String>,
    pub user_message: String,
}

fn emit_file_sort_event<F>(
    on_event: &mut F,
    sequence: &mut usize,
    started: &Instant,
    feature_id: &str,
    phase: &str,
    status: &str,
    completed_files: usize,
    total_files: usize,
    components: &[&str],
    actual_engine: Option<&str>,
    error_type: Option<&str>,
    user_message: impl Into<String>,
) where
    F: FnMut(&FileSortProgressEvent),
{
    *sequence += 1;
    on_event(&FileSortProgressEvent {
        sequence: *sequence,
        feature_id: feature_id.to_owned(),
        phase: phase.to_owned(),
        status: status.to_owned(),
        completed_files,
        total_files,
        elapsed_ms: started.elapsed().as_millis(),
        components: components
            .iter()
            .map(|component| (*component).to_owned())
            .collect(),
        actual_engine: actual_engine.map(str::to_owned),
        error_type: error_type.map(str::to_owned),
        user_message: user_message.into(),
    });
}

fn prepare_file_sort_output_dir<F>(
    output_dir: &Path,
    on_event: &mut F,
    sequence: &mut usize,
    started: &Instant,
    feature_id: &str,
    total_files: usize,
) -> Result<()>
where
    F: FnMut(&FileSortProgressEvent),
{
    emit_file_sort_event(
        on_event,
        sequence,
        started,
        feature_id,
        "output_directory_prepare",
        "started",
        0,
        total_files,
        &["Rust"],
        None,
        None,
        "正在准备输出目录。",
    );
    if let Err(error) = fs::create_dir_all(output_dir) {
        emit_file_sort_event(
            on_event,
            sequence,
            started,
            feature_id,
            "output_directory_prepare",
            "failed",
            0,
            total_files,
            &["Rust"],
            Some("RustFilesystem"),
            Some("output_directory_create_failed"),
            "输出目录准备失败。",
        );
        return Err(error)
            .with_context(|| format!("create output directory {}", output_dir.display()));
    }
    emit_file_sort_event(
        on_event,
        sequence,
        started,
        feature_id,
        "output_directory_prepare",
        "completed",
        0,
        total_files,
        &["Rust"],
        Some("RustFilesystem"),
        None,
        "输出目录已准备。",
    );
    Ok(())
}

fn file_sort_processing_engine(used_calamine: bool) -> (&'static [&'static str], &'static str) {
    if used_calamine {
        (
            &["Rust", "excel-reader", "Calamine", "RustFilesystem"],
            "Calamine + RustFilesystem",
        )
    } else {
        (&["Rust", "RustFilesystem"], "RustFilesystem")
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct ContentPrefixRenameRow {
    pub source_path: String,
    pub original_name: String,
    pub new_name: String,
    pub status: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct ContentPrefixRenameRunReport {
    pub ok: usize,
    pub skip: usize,
    pub spec: String,
    pub config_row: usize,
    pub config_remark: String,
    pub rows: Vec<ContentPrefixRenameRow>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct ContentPrefixUnrenameRow {
    pub source_path: String,
    pub original_name: String,
    pub new_name: String,
    pub status: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct ContentPrefixUnrenameRunReport {
    pub ok: usize,
    pub skip: usize,
    pub spec: String,
    pub config_row: usize,
    pub config_remark: String,
    pub rows: Vec<ContentPrefixUnrenameRow>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ContentRenameConfig {
    row_no: usize,
    spec: String,
    remark: String,
}

/// Execute pytools operation 4-9 against the supplied files and write its XLSX report.
///
/// The config workbook is read through the shared Calamine-backed Excel reader. File
/// moves reserve their destination with a hard link before removing the source, so a
/// concurrently-created target is never overwritten. If report generation fails,
/// completed moves are rolled back from their recorded source/target pairs.
pub fn run_file_sort_prefix(
    config_path: impl AsRef<Path>,
    files: &[PathBuf],
    output_dir: impl AsRef<Path>,
) -> Result<FileSortPrefixRunReport> {
    run_file_sort_prefix_with_progress(config_path, files, output_dir, |_| {})
}

pub fn run_file_sort_prefix_with_progress<F>(
    config_path: impl AsRef<Path>,
    files: &[PathBuf],
    output_dir: impl AsRef<Path>,
    on_event: F,
) -> Result<FileSortPrefixRunReport>
where
    F: FnMut(&FileSortProgressEvent),
{
    run_file_sort_prefix_with_report_writer_and_progress(
        config_path.as_ref(),
        files,
        output_dir.as_ref(),
        write_report,
        on_event,
    )
}

/// Execute pytools operation 4-10 and write its XLSX report.
///
/// Direct sort prefixes are removed without opening the workbook. When the sort
/// prefix follows a content prefix, configured cells are read with the shared
/// Calamine-backed reader. Errors while opening a candidate workbook are treated
/// as a non-match, matching the source operation's behavior. File moves never
/// replace an existing target; completed moves are rolled back if report creation
/// fails.
pub fn run_remove_file_sort_prefix(
    config_path: impl AsRef<Path>,
    files: &[PathBuf],
    output_dir: impl AsRef<Path>,
) -> Result<FileSortPrefixRemovalRunReport> {
    run_remove_file_sort_prefix_with_progress(config_path, files, output_dir, |_| {})
}

pub fn run_remove_file_sort_prefix_with_progress<F>(
    config_path: impl AsRef<Path>,
    files: &[PathBuf],
    output_dir: impl AsRef<Path>,
    on_event: F,
) -> Result<FileSortPrefixRemovalRunReport>
where
    F: FnMut(&FileSortProgressEvent),
{
    run_remove_file_sort_prefix_with_report_writer_and_progress(
        config_path.as_ref(),
        files,
        output_dir.as_ref(),
        write_remove_report,
        on_event,
    )
}

/// Execute pytools operation 4-6 for XLSX/XLSM/XLS files using the shared Calamine reader.
///
/// The Rust path replaces the source COM read for BIFF `.xls` with Calamine. Other
/// legacy template formats remain closed until their reader support is verified.
/// It reads the configured values, prefixes the filename, and never saves or
/// rewrites a workbook.
pub fn run_content_prefix_rename_xlsx_candidate(
    config_path: impl AsRef<Path>,
    files: &[PathBuf],
) -> Result<ContentPrefixRenameRunReport> {
    run_content_prefix_rename_xlsx_candidate_with_progress(config_path, files, |_| {})
}

pub fn run_content_prefix_rename_xlsx_candidate_with_progress<F>(
    config_path: impl AsRef<Path>,
    files: &[PathBuf],
    mut on_event: F,
) -> Result<ContentPrefixRenameRunReport>
where
    F: FnMut(&FileSortProgressEvent),
{
    let config_path = config_path.as_ref();
    let started = Instant::now();
    let mut sequence = 0;
    emit_file_sort_event(
        &mut on_event,
        &mut sequence,
        &started,
        "PT-4-6",
        "run",
        "started",
        0,
        files.len(),
        &["Rust"],
        None,
        None,
        format!(
            "开始根据工作簿内容添加文件前缀，共 {} 个文件。",
            files.len()
        ),
    );
    emit_file_sort_event(
        &mut on_event,
        &mut sequence,
        &started,
        "PT-4-6",
        "configuration_read",
        "started",
        0,
        files.len(),
        &["Rust", "excel-reader", "Calamine"],
        Some("Calamine"),
        None,
        "正在读取内容前缀配置。",
    );
    let config = match load_content_rename_config(config_path) {
        Ok(config) => config,
        Err(error) => {
            emit_file_sort_event(
                &mut on_event,
                &mut sequence,
                &started,
                "PT-4-6",
                "configuration_read",
                "failed",
                0,
                files.len(),
                &["Rust", "excel-reader", "Calamine"],
                Some("Calamine"),
                Some("configuration_read_failed"),
                "内容前缀配置读取失败。",
            );
            return Err(error);
        }
    };
    let specs = match parse_sheet_cell_spec(&config.spec) {
        Ok(specs) => specs,
        Err(error) => {
            emit_file_sort_event(
                &mut on_event,
                &mut sequence,
                &started,
                "PT-4-6",
                "configuration_read",
                "failed",
                0,
                files.len(),
                &["Rust", "excel-reader", "Calamine"],
                Some("Calamine"),
                Some("content_prefix_spec_invalid"),
                "内容前缀单元格规格无效。",
            );
            return Err(error);
        }
    };
    emit_file_sort_event(
        &mut on_event,
        &mut sequence,
        &started,
        "PT-4-6",
        "configuration_read",
        "completed",
        0,
        files.len(),
        &["Rust", "excel-reader", "Calamine"],
        Some("Calamine"),
        None,
        "内容前缀配置已读取。",
    );

    let mut rows = Vec::with_capacity(files.len());
    let mut ok = 0;
    let mut skip = 0;
    emit_file_sort_event(
        &mut on_event,
        &mut sequence,
        &started,
        "PT-4-6",
        "task",
        "started",
        0,
        files.len(),
        &["Rust"],
        None,
        None,
        "开始处理所选文件。",
    );

    for (index, source) in files.iter().enumerate() {
        let input_index = index + 1;
        emit_file_sort_event(
            &mut on_event,
            &mut sequence,
            &started,
            "PT-4-6",
            "file_processing",
            "started",
            index,
            files.len(),
            &["Rust"],
            None,
            None,
            format!("正在处理第 {input_index}/{} 个文件。", files.len()),
        );
        let (row, used_reader) = content_prefix_rename_row(source, &specs);
        if row.new_name.is_empty() {
            skip += 1;
        } else {
            ok += 1;
        }
        let (components, actual_engine) = if used_reader {
            (
                &["Rust", "excel-reader", "Calamine", "RustFilesystem"][..],
                "Calamine + RustFilesystem",
            )
        } else {
            (&["Rust", "RustFilesystem"][..], "RustFilesystem")
        };
        let error_type = match row.status.as_str() {
            "跳过：内容读取失败" => Some("content_read_failed"),
            "跳过：改名失败" => Some("rename_failed"),
            _ => None,
        };
        let user_message = if row.new_name.is_empty() {
            format!(
                "已处理 {input_index}/{} 个文件；本项跳过，成功 {ok}，跳过 {skip}。",
                files.len()
            )
        } else {
            format!(
                "已处理 {input_index}/{} 个文件；成功 {ok}，跳过 {skip}。",
                files.len()
            )
        };
        emit_file_sort_event(
            &mut on_event,
            &mut sequence,
            &started,
            "PT-4-6",
            "file_processing",
            "completed",
            input_index,
            files.len(),
            components,
            Some(actual_engine),
            error_type,
            user_message,
        );
        rows.push(row);
    }

    let report = ContentPrefixRenameRunReport {
        ok,
        skip,
        spec: config.spec,
        config_row: config.row_no,
        config_remark: config.remark,
        rows,
    };
    emit_file_sort_event(
        &mut on_event,
        &mut sequence,
        &started,
        "PT-4-6",
        "task",
        "completed",
        files.len(),
        files.len(),
        &["Rust", "excel-reader", "Calamine", "RustFilesystem"],
        Some("Calamine + RustFilesystem"),
        None,
        format!(
            "内容前缀处理完成：成功 {} 个，跳过 {} 个。",
            report.ok, report.skip
        ),
    );
    Ok(report)
}

/// Execute pytools operation 4-7 for XLSX/XLSM files using the shared Calamine reader.
///
/// This candidate deliberately limits content-cell reads to XLSX/XLSM until the
/// source COM paths for XLS/XLT and the remaining accepted extensions have
/// format-specific parity evidence. It renames files only; workbook bytes are not
/// modified. Existing targets are reserved without replacement.
pub fn run_content_prefix_unrename_xlsx_candidate(
    config_path: impl AsRef<Path>,
    files: &[PathBuf],
) -> Result<ContentPrefixUnrenameRunReport> {
    run_content_prefix_unrename_xlsx_candidate_with_progress(config_path, files, |_| {})
}

pub fn run_content_prefix_unrename_xlsx_candidate_with_progress<F>(
    config_path: impl AsRef<Path>,
    files: &[PathBuf],
    mut on_event: F,
) -> Result<ContentPrefixUnrenameRunReport>
where
    F: FnMut(&FileSortProgressEvent),
{
    let config_path = config_path.as_ref();
    let started = Instant::now();
    let mut sequence = 0;
    emit_file_sort_event(
        &mut on_event,
        &mut sequence,
        &started,
        "PT-4-7",
        "run",
        "started",
        0,
        files.len(),
        &["Rust"],
        None,
        None,
        format!(
            "开始根据工作簿内容取消文件前缀，共 {} 个文件。",
            files.len()
        ),
    );
    emit_file_sort_event(
        &mut on_event,
        &mut sequence,
        &started,
        "PT-4-7",
        "configuration_read",
        "started",
        0,
        files.len(),
        &["Rust", "excel-reader", "Calamine"],
        Some("Calamine"),
        None,
        "正在读取内容前缀配置。",
    );
    let config = match load_content_rename_config(config_path) {
        Ok(config) => config,
        Err(error) => {
            emit_file_sort_event(
                &mut on_event,
                &mut sequence,
                &started,
                "PT-4-7",
                "configuration_read",
                "failed",
                0,
                files.len(),
                &["Rust", "excel-reader", "Calamine"],
                Some("Calamine"),
                Some("configuration_read_failed"),
                "内容前缀配置读取失败。",
            );
            return Err(error);
        }
    };
    let specs = match parse_sheet_cell_spec(&config.spec) {
        Ok(specs) => specs,
        Err(error) => {
            emit_file_sort_event(
                &mut on_event,
                &mut sequence,
                &started,
                "PT-4-7",
                "configuration_read",
                "failed",
                0,
                files.len(),
                &["Rust", "excel-reader", "Calamine"],
                Some("Calamine"),
                Some("content_prefix_spec_invalid"),
                "内容前缀单元格规格无效。",
            );
            return Err(error);
        }
    };

    // Operation 4-7 treats an absent/invalid legacy sort-config sheet as an
    // empty list and still allows removal of a content prefix at the filename start.
    let mut sort_prefixes = load_file_sort_rules(config_path)
        .unwrap_or_default()
        .into_iter()
        .map(|rule| rule.prefix)
        .collect::<std::collections::BTreeSet<_>>()
        .into_iter()
        .collect::<Vec<_>>();
    sort_prefixes.sort_by(|left, right| {
        right
            .chars()
            .count()
            .cmp(&left.chars().count())
            .then_with(|| left.cmp(right))
    });

    emit_file_sort_event(
        &mut on_event,
        &mut sequence,
        &started,
        "PT-4-7",
        "configuration_read",
        "completed",
        0,
        files.len(),
        &["Rust", "excel-reader", "Calamine"],
        Some("Calamine"),
        None,
        "内容前缀配置已读取。",
    );

    let mut rows = Vec::with_capacity(files.len());
    let mut ok = 0;
    let mut skip = 0;
    emit_file_sort_event(
        &mut on_event,
        &mut sequence,
        &started,
        "PT-4-7",
        "task",
        "started",
        0,
        files.len(),
        &["Rust"],
        None,
        None,
        "开始处理所选文件。",
    );
    for (index, source) in files.iter().enumerate() {
        let input_index = index + 1;
        emit_file_sort_event(
            &mut on_event,
            &mut sequence,
            &started,
            "PT-4-7",
            "file_processing",
            "started",
            index,
            files.len(),
            &["Rust"],
            None,
            None,
            format!("正在处理第 {input_index}/{} 个文件。", files.len()),
        );
        let (row, used_reader) = content_prefix_unrename_row(source, &specs, &sort_prefixes);
        if row.new_name.is_empty() {
            skip += 1;
        } else {
            ok += 1;
        }
        let (components, actual_engine) = if used_reader {
            (
                &["Rust", "excel-reader", "Calamine", "RustFilesystem"][..],
                "Calamine + RustFilesystem",
            )
        } else {
            (&["Rust", "RustFilesystem"][..], "RustFilesystem")
        };
        let error_type = match row.status.as_str() {
            "跳过：内容读取失败" => Some("content_read_failed"),
            "跳过：改名失败" => Some("rename_failed"),
            _ => None,
        };
        let user_message = if row.new_name.is_empty() {
            format!(
                "已处理 {input_index}/{} 个文件；本项跳过，成功 {ok}，跳过 {skip}。",
                files.len()
            )
        } else {
            format!(
                "已处理 {input_index}/{} 个文件；成功 {ok}，跳过 {skip}。",
                files.len()
            )
        };
        emit_file_sort_event(
            &mut on_event,
            &mut sequence,
            &started,
            "PT-4-7",
            "file_processing",
            "completed",
            input_index,
            files.len(),
            components,
            Some(actual_engine),
            error_type,
            user_message,
        );
        rows.push(row);
    }

    let report = ContentPrefixUnrenameRunReport {
        ok,
        skip,
        spec: config.spec,
        config_row: config.row_no,
        config_remark: config.remark,
        rows,
    };
    emit_file_sort_event(
        &mut on_event,
        &mut sequence,
        &started,
        "PT-4-7",
        "task",
        "completed",
        files.len(),
        files.len(),
        &["Rust", "excel-reader", "Calamine", "RustFilesystem"],
        Some("Calamine + RustFilesystem"),
        None,
        format!(
            "取消内容前缀完成：成功 {} 个，跳过 {} 个。",
            report.ok, report.skip
        ),
    );
    Ok(report)
}

fn content_prefix_rename_row(
    source: &Path,
    specs: &[(String, String)],
) -> (ContentPrefixRenameRow, bool) {
    let mut row = ContentPrefixRenameRow {
        source_path: source.to_string_lossy().into_owned(),
        original_name: source
            .file_name()
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_default(),
        new_name: String::new(),
        status: String::new(),
    };
    if !source.exists() || !source.is_file() {
        row.status = "跳过：非文件".to_owned();
        return (row, false);
    }
    if !matches!(
        source
            .extension()
            .and_then(|extension| extension.to_str())
            .map(str::to_ascii_lowercase)
            .as_deref(),
        Some("xlsx" | "xlsm" | "xls")
    ) {
        row.status = "跳过：格式未验收".to_owned();
        return (row, false);
    }
    let values = match read_content_prefix_items(source, specs) {
        Ok(values) => values,
        Err(_) => {
            row.status = "跳过：内容读取失败".to_owned();
            return (row, true);
        }
    };
    let prefix_items = values
        .iter()
        .map(|value| sanitize_name_part(value))
        .filter(|value| !value.is_empty())
        .collect::<Vec<_>>();
    if already_renamed_by_content(&row.original_name, &prefix_items) {
        row.status = "跳过：内容已在文件名中".to_owned();
        return (row, true);
    }
    let prefix = prefix_items.join("_").chars().take(180).collect::<String>();
    let new_name = if prefix.is_empty() {
        row.original_name.clone()
    } else {
        format!("{prefix}_{}", row.original_name)
    };
    if new_name == row.original_name {
        row.status = "跳过：文件名未变".to_owned();
        return (row, true);
    }
    match move_to_available_path(source, &source.with_file_name(&new_name)) {
        Ok(actual_target) => {
            row.new_name = actual_target
                .file_name()
                .map(|name| name.to_string_lossy().into_owned())
                .unwrap_or_default();
            row.status = "已添加内容前缀".to_owned();
        }
        Err(_) => row.status = "跳过：改名失败".to_owned(),
    }
    (row, true)
}

fn content_prefix_unrename_row(
    source: &Path,
    specs: &[(String, String)],
    sort_prefixes: &[String],
) -> (ContentPrefixUnrenameRow, bool) {
    let mut row = ContentPrefixUnrenameRow {
        source_path: source.to_string_lossy().into_owned(),
        original_name: source
            .file_name()
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_default(),
        new_name: String::new(),
        status: String::new(),
    };
    if !source.exists() || !source.is_file() {
        row.status = "跳过：非文件".to_owned();
        return (row, false);
    }
    if !matches!(
        source
            .extension()
            .and_then(|extension| extension.to_str())
            .map(str::to_ascii_lowercase)
            .as_deref(),
        Some("xlsx" | "xlsm")
    ) {
        row.status = "跳过：格式未验收".to_owned();
        return (row, false);
    }
    let values = match read_content_prefix_items(source, specs) {
        Ok(values) => values,
        Err(_) => {
            row.status = "跳过：内容读取失败".to_owned();
            return (row, true);
        }
    };
    let prefix_items = values
        .iter()
        .map(|value| sanitize_name_part(value))
        .filter(|value| !value.is_empty())
        .collect::<Vec<_>>();
    let Some(new_name) =
        remove_content_prefix_keep_sort_prefix(&row.original_name, &prefix_items, sort_prefixes)
    else {
        row.status = "跳过：前缀不匹配".to_owned();
        return (row, true);
    };
    match move_to_available_path(source, &source.with_file_name(&new_name)) {
        Ok(actual_target) => {
            row.new_name = actual_target
                .file_name()
                .map(|name| name.to_string_lossy().into_owned())
                .unwrap_or_default();
            row.status = "已取消内容前缀".to_owned();
        }
        Err(_) => row.status = "跳过：改名失败".to_owned(),
    }
    (row, true)
}

#[cfg(test)]
fn run_remove_file_sort_prefix_with_report_writer<F>(
    config_path: &Path,
    files: &[PathBuf],
    output_dir: &Path,
    report_writer: F,
) -> Result<FileSortPrefixRemovalRunReport>
where
    F: FnOnce(&Path, &[FileSortPrefixRemovalRow]) -> Result<PathBuf>,
{
    run_remove_file_sort_prefix_with_report_writer_and_progress(
        config_path,
        files,
        output_dir,
        report_writer,
        |_| {},
    )
}

fn run_remove_file_sort_prefix_with_report_writer_and_progress<F, G>(
    config_path: &Path,
    files: &[PathBuf],
    output_dir: &Path,
    report_writer: F,
    mut on_event: G,
) -> Result<FileSortPrefixRemovalRunReport>
where
    F: FnOnce(&Path, &[FileSortPrefixRemovalRow]) -> Result<PathBuf>,
    G: FnMut(&FileSortProgressEvent),
{
    let started = Instant::now();
    let mut sequence = 0;
    emit_file_sort_event(
        &mut on_event,
        &mut sequence,
        &started,
        "PT-4-10",
        "configuration_read",
        "started",
        0,
        files.len(),
        &["Rust", "excel-reader", "Calamine"],
        Some("Calamine"),
        None,
        "正在读取文件排序配置。",
    );
    let rules = match load_file_sort_rules(config_path) {
        Ok(rules) => rules,
        Err(error) => {
            emit_file_sort_event(
                &mut on_event,
                &mut sequence,
                &started,
                "PT-4-10",
                "configuration_read",
                "failed",
                0,
                files.len(),
                &["Rust", "excel-reader", "Calamine"],
                Some("Calamine"),
                Some("configuration_read_failed"),
                "配置读取失败。",
            );
            return Err(error);
        }
    };
    if rules.is_empty() {
        emit_file_sort_event(
            &mut on_event,
            &mut sequence,
            &started,
            "PT-4-10",
            "configuration_read",
            "failed",
            0,
            files.len(),
            &["Rust", "excel-reader", "Calamine"],
            Some("Calamine"),
            Some("no_enabled_sort_rules"),
            "没有启用的排序规则。",
        );
        bail!("文件排序配置 没有启用且有效的排序规则。");
    }
    let mut prefixes = rules
        .iter()
        .map(|rule| rule.prefix.clone())
        .collect::<std::collections::BTreeSet<_>>()
        .into_iter()
        .collect::<Vec<_>>();
    prefixes.sort_by(|left, right| {
        right
            .chars()
            .count()
            .cmp(&left.chars().count())
            .then_with(|| left.cmp(right))
    });
    let content_spec_items = load_content_rename_spec(config_path)
        .and_then(|spec| parse_sheet_cell_spec(&spec))
        .unwrap_or_default();

    emit_file_sort_event(
        &mut on_event,
        &mut sequence,
        &started,
        "PT-4-10",
        "configuration_read",
        "completed",
        0,
        files.len(),
        &["Rust", "excel-reader", "Calamine"],
        Some("Calamine"),
        None,
        format!("已读取 {} 条有效排序规则及内容前缀设置。", rules.len()),
    );

    prepare_file_sort_output_dir(
        output_dir,
        &mut on_event,
        &mut sequence,
        &started,
        "PT-4-10",
        files.len(),
    )?;

    let mut rows = Vec::with_capacity(files.len());
    let mut renamed = 0;
    let mut skipped = 0;
    let mut failed = 0;
    let mut moved = Vec::<(PathBuf, PathBuf)>::new();

    for (index, source) in files.iter().enumerate() {
        let file_started = Instant::now();
        emit_file_sort_event(
            &mut on_event,
            &mut sequence,
            &started,
            "PT-4-10",
            "file_processing",
            "started",
            index,
            files.len(),
            &["Rust"],
            None,
            None,
            format!("正在处理文件 {}/{}。", index + 1, files.len()),
        );
        let mut row = FileSortPrefixRemovalRow {
            source_path: source.to_string_lossy().into_owned(),
            original_name: source
                .file_name()
                .map(|name| name.to_string_lossy().into_owned())
                .unwrap_or_default(),
            new_name: String::new(),
            status: String::new(),
            remark: String::new(),
        };
        let mut used_calamine = false;

        if !source.exists() || !source.is_file() {
            skipped += 1;
            row.status = "跳过：非文件".to_owned();
            rows.push(row);
            let (components, engine) = file_sort_processing_engine(used_calamine);
            emit_file_sort_event(
                &mut on_event,
                &mut sequence,
                &started,
                "PT-4-10",
                "file_processing",
                "completed",
                index + 1,
                files.len(),
                components,
                Some(engine),
                None,
                format!("已处理 {}/{} 个文件（跳过）。", index + 1, files.len()),
            );
            continue;
        }

        let original_name = &row.original_name;
        let direct_prefix = prefixes
            .iter()
            .find(|prefix| original_name.starts_with(&format!("{prefix}_")));

        let (base_name, mut status) = if let Some(prefix) = direct_prefix {
            (
                original_name[prefix.len() + 1..].to_owned(),
                "已取消排序前缀".to_owned(),
            )
        } else {
            match content_prefixed_base_name(
                original_name,
                &prefixes,
                &content_spec_items,
                source,
                &mut used_calamine,
            ) {
                Some(name) => (name, "已取消排序前缀（保留内容前缀）".to_owned()),
                None => {
                    skipped += 1;
                    row.status = "跳过：无匹配排序前缀".to_owned();
                    rows.push(row);
                    let (components, engine) = file_sort_processing_engine(used_calamine);
                    emit_file_sort_event(
                        &mut on_event,
                        &mut sequence,
                        &started,
                        "PT-4-10",
                        "file_processing",
                        "completed",
                        index + 1,
                        files.len(),
                        components,
                        Some(engine),
                        None,
                        format!(
                            "已处理 {}/{} 个文件（无匹配规则）。",
                            index + 1,
                            files.len()
                        ),
                    );
                    continue;
                }
            }
        };

        if base_name.is_empty() {
            skipped += 1;
            row.status = "跳过：无匹配排序前缀".to_owned();
            rows.push(row);
            let (components, engine) = file_sort_processing_engine(used_calamine);
            emit_file_sort_event(
                &mut on_event,
                &mut sequence,
                &started,
                "PT-4-10",
                "file_processing",
                "completed",
                index + 1,
                files.len(),
                components,
                Some(engine),
                None,
                format!(
                    "已处理 {}/{} 个文件（无匹配规则）。",
                    index + 1,
                    files.len()
                ),
            );
            continue;
        }

        let desired = source.with_file_name(&base_name);
        match move_to_available_path(source, &desired) {
            Ok(actual_target) => {
                if actual_target != *source {
                    moved.push((source.clone(), actual_target.clone()));
                }
                if actual_target != desired {
                    status = "目标同名：已自动后缀".to_owned();
                }
                renamed += 1;
                row.new_name = actual_target
                    .file_name()
                    .map(|name| name.to_string_lossy().into_owned())
                    .unwrap_or_default();
                row.status = status;
            }
            Err(error) => {
                failed += 1;
                row.status = "失败".to_owned();
                row.remark = error.to_string();
            }
        }
        let row_status = row.status.as_str();
        let event_status = if row_status == "失败" {
            "failed"
        } else {
            "completed"
        };
        let event_message = if row_status == "失败" {
            format!("文件 {}/{} 处理失败。", index + 1, files.len())
        } else {
            format!("已处理 {}/{} 个文件。", index + 1, files.len())
        };
        rows.push(row);
        let (components, engine) = file_sort_processing_engine(used_calamine);
        emit_file_sort_event(
            &mut on_event,
            &mut sequence,
            &started,
            "PT-4-10",
            "file_processing",
            event_status,
            index + 1,
            files.len(),
            components,
            Some(engine),
            (event_status == "failed").then_some("file_rename_failed"),
            format!(
                "{event_message} 本文件用时 {}ms。",
                file_started.elapsed().as_millis()
            ),
        );
    }

    emit_file_sort_event(
        &mut on_event,
        &mut sequence,
        &started,
        "PT-4-10",
        "report_write",
        "started",
        files.len(),
        files.len(),
        &["Rust", "excel-writer", "rust_xlsxwriter"],
        Some("rust_xlsxwriter"),
        None,
        "正在生成结果报表。",
    );
    let output = match report_writer(output_dir, &rows) {
        Ok(path) => path,
        Err(error) => {
            let rollback_errors = rollback_moves(&moved);
            if rollback_errors.is_empty() {
                emit_file_sort_event(
                    &mut on_event,
                    &mut sequence,
                    &started,
                    "PT-4-10",
                    "report_write",
                    "failed",
                    files.len(),
                    files.len(),
                    &["Rust", "excel-writer", "rust_xlsxwriter"],
                    Some("rust_xlsxwriter"),
                    Some("report_write_failed"),
                    "结果报表生成失败，已回滚文件改名。",
                );
                return Err(error)
                    .context("write file-sort-prefix-removal report; file moves rolled back");
            }
            emit_file_sort_event(
                &mut on_event,
                &mut sequence,
                &started,
                "PT-4-10",
                "report_write",
                "failed",
                files.len(),
                files.len(),
                &["Rust", "excel-writer", "rust_xlsxwriter"],
                Some("rust_xlsxwriter"),
                Some("report_write_and_rollback_failed"),
                "结果报表生成及文件回滚失败。",
            );
            return Err(anyhow!(
                "write file-sort-prefix-removal report failed: {error:#}; rollback also failed: {}",
                rollback_errors.join("; ")
            ));
        }
    };

    emit_file_sort_event(
        &mut on_event,
        &mut sequence,
        &started,
        "PT-4-10",
        "report_write",
        "completed",
        files.len(),
        files.len(),
        &["Rust", "excel-writer", "rust_xlsxwriter"],
        Some("rust_xlsxwriter"),
        None,
        "结果报表已生成。",
    );
    emit_file_sort_event(
        &mut on_event,
        &mut sequence,
        &started,
        "PT-4-10",
        "task",
        if failed == 0 {
            "completed"
        } else {
            "partial_failure"
        },
        files.len(),
        files.len(),
        &["Rust"],
        None,
        (failed > 0).then_some("one_or_more_files_failed"),
        format!(
            "完成：重命名 {} 个，跳过 {} 个，失败 {} 个。",
            renamed, skipped, failed
        ),
    );

    Ok(FileSortPrefixRemovalRunReport {
        output,
        rules: rules.len(),
        files: files.len(),
        renamed,
        skipped,
        failed,
        rows,
    })
}

fn load_content_rename_config(path: &Path) -> Result<ContentRenameConfig> {
    let workbook = read_workbook(path)
        .with_context(|| format!("read content-rename configuration {}", path.display()))?;
    let sheet = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "重命名配置")
        .ok_or_else(|| anyhow!("missing worksheet 重命名配置"))?;
    let mut enabled = Vec::<(usize, String, String)>::new();
    for row_no in 2..=max_row(sheet) {
        if !is_enabled(&field_at(sheet, row_no, 15)) {
            continue;
        }
        enabled.push((
            row_no,
            field_at(sheet, row_no, 16),
            field_at(sheet, row_no, 17),
        ));
    }
    match enabled.as_slice() {
        [] => bail!("no enabled content rename configuration"),
        [(row_no, spec, remark)] => Ok(ContentRenameConfig {
            row_no: *row_no,
            spec: spec.clone(),
            remark: remark.clone(),
        }),
        _ => {
            let rows = enabled
                .iter()
                .map(|(row, _, remark)| {
                    if remark.is_empty() {
                        format!("第{row}行")
                    } else {
                        format!("第{row}行（{remark}）")
                    }
                })
                .collect::<Vec<_>>();
            bail!(
                "content rename configuration allows one active row: {}",
                rows.join("、")
            )
        }
    }
}

fn load_content_rename_spec(path: &Path) -> Result<String> {
    load_content_rename_config(path).map(|config| config.spec)
}

fn field_at(sheet: &SheetDump, row: usize, column: usize) -> String {
    sheet
        .values
        .get(&cell_address(row, column))
        .map(cell_text)
        .unwrap_or_default()
}

fn parse_sheet_cell_spec(spec: &str) -> Result<Vec<(String, String)>> {
    let mut items = Vec::new();
    for part in spec
        .replace('；', ";")
        .split(';')
        .map(str::trim)
        .filter(|p| !p.is_empty())
    {
        let (sheet, address) = part
            .split_once('@')
            .ok_or_else(|| anyhow!("invalid content rename spec {part:?}; expected sheet@A1"))?;
        let sheet = sheet.trim();
        let address = address.trim().to_ascii_uppercase();
        let Some((row, column)) = parse_cell_address(&address) else {
            bail!("invalid content rename cell address in {part:?}");
        };
        if sheet.is_empty() || row == 0 || column == 0 {
            bail!("invalid content rename spec {part:?}; sheet and A1 address are required");
        }
        items.push((sheet.to_owned(), address));
    }
    if items.is_empty() {
        bail!("content rename spec is empty");
    }
    Ok(items)
}

fn read_content_prefix_items(path: &Path, specs: &[(String, String)]) -> Result<Vec<String>> {
    let workbook = read_workbook(path)
        .with_context(|| format!("read content-prefix workbook {}", path.display()))?;
    let uses_legacy_xls_value_coercion = path
        .extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| extension.eq_ignore_ascii_case("xls"));
    let mut items = Vec::with_capacity(specs.len());
    for (sheet_name, address) in specs {
        let raw = workbook
            .sheets
            .iter()
            .find(|sheet| sheet.name == *sheet_name)
            .and_then(|sheet| sheet.values.get(address))
            .map(|value| content_rename_cell_text(value, uses_legacy_xls_value_coercion))
            .unwrap_or_default();
        let sanitized = sanitize_name_part(&raw);
        items.push(if sanitized.is_empty() {
            format!("{sheet_name}!{address}")
        } else {
            sanitized
        });
    }
    Ok(items)
}

fn content_prefixed_base_name(
    file_name: &str,
    prefixes: &[String],
    specs: &[(String, String)],
    source: &Path,
    used_calamine: &mut bool,
) -> Option<String> {
    if specs.is_empty() {
        return None;
    }
    for sort_prefix in prefixes {
        let token = format!("_{sort_prefix}_");
        let Some(position) = file_name.find(&token) else {
            continue;
        };
        if position == 0 {
            continue;
        }
        *used_calamine = true;
        let Ok(values) = read_content_prefix_items(source, specs) else {
            continue;
        };
        let content_items = values
            .iter()
            .map(|value| sanitize_name_part(value))
            .filter(|value| !value.is_empty())
            .collect::<Vec<_>>();
        if content_items.is_empty() {
            continue;
        }
        let content_prefix = &file_name[..position + 1];
        let probe = format!("{content_prefix}__probe__.xlsx");
        if remove_content_prefix(&probe, &content_items).as_deref() != Some("__probe__.xlsx") {
            continue;
        }
        let tail = &file_name[position + token.len()..];
        return Some(format!("{content_prefix}{tail}"));
    }
    None
}

fn sanitize_name_part(value: &str) -> String {
    let mut output = String::new();
    let mut pending_space = false;
    for character in value.chars() {
        if is_python_whitespace(character) {
            pending_space = !output.is_empty();
            continue;
        }
        if pending_space {
            output.push(' ');
            pending_space = false;
        }
        if matches!(
            character,
            '<' | '>' | ':' | '"' | '/' | '\\' | '|' | '?' | '*'
        ) {
            output.push('_');
        } else {
            output.push(character);
        }
    }
    let trimmed = output.trim_matches('.');
    trimmed.to_owned()
}

fn remove_content_prefix(file_name: &str, prefix_items: &[String]) -> Option<String> {
    if prefix_items.is_empty() {
        return None;
    }
    let prefix = format!("{}_", prefix_items.join("_"));
    let path = Path::new(file_name);
    let mut stem = path.file_stem()?.to_string_lossy().into_owned();
    let suffix = path
        .extension()
        .map(|extension| format!(".{}", extension.to_string_lossy()))
        .unwrap_or_default();
    let mut changed = false;
    while stem.starts_with(&prefix) {
        stem.drain(..prefix.len());
        changed = true;
    }
    let prefix_set = prefix_items
        .iter()
        .map(String::as_str)
        .collect::<std::collections::BTreeSet<_>>();
    while let Some(separator) = stem.find('_') {
        if !prefix_set.contains(&&stem[..separator]) {
            break;
        }
        stem.drain(..=separator);
        changed = true;
    }
    if !changed || stem.is_empty() {
        return None;
    }
    Some(format!("{stem}{suffix}"))
}

fn already_renamed_by_content(file_name: &str, prefix_items: &[String]) -> bool {
    if prefix_items.is_empty() {
        return false;
    }
    let Some(stem) = Path::new(file_name)
        .file_stem()
        .and_then(|stem| stem.to_str())
    else {
        return false;
    };
    prefix_items.iter().all(|item| stem.contains(item))
}

fn remove_content_prefix_keep_sort_prefix(
    file_name: &str,
    prefix_items: &[String],
    sort_prefixes: &[String],
) -> Option<String> {
    for sort_prefix in sort_prefixes {
        let leading = format!("{sort_prefix}_");
        let Some(remainder) = file_name.strip_prefix(&leading) else {
            continue;
        };
        let restored = remove_content_prefix(remainder, prefix_items)?;
        return Some(format!("{leading}{restored}"));
    }
    remove_content_prefix(file_name, prefix_items)
}

#[cfg(test)]
fn run_file_sort_prefix_with_report_writer<F>(
    config_path: &Path,
    files: &[PathBuf],
    output_dir: &Path,
    report_writer: F,
) -> Result<FileSortPrefixRunReport>
where
    F: FnOnce(&Path, &[FileSortPrefixRow]) -> Result<PathBuf>,
{
    run_file_sort_prefix_with_report_writer_and_progress(
        config_path,
        files,
        output_dir,
        report_writer,
        |_| {},
    )
}

fn run_file_sort_prefix_with_report_writer_and_progress<F, G>(
    config_path: &Path,
    files: &[PathBuf],
    output_dir: &Path,
    report_writer: F,
    mut on_event: G,
) -> Result<FileSortPrefixRunReport>
where
    F: FnOnce(&Path, &[FileSortPrefixRow]) -> Result<PathBuf>,
    G: FnMut(&FileSortProgressEvent),
{
    let started = Instant::now();
    let mut sequence = 0;
    emit_file_sort_event(
        &mut on_event,
        &mut sequence,
        &started,
        "PT-4-9",
        "configuration_read",
        "started",
        0,
        files.len(),
        &["Rust", "excel-reader", "Calamine"],
        Some("Calamine"),
        None,
        "正在读取文件排序配置。",
    );
    let rules = match load_file_sort_rules(config_path) {
        Ok(rules) => rules,
        Err(error) => {
            emit_file_sort_event(
                &mut on_event,
                &mut sequence,
                &started,
                "PT-4-9",
                "configuration_read",
                "failed",
                0,
                files.len(),
                &["Rust", "excel-reader", "Calamine"],
                Some("Calamine"),
                Some("configuration_read_failed"),
                "配置读取失败。",
            );
            return Err(error);
        }
    };
    if rules.is_empty() {
        emit_file_sort_event(
            &mut on_event,
            &mut sequence,
            &started,
            "PT-4-9",
            "configuration_read",
            "failed",
            0,
            files.len(),
            &["Rust", "excel-reader", "Calamine"],
            Some("Calamine"),
            Some("no_enabled_sort_rules"),
            "没有启用的排序规则。",
        );
        bail!("文件排序配置 没有启用且有效的排序规则。");
    }
    emit_file_sort_event(
        &mut on_event,
        &mut sequence,
        &started,
        "PT-4-9",
        "configuration_read",
        "completed",
        0,
        files.len(),
        &["Rust", "excel-reader", "Calamine"],
        Some("Calamine"),
        None,
        format!("已读取 {} 条有效排序规则。", rules.len()),
    );

    prepare_file_sort_output_dir(
        output_dir,
        &mut on_event,
        &mut sequence,
        &started,
        "PT-4-9",
        files.len(),
    )?;

    let mut rows = Vec::with_capacity(files.len());
    let mut renamed = 0;
    let mut unmatched = 0;
    let mut skipped = 0;
    let mut failed = 0;
    let mut moved = Vec::<(PathBuf, PathBuf)>::new();

    for (index, source) in files.iter().enumerate() {
        let file_started = Instant::now();
        emit_file_sort_event(
            &mut on_event,
            &mut sequence,
            &started,
            "PT-4-9",
            "file_processing",
            "started",
            index,
            files.len(),
            &["Rust"],
            Some("RustFilesystem"),
            None,
            format!("正在处理文件 {}/{}。", index + 1, files.len()),
        );
        let source_path = source.to_string_lossy().into_owned();
        let original_name = source
            .file_name()
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_default();
        let mut row = FileSortPrefixRow {
            source_path,
            original_name: original_name.clone(),
            new_name: String::new(),
            config_row: None,
            sort_prefix: String::new(),
            matched_keywords: String::new(),
            status: String::new(),
            remark: String::new(),
        };

        if !source.exists() || !source.is_file() {
            skipped += 1;
            row.status = "跳过：非文件".to_owned();
            rows.push(row);
            emit_file_sort_event(
                &mut on_event,
                &mut sequence,
                &started,
                "PT-4-9",
                "file_processing",
                "completed",
                index + 1,
                files.len(),
                &["Rust"],
                Some("RustFilesystem"),
                None,
                format!("已处理 {}/{} 个文件（跳过）。", index + 1, files.len()),
            );
            continue;
        }

        let Some(rule) = match_file_sort_rule(&original_name, &rules) else {
            unmatched += 1;
            row.status = "未命中".to_owned();
            rows.push(row);
            emit_file_sort_event(
                &mut on_event,
                &mut sequence,
                &started,
                "PT-4-9",
                "file_processing",
                "completed",
                index + 1,
                files.len(),
                &["Rust"],
                Some("RustFilesystem"),
                None,
                format!(
                    "已处理 {}/{} 个文件（无匹配规则）。",
                    index + 1,
                    files.len()
                ),
            );
            continue;
        };

        let target_name = prefixed_file_name(&original_name, &rule.prefix);
        let target = source.with_file_name(target_name);
        let moved_to = if target == *source {
            Ok(source.clone())
        } else {
            move_to_available_path(source, &target)
        };

        match moved_to {
            Ok(actual_target) => {
                if actual_target != *source {
                    moved.push((source.clone(), actual_target.clone()));
                }
                renamed += 1;
                row.new_name = actual_target
                    .file_name()
                    .map(|name| name.to_string_lossy().into_owned())
                    .unwrap_or_default();
                row.config_row = Some(rule.row_no);
                row.sort_prefix = rule.prefix.clone();
                row.matched_keywords = rule.keywords_text.clone();
                row.remark = rule.remark.clone();
                row.status = if actual_target == target {
                    "已重命名".to_owned()
                } else {
                    "目标同名：已自动后缀".to_owned()
                };
            }
            Err(error) => {
                failed += 1;
                row.status = "失败".to_owned();
                row.remark = error.to_string();
            }
        }
        let row_status = row.status.as_str();
        let event_status = if row_status == "失败" {
            "failed"
        } else {
            "completed"
        };
        let event_message = if row_status == "失败" {
            format!("文件 {}/{} 处理失败。", index + 1, files.len())
        } else {
            format!("已处理 {}/{} 个文件。", index + 1, files.len())
        };
        rows.push(row);
        emit_file_sort_event(
            &mut on_event,
            &mut sequence,
            &started,
            "PT-4-9",
            "file_processing",
            event_status,
            index + 1,
            files.len(),
            &["Rust"],
            Some("RustFilesystem"),
            (event_status == "failed").then_some("file_rename_failed"),
            format!(
                "{event_message} 本文件用时 {}ms。",
                file_started.elapsed().as_millis()
            ),
        );
    }

    emit_file_sort_event(
        &mut on_event,
        &mut sequence,
        &started,
        "PT-4-9",
        "report_write",
        "started",
        files.len(),
        files.len(),
        &["Rust", "excel-writer", "rust_xlsxwriter"],
        Some("rust_xlsxwriter"),
        None,
        "正在生成结果报表。",
    );
    let output = match report_writer(output_dir, &rows) {
        Ok(path) => path,
        Err(error) => {
            let rollback_errors = rollback_moves(&moved);
            if rollback_errors.is_empty() {
                emit_file_sort_event(
                    &mut on_event,
                    &mut sequence,
                    &started,
                    "PT-4-9",
                    "report_write",
                    "failed",
                    files.len(),
                    files.len(),
                    &["Rust", "excel-writer", "rust_xlsxwriter"],
                    Some("rust_xlsxwriter"),
                    Some("report_write_failed"),
                    "结果报表生成失败，已回滚文件改名。",
                );
                return Err(error).context("write file-sort-prefix report; file moves rolled back");
            }
            emit_file_sort_event(
                &mut on_event,
                &mut sequence,
                &started,
                "PT-4-9",
                "report_write",
                "failed",
                files.len(),
                files.len(),
                &["Rust", "excel-writer", "rust_xlsxwriter"],
                Some("rust_xlsxwriter"),
                Some("report_write_and_rollback_failed"),
                "结果报表生成及文件回滚失败。",
            );
            return Err(anyhow!(
                "write file-sort-prefix report failed: {error:#}; rollback also failed: {}",
                rollback_errors.join("; ")
            ));
        }
    };

    emit_file_sort_event(
        &mut on_event,
        &mut sequence,
        &started,
        "PT-4-9",
        "report_write",
        "completed",
        files.len(),
        files.len(),
        &["Rust", "excel-writer", "rust_xlsxwriter"],
        Some("rust_xlsxwriter"),
        None,
        "结果报表已生成。",
    );
    emit_file_sort_event(
        &mut on_event,
        &mut sequence,
        &started,
        "PT-4-9",
        "task",
        if failed == 0 {
            "completed"
        } else {
            "partial_failure"
        },
        files.len(),
        files.len(),
        &["Rust"],
        None,
        (failed > 0).then_some("one_or_more_files_failed"),
        format!(
            "完成：重命名 {} 个，未命中 {} 个，跳过 {} 个，失败 {} 个。",
            renamed, unmatched, skipped, failed
        ),
    );

    Ok(FileSortPrefixRunReport {
        output,
        rules: rules.len(),
        files: files.len(),
        renamed,
        unmatched,
        skipped,
        failed,
        rows,
    })
}

fn load_file_sort_rules(path: &Path) -> Result<Vec<FileSortRule>> {
    let workbook = read_workbook(path)
        .with_context(|| format!("read file-sort configuration {}", path.display()))?;
    let sheet = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == CONFIG_SHEET)
        .ok_or_else(|| anyhow!("missing worksheet {CONFIG_SHEET}"))?;
    let headers = header_columns(sheet);
    let rows = (2..=max_row(sheet))
        .map(|row_no| FileSortConfigRow {
            row_no,
            enabled: field_text(sheet, &headers, row_no, "是否启用"),
            sort_prefix: field_text(sheet, &headers, row_no, "排序前缀"),
            keywords: field_text(sheet, &headers, row_no, "匹配关键词"),
            remark: field_text(sheet, &headers, row_no, "备注"),
        })
        .collect::<Vec<_>>();
    build_file_sort_rules(&rows).map_err(|message| anyhow!(message))
}

fn header_columns(sheet: &SheetDump) -> BTreeMap<String, usize> {
    sheet
        .values
        .iter()
        .filter_map(|(address, value)| {
            let (row, column) = parse_cell_address(address)?;
            (row == 1).then(|| (cell_text(value), column))
        })
        .filter(|(name, _)| !name.is_empty())
        .collect()
}

fn field_text(
    sheet: &SheetDump,
    headers: &BTreeMap<String, usize>,
    row: usize,
    header: &str,
) -> String {
    headers
        .get(header)
        .and_then(|column| sheet.values.get(&cell_address(row, *column)))
        .map(cell_text)
        .unwrap_or_default()
}

fn max_row(sheet: &SheetDump) -> usize {
    sheet
        .values
        .keys()
        .filter_map(|address| parse_cell_address(address).map(|(row, _)| row))
        .max()
        .unwrap_or(1)
}

fn parse_cell_address(address: &str) -> Option<(usize, usize)> {
    let split = address.find(|character: char| character.is_ascii_digit())?;
    let (column, row) = address.split_at(split);
    let column = column.chars().try_fold(0_usize, |value, character| {
        let letter = character.to_ascii_uppercase();
        if !letter.is_ascii_uppercase() {
            return None;
        }
        value
            .checked_mul(26)?
            .checked_add((letter as u8 - b'A' + 1) as usize)
    })?;
    Some((row.parse().ok()?, column))
}

fn cell_address(row: usize, column: usize) -> String {
    let mut column = column;
    let mut label = String::new();
    while column > 0 {
        column -= 1;
        label.insert(0, (b'A' + (column % 26) as u8) as char);
        column /= 26;
    }
    format!("{label}{row}")
}

fn cell_text(value: &serde_json::Value) -> String {
    let text = match value["type"].as_str().unwrap_or_default() {
        "string" | "datetime_iso" | "duration_iso" => {
            value["value"].as_str().unwrap_or_default().to_owned()
        }
        "integer" | "number" => value["value"].to_string(),
        "boolean" => match value["value"].as_bool() {
            Some(true) => "True".to_owned(),
            Some(false) => "False".to_owned(),
            None => String::new(),
        },
        "datetime_serial" => value["value"].to_string(),
        "error" => value["value"].as_str().unwrap_or_default().to_owned(),
        _ => String::new(),
    };
    text.trim().to_owned()
}

fn content_rename_cell_text(value: &serde_json::Value, legacy_xls_value_coercion: bool) -> String {
    if legacy_xls_value_coercion
        && matches!(
            value["type"].as_str().unwrap_or_default(),
            "integer" | "number" | "datetime_serial"
        )
    {
        if let Some(number) = value["value"].as_f64() {
            return python_float_text(number);
        }
    }
    cell_text(value)
}

fn python_float_text(value: f64) -> String {
    if value.is_nan() {
        return "nan".to_owned();
    }
    if value.is_infinite() {
        return if value.is_sign_negative() {
            "-inf".to_owned()
        } else {
            "inf".to_owned()
        };
    }
    if value == 0.0 {
        return if value.is_sign_negative() {
            "-0.0".to_owned()
        } else {
            "0.0".to_owned()
        };
    }

    let magnitude = value.abs();
    if !(1e-4..1e16).contains(&magnitude) {
        let debug = format!("{value:?}");
        if let Some((mantissa, exponent)) = debug.split_once('e') {
            if let Ok(exponent) = exponent.parse::<i32>() {
                let sign = if exponent >= 0 { "+" } else { "-" };
                return format!("{mantissa}e{sign}{:02}", exponent.unsigned_abs());
            }
        }
        return debug;
    }

    let mut text = value.to_string();
    if !text.contains('.') {
        text.push_str(".0");
    }
    text
}

pub(crate) fn move_to_available_path(source: &Path, desired: &Path) -> io::Result<PathBuf> {
    let mut index = 0_usize;
    loop {
        let candidate = if index == 0 {
            desired.to_path_buf()
        } else {
            path_with_suffix(desired, index)
        };
        match move_file_no_replace(source, &candidate) {
            Ok(()) => return Ok(candidate),
            Err(error) if error.kind() == io::ErrorKind::AlreadyExists => index += 1,
            Err(error) => return Err(error),
        }
    }
}

fn move_file_no_replace(source: &Path, target: &Path) -> io::Result<()> {
    fs::hard_link(source, target)?;
    if let Err(remove_error) = fs::remove_file(source) {
        return match fs::remove_file(target) {
            Ok(()) => Err(remove_error),
            Err(cleanup_error) => Err(io::Error::new(
                cleanup_error.kind(),
                format!("{remove_error}; could not remove reserved destination: {cleanup_error}"),
            )),
        };
    }
    Ok(())
}

fn rollback_moves(moved: &[(PathBuf, PathBuf)]) -> Vec<String> {
    moved
        .iter()
        .rev()
        .filter_map(|(source, target)| {
            move_file_no_replace(target, source)
                .err()
                .map(|error| format!("{} -> {}: {error}", target.display(), source.display()))
        })
        .collect()
}

fn path_with_suffix(path: &Path, index: usize) -> PathBuf {
    let stem = path.file_stem().unwrap_or_default().to_string_lossy();
    let extension = path.extension().map(|value| value.to_string_lossy());
    let file_name = match extension {
        Some(extension) if !extension.is_empty() => format!("{stem}_{index}.{extension}"),
        _ => format!("{stem}_{index}"),
    };
    path.with_file_name(file_name)
}

fn write_report(output_dir: &Path, rows: &[FileSortPrefixRow]) -> Result<PathBuf> {
    let timestamp = Local::now().format("%Y%m%d_%H%M%S");
    let base = output_dir.join(format!("{timestamp}_文件排序前缀结果.xlsx"));
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .context("system clock is before Unix epoch")?
        .as_nanos();
    let temporary = output_dir.join(format!(
        ".文件排序前缀结果-{}-{nonce}.tmp.xlsx",
        std::process::id()
    ));
    OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&temporary)
        .with_context(|| format!("reserve temporary report {}", temporary.display()))?;

    let result = (|| -> Result<PathBuf> {
        let mut workbook = Workbook::new();
        {
            let worksheet = workbook.add_worksheet();
            worksheet.set_name(REPORT_SHEET)?;
            for (column, header) in REPORT_HEADERS.iter().enumerate() {
                worksheet.write_string(0, column as u16, *header)?;
            }
            for (index, row) in rows.iter().enumerate() {
                let excel_row = index as u32 + 1;
                worksheet.write_string(excel_row, 0, &row.source_path)?;
                worksheet.write_string(excel_row, 1, &row.original_name)?;
                worksheet.write_string(excel_row, 2, &row.new_name)?;
                if let Some(config_row) = row.config_row {
                    worksheet.write_number(excel_row, 3, config_row as f64)?;
                }
                worksheet.write_string(excel_row, 4, &row.sort_prefix)?;
                worksheet.write_string(excel_row, 5, &row.matched_keywords)?;
                worksheet.write_string(excel_row, 6, &row.status)?;
                worksheet.write_string(excel_row, 7, &row.remark)?;
            }
        }
        workbook
            .save(&temporary)
            .with_context(|| format!("write temporary report {}", temporary.display()))?;
        commit_report(&temporary, &base)
    })();

    if result.is_err() {
        let _ = fs::remove_file(&temporary);
    }
    result
}

fn write_remove_report(output_dir: &Path, rows: &[FileSortPrefixRemovalRow]) -> Result<PathBuf> {
    let timestamp = Local::now().format("%Y%m%d_%H%M%S");
    let base = output_dir.join(format!("{timestamp}_{REMOVE_REPORT_SHEET}.xlsx"));
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .context("system clock is before Unix epoch")?
        .as_nanos();
    let temporary = output_dir.join(format!(
        ".{REMOVE_REPORT_SHEET}-{}-{nonce}.tmp.xlsx",
        std::process::id()
    ));
    OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&temporary)
        .with_context(|| format!("reserve temporary report {}", temporary.display()))?;

    let result = (|| -> Result<PathBuf> {
        let mut workbook = Workbook::new();
        {
            let worksheet = workbook.add_worksheet();
            worksheet.set_name(REMOVE_REPORT_SHEET)?;
            if !rows.is_empty() {
                for (column, header) in REMOVE_REPORT_HEADERS.iter().enumerate() {
                    worksheet.write_string(0, column as u16, *header)?;
                }
            }
            for (index, row) in rows.iter().enumerate() {
                let excel_row = index as u32 + 1;
                worksheet.write_string(excel_row, 0, &row.source_path)?;
                worksheet.write_string(excel_row, 1, &row.original_name)?;
                worksheet.write_string(excel_row, 2, &row.new_name)?;
                worksheet.write_string(excel_row, 3, &row.status)?;
                worksheet.write_string(excel_row, 4, &row.remark)?;
            }
        }
        workbook
            .save(&temporary)
            .with_context(|| format!("write temporary report {}", temporary.display()))?;
        commit_report(&temporary, &base)
    })();

    if result.is_err() {
        let _ = fs::remove_file(&temporary);
    }
    result
}

fn commit_report(temporary: &Path, base: &Path) -> Result<PathBuf> {
    let mut index = 0_usize;
    loop {
        let candidate = if index == 0 {
            base.to_path_buf()
        } else {
            path_with_suffix(base, index)
        };
        match fs::hard_link(temporary, &candidate) {
            Ok(()) => {
                if let Err(error) = fs::remove_file(temporary) {
                    return match fs::remove_file(&candidate) {
                        Ok(()) => Err(error).with_context(|| {
                            format!("remove temporary report {}", temporary.display())
                        }),
                        Err(cleanup_error) => Err(anyhow!(
                            "published report {} but could not clean temporary {} ({error}); report cleanup also failed ({cleanup_error})",
                            candidate.display(),
                            temporary.display()
                        )),
                    };
                }
                return Ok(candidate);
            }
            Err(error) if error.kind() == io::ErrorKind::AlreadyExists => index += 1,
            Err(error) => {
                return Err(error).with_context(|| {
                    format!("publish report without replacing {}", candidate.display())
                })
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn temp_dir() -> PathBuf {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let path = std::env::temp_dir().join(format!(
            "pytools-sort-prefix-rollback-{}-{nonce}",
            std::process::id()
        ));
        fs::create_dir_all(&path).unwrap();
        path
    }

    #[test]
    fn collision_move_does_not_overwrite_and_rollback_restores_source() {
        let root = temp_dir();
        let source = root.join("source.xlsx");
        let target = root.join("target.xlsx");
        fs::write(&source, b"source payload").unwrap();
        fs::write(&target, b"pre-existing target").unwrap();

        let moved_to = move_to_available_path(&source, &target).unwrap();
        let expected = root.join("target_1.xlsx");
        assert_eq!(moved_to, expected);
        assert!(!source.exists());
        assert_eq!(fs::read(&target).unwrap(), b"pre-existing target");
        assert_eq!(fs::read(&expected).unwrap(), b"source payload");

        assert!(rollback_moves(&[(source.clone(), expected.clone())]).is_empty());
        assert_eq!(fs::read(&source).unwrap(), b"source payload");
        assert!(!expected.exists());
        assert_eq!(fs::read(&target).unwrap(), b"pre-existing target");
        fs::remove_dir_all(root).unwrap();
    }

    #[test]
    fn report_write_failure_rolls_back_completed_file_moves() {
        let root = temp_dir();
        let source = root.join("202506_建设银行_存款表.xlsx");
        fs::write(&source, b"must be restored").unwrap();
        let output = root.join("output");
        let config = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/file_sort_prefix_config.xlsx");

        let error = run_file_sort_prefix_with_report_writer(
            &config,
            std::slice::from_ref(&source),
            &output,
            |_output_dir, _rows| bail!("injected report write failure"),
        )
        .unwrap_err();
        assert!(format!("{error:#}").contains("file moves rolled back"));
        assert_eq!(fs::read(&source).unwrap(), b"must be restored");
        assert!(!root.join("002_202506_建设银行_存款表.xlsx").exists());
        fs::remove_dir_all(root).unwrap();
    }

    #[test]
    fn removal_report_write_failure_rolls_back_completed_file_moves() {
        let root = temp_dir();
        let source = root.join("001_报表.xlsx");
        fs::write(&source, b"must be restored").unwrap();
        let output = root.join("output");
        let config = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/file_sort_prefix_config.xlsx");

        let error = run_remove_file_sort_prefix_with_report_writer(
            &config,
            std::slice::from_ref(&source),
            &output,
            |_output_dir, _rows| bail!("injected report write failure"),
        )
        .unwrap_err();
        assert!(format!("{error:#}").contains("file moves rolled back"));
        assert_eq!(fs::read(&source).unwrap(), b"must be restored");
        assert!(!root.join("报表.xlsx").exists());
        fs::remove_dir_all(root).unwrap();
    }

    #[test]
    fn removal_content_prefix_helpers_match_source_behavior() {
        assert_eq!(sanitize_name_part("  <机构A>  / 报表.  "), "_机构A_ _ 报表");
        assert_eq!(
            remove_content_prefix("A_B_A_file.xlsx", &["A".to_owned(), "B".to_owned()]),
            Some("file.xlsx".to_owned())
        );
        assert_eq!(remove_content_prefix("file.xlsx", &["A".to_owned()]), None);
    }

    #[test]
    fn legacy_xls_content_numbers_match_com_float_text() {
        assert_eq!(
            content_rename_cell_text(&serde_json::json!({"type":"integer","value":42}), true),
            "42.0"
        );
        assert_eq!(
            content_rename_cell_text(&serde_json::json!({"type":"number","value":42}), true),
            "42.0"
        );
        assert_eq!(
            content_rename_cell_text(
                &serde_json::json!({"type":"datetime_serial","value":42}),
                true
            ),
            "42.0"
        );
        assert_eq!(
            content_rename_cell_text(&serde_json::json!({"type":"integer","value":42}), false),
            "42"
        );
        assert_eq!(python_float_text(-0.0), "-0.0");
        assert_eq!(python_float_text(1e16), "1e+16");
        assert_eq!(python_float_text(1e-5), "1e-05");
        assert_eq!(python_float_text(1e-4), "0.0001");
    }
}

/// Port of Python `_sort_prefix_digits`.
pub fn sort_prefix_digits(value: &str) -> Result<String, String> {
    let raw = value.trim();
    if raw.is_empty() {
        return Ok(String::new());
    }

    if is_numeric_prefix(raw) {
        let ascii = decimal_digits_to_ascii(raw);
        let number = ascii
            .parse::<f64>()
            .map_err(|_| format!("invalid numeric sort prefix: {raw}"))?;
        if !number.is_finite() {
            return Err(format!("numeric sort prefix is out of range: {raw}"));
        }
        return Ok(format!("{number:.0}"));
    }

    Ok(raw
        .chars()
        .filter(|character| decimal_digit_value(*character).is_some())
        .collect())
}

/// Port of Python `_sort_prefix_width`.
pub fn sort_prefix_width(value: &str) -> usize {
    let raw = value.trim();
    if raw.is_empty() {
        return 0;
    }

    if is_numeric_prefix(raw) {
        return raw.split_once('.').map_or_else(
            || raw.chars().count(),
            |(integer, _)| integer.chars().count(),
        );
    }

    raw.chars()
        .filter(|character| decimal_digit_value(*character).is_some())
        .count()
}

/// Port of Python `_normalize_sort_prefix`.
pub fn normalize_sort_prefix(value: &str, width: usize) -> Result<String, String> {
    let digits = sort_prefix_digits(value)?;
    if digits.is_empty() {
        return Ok(String::new());
    }

    let ascii_digits = decimal_digits_to_ascii(&digits);
    let normalized = ascii_digits.trim_start_matches('0');
    let normalized = if normalized.is_empty() {
        "0"
    } else {
        normalized
    };
    let target_width = width.max(1);
    if normalized.len() >= target_width {
        return Ok(normalized.to_owned());
    }

    let mut output = String::with_capacity(target_width);
    output.extend(std::iter::repeat('0').take(target_width - normalized.len()));
    output.push_str(normalized);
    Ok(output)
}

/// Parse enabled rows and normalize their prefixes to the widest active rule.
pub fn build_file_sort_rules(rows: &[FileSortConfigRow]) -> Result<Vec<FileSortRule>, String> {
    let mut width = 1;
    for row in rows {
        if !is_enabled(&row.enabled) || split_tokens(&row.keywords).is_empty() {
            continue;
        }
        width = width.max(sort_prefix_width(&row.sort_prefix));
    }

    let mut rules = Vec::new();
    for row in rows {
        if !is_enabled(&row.enabled) {
            continue;
        }

        let prefix = normalize_sort_prefix(&row.sort_prefix, width)?;
        let keywords = split_tokens(&row.keywords);
        if prefix.is_empty() || keywords.is_empty() {
            continue;
        }

        rules.push(FileSortRule {
            row_no: row.row_no,
            prefix,
            keywords,
            keywords_text: row.keywords.trim().to_owned(),
            remark: row.remark.trim().to_owned(),
        });
    }
    Ok(rules)
}

/// Return the first enabled rule whose first or later keyword occurs in the name.
pub fn match_file_sort_rule<'a>(
    file_name: &str,
    rules: &'a [FileSortRule],
) -> Option<&'a FileSortRule> {
    let lower_name = file_name.to_lowercase();
    rules.iter().find(|rule| {
        rule.keywords
            .iter()
            .any(|keyword| lower_name.contains(&keyword.to_lowercase()))
    })
}

/// Port of Python `_strip_sort_prefix`; only a leading 1–3 digit marker is removed.
pub fn strip_sort_prefix(file_name: &str) -> String {
    let name = file_name.rsplit(['/', '\\']).next().unwrap_or(file_name);
    let (stem, suffix) = split_stem_suffix(name);
    format!("{}{}", strip_leading_marker(stem), suffix)
}

/// Produce the target name for operation 4-9 after removing any existing marker.
pub fn prefixed_file_name(file_name: &str, prefix: &str) -> String {
    format!("{prefix}_{}", strip_sort_prefix(file_name))
}

fn split_stem_suffix(file_name: &str) -> (&str, &str) {
    match file_name.rfind('.') {
        Some(index) if index > 0 && index + 1 < file_name.len() => {
            (&file_name[..index], &file_name[index..])
        }
        _ => (file_name, ""),
    }
}

fn strip_leading_marker(stem: &str) -> &str {
    let mut digit_end = 0;
    let mut digit_count = 0;
    for (index, character) in stem.char_indices() {
        if decimal_digit_value(character).is_some() && digit_count < 3 {
            digit_count += 1;
            digit_end = index + character.len_utf8();
        } else {
            break;
        }
    }

    if digit_count == 0 {
        return stem;
    }

    let mut separator_end = digit_end;
    for (offset, character) in stem[digit_end..].char_indices() {
        if is_python_whitespace(character) || character == '_' || character == '-' {
            separator_end = digit_end + offset + character.len_utf8();
        } else {
            break;
        }
    }

    if separator_end == digit_end {
        stem
    } else {
        &stem[separator_end..]
    }
}

fn split_tokens(value: &str) -> Vec<String> {
    value
        .replace('；', ";")
        .replace(',', ";")
        .split(';')
        .map(str::trim)
        .filter(|token| !token.is_empty())
        .map(str::to_owned)
        .collect()
}

fn is_enabled(value: &str) -> bool {
    matches!(
        value.trim().to_lowercase().as_str(),
        "1" | "y" | "yes" | "true" | "是" | "启用"
    )
}

fn is_numeric_prefix(value: &str) -> bool {
    if value
        .chars()
        .all(|character| decimal_digit_value(character).is_some())
    {
        return !value.is_empty();
    }

    match value.split_once('.') {
        Some((integer, fraction)) => {
            !integer.is_empty()
                && integer
                    .chars()
                    .all(|character| decimal_digit_value(character).is_some())
                && !fraction.is_empty()
                && fraction.chars().all(|character| character == '0')
        }
        None => false,
    }
}

fn decimal_digits_to_ascii(value: &str) -> String {
    let mut output = String::with_capacity(value.len());
    for character in value.chars() {
        match decimal_digit_value(character) {
            Some(digit) => output.push(char::from(b'0' + digit)),
            None => output.push(character),
        }
    }
    output
}

fn decimal_digit_value(character: char) -> Option<u8> {
    const DECIMAL_ZEROES_UNICODE_15: &[u32] = &[
        0x0030, 0x0660, 0x06f0, 0x07c0, 0x0966, 0x09e6, 0x0a66, 0x0ae6, 0x0b66, 0x0be6, 0x0c66,
        0x0ce6, 0x0d66, 0x0de6, 0x0e50, 0x0ed0, 0x0f20, 0x1040, 0x1090, 0x17e0, 0x1810, 0x1946,
        0x19d0, 0x1a80, 0x1a90, 0x1b50, 0x1bb0, 0x1c40, 0x1c50, 0xa620, 0xa8d0, 0xa900, 0xa9d0,
        0xa9f0, 0xaa50, 0xabf0, 0xff10, 0x104a0, 0x10d30, 0x11066, 0x110f0, 0x11136, 0x111d0,
        0x112f0, 0x11450, 0x114d0, 0x11650, 0x116c0, 0x11730, 0x118e0, 0x11950, 0x11c50, 0x11d50,
        0x11da0, 0x11f50, 0x16a60, 0x16ac0, 0x16b50, 0x1d7ce, 0x1d7d8, 0x1d7e2, 0x1d7ec, 0x1d7f6,
        0x1e140, 0x1e2f0, 0x1e4f0, 0x1e950, 0x1fbf0,
    ];
    let code_point = character as u32;
    DECIMAL_ZEROES_UNICODE_15.iter().find_map(|zero| {
        let digit = code_point.checked_sub(*zero)?;
        (digit < 10).then_some(digit as u8)
    })
}

fn is_python_whitespace(character: char) -> bool {
    character.is_whitespace() || matches!(character, '\u{001c}'..='\u{001f}')
}

/// Compatibility alias for the original Python 4-6 operation identifier.
#[deprecated(note = "use `run_content_prefix_rename_xlsx_candidate`")]
pub fn run_content_prefix_rename_46_xlsx_candidate(
    config_path: impl AsRef<Path>,
    files: &[PathBuf],
) -> Result<ContentPrefixRenameRunReport> {
    run_content_prefix_rename_xlsx_candidate(config_path, files)
}

/// Compatibility alias for the original Python 4-7 operation identifier.
#[deprecated(note = "use `run_content_prefix_unrename_xlsx_candidate`")]
pub fn run_content_prefix_unrename_47_xlsx_candidate(
    config_path: impl AsRef<Path>,
    files: &[PathBuf],
) -> Result<ContentPrefixUnrenameRunReport> {
    run_content_prefix_unrename_xlsx_candidate(config_path, files)
}

/// Compatibility alias for the original Python 4-9 operation identifier.
#[deprecated(note = "use `run_file_sort_prefix`")]
pub fn run_file_sort_prefix_49(
    config_path: impl AsRef<Path>,
    files: &[PathBuf],
    output_dir: impl AsRef<Path>,
) -> Result<FileSortPrefixRunReport> {
    run_file_sort_prefix(config_path, files, output_dir)
}

/// Compatibility alias for the original Python 4-10 operation identifier.
#[deprecated(note = "use `run_remove_file_sort_prefix`")]
pub fn run_remove_file_sort_prefix_410(
    config_path: impl AsRef<Path>,
    files: &[PathBuf],
    output_dir: impl AsRef<Path>,
) -> Result<FileSortPrefixRemovalRunReport> {
    run_remove_file_sort_prefix(config_path, files, output_dir)
}
