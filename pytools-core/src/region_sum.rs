//! Candidate Rust implementation of PyCode's read-only regional total check.
//!
//! The source feature reads stored worksheet values and creates a new summary
//! workbook. It does not modify the input workbooks. The Rust implementation
//! is available through the `pytools region-sum-check` CLI.

use anyhow::{bail, ensure, Context, Result};
use chrono::Local;
use excel_reader::{read_value_columns, CellValue, SheetValues};
use excel_writer::{Color, Format, FormatAlign, Workbook};
use serde::Serialize;
use std::collections::BTreeMap;
use std::fs::{self, File, OpenOptions};
use std::io::{BufWriter, Write};
use std::path::{Path, PathBuf};
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};

const REGION_SHEET_KEY: &str = "分地区";
const TOTAL_REGION: &str = "惠州市";
const SUB_REGIONS: [&str; 6] = [
    "惠阳区",
    "博罗县",
    "惠东县",
    "龙门县",
    "惠州市本部",
    "大亚湾",
];
const ABS_TOLERANCE: f64 = 0.01;
const REL_TOLERANCE: f64 = 0.001;
const RESULT_SHEET: &str = "核对结果";
const RESULT_HEADERS: [&str; 6] = [
    "序号",
    "文件名",
    "工作表名",
    "核对状态",
    "错误数量",
    "错误信息",
];
const MAX_XLSX_COLUMNS: usize = 16_384;

#[derive(Clone, Debug, PartialEq, Eq)]
struct SheetCheck {
    error_count: usize,
    detail: String,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct ResultRow {
    sequence: usize,
    file_name: String,
    sheet_name: String,
    status: String,
    error_count: usize,
    detail: String,
}

/// Aggregate returned by `run_region_sum_check`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RegionSumRun {
    pub files: usize,
    pub sheets: usize,
    pub errors: usize,
    pub failed_files: usize,
    pub path: Option<PathBuf>,
    pub log_path: PathBuf,
}

#[derive(Clone, Debug, Default, Serialize)]
pub struct RegionSumProgress {
    pub total_files: usize,
    pub processed_files: usize,
    pub completed_files: usize,
    pub failed_files: usize,
    pub matched_sheets: usize,
    pub errors: usize,
}

#[derive(Clone, Debug, Serialize)]
pub struct RegionSumRunEvent {
    pub sequence: usize,
    pub timestamp: String,
    pub product: String,
    pub feature_id: String,
    pub function: String,
    pub run_id: String,
    pub phase: String,
    pub status: String,
    pub user_message: String,
    pub progress: RegionSumProgress,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub input_index: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stage_elapsed_ms: Option<u128>,
    pub total_elapsed_ms: u128,
    pub components: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub actual_engine: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_type: Option<String>,
}

struct RegionSumProgressEmitter<'a> {
    started: Instant,
    sequence: usize,
    run_id: String,
    progress: RegionSumProgress,
    callback: Option<&'a mut dyn FnMut(RegionSumRunEvent)>,
}

impl RegionSumProgressEmitter<'_> {
    fn emit(
        &mut self,
        phase: &str,
        status: &str,
        user_message: impl Into<String>,
        input_index: Option<usize>,
        stage_elapsed: Option<Duration>,
        components: &[&str],
        actual_engine: Option<&str>,
        error_type: Option<&str>,
    ) {
        if self.callback.is_none() {
            return;
        }
        self.sequence += 1;
        let mut recorded_components = Vec::with_capacity(components.len() + 1);
        recorded_components.push("Rust".to_owned());
        recorded_components.extend(components.iter().map(|component| (*component).to_owned()));
        let event = RegionSumRunEvent {
            sequence: self.sequence,
            timestamp: Local::now().to_rfc3339(),
            product: "pytools".to_owned(),
            feature_id: "PT-7-5".to_owned(),
            function: "region-sum-check".to_owned(),
            run_id: self.run_id.clone(),
            phase: phase.to_owned(),
            status: status.to_owned(),
            user_message: user_message.into(),
            progress: self.progress.clone(),
            input_index,
            stage_elapsed_ms: stage_elapsed.map(|elapsed| elapsed.as_millis()),
            total_elapsed_ms: self.started.elapsed().as_millis(),
            components: recorded_components,
            actual_engine: actual_engine.map(str::to_owned),
            error_type: error_type.map(str::to_owned),
        };
        if let Some(callback) = self.callback.as_mut() {
            callback(event);
        }
    }
}

/// Run the standard feature using the current local time for output names.
pub fn run_region_sum_check<I, P>(
    target_workbooks: I,
    output_dir: impl AsRef<Path>,
    log_dir: impl AsRef<Path>,
) -> Result<RegionSumRun>
where
    I: IntoIterator<Item = P>,
    P: AsRef<Path>,
{
    let now = Local::now();
    let file_stamp = now.format("%Y%m%d_%H%M%S").to_string();
    let log_stamp = now.format("%Y-%m-%d %H:%M:%S").to_string();
    run_region_sum_check_with_timestamp(
        target_workbooks,
        output_dir,
        log_dir,
        &file_stamp,
        &log_stamp,
    )
}

/// Deterministic entry used by probes and tests; the timestamp must be
/// `YYYYMMDD_HHMMSS`. Product callers should use `run_region_sum_check`.
pub fn run_region_sum_check_with_timestamp<I, P>(
    target_workbooks: I,
    output_dir: impl AsRef<Path>,
    log_dir: impl AsRef<Path>,
    file_stamp: &str,
    log_stamp: &str,
) -> Result<RegionSumRun>
where
    I: IntoIterator<Item = P>,
    P: AsRef<Path>,
{
    run_region_sum_check_inner(
        target_workbooks,
        output_dir,
        log_dir,
        file_stamp,
        log_stamp,
        None,
    )
}

pub fn run_region_sum_check_with_progress<I, P, F>(
    target_workbooks: I,
    output_dir: impl AsRef<Path>,
    log_dir: impl AsRef<Path>,
    on_event: F,
) -> Result<RegionSumRun>
where
    I: IntoIterator<Item = P>,
    P: AsRef<Path>,
    F: FnMut(RegionSumRunEvent),
{
    let now = Local::now();
    let file_stamp = now.format("%Y%m%d_%H%M%S").to_string();
    let log_stamp = now.format("%Y-%m-%d %H:%M:%S").to_string();
    run_region_sum_check_with_timestamp_and_progress(
        target_workbooks,
        output_dir,
        log_dir,
        &file_stamp,
        &log_stamp,
        on_event,
    )
}

pub fn run_region_sum_check_with_timestamp_and_progress<I, P, F>(
    target_workbooks: I,
    output_dir: impl AsRef<Path>,
    log_dir: impl AsRef<Path>,
    file_stamp: &str,
    log_stamp: &str,
    mut on_event: F,
) -> Result<RegionSumRun>
where
    I: IntoIterator<Item = P>,
    P: AsRef<Path>,
    F: FnMut(RegionSumRunEvent),
{
    run_region_sum_check_inner(
        target_workbooks,
        output_dir,
        log_dir,
        file_stamp,
        log_stamp,
        Some(&mut on_event),
    )
}

fn run_region_sum_check_inner<I, P>(
    target_workbooks: I,
    output_dir: impl AsRef<Path>,
    log_dir: impl AsRef<Path>,
    file_stamp: &str,
    log_stamp: &str,
    callback: Option<&mut dyn FnMut(RegionSumRunEvent)>,
) -> Result<RegionSumRun>
where
    I: IntoIterator<Item = P>,
    P: AsRef<Path>,
{
    validate_timestamp(file_stamp)?;
    let target_workbooks = target_workbooks
        .into_iter()
        .map(|input| input.as_ref().to_path_buf())
        .collect::<Vec<_>>();
    let output_dir = output_dir.as_ref();
    let log_dir = log_dir.as_ref();
    let started = Instant::now();
    let mut emitter = RegionSumProgressEmitter {
        started,
        sequence: 0,
        run_id: region_sum_run_id(),
        progress: RegionSumProgress {
            total_files: target_workbooks.len(),
            ..Default::default()
        },
        callback,
    };
    emitter.emit(
        "run",
        "started",
        format!(
            "开始地区总分校验，共 {} 个工作簿",
            emitter.progress.total_files
        ),
        None,
        None,
        &[],
        None,
        None,
    );

    if let Err(error) = fs::create_dir_all(log_dir) {
        emitter.emit(
            "log_setup",
            "failed",
            "运行日志目录不可用，任务未启动",
            None,
            None,
            &[],
            None,
            Some("log_directory_create_failed"),
        );
        return Err(error).with_context(|| format!("create log directory {}", log_dir.display()));
    }
    let (log_path, initial_output_path) = match unique_paths(output_dir, log_dir, file_stamp) {
        Ok(paths) => paths,
        Err(error) => {
            emitter.emit(
                "output_setup",
                "failed",
                "无法分配核对结果文件名",
                None,
                None,
                &[],
                None,
                Some("output_name_allocation_failed"),
            );
            return Err(error);
        }
    };
    let log_file = match File::create(&log_path) {
        Ok(file) => file,
        Err(error) => {
            emitter.emit(
                "log_setup",
                "failed",
                "无法创建运行日志",
                None,
                None,
                &[],
                None,
                Some("legacy_log_create_failed"),
            );
            return Err(error).with_context(|| format!("create log {}", log_path.display()));
        }
    };
    let mut log = BufWriter::new(log_file);
    let mut result_rows = Vec::new();

    for (index, source) in target_workbooks.iter().enumerate() {
        let input_index = index + 1;
        if !openpyxl_compatible_extension(source) {
            emitter.progress.processed_files += 1;
            emitter.progress.failed_files += 1;
            write_log(
                &mut log,
                log_stamp,
                "ERROR",
                &format!(
                    "打开失败: {} - unsupported workbook extension",
                    source.display()
                ),
            );
            emitter.emit(
                "input_validation",
                "failed",
                format!("第 {input_index} 个工作簿格式不受支持，已跳过"),
                Some(input_index),
                None,
                &[],
                None,
                Some("unsupported_workbook_format"),
            );
            continue;
        }

        emitter.emit(
            "workbook_read",
            "started",
            format!(
                "正在读取第 {input_index}/{} 个工作簿",
                emitter.progress.total_files
            ),
            Some(input_index),
            None,
            &[],
            None,
            None,
        );
        let read_started = Instant::now();
        let sheets = match read_value_columns(source, MAX_XLSX_COLUMNS) {
            Ok(sheets) => sheets,
            Err(error) => {
                emitter.progress.processed_files += 1;
                emitter.progress.failed_files += 1;
                write_log(
                    &mut log,
                    log_stamp,
                    "ERROR",
                    &format!("打开失败: {} - {error:#}", source.display()),
                );
                emitter.emit(
                    "workbook_read",
                    "failed",
                    format!("第 {input_index} 个工作簿读取失败，已跳过"),
                    Some(input_index),
                    Some(read_started.elapsed()),
                    &["calamine"],
                    Some("calamine"),
                    Some("workbook_read_failed"),
                );
                continue;
            }
        };
        emitter.emit(
            "workbook_read",
            "completed",
            format!(
                "已读取第 {input_index}/{} 个工作簿",
                emitter.progress.total_files
            ),
            Some(input_index),
            Some(read_started.elapsed()),
            &["calamine"],
            Some("calamine"),
            None,
        );

        let check_started = Instant::now();
        let mut file_sheets = 0usize;
        let mut file_errors = 0usize;
        let file_name = source
            .file_name()
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_default();

        for sheet in &sheets {
            if !sheet.name.contains(REGION_SHEET_KEY) {
                continue;
            }
            let check = check_one_sheet(sheet);
            emitter.progress.matched_sheets += 1;
            emitter.progress.errors += check.error_count;
            file_sheets += 1;
            file_errors += check.error_count;
            result_rows.push(make_result_row(
                result_rows.len() + 1,
                &file_name,
                &sheet.name,
                check,
            ));
        }
        write_log(
            &mut log,
            log_stamp,
            "INFO",
            &format!("{file_name}: 校验 {file_sheets} 个分地区 sheet，错误 {file_errors}"),
        );
        emitter.progress.processed_files += 1;
        emitter.progress.completed_files += 1;
        emitter.emit(
            "region_check",
            "completed",
            format!("第 {input_index} 个工作簿地区核对完成"),
            Some(input_index),
            Some(check_started.elapsed()),
            &[],
            None,
            None,
        );
    }

    emitter.emit(
        "report_write",
        "started",
        "正在生成地区总分核对结果",
        None,
        None,
        &[],
        None,
        None,
    );
    let write_started = Instant::now();
    if let Err(error) = fs::create_dir_all(output_dir) {
        emitter.emit(
            "report_write",
            "failed",
            "结果目录不可用，未生成核对报表",
            None,
            Some(write_started.elapsed()),
            &[],
            None,
            Some("output_directory_create_failed"),
        );
        return Err(error)
            .with_context(|| format!("create output directory {}", output_dir.display()));
    }
    let files = emitter.progress.completed_files;
    let sheets_checked = emitter.progress.matched_sheets;
    let errors = emitter.progress.errors;
    let output_path = match save_report(
        output_dir,
        initial_output_path,
        &result_rows,
        files,
        sheets_checked,
        errors,
    ) {
        Ok(path) => path,
        Err(error) => {
            emitter.emit(
                "report_write",
                "failed",
                "生成核对报表失败",
                None,
                Some(write_started.elapsed()),
                &[],
                None,
                Some("report_write_failed"),
            );
            return Err(error);
        }
    };
    let mut report_components = vec!["rust_xlsxwriter"];
    if !result_rows.is_empty() {
        report_components.push("ooxml-patch");
    }
    emitter.emit(
        "report_write",
        "completed",
        "地区总分核对报表已生成",
        None,
        Some(write_started.elapsed()),
        &report_components,
        Some("rust_xlsxwriter"),
        None,
    );
    write_log(
        &mut log,
        log_stamp,
        "INFO",
        &format!("结果已保存: {}", output_path.display()),
    );
    if let Err(error) = log.flush() {
        emitter.emit(
            "legacy_log_write",
            "failed",
            "运行日志写入失败",
            None,
            None,
            &["RustFilesystem"],
            Some("RustFilesystem"),
            Some("legacy_log_flush_failed"),
        );
        return Err(error).context("flush feature log");
    }

    let run = RegionSumRun {
        files,
        sheets: sheets_checked,
        errors,
        failed_files: emitter.progress.failed_files,
        path: Some(output_path),
        log_path,
    };
    emitter.emit(
        "run",
        "completed",
        format!(
            "地区总分校验完成：成功 {} 个，失败 {} 个，核对 {} 张工作表，发现 {} 项错误",
            run.files, run.failed_files, run.sheets, run.errors
        ),
        None,
        Some(started.elapsed()),
        &[],
        None,
        None,
    );
    Ok(run)
}

fn region_sum_run_id() -> String {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_nanos())
        .unwrap_or_default();
    format!("pytools-pt-7-5-{}-{nonce}", std::process::id())
}

fn validate_timestamp(value: &str) -> Result<()> {
    ensure!(
        value.len() == 15
            && value.as_bytes().get(8) == Some(&b'_')
            && value
                .bytes()
                .enumerate()
                .all(|(index, byte)| index == 8 || byte.is_ascii_digit()),
        "timestamp must use YYYYMMDD_HHMMSS"
    );
    Ok(())
}

fn openpyxl_compatible_extension(path: &Path) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| {
            matches!(
                extension.to_ascii_lowercase().as_str(),
                "xlsx" | "xlsm" | "xltx" | "xltm"
            )
        })
}

fn unique_paths(output_dir: &Path, log_dir: &Path, stamp: &str) -> Result<(PathBuf, PathBuf)> {
    for suffix in 0u32..10_000 {
        let suffix = if suffix == 0 {
            String::new()
        } else {
            format!("_{suffix}")
        };
        let output = output_dir.join(format!("地区数值核对结果_{stamp}{suffix}.xlsx"));
        let log = log_dir.join(format!("feature_1_5_region_sum_check_{stamp}{suffix}.log"));
        if !output.exists() && !log.exists() {
            return Ok((log, output));
        }
    }
    bail!("could not allocate a unique regional-check output name")
}

fn write_log(log: &mut BufWriter<File>, timestamp: &str, level: &str, message: &str) {
    let _ = writeln!(log, "{timestamp} [{level}] {message}");
}

fn make_result_row(
    sequence: usize,
    file_name: &str,
    sheet_name: &str,
    check: SheetCheck,
) -> ResultRow {
    let status = if check.error_count > 0 {
        "错误"
    } else {
        "正确"
    }
    .to_owned();
    let detail = if check.error_count > 0 {
        check.detail
    } else {
        "核对通过".to_owned()
    };
    ResultRow {
        sequence,
        file_name: file_name.to_owned(),
        sheet_name: sheet_name.to_owned(),
        status,
        error_count: check.error_count,
        detail,
    }
}

fn check_one_sheet(sheet: &SheetValues) -> SheetCheck {
    if sheet.max_row == 0 || sheet.max_column == 0 {
        return SheetCheck {
            error_count: 1,
            detail: "工作表为空或没有数据".to_owned(),
        };
    }

    let mut rows = BTreeMap::<String, usize>::new();
    for row in 1..=sheet.max_row {
        let Some(value) = cell_at(sheet, row, 2) else {
            continue;
        };
        let Some(name) = python_text(value) else {
            continue;
        };
        let name = name.trim();
        if (name == TOTAL_REGION || SUB_REGIONS.contains(&name)) && !rows.contains_key(name) {
            rows.insert(name.to_owned(), row);
        }
    }

    let Some(&total_row) = rows.get(TOTAL_REGION) else {
        return SheetCheck {
            error_count: 1,
            detail: format!("未找到'{TOTAL_REGION}'行"),
        };
    };
    if sheet.max_column < 3 {
        return SheetCheck {
            error_count: 1,
            detail: "未找到有效数据区域".to_owned(),
        };
    }

    let mut error_details = Vec::new();
    let mut numeric_error_count = 0usize;
    for column in 3..=sheet.max_column {
        let Some(total_value) = number_or_none(cell_at(sheet, total_row, column)) else {
            continue;
        };

        let mut sub_sum = 0.0;
        let mut any_sub_present = false;
        for name in SUB_REGIONS {
            let Some(&row) = rows.get(name) else {
                continue;
            };
            if let Some(value) = number_or_none(cell_at(sheet, row, column)) {
                sub_sum += value;
                any_sub_present = true;
            }
        }
        if !any_sub_present {
            continue;
        }

        let difference = total_value - sub_sum;
        let significant = if difference.abs() > ABS_TOLERANCE {
            true
        } else if total_value != 0.0 {
            (difference / total_value).abs() > REL_TOLERANCE
        } else {
            sub_sum != 0.0
        };
        if significant {
            numeric_error_count += 1;
            let header = cell_at(sheet, 2, column)
                .and_then(python_text)
                .map(|value| value.trim().to_owned())
                .filter(|value| !value.is_empty())
                .unwrap_or_else(|| format!("列{}", column_letters(column)));
            error_details.push(format!(
                "{header}: {} ≠ {} (差{})",
                fixed_2(total_value),
                fixed_2(sub_sum),
                fixed_2(difference)
            ));
        }
    }

    let missing = SUB_REGIONS
        .iter()
        .filter(|name| !rows.contains_key(**name))
        .copied()
        .collect::<Vec<_>>();
    if !missing.is_empty() {
        error_details.push(format!("未找到: {}", missing.join(", ")));
    }
    SheetCheck {
        error_count: numeric_error_count,
        detail: error_details.join("; "),
    }
}

fn cell_at(sheet: &SheetValues, row: usize, column: usize) -> Option<&CellValue> {
    sheet.values.get(&cell_address(row, column))
}

fn cell_address(row: usize, column: usize) -> String {
    format!("{}{}", column_letters(column), row)
}

fn column_letters(mut column: usize) -> String {
    let mut letters = String::new();
    while column > 0 {
        column -= 1;
        letters.insert(0, char::from(b'A' + (column % 26) as u8));
        column /= 26;
    }
    letters
}

fn python_text(value: &CellValue) -> Option<String> {
    match value {
        CellValue::Empty => None,
        CellValue::Text(value)
        | CellValue::DateTimeIso(value)
        | CellValue::DurationIso(value)
        | CellValue::Error(value) => Some(value.clone()),
        CellValue::Integer(value) => Some(value.to_string()),
        CellValue::Number(value) => {
            let mut text = value.to_string();
            if value.is_finite()
                && value.fract() == 0.0
                && !text.contains('e')
                && !text.contains('E')
            {
                text.push_str(".0");
            }
            Some(text)
        }
        CellValue::Boolean(value) => Some(if *value { "True" } else { "False" }.to_owned()),
        CellValue::DateTimeSerial(value) => Some(value.to_string()),
    }
}

fn number_or_none(value: Option<&CellValue>) -> Option<f64> {
    match value {
        None | Some(CellValue::Empty) => Some(0.0),
        Some(CellValue::Integer(value)) => Some(*value as f64),
        Some(CellValue::Number(value)) => Some(*value),
        Some(CellValue::Boolean(value)) => Some(if *value { 1.0 } else { 0.0 }),
        Some(CellValue::Text(value)) => parse_python_float(value),
        Some(
            CellValue::DateTimeSerial(_)
            | CellValue::DateTimeIso(_)
            | CellValue::DurationIso(_)
            | CellValue::Error(_),
        ) => None,
    }
}

fn parse_python_float(value: &str) -> Option<f64> {
    let value = value.trim();
    if value.is_empty() || value == "-" {
        return Some(0.0);
    }

    // The Python 3.12 baseline accepts Unicode 15.0 decimal digits and
    // underscores only between decimal digits. Normalize those digits and
    // separators before delegating the numeric conversion to Rust.
    let characters = value.chars().collect::<Vec<_>>();
    let mut normalized = String::with_capacity(value.len());
    for (index, character) in characters.iter().copied().enumerate() {
        if character == '_' {
            if index == 0
                || index + 1 == characters.len()
                || decimal_digit_value(characters[index - 1]).is_none()
                || decimal_digit_value(characters[index + 1]).is_none()
            {
                return None;
            }
            continue;
        }
        if let Some(digit) = decimal_digit_value(character) {
            normalized.push(char::from(b'0' + digit));
        } else {
            normalized.push(character);
        }
    }
    normalized.parse::<f64>().ok()
}

const PYTHON_UNICODE_15_DECIMAL_ZEROES: [u32; 68] = [
    0x0030, 0x0660, 0x06F0, 0x07C0, 0x0966, 0x09E6, 0x0A66, 0x0AE6, 0x0B66, 0x0BE6, 0x0C66, 0x0CE6,
    0x0D66, 0x0DE6, 0x0E50, 0x0ED0, 0x0F20, 0x1040, 0x1090, 0x17E0, 0x1810, 0x1946, 0x19D0, 0x1A80,
    0x1A90, 0x1B50, 0x1BB0, 0x1C40, 0x1C50, 0xA620, 0xA8D0, 0xA900, 0xA9D0, 0xA9F0, 0xAA50, 0xABF0,
    0xFF10, 0x104A0, 0x10D30, 0x11066, 0x110F0, 0x11136, 0x111D0, 0x112F0, 0x11450, 0x114D0,
    0x11650, 0x116C0, 0x11730, 0x118E0, 0x11950, 0x11C50, 0x11D50, 0x11DA0, 0x11F50, 0x16A60,
    0x16AC0, 0x16B50, 0x1D7CE, 0x1D7D8, 0x1D7E2, 0x1D7EC, 0x1D7F6, 0x1E140, 0x1E2F0, 0x1E4F0,
    0x1E950, 0x1FBF0,
];

fn decimal_digit_value(character: char) -> Option<u8> {
    let codepoint = character as u32;
    let index = PYTHON_UNICODE_15_DECIMAL_ZEROES.partition_point(|zero| *zero <= codepoint);
    let zero = *PYTHON_UNICODE_15_DECIMAL_ZEROES.get(index.checked_sub(1)?)?;
    let value = codepoint.checked_sub(zero)?;
    (value < 10).then_some(value as u8)
}

fn fixed_2(value: f64) -> String {
    if value.is_nan() {
        "nan".to_owned()
    } else if value == f64::INFINITY {
        "inf".to_owned()
    } else if value == f64::NEG_INFINITY {
        "-inf".to_owned()
    } else {
        format!("{value:.2}")
    }
}

fn save_report(
    output_dir: &Path,
    output_path: PathBuf,
    rows: &[ResultRow],
    files: usize,
    sheets: usize,
    errors: usize,
) -> Result<PathBuf> {
    let mut report_rows = Vec::<Vec<String>>::new();
    report_rows.push(
        RESULT_HEADERS
            .iter()
            .map(|value| (*value).to_owned())
            .collect(),
    );
    for row in rows {
        report_rows.push(vec![
            row.sequence.to_string(),
            row.file_name.clone(),
            row.sheet_name.clone(),
            row.status.clone(),
            row.error_count.to_string(),
            row.detail.clone(),
        ]);
    }
    if !rows.is_empty() {
        report_rows.push(vec![
            "=== 统计汇总 ===".to_owned(),
            format!("已检查文件数: {files}"),
            format!("已检查工作表数: {sheets}"),
            format!("发现错误总数: {errors}"),
            String::new(),
            String::new(),
        ]);
    }

    let header_format = Format::new()
        .set_bold()
        .set_align(FormatAlign::Center)
        .set_background_color(Color::RGB(0xC8DCF0));
    let error_format = Format::new()
        .set_font_color(Color::RGB(0xC00000))
        .set_background_color(Color::RGB(0xFFC8C8));
    let ok_format = Format::new()
        .set_font_color(Color::RGB(0x006400))
        .set_background_color(Color::RGB(0xC8FFC8));

    let mut workbook = Workbook::new();
    {
        let worksheet = workbook.add_worksheet();
        worksheet.set_name(RESULT_SHEET)?;
        // The Python baseline stores this marker as a formula with no cached
        // value, so preserve that cache state instead of rust_xlsxwriter's 0.
        worksheet.set_formula_result_default("");
        for (column, header) in RESULT_HEADERS.iter().enumerate() {
            worksheet.write_string_with_format(0, column as u16, *header, &header_format)?;
        }
        for (index, row) in rows.iter().enumerate() {
            let excel_row = index as u32 + 1;
            let format = if row.error_count > 0 {
                &error_format
            } else {
                &ok_format
            };
            worksheet.write_number_with_format(excel_row, 0, row.sequence as f64, format)?;
            worksheet.write_string_with_format(excel_row, 1, &row.file_name, format)?;
            worksheet.write_string_with_format(excel_row, 2, &row.sheet_name, format)?;
            worksheet.write_string_with_format(excel_row, 3, &row.status, format)?;
            worksheet.write_number_with_format(excel_row, 4, row.error_count as f64, format)?;
            worksheet.write_string_with_format(excel_row, 5, &row.detail, format)?;
        }
        if !rows.is_empty() {
            let row = rows.len() as u32 + 1;
            worksheet.write_formula(row, 0, "=== 统计汇总 ===")?;
            worksheet.write_string(row, 1, format!("已检查文件数: {files}"))?;
            worksheet.write_string(row, 2, format!("已检查工作表数: {sheets}"))?;
            worksheet.write_string(row, 3, format!("发现错误总数: {errors}"))?;
        }
        for column in 0..RESULT_HEADERS.len() {
            let max_length = report_rows
                .iter()
                .filter_map(|row| row.get(column))
                .map(|value| display_width(value))
                .max()
                .unwrap_or_default();
            let width_units = (max_length + 2).clamp(8, 80) as u32;
            // The source writes openpyxl's raw width value. rust_xlsxwriter's
            // character-unit API adds 5 px of padding, so use its pixel API to
            // reproduce the same OOXML width attribute for the default 7 px digit.
            worksheet.set_column_width_pixels(column as u16, width_units * 7)?;
        }
    }

    let temporary = create_temporary_path(output_dir)?;
    let save_result = workbook
        .save(&temporary)
        .with_context(|| format!("write temporary report {}", temporary.display()));
    if let Err(error) = save_result {
        let _ = fs::remove_file(&temporary);
        return Err(error);
    }

    let committed_source = if rows.is_empty() {
        temporary.clone()
    } else {
        let patched = match create_temporary_path(output_dir) {
            Ok(path) => path,
            Err(error) => {
                let _ = fs::remove_file(&temporary);
                return Err(error);
            }
        };
        let formula_marker_cell = format!("A{}", rows.len() + 2);
        if let Err(error) = ooxml_patch::clear_empty_formula_cache(
            &temporary,
            &patched,
            RESULT_SHEET,
            &formula_marker_cell,
        ) {
            let _ = fs::remove_file(&temporary);
            let _ = fs::remove_file(&patched);
            return Err(error).context("preserve empty formula cache in result marker");
        }
        fs::remove_file(&temporary).with_context(|| {
            format!("remove unpatched temporary report {}", temporary.display())
        })?;
        patched
    };

    let committed_path = commit_without_overwrite(&committed_source, output_path)?;
    Ok(committed_path)
}

fn display_width(value: &str) -> usize {
    value
        .chars()
        .map(|character| if character as u32 > 127 { 2 } else { 1 })
        .sum()
}

fn create_temporary_path(output_dir: &Path) -> Result<PathBuf> {
    for attempt in 0u32..10_000 {
        let path = output_dir.join(format!(
            ".region_sum_{}_{}.tmp.xlsx",
            std::process::id(),
            attempt
        ));
        match OpenOptions::new().write(true).create_new(true).open(&path) {
            Ok(file) => {
                drop(file);
                return Ok(path);
            }
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => return Err(error).with_context(|| format!("reserve {}", path.display())),
        }
    }
    bail!("could not allocate temporary report path")
}

fn commit_without_overwrite(temporary: &Path, desired: PathBuf) -> Result<PathBuf> {
    for suffix in 0u32..10_000 {
        let candidate = if suffix == 0 {
            desired.clone()
        } else {
            let stem = desired
                .file_stem()
                .and_then(|name| name.to_str())
                .unwrap_or("report");
            let extension = desired
                .extension()
                .and_then(|name| name.to_str())
                .unwrap_or("xlsx");
            desired.with_file_name(format!("{stem}_{suffix}.{extension}"))
        };
        match fs::hard_link(temporary, &candidate) {
            Ok(()) => {
                let _ = fs::remove_file(temporary);
                return Ok(candidate);
            }
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => {
                let _ = fs::remove_file(temporary);
                return Err(error).with_context(|| {
                    format!("commit report without overwriting {}", candidate.display())
                });
            }
        }
    }
    let _ = fs::remove_file(temporary);
    bail!("could not allocate a unique report output path")
}

#[cfg(test)]
mod tests {
    use super::{check_one_sheet, number_or_none, parse_python_float};
    use excel_reader::{CellValue, SheetValues};
    use std::collections::BTreeMap;

    fn sheet(values: &[(&str, CellValue)], max_row: usize, max_column: usize) -> SheetValues {
        SheetValues {
            name: "分地区测试".to_owned(),
            values: values
                .iter()
                .map(|(address, value)| ((*address).to_owned(), value.clone()))
                .collect::<BTreeMap<_, _>>(),
            max_row,
            max_column,
        }
    }

    #[test]
    fn numeric_conversion_preserves_source_blank_boolean_and_date_rules() {
        assert_eq!(number_or_none(None), Some(0.0));
        assert_eq!(number_or_none(Some(&CellValue::Boolean(true))), Some(1.0));
        assert_eq!(
            number_or_none(Some(&CellValue::Text(" - ".to_owned()))),
            Some(0.0)
        );
        assert_eq!(
            number_or_none(Some(&CellValue::Text("2.5".to_owned()))),
            Some(2.5)
        );
        assert_eq!(
            number_or_none(Some(&CellValue::Text("bad".to_owned()))),
            None
        );
        assert_eq!(
            number_or_none(Some(&CellValue::DateTimeSerial(45_000.0))),
            None
        );
        assert_eq!(parse_python_float(" Infinity "), Some(f64::INFINITY));
    }

    #[test]
    fn numeric_text_underscores_match_python_float_grammar() {
        assert_eq!(parse_python_float("1_000"), Some(1_000.0));
        assert_eq!(parse_python_float("1_000.25"), Some(1_000.25));
        assert_eq!(parse_python_float("1.25e1_0"), Some(12_500_000_000.0));
        assert_eq!(parse_python_float("١٢"), Some(12.0));
        assert_eq!(parse_python_float("١_٢"), Some(12.0));
        assert_eq!(parse_python_float("１２３"), Some(123.0));
        assert_eq!(parse_python_float("１_２"), Some(12.0));
        assert_eq!(parse_python_float("1__000"), None);
        assert_eq!(parse_python_float("1_.0"), None);
        assert_eq!(parse_python_float("1._0"), None);
        assert_eq!(parse_python_float("1e_10"), None);
        assert_eq!(parse_python_float("١__٢"), None);

        // Unicode 15 includes five mathematical decimal digit blocks. Keep
        // all blocks in sync with Python 3.12's float(string) behavior.
        for zero in [0x1D7CE_u32, 0x1D7D8, 0x1D7E2, 0x1D7EC, 0x1D7F6] {
            let digit = |offset: u32| char::from_u32(zero + offset).unwrap();
            let value = format!(
                "{}_{}{}{}.{}{}",
                digit(1),
                digit(0),
                digit(0),
                digit(0),
                digit(2),
                digit(5)
            );
            assert_eq!(parse_python_float(&value), Some(1_000.25), "{value:?}");
        }
    }

    #[test]
    fn sheet_check_parses_valid_numeric_separators_and_ignores_malformed_values() {
        let input = sheet(
            &[
                ("B2", CellValue::Text("地区".to_owned())),
                ("C2", CellValue::Text("金额".to_owned())),
                ("D2", CellValue::Text("无效格式".to_owned())),
                ("B3", CellValue::Text("惠州市".to_owned())),
                ("C3", CellValue::Text("1_000.25".to_owned())),
                ("D3", CellValue::Text("1__000".to_owned())),
                ("B4", CellValue::Text("惠阳区".to_owned())),
                ("C4", CellValue::Integer(500)),
                ("D4", CellValue::Integer(1)),
                ("B5", CellValue::Text("博罗县".to_owned())),
                ("C5", CellValue::Integer(499)),
                ("D5", CellValue::Integer(2)),
                ("B6", CellValue::Text("惠东县".to_owned())),
                ("C6", CellValue::Integer(0)),
                ("D6", CellValue::Integer(0)),
                ("B7", CellValue::Text("龙门县".to_owned())),
                ("C7", CellValue::Integer(0)),
                ("D7", CellValue::Integer(0)),
                ("B8", CellValue::Text("惠州市本部".to_owned())),
                ("C8", CellValue::Integer(0)),
                ("D8", CellValue::Integer(0)),
                ("B9", CellValue::Text("大亚湾".to_owned())),
                ("C9", CellValue::Integer(0)),
                ("D9", CellValue::Integer(0)),
            ],
            9,
            4,
        );

        let result = check_one_sheet(&input);
        assert_eq!(result.error_count, 1);
        assert_eq!(result.detail, "金额: 1000.25 ≠ 999.00 (差1.25)");
    }

    #[test]
    fn sheet_check_counts_numeric_mismatches_but_preserves_missing_region_text() {
        let input = sheet(
            &[
                ("B2", CellValue::Text("地区".to_owned())),
                ("C2", CellValue::Text("金额".to_owned())),
                ("B3", CellValue::Text("惠州市".to_owned())),
                ("C3", CellValue::Integer(10)),
                ("B4", CellValue::Text("惠阳区".to_owned())),
                ("C4", CellValue::Integer(3)),
            ],
            4,
            3,
        );
        let result = check_one_sheet(&input);
        assert_eq!(result.error_count, 1);
        assert!(result.detail.starts_with("金额: 10.00 ≠ 3.00 (差7.00)"));
        assert!(result
            .detail
            .contains("未找到: 博罗县, 惠东县, 龙门县, 惠州市本部, 大亚湾"));
    }
}
