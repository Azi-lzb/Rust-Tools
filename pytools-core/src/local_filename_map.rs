//! Ordered substring replacement for filenames, migrated from pytools operation 4-8.

use anyhow::{anyhow, Result};
use excel_reader::{read_workbook, SheetDump};
use serde::Serialize;
use std::path::{Path, PathBuf};
use std::time::Instant;

const CONFIG_SHEET: &str = "重命名配置";
const SOURCE_HEADER: &str = "原文件名片段";
const TARGET_HEADER: &str = "新文件名片段";

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct LocalFilenameMapRow {
    pub source_path: String,
    pub original_name: String,
    pub new_name: String,
    pub status: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct LocalFilenameMapRunReport {
    pub ok: usize,
    pub skip: usize,
    pub rules: usize,
    pub rows: Vec<LocalFilenameMapRow>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct LocalFilenameMapProgressEvent {
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

/// Execute operation 4-8 using ordered M:N filename fragment mappings.
///
/// The configuration is read through the shared Calamine reader. The function
/// accepts regular files of any extension, does not open or rewrite them, and
/// reserves rename targets without replacing existing files.
pub fn run_local_filename_map_candidate(
    config_path: impl AsRef<Path>,
    files: &[PathBuf],
) -> Result<LocalFilenameMapRunReport> {
    run_local_filename_map_with_progress(config_path, files, |_| {})
}

pub fn run_local_filename_map_with_progress<F>(
    config_path: impl AsRef<Path>,
    files: &[PathBuf],
    mut on_event: F,
) -> Result<LocalFilenameMapRunReport>
where
    F: FnMut(&LocalFilenameMapProgressEvent),
{
    let started = Instant::now();
    let mut sequence = 0;
    emit_progress(
        &mut on_event,
        &mut sequence,
        &started,
        "run",
        "started",
        0,
        files.len(),
        &["Rust"],
        None,
        None,
        format!("开始按配置映射文件名，共 {} 个文件。", files.len()),
    );
    emit_progress(
        &mut on_event,
        &mut sequence,
        &started,
        "configuration_read",
        "started",
        0,
        files.len(),
        &["Rust", "excel-reader", "Calamine"],
        Some("Calamine"),
        None,
        "正在读取局部文件名映射配置。",
    );
    let mappings = match load_mappings(config_path.as_ref()) {
        Ok(mappings) if !mappings.is_empty() => mappings,
        Ok(_) => {
            emit_progress(
                &mut on_event,
                &mut sequence,
                &started,
                "configuration_read",
                "failed",
                0,
                files.len(),
                &["Rust", "excel-reader", "Calamine"],
                Some("Calamine"),
                Some("mapping_rules_missing"),
                "配置中没有可用的局部文件名映射。",
            );
            return Err(anyhow!("重命名配置 缺少 M/N 列局部文件名映射。"));
        }
        Err(error) => {
            emit_progress(
                &mut on_event,
                &mut sequence,
                &started,
                "configuration_read",
                "failed",
                0,
                files.len(),
                &["Rust", "excel-reader", "Calamine"],
                Some("Calamine"),
                Some("configuration_read_failed"),
                "局部文件名映射配置读取失败。",
            );
            return Err(error);
        }
    };
    emit_progress(
        &mut on_event,
        &mut sequence,
        &started,
        "configuration_read",
        "completed",
        0,
        files.len(),
        &["Rust", "excel-reader", "Calamine"],
        Some("Calamine"),
        None,
        format!("已读取 {} 条局部文件名映射规则。", mappings.len()),
    );

    let mut ok = 0;
    let mut skip = 0;
    let mut rows = Vec::with_capacity(files.len());
    emit_progress(
        &mut on_event,
        &mut sequence,
        &started,
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
        emit_progress(
            &mut on_event,
            &mut sequence,
            &started,
            "file_processing",
            "started",
            index,
            files.len(),
            &["Rust"],
            None,
            None,
            format!("正在处理第 {input_index}/{} 个文件。", files.len()),
        );
        let row = map_local_filename(source, &mappings);
        if row.new_name.is_empty() {
            skip += 1;
        } else {
            ok += 1;
        }
        let error_type = match row.status.as_str() {
            "跳过：非文件或文件不存在" => Some("input_not_file"),
            "跳过：文件名不是 UTF-8" => Some("filename_not_utf8"),
            "跳过：改名失败" => Some("rename_failed"),
            _ => None,
        };
        emit_progress(
            &mut on_event,
            &mut sequence,
            &started,
            "file_processing",
            "completed",
            input_index,
            files.len(),
            &["Rust", "RustFilesystem"],
            Some("RustFilesystem"),
            error_type,
            format!(
                "已处理 {input_index}/{} 个文件；改名成功 {ok}，跳过 {skip}。",
                files.len()
            ),
        );
        rows.push(row);
    }

    let report = LocalFilenameMapRunReport {
        ok,
        skip,
        rules: mappings.len(),
        rows,
    };
    let (components, actual_engine) = if files.is_empty() {
        (&["Rust", "excel-reader", "Calamine"][..], "Calamine")
    } else {
        (
            &["Rust", "excel-reader", "Calamine", "RustFilesystem"][..],
            "Calamine + RustFilesystem",
        )
    };
    emit_progress(
        &mut on_event,
        &mut sequence,
        &started,
        "task",
        "completed",
        files.len(),
        files.len(),
        components,
        Some(actual_engine),
        None,
        format!(
            "局部文件名映射完成：规则 {} 条，改名 {} 个，跳过 {} 个。",
            report.rules, report.ok, report.skip
        ),
    );
    Ok(report)
}

fn map_local_filename(source: &Path, mappings: &[(String, String)]) -> LocalFilenameMapRow {
    let mut row = LocalFilenameMapRow {
        source_path: source.to_string_lossy().into_owned(),
        original_name: source
            .file_name()
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_default(),
        new_name: String::new(),
        status: String::new(),
    };
    if !source.exists() || !source.is_file() {
        row.status = "跳过：非文件或文件不存在".to_owned();
        return row;
    }

    let Some(file_name) = source.file_name().and_then(|name| name.to_str()) else {
        row.status = "跳过：文件名不是 UTF-8".to_owned();
        return row;
    };
    let new_name = apply_mapping(file_name, mappings);
    if new_name == row.original_name {
        row.status = "跳过：没有匹配片段".to_owned();
        return row;
    }
    if !is_single_filename(&new_name) {
        row.status = "跳过：改名失败".to_owned();
        return row;
    }

    match crate::file_sort_prefix::move_to_available_path(source, &source.with_file_name(&new_name))
    {
        Ok(target) => {
            row.new_name = target
                .file_name()
                .map(|name| name.to_string_lossy().into_owned())
                .unwrap_or_default();
            row.status = "已按局部映射改名".to_owned();
        }
        Err(_) => row.status = "跳过：改名失败".to_owned(),
    }
    row
}

fn emit_progress<F>(
    on_event: &mut F,
    sequence: &mut usize,
    started: &Instant,
    phase: &str,
    status: &str,
    completed_files: usize,
    total_files: usize,
    components: &[&str],
    actual_engine: Option<&str>,
    error_type: Option<&str>,
    user_message: impl Into<String>,
) where
    F: FnMut(&LocalFilenameMapProgressEvent),
{
    *sequence += 1;
    on_event(&LocalFilenameMapProgressEvent {
        sequence: *sequence,
        feature_id: "PT-4-8".to_owned(),
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

fn load_mappings(path: &Path) -> Result<Vec<(String, String)>> {
    let workbook = read_workbook(path)?;
    let sheet = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == CONFIG_SHEET)
        .ok_or_else(|| anyhow!("重命名配置 缺少 M/N 列局部文件名映射。"))?;

    let source_column = header_column(sheet, SOURCE_HEADER);
    let target_column = header_column(sheet, TARGET_HEADER);
    let (Some(source_column), Some(target_column)) = (source_column, target_column) else {
        return Err(anyhow!("重命名配置 缺少 M/N 列局部文件名映射。"));
    };

    let max_row = sheet
        .values
        .keys()
        .filter_map(|address| parse_cell_address(address).map(|(row, _)| row))
        .max()
        .unwrap_or(1);
    let mappings = (2..=max_row)
        .filter_map(|row| {
            let source = pandas_object_cell_text(sheet, row, source_column);
            if source.is_empty() {
                return None;
            }
            Some((source, pandas_object_cell_text(sheet, row, target_column)))
        })
        .collect();
    Ok(mappings)
}

fn header_column(sheet: &SheetDump, expected: &str) -> Option<usize> {
    sheet
        .values
        .iter()
        .filter_map(|(address, value)| {
            let (row, column) = parse_cell_address(address)?;
            (row == 1 && value_text(value) == expected).then_some(column)
        })
        .min()
}

/// Match pandas read_excel(dtype=object) for blank cells in the PyCode config.
/// Missing cells become NaN, then _norm stringifies that value as "nan".
fn pandas_object_cell_text(sheet: &SheetDump, row: usize, column: usize) -> String {
    let Some(value) = sheet.values.get(&cell_address(row, column)) else {
        return "nan".to_owned();
    };

    let value_type = value["type"].as_str().unwrap_or_default();
    if value_type == "empty" || (value_type == "string" && value["value"].as_str() == Some("")) {
        return "nan".to_owned();
    }

    value_text(value)
}

fn value_text(value: &serde_json::Value) -> String {
    let text = match value["type"].as_str().unwrap_or_default() {
        "string" | "datetime_iso" | "duration_iso" | "error" => {
            value["value"].as_str().unwrap_or_default().to_owned()
        }
        "integer" | "number" | "datetime_serial" => value["value"].to_string(),
        "boolean" => match value["value"].as_bool() {
            Some(true) => "True".to_owned(),
            Some(false) => "False".to_owned(),
            None => String::new(),
        },
        _ => String::new(),
    };
    text.trim().to_owned()
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

fn apply_mapping(name: &str, mappings: &[(String, String)]) -> String {
    let (stem, suffix) = stem_and_suffix(name);
    let mut new_stem = stem.to_owned();
    for (source, target) in mappings {
        new_stem = new_stem.replace(source, target);
    }
    format!("{new_stem}{suffix}")
}

fn stem_and_suffix(name: &str) -> (&str, &str) {
    match name.rfind('.') {
        Some(index) if index > 0 && index + 1 < name.len() => (&name[..index], &name[index..]),
        _ => (name, ""),
    }
}

fn is_single_filename(name: &str) -> bool {
    !name.is_empty() && name != "." && name != ".." && !name.contains('/') && !name.contains('\\')
}

#[cfg(test)]
mod tests {
    use super::apply_mapping;

    #[test]
    fn mappings_apply_in_order_to_stem_and_preserve_suffix() {
        let mappings = vec![
            ("aa".to_owned(), "bb".to_owned()),
            ("bb".to_owned(), "cc".to_owned()),
        ];
        assert_eq!(
            apply_mapping("file_aa_aa.docx", &mappings),
            "file_cc_cc.docx"
        );
        assert_eq!(apply_mapping("unmatched.pdf", &mappings), "unmatched.pdf");
    }

    #[test]
    fn filename_suffix_matches_python_pathlib_for_trailing_dot_and_dotfile() {
        let mappings = vec![("aa".to_owned(), "bb".to_owned())];
        assert_eq!(apply_mapping("file.aa.", &mappings), "file.bb.");
        assert_eq!(apply_mapping(".hidden_aa", &mappings), ".hidden_bb");
    }
}

/// Compatibility alias for the original Python 4-8 operation identifier.
#[deprecated(note = "use `run_local_filename_map_candidate`")]
pub fn run_local_filename_map_48_candidate(
    config_path: impl AsRef<Path>,
    files: &[PathBuf],
) -> Result<LocalFilenameMapRunReport> {
    run_local_filename_map_candidate(config_path, files)
}
