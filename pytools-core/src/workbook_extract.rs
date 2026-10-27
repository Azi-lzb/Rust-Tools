//! OOXML/BIFF workbook input adapter for the timeline extraction rules.

use crate::config::{PathMapRule, TimelineRule};
use crate::rule_engine::{extract_cells_from_matrix, ExtractedCell, MergeRange, SheetValue};
use anyhow::{anyhow, Context, Result};
use excel_reader::{read_workbook, SheetDump};
use serde_json::Value;
use std::path::Path;

#[derive(Debug, Clone, PartialEq)]
pub enum WorkbookCell {
    Text(String),
    Integer(i64),
    Number(f64),
    Boolean(bool),
    DateTimeSerial(f64),
    DurationSerial(f64),
    DateTimeIso(String),
    DurationIso(String),
    Error(String),
}

impl WorkbookCell {
    pub(crate) fn from_json(value: &Value) -> Option<Self> {
        let kind = value.get("type")?.as_str()?;
        let data = value.get("value")?;
        Some(match kind {
            "string" => Self::Text(data.as_str()?.to_owned()),
            "integer" => Self::Integer(data.as_i64()?),
            "number" => {
                let number = data.as_f64()?;
                if number.is_finite()
                    && number.fract() == 0.0
                    && number >= i64::MIN as f64
                    && number <= i64::MAX as f64
                {
                    Self::Integer(number as i64)
                } else {
                    Self::Number(number)
                }
            }
            "boolean" => Self::Boolean(data.as_bool()?),
            "datetime_serial" => Self::DateTimeSerial(data.as_f64()?),
            "duration_serial" => Self::DurationSerial(data.as_f64()?),
            "datetime_iso" => Self::DateTimeIso(data.as_str()?.to_owned()),
            "duration_iso" => Self::DurationIso(data.as_str()?.to_owned()),
            "error" => Self::Error(data.as_str()?.to_owned()),
            _ => return None,
        })
    }

    pub fn to_json_value(&self) -> Value {
        match self {
            Self::Text(value)
            | Self::DateTimeIso(value)
            | Self::DurationIso(value)
            | Self::Error(value) => Value::String(value.clone()),
            Self::Integer(value) => Value::from(*value),
            Self::Number(value) | Self::DateTimeSerial(value) | Self::DurationSerial(value) => {
                serde_json::Number::from_f64(*value)
                    .map(Value::Number)
                    .unwrap_or(Value::Null)
            }
            Self::Boolean(value) => Value::Bool(*value),
        }
    }
}

impl SheetValue for WorkbookCell {
    fn to_sheet_text(&self) -> String {
        match self {
            Self::Text(value)
            | Self::DateTimeIso(value)
            | Self::DurationIso(value)
            | Self::Error(value) => value.clone(),
            Self::Integer(value) => value.to_string(),
            Self::Number(value) | Self::DateTimeSerial(value) | Self::DurationSerial(value) => {
                format!("{value:?}")
            }
            Self::Boolean(value) => (if *value { "True" } else { "False" }).to_owned(),
        }
    }
}

/// Read one worksheet with the shared Rust Excel reader and expand it using the timeline rules.
pub fn extract_cells_from_workbook(
    path: impl AsRef<Path>,
    sheet_name: &str,
    rule: &TimelineRule,
    path_maps: &[PathMapRule],
    row_suffix_enabled: bool,
) -> Result<Vec<ExtractedCell<WorkbookCell>>> {
    let path = path.as_ref();
    let workbook =
        read_workbook(path).with_context(|| format!("read source workbook {}", path.display()))?;
    let sheet = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == sheet_name)
        .ok_or_else(|| anyhow!("worksheet not found: {sheet_name}"))?;
    let (matrix, merges) = matrix_from_dump(sheet)?;
    let file_name = path
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or_else(|| anyhow!("source workbook has no valid file name"))?;
    extract_cells_from_matrix(
        &matrix,
        &merges,
        file_name,
        rule,
        sheet_name,
        path_maps,
        row_suffix_enabled,
        &Default::default(),
    )
    .map_err(|message| anyhow!(message))
}

pub(crate) fn matrix_from_dump(
    sheet: &SheetDump,
) -> Result<(Vec<Vec<Option<WorkbookCell>>>, Vec<MergeRange>)> {
    let merges = sheet
        .merges
        .iter()
        .map(|range| parse_merge_range(range))
        .collect::<Result<Vec<_>>>()?;
    let cells = sheet
        .values
        .iter()
        .filter_map(|(address, value)| parse_address(address).map(|coord| (coord, value)))
        .collect::<Vec<_>>();
    let rows = cells
        .iter()
        .map(|((row, _), _)| *row + 1)
        .chain(merges.iter().map(|range| range.max_row + 1))
        .max()
        .unwrap_or(0);
    let columns = cells
        .iter()
        .map(|((_, column), _)| *column + 1)
        .chain(merges.iter().map(|range| range.max_col + 1))
        .max()
        .unwrap_or(0);
    let mut matrix = vec![vec![None; columns]; rows];
    for ((row, column), value) in cells {
        if let Some(value) = WorkbookCell::from_json(value) {
            matrix[row][column] = Some(value);
        }
    }
    Ok((matrix, merges))
}

fn parse_merge_range(value: &str) -> Result<MergeRange> {
    let (start, end) = value.split_once(':').unwrap_or((value, value));
    let (min_row, min_col) =
        parse_address(start).ok_or_else(|| anyhow!("invalid merge range start: {value}"))?;
    let (max_row, max_col) =
        parse_address(end).ok_or_else(|| anyhow!("invalid merge range end: {value}"))?;
    Ok(MergeRange {
        min_row,
        max_row,
        min_col,
        max_col,
    })
}

fn parse_address(address: &str) -> Option<(usize, usize)> {
    let bytes = address.as_bytes();
    let letters = bytes
        .iter()
        .take_while(|byte| byte.is_ascii_alphabetic())
        .count();
    if letters == 0 || letters == bytes.len() {
        return None;
    }
    let mut column = 0usize;
    for letter in &bytes[..letters] {
        column = column
            .checked_mul(26)?
            .checked_add((letter.to_ascii_uppercase() - b'A' + 1) as usize)?;
    }
    let row = std::str::from_utf8(&bytes[letters..])
        .ok()?
        .parse::<usize>()
        .ok()?;
    if row == 0 {
        return None;
    }
    Some((row - 1, column - 1))
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::{json, Value};
    use std::path::PathBuf;

    struct TestDir(PathBuf);

    impl TestDir {
        fn new() -> Self {
            let nonce = std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos();
            let path = std::env::temp_dir().join(format!(
                "pytools-workbook-extract-{}-{nonce}",
                std::process::id()
            ));
            std::fs::create_dir(&path).unwrap();
            Self(path)
        }

        fn path(&self) -> &Path {
            &self.0
        }
    }

    impl Drop for TestDir {
        fn drop(&mut self) {
            let _ = std::fs::remove_dir_all(&self.0);
        }
    }

    fn string_vec(value: &Value) -> Vec<String> {
        value
            .as_array()
            .unwrap()
            .iter()
            .map(|item| item.as_str().unwrap().to_owned())
            .collect()
    }

    #[test]
    fn workbook_extraction_matches_original_python_helper() {
        let cases: Value = serde_json::from_str(include_str!(
            "../tests/fixtures/workbook_extract_cases.json"
        ))
        .unwrap();
        let test_dir = TestDir::new();
        let fixture = test_dir.path().join(cases["filename"].as_str().unwrap());
        let mut workbook = rust_xlsxwriter::Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet
            .set_name(cases["sheet_name"].as_str().unwrap())
            .unwrap();
        let rows = cases["sheet"].as_array().unwrap();
        for (row_index, row) in rows.iter().enumerate() {
            for (column_index, value) in row.as_array().unwrap().iter().enumerate() {
                // The fixture contains one merged B2:C2 header; write it
                // through merge_range below so the generated workbook has
                // the same merge metadata as the Python baseline input.
                if row_index == 1 && matches!(column_index, 1 | 2) {
                    continue;
                }
                match value {
                    Value::Null => {}
                    Value::String(value) => {
                        sheet
                            .write_string(row_index as u32, column_index as u16, value)
                            .unwrap();
                    }
                    Value::Number(value) => {
                        sheet
                            .write_number(
                                row_index as u32,
                                column_index as u16,
                                value.as_f64().unwrap(),
                            )
                            .unwrap();
                    }
                    Value::Bool(value) => {
                        sheet
                            .write_boolean(row_index as u32, column_index as u16, *value)
                            .unwrap();
                    }
                    _ => panic!("unsupported workbook extraction fixture cell: {value}"),
                }
            }
        }
        let merge_header = rows[1][1].as_str().unwrap();
        sheet
            .merge_range(1, 1, 1, 2, merge_header, &rust_xlsxwriter::Format::new())
            .unwrap();
        workbook.save(&fixture).unwrap();
        let rule_case = &cases["rule"];
        let rule = TimelineRule {
            enabled: true,
            name: rule_case["name"].as_str().unwrap().to_owned(),
            wb_keyword: rule_case["wb_keyword"].as_str().unwrap().to_owned(),
            sheet_keyword: rule_case["sheet_keyword"].as_str().unwrap().to_owned(),
            row_header_col: rule_case["row_header_col"].as_u64().unwrap() as usize,
            col_header_row: rule_case["col_header_rows"][0].as_u64().unwrap() as usize,
            row_header_cols: rule_case["row_header_cols"]
                .as_array()
                .unwrap()
                .iter()
                .map(|value| value.as_u64().unwrap() as usize)
                .collect(),
            row_header_col_specified: rule_case["row_header_col_specified"].as_bool().unwrap(),
            col_header_rows: rule_case["col_header_rows"]
                .as_array()
                .unwrap()
                .iter()
                .map(|value| value.as_u64().unwrap() as usize)
                .collect(),
            required_col_headers: string_vec(&rule_case["required_col_headers"]),
            required_row_headers: string_vec(&rule_case["required_row_headers"]),
            data_row_start: rule_case["data_row_start"].as_u64().unwrap() as usize,
            data_row_end: rule_case["data_row_end"]
                .as_u64()
                .map(|value| value as usize),
            data_col_start: rule_case["data_col_start"].as_u64().unwrap() as usize,
            data_col_end: rule_case["data_col_end"]
                .as_u64()
                .map(|value| value as usize),
            skip_keywords: string_vec(&rule_case["skip_keywords"]),
            target_wb_path: None,
            target_sheet: None,
            target_write_enabled: true,
            set_items: Vec::new(),
            set_parse_error: String::new(),
            target_dedup_cols: Vec::new(),
            output_row_no: false,
        };
        let path_maps = cases["path_maps"]
            .as_array()
            .unwrap()
            .iter()
            .map(|map| PathMapRule {
                enabled: true,
                name: map["name"].as_str().unwrap().to_owned(),
                applicable_rule: map["applicable_rule"].as_str().unwrap().to_owned(),
                wb_keyword: map["wb_keyword"].as_str().unwrap().to_owned(),
                sheet_keyword: map["sheet_keyword"].as_str().unwrap().to_owned(),
                target: map["target"].as_str().unwrap().to_owned(),
                match_mode: map["match_mode"].as_str().unwrap().to_owned(),
                original: map["original"].as_str().unwrap().to_owned(),
                standard: map["standard"].as_str().unwrap().to_owned(),
            })
            .collect::<Vec<_>>();

        let extracted = extract_cells_from_workbook(
            &fixture,
            cases["sheet_name"].as_str().unwrap(),
            &rule,
            &path_maps,
            true,
        )
        .unwrap();
        let actual = extracted
            .iter()
            .map(|cell| {
                json!({
                    "source_file": cell.source_file,
                    "sheet_name": cell.sheet_name,
                    "rule_name": cell.rule_name,
                    "data_date": cell.data_date,
                    "row_path": cell.row_path,
                    "col_path": cell.col_path,
                    "value": cell.value.to_json_value(),
                    "cell_addr": cell.cell_addr,
                })
            })
            .collect::<Vec<_>>();
        let baseline: Value = serde_json::from_str(include_str!(
            "../probe/reports/workbook_extract_baseline.json"
        ))
        .expect("Python baseline report must be valid JSON");

        assert_eq!(baseline["python_baseline"]["status"], "保留");
        assert_eq!(Value::Array(actual), baseline["result"]);
    }
}
