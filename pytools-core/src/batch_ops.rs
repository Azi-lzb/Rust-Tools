//! Excel edit plans for the remaining first/second batch operations.

use anyhow::{bail, Context, Result};
use excel_editor::{
    inspect_xlsx_profile, CellValue, EditOperation, ExcelEditPlan, PreservationRequirement,
    WorkbookFormat,
};
use excel_reader::{read_workbook, SheetDump};
use serde::Serialize;
use std::{
    collections::{BTreeSet, HashMap},
    path::{Path, PathBuf},
};

const BRANCH_KEY: &str = "分机构";
const VILLAGE_KEY: &str = "村镇银行";
const SPLIT_NAMES: [&str; 2] = ["博罗长江村镇银行", "惠东惠民村镇银行"];
const INST_KEYS: [&str; 2] = ["惠州市", "涉农贷款分机构"];
const AREA_KEYS: [&str; 2] = ["本外币", "涉农贷款分地区"];
const QUARTER_AREA_KEY: &str = "涉农贷款分地区统计";
const QUARTER_INST_KEY: &str = "涉农贷款分机构";
const HEADER_TEXT: &str = "涉农贷款比上月";
const TOLERANCE: f64 = 0.01;

#[derive(Clone, Debug, Default, Serialize)]
pub struct PlanBuildReport {
    pub input: PathBuf,
    pub output: PathBuf,
    pub matched_sheets: usize,
    pub changed_cells: usize,
    pub plan: Option<ExcelEditPlan>,
    pub warnings: Vec<String>,
}

/// Plan 7-1, using the prior and current workbooks as read-only data sources.
pub fn build_village_split_plan(
    current: &Path,
    previous: &Path,
    target: &Path,
    output: &Path,
) -> Result<PlanBuildReport> {
    let cur = read_workbook(current)
        .with_context(|| format!("read current workbook {}", current.display()))?;
    let prev = read_workbook(previous)
        .with_context(|| format!("read previous workbook {}", previous.display()))?;
    let target_dump = read_workbook(target)
        .with_context(|| format!("read target workbook {}", target.display()))?;
    let matrix = village_matrix(&cur, &prev);
    anyhow::ensure!(
        !(numeric_plan(&matrix[0][0]) == 0.0 && numeric_plan(&matrix[0][4]) == 0.0),
        "village-bank matrix is empty; no output plan created"
    );
    let mut operations = Vec::new();
    let mut matched_sheets = 0;
    let mut warnings = Vec::new();
    for sheet in target_dump
        .sheets
        .iter()
        .filter(|sheet| sheet.name.contains(BRANCH_KEY))
    {
        let kind = if sheet.name.contains("本外币") {
            0
        } else if sheet.name.contains("人民币") {
            1
        } else if sheet.name.contains("外汇") {
            2
        } else {
            warnings.push(format!("unrecognized branch sheet: {}", sheet.name));
            continue;
        };
        let row =
            (1..=used_rows(sheet)).find(|row| normalize(&text(sheet, *row, 1)) == VILLAGE_KEY);
        let Some(row) = row else {
            warnings.push(format!("village bank row not found: {}", sheet.name));
            continue;
        };
        operations.push(EditOperation::DeleteRows {
            sheet: sheet.name.clone(),
            first_row: row as u32,
            count: 1,
        });
        operations.push(EditOperation::InsertRows {
            sheet: sheet.name.clone(),
            first_row: row as u32,
            count: 2,
        });
        operations.push(EditOperation::SetCellValue {
            sheet: sheet.name.clone(),
            cell: format!("A{row}"),
            value: CellValue::Text(SPLIT_NAMES[0].into()),
        });
        operations.push(EditOperation::SetCellValue {
            sheet: sheet.name.clone(),
            cell: format!("A{}", row + 1),
            value: CellValue::Text(SPLIT_NAMES[1].into()),
        });
        for column in 0..8 {
            let target_column = column + 2;
            operations.push(EditOperation::SetCellValue {
                sheet: sheet.name.clone(),
                cell: cell_address(row, target_column),
                value: matrix[kind][column].clone(),
            });
            operations.push(EditOperation::SetCellValue {
                sheet: sheet.name.clone(),
                cell: cell_address(row + 1, target_column),
                value: matrix[kind + 3][column].clone(),
            });
        }
        matched_sheets += 1;
    }
    let changed_cells = operations.len();
    let plan = make_plan(target, output, operations, "split-village-bank")?;
    Ok(PlanBuildReport {
        input: target.into(),
        output: output.into(),
        matched_sheets,
        changed_cells,
        plan,
        warnings,
    })
}

/// Plan 7-3; output path follows the legacy `(金融局)` suffix convention.
pub fn build_remove_foreign_plan(
    input: &Path,
    output: &Path,
    foreign_names: &BTreeSet<String>,
) -> Result<PlanBuildReport> {
    anyhow::ensure!(
        !foreign_names.is_empty(),
        "foreign-institution mapping is empty"
    );
    let dump = read_workbook(input)?;
    let mut operations = Vec::new();
    let mut matched_sheets = 0;
    for sheet in dump
        .sheets
        .iter()
        .filter(|sheet| sheet.name.contains(BRANCH_KEY))
    {
        matched_sheets += 1;
        let mut rows = Vec::new();
        for row in 1..=used_rows(sheet) {
            if foreign_names.contains(&normalize(&text(sheet, row, 1))) {
                rows.push(row);
            }
        }
        for (start, end) in descending_runs(rows) {
            operations.push(EditOperation::DeleteRows {
                sheet: sheet.name.clone(),
                first_row: start as u32,
                count: (end - start + 1) as u32,
            });
        }
    }
    let changed_cells = operations.len();
    let plan = make_plan(input, output, operations, "remove-foreign-bank")?;
    Ok(PlanBuildReport {
        input: input.into(),
        output: output.into(),
        matched_sheets,
        changed_cells,
        plan,
        warnings: Vec::new(),
    })
}

/// Plan 7-7 for the city-level rural-loan check.
pub fn build_rural_loan_plan(
    current: &Path,
    previous: &Path,
    output: &Path,
) -> Result<(PlanBuildReport, RuralLoanSummary)> {
    let current_dump = read_workbook(current)?;
    let previous_dump = read_workbook(previous)?;
    let current_inst = unique_sheet(&current_dump, &INST_KEYS)
        .context("current workbook lacks a unique city institution sheet")?;
    let previous_inst = unique_sheet(&previous_dump, &INST_KEYS)
        .context("previous workbook lacks a unique city institution sheet")?;
    let current_area = current_dump
        .sheets
        .iter()
        .find(|sheet| AREA_KEYS.iter().all(|key| sheet.name.contains(key)))
        .context("current workbook lacks a city area statistics sheet")?;
    let (current_values, current_rows, current_duplicates) =
        institution_values(current_inst, false);
    let (previous_values, previous_rows, previous_duplicates) =
        institution_values(previous_inst, true);
    anyhow::ensure!(
        current_duplicates.is_empty(),
        "duplicate institution names in current workbook: {}",
        current_duplicates.join("、")
    );
    anyhow::ensure!(
        previous_duplicates.is_empty(),
        "duplicate institution names in previous workbook: {}",
        previous_duplicates.join("、")
    );
    let mut matched = 0;
    let mut fill_total = 0.0;
    let mut unmatched = Vec::new();
    let mut operations = Vec::new();
    for (name, row) in institution_rows_in_order(&current_rows) {
        if let Some(previous_value) = previous_values.get(&name) {
            let difference = current_values[&name] - previous_value;
            fill_total += difference;
            matched += 1;
            operations.push(EditOperation::SetCellValue {
                sheet: current_inst.name.clone(),
                cell: format!("D{row}"),
                value: CellValue::Decimal(exact_decimal_text(difference)),
            });
        } else {
            unmatched.push(name);
            // Preserve the source implementation's openpyxl behavior:
            // assigning cell(..., value=None) does not clear an existing value.
        }
    }
    let header_col = find_header_column(current_area, HEADER_TEXT)
        .context("area sheet lacks the 涉农贷款比上月 header")?;
    let city_row = find_exact_row(current_area, &[1, 2, 3], "惠州市")
        .context("area sheet lacks 惠州市 row")?;
    let area_total = numeric_with_formula_mode(current_area, city_row, header_col, false);
    let previous_only = institution_rows_in_order(&previous_rows)
        .into_iter()
        .filter_map(|(name, _)| (!current_values.contains_key(&name)).then_some(name))
        .collect();
    let summary = RuralLoanSummary {
        matched,
        fill_total,
        area_total,
        consistent: (fill_total - area_total).abs() <= TOLERANCE,
        unmatched_current: unmatched,
        previous_only,
        error: None,
    };
    let changed_cells = operations.len();
    let plan = make_plan(current, output, operations, "adjust-rural-loan")?;
    Ok((
        PlanBuildReport {
            input: current.into(),
            output: output.into(),
            matched_sheets: 1,
            changed_cells,
            plan,
            warnings: Vec::new(),
        },
        summary,
    ))
}

/// Plan 7-8 for each region, collecting all row updates into one workbook plan.
pub fn build_quarterly_rural_loan_plan(
    current: &Path,
    previous: &Path,
    output: &Path,
) -> Result<(PlanBuildReport, QuarterlyLoanSummary)> {
    let current_dump = read_workbook(current)?;
    let previous_dump = read_workbook(previous)?;
    let current_sheets = region_sheets(&current_dump)?;
    let previous_sheets = region_sheets(&previous_dump)?
        .into_iter()
        .collect::<HashMap<_, _>>();
    anyhow::ensure!(
        !current_sheets.is_empty(),
        "current workbook has no rural-loan institution sheets"
    );
    let area = current_dump
        .sheets
        .iter()
        .filter(|sheet| sheet.name.contains(QUARTER_AREA_KEY))
        .collect::<Vec<_>>();
    anyhow::ensure!(
        area.len() == 1,
        "expected one rural-loan regional statistics sheet; found {}",
        area.len()
    );
    let area = area[0];
    let header_col = find_header_column(area, HEADER_TEXT)
        .context("regional statistics sheet lacks the month-on-month header")?;
    let mut operations = Vec::new();
    let mut regions = Vec::new();
    let mut total_fill = 0.0;
    let mut total_area = 0.0;
    let mut total_matched = 0;
    let mut warnings = Vec::new();
    for (region, current_sheet) in current_sheets {
        let Some(previous_sheet) = previous_sheets.get(&region) else {
            let error = format!("上期文件缺少{region}分机构表");
            warnings.push(error.clone());
            regions.push(QuarterlyRegionSummary {
                region,
                sheet: current_sheet.name.clone(),
                matched: 0,
                fill_total: 0.0,
                area_total: None,
                consistent: false,
                unmatched_current: Vec::new(),
                previous_only: Vec::new(),
                error: Some(error),
            });
            continue;
        };
        let (cur_values, cur_rows, cur_dups) = institution_values(current_sheet, false);
        let (prev_values, prev_rows, prev_dups) = institution_values(previous_sheet, true);
        if !cur_dups.is_empty() || !prev_dups.is_empty() {
            let duplicates = cur_dups.into_iter().chain(prev_dups).collect::<Vec<_>>();
            let error = format!("分机构表存在重复机构名称：{}", duplicates.join("、"));
            warnings.push(format!("{region}: {error}"));
            regions.push(QuarterlyRegionSummary {
                region,
                sheet: current_sheet.name.clone(),
                matched: 0,
                fill_total: 0.0,
                area_total: None,
                consistent: false,
                unmatched_current: Vec::new(),
                previous_only: Vec::new(),
                error: Some(error),
            });
            continue;
        }
        let mut region_fill = 0.0;
        let mut matched = 0;
        let mut unmatched = Vec::new();
        for (name, row) in institution_rows_in_order(&cur_rows) {
            if let Some(previous_value) = prev_values.get(&name) {
                let delta = cur_values[&name] - previous_value;
                region_fill += delta;
                matched += 1;
                operations.push(EditOperation::SetCellValue {
                    sheet: current_sheet.name.clone(),
                    cell: format!("D{row}"),
                    value: CellValue::Decimal(exact_decimal_text(delta)),
                });
            } else {
                unmatched.push(name);
                // Match the old openpyxl call cell(..., value=None), which
                // leaves a previously stored D-column value untouched.
            }
        }
        let area_row = find_exact_row(area, &[1, 2, 3], &region);
        let area_total =
            area_row.map(|row| numeric_with_formula_mode(area, row, header_col, false));
        let consistent = area_total.is_some_and(|total| (region_fill - total).abs() <= TOLERANCE);
        let error = if let Some(total) = area_total {
            total_area += total;
            None
        } else {
            let message = format!("“{}”未找到地区行：{region}", area.name);
            warnings.push(message.clone());
            Some(message)
        };
        total_fill += region_fill;
        total_matched += matched;
        regions.push(QuarterlyRegionSummary {
            region,
            sheet: current_sheet.name.clone(),
            matched,
            fill_total: region_fill,
            area_total,
            consistent,
            unmatched_current: unmatched,
            previous_only: institution_rows_in_order(&prev_rows)
                .into_iter()
                .filter_map(|(name, _)| (!cur_values.contains_key(&name)).then_some(name))
                .collect(),
            error,
        });
    }
    let consistent = !regions.is_empty()
        && regions
            .iter()
            .all(|region| region.consistent && region.error.is_none());
    let summary = QuarterlyLoanSummary {
        regions,
        matched: total_matched,
        fill_total: total_fill,
        area_total: total_area,
        consistent,
        error: None,
    };
    let changed_cells = operations.len();
    let plan = make_plan(current, output, operations, "adjust-quarterly-rural-loan")?;
    Ok((
        PlanBuildReport {
            input: current.into(),
            output: output.into(),
            matched_sheets: summary.regions.len(),
            changed_cells,
            plan,
            warnings,
        },
        summary,
    ))
}

#[derive(Clone, Debug, Serialize)]
pub struct RuralLoanSummary {
    pub matched: usize,
    pub fill_total: f64,
    pub area_total: f64,
    pub consistent: bool,
    pub unmatched_current: Vec<String>,
    pub previous_only: Vec<String>,
    pub error: Option<String>,
}
#[derive(Clone, Debug, Serialize)]
pub struct QuarterlyRegionSummary {
    pub region: String,
    pub sheet: String,
    pub matched: usize,
    pub fill_total: f64,
    pub area_total: Option<f64>,
    pub consistent: bool,
    pub unmatched_current: Vec<String>,
    pub previous_only: Vec<String>,
    pub error: Option<String>,
}
#[derive(Clone, Debug, Serialize)]
pub struct QuarterlyLoanSummary {
    pub regions: Vec<QuarterlyRegionSummary>,
    pub matched: usize,
    pub fill_total: f64,
    pub area_total: f64,
    pub consistent: bool,
    pub error: Option<String>,
}

fn village_matrix(
    current: &excel_reader::WorkbookDump,
    previous: &excel_reader::WorkbookDump,
) -> Vec<Vec<CellValue>> {
    let indicators = ["村镇银行本外币", "村镇银行人民币", "村镇银行外汇"];
    let mut matrix = vec![vec![CellValue::Decimal("0".into()); 8]; 6];
    for (index, indicator) in indicators.iter().enumerate() {
        let Some(cur) = current
            .sheets
            .iter()
            .find(|sheet| sheet.name.contains(indicator))
        else {
            continue;
        };
        let Some(prev) = previous.sheets.iter().find(|sheet| sheet.name == cur.name) else {
            continue;
        };
        for (offset, row) in [5usize, 6].iter().enumerate() {
            let target = index + offset * 3;
            let curr_dep = village_numeric(&cell_value(cur, *row, 3));
            let prev_dep = village_numeric(&cell_value(prev, *row, 3));
            let curr_loan = village_numeric(&cell_value(cur, *row, 87));
            let prev_loan = village_numeric(&cell_value(prev, *row, 87));
            matrix[target][0] = CellValue::Decimal(decimal_text(curr_dep));
            matrix[target][1] = value_for_cell(&cell_value(cur, *row, 4));
            matrix[target][2] = value_for_cell(&cell_value(cur, *row, 5));
            matrix[target][3] = yoy(curr_dep, prev_dep);
            matrix[target][4] = CellValue::Decimal(decimal_text(curr_loan));
            matrix[target][5] = value_for_cell(&cell_value(cur, *row, 88));
            matrix[target][6] = value_for_cell(&cell_value(cur, *row, 89));
            matrix[target][7] = yoy(curr_loan, prev_loan);
        }
    }
    matrix
}

fn institution_values(
    sheet: &SheetDump,
    data_only: bool,
) -> (HashMap<String, f64>, HashMap<String, usize>, Vec<String>) {
    let mut values = HashMap::new();
    let mut rows = HashMap::new();
    let mut duplicates = Vec::new();
    for row in 3..=used_rows(sheet) {
        let name = normalize(&text(sheet, row, 2));
        if name.is_empty() {
            continue;
        }
        if values.contains_key(&name) {
            duplicates.push(name);
            continue;
        }
        values.insert(
            name.clone(),
            numeric_with_formula_mode(sheet, row, 3, data_only),
        );
        rows.insert(name, row);
    }
    (values, rows, duplicates)
}
fn institution_rows_in_order(rows: &HashMap<String, usize>) -> Vec<(String, usize)> {
    let mut entries = rows
        .iter()
        .map(|(name, row)| (name.clone(), *row))
        .collect::<Vec<_>>();
    entries.sort_by_key(|(_, row)| *row);
    entries
}
fn unique_sheet<'a>(book: &'a excel_reader::WorkbookDump, keys: &[&str]) -> Option<&'a SheetDump> {
    let hits = book
        .sheets
        .iter()
        .filter(|sheet| keys.iter().all(|key| sheet.name.contains(key)))
        .collect::<Vec<_>>();
    (hits.len() == 1).then_some(hits[0])
}
fn region_sheets<'a>(book: &'a excel_reader::WorkbookDump) -> Result<Vec<(String, &'a SheetDump)>> {
    let mut result = Vec::new();
    let mut seen = BTreeSet::new();
    for sheet in book
        .sheets
        .iter()
        .filter(|sheet| sheet.name.contains(QUARTER_INST_KEY))
    {
        let index = sheet
            .name
            .find(QUARTER_INST_KEY)
            .context("invalid region sheet name")?;
        let region = normalize(&sheet.name[..index])
            .trim_end_matches([' ', '-', '—', '_', '－'])
            .to_owned();
        anyhow::ensure!(
            !region.is_empty(),
            "cannot determine region from worksheet {:?}",
            sheet.name
        );
        anyhow::ensure!(
            seen.insert(region.clone()),
            "duplicate institution worksheet for region {region}"
        );
        result.push((region, sheet));
    }
    Ok(result)
}
fn find_header_column(sheet: &SheetDump, header: &str) -> Option<usize> {
    for row in 1..=used_rows(sheet).min(10) {
        for col in 1..=used_columns(sheet) {
            if normalize(&text(sheet, row, col)) == header {
                return Some(col);
            }
        }
    }
    None
}
fn find_exact_row(sheet: &SheetDump, columns: &[usize], value: &str) -> Option<usize> {
    (1..=used_rows(sheet)).find(|row| {
        columns
            .iter()
            .any(|column| normalize(&text(sheet, *row, *column)) == value)
    })
}
fn cell_value(sheet: &SheetDump, row: usize, col: usize) -> serde_json::Value {
    let address = cell_address(row, col);
    sheet
        .values
        .get(&address)
        .cloned()
        .unwrap_or(serde_json::Value::Null)
}
/// Reproduces the old report code's different openpyxl modes for current and
/// previous workbooks: current formulas are read as formula text (and parse as
/// zero in `_to_num`), while previous formulas use their stored cached values.
fn numeric_with_formula_mode(sheet: &SheetDump, row: usize, col: usize, data_only: bool) -> f64 {
    let address = cell_address(row, col);
    if !data_only && sheet.formulas.contains_key(&address) {
        return 0.0;
    }
    numeric(&cell_value(sheet, row, col))
}
fn value_for_cell(value: &serde_json::Value) -> CellValue {
    match value.get("type").and_then(serde_json::Value::as_str) {
        Some("string" | "datetime_iso" | "duration_iso") => CellValue::Text(
            value
                .get("value")
                .and_then(serde_json::Value::as_str)
                .unwrap_or_default()
                .to_owned(),
        ),
        Some("boolean") => CellValue::Boolean(
            value
                .get("value")
                .and_then(serde_json::Value::as_bool)
                .unwrap_or(false),
        ),
        Some("number" | "integer" | "datetime_serial") => CellValue::Decimal(
            value
                .get("value")
                .map(ToString::to_string)
                .unwrap_or_else(|| "0".into()),
        ),
        Some("error") => CellValue::Error(
            value
                .get("value")
                .and_then(serde_json::Value::as_str)
                .unwrap_or_default()
                .into(),
        ),
        _ => CellValue::Blank,
    }
}
fn numeric(value: &serde_json::Value) -> f64 {
    match value.get("type").and_then(serde_json::Value::as_str) {
        Some("number" | "integer" | "datetime_serial") => value
            .get("value")
            .and_then(serde_json::Value::as_f64)
            .unwrap_or(0.0),
        Some("boolean") => {
            if value
                .get("value")
                .and_then(serde_json::Value::as_bool)
                .unwrap_or(false)
            {
                1.0
            } else {
                0.0
            }
        }
        Some("string" | "datetime_iso" | "duration_iso") => value
            .get("value")
            .and_then(serde_json::Value::as_str)
            .unwrap_or("")
            .trim()
            .replace(',', "")
            .parse()
            .unwrap_or(0.0),
        _ => 0.0,
    }
}

/// Mirrors the legacy Python `_to_num` used by the village-bank splitter.
/// In particular, numeric-looking strings with separators such as `"1,000"`
/// are not accepted by Python's `float()` and therefore become zero.
fn village_numeric(value: &serde_json::Value) -> f64 {
    match value.get("type").and_then(serde_json::Value::as_str) {
        Some("number" | "integer") => value
            .get("value")
            .and_then(serde_json::Value::as_f64)
            .unwrap_or(0.0),
        Some("boolean") => {
            if value
                .get("value")
                .and_then(serde_json::Value::as_bool)
                .unwrap_or(false)
            {
                1.0
            } else {
                0.0
            }
        }
        Some("string") => value
            .get("value")
            .and_then(serde_json::Value::as_str)
            .unwrap_or("")
            .trim()
            .parse()
            .unwrap_or(0.0),
        _ => 0.0,
    }
}

fn numeric_plan(value: &CellValue) -> f64 {
    match value {
        CellValue::Decimal(value) | CellValue::Text(value) => {
            value.trim().replace(',', "").parse::<f64>().unwrap_or(0.0)
        }
        CellValue::Boolean(value) => f64::from(*value),
        CellValue::DateTime { serial, .. } => serial.parse().unwrap_or(0.0),
        CellValue::Error(_) | CellValue::Blank => 0.0,
    }
}
fn text(sheet: &SheetDump, row: usize, col: usize) -> String {
    let value = cell_value(sheet, row, col);
    match value.get("type").and_then(serde_json::Value::as_str) {
        Some("string" | "datetime_iso" | "duration_iso" | "error") => value
            .get("value")
            .and_then(serde_json::Value::as_str)
            .unwrap_or_default()
            .into(),
        Some("number" | "integer") => value
            .get("value")
            .map(ToString::to_string)
            .unwrap_or_default(),
        Some("boolean") => value
            .get("value")
            .and_then(serde_json::Value::as_bool)
            .map(|b| if b { "True" } else { "False" }.into())
            .unwrap_or_default(),
        _ => String::new(),
    }
}
fn normalize(text: &str) -> String {
    text.chars()
        .map(|c| {
            if matches!(c, '\r' | '\n' | '\t' | '　' | '\u{00a0}') {
                ' '
            } else {
                c
            }
        })
        .collect::<String>()
        .split_whitespace()
        .collect::<Vec<_>>()
        .join(" ")
}
fn yoy(current: f64, previous: f64) -> CellValue {
    if previous == 0.0 {
        if current == 0.0 {
            CellValue::Decimal("0".into())
        } else {
            CellValue::Text("∞".into())
        }
    } else {
        CellValue::Decimal(decimal_text(((current / previous) - 1.0) * 100.0))
    }
}
fn decimal_text(value: f64) -> String {
    format!("{:.2}", value)
}
fn exact_decimal_text(value: f64) -> String {
    value.to_string()
}
fn used_rows(sheet: &SheetDump) -> usize {
    sheet
        .values
        .keys()
        .chain(sheet.formulas.keys())
        .filter_map(|address| parse_a1(address).map(|value| value.0))
        .max()
        .unwrap_or(0)
}
fn used_columns(sheet: &SheetDump) -> usize {
    sheet
        .values
        .keys()
        .chain(sheet.formulas.keys())
        .filter_map(|address| parse_a1(address).map(|value| value.1))
        .max()
        .unwrap_or(1)
}
fn parse_a1(value: &str) -> Option<(usize, usize)> {
    let split = value.find(|c: char| c.is_ascii_digit())?;
    let col = value[..split].bytes().try_fold(0usize, |acc, b| {
        acc.checked_mul(26)?
            .checked_add((b.to_ascii_uppercase() - b'A' + 1) as usize)
    })?;
    Some((value[split..].parse().ok()?, col))
}
fn cell_address(row: usize, mut col: usize) -> String {
    let mut letters = String::new();
    while col > 0 {
        letters.insert(0, (b'A' + ((col - 1) % 26) as u8) as char);
        col = (col - 1) / 26;
    }
    format!("{letters}{row}")
}
fn descending_runs(mut rows: Vec<usize>) -> Vec<(usize, usize)> {
    rows.sort_unstable();
    rows.dedup();
    let mut runs = Vec::new();
    let mut iter = rows.into_iter().rev().peekable();
    while let Some(end) = iter.next() {
        let mut start = end;
        while iter.peek().is_some_and(|row| *row + 1 == start) {
            start = iter.next().expect("peeked row");
        }
        runs.push((start, end));
    }
    runs
}
fn make_plan(
    input: &Path,
    output: &Path,
    operations: Vec<EditOperation>,
    request: &str,
) -> Result<Option<ExcelEditPlan>> {
    if operations.is_empty() {
        return Ok(None);
    }
    anyhow::ensure!(
        input
            .extension()
            .and_then(|e| e.to_str())
            .is_some_and(|e| e.eq_ignore_ascii_case("xlsx"))
            && output
                .extension()
                .and_then(|e| e.to_str())
                .is_some_and(|e| e.eq_ignore_ascii_case("xlsx")),
        "batch edit plans currently accept .xlsx only"
    );
    let preserve = BTreeSet::from([
        PreservationRequirement::CellValuesAndTypes,
        PreservationRequirement::FormulaText,
        PreservationRequirement::CommentsAndPositions,
        PreservationRequirement::Styles,
        PreservationRequirement::MergedRanges,
        PreservationRequirement::DefinedNames,
        PreservationRequirement::RowAndColumnDimensions,
        PreservationRequirement::PageSettings,
        PreservationRequirement::HeaderFooterText,
        PreservationRequirement::Relationships,
        PreservationRequirement::UnmodifiedPackageParts,
        PreservationRequirement::SheetIdentityAndOrder,
    ]);
    let plan = ExcelEditPlan {
        request_id: format!("pytools-{request}"),
        input: input.into(),
        output: output.into(),
        format: WorkbookFormat::Xlsx,
        workbook: inspect_xlsx_profile(input)?,
        operations,
        preserve,
        drop_allowed: Default::default(),
    };
    plan.validate()?;
    Ok(Some(plan))
}
