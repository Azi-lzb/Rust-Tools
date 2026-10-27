//! Pure rule and header helpers used by workbook extraction workflows.

use crate::{
    config::{PathMapRule, TimelineRule},
    date_parse::parse_data_date,
    path_map::match_all_keywords,
};
use std::collections::{BTreeMap, HashMap};

pub trait SheetValue: Clone {
    fn to_sheet_text(&self) -> String;

    fn is_missing(&self) -> bool {
        false
    }
}

impl SheetValue for String {
    fn to_sheet_text(&self) -> String {
        self.clone()
    }
}

impl SheetValue for bool {
    fn to_sheet_text(&self) -> String {
        (if *self { "True" } else { "False" }).to_owned()
    }
}

macro_rules! impl_integer_sheet_value {
    ($($value_type:ty),* $(,)?) => {
        $(
            impl SheetValue for $value_type {
                fn to_sheet_text(&self) -> String {
                    self.to_string()
                }
            }
        )*
    };
}

impl_integer_sheet_value!(i8, i16, i32, i64, isize, u8, u16, u32, u64, usize);

impl SheetValue for f32 {
    fn to_sheet_text(&self) -> String {
        format!("{self:?}")
    }

    fn is_missing(&self) -> bool {
        self.is_nan()
    }
}

impl SheetValue for f64 {
    fn to_sheet_text(&self) -> String {
        format!("{self:?}")
    }

    fn is_missing(&self) -> bool {
        self.is_nan()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct MergeRange {
    pub min_row: usize,
    pub max_row: usize,
    pub min_col: usize,
    pub max_col: usize,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ExtractedCell<T> {
    pub source_file: String,
    pub sheet_name: String,
    pub rule_name: String,
    pub data_date: String,
    pub row_path: String,
    pub col_path: String,
    pub value: T,
    pub cell_addr: String,
}

pub fn column_letter(mut column_1_based: usize) -> String {
    let mut letters = Vec::new();
    while column_1_based > 0 {
        let remainder = (column_1_based - 1) % 26;
        letters.push((b'A' + remainder as u8) as char);
        column_1_based = (column_1_based - 1) / 26;
    }
    letters.into_iter().rev().collect()
}

pub fn cell_address(row_1_based: usize, column_1_based: usize) -> String {
    format!("{}{}", column_letter(column_1_based), row_1_based)
}

pub fn fallback_column_path(column_0_based: usize) -> String {
    format!("列_{}", column_letter(column_0_based + 1))
}

pub fn pick_rules_for_workbook<'a>(
    rules: &'a [TimelineRule],
    workbook_name: &str,
    mode: &str,
) -> Vec<&'a TimelineRule> {
    let mut matched = rules
        .iter()
        .filter(|rule| match_all_keywords(workbook_name, &rule.wb_keyword))
        .collect::<Vec<_>>();
    if mode == "first_match" {
        matched.truncate(1);
    }
    matched
}

pub fn skip_match(skip_keywords: &[String], text: &str) -> bool {
    !text.is_empty() && skip_keywords.iter().any(|keyword| text.contains(keyword))
}

pub fn required_match(required: &[String], values: &[String]) -> bool {
    if required.is_empty() {
        return true;
    }
    let joined = values.join("|");
    required.iter().all(|keyword| joined.contains(keyword))
}

/// Build one standardized column path per column in the requested inclusive range.
/// Row/column indexes and merge coordinates are zero-based.
pub fn compute_wide_col_paths<T: SheetValue>(
    sheet: &[Vec<Option<T>>],
    merge_ranges: &[MergeRange],
    rule: &TimelineRule,
    sheet_name: &str,
    path_maps: &[PathMapRule],
    workbook_name: &str,
    c_start: usize,
    c_end: usize,
    header_rows: &[usize],
    overrides: &HashMap<(usize, usize), String>,
) -> (BTreeMap<usize, String>, Vec<String>) {
    let column_count = sheet.iter().map(Vec::len).max().unwrap_or(0);
    if c_start > c_end || c_end >= column_count || header_rows.iter().any(|row| *row >= sheet.len())
    {
        return (BTreeMap::new(), Vec::new());
    }

    let header_rows_text =
        collect_header_rows(sheet, merge_ranges, overrides, c_start, c_end, header_rows);
    compute_col_paths_from_headers(
        &header_rows_text,
        rule,
        sheet_name,
        path_maps,
        workbook_name,
        c_start,
        c_end,
        header_rows,
    )
}

fn compute_col_paths_from_headers(
    header_rows_text: &BTreeMap<usize, Vec<String>>,
    rule: &TimelineRule,
    sheet_name: &str,
    path_maps: &[PathMapRule],
    workbook_name: &str,
    c_start: usize,
    c_end: usize,
    header_rows: &[usize],
) -> (BTreeMap<usize, String>, Vec<String>) {
    let mut raw_paths = BTreeMap::new();
    for col in c_start..=c_end {
        let offset = col - c_start;
        let parts = header_rows
            .iter()
            .filter_map(|row| {
                header_rows_text
                    .get(row)
                    .and_then(|texts| texts.get(offset))
            })
            .filter(|text| !text.is_empty())
            .cloned()
            .collect::<Vec<_>>();
        raw_paths.insert(col, parts.join("_"));
    }

    let mut raw_counts = HashMap::<String, usize>::new();
    for raw in raw_paths.values().filter(|path| !path.is_empty()) {
        *raw_counts.entry(raw.as_str().to_owned()).or_default() += 1;
    }

    let mut final_paths = BTreeMap::new();
    let mut seen = HashMap::<String, usize>::new();
    for (&col, raw) in &raw_paths {
        if raw.is_empty() {
            final_paths.insert(col, fallback_column_path(col));
            continue;
        }
        let count = seen.entry(raw.as_str().to_owned()).or_default();
        *count += 1;
        let path = if raw_counts.get(raw).copied().unwrap_or(0) > 1 {
            format!("{raw}_{}", *count)
        } else {
            raw.as_str().to_owned()
        };
        final_paths.insert(col, path);
    }

    let column_maps = filter_path_maps(path_maps, "列头", &rule.name, workbook_name, sheet_name);
    let mut ordered = Vec::with_capacity(c_end - c_start + 1);
    for col in c_start..=c_end {
        let path = final_paths.get(&col).map(String::as_str).unwrap_or("");
        ordered.push(apply_path_maps(path, &column_maps));
    }
    let standardized = (c_start..=c_end).zip(ordered.iter().cloned()).collect();
    (standardized, ordered)
}

/// Expand an in-memory worksheet using the configured row/column header rules.
/// Cell values remain typed and are cloned only for cells included in the result.
pub fn extract_cells_from_matrix<T: SheetValue>(
    sheet: &[Vec<Option<T>>],
    merge_ranges: &[MergeRange],
    source_file_name: &str,
    rule: &TimelineRule,
    sheet_name: &str,
    path_maps: &[PathMapRule],
    row_suffix_enabled: bool,
    overrides: &HashMap<(usize, usize), String>,
) -> Result<Vec<ExtractedCell<T>>, &'static str> {
    let rows = sheet.len();
    let column_count = sheet.iter().map(Vec::len).max().unwrap_or(0);
    if rows == 0 || column_count == 0 {
        return Ok(Vec::new());
    }

    let a1 = cell_text(sheet, 0, 0);
    let data_date = parse_data_date(source_file_name, Some(&a1))
        .map(|parsed| parsed.date.to_string())
        .unwrap_or_default();

    let row_header_columns = if rule.row_header_col_specified {
        let configured = if rule.row_header_cols.is_empty() {
            vec![rule.row_header_col]
        } else {
            rule.row_header_cols.clone()
        };
        let configured_columns = configured
            .into_iter()
            .filter_map(|column| column.checked_sub(1))
            .collect::<Vec<_>>();
        if configured_columns.is_empty()
            || configured_columns
                .iter()
                .any(|column| *column >= column_count)
        {
            return Ok(Vec::new());
        }
        configured_columns
    } else {
        Vec::new()
    };
    if rule.row_header_col_specified && row_header_columns.is_empty() {
        return Ok(Vec::new());
    }

    let header_rows = rule
        .col_header_rows
        .iter()
        .map(|row| row.checked_sub(1))
        .collect::<Option<Vec<_>>>();
    let Some(header_rows) = header_rows else {
        return Ok(Vec::new());
    };
    if header_rows.is_empty() || header_rows.iter().any(|row| *row >= rows) {
        return Ok(Vec::new());
    }

    let row_start = if rule.data_row_start > 0 {
        rule.data_row_start - 1
    } else {
        *header_rows.iter().max().unwrap_or(&0) + 1
    };
    let row_end = rule
        .data_row_end
        .filter(|end| *end > 0)
        .map(|end| end - 1)
        .unwrap_or(rows - 1);
    let col_start = if rule.data_col_start > 0 {
        rule.data_col_start - 1
    } else if rule.row_header_col_specified {
        row_header_columns.iter().copied().max().unwrap_or(0) + 1
    } else {
        0
    };
    let col_end = rule
        .data_col_end
        .filter(|end| *end > 0)
        .map(|end| end - 1)
        .unwrap_or(column_count - 1);

    if row_start > row_end || col_start > col_end {
        return Ok(Vec::new());
    }
    // PyCode reads out-of-range header cells as empty strings, so rows beyond
    // the stored data silently drop out (empty row header), and the wide
    // pipeline clamps its configured column bound to the sheet. Real VBA
    // rules routinely configure the template's maximal row span while actual
    // reports are shorter, so a configured range past the sheet must
    // truncate instead of failing.
    let row_end = row_end.min(rows.saturating_sub(1));
    let col_end = col_end.min(column_count.saturating_sub(1));
    if row_start > row_end || col_start > col_end {
        return Ok(Vec::new());
    }

    let header_texts = collect_header_rows(
        sheet,
        merge_ranges,
        overrides,
        col_start,
        col_end,
        &header_rows,
    );

    let mut row_header_values = Vec::new();
    if rule.row_header_col_specified {
        for row in row_start..=row_end {
            let parts = row_header_columns
                .iter()
                .map(|&col| row_header_text(sheet, merge_ranges, overrides, row, col))
                .filter(|text| !text.is_empty())
                .collect::<Vec<_>>();
            row_header_values.push(parts.join("_"));
        }
    }
    let mut column_header_values = Vec::new();
    for header_row in &header_rows {
        if let Some(values) = header_texts.get(header_row) {
            column_header_values.extend(values.iter().cloned());
        }
    }
    if (rule.row_header_col_specified
        && !required_match(&rule.required_row_headers, &row_header_values))
        || !required_match(&rule.required_col_headers, &column_header_values)
    {
        return Ok(Vec::new());
    }

    let (column_paths, _) = compute_col_paths_from_headers(
        &header_texts,
        rule,
        sheet_name,
        path_maps,
        source_file_name,
        col_start,
        col_end,
        &header_rows,
    );

    let mut raw_row_paths = Vec::with_capacity(row_end - row_start + 1);
    let mut raw_row_counts = HashMap::<String, usize>::new();
    for row in row_start..=row_end {
        let raw = if rule.row_header_col_specified {
            row_header_columns
                .iter()
                .map(|&col| row_header_text(sheet, merge_ranges, overrides, row, col))
                .filter(|text| !text.is_empty())
                .collect::<Vec<_>>()
                .join("_")
        } else {
            format!("ROW#{}", row + 1)
        };
        if !raw.is_empty() {
            *raw_row_counts.entry(raw.clone()).or_default() += 1;
        }
        raw_row_paths.push(raw);
    }

    let row_maps = filter_path_maps(path_maps, "行头", &rule.name, source_file_name, sheet_name);
    let mut seen_rows = HashMap::<String, usize>::new();
    let mut extracted = Vec::new();
    for (offset, raw_path) in raw_row_paths.into_iter().enumerate() {
        let row = row_start + offset;
        if raw_path.is_empty() || skip_match(&rule.skip_keywords, &raw_path) {
            continue;
        }
        let row_path =
            if row_suffix_enabled && raw_row_counts.get(&raw_path).copied().unwrap_or(0) > 1 {
                let count = seen_rows.entry(raw_path.clone()).or_default();
                *count += 1;
                format!("{raw_path}_{}", *count)
            } else {
                raw_path
            };
        let row_path = apply_path_maps(&row_path, &row_maps);

        for col in col_start..=col_end {
            let col_path = column_paths.get(&col).map(String::as_str).unwrap_or("");
            if col_path.is_empty() {
                continue;
            }
            let Some(value) = sheet
                .get(row)
                .and_then(|cells| cells.get(col))
                .and_then(Option::as_ref)
            else {
                continue;
            };
            if value.is_missing() {
                continue;
            }
            extracted.push(ExtractedCell {
                source_file: source_file_name.to_owned(),
                sheet_name: sheet_name.to_owned(),
                rule_name: rule.name.clone(),
                data_date: data_date.clone(),
                row_path: row_path.clone(),
                col_path: col_path.to_owned(),
                value: value.clone(),
                cell_addr: cell_address(row + 1, col + 1),
            });
        }
    }
    Ok(extracted)
}

fn cell_text<T: SheetValue>(sheet: &[Vec<Option<T>>], row: usize, col: usize) -> String {
    sheet
        .get(row)
        .and_then(|cells| cells.get(col))
        .and_then(Option::as_ref)
        .filter(|value| !value.is_missing())
        .map(|value| value.to_sheet_text())
        .map(|text| text.trim().to_owned())
        .unwrap_or_default()
}

fn collect_header_rows<T: SheetValue>(
    sheet: &[Vec<Option<T>>],
    merge_ranges: &[MergeRange],
    overrides: &HashMap<(usize, usize), String>,
    c_start: usize,
    c_end: usize,
    header_rows: &[usize],
) -> BTreeMap<usize, Vec<String>> {
    header_rows
        .iter()
        .map(|&row| {
            let texts = (c_start..=c_end)
                .map(|col| header_text_with_left_fill(sheet, merge_ranges, overrides, row, col))
                .collect();
            (row, texts)
        })
        .collect()
}

fn header_cell_text<T: SheetValue>(
    sheet: &[Vec<Option<T>>],
    merge_ranges: &[MergeRange],
    row: usize,
    col: usize,
) -> String {
    if let Some(merge) = merge_ranges.iter().find(|merge| {
        merge.min_row <= row && row <= merge.max_row && merge.min_col <= col && col <= merge.max_col
    }) {
        let merged_value = cell_text(sheet, merge.min_row, merge.min_col);
        if !merged_value.is_empty() {
            return merged_value;
        }
    }
    cell_text(sheet, row, col)
}

fn header_text_with_left_fill<T: SheetValue>(
    sheet: &[Vec<Option<T>>],
    merge_ranges: &[MergeRange],
    overrides: &HashMap<(usize, usize), String>,
    row: usize,
    col: usize,
) -> String {
    if let Some(text) = overrides.get(&(row, col)) {
        return text.clone();
    }
    let text = header_cell_text(sheet, merge_ranges, row, col);
    if !text.is_empty() || !merge_ranges.is_empty() {
        return text;
    }
    (0..col)
        .rev()
        .map(|previous_col| header_cell_text(sheet, merge_ranges, row, previous_col))
        .find(|text| !text.is_empty())
        .unwrap_or_default()
}

fn row_header_text<T: SheetValue>(
    sheet: &[Vec<Option<T>>],
    merge_ranges: &[MergeRange],
    overrides: &HashMap<(usize, usize), String>,
    row: usize,
    col: usize,
) -> String {
    overrides
        .get(&(row, col))
        .cloned()
        .unwrap_or_else(|| header_cell_text(sheet, merge_ranges, row, col))
}

fn filter_path_maps<'a>(
    path_maps: &'a [PathMapRule],
    target: &str,
    rule_name: &str,
    workbook_name: &str,
    sheet_name: &str,
) -> Vec<&'a PathMapRule> {
    path_maps
        .iter()
        .filter(|map| {
            (map.applicable_rule.is_empty() || map.applicable_rule == rule_name)
                && (map.target == target || map.target == "两者" || map.target.is_empty())
                && match_all_keywords(workbook_name, &map.wb_keyword)
                && match_all_keywords(sheet_name, &map.sheet_keyword)
        })
        .collect()
}

fn apply_path_maps(path: &str, maps: &[&PathMapRule]) -> String {
    if path.is_empty() {
        return path.to_owned();
    }
    for map in maps {
        if map.match_mode == "包含" {
            if !map.original.is_empty() && path.contains(&map.original) {
                return path.replace(&map.original, &map.standard);
            }
        } else if path == map.original {
            return map.standard.clone();
        }
    }
    path.to_owned()
}

#[cfg(test)]
mod path_map_tests {
    use super::*;
    use serde_json::Value;

    fn make_map(value: &Value) -> PathMapRule {
        let text = |key: &str| {
            value
                .get(key)
                .and_then(Value::as_str)
                .unwrap_or_default()
                .to_owned()
        };
        PathMapRule {
            enabled: value
                .get("enabled")
                .and_then(Value::as_bool)
                .unwrap_or(true),
            name: text("name"),
            applicable_rule: text("applicable_rule"),
            wb_keyword: text("wb_keyword"),
            sheet_keyword: text("sheet_keyword"),
            target: text("target"),
            match_mode: text("match_mode"),
            original: text("original"),
            standard: text("standard"),
        }
    }

    #[test]
    fn standardization_matches_shared_python_vectors() {
        let vectors: Value =
            serde_json::from_str(include_str!("../tests/fixtures/path_map_cases.json")).unwrap();
        for case in vectors["standardize"].as_array().unwrap() {
            let maps = case["maps"]
                .as_array()
                .unwrap()
                .iter()
                .map(make_map)
                .collect::<Vec<_>>();
            let filtered = filter_path_maps(
                &maps,
                case["target"].as_str().unwrap(),
                case["rule_name"].as_str().unwrap(),
                case["workbook_name"].as_str().unwrap(),
                case["sheet_name"].as_str().unwrap(),
            );
            let actual = apply_path_maps(case["path"].as_str().unwrap(), &filtered);
            assert_eq!(
                actual,
                case["expected"].as_str().unwrap(),
                "path-map case {}",
                case["id"].as_str().unwrap()
            );
        }
    }
}

#[cfg(test)]
mod configured_range_boundary_tests {
    use super::*;

    #[test]
    fn configured_data_start_after_sheet_tail_returns_no_cells() {
        let sheet = vec![
            vec![Some("标题".to_owned()), Some("数据".to_owned())],
            vec![Some("地区".to_owned()), Some("余额".to_owned())],
            vec![None, None],
        ];
        let rule = TimelineRule {
            enabled: true,
            name: "short-sheet".to_owned(),
            wb_keyword: String::new(),
            sheet_keyword: String::new(),
            row_header_col: 1,
            col_header_row: 2,
            row_header_cols: vec![1],
            row_header_col_specified: true,
            col_header_rows: vec![2],
            required_col_headers: Vec::new(),
            required_row_headers: Vec::new(),
            data_row_start: 4,
            data_row_end: Some(8),
            data_col_start: 2,
            data_col_end: Some(2),
            skip_keywords: Vec::new(),
            target_wb_path: None,
            target_sheet: None,
            target_write_enabled: false,
            set_items: Vec::new(),
            set_parse_error: String::new(),
            target_dedup_cols: Vec::new(),
            output_row_no: false,
        };

        let cells = extract_cells_from_matrix(
            &sheet,
            &[],
            "sample.xlsx",
            &rule,
            "Report",
            &[],
            false,
            &HashMap::new(),
        )
        .expect("a data range starting after the sheet should produce no rows");

        assert!(cells.is_empty());
    }
}
