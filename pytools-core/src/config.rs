//! Business configuration types loaded from the tools configuration workbook.

use crate::path_map::{
    parse_col_spec_value as parse_column_spec, parse_col_specs_value as parse_column_specs,
    ColumnSpecValue,
};
use anyhow::{anyhow, Context, Result};
use excel_reader::{read_workbook, SheetDump, WorkbookDump};
use serde_json::Value;
use std::collections::{BTreeMap, BTreeSet};
use std::path::{Component, Path, PathBuf};

const SHEET_GLOBAL: &str = "全局配置";
const SHEET_TIMELINE_RULE: &str = "时序提取规则";
const SHEET_PATH_MAP: &str = "路径标准化映射";
const SHEET_DEDUP_TASK: &str = "去重追加数据配置";
const SHEET_DEDUP_TASK_LEGACY: &str = "数据处理配置";
const SHEET_PRINT_CONFIG: &str = "打印配置";
const SHEET_INSTITUTION_MAPPING: &str = "机构映射表";
const SHEET_EXTRACT_CONFIG: &str = "工作表提取";

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GlobalConfig {
    pub input_dir: PathBuf,
    pub output_dir: PathBuf,
    pub log_dir: PathBuf,
    pub error_policy: String,
    pub default_encoding: String,
    pub source_exts: Vec<String>,
    pub timeline_rule_match_mode: String,
    /// 开发期审核用的计划 JSON 等附带文件是否落盘；实际使用可在全局配置
    /// 以 `输出附带计划文件=否` 关闭，仅保留候选副本等核心输出。
    pub emit_plan_files: bool,
}

impl GlobalConfig {
    pub fn new(input_dir: PathBuf, output_dir: PathBuf, log_dir: PathBuf) -> Self {
        Self {
            input_dir,
            output_dir,
            log_dir,
            error_policy: "continue".to_owned(),
            default_encoding: "utf-8".to_owned(),
            source_exts: vec![".xlsx".to_owned(), ".xlsm".to_owned(), ".csv".to_owned()],
            timeline_rule_match_mode: "all_match".to_owned(),
            emit_plan_files: true,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, serde::Serialize)]
pub struct DedupTask {
    pub row_no: usize,
    pub enabled: bool,
    pub task_name: String,
    pub source_wb: Option<PathBuf>,
    pub source_ws: String,
    pub key_cols: Vec<usize>,
    pub target_wb: Option<PathBuf>,
    pub target_ws: String,
    pub write_mode: String,
}

#[derive(Debug, Clone, PartialEq, Eq, serde::Serialize)]
pub struct PrintTask {
    pub row_no: usize,
    pub enabled: bool,
    pub mode: i64,
    pub source_wb: Option<PathBuf>,
    pub source_ws: String,
    pub source_ranges: String,
    pub target_wb: Option<PathBuf>,
    pub target_ws: String,
    pub fit_wide: i64,
    pub fit_tall: i64,
    pub orientation: String,
    pub center_h: bool,
    pub center_v: bool,
    pub drop_comments: bool,
    pub hide_zero: bool,
    pub remark: String,
}

#[derive(Debug, Clone, PartialEq, Eq, serde::Serialize)]
pub struct ExtractTask {
    pub row_no: usize,
    pub workbook_keywords: String,
    pub sheet_keywords: String,
    pub exclude_keywords: String,
    pub area_spec: String,
    pub stitch_mode: String,
    pub output_folder: String,
    pub output_name: String,
    pub target_sheet: String,
    pub remark: String,
}

/// Load the global key/value settings from the configuration workbook.
///
/// Relative input/output/log paths use the same base as the Python tools:
/// the parent directory of the `PyCode` directory containing the workbook.
pub fn load_global(config_path: impl AsRef<Path>) -> Result<GlobalConfig> {
    let config_path = config_path.as_ref();
    let workbook = read_workbook(config_path)
        .with_context(|| format!("read configuration workbook {}", config_path.display()))?;
    let sheet = required_sheet(&workbook, SHEET_GLOBAL)?;
    let mut values = BTreeMap::new();
    for row in 2..=max_row(sheet) {
        let key = cell_text(sheet, row, 1);
        if !key.is_empty() {
            values.insert(key, cell_text(sheet, row, 2));
        }
    }

    let base = config_path
        .parent()
        .and_then(Path::parent)
        .unwrap_or_else(|| Path::new("."));
    let base = if base.is_absolute() {
        base.to_path_buf()
    } else {
        std::env::current_dir()?.join(base)
    };
    let resolve = |key: &str, default: &str| {
        resolve_config_path(
            &base,
            values.get(key).map(String::as_str).unwrap_or(""),
            default,
        )
    };

    let mut source_exts = split(values.get("源文件扩展名").map(String::as_str).unwrap_or(""));
    if source_exts.is_empty() {
        source_exts = vec![".xlsx".to_owned(), ".xlsm".to_owned(), ".csv".to_owned()];
    }
    for extension in &mut source_exts {
        *extension = extension.to_lowercase();
        if !extension.starts_with('.') {
            extension.insert(0, '.');
        }
    }

    let mode = values
        .get("时序规则命中策略")
        .filter(|value| !value.trim().is_empty())
        .or_else(|| values.get("规则命中策略"))
        .map(|value| value.trim().to_lowercase())
        .unwrap_or_default();
    let timeline_rule_match_mode = match mode.as_str() {
        "first_match" => "first_match".to_owned(),
        "all_match" => "all_match".to_owned(),
        _ => "all_match".to_owned(),
    };

    Ok(GlobalConfig {
        input_dir: resolve("输入目录", "input"),
        output_dir: resolve("输出目录", "output"),
        log_dir: resolve("日志目录", "logs"),
        error_policy: values
            .get("错误策略")
            .filter(|value| !value.is_empty())
            .cloned()
            .unwrap_or_else(|| "continue".to_owned()),
        default_encoding: values
            .get("默认编码")
            .filter(|value| !value.is_empty())
            .cloned()
            .unwrap_or_else(|| "utf-8".to_owned()),
        source_exts,
        timeline_rule_match_mode,
        emit_plan_files: values
            .get("输出附带计划文件")
            .map(|value| !matches!(value.trim(), "否" | "N" | "n" | "false" | "FALSE"))
            .unwrap_or(true),
    })
}

/// Return one global setting, using the first matching key just like the source loader.
pub fn load_global_value(
    config_path: impl AsRef<Path>,
    key: &str,
    default: &str,
) -> Result<String> {
    let config_path = config_path.as_ref();
    let workbook = read_workbook(config_path)
        .with_context(|| format!("read configuration workbook {}", config_path.display()))?;
    let sheet = required_sheet(&workbook, SHEET_GLOBAL)?;
    for row in 2..=max_row(sheet) {
        if cell_text(sheet, row, 1) == key {
            return Ok(cell_text(sheet, row, 2));
        }
    }
    Ok(default.to_owned())
}

/// Load enabled timeline extraction rules from the configuration workbook.
pub fn load_timeline_rules(config_path: impl AsRef<Path>) -> Result<Vec<TimelineRule>> {
    let config_path = config_path.as_ref();
    let workbook = read_workbook(config_path)
        .with_context(|| format!("read configuration workbook {}", config_path.display()))?;
    let sheet = required_sheet(&workbook, SHEET_TIMELINE_RULE)?;
    let headers = header_columns(sheet);
    let has_target_write_column = headers.contains_key("启用目标写入");
    let mut rules = Vec::new();

    for row in 2..=max_row(sheet) {
        if !truthy(&field_text(sheet, &headers, row, "是否启用")) {
            continue;
        }

        let col_header_raw = field_text(sheet, &headers, row, "列表头行");
        let mut col_header_rows = col_header_raw
            .replace('，', ",")
            .replace('；', ";")
            .replace(' ', ",")
            .replace(';', ",")
            .split(',')
            .filter_map(|token| to_int(token, 0))
            .filter(|number| *number > 0)
            .map(|number| number as usize)
            .collect::<Vec<_>>();
        if col_header_rows.is_empty() {
            col_header_rows.push(to_int(&col_header_raw, 1).unwrap_or(1).max(0) as usize);
        }

        let target_wb = field_text(sheet, &headers, row, "目标工作簿路径");
        let row_header_col_value = field_value(sheet, &headers, row, "行头列");
        let row_header_cols = parse_col_specs_value(row_header_col_value);
        let row_header_col = row_header_cols.first().copied().unwrap_or(0).max(1);
        let row_header_col_raw = value_text(row_header_col_value);
        let data_col_start = parse_col_spec_value(field_value(sheet, &headers, row, "数据起始列"));
        let data_col_end =
            match parse_col_spec_value(field_value(sheet, &headers, row, "数据结束列")) {
                0 => None,
                column => Some(column),
            };
        let (set_items, set_parse_error) =
            parse_set_items(&field_text(sheet, &headers, row, "set区域"));
        let target_write_enabled = if has_target_write_column {
            truthy(&field_text(sheet, &headers, row, "启用目标写入"))
        } else {
            true
        };

        rules.push(TimelineRule {
            enabled: true,
            name: field_text(sheet, &headers, row, "规则名称"),
            wb_keyword: field_text(sheet, &headers, row, "工作簿关键字"),
            sheet_keyword: field_text(sheet, &headers, row, "工作表关键字"),
            row_header_col,
            col_header_row: col_header_rows[0],
            row_header_col_specified: !row_header_col_raw.trim().is_empty()
                && !row_header_cols.is_empty(),
            row_header_cols,
            col_header_rows,
            required_col_headers: split(&field_text(sheet, &headers, row, "必含列头")),
            required_row_headers: split(&field_text(sheet, &headers, row, "必含行头")),
            data_row_start: to_int(&field_text(sheet, &headers, row, "数据起始行"), 0)
                .unwrap_or(0)
                .max(0) as usize,
            data_row_end: to_optional_positive_int(&field_text(sheet, &headers, row, "数据结束行")),
            data_col_start,
            data_col_end,
            skip_keywords: split(&field_text(sheet, &headers, row, "跳过关键字")),
            target_wb_path: (!target_wb.is_empty()).then(|| PathBuf::from(target_wb)),
            target_sheet: non_empty_option(field_text(sheet, &headers, row, "目标工作表")),
            target_write_enabled,
            set_items,
            set_parse_error,
            target_dedup_cols: split(&field_text(sheet, &headers, row, "目标去重列")),
            output_row_no: truthy(&field_text(sheet, &headers, row, "是否输出行序号")),
        });
    }
    Ok(rules)
}

/// Load enabled path-standardization mappings. The sheet is optional in the source project.
pub fn load_path_maps(config_path: impl AsRef<Path>) -> Result<Vec<PathMapRule>> {
    let config_path = config_path.as_ref();
    let workbook = read_workbook(config_path)
        .with_context(|| format!("read configuration workbook {}", config_path.display()))?;
    let Some(sheet) = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == SHEET_PATH_MAP)
    else {
        return Ok(Vec::new());
    };
    let headers = header_columns(sheet);
    let mut maps = Vec::new();
    for row in 2..=max_row(sheet) {
        if !truthy(&field_text(sheet, &headers, row, "是否启用")) {
            continue;
        }
        let target = field_text(sheet, &headers, row, "作用对象");
        let match_mode = field_text(sheet, &headers, row, "匹配方式");
        maps.push(PathMapRule {
            enabled: true,
            name: field_text(sheet, &headers, row, "映射名称"),
            applicable_rule: field_text(sheet, &headers, row, "适用规则名"),
            wb_keyword: field_text(sheet, &headers, row, "工作簿关键字"),
            sheet_keyword: field_text(sheet, &headers, row, "工作表关键字"),
            target: if target.is_empty() {
                "两者".to_owned()
            } else {
                target
            },
            match_mode: if match_mode.is_empty() {
                "精确".to_owned()
            } else {
                match_mode
            },
            original: field_text(sheet, &headers, row, "原始路径"),
            standard: field_text(sheet, &headers, row, "标准路径"),
        });
    }
    Ok(maps)
}

/// Load de-duplication tasks from the current sheet, falling back to the
/// legacy sheet when the current sheet is absent.
pub fn load_dedup_tasks(config_path: impl AsRef<Path>) -> Result<Vec<DedupTask>> {
    let config_path = config_path.as_ref();
    let workbook = read_workbook(config_path)
        .with_context(|| format!("read configuration workbook {}", config_path.display()))?;
    let Some(sheet) = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == SHEET_DEDUP_TASK)
        .or_else(|| {
            workbook
                .sheets
                .iter()
                .find(|sheet| sheet.name == SHEET_DEDUP_TASK_LEGACY)
        })
    else {
        return Ok(Vec::new());
    };
    let headers = header_columns(sheet);
    let enabled_col = find_column(&headers, &["是否启用", "启用"]);
    let task_name_col = find_column(&headers, &["任务名", "任务名称"]);
    let source_wb_col = find_column(&headers, &["源数据工作簿", "源工作簿路径", "源工作簿"]);
    let source_ws_col = find_column(&headers, &["源数据工作表", "源工作表名", "源工作表"]);
    let key_cols_col = find_column(&headers, &["标识列序号", "去重列序号", "去重列"]);
    let target_wb_col = find_column(&headers, &["目标工作簿", "目标工作簿路径"]);
    let target_ws_col = find_column(&headers, &["目标工作表", "目标工作表名"]);
    let write_mode_col = find_column(&headers, &["写入方式", "覆盖方式"]);
    let base = config_path_base(config_path)?;
    let mut tasks = Vec::new();

    for row in 2..=max_row(sheet) {
        let source_wb_text = column_text(sheet, source_wb_col, row);
        let target_wb_text = column_text(sheet, target_wb_col, row);
        tasks.push(DedupTask {
            row_no: row,
            enabled: enabled_col
                .map(|column| cell_truthy(sheet, row, column))
                .unwrap_or(true),
            task_name: column_text(sheet, task_name_col, row),
            source_wb: resolve_optional_task_path(&base, &source_wb_text),
            source_ws: column_text(sheet, source_ws_col, row),
            key_cols: parse_key_columns(&column_text(sheet, key_cols_col, row)),
            target_wb: resolve_optional_task_path(&base, &target_wb_text),
            target_ws: column_text(sheet, target_ws_col, row),
            write_mode: column_text(sheet, write_mode_col, row),
        });
    }
    Ok(tasks)
}

/// Load print preparation settings. This reads configuration only; rendering
/// and PDF export are outside this API.
pub fn load_print_tasks(config_path: impl AsRef<Path>) -> Result<Vec<PrintTask>> {
    let config_path = config_path.as_ref();
    let workbook = read_workbook(config_path)
        .with_context(|| format!("read configuration workbook {}", config_path.display()))?;
    let Some(sheet) = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == SHEET_PRINT_CONFIG)
    else {
        return Ok(Vec::new());
    };
    let headers = header_columns(sheet);
    let base = config_path_base(config_path)?;
    let mut tasks = Vec::new();
    for row in 2..=max_row(sheet) {
        let source_wb_text = field_text(sheet, &headers, row, "源工作簿");
        let target_wb_text = field_text(sheet, &headers, row, "目标工作簿");
        tasks.push(PrintTask {
            row_no: row,
            enabled: field_truthy(sheet, &headers, row, "是否启用"),
            mode: to_int(&field_text(sheet, &headers, row, "打印模式"), 0).unwrap_or(0),
            source_wb: resolve_optional_task_path(&base, &source_wb_text),
            source_ws: field_text(sheet, &headers, row, "源工作表"),
            source_ranges: field_text(sheet, &headers, row, "源工作表打印区域"),
            target_wb: resolve_optional_task_path(&base, &target_wb_text),
            target_ws: field_text(sheet, &headers, row, "目标工作表"),
            fit_wide: to_int(&field_text(sheet, &headers, row, "FitToPagesWide"), 1).unwrap_or(1),
            fit_tall: to_int(&field_text(sheet, &headers, row, "FitToPagesTall"), 1).unwrap_or(1),
            orientation: field_text(sheet, &headers, row, "打印方向"),
            center_h: field_truthy(sheet, &headers, row, "水平居中"),
            center_v: field_truthy(sheet, &headers, row, "垂直居中"),
            drop_comments: field_truthy(sheet, &headers, row, "不输出批注"),
            hide_zero: field_truthy(sheet, &headers, row, "零值不输出"),
            remark: field_text(sheet, &headers, row, "备注"),
        });
    }
    Ok(tasks)
}

/// Load enabled, non-empty worksheet extraction tasks.
pub fn load_extract_tasks(config_path: impl AsRef<Path>) -> Result<Vec<ExtractTask>> {
    let config_path = config_path.as_ref();
    let workbook = read_workbook(config_path)
        .with_context(|| format!("read configuration workbook {}", config_path.display()))?;
    let Some(sheet) = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == SHEET_EXTRACT_CONFIG)
    else {
        return Ok(Vec::new());
    };
    let headers = header_columns(sheet);
    let mut tasks = Vec::new();
    for row in 2..=max_row(sheet) {
        if !field_truthy(sheet, &headers, row, "是否启用") {
            continue;
        }
        let workbook_keywords = field_text(sheet, &headers, row, "工作簿关键字");
        let sheet_keywords = field_text(sheet, &headers, row, "工作表关键字");
        let exclude_keywords = field_text(sheet, &headers, row, "排除关键字");
        let area_spec = field_text(sheet, &headers, row, "提取区域");
        if workbook_keywords.is_empty()
            && sheet_keywords.is_empty()
            && exclude_keywords.is_empty()
            && area_spec.is_empty()
        {
            continue;
        }
        tasks.push(ExtractTask {
            row_no: row,
            workbook_keywords,
            sheet_keywords,
            exclude_keywords,
            area_spec,
            stitch_mode: field_text(sheet, &headers, row, "多区域拼接方式"),
            output_folder: field_text(sheet, &headers, row, "输出文件夹"),
            output_name: field_text(sheet, &headers, row, "文件名"),
            target_sheet: field_text(sheet, &headers, row, "目标工作表"),
            remark: field_text(sheet, &headers, row, "备注"),
        });
    }
    Ok(tasks)
}

/// Return the institution names marked as foreign banks in column C.
pub fn load_foreign_banks(config_path: impl AsRef<Path>) -> Result<BTreeSet<String>> {
    let config_path = config_path.as_ref();
    let workbook = read_workbook(config_path)
        .with_context(|| format!("read configuration workbook {}", config_path.display()))?;
    let Some(sheet) = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == SHEET_INSTITUTION_MAPPING)
    else {
        return Ok(BTreeSet::new());
    };
    if max_column(sheet) < 3 {
        return Ok(BTreeSet::new());
    }
    let mut names = BTreeSet::new();
    for row in 2..=max_row(sheet) {
        let name = cell_text(sheet, row, 2);
        if name.is_empty() {
            continue;
        }
        if to_int(&cell_text(sheet, row, 3), 0).unwrap_or(0) == 1 {
            names.insert(name);
        }
    }
    Ok(names)
}

/// Load the original-to-standard institution mapping from columns A and B.
pub fn load_institution_mapping(config_path: impl AsRef<Path>) -> Result<BTreeMap<String, String>> {
    let config_path = config_path.as_ref();
    let workbook = read_workbook(config_path)
        .with_context(|| format!("read configuration workbook {}", config_path.display()))?;
    let Some(sheet) = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == SHEET_INSTITUTION_MAPPING)
    else {
        return Ok(BTreeMap::new());
    };
    let mut mapping = BTreeMap::new();
    for row in 2..=max_row(sheet) {
        let key = cell_text(sheet, row, 1);
        let value = cell_text(sheet, row, 2);
        if !key.is_empty() && !value.is_empty() {
            mapping.insert(key, value);
        }
    }
    Ok(mapping)
}

/// Ordered institution mapping used when the first matching substring rule
/// must win, matching the configuration sheet row order.
pub fn load_institution_mapping_ordered(
    config_path: impl AsRef<Path>,
) -> Result<Vec<(String, String)>> {
    let workbook = read_workbook(config_path.as_ref()).with_context(|| {
        format!(
            "read configuration workbook {}",
            config_path.as_ref().display()
        )
    })?;
    let Some(sheet) = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == SHEET_INSTITUTION_MAPPING)
    else {
        return Ok(Vec::new());
    };
    let mut mapping: Vec<(String, String)> = Vec::new();
    let mut positions = BTreeMap::<String, usize>::new();
    for row in 2..=max_row(sheet) {
        let source = cell_text(sheet, row, 1);
        let target = cell_text(sheet, row, 2);
        if !source.is_empty() && !target.is_empty() {
            if let Some(index) = positions.get(&source).copied() {
                // Python's dict assignment replaces a duplicate key's value
                // without moving the key from its first insertion position.
                mapping[index].1 = target;
            } else {
                positions.insert(source.clone(), mapping.len());
                mapping.push((source, target));
            }
        }
    }
    Ok(mapping)
}

fn required_sheet<'a>(workbook: &'a WorkbookDump, name: &str) -> Result<&'a SheetDump> {
    workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == name)
        .ok_or_else(|| anyhow!("missing configuration sheet: {name}"))
}

fn max_row(sheet: &SheetDump) -> usize {
    sheet
        .values
        .keys()
        .filter_map(|address| parse_row(address))
        .max()
        .unwrap_or(1)
}

fn parse_row(address: &str) -> Option<usize> {
    let digits = address
        .bytes()
        .skip_while(u8::is_ascii_alphabetic)
        .collect::<Vec<_>>();
    if digits.is_empty() || !digits.iter().all(u8::is_ascii_digit) {
        return None;
    }
    std::str::from_utf8(&digits).ok()?.parse().ok()
}

fn column_letters(mut column: usize) -> String {
    let mut letters = String::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        letters.insert(0, (b'A' + remainder as u8) as char);
        column = (column - 1) / 26;
    }
    letters
}

fn cell_value(sheet: &SheetDump, row: usize, column: usize) -> Option<&Value> {
    sheet
        .values
        .get(&format!("{}{}", column_letters(column), row))
}

fn cell_text(sheet: &SheetDump, row: usize, column: usize) -> String {
    cell_value(sheet, row, column)
        .map(|value| value_text(Some(value)))
        .unwrap_or_default()
        .trim()
        .to_owned()
}

fn value_text(value: Option<&Value>) -> String {
    let Some(value) = value else {
        return String::new();
    };
    let Some(kind) = value.get("type").and_then(Value::as_str) else {
        return value.as_str().map(str::to_owned).unwrap_or_default();
    };
    let Some(data) = value.get("value") else {
        return String::new();
    };
    match kind {
        "string" | "datetime_iso" | "duration_iso" | "error" => {
            data.as_str().map(str::to_owned).unwrap_or_default()
        }
        "boolean" => data
            .as_bool()
            .map(|value| if value { "True" } else { "False" }.to_owned())
            .unwrap_or_default(),
        "integer" => data
            .as_i64()
            .map(|number| number.to_string())
            .unwrap_or_default(),
        "number" | "datetime_serial" => data
            .as_f64()
            .map(|number| {
                if number.is_finite() && number.fract() == 0.0 {
                    format!("{number:.1}")
                } else {
                    number.to_string()
                }
            })
            .unwrap_or_default(),
        _ => String::new(),
    }
}

fn max_column(sheet: &SheetDump) -> usize {
    sheet
        .values
        .keys()
        .filter_map(|address| parse_column(address))
        .max()
        .unwrap_or(1)
}

fn parse_column(address: &str) -> Option<usize> {
    let mut column = 0usize;
    let mut letters = address
        .bytes()
        .take_while(u8::is_ascii_alphabetic)
        .peekable();
    if letters.peek().is_none() {
        return None;
    }
    for letter in letters {
        column = column
            .checked_mul(26)?
            .checked_add((letter.to_ascii_uppercase() - b'A' + 1) as usize)?;
    }
    Some(column)
}

fn header_columns(sheet: &SheetDump) -> BTreeMap<String, usize> {
    let mut headers = BTreeMap::new();
    let mut duplicate_counts = BTreeMap::<String, usize>::new();
    for column in 1..=max_column(sheet) {
        let header = cell_text(sheet, 1, column);
        if header.is_empty() {
            continue;
        }
        let count = duplicate_counts.entry(header.clone()).or_default();
        *count += 1;
        let unique = if *count == 1 {
            header
        } else {
            format!("{header}.{}", *count - 1)
        };
        headers.insert(unique, column);
    }
    headers
}

fn field_value<'a>(
    sheet: &'a SheetDump,
    headers: &BTreeMap<String, usize>,
    row: usize,
    name: &str,
) -> Option<&'a Value> {
    headers
        .get(name)
        .and_then(|column| cell_value(sheet, row, *column))
}

fn field_text(
    sheet: &SheetDump,
    headers: &BTreeMap<String, usize>,
    row: usize,
    name: &str,
) -> String {
    value_text(field_value(sheet, headers, row, name))
        .trim()
        .to_owned()
}

fn find_column(headers: &BTreeMap<String, usize>, candidates: &[&str]) -> Option<usize> {
    candidates
        .iter()
        .find_map(|candidate| headers.get(*candidate).copied())
}

fn column_text(sheet: &SheetDump, column: Option<usize>, row: usize) -> String {
    column
        .map(|column| cell_text(sheet, row, column))
        .unwrap_or_default()
}

fn field_truthy(
    sheet: &SheetDump,
    headers: &BTreeMap<String, usize>,
    row: usize,
    name: &str,
) -> bool {
    value_truthy(field_value(sheet, headers, row, name))
}

fn cell_truthy(sheet: &SheetDump, row: usize, column: usize) -> bool {
    value_truthy(cell_value(sheet, row, column))
}

fn value_truthy(value: Option<&Value>) -> bool {
    let Some(value) = value else {
        return false;
    };
    match value.get("type").and_then(Value::as_str) {
        Some("integer") => value["value"].as_i64() == Some(1),
        Some("number") => value["value"].as_f64() == Some(1.0),
        _ => truthy(&value_text(Some(value))),
    }
}

fn config_path_base(config_path: &Path) -> Result<PathBuf> {
    let base = config_path
        .parent()
        .and_then(Path::parent)
        .unwrap_or_else(|| Path::new("."));
    let absolute = if base.is_absolute() {
        base.to_path_buf()
    } else {
        std::env::current_dir()?.join(base)
    };
    Ok(std::fs::canonicalize(&absolute).unwrap_or(absolute))
}

fn resolve_optional_task_path(base: &Path, path_text: &str) -> Option<PathBuf> {
    let path_text = path_text.trim();
    if path_text.is_empty() {
        return None;
    }
    let path = Path::new(path_text);
    if path.is_absolute() {
        Some(path.to_path_buf())
    } else {
        Some(resolve_config_path(base, path_text, ""))
    }
}

fn parse_key_columns(raw: &str) -> Vec<usize> {
    raw.replace('，', ";")
        .replace(',', ";")
        .replace('；', ";")
        .replace(':', ";")
        .split(';')
        .filter_map(|part| to_int(part, 0))
        .filter(|column| *column > 0)
        .map(|column| column as usize)
        .collect()
}

fn column_spec_value(value: Option<&Value>) -> ColumnSpecValue<'_> {
    let Some(value) = value else {
        return ColumnSpecValue::Missing;
    };
    match value.get("type").and_then(Value::as_str) {
        Some("string") => ColumnSpecValue::Text(value["value"].as_str().unwrap_or_default()),
        Some("integer") => ColumnSpecValue::Integer(value["value"].as_i64().unwrap_or_default()),
        Some("number") => ColumnSpecValue::Number(value["value"].as_f64().unwrap_or_default()),
        _ => ColumnSpecValue::Missing,
    }
}

fn parse_col_spec_value(value: Option<&Value>) -> usize {
    parse_column_spec(column_spec_value(value))
}

fn parse_col_specs_value(value: Option<&Value>) -> Vec<usize> {
    parse_column_specs(column_spec_value(value))
}

fn truthy(value: &str) -> bool {
    matches!(
        value.trim().to_lowercase().as_str(),
        "是" | "1" | "true" | "y" | "yes" | "on"
    )
}

fn split(value: &str) -> Vec<String> {
    value
        .replace('；', ";")
        .replace('，', ";")
        .replace(',', ";")
        .split(';')
        .map(str::trim)
        .filter(|part| !part.is_empty())
        .map(str::to_owned)
        .collect()
}

fn to_int(value: &str, default: i64) -> Option<i64> {
    let value = value.trim();
    if value.is_empty() {
        return Some(default);
    }
    value
        .parse::<f64>()
        .ok()
        .filter(|number| number.is_finite())
        .map(|number| number.trunc() as i64)
        .or(Some(default))
}

fn to_optional_positive_int(value: &str) -> Option<usize> {
    let value = value.trim();
    if value.is_empty() {
        return None;
    }
    value
        .parse::<f64>()
        .ok()
        .filter(|number| number.is_finite())
        .map(|number| number.trunc() as i64)
        .filter(|number| *number > 0)
        .map(|number| number as usize)
}

fn parse_set_items(raw: &str) -> (Vec<(String, String)>, String) {
    let text = raw.trim();
    if text.is_empty() {
        return (Vec::new(), String::new());
    }
    for part in text
        .replace('；', ";")
        .split(';')
        .map(str::trim)
        .filter(|part| !part.is_empty())
    {
        let Some((alias, address)) = part.split_once('@') else {
            return (
                Vec::new(),
                format!("set区域项格式非法: {part}（应为 别名@地址）"),
            );
        };
        let alias = alias.trim();
        let address = address.trim().to_uppercase();
        if alias.is_empty() {
            return (Vec::new(), format!("set区域项别名为空: {part}"));
        }
        if !valid_a1_cell(&address) {
            return (Vec::new(), format!("set区域项地址非法: {part}"));
        }
    }
    let items = text
        .replace('；', ";")
        .split(';')
        .map(str::trim)
        .filter(|part| !part.is_empty())
        .filter_map(|part| part.split_once('@'))
        .map(|(alias, address)| (alias.trim().to_owned(), address.trim().to_uppercase()))
        .collect();
    (items, String::new())
}

fn valid_a1_cell(address: &str) -> bool {
    let letters = address.bytes().take_while(u8::is_ascii_alphabetic).count();
    let bytes = address.as_bytes();
    letters > 0
        && letters < bytes.len()
        && (b'1'..=b'9').contains(&bytes[letters])
        && bytes[letters..].iter().all(u8::is_ascii_digit)
}

fn non_empty_option(value: String) -> Option<String> {
    (!value.is_empty()).then_some(value)
}

fn resolve_config_path(base: &Path, configured: &str, default: &str) -> PathBuf {
    let value = if configured.trim().is_empty() {
        default
    } else {
        configured.trim()
    };
    let path = Path::new(value);
    let joined = if path.is_absolute() {
        path.to_path_buf()
    } else {
        base.join(path)
    };
    let mut normalized = PathBuf::new();
    for component in joined.components() {
        match component {
            Component::CurDir => {}
            Component::ParentDir => {
                if !normalized.pop() && !joined.is_absolute() {
                    normalized.push(component.as_os_str());
                }
            }
            other => normalized.push(other.as_os_str()),
        }
    }
    std::fs::canonicalize(&normalized).unwrap_or(normalized)
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    fn normalized_path(path: &Path, base: &Path) -> String {
        let relative = path
            .strip_prefix(base)
            .unwrap_or_else(|_| panic!("{} is not under {}", path.display(), base.display()));
        format!("$BASE/{}", relative.to_string_lossy().replace('\\', "/"))
    }

    #[test]
    fn config_workbook_matches_original_python_loader_report() {
        let config_path =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/config_loader.xlsx");
        let base = config_path.parent().unwrap().parent().unwrap();
        let global = load_global(&config_path).unwrap();
        let rules = load_timeline_rules(&config_path).unwrap();
        let maps = load_path_maps(&config_path).unwrap();

        let global_json = json!({
            "input_dir": normalized_path(&global.input_dir, base),
            "output_dir": normalized_path(&global.output_dir, base),
            "log_dir": normalized_path(&global.log_dir, base),
            "error_policy": global.error_policy,
            "default_encoding": global.default_encoding,
            "source_exts": global.source_exts,
            "timeline_rule_match_mode": global.timeline_rule_match_mode,
        });
        let rules_json = rules
            .iter()
            .map(|rule| {
                json!({
                    "enabled": rule.enabled,
                    "name": rule.name,
                    "wb_keyword": rule.wb_keyword,
                    "sheet_keyword": rule.sheet_keyword,
                    "row_header_col": rule.row_header_col,
                    "col_header_row": rule.col_header_row,
                    "row_header_cols": rule.row_header_cols,
                    "row_header_col_specified": rule.row_header_col_specified,
                    "col_header_rows": rule.col_header_rows,
                    "required_col_headers": rule.required_col_headers,
                    "required_row_headers": rule.required_row_headers,
                    "data_row_start": rule.data_row_start,
                    "data_row_end": rule.data_row_end,
                    "data_col_start": rule.data_col_start,
                    "data_col_end": rule.data_col_end,
                    "skip_keywords": rule.skip_keywords,
                    "target_wb_path": rule.target_wb_path.as_ref().map(|path| path.to_string_lossy()),
                    "target_sheet": rule.target_sheet,
                    "target_write_enabled": rule.target_write_enabled,
                    "set_items": rule.set_items,
                    "set_parse_error": rule.set_parse_error,
                    "target_dedup_cols": rule.target_dedup_cols,
                    "output_row_no": rule.output_row_no,
                })
            })
            .collect::<Vec<_>>();
        let maps_json = maps
            .iter()
            .map(|map| {
                json!({
                    "enabled": map.enabled,
                    "name": map.name,
                    "applicable_rule": map.applicable_rule,
                    "wb_keyword": map.wb_keyword,
                    "sheet_keyword": map.sheet_keyword,
                    "target": map.target,
                    "match_mode": map.match_mode,
                    "original": map.original,
                    "standard": map.standard,
                })
            })
            .collect::<Vec<_>>();

        let actual = json!({
            "global": global_json,
            "global_value_existing": load_global_value(&config_path, "错误策略", "").unwrap(),
            "global_value_missing": load_global_value(&config_path, "missing", "fallback").unwrap(),
            "timeline_rules": rules_json,
            "path_maps": maps_json,
        });
        let baseline: Value =
            serde_json::from_str(include_str!("../probe/reports/config_loader_baseline.json"))
                .expect("Python baseline report must be valid JSON");

        assert_eq!(baseline["python_baseline"]["status"], "保留");
        assert_eq!(actual, baseline["result"]);
    }
}

#[derive(Debug, Clone, PartialEq, Eq, serde::Serialize)]
pub struct TimelineRule {
    pub enabled: bool,
    pub name: String,
    pub wb_keyword: String,
    pub sheet_keyword: String,
    pub row_header_col: usize,
    pub col_header_row: usize,
    pub row_header_cols: Vec<usize>,
    pub row_header_col_specified: bool,
    pub col_header_rows: Vec<usize>,
    pub required_col_headers: Vec<String>,
    pub required_row_headers: Vec<String>,
    pub data_row_start: usize,
    pub data_row_end: Option<usize>,
    pub data_col_start: usize,
    pub data_col_end: Option<usize>,
    pub skip_keywords: Vec<String>,
    pub target_wb_path: Option<PathBuf>,
    pub target_sheet: Option<String>,
    pub target_write_enabled: bool,
    pub set_items: Vec<(String, String)>,
    pub set_parse_error: String,
    pub target_dedup_cols: Vec<String>,
    pub output_row_no: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, serde::Serialize)]
pub struct PathMapRule {
    pub enabled: bool,
    pub name: String,
    pub applicable_rule: String,
    pub wb_keyword: String,
    pub sheet_keyword: String,
    /// `行头`, `列头`, `两者`, or empty for both.
    pub target: String,
    /// `包含` performs substring replacement; other values use exact matching.
    pub match_mode: String,
    pub original: String,
    pub standard: String,
}
