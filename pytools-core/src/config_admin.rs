//! Configuration workbook creation and inspection for the pytools CLI.

use anyhow::{bail, Context, Result};
use excel_editor::{
    inspect_xlsx_profile, CellValue as PlannedCellValue, EditOperation, ExcelEditPlan,
    PreservationRequirement, WorkbookFormat,
};
use excel_reader::{
    read_value_columns_with_iso_dates, read_workbook, read_xlsx_worksheet_extents, CellValue,
    SheetDump, WorksheetExtent,
};
use excel_writer::{Note, Workbook};
use serde::Serialize;
use serde_json::Value;
use std::{
    collections::BTreeMap,
    fs,
    path::{Path, PathBuf},
};

#[derive(Clone, Debug, Serialize)]
pub struct ConfigView {
    pub path: PathBuf,
    pub sheets: Vec<ConfigSheetView>,
}
#[derive(Clone, Debug, Serialize)]
pub struct ConfigSheetView {
    pub name: String,
    pub rows: Vec<Vec<Value>>,
    pub comments: BTreeMap<String, String>,
}
#[derive(Clone, Debug, Serialize)]
pub struct ConfigInitResult {
    pub created: bool,
    pub path: PathBuf,
    pub sheets: usize,
    pub message: String,
}

const SHEET_SPECS: &[(&str, &[&str])] = &[
    ("全局配置", &["键", "值", "备注"]),
    (
        "时序提取规则",
        &[
            "是否启用",
            "规则名称",
            "工作簿关键字",
            "工作表关键字",
            "行头列",
            "列表头行",
            "必含列头",
            "必含行头",
            "数据起始行",
            "数据结束行",
            "数据起始列",
            "数据结束列",
            "跳过关键字",
            "目标工作簿路径",
            "目标工作表",
            "启用目标写入",
            "set区域",
            "目标去重列",
            "是否输出行序号",
        ],
    ),
    (
        "路径标准化映射",
        &[
            "是否启用",
            "映射名称",
            "适用规则名",
            "工作簿关键字",
            "工作表关键字",
            "作用对象",
            "匹配方式",
            "原始路径",
            "标准路径",
            "备注",
        ],
    ),
    (
        "去重追加数据配置",
        &[
            "是否启用",
            "源数据工作簿",
            "源数据工作表",
            "标识列序号",
            "目标工作簿",
            "目标工作表",
            "执行模式",
            "备注",
        ],
    ),
    (
        "打印配置",
        &[
            "是否启用",
            "打印模式",
            "源工作簿",
            "源工作表",
            "源工作表打印区域",
            "目标工作簿",
            "目标工作表",
            "FitToPagesWide",
            "FitToPagesTall",
            "打印方向",
            "水平居中",
            "垂直居中",
            "不输出批注",
            "零值不输出",
            "备注",
        ],
    ),
    (
        "重命名配置",
        &[
            "简称",
            "全称",
            "占位C",
            "代码",
            "全称(代码映射)",
            "占位F",
            "键",
            "值",
            "占位I",
            "原表名",
            "新表名",
            "占位L",
            "原文件名片段",
            "新文件名片段",
            "内容重命名启用",
            "内容重命名取值规则",
            "备注",
        ],
    ),
    (
        "归档类型配置",
        &[
            "是否启用",
            "匹配关键词",
            "归档文件夹",
            "备注",
            "需要归档的后缀",
            "排除归档的后缀",
        ],
    ),
    (
        "提交检查配置",
        &[
            "是否启用",
            "日期关键词",
            "县区关键词",
            "机构关键词",
            "备用关键词",
            "命中次数",
            "命中的文件名",
            "备注",
        ],
    ),
    (
        "文件排序配置",
        &["是否启用", "排序前缀", "匹配关键词", "备注"],
    ),
    (
        "机构映射表",
        &["原始机构名称", "映射后机构名称", "是否为外资行"],
    ),
    (
        "工作表提取",
        &[
            "是否启用",
            "工作簿关键字",
            "工作表关键字",
            "排除关键字",
            "提取区域",
            "多区域拼接方式",
            "输出文件夹",
            "文件名",
            "目标工作表",
            "备注",
        ],
    ),
];

const GLOBAL_DEFAULTS: &[(&str, &str)] = &[
    ("输入目录", "input"),
    ("输出目录", "output"),
    // Blank selects run_logging::default_log_dir() at runtime; an explicit
    // user value remains relative to the configuration workbook as before.
    ("日志目录", ""),
    ("日志模式", "普通"),
    ("日志保留天数", ""),
    ("日志最大占用MB", ""),
    ("源文件扩展名", ".xlsx;.xlsm;.csv"),
    ("错误策略", "continue"),
    ("默认编码", "utf-8"),
    ("打印零值不输出", "否"),
    ("时序规则命中策略", "all_match"),
    ("excel目的格式", "xlsx"),
    ("word目的格式", "docx"),
];

/// Create a new configuration workbook without overwriting user data. Existing
/// workbooks are rejected here and must go through the ExcelModifier repair
/// plan, which is wired separately from new-file generation.
pub fn initialize_new_config(path: impl AsRef<Path>) -> Result<ConfigInitResult> {
    let path = path.as_ref();
    if path.exists() {
        bail!("configuration already exists; repair must run through ExcelModifier and will not be replaced by a new workbook: {}",path.display());
    }
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)?;
    }
    let mut workbook = Workbook::new();
    for (sheet_name, headers) in SHEET_SPECS {
        let sheet = workbook.add_worksheet();
        sheet.set_name(*sheet_name)?;
        for (col, header) in headers.iter().enumerate() {
            sheet.write_string(0, col as u16, *header)?;
            if let Some(note) = header_note(sheet_name, header) {
                // Keep the author as comment metadata; rust_xlsxwriter prefixes
                // it into note text by default, unlike the Python baseline.
                let note = Note::new(note)
                    .add_author_prefix(false)
                    .set_author("pytools");
                sheet.insert_note(0, col as u16, &note)?;
            }
        }
    }
    let global = workbook
        .worksheet_from_name("全局配置")
        .context("global worksheet missing")?;
    for (row, (key, value)) in GLOBAL_DEFAULTS.iter().enumerate() {
        global.write_string((row + 1) as u32, 0, *key)?;
        global.write_string((row + 1) as u32, 1, *value)?;
    }
    let rename = workbook
        .worksheet_from_name("重命名配置")
        .context("rename worksheet missing")?;
    rename.write_string(1, 14, "Y")?;
    rename.write_string(1, 15, "sheet1@A1;sheet2@A2")?;
    rename.write_string(1, 16, "示例：按文件内容单元格生成前缀")?;
    write_examples(&mut workbook)?;
    workbook
        .save(path)
        .with_context(|| format!("save new configuration {}", path.display()))?;
    Ok(ConfigInitResult {
        created: true,
        path: path.to_path_buf(),
        sheets: SHEET_SPECS.len(),
        message: "new configuration template created; existing files are never overwritten".into(),
    })
}

pub fn view_config(path: impl AsRef<Path>) -> Result<ConfigView> {
    let path = path.as_ref();
    let dump =
        read_workbook(path).with_context(|| format!("read configuration {}", path.display()))?;
    let max_columns = dump
        .sheets
        .iter()
        .map(sheet_max_column)
        .max()
        .unwrap_or(1)
        .max(1);
    let date_values = read_value_columns_with_iso_dates(path, max_columns)?
        .into_iter()
        .flat_map(|sheet| {
            sheet
                .values
                .into_iter()
                .filter_map(move |(address, value)| {
                    let CellValue::DateTimeIso(value) = value else {
                        return None;
                    };
                    Some(((sheet.name.clone(), address), value))
                })
        })
        .collect::<BTreeMap<_, _>>();
    let sheets = dump
        .sheets
        .iter()
        .map(|sheet| sheet_view(sheet, &date_values))
        .collect();
    Ok(ConfigView {
        path: path.to_path_buf(),
        sheets,
    })
}

fn ensure_header_layout_unambiguous(
    sheet_name: &str,
    sheet: &SheetDump,
    headers: &[&str],
) -> Result<()> {
    let mut expected_columns = BTreeMap::<&str, Vec<usize>>::new();
    for (index, header) in headers.iter().enumerate() {
        expected_columns.entry(header).or_default().push(index + 1);
    }

    for address in sheet.formulas.keys() {
        if let Some((row, column)) = parse_a1(address) {
            if row == 1 && column <= headers.len() {
                bail!(
                    "configuration header layout is ambiguous: {sheet_name}!{address} contains a formula; refusing repair to protect existing configuration"
                );
            }
        }
    }

    for address in sheet.values.keys() {
        let Some((row, column)) = parse_a1(address) else {
            continue;
        };
        if row != 1 {
            continue;
        }
        let value = cell_text(sheet, address).trim().to_owned();
        let Some(expected) = expected_columns.get(value.as_str()) else {
            continue;
        };
        if !expected.contains(&column) {
            bail!(
                "configuration header layout is ambiguous: {sheet_name}!{address} contains {value:?}, expected column(s) {expected:?}; refusing repair to protect existing configuration"
            );
        }
    }
    Ok(())
}

/// Build a non-destructive repair plan using the legacy configuration schema:
/// correct configured headers/comments, add missing sheets/defaults, and fill
/// named examples only when their sheets have no data rows. Existing sheets,
/// including retired configuration sheets, are retained to protect data.
pub fn build_config_repair_plan(
    input: impl AsRef<Path>,
    output: impl AsRef<Path>,
) -> Result<Option<ExcelEditPlan>> {
    let input = input.as_ref();
    let output = output.as_ref();
    if !input.is_file() {
        bail!("configuration does not exist: {}", input.display());
    }
    if !input
        .extension()
        .and_then(|e| e.to_str())
        .is_some_and(|e| e.eq_ignore_ascii_case("xlsx"))
        || !output
            .extension()
            .and_then(|e| e.to_str())
            .is_some_and(|e| e.eq_ignore_ascii_case("xlsx"))
    {
        bail!("configuration repair plans accept .xlsx only");
    }
    if input.canonicalize().ok() == output.canonicalize().ok() && output.exists() {
        bail!("configuration repair cannot overwrite the original file");
    }
    let dump =
        read_workbook(input).with_context(|| format!("read configuration {}", input.display()))?;
    let profile = inspect_xlsx_profile(input)?;
    let extents = read_xlsx_worksheet_extents(input)?;
    let mut operations = Vec::new();
    for (sheet_name, headers) in SHEET_SPECS {
        if let Some(sheet) = dump.sheets.iter().find(|sheet| sheet.name == *sheet_name) {
            ensure_header_layout_unambiguous(sheet_name, sheet, headers)?;
            for (column, header) in headers.iter().enumerate() {
                let address = format!("{}1", column_letters(column + 1));
                if sheet.formulas.contains_key(&address) || cell_text(sheet, &address) != *header {
                    operations.push(EditOperation::SetCellValue {
                        sheet: (*sheet_name).to_owned(),
                        cell: address.clone(),
                        value: PlannedCellValue::Text((*header).to_owned()),
                    });
                }
                if let Some(note) = header_note(sheet_name, header) {
                    if sheet.comments.get(&address).map(String::as_str) != Some(note) {
                        operations.push(EditOperation::SetComment {
                            sheet: (*sheet_name).to_owned(),
                            cell: address,
                            text: Some(note.to_owned()),
                            author: Some("pytools".to_owned()),
                        });
                    }
                }
            }
        } else {
            operations.push(EditOperation::CreateSheet {
                name: (*sheet_name).to_owned(),
            });
            for (column, header) in headers.iter().enumerate() {
                let address = format!("{}1", column_letters(column + 1));
                operations.push(EditOperation::SetCellValue {
                    sheet: (*sheet_name).to_owned(),
                    cell: address.clone(),
                    value: PlannedCellValue::Text((*header).to_owned()),
                });
                if let Some(note) = header_note(sheet_name, header) {
                    operations.push(EditOperation::SetComment {
                        sheet: (*sheet_name).to_owned(),
                        cell: address,
                        text: Some(note.to_owned()),
                        author: Some("pytools".to_owned()),
                    });
                }
            }
        }
    }
    // V2 repair preserves existing workbook data. Although the legacy Python
    // repair removed these retired sheets, their rows may still contain user
    // configuration and must not be discarded by an automatic repair.
    if let Some(sheet) = dump.sheets.iter().find(|sheet| sheet.name == "全局配置") {
        let mut rows = std::collections::BTreeMap::new();
        let mut key_cells = sheet
            .values
            .iter()
            .filter_map(|(address, _)| {
                parse_a1(address)
                    .filter(|(_, column)| *column == 1)
                    .map(|(row, _)| (row, cell_text(sheet, address).trim().to_owned()))
            })
            .collect::<Vec<_>>();
        key_cells.sort_by_key(|(row, _)| *row);
        for (row, key) in key_cells {
            if !key.is_empty() {
                rows.insert(key, row);
            }
        }
        let mut next = sheet_max_row(sheet, &extents) + 1;
        for (key, value) in GLOBAL_DEFAULTS {
            if let Some(row) = rows.get(*key).copied() {
                if cell_is_blank(sheet, &format!("B{row}")) {
                    operations.push(EditOperation::SetCellValue {
                        sheet: "全局配置".into(),
                        cell: format!("B{row}"),
                        value: PlannedCellValue::Text((*value).to_owned()),
                    });
                }
            } else {
                operations.push(EditOperation::SetCellValue {
                    sheet: "全局配置".into(),
                    cell: format!("A{next}"),
                    value: PlannedCellValue::Text((*key).to_owned()),
                });
                operations.push(EditOperation::SetCellValue {
                    sheet: "全局配置".into(),
                    cell: format!("B{next}"),
                    value: PlannedCellValue::Text((*value).to_owned()),
                });
                next += 1;
            }
        }
    } else {
        for (row, (key, value)) in GLOBAL_DEFAULTS.iter().enumerate() {
            let target_row = row + 2;
            operations.push(EditOperation::SetCellValue {
                sheet: "全局配置".into(),
                cell: format!("A{target_row}"),
                value: PlannedCellValue::Text((*key).to_owned()),
            });
            operations.push(EditOperation::SetCellValue {
                sheet: "全局配置".into(),
                cell: format!("B{target_row}"),
                value: PlannedCellValue::Text((*value).to_owned()),
            });
        }
    }
    let examples: [(&str, Vec<Vec<&str>>); 3] = [
        (
            "归档类型配置",
            vec![
                vec!["是", "普惠贷款表", "普惠贷款表", "更具体的关键词放前面"],
                vec!["是", "科技贷款表", "科技贷款表", ""],
                vec!["是", "存款表", "存款表", ""],
                vec!["是", "贷款表", "贷款表", "短关键词放后面"],
            ],
        ),
        (
            "提交检查配置",
            vec![
                vec![
                    "Y",
                    "202506",
                    "惠城区",
                    "建设银行",
                    "存款表",
                    "",
                    "",
                    "示例：四个非空关键词都要出现在同一文件名",
                ],
                vec![
                    "Y",
                    "202506",
                    "",
                    "工商银行",
                    "贷款表",
                    "",
                    "",
                    "示例：文件名没有县区时，县区关键词可留空",
                ],
            ],
        ),
        (
            "文件排序配置",
            vec![
                vec!["Y", "001", "政策性银行;国开行;农发行", "政策性银行"],
                vec![
                    "Y",
                    "002",
                    "工商银行;农业银行;中国银行;建设银行;交通银行",
                    "国有银行",
                ],
                vec!["Y", "003", "招商银行;浦发银行;中信银行", "股份制银行"],
            ],
        ),
    ];
    for (sheet_name, rows) in examples {
        let empty = dump
            .sheets
            .iter()
            .find(|sheet| sheet.name == sheet_name)
            .map(|sheet| sheet_max_row(sheet, &extents) <= 1)
            .unwrap_or(true);
        if empty {
            for (row, values) in rows.iter().enumerate() {
                for (col, value) in values.iter().enumerate() {
                    operations.push(EditOperation::SetCellValue {
                        sheet: sheet_name.to_owned(),
                        cell: format!("{}{}", column_letters(col + 1), row + 2),
                        value: if value.is_empty() {
                            // openpyxl writes an empty string as a blank XLSX cell.
                            PlannedCellValue::Blank
                        } else {
                            PlannedCellValue::Text((*value).to_owned())
                        },
                    });
                }
            }
        }
    }
    let rename_sheet = dump.sheets.iter().find(|sheet| sheet.name == "重命名配置");
    for (address, value) in [
        ("O2", "Y"),
        ("P2", "sheet1@A1;sheet2@A2"),
        ("Q2", "示例：按文件内容单元格生成前缀"),
    ] {
        if rename_sheet
            .map(|sheet| cell_is_blank(sheet, address))
            .unwrap_or(true)
        {
            operations.push(EditOperation::SetCellValue {
                sheet: "重命名配置".into(),
                cell: address.into(),
                value: PlannedCellValue::Text(value.into()),
            });
        }
    }
    if let Some(rename_sheet) = rename_sheet {
        for address in ["L1", "L2"] {
            if rename_sheet.comments.contains_key(address) {
                operations.push(EditOperation::SetComment {
                    sheet: "重命名配置".into(),
                    cell: address.into(),
                    text: None,
                    author: None,
                });
            }
        }
    }
    if operations.is_empty() {
        return Ok(None);
    }
    let preserve = std::collections::BTreeSet::from([
        PreservationRequirement::CellValuesAndTypes,
        PreservationRequirement::FormulaText,
        PreservationRequirement::CommentsAndPositions,
        PreservationRequirement::Styles,
        PreservationRequirement::MergedRanges,
        PreservationRequirement::SheetIdentityAndOrder,
    ]);
    // Repair changes workbook structure. Require only semantic properties that
    // the selected validator can check; package-level and page-layout fidelity
    // remain capability-specific and must be added after dedicated validation.
    let plan = ExcelEditPlan {
        request_id: "pytools-config-repair".into(),
        input: input.to_path_buf(),
        output: output.to_path_buf(),
        format: WorkbookFormat::Xlsx,
        workbook: profile,
        operations,
        preserve,
        drop_allowed: Default::default(),
    };
    plan.validate()
        .context("validate configuration repair plan")?;
    Ok(Some(plan))
}

fn sheet_view(
    sheet: &SheetDump,
    date_values: &BTreeMap<(String, String), String>,
) -> ConfigSheetView {
    let max_row = sheet
        .values
        .keys()
        .filter_map(|a| parse_a1(a).map(|x| x.0))
        .chain(
            sheet
                .comments
                .keys()
                .filter_map(|a| parse_a1(a).map(|x| x.0)),
        )
        .max()
        .unwrap_or(0);
    let max_col = sheet
        .values
        .keys()
        .filter_map(|a| parse_a1(a).map(|x| x.1))
        .chain(
            sheet
                .comments
                .keys()
                .filter_map(|a| parse_a1(a).map(|x| x.1)),
        )
        .max()
        .unwrap_or(0);
    let mut rows = Vec::with_capacity(max_row);
    for row in 1..=max_row {
        let mut values = Vec::with_capacity(max_col);
        for col in 1..=max_col {
            let address = format!("{}{}", column_letters(col), row);
            values.push(
                match date_values.get(&(sheet.name.clone(), address.clone())) {
                    Some(value) => Value::String(normalize_iso_datetime(value)),
                    None => sheet
                        .values
                        .get(&address)
                        .map(value_as_json)
                        .unwrap_or(Value::Null),
                },
            );
        }
        rows.push(values);
    }
    ConfigSheetView {
        name: sheet.name.clone(),
        rows,
        comments: sheet.comments.clone(),
    }
}
fn sheet_max_column(sheet: &SheetDump) -> usize {
    sheet
        .values
        .keys()
        .filter_map(|address| parse_a1(address).map(|(_, col)| col))
        .chain(
            sheet
                .comments
                .keys()
                .filter_map(|address| parse_a1(address).map(|(_, col)| col)),
        )
        .max()
        .unwrap_or(0)
}
fn normalize_iso_datetime(value: &str) -> String {
    value.strip_suffix(".000").unwrap_or(value).to_owned()
}
fn value_as_json(value: &Value) -> Value {
    match value.get("value") {
        Some(value) => value.clone(),
        None => Value::Null,
    }
}
fn cell_text(sheet: &SheetDump, address: &str) -> String {
    sheet
        .values
        .get(address)
        .and_then(|value| value.get("value"))
        .map(|value| {
            value
                .as_str()
                .map(str::to_owned)
                .unwrap_or_else(|| value.to_string())
        })
        .unwrap_or_default()
}
fn cell_is_blank(sheet: &SheetDump, address: &str) -> bool {
    !sheet.formulas.contains_key(address) && cell_text(sheet, address).trim().is_empty()
}
fn extent_for<'a>(extents: &'a [WorksheetExtent], sheet_name: &str) -> Option<&'a WorksheetExtent> {
    extents.iter().find(|extent| extent.name == sheet_name)
}
fn sheet_max_row(sheet: &SheetDump, extents: &[WorksheetExtent]) -> usize {
    sheet
        .values
        .keys()
        .filter_map(|address| parse_a1(address).map(|(row, _)| row))
        .chain(
            sheet
                .comments
                .keys()
                .filter_map(|address| parse_a1(address).map(|(row, _)| row)),
        )
        .chain(extent_for(extents, &sheet.name).map(|extent| extent.max_row))
        .max()
        .unwrap_or(1)
}
fn parse_a1(address: &str) -> Option<(usize, usize)> {
    let index = address.find(|c: char| c.is_ascii_digit())?;
    let column = address[..index].chars().try_fold(0usize, |value, char| {
        value
            .checked_mul(26)?
            .checked_add(char.to_ascii_uppercase() as usize - 'A' as usize + 1)
    })?;
    Some((address[index..].parse().ok()?, column))
}
fn column_letters(mut column: usize) -> String {
    let mut result = String::new();
    while column > 0 {
        result.insert(0, (b'A' + ((column - 1) % 26) as u8) as char);
        column = (column - 1) / 26;
    }
    result
}

fn write_examples(workbook: &mut Workbook) -> Result<()> {
    let archive = workbook
        .worksheet_from_name("归档类型配置")
        .context("archive worksheet missing")?;
    for (row, values) in [
        ["是", "普惠贷款表", "普惠贷款表", "更具体的关键词放前面"],
        ["是", "科技贷款表", "科技贷款表", ""],
        ["是", "存款表", "存款表", ""],
        ["是", "贷款表", "贷款表", "短关键词放后面"],
    ]
    .iter()
    .enumerate()
    {
        for (col, value) in values.iter().enumerate() {
            archive.write_string((row + 1) as u32, col as u16, *value)?;
        }
    }
    let submission = workbook
        .worksheet_from_name("提交检查配置")
        .context("submission worksheet missing")?;
    let rows = [
        [
            "Y",
            "202506",
            "惠城区",
            "建设银行",
            "存款表",
            "",
            "",
            "示例：四个非空关键词都要出现在同一文件名",
        ],
        [
            "Y",
            "202506",
            "",
            "工商银行",
            "贷款表",
            "",
            "",
            "示例：文件名没有县区时，县区关键词可留空",
        ],
    ];
    for (row, values) in rows.iter().enumerate() {
        for (col, value) in values.iter().enumerate() {
            submission.write_string((row + 1) as u32, col as u16, *value)?;
        }
    }
    let sort = workbook
        .worksheet_from_name("文件排序配置")
        .context("file-sort worksheet missing")?;
    for (row, values) in [
        ["Y", "001", "政策性银行;国开行;农发行", "政策性银行"],
        [
            "Y",
            "002",
            "工商银行;农业银行;中国银行;建设银行;交通银行",
            "国有银行",
        ],
        ["Y", "003", "招商银行;浦发银行;中信银行", "股份制银行"],
    ]
    .iter()
    .enumerate()
    {
        for (col, value) in values.iter().enumerate() {
            sort.write_string((row + 1) as u32, col as u16, *value)?;
        }
    }
    Ok(())
}
const HEADER_NOTES: &[(&str, &str, &str)] = &[
        ("全局配置", "键", "关键字段（程序读取）：配置键名，例如 输入目录/输出目录/日志目录。"),
        ("全局配置", "日志模式", "普通、详细或关闭；交互菜单启动时读取，详细模式会显示开发者事件。"),
        ("全局配置", "日志保留天数", "正整数；仅清理受管理、已经完成的诊断 JSONL，留空则不按天清理。"),
        ("全局配置", "日志最大占用MB", "正整数；仅约束受管理、已经完成的诊断 JSONL，留空则不按大小清理。"),
        ("全局配置", "值", "关键字段（程序读取）：对应配置值。"),
        ("全局配置", "备注", "备注字段（仅人工说明）：程序不读取。"),
        ("时序提取规则", "是否启用", "关键字段（程序读取）：Y/是/1 表示启用。"),
        ("时序提取规则", "规则名称", "关键字段（程序读取）：规则唯一标识，建议不重复。"),
        ("时序提取规则", "工作簿关键字", "关键字段（程序读取）：支持分号分隔的关键字匹配。"),
        ("时序提取规则", "工作表关键字", "关键字段（程序读取）：支持分号分隔的关键字匹配。"),
        ("时序提取规则", "行头列", "关键字段（程序读取）：行头所在列（如 B 或 2）。"),
        ("时序提取规则", "列表头行", "关键字段（程序读取）：列表头行号，可多行（如 2,3）。"),
        ("时序提取规则", "必含列头", "关键字段（程序读取）：可空；用于前置命中校验。"),
        ("时序提取规则", "必含行头", "关键字段（程序读取）：可空；用于前置命中校验。"),
        ("时序提取规则", "数据起始行", "关键字段（程序读取）：可空；空则按表头后自动推断。"),
        ("时序提取规则", "数据结束行", "关键字段（程序读取）：可空。"),
        ("时序提取规则", "数据起始列", "关键字段（程序读取）：可空；空则按行头列后自动推断。"),
        ("时序提取规则", "数据结束列", "关键字段（程序读取）：可空。"),
        ("时序提取规则", "跳过关键字", "关键字段（程序读取）：可空；命中则跳过该行头。"),
        ("时序提取规则", "目标工作簿路径", "关键字段（程序读取）：可空；与目标工作表一起决定是否写入目标。"),
        ("时序提取规则", "目标工作表", "关键字段（程序读取）：可空；与目标工作簿路径一起生效。"),
        ("时序提取规则", "启用目标写入", "关键字段（程序读取）：Y/是/1 才写目标；空/否不写目标。"),
        ("时序提取规则", "set区域", "关键字段（程序读取）：可空；格式 别名@地址;别名@地址（如 机构代码@A3;报表口径@A5），仅宽表汇总使用。"),
        ("时序提取规则", "目标去重列", "关键字段（程序读取）：可空；宽表目标写入去重列，填写输出表列名或列标/列号（如 数据日期;行头路径 或 A;B;C），支持 分号/逗号 分隔。空则默认按 数据日期+set列+行头路径 去重。"),
        ("时序提取规则", "是否输出行序号", "关键字段（程序读取）：Y/是/1 时宽表汇总输出增加'行序号'列（源表行号），用于行头重复的报表区分同名列；建议配合 目标去重列=数据日期;行序号 使用。空/否不输出。"),
        ("路径标准化映射", "是否启用", "关键字段（程序读取）：Y/是/1 表示启用映射。"),
        ("路径标准化映射", "映射名称", "关键字段（程序读取）：映射规则名称。"),
        ("路径标准化映射", "适用规则名", "关键字段（程序读取）：可填具体规则名/全部/*。"),
        ("路径标准化映射", "工作簿关键字", "关键字段（程序读取）：可空；用于缩小映射生效范围。"),
        ("路径标准化映射", "工作表关键字", "关键字段（程序读取）：可空；用于缩小映射生效范围。"),
        ("路径标准化映射", "作用对象", "关键字段（程序读取）：行头/列头/两者。"),
        ("路径标准化映射", "匹配方式", "关键字段（程序读取）：精确/包含。"),
        ("路径标准化映射", "原始路径", "关键字段（程序读取）：待匹配文本。"),
        ("路径标准化映射", "标准路径", "关键字段（程序读取）：替换后的标准文本。"),
        ("路径标准化映射", "备注", "备注字段（仅人工说明）：程序不读取。"),
        ("去重追加数据配置", "是否启用", "关键字段（程序读取）：Y/是/1 表示启用。"),
        ("去重追加数据配置", "源数据工作簿", "关键字段（程序读取）：源工作簿完整路径。"),
        ("去重追加数据配置", "源数据工作表", "关键字段（程序读取）：源工作表名。"),
        ("去重追加数据配置", "标识列序号", "关键字段（程序读取）：去重键列，如 1;2;5；留空默认全列。"),
        ("去重追加数据配置", "目标工作簿", "关键字段（程序读取）：目标工作簿完整路径。"),
        ("去重追加数据配置", "目标工作表", "关键字段（程序读取）：目标工作表名。"),
        ("去重追加数据配置", "执行模式", "关键字段（程序读取）：1=正常执行；2=仅校验不写入；3=备份后执行。"),
        ("去重追加数据配置", "备注", "备注字段（仅人工说明）：程序不读取。"),
        ("打印配置", "是否启用", "关键字段（程序读取）：Y/是/1 表示启用。"),
        ("打印配置", "打印模式", "关键字段（程序读取）：1=保留源格式；2=保留源格式+自动扩列（减少####）；3=快速复制（值+基础格式，速度更快）。3.10.7 会按 1/2/3 全部执行。"),
        ("打印配置", "源工作簿", "关键字段（程序读取）：源工作簿完整路径。"),
        ("打印配置", "源工作表", "关键字段（程序读取）：源工作表名（精确匹配）。"),
        ("打印配置", "源工作表打印区域", "关键字段（程序读取）：可空；支持 A1:H30;A35:H70，多段分号分隔。"),
        ("打印配置", "目标工作簿", "关键字段（程序读取）：目标工作簿完整路径；不存在会自动创建。"),
        ("打印配置", "目标工作表", "关键字段（程序读取）：目标工作表名；不存在会自动创建。"),
        ("打印配置", "FitToPagesWide", "关键字段（程序读取）：分页宽度页数，留空默认 1。"),
        ("打印配置", "FitToPagesTall", "关键字段（程序读取）：分页高度页数，留空默认 1。"),
        ("打印配置", "打印方向", "关键字段（程序读取）：横向/纵向；留空按宽高比自动。"),
        ("打印配置", "水平居中", "关键字段（程序读取）：Y/是/1 时打印页面水平居中。"),
        ("打印配置", "垂直居中", "关键字段（程序读取）：Y/是/1 时打印页面垂直居中。"),
        ("打印配置", "不输出批注", "关键字段（程序读取）：Y/是/1 时清除输出文件中的单元格批注。"),
        ("打印配置", "零值不输出", "关键字段（程序读取）：Y/是/1 时将 0 值输出为空白。"),
        ("打印配置", "备注", "备注字段（仅人工说明）：程序不读取。"),
        ("重命名配置", "简称", "关键字段（程序读取）：3.3 使用，简称->全称映射（A->B）。"),
        ("重命名配置", "全称", "关键字段（程序读取）：3.3 使用。"),
        ("重命名配置", "代码", "关键字段（程序读取）：3.3 使用，与 E 列全称做全称->代码映射（E->D）。"),
        ("重命名配置", "全称(代码映射)", "关键字段（程序读取）：3.3 使用。"),
        ("重命名配置", "键", "关键字段（程序读取）：3.3 使用（如 数据日期、报表名称）。"),
        ("重命名配置", "值", "关键字段（程序读取）：3.3 使用。"),
        ("重命名配置", "原表名", "关键字段（程序读取）：3.7 使用，原 Sheet 名。"),
        ("重命名配置", "新表名", "关键字段（程序读取）：3.7 使用，新 Sheet 名。"),
        ("重命名配置", "内容重命名启用", "关键字段（程序读取）：4.5/4.6 使用，Y/是/1/true 表示启用。只能启用一条。"),
        ("重命名配置", "内容重命名取值规则", "关键字段（程序读取）：4.5/4.6 使用，填写 sheet@单元格 规则，多个用分号分隔。"),
        ("重命名配置", "原文件名片段", "关键字段（程序读取）：4.8 使用，文件名中要替换的片段（M列）。"),
        ("重命名配置", "新文件名片段", "关键字段（程序读取）：4.8 使用，替换后的文件名片段（N列）。"),
        ("重命名配置", "备注", "备注字段（仅人工说明）：程序不读取。建议写适用文件或报表类型。"),
        ("归档类型配置", "是否启用", "关键字段（程序读取）：Y/是/1 表示启用该类型规则。"),
        ("归档类型配置", "匹配关键词", "关键字段（程序读取）：按类型归档时，在文件名中查找的关键词。"),
        ("归档类型配置", "归档文件夹", "关键字段（程序读取）：命中关键词后放入的单层文件夹名称。"),
        ("归档类型配置", "备注", "备注字段（仅人工说明）：程序不读取。建议把更具体的关键词放在前面。"),
        ("归档类型配置", "需要归档的后缀", "关键字段（程序读取）：E2 填写允许归档的后缀，支持 xlsx;.xls;docx 多个用分号/逗号/空格分隔；空表示不限制。"),
        ("归档类型配置", "排除归档的后缀", "关键字段（程序读取）：F2 填写要排除的后缀，支持 exe;dll;tmp 多个用分号/逗号/空格分隔；优先级高于 E2。"),
        ("提交检查配置", "是否启用", "关键字段（程序读取）：Y/是/1/true 表示启用该检查项。"),
        ("提交检查配置", "日期关键词", "关键字段（程序读取）：可空；非空时必须出现在文件名中。"),
        ("提交检查配置", "县区关键词", "关键字段（程序读取）：可空；非空时必须出现在文件名中。"),
        ("提交检查配置", "机构关键词", "关键字段（程序读取）：可空；非空时必须出现在文件名中。"),
        ("提交检查配置", "备用关键词", "关键字段（程序读取）：可空；非空时必须出现在文件名中，可填报表类型等。"),
        ("提交检查配置", "命中次数", "输出字段（程序写入）：匹配到的文件数量。"),
        ("提交检查配置", "命中的文件名", "输出字段（程序写入）：匹配到的文件名，多个用分号分隔。"),
        ("提交检查配置", "备注", "备注字段（仅人工说明）：程序不读取。"),
        ("文件排序配置", "是否启用", "关键字段（程序读取）：Y/是/1/true 表示启用该排序规则。"),
        ("文件排序配置", "排序前缀", "关键字段（程序读取）：支持 1/01/001，程序统一补齐为三位数字。"),
        ("文件排序配置", "匹配关键词", "关键字段（程序读取）：文件名命中任一关键词即使用该前缀，多个用分号/逗号分隔。"),
        ("文件排序配置", "备注", "备注字段（仅人工说明）：程序不读取。"),
        ("机构映射表", "原始机构名称", "关键字段（程序读取）：1.2 使用，源文件中待标准化的机构名。"),
        ("机构映射表", "映射后机构名称", "关键字段（程序读取）：1.2 使用，标准化后的机构名。"),
        ("机构映射表", "是否为外资行", "关键字段（程序读取）：1.3 使用，1=外资行（将被剔除），0/空=否。"),
        ("工作表提取", "是否启用", "关键字段（程序读取）：Y/是/1/true 启用该行。"),
        ("工作表提取", "工作簿关键字", "关键字段（程序读取）：匹配源文件名；分号/中文分号=必须同时命中，逗号/中文逗号/|=同组任选；空=不限制。"),
        ("工作表提取", "工作表关键字", "关键字段（程序读取）：匹配工作表名；分号/中文分号=必须同时命中，逗号/中文逗号/|=同组任选；空=不限制。"),
        ("工作表提取", "排除关键字", "关键字段（程序读取）：匹配源文件名或工作表名；命中则跳过；语法同关键词列；空=不排除。"),
        ("工作表提取", "提取区域", "关键字段（程序读取）：空/整表/全部=整表；支持 A1:D20、A:D、1:10；多个区域用分号分隔。"),
        ("工作表提取", "多区域拼接方式", "关键字段（程序读取）：空/横向拼接=左右拼接；纵向追加=上下追加。"),
        ("工作表提取", "输出文件夹", "关键字段（程序读取）：空=默认 output；相对路径=output 下子目录；绝对路径=直接使用。"),
        ("工作表提取", "文件名", "关键字段（程序读取）：输出工作簿基名（不含扩展名）；空则自动命名。"),
        ("工作表提取", "目标工作表", "关键字段（程序读取）：输出表名；空则沿用源工作表名。结果写入新输出表，不覆盖已有工作表；重名时自动添加序号。"),
        ("工作表提取", "备注", "备注字段（仅人工说明）：程序不读取。"),
];

fn header_note(sheet: &str, header: &str) -> Option<&'static str> {
    HEADER_NOTES
        .iter()
        .find_map(|(candidate_sheet, candidate_header, text)| {
            (*candidate_sheet == sheet && *candidate_header == header).then_some(*text)
        })
}
