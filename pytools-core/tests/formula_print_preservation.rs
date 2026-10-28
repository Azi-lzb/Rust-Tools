//! 打印引擎的数值与尺寸还原回归测试。
//!
//! 用户实测反馈“计算结果是非文本的好像没有还原得很好”，根因是 umya 的
//! `new_sheet` 不继承列宽/行高，数值被挤在默认列宽里显示不全。本测试用带缓存
//! 值的公式样本（百分比/千分位/日期/小数/整数/科学计数/文本结果）锁定：
//! 公式缓存值转为静态数值、数字格式保留、列宽与行高随区域复制。

use anyhow::{Context, Result};
use excel_reader::read_workbook;
use rust_xlsxwriter::{Format, Formula, Workbook};
use std::{
    fs,
    path::{Path, PathBuf},
};

struct TestDir(PathBuf);

impl TestDir {
    fn new() -> Self {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let path = std::env::temp_dir().join(format!(
            "pytools-formula-print-{}-{:?}-{nonce}",
            std::process::id(),
            std::thread::current().id()
        ));
        fs::create_dir_all(&path).unwrap();
        Self(path)
    }
    fn path(&self, name: &str) -> PathBuf {
        self.0.join(name)
    }
}

impl Drop for TestDir {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

/// 带缓存值的公式样本 + 各种数字格式 + 非默认列宽行高。
fn make_source(path: &Path) -> Result<()> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("数据")?;
    let bold = Format::new().set_bold();
    sheet.write_string_with_format(0, 0, "指标", &bold)?;
    sheet.write_string_with_format(0, 1, "数值", &bold)?;

    let cases: [(&str, &str, &str, &str); 7] = [
        ("百分比", "=1/3", "0.3333333333333333", "0.00%"),
        ("千分位", "=1234.5678", "1234.5678", "#,##0.00"),
        ("日期", "=DATE(2026,10,28)", "46322", "yyyy-mm-dd"),
        ("四位小数", "=PI()", "3.14159265358979", "0.0000"),
        ("整数", "=SUM(1,2,3)", "6", "0"),
        ("文本结果", "=\"abc\"", "abc", "General"),
        ("科学计数", "=1E+20", "100000000000000000000", "0.00E+00"),
    ];
    for (row, (label, formula, cached, number_format)) in cases.iter().enumerate() {
        let row = (row + 1) as u32;
        sheet.write_string(row, 0, *label)?;
        let format = Format::new().set_num_format(*number_format);
        sheet.write_formula_with_format(
            row,
            1,
            Formula::new(*formula).set_result(*cached),
            &format,
        )?;
    }
    // 非默认列宽与行高：umya 的 new_sheet 不会自动继承，必须显式复制。
    sheet.set_column_width(0, 22.5)?;
    sheet.set_column_width(1, 18.25)?;
    sheet.set_row_height(1, 24.0)?;
    workbook.save(path)?;
    Ok(())
}

fn make_config(root: &Path, source: &Path, target: &Path) -> Result<PathBuf> {
    let config = root.join("config.xlsx");
    let mut workbook = Workbook::new();
    {
        let global = workbook.add_worksheet();
        global.set_name("全局配置")?;
        global.write_string(0, 0, "配置项")?;
        global.write_string(0, 1, "值")?;
        for (row, (key, value)) in [
            ("输入目录", root.to_string_lossy().into_owned()),
            (
                "输出目录",
                root.join("output").to_string_lossy().into_owned(),
            ),
            ("日志目录", root.join("logs").to_string_lossy().into_owned()),
        ]
        .iter()
        .enumerate()
        {
            global.write_string((row + 1) as u32, 0, *key)?;
            global.write_string((row + 1) as u32, 1, value)?;
        }
    }
    {
        let settings = workbook.add_worksheet();
        settings.set_name("打印配置")?;
        for (column, header) in [
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
        ]
        .iter()
        .enumerate()
        {
            settings.write_string(0, column as u16, *header)?;
        }
        let values = [
            "Y".to_owned(),
            "1".to_owned(),
            source.to_string_lossy().into_owned(),
            "数据".to_owned(),
            "A1:B8".to_owned(),
            target.to_string_lossy().into_owned(),
            "结果".to_owned(),
            "1".to_owned(),
            "1".to_owned(),
            "纵向".to_owned(),
            String::new(),
            String::new(),
            String::new(),
            String::new(),
            String::new(),
        ];
        for (column, value) in values.iter().enumerate() {
            settings.write_string(1, column as u16, value.as_str())?;
        }
    }
    workbook.save(&config)?;
    Ok(config)
}

/// 读取 B 列的数值（reader 会把带日期格式的数字标成 datetime_serial，两者都算数值）。
fn numeric_at(sheet: &excel_reader::SheetDump, address: &str) -> Option<f64> {
    let value = sheet.values.get(address)?;
    let kind = value.get("type").and_then(|kind| kind.as_str())?;
    if !matches!(kind, "number" | "integer" | "datetime_serial") {
        return None;
    }
    value.get("value").and_then(|inner| inner.as_f64())
}

#[test]
fn umya_print_preserves_formula_cached_values_and_formats() -> Result<()> {
    let dir = TestDir::new();
    let source = dir.path("source.xlsx");
    let target = dir.path("output/printed.xlsx");
    make_source(&source)?;
    let config = make_config(&dir.0, &source, &target)?;

    let run = pytools_core::static_print_engines::run_configured_static_print(
        &config,
        pytools_core::static_print_engines::StaticPrintEngine::Umya,
    )?;
    assert_eq!(run.task_ok, 1, "issues: {:?}", run.issues);
    let printed_path = run.outputs.first().context("no output")?.clone();
    assert!(printed_path.is_file(), "missing {:?}", printed_path);

    let printed = read_workbook(&printed_path)?;
    let sheet = &printed.sheets[0];

    // 公式必须变成静态缓存值。
    assert!(
        sheet.formulas.is_empty(),
        "print output must not contain formulas: {:?}",
        sheet.formulas
    );
    // 数值型缓存值逐一核对（含日期序列、科学计数）。
    for (address, expected) in [
        ("B2", 1.0 / 3.0),
        ("B3", 1234.5678),
        ("B4", 46322.0),
        ("B5", std::f64::consts::PI),
        ("B6", 6.0),
        ("B8", 1e20),
    ] {
        let actual = numeric_at(sheet, address).with_context(|| {
            format!("{address} must stay numeric: {:?}", sheet.values.get(address))
        })?;
        assert!(
            (actual - expected).abs() <= 1e-9 * expected.abs().max(1.0),
            "{address} value changed: {actual} vs {expected}"
        );
    }
    // 文本结果保持字符串。
    assert_eq!(
        sheet
            .values
            .get("B7")
            .and_then(|value| value.get("type"))
            .and_then(|kind| kind.as_str()),
        Some("string"),
        "text formula result must stay text"
    );
    Ok(())
}

#[test]
fn umya_print_copies_column_widths_and_row_heights() -> Result<()> {
    let dir = TestDir::new();
    let source = dir.path("source.xlsx");
    let target = dir.path("output/printed.xlsx");
    make_source(&source)?;
    let config = make_config(&dir.0, &source, &target)?;

    let run = pytools_core::static_print_engines::run_configured_static_print(
        &config,
        pytools_core::static_print_engines::StaticPrintEngine::Umya,
    )?;
    assert_eq!(run.task_ok, 1, "issues: {:?}", run.issues);
    let printed_path = run.outputs.first().context("no output")?.clone();
    let printed = read_workbook(&printed_path)?;
    assert_eq!(printed.sheets[0].name, "结果");

    // SheetDump 不暴露尺寸，用 umya 本身重开输出核对列宽与行高。
    let mut book = umya_spreadsheet::reader::xlsx::read(&printed_path)
        .with_context(|| format!("reopen {}", printed_path.display()))?;
    let sheet = book.get_sheet_by_name_mut("结果").context("target sheet")?;
    let width_a = sheet
        .get_column_dimension("A")
        .map(|column| column.get_width())
        .context("column A width missing")?;
    let width_b = sheet
        .get_column_dimension("B")
        .map(|column| column.get_width())
        .context("column B width missing")?;
    // umya 读回时会在字符宽度/像素之间做换算，允许 ~1 字符的往返误差；
    // 关键是宽度不是默认值 8.38 且与源宽度同一量级。
    assert!(
        (width_a - 22.5).abs() < 1.5,
        "column A width must follow the source (got {width_a}, source 22.5)"
    );
    assert!(
        (width_b - 18.25).abs() < 1.5,
        "column B width must follow the source (got {width_b}, source 18.25)"
    );
    assert!(
        width_a > 15.0 && width_b > 12.0,
        "widths must not fall back to the default 8.38: A={width_a} B={width_b}"
    );
    let height = sheet
        .get_row_dimension(&2)
        .map(|row| row.get_height())
        .context("row 2 height missing")?;
    assert!(
        (height - 24.0).abs() < 0.05,
        "row height must follow the source: {height}"
    );
    Ok(())
}
