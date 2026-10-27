use anyhow::Result;
use pytools_core::batch_ops::{build_quarterly_rural_loan_plan, build_rural_loan_plan};
use rust_xlsxwriter::{Formula, Workbook};
use std::{fs, path::PathBuf};

struct TestDir(PathBuf);

impl TestDir {
    fn new() -> Self {
        let path = std::env::temp_dir().join(format!(
            "pytools-rural-loan-regressions-{}-{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
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

#[derive(Clone)]
enum Cell {
    Text(&'static str),
    Number(f64),
}

fn write_book(path: &std::path::Path, sheets: Vec<(&str, Vec<Vec<Cell>>)>) -> Result<()> {
    let mut workbook = Workbook::new();
    for (name, rows) in sheets {
        let worksheet = workbook.add_worksheet();
        worksheet.set_name(name)?;
        for (row, values) in rows.iter().enumerate() {
            for (column, value) in values.iter().enumerate() {
                match value {
                    Cell::Text(text) => {
                        worksheet.write_string(row as u32, column as u16, *text)?;
                    }
                    Cell::Number(number) => {
                        worksheet.write_number(row as u32, column as u16, *number)?;
                    }
                }
            }
        }
    }
    workbook.save(path)?;
    Ok(())
}

fn institution_sheet(entries: &[(&'static str, f64, f64)]) -> Vec<Vec<Cell>> {
    let mut rows = vec![
        vec![
            Cell::Text("序号"),
            Cell::Text("机构名称"),
            Cell::Text("涉农贷款余额"),
            Cell::Text("旧值"),
        ],
        vec![
            Cell::Text("序号"),
            Cell::Text("机构名称"),
            Cell::Text("涉农贷款余额"),
            Cell::Text("旧值"),
        ],
    ];
    rows.extend(
        entries
            .iter()
            .enumerate()
            .map(|(index, (name, value, old))| {
                vec![
                    Cell::Number((index + 1) as f64),
                    Cell::Text(name),
                    Cell::Number(*value),
                    Cell::Number(*old),
                ]
            }),
    );
    rows
}

#[test]
fn rural_loan_plan_preserves_unrounded_difference() -> Result<()> {
    let dir = TestDir::new();
    let current = dir.path("current.xlsx");
    let previous = dir.path("previous.xlsx");
    write_book(
        &current,
        vec![
            (
                "惠州市涉农贷款分机构",
                institution_sheet(&[("甲行", 14.123456789, 999.0)]),
            ),
            (
                "本外币涉农贷款分地区",
                vec![
                    vec![Cell::Text("地区"), Cell::Text("涉农贷款比上月")],
                    vec![Cell::Text("惠州市"), Cell::Number(5.123456789)],
                ],
            ),
        ],
    )?;
    write_book(
        &previous,
        vec![(
            "惠州市涉农贷款分机构",
            institution_sheet(&[("甲行", 9.0, 0.0)]),
        )],
    )?;

    let (report, summary) =
        build_rural_loan_plan(&current, &previous, &dir.path("candidate.xlsx"))?;
    let plan = report.plan.expect("rural loan plan");
    let value = plan
        .operations
        .iter()
        .find_map(|operation| match operation {
            excel_editor::EditOperation::SetCellValue {
                cell,
                value: excel_editor::CellValue::Decimal(value),
                ..
            } if cell == "D3" => Some(value),
            _ => None,
        })
        .expect("D3 difference operation");
    let actual = value.parse::<f64>().expect("decimal value");
    assert!((actual - 5.123456789).abs() <= 1e-15);
    assert!((summary.fill_total - actual).abs() <= 1e-15);
    Ok(())
}

#[test]
fn rural_loan_plans_match_legacy_formula_read_modes() -> Result<()> {
    let dir = TestDir::new();
    let current = dir.path("current.xlsx");
    let previous = dir.path("previous.xlsx");
    write_formula_rural_book(&current, "99", "80")?;
    write_formula_rural_book(&previous, "9", "70")?;
    let current_dump = excel_reader::read_workbook(&current)?;
    let previous_dump = excel_reader::read_workbook(&previous)?;
    assert_eq!(current_dump.sheets[0].values["C3"]["value"], 99.0);
    assert!(current_dump.sheets[0].formulas["C3"].contains("1+1"));
    assert_eq!(previous_dump.sheets[0].values["C3"]["value"], 9.0);
    let current_before = fs::read(&current)?;
    let previous_before = fs::read(&previous)?;

    let (city_report, city_summary) =
        build_rural_loan_plan(&current, &previous, &dir.path("city-candidate.xlsx"))?;
    let city_plan = city_report.plan.expect("city rural-loan plan");
    assert_eq!(planned_decimal(&city_plan, "D3"), Some("-9"));
    assert_eq!(city_summary.fill_total, -9.0);
    assert_eq!(city_summary.area_total, 0.0);
    assert!(!city_summary.consistent);

    let (quarter_report, quarter_summary) =
        build_quarterly_rural_loan_plan(&current, &previous, &dir.path("quarter-candidate.xlsx"))?;
    let quarter_plan = quarter_report.plan.expect("quarter rural-loan plan");
    assert_eq!(planned_decimal(&quarter_plan, "D3"), Some("-9"));
    assert_eq!(quarter_summary.regions[0].fill_total, -9.0);
    assert_eq!(quarter_summary.regions[0].area_total, Some(0.0));
    assert!(!quarter_summary.regions[0].consistent);
    assert_eq!(fs::read(&current)?, current_before);
    assert_eq!(fs::read(&previous)?, previous_before);
    Ok(())
}

fn write_formula_rural_book(
    path: &std::path::Path,
    institution_cache: &str,
    area_cache: &str,
) -> Result<()> {
    let mut workbook = Workbook::new();
    let institution = workbook.add_worksheet();
    institution.set_name("惠州市涉农贷款分机构")?;
    for row in 0..2 {
        institution.write_string(row, 0, "序号")?;
        institution.write_string(row, 1, "机构名称")?;
        institution.write_string(row, 2, "涉农贷款余额")?;
        institution.write_string(row, 3, "旧值")?;
    }
    institution.write_number(2, 0, 1.0)?;
    institution.write_string(2, 1, "甲行")?;
    institution.write_formula(2, 2, Formula::new("1+1").set_result(institution_cache))?;
    institution.write_number(2, 3, 999.0)?;

    let area = workbook.add_worksheet();
    area.set_name("本外币涉农贷款分地区统计")?;
    area.write_string(0, 0, "地区")?;
    area.write_string(0, 1, "涉农贷款比上月")?;
    area.write_string(1, 0, "惠州市")?;
    area.write_formula(1, 1, Formula::new("40+40").set_result(area_cache))?;
    workbook.save(path)?;
    Ok(())
}

fn planned_decimal<'a>(plan: &'a excel_editor::ExcelEditPlan, cell: &str) -> Option<&'a str> {
    plan.operations
        .iter()
        .find_map(|operation| match operation {
            excel_editor::EditOperation::SetCellValue {
                cell: address,
                value: excel_editor::CellValue::Decimal(value),
                ..
            } if address == cell => Some(value.as_str()),
            _ => None,
        })
}

#[test]
fn rural_loan_reports_and_plan_follow_institution_row_order() -> Result<()> {
    let dir = TestDir::new();
    let current = dir.path("current.xlsx");
    let previous = dir.path("previous.xlsx");
    write_book(
        &current,
        vec![
            (
                "惠州市涉农贷款分机构",
                institution_sheet(&[
                    ("机构乙", 15.0, 999.0),
                    ("未匹配甲", 4.0, 998.0),
                    ("机构甲", 11.0, 997.0),
                    ("未匹配乙", 5.0, 996.0),
                ]),
            ),
            (
                "本外币涉农贷款分地区",
                vec![
                    vec![Cell::Text("地区"), Cell::Text("涉农贷款比上月")],
                    vec![Cell::Text("惠州市"), Cell::Number(15.0)],
                ],
            ),
        ],
    )?;
    write_book(
        &previous,
        vec![(
            "惠州市涉农贷款分机构",
            institution_sheet(&[
                ("机构乙", 10.0, 0.0),
                ("上期乙", 3.0, 0.0),
                ("机构甲", 1.0, 0.0),
                ("上期甲", 5.0, 0.0),
            ]),
        )],
    )?;

    let (report, summary) =
        build_rural_loan_plan(&current, &previous, &dir.path("candidate.xlsx"))?;
    assert_eq!(summary.unmatched_current, ["未匹配甲", "未匹配乙"]);
    assert_eq!(summary.previous_only, ["上期乙", "上期甲"]);
    let plan = report.plan.expect("rural-loan plan missing");
    let edited_cells = plan
        .operations
        .iter()
        .filter_map(|operation| match operation {
            excel_editor::EditOperation::SetCellValue { cell, .. } => Some(cell.as_str()),
            _ => None,
        })
        .collect::<Vec<_>>();
    assert_eq!(edited_cells, ["D3", "D5"]);
    Ok(())
}

#[test]
fn quarterly_summary_keeps_invalid_regions_and_fails_overall_check() -> Result<()> {
    let dir = TestDir::new();
    let current = dir.path("current.xlsx");
    let previous = dir.path("previous.xlsx");
    write_book(
        &current,
        vec![
            (
                "惠州市涉农贷款分机构",
                institution_sheet(&[("甲行", 14.123456789, 999.0)]),
            ),
            (
                "江门市涉农贷款分机构",
                institution_sheet(&[("江门行", 8.0, 777.0)]),
            ),
            (
                "珠海市涉农贷款分机构",
                institution_sheet(&[("重复行", 4.0, 777.0), ("重复行", 5.0, 888.0)]),
            ),
            (
                "深圳市涉农贷款分机构",
                institution_sheet(&[("深行", 5.0, 999.0)]),
            ),
            (
                "涉农贷款分地区统计表",
                vec![
                    vec![Cell::Text("地区"), Cell::Text("涉农贷款比上月")],
                    vec![Cell::Text("惠州市"), Cell::Number(5.123456789)],
                    vec![Cell::Text("江门市"), Cell::Number(8.0)],
                    vec![Cell::Text("珠海市"), Cell::Number(0.0)],
                ],
            ),
        ],
    )?;
    write_book(
        &previous,
        vec![
            (
                "惠州市涉农贷款分机构",
                institution_sheet(&[("甲行", 9.0, 888.0)]),
            ),
            (
                "珠海市涉农贷款分机构",
                institution_sheet(&[("重复行", 3.0, 999.0)]),
            ),
            (
                "深圳市涉农贷款分机构",
                institution_sheet(&[("深行", 3.0, 999.0)]),
            ),
        ],
    )?;

    let (report, summary) =
        build_quarterly_rural_loan_plan(&current, &previous, &dir.path("candidate.xlsx"))?;
    assert_eq!(summary.regions.len(), 4);
    assert!(!summary.consistent);
    assert_eq!(summary.error, None);
    assert_eq!(summary.regions[0].error, None);
    assert_eq!(
        summary.regions[1].error.as_deref(),
        Some("上期文件缺少江门市分机构表")
    );
    assert_eq!(
        summary.regions[2].error.as_deref(),
        Some("分机构表存在重复机构名称：重复行")
    );
    assert_eq!(
        summary.regions[3].error.as_deref(),
        Some("“涉农贷款分地区统计表”未找到地区行：深圳市")
    );
    assert_eq!(report.changed_cells, 2);
    assert_eq!(report.plan.as_ref().unwrap().operations.len(), 2);
    Ok(())
}

#[test]
fn quarterly_rural_loan_reports_follow_institution_row_order() -> Result<()> {
    let dir = TestDir::new();
    let current = dir.path("current.xlsx");
    let previous = dir.path("previous.xlsx");
    write_book(
        &current,
        vec![
            (
                "江门市涉农贷款分机构",
                institution_sheet(&[
                    ("机构乙", 15.0, 999.0),
                    ("未匹配甲", 4.0, 998.0),
                    ("机构甲", 11.0, 997.0),
                    ("未匹配乙", 5.0, 996.0),
                ]),
            ),
            (
                "涉农贷款分地区统计表",
                vec![
                    vec![Cell::Text("地区"), Cell::Text("涉农贷款比上月")],
                    vec![Cell::Text("江门市"), Cell::Number(15.0)],
                ],
            ),
        ],
    )?;
    write_book(
        &previous,
        vec![(
            "江门市涉农贷款分机构",
            institution_sheet(&[
                ("机构乙", 10.0, 0.0),
                ("上期乙", 3.0, 0.0),
                ("机构甲", 1.0, 0.0),
                ("上期甲", 5.0, 0.0),
            ]),
        )],
    )?;

    let (report, summary) =
        build_quarterly_rural_loan_plan(&current, &previous, &dir.path("candidate.xlsx"))?;
    let region = &summary.regions[0];
    assert_eq!(region.unmatched_current, ["未匹配甲", "未匹配乙"]);
    assert_eq!(region.previous_only, ["上期乙", "上期甲"]);
    assert_eq!(report.changed_cells, 2);
    let plan = report.plan.expect("quarterly rural-loan plan missing");
    let edited_cells = plan
        .operations
        .iter()
        .filter_map(|operation| match operation {
            excel_editor::EditOperation::SetCellValue { cell, .. } => Some(cell.as_str()),
            _ => None,
        })
        .collect::<Vec<_>>();
    assert_eq!(edited_cells, ["D3", "D5"]);
    Ok(())
}
