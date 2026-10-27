//! Regression coverage for row-operation coordinate mapping in the
//! engine-neutral validator, registered by the Audit S1 flow: the complete
//! history plan (delete old rows, re-insert, write new content) was falsely
//! rejected because style-only cells inside the deleted range and rows
//! shifted by insertions were compared at their source coordinates.

use anyhow::Result;
use excel_editor::{
    inspect_xlsx_profile, BackendCapability, CapabilityStatus, EditBackend, EditOperation,
    EditOperationKind, EngineId, ExcelEditPlan, OpenpyxlXlsxValidator, OutputValidator,
    PreservationRequirement, UmyaXlsxBackend, UmyaXlsxValidator, WorkbookFormat,
};
use std::{
    fs,
    path::{Path, PathBuf},
    time::{SystemTime, UNIX_EPOCH},
};

struct TestDir(PathBuf);

impl TestDir {
    fn new() -> Self {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("system clock after epoch")
            .as_nanos();
        let path = std::env::temp_dir().join(format!("excel-editor-row-map-{nonce}"));
        fs::create_dir_all(&path).expect("create test directory");
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

fn fill_style(style: &mut umya_spreadsheet::Style, argb: &str) {
    let mut color = umya_spreadsheet::Color::default();
    color.set_argb_str(argb);
    style
        .get_fill_mut()
        .get_pattern_fill_mut()
        .set_pattern_type(umya_spreadsheet::PatternValues::Solid)
        .set_foreground_color(color);
}

/// History sheet shaped like the real S1 workbook: values in rows 2..3,
/// style-only cells (L3, M6) inside the rows the plan deletes, and content
/// below the deleted range (A7/B7/A8) that must survive in place.
fn write_history_input(path: &Path) -> Result<()> {
    let mut workbook = umya_spreadsheet::new_file();
    let sheet = workbook.sheet_by_name_mut("Sheet1")?;
    sheet.set_name("History");
    sheet.cell_mut("A1").set_value_string("标题");
    sheet.cell_mut("A2").set_value_string("旧数据1");
    sheet.cell_mut("A3").set_value_string("旧数据2");
    fill_style(sheet.get_style_mut("L3"), "FFFFC7CE");
    fill_style(sheet.get_style_mut("M6"), "FFC6EFCE");
    sheet.cell_mut("A7").set_value_string("below");
    fill_style(sheet.get_style_mut("B7"), "FFDDEBF7");
    sheet.cell_mut("A8").set_value_string("below2");
    let mut other = workbook.new_sheet("Other")?;
    other.cell_mut("A2").set_value_string("other stays");
    umya_spreadsheet::writer::xlsx::write(&workbook, path)?;
    Ok(())
}

#[test]
fn delete_rewrite_plan_validates_without_deleted_range_false_positives() -> Result<()> {
    let dir = TestDir::new();
    let input = dir.path("input.xlsx");
    let output = dir.path("output.xlsx");
    write_history_input(&input)?;

    let profile = inspect_xlsx_profile(&input)?;
    let plan = ExcelEditPlan {
        request_id: "row-map-delete-rewrite".to_owned(),
        input: input.clone(),
        output: output.clone(),
        format: WorkbookFormat::Xlsx,
        workbook: profile.clone(),
        operations: vec![
            EditOperation::DeleteRows {
                sheet: "History".to_owned(),
                first_row: 2,
                count: 5,
            },
            EditOperation::InsertRows {
                sheet: "History".to_owned(),
                first_row: 2,
                count: 5,
            },
            EditOperation::SetCellValue {
                sheet: "History".to_owned(),
                cell: "A2".to_owned(),
                value: excel_editor::CellValue::Text("新内容".to_owned()),
            },
            EditOperation::SetCellValue {
                sheet: "History".to_owned(),
                cell: "A3".to_owned(),
                value: excel_editor::CellValue::Text("第二条".to_owned()),
            },
        ],
        preserve: [
            PreservationRequirement::CellValuesAndTypes,
            PreservationRequirement::FormulaText,
            PreservationRequirement::CommentsAndPositions,
            PreservationRequirement::Styles,
        ]
        .into_iter()
        .collect(),
        drop_allowed: Default::default(),
    };
    let backend = UmyaXlsxBackend::new(BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Unverified,
        format: WorkbookFormat::Xlsx,
        operations: [
            EditOperationKind::DeleteRows,
            EditOperationKind::InsertRows,
            EditOperationKind::SetCellValue,
        ]
        .into_iter()
        .collect(),
        supported_features: profile.features_present,
        preserved: plan.preserve.clone(),
        platform: "synthetic-row-map-test".to_owned(),
        evidence: None,
    })?;
    backend
        .apply(&plan)
        .map_err(|failure| anyhow::anyhow!(failure.message))?;

    // The engine-neutral validator previously false-positived on the
    // style-only cells (L3, M6) inside the deleted range.
    OpenpyxlXlsxValidator.validate(&plan, &output)?;
    UmyaXlsxValidator.validate(&plan, &output)?;

    let reopened = excel_editor::open_xlsx(&output)?;
    let history = reopened.sheet_by_name("History")?;
    assert_eq!(
        history.cell("A2").map(|cell| cell.value().to_string()),
        Some("新内容".to_owned())
    );
    // Net delete+insert of the same count keeps rows below the range in
    // place; their values and styles survive.
    assert_eq!(
        history.cell("A7").map(|cell| cell.value().to_string()),
        Some("below".to_owned())
    );
    assert_eq!(
        history.cell("A8").map(|cell| cell.value().to_string()),
        Some("below2".to_owned())
    );
    assert!(
        history.cell("L3").is_none() || history.cell("L3").unwrap().value().to_string().is_empty()
    );
    Ok(())
}

/// Data sheet with content below the insertion point: values, a style-only
/// cell, a formula and a comment all shift down by the inserted count.
fn write_shift_input(path: &Path) -> Result<()> {
    let mut workbook = umya_spreadsheet::new_file();
    let sheet = workbook.sheet_by_name_mut("Sheet1")?;
    sheet.set_name("Data");
    sheet.cell_mut("A1").set_value_string("head");
    sheet.cell_mut("A2").set_value_string("old2");
    sheet.cell_mut("A3").set_value_string("old3");
    sheet.cell_mut("A5").set_value_string("old5");
    fill_style(sheet.get_style_mut("B3"), "FFFFEB9C");
    sheet
        .cell_mut("C3")
        .set_formula("A3+1")
        .set_formula_result_number(4.0);
    let mut comment = umya_spreadsheet::Comment::default();
    comment.new_comment("A3");
    comment.set_author("test");
    comment.set_text_string("shifts down");
    sheet.add_comments(comment);
    umya_spreadsheet::writer::xlsx::write(&workbook, path)?;
    Ok(())
}

#[test]
fn insertion_shift_rows_validates_through_mapped_coordinates() -> Result<()> {
    let dir = TestDir::new();
    let input = dir.path("input.xlsx");
    let output = dir.path("output.xlsx");
    write_shift_input(&input)?;

    let profile = inspect_xlsx_profile(&input)?;
    let plan = ExcelEditPlan {
        request_id: "row-map-insert-shift".to_owned(),
        input: input.clone(),
        output: output.clone(),
        format: WorkbookFormat::Xlsx,
        workbook: profile.clone(),
        operations: vec![EditOperation::InsertRows {
            sheet: "Data".to_owned(),
            first_row: 2,
            count: 2,
        }],
        preserve: [
            PreservationRequirement::CellValuesAndTypes,
            PreservationRequirement::FormulaText,
            PreservationRequirement::CommentsAndPositions,
            PreservationRequirement::Styles,
        ]
        .into_iter()
        .collect(),
        drop_allowed: Default::default(),
    };
    let backend = UmyaXlsxBackend::new(BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Unverified,
        format: WorkbookFormat::Xlsx,
        operations: [EditOperationKind::InsertRows].into_iter().collect(),
        supported_features: profile.features_present,
        preserved: plan.preserve.clone(),
        platform: "synthetic-row-map-test".to_owned(),
        evidence: None,
    })?;
    backend
        .apply(&plan)
        .map_err(|failure| anyhow::anyhow!(failure.message))?;

    OpenpyxlXlsxValidator.validate(&plan, &output)?;
    UmyaXlsxValidator.validate(&plan, &output)?;

    // The engine really shifted the rows; the validator compared them through
    // the same mapping instead of failing at the source coordinates.
    let reopened = excel_editor::open_xlsx(&output)?;
    let data = reopened.sheet_by_name("Data")?;
    assert_eq!(
        data.cell("A4").map(|cell| cell.value().to_string()),
        Some("old2".to_owned())
    );
    assert_eq!(
        data.cell("A5").map(|cell| cell.value().to_string()),
        Some("old3".to_owned())
    );
    assert_eq!(
        data.cell("A7").map(|cell| cell.value().to_string()),
        Some("old5".to_owned())
    );
    assert_eq!(data.comments()[0].coordinate().to_string(), "A5".to_owned());
    assert_eq!(data.cell("C5").map(|cell| cell.formula()), Some("A5+1"));
    Ok(())
}
