use anyhow::Result;
use excel_editor::{
    inspect_xlsx_profile, BackendCapability, CapabilityStatus, EditBackend, EditOperation,
    EditOperationKind, EngineId, ExcelEditPlan, OutputValidator, PreservationRequirement,
    UmyaXlsxBackend, UmyaXlsxValidator, WorkbookFormat,
};
use std::{
    fs,
    path::PathBuf,
    time::{SystemTime, UNIX_EPOCH},
};

struct TestDir(PathBuf);

impl TestDir {
    fn new() -> Self {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("system clock after epoch")
            .as_nanos();
        let path = std::env::temp_dir().join(format!("excel-editor-insert-rows-{nonce}"));
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

#[test]
fn insertion_is_scoped_to_target_sheet_and_updates_references() -> Result<()> {
    let dir = TestDir::new();
    let input = dir.path("input.xlsx");
    let output = dir.path("output.xlsx");

    let mut workbook = umya_spreadsheet::new_file();
    let target = workbook.sheet_by_name_mut("Sheet1")?;
    target.set_name("Target");
    target.cell_mut("A1").set_value_string("above");
    target.cell_mut("A2").set_value_string("moves");
    target
        .cell_mut("B2")
        .set_formula("A2+10")
        .set_formula_result_number(10.0);
    target
        .cell_mut("C2")
        .set_formula("$A$2+20")
        .set_formula_result_number(20.0);
    target.cell_mut("A5").set_value_string("lower");
    target.add_merge_cells("D2:E2");
    target.row_dimension_mut(2).set_height(24.0);
    let mut target_comment = umya_spreadsheet::Comment::default();
    target_comment.new_comment("F2");
    target_comment.set_author("test");
    target_comment.set_text_string("target comment");
    target.add_comments(target_comment);

    let other = workbook.new_sheet("Other")?;
    other.cell_mut("A2").set_value_string("must stay at row 2");
    other.cell_mut("A5").set_value_string("must stay at row 5");
    other
        .cell_mut("B2")
        .set_formula("Target!A2")
        .set_formula_result_number(0.0);
    other
        .cell_mut("C2")
        .set_formula("A2+1")
        .set_formula_result_number(0.0);
    other.add_merge_cells("D2:E2");
    other.row_dimension_mut(5).set_height(19.0);
    let mut other_comment = umya_spreadsheet::Comment::default();
    other_comment.new_comment("F5");
    other_comment.set_author("test");
    other_comment.set_text_string("other comment");
    other.add_comments(other_comment);

    let unmatched = workbook.new_sheet("Unmatched")?;
    unmatched.cell_mut("A2").set_value_string("unmatched stays");
    unmatched.cell_mut("A4").set_value_string("also stays");

    let mut target_name = umya_spreadsheet::DefinedName::default();
    target_name
        .set_name("TargetSpan")
        .set_address("Target!$A$2:$A$5");
    workbook.add_defined_names(target_name);
    umya_spreadsheet::writer::xlsx::write(&workbook, &input)?;
    assert_eq!(
        excel_editor::open_xlsx(&input)?
            .sheet_by_name("Target")?
            .defined_names()
            .len(),
        1
    );

    let profile = inspect_xlsx_profile(&input)?;
    let plan = ExcelEditPlan {
        request_id: "insert-rows-sheet-scope".to_owned(),
        input: input.clone(),
        output: output.clone(),
        format: WorkbookFormat::Xlsx,
        workbook: profile.clone(),
        operations: vec![EditOperation::InsertRows {
            sheet: "Target".to_owned(),
            first_row: 2,
            count: 2,
        }],
        preserve: [
            PreservationRequirement::CellValuesAndTypes,
            PreservationRequirement::FormulaText,
            PreservationRequirement::CommentsAndPositions,
            PreservationRequirement::MergedRanges,
            PreservationRequirement::DefinedNames,
            PreservationRequirement::RowAndColumnDimensions,
            PreservationRequirement::SheetIdentityAndOrder,
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
        platform: "synthetic-insert-rows-test".to_owned(),
        evidence: None,
    })?;

    backend
        .apply(&plan)
        .map_err(|failure| anyhow::anyhow!(failure.message))?;
    UmyaXlsxValidator.validate(&plan, &output)?;

    let actual = excel_editor::open_xlsx(&output)?;
    let target = actual.sheet_by_name("Target")?;
    assert_eq!(
        target.cell("A1").map(|cell| cell.value().to_string()),
        Some("above".to_owned())
    );
    assert!(target.cell("A2").is_none());
    assert_eq!(
        target.cell("A4").map(|cell| cell.value().to_string()),
        Some("moves".to_owned())
    );
    assert_eq!(
        target.cell("A7").map(|cell| cell.value().to_string()),
        Some("lower".to_owned())
    );
    assert_eq!(target.cell("B4").map(|cell| cell.formula()), Some("A4+10"));
    assert_eq!(
        target.cell("C4").map(|cell| cell.formula()),
        Some("$A$4+20")
    );
    assert_eq!(target.merge_cells()[0].range(), "D4:E4");
    assert_eq!(target.row_dimension(4).map(|row| row.height()), Some(24.0));
    assert_eq!(target.comments()[0].coordinate().to_string(), "F4");

    let other = actual.sheet_by_name("Other")?;
    assert_eq!(
        other.cell("A2").map(|cell| cell.value().to_string()),
        Some("must stay at row 2".to_owned())
    );
    assert_eq!(
        other.cell("A5").map(|cell| cell.value().to_string()),
        Some("must stay at row 5".to_owned())
    );
    assert_eq!(
        other.cell("B2").map(|cell| cell.formula()),
        Some("Target!A4")
    );
    assert_eq!(other.cell("C2").map(|cell| cell.formula()), Some("A2+1"));
    assert_eq!(other.merge_cells()[0].range(), "D2:E2");
    assert_eq!(other.row_dimension(5).map(|row| row.height()), Some(19.0));
    assert_eq!(other.comments()[0].coordinate().to_string(), "F5");

    let unmatched = actual.sheet_by_name("Unmatched")?;
    assert_eq!(
        unmatched.cell("A2").map(|cell| cell.value().to_string()),
        Some("unmatched stays".to_owned())
    );
    assert_eq!(
        unmatched.cell("A4").map(|cell| cell.value().to_string()),
        Some("also stays".to_owned())
    );
    assert_eq!(
        actual.sheet_by_name("Target")?.defined_names()[0].address(),
        "Target!$A$4:$A$7"
    );
    Ok(())
}
