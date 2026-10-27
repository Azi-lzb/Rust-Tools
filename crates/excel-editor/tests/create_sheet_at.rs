use anyhow::Result;
use excel_editor::{
    BackendCapability, CapabilityEvidence, CapabilityStatus, EditOperation, EditOperationKind,
    EngineId, ExcelEditPlan, ExcelModifier, OutputValidator, PreservationRequirement,
    UmyaXlsxBackend, UmyaXlsxValidator, WorkbookFormat,
};
use excel_reader::read_workbook;
use rust_xlsxwriter::Workbook;
use std::{
    fs,
    path::PathBuf,
    sync::Arc,
    time::{SystemTime, UNIX_EPOCH},
};

struct TestDir(PathBuf);

impl TestDir {
    fn new() -> Self {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("system time is after epoch")
            .as_nanos();
        let path = std::env::temp_dir().join(format!("excel-editor-create-sheet-at-{nonce}"));
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
fn create_sheet_at_inserts_and_preserves_existing_sheet_order() -> Result<()> {
    let dir = TestDir::new();
    let input = dir.path("input.xlsx");
    let output = dir.path("output.xlsx");
    let mut workbook = Workbook::new();
    for name in ["Before", "Target", "After"] {
        let sheet = workbook.add_worksheet();
        sheet.set_name(name)?;
        sheet.write_string(0, 0, name)?;
    }
    workbook.save(&input)?;

    let profile = excel_editor::inspect_xlsx_profile(&input)?;
    let platform = "synthetic-create-sheet-at-test";
    let plan = ExcelEditPlan {
        request_id: "create-sheet-at-test".to_owned(),
        input: input.clone(),
        output: output.clone(),
        format: WorkbookFormat::Xlsx,
        workbook: profile.clone(),
        operations: vec![EditOperation::CreateSheetAt {
            name: "Inserted".to_owned(),
            index: 1,
        }],
        preserve: [PreservationRequirement::SheetIdentityAndOrder]
            .into_iter()
            .collect(),
        drop_allowed: Default::default(),
    };
    let capability = BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Passed,
        format: WorkbookFormat::Xlsx,
        operations: [EditOperationKind::CreateSheetAt].into_iter().collect(),
        supported_features: profile.features_present,
        preserved: plan.preserve.clone(),
        platform: platform.to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: "synthetic-sheet-order".to_owned(),
            report_path: "crates/excel-editor/tests/create_sheet_at.rs".to_owned(),
            adapter_version: "umya-spreadsheet-3.1.1".to_owned(),
            engine_version: "umya-spreadsheet-3.1.1".to_owned(),
            platform: platform.to_owned(),
        }),
    };
    let backend = Arc::new(UmyaXlsxBackend::new(capability)?);
    let validator: Arc<dyn OutputValidator> = Arc::new(UmyaXlsxValidator);
    let modifier = ExcelModifier::new(platform, vec![backend], validator)?;
    let report = modifier.execute(&plan)?;

    assert_eq!(report.engine, EngineId::UmyaSpreadsheet);
    let actual = read_workbook(&output)?
        .sheets
        .into_iter()
        .map(|sheet| sheet.name)
        .collect::<Vec<_>>();
    assert_eq!(actual, ["Before", "Inserted", "Target", "After"]);
    Ok(())
}
