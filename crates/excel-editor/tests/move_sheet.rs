use anyhow::Result;
use excel_editor::{
    inspect_xlsx_profile, BackendCapability, CapabilityEvidence, CapabilityStatus, EditBackend,
    EditOperation, EditOperationKind, EngineId, ExcelEditPlan, ExcelModifier, OutputValidator,
    PreservationRequirement, UmyaXlsxBackend, UmyaXlsxValidator, WorkbookFormat,
};
use excel_reader::{read_defined_names, read_sheet_names, DefinedNameScope};
use rust_xlsxwriter::Workbook;
use std::{
    fs,
    path::{Path, PathBuf},
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
        let path = std::env::temp_dir().join(format!("excel-editor-move-sheet-{nonce}"));
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

/// Three sheets, each with a sheet-scoped print area and its own value; the
/// History sheet is landscape. The named worksheet is left active so a
/// structural change must keep following it with the workbook `activeTab`.
fn write_input(path: &Path, active_sheet: &str) -> Result<()> {
    let mut workbook = Workbook::new();
    {
        let summary = workbook.add_worksheet();
        summary.set_name("Summary")?;
        summary.write_string(0, 0, "summary-data")?;
        summary.set_print_area(0, 0, 1, 1)?;
        if active_sheet == "Summary" {
            summary.set_active(true);
        }
    }
    {
        let config = workbook.add_worksheet();
        config.set_name("Config")?;
        config.write_string(0, 0, "config-data")?;
        config.set_print_area(0, 0, 2, 2)?;
        if active_sheet == "Config" {
            config.set_active(true);
        }
    }
    {
        let history = workbook.add_worksheet();
        history.set_name("History")?;
        history.write_string(0, 0, "history-data")?;
        history.set_print_area(0, 0, 3, 3)?;
        history.set_landscape();
        if active_sheet == "History" {
            history.set_active(true);
        }
    }
    workbook.save(path)?;
    Ok(())
}

/// (defined name, scoped worksheet or empty for workbook scope, raw reference)
/// for every defined name whose scope resolves to the workbook or a named
/// worksheet.
fn scoped_defined_name_keys(path: &Path) -> Result<Vec<(String, String, String)>> {
    let mut keys = Vec::new();
    for record in read_defined_names(path)? {
        let scoped_sheet = match &record.scope {
            DefinedNameScope::Workbook => String::new(),
            DefinedNameScope::Worksheet {
                sheet_name: Some(sheet),
                ..
            } => sheet.clone(),
            _ => continue,
        };
        keys.push((record.name, scoped_sheet, record.raw_reference));
    }
    keys.sort();
    Ok(keys)
}

/// (worksheet, localSheetId) recorded for each sheet-scoped `_xlnm.Print_Area`.
fn print_area_scopes(path: &Path) -> Result<Vec<(String, usize)>> {
    let mut scopes = Vec::new();
    for record in read_defined_names(path)? {
        if record.name != "_xlnm.Print_Area" {
            continue;
        }
        match &record.scope {
            DefinedNameScope::Worksheet {
                local_sheet_id,
                sheet_name: Some(sheet),
                ..
            } => scopes.push((sheet.clone(), *local_sheet_id)),
            other => anyhow::bail!("unexpected Print_Area scope: {other:?}"),
        }
    }
    scopes.sort();
    Ok(scopes)
}

#[test]
fn move_sheet_to_front_preserves_sheet_metadata_and_active_tab() -> Result<()> {
    let dir = TestDir::new();
    let input = dir.path("input.xlsx");
    let output = dir.path("output.xlsx");
    write_input(&input, "History")?;

    let profile = inspect_xlsx_profile(&input)?;
    let platform = "synthetic-move-sheet-test";
    let plan = ExcelEditPlan {
        request_id: "move-sheet-to-front".to_owned(),
        input: input.clone(),
        output: output.clone(),
        format: WorkbookFormat::Xlsx,
        workbook: profile.clone(),
        operations: vec![EditOperation::MoveSheet {
            sheet: "History".to_owned(),
            new_index: 0,
        }],
        preserve: [
            PreservationRequirement::SheetIdentityAndOrder,
            PreservationRequirement::DefinedNames,
            PreservationRequirement::CellValuesAndTypes,
        ]
        .into_iter()
        .collect(),
        drop_allowed: Default::default(),
    };
    let capability = BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Passed,
        format: WorkbookFormat::Xlsx,
        operations: [EditOperationKind::MoveSheet].into_iter().collect(),
        supported_features: profile.features_present,
        preserved: plan.preserve.clone(),
        platform: platform.to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: "synthetic-move-sheet".to_owned(),
            report_path: "crates/excel-editor/tests/move_sheet.rs".to_owned(),
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

    assert_eq!(read_sheet_names(&output)?, ["History", "Summary", "Config"]);
    assert_eq!(
        scoped_defined_name_keys(&input)?,
        scoped_defined_name_keys(&output)?
    );
    assert_eq!(
        print_area_scopes(&output)?,
        [
            ("Config".to_owned(), 2usize),
            ("History".to_owned(), 0),
            ("Summary".to_owned(), 1),
        ]
    );

    let reopened = excel_editor::open_xlsx(&output)?;
    assert_eq!(reopened.active_sheet().name(), "History");
    assert_eq!(
        reopened
            .sheet_by_name("History")?
            .cell("A1")
            .map(|cell| cell.value().to_string()),
        Some("history-data".to_owned())
    );
    // umya normalizes a missing pageSetup orientation to `Portrait` on read,
    // so compare source and candidate per sheet instead of asserting defaults.
    let reopened_input = excel_editor::open_xlsx(&input)?;
    for name in ["History", "Summary", "Config"] {
        assert_eq!(
            format!(
                "{:?}",
                reopened_input
                    .sheet_by_name(name)?
                    .page_setup()
                    .orientation()
            ),
            format!(
                "{:?}",
                reopened.sheet_by_name(name)?.page_setup().orientation()
            ),
            "orientation of worksheet {name} changed by the move"
        );
    }
    assert_eq!(
        format!(
            "{:?}",
            reopened
                .sheet_by_name("History")?
                .page_setup()
                .orientation()
        ),
        "Landscape"
    );
    Ok(())
}

#[test]
fn move_sheet_forward_keeps_scoped_names_and_active_tab_following() -> Result<()> {
    let dir = TestDir::new();
    let input = dir.path("input.xlsx");
    let output = dir.path("output.xlsx");
    write_input(&input, "History")?;

    let profile = inspect_xlsx_profile(&input)?;
    let platform = "synthetic-move-sheet-test";
    let plan = ExcelEditPlan {
        request_id: "move-sheet-forward".to_owned(),
        input: input.clone(),
        output: output.clone(),
        format: WorkbookFormat::Xlsx,
        workbook: profile.clone(),
        operations: vec![EditOperation::MoveSheet {
            sheet: "Summary".to_owned(),
            new_index: 2,
        }],
        preserve: [
            PreservationRequirement::SheetIdentityAndOrder,
            PreservationRequirement::DefinedNames,
            PreservationRequirement::CellValuesAndTypes,
        ]
        .into_iter()
        .collect(),
        drop_allowed: Default::default(),
    };
    let backend = Arc::new(UmyaXlsxBackend::new(BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Passed,
        format: WorkbookFormat::Xlsx,
        operations: [EditOperationKind::MoveSheet].into_iter().collect(),
        supported_features: profile.features_present,
        preserved: plan.preserve.clone(),
        platform: platform.to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: "synthetic-move-sheet".to_owned(),
            report_path: "crates/excel-editor/tests/move_sheet.rs".to_owned(),
            adapter_version: "umya-spreadsheet-3.1.1".to_owned(),
            engine_version: "umya-spreadsheet-3.1.1".to_owned(),
            platform: platform.to_owned(),
        }),
    })?);
    let validator: Arc<dyn OutputValidator> = Arc::new(UmyaXlsxValidator);
    let modifier = ExcelModifier::new(platform, vec![backend], validator)?;
    modifier.execute(&plan)?;

    assert_eq!(read_sheet_names(&output)?, ["Config", "History", "Summary"]);
    assert_eq!(
        scoped_defined_name_keys(&input)?,
        scoped_defined_name_keys(&output)?
    );
    assert_eq!(
        print_area_scopes(&output)?,
        [
            ("Config".to_owned(), 0usize),
            ("History".to_owned(), 1),
            ("Summary".to_owned(), 2),
        ]
    );

    let reopened = excel_editor::open_xlsx(&output)?;
    assert_eq!(reopened.active_sheet().name(), "History");
    Ok(())
}

#[test]
fn move_sheet_rejects_unknown_sheets_and_out_of_range_targets_before_write() -> Result<()> {
    let dir = TestDir::new();
    let input = dir.path("input.xlsx");
    let output = dir.path("output.xlsx");
    write_input(&input, "History")?;

    let profile = inspect_xlsx_profile(&input)?;
    let backend = UmyaXlsxBackend::new(BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Unverified,
        format: WorkbookFormat::Xlsx,
        operations: [EditOperationKind::MoveSheet].into_iter().collect(),
        supported_features: profile.features_present.clone(),
        preserved: Default::default(),
        platform: "synthetic-move-sheet-test".to_owned(),
        evidence: None,
    })?;

    let missing = ExcelEditPlan {
        request_id: "move-sheet-missing".to_owned(),
        input: input.clone(),
        output: output.clone(),
        format: WorkbookFormat::Xlsx,
        workbook: profile.clone(),
        operations: vec![EditOperation::MoveSheet {
            sheet: "Missing".to_owned(),
            new_index: 0,
        }],
        preserve: [PreservationRequirement::SheetIdentityAndOrder]
            .into_iter()
            .collect(),
        drop_allowed: Default::default(),
    };
    let failure = backend.apply(&missing).unwrap_err();
    assert!(
        failure.message.contains("not found"),
        "unexpected failure: {}",
        failure.message
    );
    assert!(!output.exists());

    let out_of_range = ExcelEditPlan {
        request_id: "move-sheet-out-of-range".to_owned(),
        input: input.clone(),
        output: output.clone(),
        format: WorkbookFormat::Xlsx,
        workbook: profile,
        operations: vec![EditOperation::MoveSheet {
            sheet: "History".to_owned(),
            new_index: 5,
        }],
        preserve: [PreservationRequirement::SheetIdentityAndOrder]
            .into_iter()
            .collect(),
        drop_allowed: Default::default(),
    };
    let failure = backend.apply(&out_of_range).unwrap_err();
    assert!(
        failure.message.contains("exceeds workbook sheet count"),
        "unexpected failure: {}",
        failure.message
    );
    assert!(!output.exists());
    Ok(())
}
