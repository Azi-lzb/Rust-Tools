//! Regression coverage for workbook-position metadata after structural sheet
//! changes: umya-spreadsheet re-emits the `localSheetId` captured at read time
//! and stores `activeTab` as a raw index, so every operation that shifts sheet
//! positions must rewrite both. Without the rewrite, sheet-scoped print areas
//! silently re-bind to whichever worksheet ends up at the stale index.

use anyhow::Result;
use excel_editor::{
    inspect_xlsx_profile, BackendCapability, CapabilityStatus, EditBackend, EditOperation,
    EditOperationKind, EngineId, ExcelEditPlan, OutputValidator, PreservationRequirement,
    UmyaXlsxBackend, UmyaXlsxValidator, WorkbookFormat,
};
use excel_reader::{read_defined_names, read_sheet_names, DefinedNameScope};
use rust_xlsxwriter::Workbook;
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
            .expect("system time is after epoch")
            .as_nanos();
        let path = std::env::temp_dir().join(format!("excel-editor-sheet-index-{nonce}"));
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
        if active_sheet == "History" {
            history.set_active(true);
        }
    }
    workbook.save(path)?;
    Ok(())
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

fn structural_plan(
    request_id: &str,
    input: &Path,
    output: &Path,
    operation: EditOperation,
) -> Result<ExcelEditPlan> {
    let profile = inspect_xlsx_profile(input)?;
    Ok(ExcelEditPlan {
        request_id: request_id.to_owned(),
        input: input.to_path_buf(),
        output: output.to_path_buf(),
        format: WorkbookFormat::Xlsx,
        workbook: profile,
        operations: vec![operation],
        preserve: [
            PreservationRequirement::SheetIdentityAndOrder,
            PreservationRequirement::DefinedNames,
        ]
        .into_iter()
        .collect(),
        drop_allowed: Default::default(),
    })
}

fn umya_backend(
    profile_features: std::collections::BTreeSet<excel_editor::WorkbookFeature>,
) -> Result<UmyaXlsxBackend> {
    UmyaXlsxBackend::new(BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Unverified,
        format: WorkbookFormat::Xlsx,
        operations: [
            EditOperationKind::CreateSheetAt,
            EditOperationKind::DeleteSheet,
        ]
        .into_iter()
        .collect(),
        supported_features: profile_features,
        preserved: Default::default(),
        platform: "synthetic-sheet-index-metadata-test".to_owned(),
        evidence: None,
    })
}

#[test]
fn create_sheet_at_remaps_scoped_print_areas_and_active_tab() -> Result<()> {
    let dir = TestDir::new();
    let input = dir.path("input.xlsx");
    let output = dir.path("output.xlsx");
    write_input(&input, "History")?;

    let profile = inspect_xlsx_profile(&input)?;
    let plan = structural_plan(
        "create-sheet-at-print-areas",
        &input,
        &output,
        EditOperation::CreateSheetAt {
            name: "New".to_owned(),
            index: 0,
        },
    )?;
    let backend = umya_backend(profile.features_present)?;
    backend
        .apply(&plan)
        .map_err(|failure| anyhow::anyhow!(failure.message))?;
    UmyaXlsxValidator.validate(&plan, &output)?;

    assert_eq!(
        read_sheet_names(&output)?,
        ["New", "Summary", "Config", "History"]
    );
    assert_eq!(
        print_area_scopes(&output)?,
        [
            ("Config".to_owned(), 2usize),
            ("History".to_owned(), 3),
            ("Summary".to_owned(), 1),
        ]
    );
    let reopened = excel_editor::open_xlsx(&output)?;
    assert_eq!(reopened.active_sheet().name(), "History");
    Ok(())
}

#[test]
fn delete_sheet_remaps_scoped_print_areas_and_follows_active_tab() -> Result<()> {
    let dir = TestDir::new();
    let input = dir.path("input.xlsx");
    let output = dir.path("output.xlsx");
    write_input(&input, "History")?;

    let profile = inspect_xlsx_profile(&input)?;
    let plan = structural_plan(
        "delete-sheet-print-areas",
        &input,
        &output,
        EditOperation::DeleteSheet {
            name: "Config".to_owned(),
        },
    )?;
    let backend = umya_backend(profile.features_present)?;
    backend
        .apply(&plan)
        .map_err(|failure| anyhow::anyhow!(failure.message))?;
    UmyaXlsxValidator.validate(&plan, &output)?;

    assert_eq!(read_sheet_names(&output)?, ["Summary", "History"]);
    assert_eq!(
        print_area_scopes(&output)?,
        [("History".to_owned(), 1usize), ("Summary".to_owned(), 0),]
    );
    let reopened = excel_editor::open_xlsx(&output)?;
    assert_eq!(reopened.active_sheet().name(), "History");
    Ok(())
}

#[test]
fn delete_active_sheet_clamps_the_active_tab_into_range() -> Result<()> {
    let dir = TestDir::new();
    let input = dir.path("input.xlsx");
    let output = dir.path("output.xlsx");
    write_input(&input, "Config")?;

    let profile = inspect_xlsx_profile(&input)?;
    let plan = structural_plan(
        "delete-active-sheet-clamp",
        &input,
        &output,
        EditOperation::DeleteSheet {
            name: "Config".to_owned(),
        },
    )?;
    let backend = umya_backend(profile.features_present)?;
    backend
        .apply(&plan)
        .map_err(|failure| anyhow::anyhow!(failure.message))?;
    UmyaXlsxValidator.validate(&plan, &output)?;

    assert_eq!(read_sheet_names(&output)?, ["Summary", "History"]);
    assert_eq!(
        print_area_scopes(&output)?,
        [("History".to_owned(), 1usize), ("Summary".to_owned(), 0),]
    );
    let reopened = excel_editor::open_xlsx(&output)?;
    assert_eq!(reopened.active_sheet().name(), "History");
    Ok(())
}
