//! Synthetic coverage for umya freeze-pane and AutoFilter edits, registered
//! by the Audit S1 flow: freeze panes follow openpyxl semantics, and the
//! hidden sheet-scoped `_xlnm._FilterDatabase` name must follow the requested
//! AutoFilter range the way Excel and the openpyxl worker maintain it.

use anyhow::Result;
use excel_editor::{
    inspect_xlsx_profile, BackendCapability, BackendErrorKind, CapabilityStatus, EditBackend,
    EditOperation, EditOperationKind, EngineId, ExcelEditPlan, OutputValidator,
    PreservationRequirement, UmyaXlsxBackend, UmyaXlsxValidator, WorkbookFormat,
};
use excel_reader::{read_defined_names, DefinedNameScope};
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
        let path = std::env::temp_dir().join(format!("excel-editor-freeze-filter-{nonce}"));
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

fn write_plain_input(path: &Path) -> Result<()> {
    let mut workbook = Workbook::new();
    {
        let history = workbook.add_worksheet();
        history.set_name("历史")?;
        history.write_string(0, 0, "表头")?;
        history.write_string(1, 0, "数据")?;
    }
    {
        let summary = workbook.add_worksheet();
        summary.set_name("Summary")?;
        summary.write_string(0, 0, "summary")?;
    }
    workbook.save(path)?;
    Ok(())
}

/// History carries an existing freeze at A2 and an AutoFilter over A1:L10
/// (which rust_xlsxwriter records as a sheet-scoped `_xlnm._FilterDatabase`).
fn write_existing_input(path: &Path) -> Result<()> {
    let mut workbook = Workbook::new();
    {
        let history = workbook.add_worksheet();
        history.set_name("History")?;
        history.write_string(0, 0, "header")?;
        history.write_string(1, 0, "data")?;
        history.set_freeze_panes(1, 0)?;
        history.autofilter(0, 0, 9, 11)?;
    }
    {
        let plain = workbook.add_worksheet();
        plain.set_name("Plain")?;
        plain.write_string(0, 0, "plain")?;
    }
    workbook.save(path)?;
    Ok(())
}

fn umya_backend(
    features: std::collections::BTreeSet<excel_editor::WorkbookFeature>,
    operations: &[EditOperationKind],
) -> Result<UmyaXlsxBackend> {
    UmyaXlsxBackend::new(BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Unverified,
        format: WorkbookFormat::Xlsx,
        operations: operations.iter().copied().collect(),
        supported_features: features,
        preserved: Default::default(),
        platform: "synthetic-freeze-filter-test".to_owned(),
        evidence: None,
    })
}

/// (localSheetId, raw reference) of the sheet-scoped `_xlnm._FilterDatabase`
/// record for the named worksheet, if any.
fn filter_database_record(path: &Path, sheet: &str) -> Result<Option<(usize, String)>> {
    for record in read_defined_names(path)? {
        if record.name != "_xlnm._FilterDatabase" {
            continue;
        }
        if let DefinedNameScope::Worksheet {
            local_sheet_id,
            sheet_name: Some(name),
            ..
        } = &record.scope
        {
            if name == sheet {
                return Ok(Some((*local_sheet_id, record.raw_reference.clone())));
            }
        }
    }
    Ok(None)
}

fn filter_database_hidden(path: &Path, sheet: &str) -> Result<Option<bool>> {
    let workbook = excel_editor::open_xlsx(path)?;
    let worksheet = workbook.sheet_by_name(sheet)?;
    Ok(worksheet
        .defined_names()
        .iter()
        .find(|record| record.name() == "_xlnm._FilterDatabase")
        .map(|record| record.hidden()))
}

#[test]
fn set_creates_pane_filter_and_filter_database_on_non_ascii_sheet() -> Result<()> {
    let dir = TestDir::new();
    let input = dir.path("input.xlsx");
    let output = dir.path("output.xlsx");
    write_plain_input(&input)?;

    let profile = inspect_xlsx_profile(&input)?;
    let plan = ExcelEditPlan {
        request_id: "freeze-filter-set".to_owned(),
        input: input.clone(),
        output: output.clone(),
        format: WorkbookFormat::Xlsx,
        workbook: profile.clone(),
        operations: vec![
            EditOperation::SetFreezePanes {
                sheet: "历史".to_owned(),
                cell: Some("A2".to_owned()),
            },
            EditOperation::SetAutoFilter {
                sheet: "历史".to_owned(),
                range: Some("A1:M5".to_owned()),
            },
        ],
        preserve: [
            PreservationRequirement::CellValuesAndTypes,
            PreservationRequirement::FreezePanes,
            PreservationRequirement::AutoFilter,
        ]
        .into_iter()
        .collect(),
        drop_allowed: Default::default(),
    };
    let backend = umya_backend(
        profile.features_present,
        &[
            EditOperationKind::SetFreezePanes,
            EditOperationKind::SetAutoFilter,
        ],
    )?;
    backend
        .apply(&plan)
        .map_err(|failure| anyhow::anyhow!(failure.message))?;
    UmyaXlsxValidator.validate(&plan, &output)?;

    let reopened = excel_editor::open_xlsx(&output)?;
    let history = reopened.sheet_by_name("历史")?;
    let pane = history
        .sheets_views()
        .sheet_view_list()
        .iter()
        .find_map(|view| view.pane())
        .expect("frozen pane after SetFreezePanes");
    assert_eq!(format!("{:?}", pane.state()), "Frozen");
    assert_eq!(format!("{:?}", pane.active_pane()), "BottomLeft");
    assert_eq!(pane.horizontal_split(), 0.0);
    assert_eq!(pane.vertical_split(), 1.0);
    assert_eq!(
        (
            pane.top_left_cell().row_num(),
            pane.top_left_cell().col_num()
        ),
        (2, 1)
    );
    assert_eq!(
        history.auto_filter().map(|filter| filter.range().range()),
        Some("A1:M5".to_owned())
    );

    let summary = reopened.sheet_by_name("Summary")?;
    assert!(summary
        .sheets_views()
        .sheet_view_list()
        .iter()
        .all(|view| view.pane().is_none()));
    assert!(summary.auto_filter().is_none());

    // umya re-renders the non-ASCII sheet name quoted; the semantic contract
    // is the dequoted reference.
    let (local_sheet_id, raw_reference) =
        filter_database_record(&output, "历史")?.expect("filter database name after SetAutoFilter");
    assert_eq!(local_sheet_id, 0);
    assert_eq!(raw_reference.replace('\'', ""), "历史!$A$1:$M$5");
    assert_eq!(filter_database_hidden(&output, "历史")?, Some(true));
    Ok(())
}

#[test]
fn set_replaces_existing_pane_and_restores_filter_database_range() -> Result<()> {
    let dir = TestDir::new();
    let input = dir.path("input.xlsx");
    let output = dir.path("output.xlsx");
    write_existing_input(&input)?;
    let (initial_local_sheet_id, initial_reference) =
        filter_database_record(&input, "History")?.expect("input filter database name");
    assert_eq!(initial_reference.replace('\'', ""), "History!$A$1:$L$10");
    assert_eq!(filter_database_hidden(&input, "History")?, Some(true));

    let profile = inspect_xlsx_profile(&input)?;
    let plan = ExcelEditPlan {
        request_id: "freeze-filter-replace".to_owned(),
        input: input.clone(),
        output: output.clone(),
        format: WorkbookFormat::Xlsx,
        workbook: profile.clone(),
        operations: vec![
            EditOperation::SetFreezePanes {
                sheet: "History".to_owned(),
                cell: Some("B2".to_owned()),
            },
            EditOperation::SetAutoFilter {
                sheet: "History".to_owned(),
                range: Some("A1:C3".to_owned()),
            },
        ],
        preserve: [
            PreservationRequirement::CellValuesAndTypes,
            PreservationRequirement::FreezePanes,
            PreservationRequirement::AutoFilter,
        ]
        .into_iter()
        .collect(),
        drop_allowed: Default::default(),
    };
    let backend = umya_backend(
        profile.features_present,
        &[
            EditOperationKind::SetFreezePanes,
            EditOperationKind::SetAutoFilter,
        ],
    )?;
    backend
        .apply(&plan)
        .map_err(|failure| anyhow::anyhow!(failure.message))?;
    UmyaXlsxValidator.validate(&plan, &output)?;

    let reopened = excel_editor::open_xlsx(&output)?;
    let history = reopened.sheet_by_name("History")?;
    let pane = history
        .sheets_views()
        .sheet_view_list()
        .iter()
        .find_map(|view| view.pane())
        .expect("frozen pane after SetFreezePanes");
    assert_eq!(format!("{:?}", pane.active_pane()), "BottomRight");
    assert_eq!(pane.horizontal_split(), 1.0);
    assert_eq!(pane.vertical_split(), 1.0);
    assert_eq!(
        (
            pane.top_left_cell().row_num(),
            pane.top_left_cell().col_num()
        ),
        (2, 2)
    );
    assert_eq!(
        history.auto_filter().map(|filter| filter.range().range()),
        Some("A1:C3".to_owned())
    );
    let (local_sheet_id, raw_reference) = filter_database_record(&output, "History")?
        .expect("filter database name after SetAutoFilter");
    assert_eq!(local_sheet_id, initial_local_sheet_id);
    assert_eq!(raw_reference.replace('\'', ""), "History!$A$1:$C$3");
    assert_eq!(filter_database_hidden(&output, "History")?, Some(true));
    Ok(())
}

#[test]
fn clears_autofilter_and_rejects_freeze_clear_with_existing_pane() -> Result<()> {
    let dir = TestDir::new();
    let input = dir.path("input.xlsx");
    write_existing_input(&input)?;

    let profile = inspect_xlsx_profile(&input)?;
    let backend = umya_backend(
        profile.features_present.clone(),
        &[
            EditOperationKind::SetFreezePanes,
            EditOperationKind::SetAutoFilter,
        ],
    )?;

    // Clearing the AutoFilter removes its generated _FilterDatabase name and
    // leaves the frozen pane untouched.
    let filter_output = dir.path("filter-clear.xlsx");
    let filter_plan = ExcelEditPlan {
        request_id: "filter-clear".to_owned(),
        input: input.clone(),
        output: filter_output.clone(),
        format: WorkbookFormat::Xlsx,
        workbook: profile.clone(),
        operations: vec![EditOperation::SetAutoFilter {
            sheet: "History".to_owned(),
            range: None,
        }],
        preserve: [PreservationRequirement::FreezePanes].into_iter().collect(),
        drop_allowed: Default::default(),
    };
    backend
        .apply(&filter_plan)
        .map_err(|failure| anyhow::anyhow!(failure.message))?;
    UmyaXlsxValidator.validate(&filter_plan, &filter_output)?;
    let reopened = excel_editor::open_xlsx(&filter_output)?;
    assert!(reopened.sheet_by_name("History")?.auto_filter().is_none());
    assert!(filter_database_record(&filter_output, "History")?.is_none());
    assert!(reopened
        .sheet_by_name("History")?
        .sheets_views()
        .sheet_view_list()
        .iter()
        .any(|view| view.pane().is_some()));

    // Clearing the freeze on a sheet without a pane is a no-op.
    let freeze_clear_output = dir.path("freeze-clear-plain.xlsx");
    let freeze_clear_plan = ExcelEditPlan {
        request_id: "freeze-clear-plain".to_owned(),
        input: input.clone(),
        output: freeze_clear_output.clone(),
        format: WorkbookFormat::Xlsx,
        workbook: profile.clone(),
        operations: vec![EditOperation::SetFreezePanes {
            sheet: "Plain".to_owned(),
            cell: None,
        }],
        preserve: [PreservationRequirement::FreezePanes].into_iter().collect(),
        drop_allowed: Default::default(),
    };
    backend
        .apply(&freeze_clear_plan)
        .map_err(|failure| anyhow::anyhow!(failure.message))?;
    UmyaXlsxValidator.validate(&freeze_clear_plan, &freeze_clear_output)?;
    assert!(excel_editor::open_xlsx(&freeze_clear_output)?
        .sheet_by_name("Plain")?
        .sheets_views()
        .sheet_view_list()
        .iter()
        .all(|view| view.pane().is_none()));

    // Clearing the freeze on a sheet that still carries a pane is rejected
    // before any write: umya cannot drop the pane element.
    let pane_clear_output = dir.path("freeze-clear-pane.xlsx");
    let pane_clear_plan = ExcelEditPlan {
        request_id: "freeze-clear-pane".to_owned(),
        input: input.clone(),
        output: pane_clear_output.clone(),
        format: WorkbookFormat::Xlsx,
        workbook: profile,
        operations: vec![EditOperation::SetFreezePanes {
            sheet: "History".to_owned(),
            cell: None,
        }],
        preserve: [PreservationRequirement::FreezePanes].into_iter().collect(),
        drop_allowed: Default::default(),
    };
    let failure = backend.apply(&pane_clear_plan).unwrap_err();
    assert_eq!(failure.kind, BackendErrorKind::UnsupportedBeforeWrite);
    assert!(
        failure
            .message
            .contains("cannot clear an existing frozen pane"),
        "unexpected failure: {}",
        failure.message
    );
    assert!(!pane_clear_output.exists());
    Ok(())
}
