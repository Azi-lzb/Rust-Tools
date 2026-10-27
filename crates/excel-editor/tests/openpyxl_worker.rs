use anyhow::{Context, Result};
use excel_editor::{
    inspect_xlsx_profile, BackendCapability, CapabilityEvidence, CapabilityStatus, EditBackend,
    EditOperation, EditOperationKind, EngineId, ExcelEditPlan, ExcelModifier, OpenpyxlWorkerConfig,
    OpenpyxlXlsxBackend, OpenpyxlXlsxValidator, PreservationRequirement, WorkbookFeature,
    WorkbookFormat,
};
use rust_xlsxwriter::{Format, Workbook};
use std::{
    collections::BTreeSet,
    fs,
    path::{Path, PathBuf},
    time::Duration,
};

struct TestDir(PathBuf);

impl TestDir {
    fn new(name: &str) -> Self {
        let path = std::env::temp_dir().join(format!(
            "excel-editor-openpyxl-{}-{name}",
            std::process::id()
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

fn python_path() -> Option<PathBuf> {
    std::env::var_os("RUST_WORKSPACE_OPENPYXL_PYTHON").map(PathBuf::from)
}

fn required_python_path() -> PathBuf {
    python_path().expect("set RUST_WORKSPACE_OPENPYXL_PYTHON to an isolated openpyxl test runtime")
}

fn worker_path() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("python/openpyxl_worker.py")
}

fn make_source(path: &Path) -> Result<()> {
    let mut workbook = Workbook::new();
    {
        let sheet = workbook.add_worksheet();
        sheet.set_name("Data")?;
        sheet.write_string(0, 0, "original")?;
        sheet.write_number(0, 1, 3.0)?;
        sheet.merge_range(0, 3, 0, 4, "merged", &Format::new())?;
    }
    workbook.add_worksheet().set_name("Notes")?;
    workbook.save(path)?;
    Ok(())
}

fn make_import_source(path: &Path) -> Result<()> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Source")?;
    sheet.write_string_with_format(0, 0, "imported", &Format::new().set_bold())?;
    sheet.write_formula(0, 2, "=1+1")?;
    sheet.merge_range(1, 0, 1, 1, "import merge", &Format::new())?;
    workbook.save(path)?;
    Ok(())
}

fn make_perf_source(path: &Path, rows: u32, columns: u16) -> Result<()> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Data")?;
    for row in 0..rows {
        for column in 0..columns {
            sheet.write_number(row, column, (row as f64) * 10.0 + f64::from(column))?;
        }
    }
    workbook.save(path)?;
    Ok(())
}

fn plan(
    input: &Path,
    output: &Path,
    import_source: &Path,
    profile: excel_editor::WorkbookProfile,
) -> ExcelEditPlan {
    ExcelEditPlan {
        request_id: "openpyxl-rust-integration".to_owned(),
        input: input.to_path_buf(),
        output: output.to_path_buf(),
        format: WorkbookFormat::Xlsx,
        workbook: profile,
        operations: vec![
            EditOperation::SetCellValue {
                sheet: "Data".to_owned(),
                cell: "A1".to_owned(),
                value: excel_editor::CellValue::Text("changed".to_owned()),
            },
            EditOperation::SetFormulaText {
                sheet: "Data".to_owned(),
                cell: "C1".to_owned(),
                formula: "=B1*2".to_owned(),
            },
            EditOperation::SetComment {
                sheet: "Data".to_owned(),
                cell: "A1".to_owned(),
                text: Some("worker comment".to_owned()),
                author: Some("integration".to_owned()),
            },
            EditOperation::SetColumnWidth {
                sheet: "Data".to_owned(),
                column_index: 2,
                width: 18.0,
            },
            EditOperation::SetRowHeight {
                sheet: "Data".to_owned(),
                row_index: 2,
                height: 22.0,
            },
            EditOperation::SetPageSettings {
                sheet: "Data".to_owned(),
                settings: excel_editor::PageSettings {
                    orientation: Some(excel_editor::PageOrientation::Landscape),
                    fit_to_page: Some(true),
                    fit_to_width: Some(1),
                    fit_to_height: Some(0),
                    clear_scale: true,
                    paper_size: Some(9),
                    print_area: Some("A1:E3".to_owned()),
                    horizontal_centered: Some(true),
                    vertical_centered: Some(false),
                    footer: Some("&P".to_owned()),
                },
            },
            EditOperation::SetCellFillColor {
                sheet: "Data".to_owned(),
                cell: "A1".to_owned(),
                argb: "FF112233".to_owned(),
            },
            EditOperation::CopyRange {
                source_sheet: "Data".to_owned(),
                source_range: "A1:B1".to_owned(),
                target_sheet: "Data".to_owned(),
                target_start: "H5".to_owned(),
                copy_values: true,
                copy_formulas: false,
                copy_styles: true,
                copy_comments: true,
            },
            EditOperation::ImportRange {
                source_workbook: import_source.to_path_buf(),
                source_sheet: "Source".to_owned(),
                source_range: "A1:C2".to_owned(),
                target_sheet: "Data".to_owned(),
                target_start: "J1".to_owned(),
                copy_values: true,
                copy_formulas: true,
                style_mode: excel_editor::RangeStyleMode::Full,
                copy_comments: false,
                copy_merges: true,
                copy_row_heights: false,
                copy_column_widths: false,
                copy_row_hidden: false,
                copy_column_hidden: false,
                copy_outline_levels: false,
                copy_all_row_dimensions: false,
                copy_all_column_dimensions: false,
                materialize_blank_cells: true,
                hide_zero: false,
            },
            EditOperation::RenameSheet {
                old_name: "Notes".to_owned(),
                new_name: "Information".to_owned(),
            },
            EditOperation::CreateSheet {
                name: "Generated".to_owned(),
            },
            EditOperation::CreateSheetAt {
                name: "History".to_owned(),
                index: 0,
            },
            EditOperation::SetFreezePanes {
                sheet: "History".to_owned(),
                cell: Some("A2".to_owned()),
            },
            EditOperation::SetAutoFilter {
                sheet: "History".to_owned(),
                range: Some("A1:M2".to_owned()),
            },
            EditOperation::SetWorkbookCalculationProperties {
                full_calc_on_load: Some(true),
                force_full_calc: Some(true),
            },
        ],
        preserve: [
            PreservationRequirement::CellValuesAndTypes,
            PreservationRequirement::FormulaText,
            PreservationRequirement::CommentsAndPositions,
            PreservationRequirement::MergedRanges,
            PreservationRequirement::SheetIdentityAndOrder,
            PreservationRequirement::FreezePanes,
            PreservationRequirement::AutoFilter,
        ]
        .into_iter()
        .collect(),
        drop_allowed: Default::default(),
    }
}

fn single_value_plan(
    input: &Path,
    output: &Path,
    profile: excel_editor::WorkbookProfile,
    request_id: String,
) -> ExcelEditPlan {
    ExcelEditPlan {
        request_id,
        input: input.to_path_buf(),
        output: output.to_path_buf(),
        format: WorkbookFormat::Xlsx,
        workbook: profile,
        operations: vec![EditOperation::SetCellValue {
            sheet: "Data".to_owned(),
            cell: "A1".to_owned(),
            value: excel_editor::CellValue::Text("changed".to_owned()),
        }],
        preserve: [PreservationRequirement::CellValuesAndTypes]
            .into_iter()
            .collect(),
        drop_allowed: Default::default(),
    }
}

fn capability(plan: &ExcelEditPlan, python_version: &str) -> BackendCapability {
    let mut operations: BTreeSet<_> = plan.operations.iter().map(EditOperation::kind).collect();
    operations.remove(&EditOperationKind::DeleteRows);
    BackendCapability {
        engine: EngineId::OpenpyxlWorker,
        status: CapabilityStatus::Passed,
        format: WorkbookFormat::Xlsx,
        operations,
        supported_features: plan
            .workbook
            .features_present
            .union(
                &[
                    WorkbookFeature::Comments,
                    WorkbookFeature::Formulas,
                    WorkbookFeature::HeadersFooters,
                ]
                .into_iter()
                .collect(),
            )
            .copied()
            .collect(),
        preserved: plan.preserve.clone(),
        platform: "wsl2-x86_64-test".to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: "synthetic-worker-integration".to_owned(),
            report_path: "crates/excel-editor/tests/openpyxl_worker.rs".to_owned(),
            adapter_version: "1".to_owned(),
            engine_version: python_version.to_owned(),
            platform: "wsl2-x86_64-test".to_owned(),
        }),
    }
}

#[test]
fn missing_bundled_worker_is_not_routable() -> Result<()> {
    let dir = TestDir::new("missing-runtime");
    let input = dir.path("source.xlsx");
    let import_source = dir.path("import-source.xlsx");
    let output = dir.path("result.xlsx");
    make_source(&input)?;
    make_import_source(&import_source)?;
    let edit_plan = plan(
        &input,
        &output,
        &import_source,
        inspect_xlsx_profile(&input)?,
    );
    let absent_python = dir.path("missing-python");
    let backend = OpenpyxlXlsxBackend::new(
        capability(&edit_plan, "Python 3.12.3 / openpyxl 3.1.5"),
        OpenpyxlWorkerConfig::new(absent_python, worker_path())?,
    )?;
    assert_eq!(backend.capability().status, CapabilityStatus::Unsupported);
    let modifier = ExcelModifier::new(
        "wsl2-x86_64-test",
        vec![std::sync::Arc::new(backend)],
        std::sync::Arc::new(OpenpyxlXlsxValidator),
    )?;
    let error = modifier
        .execute(&edit_plan)
        .expect_err("missing worker must not be routable");
    assert_eq!(
        error.kind,
        excel_editor::ExcelModifierErrorKind::NoEligibleBackend
    );
    assert!(!output.exists());
    Ok(())
}

#[test]
#[ignore = "requires an isolated openpyxl test runtime; run with --ignored"]
fn rust_modifier_invokes_one_batch_worker_and_preserves_source() -> Result<()> {
    let python = required_python_path();
    let dir = TestDir::new("batch");
    let input = dir.path("source.xlsx");
    let import_source = dir.path("import-source.xlsx");
    let output = dir.path("result.xlsx");
    make_source(&input)?;
    make_import_source(&import_source)?;
    let source_bytes = fs::read(&input)?;
    let import_bytes = fs::read(&import_source)?;
    let profile = inspect_xlsx_profile(&input)?;
    let edit_plan = plan(&input, &output, &import_source, profile);
    let backend = OpenpyxlXlsxBackend::new(
        capability(&edit_plan, "Python 3.12.3 / openpyxl 3.1.5"),
        OpenpyxlWorkerConfig::new(python, worker_path())?,
    )?;
    let modifier = ExcelModifier::new(
        "wsl2-x86_64-test",
        vec![std::sync::Arc::new(backend)],
        std::sync::Arc::new(OpenpyxlXlsxValidator),
    )?;
    let report = modifier.execute(&edit_plan).map_err(anyhow::Error::from)?;
    assert_eq!(report.engine, EngineId::OpenpyxlWorker);
    assert!(
        report.elapsed_ms > 0,
        "execution report must include worker/process and validation time"
    );
    assert_eq!(
        report.attempts.len(),
        1,
        "one worker process must execute the complete workbook plan"
    );
    assert_eq!(
        fs::read(&input)?,
        source_bytes,
        "source must remain byte-identical"
    );
    assert_eq!(
        fs::read(&import_source)?,
        import_bytes,
        "ImportRange source must remain byte-identical"
    );
    let dump = excel_reader::read_workbook(&output)?;
    assert!(dump.sheets.iter().any(|sheet| sheet.name == "Information"));
    assert_eq!(dump.sheets[0].name, "History");
    assert_eq!(dump.sheets[1].name, "Data");
    let data = dump
        .sheets
        .iter()
        .find(|sheet| sheet.name == "Data")
        .context("Data sheet missing")?;
    assert_eq!(data.values["A1"]["value"], "changed");
    assert_eq!(data.formulas["C1"].trim_start_matches('='), "B1*2");
    assert_eq!(data.comments["A1"], "worker comment");
    assert!(data.merges.contains(&"D1:E1".to_owned()));
    assert!(data.merges.contains(&"J2:K2".to_owned()));
    assert_eq!(data.values["J1"]["value"], "imported");
    assert_eq!(data.formulas["L1"].trim_start_matches('='), "1+1");
    assert_eq!(data.values["H5"]["value"], "changed");
    let calculation = ooxml_patch::read_workbook_calculation_properties(&output)?;
    assert_eq!(calculation.full_calc_on_load, Some(true));
    assert_eq!(calculation.force_full_calc, Some(true));
    Ok(())
}

#[test]
#[ignore = "requires an isolated openpyxl test runtime; run with --ignored"]
fn reports_end_to_end_worker_latency_for_a_fixed_xlsx_sample() -> Result<()> {
    let python = required_python_path();
    let dir = TestDir::new("e2e-performance");
    let input = dir.path("source-5000x10.xlsx");
    make_perf_source(&input, 5_000, 10)?;
    let profile = inspect_xlsx_profile(&input)?;
    let sample_plan = single_value_plan(
        &input,
        &dir.path("capability.xlsx"),
        profile.clone(),
        "capability".into(),
    );
    let backend = OpenpyxlXlsxBackend::new(
        capability(&sample_plan, "Python 3.12.3 / openpyxl 3.1.5"),
        OpenpyxlWorkerConfig::new(python, worker_path())?,
    )?;
    let modifier = ExcelModifier::new(
        "wsl2-x86_64-test",
        vec![std::sync::Arc::new(backend)],
        std::sync::Arc::new(OpenpyxlXlsxValidator),
    )?;
    let mut samples = Vec::new();
    for index in 0..5 {
        let output = dir.path(&format!("result-{index}.xlsx"));
        let edit_plan = single_value_plan(
            &input,
            &output,
            profile.clone(),
            format!("e2e-perf-{index}"),
        );
        let report = modifier.execute(&edit_plan).map_err(anyhow::Error::from)?;
        assert_eq!(report.engine, EngineId::OpenpyxlWorker);
        samples.push(report.elapsed_ms);
    }
    samples.sort_unstable();
    eprintln!(
        "openpyxl end-to-end 5000x10 XLSX execution elapsed_ms={samples:?}, median_ms={}",
        samples[2]
    );
    assert!(samples.iter().all(|elapsed| *elapsed > 0));
    Ok(())
}

#[test]
#[ignore = "requires an isolated openpyxl test runtime; run with --ignored"]
fn worker_timeout_is_terminal_and_does_not_publish_output() -> Result<()> {
    let python = required_python_path();
    let dir = TestDir::new("timeout");
    let input = dir.path("source.xlsx");
    let import_source = dir.path("import-source.xlsx");
    let output = dir.path("result.xlsx");
    let delayed_worker = dir.path("delayed.py");
    make_source(&input)?;
    make_import_source(&import_source)?;
    fs::write(&delayed_worker, "import time\ntime.sleep(2)\n")?;
    let edit_plan = plan(
        &input,
        &output,
        &import_source,
        inspect_xlsx_profile(&input)?,
    );
    let config = OpenpyxlWorkerConfig::new(python, delayed_worker)?
        .with_timeout(Duration::from_millis(30))?;
    let backend = OpenpyxlXlsxBackend::new(capability(&edit_plan, "test"), config)?;
    let modifier = ExcelModifier::new(
        "wsl2-x86_64-test",
        vec![std::sync::Arc::new(backend)],
        std::sync::Arc::new(OpenpyxlXlsxValidator),
    )?;
    let error = modifier.execute(&edit_plan).expect_err("timeout must fail");
    assert_eq!(
        error.kind,
        excel_editor::ExcelModifierErrorKind::BackendFailed
    );
    assert!(!output.exists());
    Ok(())
}

#[test]
#[ignore = "requires an isolated openpyxl test runtime; run with --ignored"]
fn abnormal_worker_exit_is_not_mistaken_for_unsupported_input() -> Result<()> {
    let python = required_python_path();
    let dir = TestDir::new("crash");
    let input = dir.path("source.xlsx");
    let import_source = dir.path("import-source.xlsx");
    let output = dir.path("result.xlsx");
    let crashing_worker = dir.path("crash.py");
    make_source(&input)?;
    make_import_source(&import_source)?;
    fs::write(&crashing_worker, "raise SystemExit(23)\n")?;
    let edit_plan = plan(
        &input,
        &output,
        &import_source,
        inspect_xlsx_profile(&input)?,
    );
    let backend = OpenpyxlXlsxBackend::new(
        capability(&edit_plan, "test"),
        OpenpyxlWorkerConfig::new(python, crashing_worker)?,
    )?;
    let modifier = ExcelModifier::new(
        "wsl2-x86_64-test",
        vec![std::sync::Arc::new(backend)],
        std::sync::Arc::new(OpenpyxlXlsxValidator),
    )?;
    let error = modifier
        .execute(&edit_plan)
        .expect_err("abnormal worker exit must fail");
    assert_eq!(
        error.kind,
        excel_editor::ExcelModifierErrorKind::BackendFailed
    );
    assert!(!output.exists());
    Ok(())
}
