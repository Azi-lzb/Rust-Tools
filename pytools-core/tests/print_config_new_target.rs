use anyhow::{Context, Result};
use excel_editor::{
    BackendCapability, CapabilityEvidence, CapabilityStatus, EditOperation, EngineId, ExcelEditPlan,
};
use excel_reader::read_workbook;
use rust_xlsxwriter::{Color, Format, Formula, Note, Workbook};
use serde_json::Value;
use std::{
    fs,
    path::{Path, PathBuf},
    process::{Command, Output},
    sync::atomic::{AtomicU64, Ordering},
};

static NEXT_TEST_DIR_ID: AtomicU64 = AtomicU64::new(0);

struct TestDir(PathBuf);

impl TestDir {
    fn new() -> Self {
        let serial = NEXT_TEST_DIR_ID.fetch_add(1, Ordering::Relaxed);
        let path = std::env::temp_dir().join(format!(
            "pytools-print-config-{}-{}-{serial}",
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

fn make_source(path: &Path) -> Result<()> {
    make_source_with_formula_cache(path, Some("12.5"))
}

fn make_source_without_formula_cache(path: &Path) -> Result<()> {
    make_source_with_formula_cache(path, None)
}

fn make_source_with_formula_cache(path: &Path, cached_result: Option<&str>) -> Result<()> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Source")?;
    sheet.set_formula_result_default("");
    sheet.write_string_with_format(0, 0, "Report", &Format::new().set_bold())?;
    sheet.merge_range(0, 0, 0, 1, "Report", &Format::new().set_bold())?;
    sheet.write_string(1, 0, "A")?;
    let formula = cached_result.map_or_else(
        || Formula::new("6+6.5"),
        |result| Formula::new("6+6.5").set_result(result),
    );
    sheet.write_formula_with_format(
        1,
        1,
        formula,
        &Format::new().set_background_color(Color::RGB(0x00_99_CC)),
    )?;
    sheet.write_formula(1, 2, Formula::new("1-1").set_result("0"))?;
    sheet.write_formula(1, 3, Formula::new("1/10000").set_result("0.0001"))?;
    sheet.write_string(1, 4, "0")?;
    sheet.write_number(1, 5, 0.0)?;
    sheet.insert_note(1, 0, &Note::new("source note"))?;
    sheet.set_row_height(1, 24.0)?;
    sheet.set_column_width(0, 18.0)?;
    workbook.save(path)?;
    Ok(())
}

fn make_print_config(path: &Path, source: &Path, target: &Path) -> Result<()> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("打印配置")?;
    let headers = [
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
    ];
    for (column, header) in headers.iter().enumerate() {
        sheet.write_string(0, column as u16, *header)?;
    }
    sheet.write_boolean(1, 0, true)?;
    sheet.write_number(1, 1, 1)?;
    sheet.write_string(1, 2, source.to_string_lossy())?;
    sheet.write_string(1, 3, "Source")?;
    sheet.write_string(1, 4, "A1:F2")?;
    sheet.write_string(1, 5, target.to_string_lossy())?;
    sheet.write_string(1, 6, "Printed")?;
    sheet.write_number(1, 7, 1)?;
    sheet.write_number(1, 8, 1)?;
    sheet.write_string(1, 9, "横向")?;
    sheet.write_boolean(1, 10, true)?;
    sheet.write_boolean(1, 11, true)?;
    workbook.save(path)?;
    Ok(())
}

fn success_json(output: &Output) -> Result<Value> {
    anyhow::ensure!(
        output.status.success(),
        "pytools failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    Ok(serde_json::from_slice(&output.stdout)?)
}

fn capability_for(plan: &ExcelEditPlan) -> BackendCapability {
    BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Passed,
        format: plan.format,
        operations: plan.operations.iter().map(EditOperation::kind).collect(),
        supported_features: plan.workbook.features_present.clone(),
        preserved: plan.preserve.clone(),
        platform: "synthetic-pt-3-3-new-target-test".to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: "synthetic-pt-3-3-print-config".to_owned(),
            report_path: "pytools-core/tests/print_config_new_target.rs".to_owned(),
            adapter_version: "umya-spreadsheet-3.1.1".to_owned(),
            engine_version: "umya-spreadsheet-3.1.1".to_owned(),
            platform: "synthetic-pt-3-3-new-target-test".to_owned(),
        }),
    }
}

fn style_fill(
    book: &excel_editor::Workbook,
    sheet_name: &str,
    address: &str,
) -> Result<Option<String>> {
    let sheet = book
        .sheet_by_name(sheet_name)
        .with_context(|| format!("worksheet {sheet_name:?} missing"))?;
    Ok(sheet
        .cell(address)
        .and_then(|cell| cell.style().fill())
        .and_then(|fill| fill.pattern_fill())
        .and_then(|fill| fill.foreground_color())
        .map(|color| color.argb_str().to_ascii_uppercase()))
}

#[test]
fn pt_3_3_rejects_existing_target_before_output_side_effects() -> Result<()> {
    let dir = TestDir::new();
    let source = dir.path("source.xlsx");
    let config = dir.path("config.xlsx");
    let target = dir.path("already-exists.xlsx");
    let output_dir = dir.path("must-not-be-created");
    make_source(&source)?;
    let mut target_book = Workbook::new();
    target_book
        .add_worksheet()
        .write_string(0, 0, "preserve this target")?;
    target_book.save(&target)?;
    make_print_config(&config, &source, &target)?;
    let original_target = fs::read(&target)?;

    let precheck = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["print-config-precheck", "--config"])
        .arg(&config)
        .output()?;
    anyhow::ensure!(!precheck.status.success());
    anyhow::ensure!(
        String::from_utf8_lossy(&precheck.stdout).contains("PT-3-3 only writes to a new file"),
        "precheck did not explain the output contract: {}",
        String::from_utf8_lossy(&precheck.stdout)
    );

    let plan = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["print-config-plan", "--config"])
        .arg(&config)
        .args(["--output-dir"])
        .arg(&output_dir)
        .output()?;
    anyhow::ensure!(!plan.status.success());
    anyhow::ensure!(
        String::from_utf8_lossy(&plan.stderr).contains("PT-3-3 only writes to a new file"),
        "plan builder did not reject the existing target: {}",
        String::from_utf8_lossy(&plan.stderr)
    );
    anyhow::ensure!(
        !output_dir.exists(),
        "rejection created the output directory"
    );
    anyhow::ensure!(
        fs::read(&target)? == original_target,
        "existing target bytes changed after rejection"
    );
    Ok(())
}

#[test]
fn pt_3_3_rejects_formula_without_cached_value_before_output() -> Result<()> {
    let dir = TestDir::new();
    let source = dir.path("source-uncached.xlsx");
    let config = dir.path("config-uncached.xlsx");
    let target = dir.path("nested-output").join("uncached-target.xlsx");
    let output_dir = dir.path("output");
    make_source_without_formula_cache(&source)?;
    make_print_config(&config, &source, &target)?;
    let source_before = fs::read(&source)?;
    let config_before = fs::read(&config)?;

    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["print-config-plan", "--config"])
        .arg(&config)
        .args(["--output-dir"])
        .arg(&output_dir)
        .output()?;
    assert_eq!(output.status.code(), Some(2));
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(
        stderr.contains("formula cached value is blank or unavailable"),
        "{stderr}"
    );
    assert!(!target.exists(), "uncached formula must not create output");
    assert!(
        !output_dir.exists(),
        "validation failure must not publish a plan directory"
    );
    assert_eq!(fs::read(&source)?, source_before);
    assert_eq!(fs::read(&config)?, config_before);
    Ok(())
}

#[test]
fn pt_3_3_new_target_creates_a_plan_without_touching_source_or_target() -> Result<()> {
    let dir = TestDir::new();
    let source = dir.path("source.xlsx");
    let config = dir.path("config.xlsx");
    let target = dir.path("nested-output").join("new-target.xlsx");
    let target_parent = target.parent().unwrap().to_path_buf();
    let output_dir = dir.path("output");
    let capability_path = dir.path("test-capability.json");
    make_source(&source)?;
    make_print_config(&config, &source, &target)?;
    let source_before = fs::read(&source)?;
    let config_before = fs::read(&config)?;

    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["print-config-plan", "--config"])
        .arg(&config)
        .args(["--output-dir"])
        .arg(&output_dir)
        .output()?;
    let response = success_json(&output)?;
    anyhow::ensure!(response.as_array().is_some_and(|items| items.len() == 1));
    assert_eq!(response[0]["execution"], "plan_only");
    let report_output = PathBuf::from(
        response[0]["report"]["output"]
            .as_str()
            .context("print output path missing")?,
    );
    let plan_path = PathBuf::from(
        response[0]["plan_file"]
            .as_str()
            .context("plan path missing")?,
    );
    let plan: ExcelEditPlan = serde_json::from_slice(&fs::read(plan_path)?)?;

    assert_eq!(plan.input, source);
    assert_eq!(plan.output, report_output);
    assert_eq!(plan.output, target);
    assert!(
        !target_parent.exists(),
        "plan-only must not create the requested target directory"
    );
    assert!(
        !target.exists(),
        "plan-only must not create requested target"
    );
    assert!(
        !report_output.exists(),
        "plan-only must not create workbook output"
    );
    assert!(plan
        .operations
        .iter()
        .any(|operation| matches!(operation, EditOperation::ImportRange { .. })));
    assert!(plan
        .operations
        .iter()
        .any(|operation| matches!(operation, EditOperation::SetPageSettings { .. })));
    let print_area = plan
        .operations
        .iter()
        .find_map(|operation| match operation {
            EditOperation::SetPageSettings { settings, .. } => settings.print_area.as_deref(),
            _ => None,
        });
    assert_eq!(print_area, Some("A1:F2"));
    assert_eq!(fs::read(&source)?, source_before);
    assert_eq!(fs::read(&config)?, config_before);

    fs::write(
        &capability_path,
        serde_json::to_vec_pretty(&capability_for(&plan))?,
    )?;
    let executed = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["print-config-plan", "--config"])
        .arg(&config)
        .args(["--output-dir"])
        .arg(&output_dir)
        .args(["--capabilities"])
        .arg(&capability_path)
        .args(["--platform", "synthetic-pt-3-3-new-target-test"])
        .output()?;
    let execution = success_json(&executed)?;
    assert_eq!(execution[0]["execution"]["engine"], "umya_spreadsheet");
    let candidate = PathBuf::from(
        execution[0]["report"]["output"]
            .as_str()
            .context("executed print output path missing")?,
    );
    assert_eq!(candidate, target);
    assert!(
        candidate.is_file(),
        "validated requested target was not created"
    );
    assert!(target_parent.is_dir());

    let printed = read_workbook(&candidate)?;
    assert_eq!(printed.sheets.len(), 1);
    let sheet = &printed.sheets[0];
    assert_eq!(sheet.name, "Printed");
    assert_eq!(sheet.values["A1"]["value"], "Report");
    assert_eq!(sheet.values["A2"]["value"], "A");
    assert_eq!(sheet.values["B2"]["value"], 12.5);
    assert!(
        !sheet.values.contains_key("C2"),
        "cached formula zero must be blank"
    );
    assert_eq!(sheet.values["D2"]["value"], 0.0001);
    assert_eq!(sheet.values["E2"]["value"], "0");
    assert!(
        !sheet.values.contains_key("F2"),
        "numeric zero must be blank"
    );
    assert!(
        sheet.formulas.is_empty(),
        "static print must materialize formula values"
    );
    assert_eq!(sheet.merges, vec!["A1:B1".to_owned()]);
    assert!(
        sheet.comments.is_empty(),
        "static print must remove all comments"
    );
    let output_book = excel_editor::open_xlsx(&candidate)?;
    let output_sheet = output_book
        .sheet_by_name("Printed")
        .context("Printed worksheet missing from Umya readback")?;
    let source_book = excel_editor::open_xlsx(&source)?;
    let source_sheet = source_book
        .sheet_by_name("Source")
        .context("Source worksheet missing from Umya readback")?;
    assert_eq!(
        style_fill(&output_book, "Printed", "B2")?,
        style_fill(&source_book, "Source", "B2")?,
        "print mode 1 must preserve cell fills"
    );
    assert_eq!(
        output_sheet.row_dimension(2).map(|row| row.height()),
        source_sheet.row_dimension(2).map(|row| row.height()),
        "print mode 1 must copy row height"
    );
    assert_eq!(
        output_sheet
            .column_dimension_by_number(1)
            .map(|column| column.width()),
        source_sheet
            .column_dimension_by_number(1)
            .map(|column| column.width()),
        "print mode 1 must copy column width"
    );
    assert_eq!(
        format!("{:?}", output_sheet.page_setup().orientation()),
        "Landscape"
    );
    assert!(output_sheet.page_setup().fit_to_page());
    assert_eq!(output_sheet.page_setup().fit_to_width(), 1);
    assert_eq!(output_sheet.page_setup().fit_to_height(), 1);
    assert_eq!(output_sheet.page_setup().paper_size(), 9);
    assert!(output_sheet.print_options().horizontal_centered());
    assert!(output_sheet.print_options().vertical_centered());
    assert!(output_sheet
        .header_footer()
        .odd_footer()
        .value()
        .contains("第 1 页 / 共 1 页"));
    assert_eq!(fs::read(&source)?, source_before);
    assert_eq!(fs::read(&config)?, config_before);
    Ok(())
}
