use anyhow::{Context, Result};
use excel_editor::{
    BackendCapability, CapabilityEvidence, CapabilityStatus, EditOperation, EngineId,
    ExcelEditPlan, PreservationRequirement,
};
use excel_reader::{read_workbook, SheetDump};
use pytools_core::dedup_ops::AppendTargetPlan;
use rust_xlsxwriter::{Color, Format, Note, Workbook};
use serde_json::Value;
use std::{
    fs,
    path::{Path, PathBuf},
    process::{Command, Output},
};

struct TestDir(PathBuf);

impl TestDir {
    fn new() -> Self {
        let path = std::env::temp_dir().join(format!(
            "pytools-dedup-append-{}-{}",
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
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Source")?;
    for (column, header) in ["ID", "Amount", "Derived"].iter().enumerate() {
        sheet.write_string(0, column as u16, *header)?;
    }

    sheet.write_string(1, 0, "A")?;
    sheet.write_number(1, 1, 10.0)?;
    sheet.write_formula(1, 2, "=B2*2")?;

    sheet.write_string(2, 0, "A")?;
    sheet.write_number(2, 1, 20.0)?;
    sheet.write_formula(2, 2, "=B3*2")?;

    sheet.write_string(3, 0, "B")?;
    sheet.write_number(3, 1, 10.0)?;
    sheet.write_formula(3, 2, "=B4*2")?;

    // Rows with a blank key are retained by the Python contract and are not
    // inserted into the de-duplication key set.
    sheet.write_number(4, 1, 30.0)?;
    sheet.write_formula(4, 2, "=B5*2")?;
    workbook.save(path)?;
    Ok(())
}

fn make_existing_target(path: &Path) -> Result<()> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Combined")?;
    for (column, header) in ["ID", "Amount", "Derived"].iter().enumerate() {
        sheet.write_string(0, column as u16, *header)?;
    }
    sheet.write_string(1, 0, "Z")?;
    sheet.write_number_with_format(
        1,
        1,
        99.0,
        &Format::new().set_background_color(Color::RGB(0x00_99_CC)),
    )?;
    sheet.write_formula(1, 2, "=B2*2")?;

    let untouched = workbook.add_worksheet();
    untouched.set_name("Untouched")?;
    untouched.write_string(0, 0, "keep-me")?;
    untouched.insert_note(0, 0, &Note::new("retained-note"))?;
    untouched.merge_range(0, 2, 0, 3, "retained-merge", &Format::new())?;
    untouched.set_row_height(0, 24.0)?;
    workbook.save(path)?;
    Ok(())
}

fn make_header_only_target(path: &Path, headers: &[&str]) -> Result<()> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Combined")?;
    for (column, header) in headers.iter().enumerate() {
        if !header.is_empty() {
            sheet.write_string(0, column as u16, *header)?;
        }
    }
    workbook.save(path)?;
    Ok(())
}

fn style_fill(
    book: &excel_editor::Workbook,
    sheet_name: &str,
    address: &str,
) -> Result<Option<String>> {
    let sheet = book.sheet_by_name(sheet_name)?;
    Ok(sheet
        .get_cell(address)
        .and_then(|cell| cell.style().fill())
        .and_then(|fill| fill.pattern_fill())
        .and_then(|fill| fill.foreground_color())
        .map(|color| color.argb_str().to_ascii_uppercase()))
}

fn make_config(path: &Path, source: &Path, target: &Path) -> Result<()> {
    let mut workbook = Workbook::new();
    let global = workbook.add_worksheet();
    global.set_name("全局配置")?;
    for (column, value) in ["配置项", "值"].iter().enumerate() {
        global.write_string(0, column as u16, *value)?;
    }
    // absolute per-test output: parallel CLI tests must not share the
    // crate-root relative "output" directory
    let output_value = path
        .parent()
        .context("configuration parent missing")?
        .join("output")
        .to_string_lossy()
        .into_owned();
    for (row, (key, value)) in [
        ("输入目录", "."),
        ("输出目录", output_value.as_str()),
        ("日志目录", "logs"),
    ]
    .iter()
    .enumerate()
    {
        global.write_string((row + 1) as u32, 0, *key)?;
        global.write_string((row + 1) as u32, 1, *value)?;
    }

    let tasks = workbook.add_worksheet();
    tasks.set_name("去重追加数据配置")?;
    let headers = [
        "是否启用",
        "任务名",
        "源数据工作簿",
        "源数据工作表",
        "标识列序号",
        "目标工作簿",
        "目标工作表",
        "执行模式",
    ];
    for (column, header) in headers.iter().enumerate() {
        tasks.write_string(0, column as u16, *header)?;
    }
    for (row, (name, key_column)) in [("By-ID", "1"), ("By-amount", "2")].iter().enumerate() {
        let row = (row + 1) as u32;
        tasks.write_string(row, 0, "是")?;
        tasks.write_string(row, 1, *name)?;
        tasks.write_string(row, 2, source.to_string_lossy().as_ref())?;
        tasks.write_string(row, 3, "Source")?;
        tasks.write_string(row, 4, *key_column)?;
        tasks.write_string(row, 5, target.to_string_lossy().as_ref())?;
        tasks.write_string(row, 6, "Combined")?;
        tasks.write_string(row, 7, "追加")?;
    }
    workbook.save(path)?;
    Ok(())
}

fn assert_success(output: &Output) -> Result<Value> {
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
        platform: "synthetic-pt-2-5-umya-test".to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: "synthetic-pt-2-5-existing-target".to_owned(),
            report_path: "pytools-core/tests/dedup_append_execution.rs".to_owned(),
            adapter_version: "umya-spreadsheet-3.1.1".to_owned(),
            engine_version: "umya-spreadsheet-3.1.1".to_owned(),
            platform: "synthetic-pt-2-5-umya-test".to_owned(),
        }),
    }
}

fn sheet<'a>(book: &'a excel_reader::WorkbookDump, name: &str) -> Result<&'a SheetDump> {
    book.sheets
        .iter()
        .find(|candidate| candidate.name == name)
        .with_context(|| format!("worksheet {name:?} missing"))
}

fn text_at<'a>(sheet: &'a SheetDump, address: &str) -> Option<&'a str> {
    sheet
        .values
        .get(address)
        .and_then(|value| value.get("value"))
        .and_then(Value::as_str)
}

fn number_at(sheet: &SheetDump, address: &str) -> Option<f64> {
    sheet
        .values
        .get(address)
        .and_then(|value| value.get("value"))
        .and_then(Value::as_f64)
}

#[test]
fn pt_2_5_absent_target_appends_unique_rows_through_the_rust_cli() -> Result<()> {
    let dir = TestDir::new();
    let source = dir.path("source.xlsx");
    let target = dir.path("target.xlsx");
    let config = dir.path("config.xlsx");
    let output_dir = dir.path("output");
    make_source(&source)?;
    make_config(&config, &source, &target)?;

    let source_before = fs::read(&source)?;
    let config_before = fs::read(&config)?;
    let planned = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-append-plan", "--config"])
        .arg(&config)
        .args(["--output-dir"])
        .arg(&output_dir)
        .output()?;
    let reports = assert_success(&planned)?;
    anyhow::ensure!(reports.as_array().is_some_and(|items| items.len() == 1));
    assert_eq!(reports[0]["target_existed"], false);
    assert_eq!(reports[0]["task_ok"], 2);
    assert_eq!(reports[0]["task_skipped"], 0);
    assert_eq!(reports[0]["appended_rows"], 4);
    assert_eq!(reports[0]["duplicate_rows_filtered"], 4);
    assert_eq!(reports[0]["merged_tasks"], 0);
    assert!(
        !target.exists(),
        "planning must not create or alter the target"
    );
    assert_eq!(fs::read(&source)?, source_before);
    assert_eq!(fs::read(&config)?, config_before);

    let plan_path = PathBuf::from(
        reports[0]["plan_file"]
            .as_str()
            .context("plan path missing from CLI output")?,
    );
    let plan: AppendTargetPlan = serde_json::from_slice(&fs::read(&plan_path)?)?;
    assert!(!plan.target_existed);
    assert_eq!(plan.appended_rows, 4);
    assert_eq!(plan.duplicate_rows_filtered, 4);
    assert!(plan.edit_plan.is_none());
    assert_eq!(plan.new_workbook.as_ref().unwrap().sheets.len(), 1);

    let executed = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-append-execute", "--plan"])
        .arg(&plan_path)
        .output()?;
    let execution = assert_success(&executed)?;
    assert_eq!(execution["status"], "created_and_reopened");
    let candidate = PathBuf::from(
        execution["candidate"]
            .as_str()
            .context("candidate path missing from CLI output")?,
    );
    assert!(candidate.is_file());
    assert!(
        !target.exists(),
        "candidate generation must leave target absent"
    );
    assert_eq!(fs::read(&source)?, source_before);
    assert_eq!(fs::read(&config)?, config_before);

    let output = read_workbook(&candidate)?;
    assert_eq!(output.sheets.len(), 1);
    let combined = sheet(&output, "Combined")?;
    assert_eq!(text_at(combined, "A1"), Some("ID"));
    assert_eq!(text_at(combined, "B1"), Some("Amount"));
    assert_eq!(text_at(combined, "C1"), Some("Derived"));
    assert_eq!(text_at(combined, "A2"), Some("A"));
    assert_eq!(number_at(combined, "B2"), Some(10.0));
    assert_eq!(text_at(combined, "A3"), Some("B"));
    assert_eq!(number_at(combined, "B3"), Some(10.0));
    assert_eq!(number_at(combined, "B4"), Some(30.0));
    assert_eq!(text_at(combined, "A5"), Some("A"));
    assert_eq!(number_at(combined, "B5"), Some(20.0));
    assert_eq!(
        combined.formulas.get("C2").map(String::as_str),
        Some("B2*2")
    );
    assert_eq!(
        combined.formulas.get("C3").map(String::as_str),
        Some("B4*2")
    );
    assert_eq!(
        combined.formulas.get("C4").map(String::as_str),
        Some("B5*2")
    );
    assert_eq!(
        combined.formulas.get("C5").map(String::as_str),
        Some("B3*2")
    );

    Ok(())
}

#[test]
fn pt_2_5_skips_existing_targets_when_header_names_order_or_width_differ() -> Result<()> {
    let dir = TestDir::new();
    let source = dir.path("source.xlsx");
    make_source(&source)?;
    let source_before = fs::read(&source)?;

    for (case, headers) in [
        (
            "extra-target-column",
            vec!["ID", "Amount", "Derived", "Extra"],
        ),
        (
            "leading-header-whitespace",
            vec![" ID", "Amount", "Derived"],
        ),
        ("reordered-target-columns", vec!["Amount", "ID", "Derived"]),
        (
            "blank-first-but-not-empty-header",
            vec!["", "ExistingHeader"],
        ),
    ] {
        let case_dir = dir.path(case);
        fs::create_dir_all(&case_dir)?;
        let target = case_dir.join("target.xlsx");
        let config = case_dir.join("config.xlsx");
        let output_dir = case_dir.join("output");
        make_header_only_target(&target, &headers)?;
        make_config(&config, &source, &target)?;
        let target_before = fs::read(&target)?;

        let planned = Command::new(env!("CARGO_BIN_EXE_pytools"))
            .args(["dedup-append-plan", "--config"])
            .arg(&config)
            .args(["--output-dir"])
            .arg(&output_dir)
            .output()?;
        let reports = assert_success(&planned)?;
        assert_eq!(reports.as_array().map(Vec::len), Some(1));
        assert_eq!(reports[0]["target_existed"], true);
        assert_eq!(reports[0]["task_ok"], 0);
        assert_eq!(reports[0]["task_skipped"], 2);
        assert_eq!(reports[0]["appended_rows"], 0);
        assert!(reports[0]["issues"].as_array().is_some_and(|issues| {
            issues.len() == 2
                && issues.iter().all(|issue| {
                    issue
                        .as_str()
                        .is_some_and(|text| text.contains("header_mismatch"))
                })
        }));
        assert_eq!(
            fs::read(&target)?,
            target_before,
            "{case} target was modified"
        );
        assert_eq!(fs::read(&source)?, source_before);
    }

    Ok(())
}

#[test]
fn pt_2_5_existing_target_edits_a_candidate_without_overwriting_the_target() -> Result<()> {
    let dir = TestDir::new();
    let source = dir.path("source.xlsx");
    let target = dir.path("target.xlsx");
    let config = dir.path("config.xlsx");
    let output_dir = dir.path("output");
    let capability_path = dir.path("test-capability.json");
    make_source(&source)?;
    make_existing_target(&target)?;
    make_config(&config, &source, &target)?;

    let source_before = fs::read(&source)?;
    let target_before = fs::read(&target)?;
    let config_before = fs::read(&config)?;
    let planned = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-append-plan", "--config"])
        .arg(&config)
        .args(["--output-dir"])
        .arg(&output_dir)
        .output()?;
    let reports = assert_success(&planned)?;
    assert_eq!(reports[0]["target_existed"], true);
    assert_eq!(reports[0]["task_ok"], 2);
    assert_eq!(reports[0]["task_skipped"], 0);
    assert_eq!(reports[0]["appended_rows"], 4);
    assert_eq!(reports[0]["duplicate_rows_filtered"], 4);
    assert_eq!(fs::read(&source)?, source_before);
    assert_eq!(fs::read(&target)?, target_before);

    let plan_path = PathBuf::from(
        reports[0]["plan_file"]
            .as_str()
            .context("plan path missing from CLI output")?,
    );
    let plan: AppendTargetPlan = serde_json::from_slice(&fs::read(&plan_path)?)?;
    let edit_plan = plan.edit_plan.as_ref().context("edit plan missing")?;
    assert!(plan.target_existed);
    assert!(plan.new_workbook.is_none());
    assert_eq!(edit_plan.input, target);
    assert!(!edit_plan
        .preserve
        .contains(&PreservationRequirement::UnmodifiedPackageParts));
    fs::write(
        &capability_path,
        serde_json::to_vec_pretty(&capability_for(edit_plan))?,
    )?;

    let executed = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-append-execute", "--plan"])
        .arg(&plan_path)
        .args(["--capabilities"])
        .arg(&capability_path)
        .args(["--platform", "synthetic-pt-2-5-umya-test"])
        .output()?;
    let execution = assert_success(&executed)?;
    assert_eq!(execution["engine"], "umya_spreadsheet");
    assert_eq!(execution["attempts"][0]["outcome"], "succeeded");
    let candidate = PathBuf::from(&plan.candidate);
    assert!(candidate.is_file());
    assert_eq!(fs::read(&source)?, source_before);
    assert_eq!(fs::read(&target)?, target_before);
    assert_eq!(fs::read(&config)?, config_before);

    let output = read_workbook(&candidate)?;
    assert_eq!(
        output
            .sheets
            .iter()
            .map(|sheet| sheet.name.as_str())
            .collect::<Vec<_>>(),
        ["Combined", "Untouched"]
    );
    let combined = sheet(&output, "Combined")?;
    let untouched = sheet(&output, "Untouched")?;
    assert_eq!(text_at(combined, "A2"), Some("Z"));
    assert_eq!(number_at(combined, "B2"), Some(99.0));
    assert_eq!(text_at(combined, "A3"), Some("A"));
    assert_eq!(number_at(combined, "B3"), Some(10.0));
    assert_eq!(text_at(combined, "A4"), Some("B"));
    assert_eq!(number_at(combined, "B4"), Some(10.0));
    assert_eq!(number_at(combined, "B5"), Some(30.0));
    assert_eq!(text_at(combined, "A6"), Some("A"));
    assert_eq!(number_at(combined, "B6"), Some(20.0));
    assert_eq!(
        combined.formulas.get("C2").map(String::as_str),
        Some("B2*2")
    );
    assert_eq!(
        combined.formulas.get("C3").map(String::as_str),
        Some("B2*2")
    );
    assert_eq!(
        combined.formulas.get("C4").map(String::as_str),
        Some("B4*2")
    );
    assert_eq!(
        combined.formulas.get("C5").map(String::as_str),
        Some("B5*2")
    );
    assert_eq!(
        combined.formulas.get("C6").map(String::as_str),
        Some("B3*2")
    );
    assert_eq!(text_at(untouched, "A1"), Some("keep-me"));
    assert!(untouched
        .comments
        .get("A1")
        .is_some_and(|comment| comment.contains("retained-note")));
    assert_eq!(untouched.merges, ["C1:D1"]);

    let target_umya = excel_editor::open_xlsx(&target)?;
    let output_umya = excel_editor::open_xlsx(&candidate)?;
    assert_eq!(
        style_fill(&output_umya, "Combined", "B2")?,
        style_fill(&target_umya, "Combined", "B2")?,
        "existing target cell style must remain unchanged"
    );
    let output_untouched = output_umya
        .sheet_by_name("Untouched")
        .context("Untouched sheet missing from Umya readback")?;
    let target_untouched = target_umya
        .sheet_by_name("Untouched")
        .context("Untouched target sheet missing from Umya readback")?;
    assert_eq!(
        output_untouched.row_dimension(1).map(|row| row.height()),
        target_untouched.row_dimension(1).map(|row| row.height()),
        "untouched row height must remain unchanged"
    );

    Ok(())
}

#[test]
fn pt_1_3_pt_1_4_explicitly_commit_existing_target_after_capability_validation() -> Result<()> {
    let dir = TestDir::new();
    let source = dir.path("source.xlsx");
    let target = dir.path("target.xlsx");
    let config = dir.path("config.xlsx");
    let output_dir = dir.path("output");
    let capability_path = dir.path("test-capability.json");
    make_source(&source)?;
    make_existing_target(&target)?;
    make_config(&config, &source, &target)?;

    let source_before = fs::read(&source)?;
    let target_before = fs::read(&target)?;
    let target_hash = pytools_core::file_commit::sha256_file(&target)?;
    let config_before = fs::read(&config)?;
    let planned = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-append-plan", "--config"])
        .arg(&config)
        .args(["--output-dir"])
        .arg(&output_dir)
        .output()?;
    let reports = assert_success(&planned)?;
    let plan_path = PathBuf::from(
        reports[0]["plan_file"]
            .as_str()
            .context("plan path missing from CLI output")?,
    );
    let plan: AppendTargetPlan = serde_json::from_slice(&fs::read(&plan_path)?)?;
    assert_eq!(plan.target_sha256.as_deref(), Some(target_hash.as_str()));
    let edit_plan = plan.edit_plan.as_ref().context("edit plan missing")?;
    fs::write(
        &capability_path,
        serde_json::to_vec_pretty(&capability_for(edit_plan))?,
    )?;

    let executed = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-append-execute", "--plan"])
        .arg(&plan_path)
        .args(["--capabilities"])
        .arg(&capability_path)
        .args([
            "--platform",
            "synthetic-pt-2-5-umya-test",
            "--commit-target",
        ])
        .output()?;
    let execution = assert_success(&executed)?;
    assert_eq!(execution["engine"], "umya_spreadsheet");
    assert_eq!(execution["target_commit"]["operation"], "replaced");
    assert_eq!(execution["target_commit"]["sha256_before"], target_hash);
    assert_eq!(execution["execution_summary"]["status"], "succeeded");
    assert_eq!(
        execution["execution_summary"]["stages"][0]["engine"],
        "umya_spreadsheet"
    );
    assert_eq!(execution["execution_summary"]["task_success_count"], 2);
    let candidate = PathBuf::from(&plan.candidate);
    assert!(
        candidate.is_file(),
        "candidate remains as an audit artifact"
    );
    assert_eq!(fs::read(&target)?, fs::read(&candidate)?);
    assert_eq!(
        execution["target_commit"]["sha256_after"],
        pytools_core::file_commit::sha256_file(&target)?
    );
    assert!(execution["target_commit"]["elapsed_ms"].as_u64().is_some());
    assert_eq!(fs::read(&source)?, source_before);
    assert_eq!(fs::read(&config)?, config_before);
    assert_ne!(fs::read(&target)?, target_before);

    let output = read_workbook(&target)?;
    assert_eq!(
        output
            .sheets
            .iter()
            .map(|sheet| sheet.name.as_str())
            .collect::<Vec<_>>(),
        ["Combined", "Untouched"]
    );
    let combined = sheet(&output, "Combined")?;
    assert_eq!(text_at(combined, "A2"), Some("Z"));
    assert_eq!(text_at(combined, "A3"), Some("A"));
    assert_eq!(number_at(combined, "B3"), Some(10.0));
    assert_eq!(
        combined.formulas.get("C3").map(String::as_str),
        Some("B2*2")
    );
    let untouched = sheet(&output, "Untouched")?;
    assert_eq!(text_at(untouched, "A1"), Some("keep-me"));
    assert!(untouched
        .comments
        .get("A1")
        .is_some_and(|comment| comment.contains("retained-note")));
    assert_eq!(untouched.merges, ["C1:D1"]);
    Ok(())
}

#[test]
fn pt_1_3_pt_1_4_refuse_commit_if_target_changed_after_planning() -> Result<()> {
    let dir = TestDir::new();
    let source = dir.path("source.xlsx");
    let target = dir.path("target.xlsx");
    let config = dir.path("config.xlsx");
    let output_dir = dir.path("output");
    let capability_path = dir.path("test-capability.json");
    make_source(&source)?;
    make_existing_target(&target)?;
    make_config(&config, &source, &target)?;

    let planned = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-append-plan", "--config"])
        .arg(&config)
        .args(["--output-dir"])
        .arg(&output_dir)
        .output()?;
    let reports = assert_success(&planned)?;
    let plan_path = PathBuf::from(
        reports[0]["plan_file"]
            .as_str()
            .context("plan path missing from CLI output")?,
    );
    let plan: AppendTargetPlan = serde_json::from_slice(&fs::read(&plan_path)?)?;
    let edit_plan = plan.edit_plan.as_ref().context("edit plan missing")?;
    fs::write(
        &capability_path,
        serde_json::to_vec_pretty(&capability_for(edit_plan))?,
    )?;

    // Simulate another program/user changing the configured workbook after the
    // plan was issued. The commit flag must reject it before running the engine.
    fs::write(&target, b"changed after planning")?;
    let changed_target = fs::read(&target)?;
    let executed = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-append-execute", "--plan"])
        .arg(&plan_path)
        .args(["--capabilities"])
        .arg(&capability_path)
        .args([
            "--platform",
            "synthetic-pt-2-5-umya-test",
            "--commit-target",
        ])
        .output()?;
    anyhow::ensure!(
        !executed.status.success(),
        "commit unexpectedly succeeded for a changed target"
    );
    anyhow::ensure!(
        String::from_utf8_lossy(&executed.stderr).contains("changed after planning"),
        "unexpected commit rejection: {}",
        String::from_utf8_lossy(&executed.stderr)
    );
    assert_eq!(fs::read(&target)?, changed_target);
    assert!(
        !plan.candidate.exists(),
        "stale plans are rejected before editing"
    );
    Ok(())
}

#[test]
fn pt_2_5_explicitly_commit_new_target_and_return_commit_receipt() -> Result<()> {
    let dir = TestDir::new();
    let source = dir.path("source.xlsx");
    let target = dir.path("target.xlsx");
    let config = dir.path("config.xlsx");
    let output_dir = dir.path("output");
    make_source(&source)?;
    make_config(&config, &source, &target)?;

    let planned = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-append-plan", "--config"])
        .arg(&config)
        .args(["--output-dir"])
        .arg(&output_dir)
        .output()?;
    let reports = assert_success(&planned)?;
    let plan_path = PathBuf::from(
        reports[0]["plan_file"]
            .as_str()
            .context("plan path missing from CLI output")?,
    );
    let plan: AppendTargetPlan = serde_json::from_slice(&fs::read(&plan_path)?)?;
    assert!(plan.target_sha256.is_none());
    let executed = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-append-execute", "--plan"])
        .arg(&plan_path)
        .arg("--commit-target")
        .output()?;
    let execution = assert_success(&executed)?;
    assert_eq!(execution["target_commit"]["operation"], "created");
    assert_eq!(
        execution["execution_summary"]["stages"][0]["engine"],
        "rust_xlsxwriter"
    );
    let candidate = PathBuf::from(
        execution["candidate"]
            .as_str()
            .context("candidate path missing")?,
    );
    assert!(target.is_file());
    assert_eq!(fs::read(&target)?, fs::read(&candidate)?);
    assert_eq!(read_workbook(&target)?.sheets.len(), 1);
    Ok(())
}

#[test]
fn pt_2_5_refuses_to_overwrite_target_created_after_plan() -> Result<()> {
    let dir = TestDir::new();
    let source = dir.path("source.xlsx");
    let target = dir.path("target.xlsx");
    let config = dir.path("config.xlsx");
    let output_dir = dir.path("output");
    make_source(&source)?;
    make_config(&config, &source, &target)?;

    let planned = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-append-plan", "--config"])
        .arg(&config)
        .args(["--output-dir"])
        .arg(&output_dir)
        .output()?;
    let reports = assert_success(&planned)?;
    let plan_path = PathBuf::from(
        reports[0]["plan_file"]
            .as_str()
            .context("plan path missing from CLI output")?,
    );
    let plan: AppendTargetPlan = serde_json::from_slice(&fs::read(&plan_path)?)?;
    assert!(!plan.target_existed);

    // Simulate another writer claiming the absent destination after planning.
    fs::write(&target, b"another writer owns this path")?;
    let competing_bytes = fs::read(&target)?;
    let executed = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-append-execute", "--plan"])
        .arg(&plan_path)
        .arg("--commit-target")
        .output()?;
    anyhow::ensure!(
        !executed.status.success(),
        "commit unexpectedly replaced a target created after planning"
    );
    anyhow::ensure!(
        String::from_utf8_lossy(&executed.stderr).contains("appeared after planning"),
        "unexpected collision error: {}",
        String::from_utf8_lossy(&executed.stderr)
    );
    assert_eq!(fs::read(&target)?, competing_bytes);
    assert!(
        !plan.candidate.exists(),
        "candidate generation must stop when the target path is already claimed"
    );
    Ok(())
}
