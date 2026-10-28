//! Menu 2-3 append edge cases on existing historical targets.
//!
//! Expectations are hand-derived from the documented append contract: the
//! first occurrence of a duplicate key survives, the duplicate key is the
//! tuple of all configured 标识列序号 columns, and planning never writes to
//! the source or the target workbook.

use anyhow::{Context, Result};
use excel_editor::{
    BackendCapability, CapabilityEvidence, CapabilityStatus, EditOperation, EngineId,
    ExcelEditPlan,
};
use excel_reader::{read_workbook, SheetDump};
use pytools_core::dedup_ops::AppendTargetPlan;
use rust_xlsxwriter::Workbook;
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
            "pytools-dedup-append-edges-{}-{:?}-{}",
            std::process::id(),
            std::thread::current().id(),
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

fn make_source(path: &Path, rows: &[(&str, f64)]) -> Result<()> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Source")?;
    for (column, header) in ["ID", "Amount"].iter().enumerate() {
        sheet.write_string(0, column as u16, *header)?;
    }
    for (row, (id, amount)) in rows.iter().enumerate() {
        let row = (row + 1) as u32;
        sheet.write_string(row, 0, *id)?;
        sheet.write_number(row, 1, *amount)?;
    }
    workbook.save(path)?;
    Ok(())
}

fn make_target(path: &Path, rows: &[(&str, f64)]) -> Result<()> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Combined")?;
    for (column, header) in ["ID", "Amount"].iter().enumerate() {
        sheet.write_string(0, column as u16, *header)?;
    }
    for (row, (id, amount)) in rows.iter().enumerate() {
        let row = (row + 1) as u32;
        sheet.write_string(row, 0, *id)?;
        sheet.write_number(row, 1, *amount)?;
    }
    workbook.save(path)?;
    Ok(())
}

fn make_config(path: &Path, source: &Path, target: &Path, key_spec: &str) -> Result<()> {
    let mut workbook = Workbook::new();
    let global = workbook.add_worksheet();
    global.set_name("全局配置")?;
    for (column, value) in ["配置项", "值"].iter().enumerate() {
        global.write_string(0, column as u16, *value)?;
    }
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
    tasks.write_string(1, 0, "是")?;
    tasks.write_string(1, 1, "Edge")?;
    tasks.write_string(1, 2, source.to_string_lossy().as_ref())?;
    tasks.write_string(1, 3, "Source")?;
    tasks.write_string(1, 4, key_spec)?;
    tasks.write_string(1, 5, target.to_string_lossy().as_ref())?;
    tasks.write_string(1, 6, "Combined")?;
    tasks.write_string(1, 7, "追加")?;
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

fn plan_report(config: &Path, output_dir: &Path) -> Result<(Value, PathBuf)> {
    let planned = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-append-plan", "--config"])
        .arg(config)
        .args(["--output-dir"])
        .arg(output_dir)
        .output()?;
    let reports = assert_success(&planned)?;
    anyhow::ensure!(reports.as_array().is_some_and(|items| items.len() == 1));
    let plan_path = PathBuf::from(
        reports[0]["plan_file"]
            .as_str()
            .context("plan path missing from CLI output")?,
    );
    Ok((reports[0].clone(), plan_path))
}

fn capability_for(plan: &ExcelEditPlan) -> BackendCapability {
    BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Passed,
        format: plan.format,
        operations: plan.operations.iter().map(EditOperation::kind).collect(),
        supported_features: plan.workbook.features_present.clone(),
        preserved: plan.preserve.clone(),
        platform: "synthetic-pt-2-5-edges-test".to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: "synthetic-pt-2-5-append-edges".to_owned(),
            report_path: "pytools-core/tests/dedup_append_target_edges.rs".to_owned(),
            adapter_version: "umya-spreadsheet-3.1.1".to_owned(),
            engine_version: "umya-spreadsheet-3.1.1".to_owned(),
            platform: "synthetic-pt-2-5-edges-test".to_owned(),
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

/// A historical workbook that exists but only carries the header row must
/// accept every unique source row. Hand-derived expectation: two source rows
/// with distinct IDs are appended in source order after the header.
#[test]
fn existing_header_only_target_appends_all_unique_rows() -> Result<()> {
    let dir = TestDir::new();
    let source = dir.path("source.xlsx");
    let target = dir.path("target.xlsx");
    let config = dir.path("config.xlsx");
    let output_dir = dir.path("output");
    make_source(&source, &[("A", 10.0), ("B", 20.0)])?;
    make_target(&target, &[])?;
    make_config(&config, &source, &target, "1")?;
    let source_before = fs::read(&source)?;
    let target_before = fs::read(&target)?;

    let (report, plan_path) = plan_report(&config, &output_dir)?;
    assert_eq!(report["target_existed"], true);
    assert_eq!(report["task_ok"], 1);
    assert_eq!(report["task_skipped"], 0);
    assert_eq!(report["appended_rows"], 2);
    assert_eq!(report["duplicate_rows_filtered"], 0);
    assert_eq!(fs::read(&source)?, source_before, "planning touched source");
    assert_eq!(fs::read(&target)?, target_before, "planning touched target");

    let plan: AppendTargetPlan = serde_json::from_slice(&fs::read(&plan_path)?)?;
    let edit_plan = plan.edit_plan.as_ref().context("edit plan missing")?;
    let capability_path = dir.path("capability.json");
    fs::write(
        &capability_path,
        serde_json::to_vec_pretty(&capability_for(edit_plan))?,
    )?;
    let executed = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-append-execute", "--plan"])
        .arg(&plan_path)
        .args(["--capabilities"])
        .arg(&capability_path)
        .args(["--platform", "synthetic-pt-2-5-edges-test"])
        .output()?;
    let execution = assert_success(&executed)?;
    assert_eq!(execution["attempts"][0]["outcome"], "succeeded");

    // The verified candidate carries the appended rows; the historical file
    // itself is still untouched until an explicit target replacement.
    let candidate = read_workbook(PathBuf::from(&plan.candidate))?;
    let combined = sheet(&candidate, "Combined")?;
    assert_eq!(text_at(combined, "A2"), Some("A"));
    assert_eq!(number_at(combined, "B2"), Some(10.0));
    assert_eq!(text_at(combined, "A3"), Some("B"));
    assert_eq!(number_at(combined, "B3"), Some(20.0));
    assert_eq!(fs::read(&source)?, source_before);
    assert_eq!(fs::read(&target)?, target_before);
    Ok(())
}

/// A historical workbook that already contains every source record must yield
/// an empty append plan instead of duplicating rows.
#[test]
fn target_with_all_records_appends_nothing() -> Result<()> {
    let dir = TestDir::new();
    let source = dir.path("source.xlsx");
    let target = dir.path("target.xlsx");
    let config = dir.path("config.xlsx");
    let output_dir = dir.path("output");
    make_source(&source, &[("A", 10.0), ("B", 20.0)])?;
    make_target(&target, &[("A", 10.0), ("B", 20.0)])?;
    make_config(&config, &source, &target, "1")?;
    let source_before = fs::read(&source)?;
    let target_before = fs::read(&target)?;

    let (report, _plan_path) = plan_report(&config, &output_dir)?;
    assert_eq!(report["target_existed"], true);
    assert_eq!(report["task_ok"], 1);
    assert_eq!(report["task_skipped"], 0);
    assert_eq!(report["appended_rows"], 0);
    assert_eq!(report["duplicate_rows_filtered"], 2);
    assert_eq!(fs::read(&source)?, source_before, "planning touched source");
    assert_eq!(fs::read(&target)?, target_before, "planning touched target");
    Ok(())
}

/// With the composite key 标识列序号=1;2 two rows only collide when every
/// key column matches: a repeated (ID, Amount) pair is filtered after its
/// first occurrence, while a different institution with the same amount is
/// kept as a distinct record.
#[test]
fn composite_key_keeps_first_occurrence_and_distinct_institutions() -> Result<()> {
    let dir = TestDir::new();
    let source = dir.path("source.xlsx");
    let target = dir.path("target.xlsx");
    let config = dir.path("config.xlsx");
    let output_dir = dir.path("output");
    make_source(&source, &[("A", 10.0), ("A", 10.0), ("B", 10.0)])?;
    make_config(&config, &source, &target, "1;2")?;
    let source_before = fs::read(&source)?;

    let (report, _plan_path) = plan_report(&config, &output_dir)?;
    assert_eq!(report["target_existed"], false);
    assert_eq!(report["task_ok"], 1);
    assert_eq!(report["appended_rows"], 2);
    assert_eq!(report["duplicate_rows_filtered"], 1);
    assert_eq!(fs::read(&source)?, source_before, "planning touched source");
    Ok(())
}
