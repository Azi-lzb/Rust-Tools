use anyhow::{Context, Result};
use rust_xlsxwriter::Workbook;
use serde_json::{json, Value};
use std::{
    fs,
    path::{Path, PathBuf},
    process::{Command, Output},
    time::{SystemTime, UNIX_EPOCH},
};

struct TempDirectory(PathBuf);

impl TempDirectory {
    fn new(name: &str) -> Result<Self> {
        let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let path = std::env::temp_dir().join(format!("{name}-{}-{nonce}", std::process::id()));
        fs::create_dir_all(&path)?;
        Ok(Self(path))
    }

    fn path(&self) -> &Path {
        &self.0
    }
}

impl Drop for TempDirectory {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

#[test]
fn batch_plan_cli_matches_python_baseline_and_filters_user_view() -> Result<()> {
    let root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let fixture = root.join("tests/fixtures/institution_normalize.xlsx");
    let baseline_path = root.join("probe/reports/institution_normalize_baseline.json");
    let baseline: Value = serde_json::from_slice(&fs::read(&baseline_path)?)?;
    let temp = TempDirectory::new("pytools-pt-7-2-batch-cli")?;
    let config = temp.path().join("institution-config.xlsx");
    write_institution_config(&config)?;
    let first = temp.path().join("first.xlsx");
    let unsupported = temp.path().join("unsupported.xls");
    let last = temp.path().join("last.xlsx");
    fs::copy(&fixture, &first)?;
    fs::write(
        &unsupported,
        b"not an xls workbook; extension is intentionally unsupported",
    )?;
    fs::copy(&fixture, &last)?;
    let first_hash = pytools_core::file_commit::sha256_file(&first)?;
    let last_hash = pytools_core::file_commit::sha256_file(&last)?;
    let plans = temp.path().join("plans");
    let report_path = temp.path().join("batch-run.json");

    let output = run_cli([
        "institution-normalize",
        "--config",
        config.to_str().context("temporary path is UTF-8")?,
        "--plan-dir",
        plans.to_str().context("temporary path is UTF-8")?,
        "--report-output",
        report_path.to_str().context("temporary path is UTF-8")?,
        "--view",
        "user",
        first.to_str().context("temporary path is UTF-8")?,
        unsupported.to_str().context("temporary path is UTF-8")?,
        last.to_str().context("temporary path is UTF-8")?,
    ])?;
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let user_view: Value = serde_json::from_slice(&output.stdout)?;
    assert_eq!(user_view["status"], "部分文件失败");
    assert_eq!(user_view["summary"]["total_inputs"], 3);
    assert_eq!(user_view["summary"]["completed_files"], 2);
    assert_eq!(user_view["summary"]["failed_files"], 1);
    assert_eq!(user_view["summary"]["mapped_rows"], 6);
    assert_eq!(user_view["files"][0]["mapped_rows"], 3);
    assert_eq!(user_view["files"][0]["plan_saved"], true);
    assert_eq!(user_view["files"][1]["error_stage"], "文件格式检查");
    assert_eq!(user_view["files"][2]["mapped_rows"], 3);
    assert!(!user_view.to_string().contains("calamine"));
    assert!(!user_view.to_string().contains("umya_spreadsheet"));
    assert!(!user_view.to_string().contains("profile_inspection"));
    assert!(!user_view.to_string().contains("plan_build"));
    assert!(user_view["events"].as_array().unwrap().iter().any(|event| {
        event["phase"] == "工作簿结构检查" && event["status"] == "已完成"
    }));
    let live_user_messages = String::from_utf8_lossy(&output.stderr);
    assert!(live_user_messages.contains("正在处理 1/3"));
    assert!(live_user_messages.contains("原文件未修改"));

    let stored_report: Value = serde_json::from_slice(&fs::read(&report_path)?)?;
    assert_eq!(stored_report["product"], "pytools");
    assert_eq!(stored_report["feature_id"], "PT-7-2");
    let events = stored_report["events"].as_array().unwrap();
    assert!(events.len() >= 21, "expected per-stage progress events");
    for (index, event) in events.iter().enumerate() {
        assert_eq!(event["sequence"].as_u64(), Some(index as u64 + 1));
    }
    assert!(events
        .iter()
        .any(|event| { event["phase"] == "profile_inspection" && event["status"] == "completed" }));
    assert!(events
        .iter()
        .any(|event| { event["phase"] == "plan_build" && event["status"] == "completed" }));
    assert_eq!(
        user_view["events"].as_array().unwrap().last().unwrap()["status"],
        "部分文件失败"
    );
    assert_eq!(stored_report["events"][0]["components"][0], "Rust");
    assert_eq!(stored_report["files"][0]["actual_engine"], Value::Null);
    assert_eq!(stored_report["files"][0]["operation_count"], 3);
    assert_eq!(
        stored_report["files"][1]["error_type"],
        "unsupported_format"
    );
    assert!(stored_report.to_string().contains("calamine"));
    assert!(!stored_report.to_string().contains("Exact Standard"));

    let plan_files = fs::read_dir(&plans)?.collect::<std::io::Result<Vec<_>>>()?;
    assert_eq!(plan_files.len(), 2);
    for plan_file in plan_files {
        let plan: Value = serde_json::from_slice(&fs::read(plan_file.path())?)?;
        assert_eq!(plan["operations"].as_array().unwrap().len(), 3);
        assert_eq!(
            plan["request_id"].as_str().unwrap().split('-').next(),
            Some("pytools")
        );
    }
    assert_eq!(pytools_core::file_commit::sha256_file(&first)?, first_hash);
    assert_eq!(pytools_core::file_commit::sha256_file(&last)?, last_hash);
    assert_eq!(
        baseline["stats"]["mapped_rows"].as_u64().unwrap() as usize,
        user_view["files"][0]["mapped_rows"].as_u64().unwrap() as usize
    );

    if let Some(path) = std::env::var_os("PYTOOLS_INSTITUTION_BATCH_REPORT") {
        let report = json!({
            "schema_version": 1,
            "suite": "pytools.institution_normalize_batch_cli",
            "status": "candidate_passed",
            "baseline": "probe/reports/institution_normalize_baseline.json",
            "fixture": "tests/fixtures/institution_normalize.xlsx (synthetic)",
            "python_baseline_stats": baseline["stats"],
            "rust_user_view": user_view,
            "rust_developer_history": stored_report,
            "security_checks": {
                "input_workbooks_unchanged_in_plan_mode": true,
                "plan_report_contains_no_cell_values": true,
                "user_view_hides_engine_details": true
            },
            "limitations": [
                "synthetic XLSX fixture only; no authorized real business sample",
                "tested planning and capability rejection, not a production engine execution",
                "PT-7-2 remains unverified and no production capability was created",
                "Windows native Rust and target Linux/deb are not verified"
            ]
        });
        fs::write(path, serde_json::to_vec_pretty(&report)?)?;
    }
    Ok(())
}

#[test]
fn apply_without_a_passed_capability_refuses_to_modify_source() -> Result<()> {
    let root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let fixture = root.join("tests/fixtures/institution_normalize.xlsx");
    let temp = TempDirectory::new("pytools-pt-7-2-capability-gate")?;
    let config = temp.path().join("institution-config.xlsx");
    let source = temp.path().join("source.xlsx");
    let capability_file = temp.path().join("empty-capabilities.json");
    let report_path = temp.path().join("gate-report.json");
    write_institution_config(&config)?;
    fs::copy(&fixture, &source)?;
    fs::write(&capability_file, b"[]")?;
    let before = pytools_core::file_commit::sha256_file(&source)?;

    let output = run_cli([
        "institution-normalize",
        "--config",
        config.to_str().context("temporary path is UTF-8")?,
        "--apply",
        "--capabilities",
        capability_file
            .to_str()
            .context("temporary path is UTF-8")?,
        "--platform",
        "linux-x86_64-test",
        "--report-output",
        report_path.to_str().context("temporary path is UTF-8")?,
        "--view",
        "developer",
        source.to_str().context("temporary path is UTF-8")?,
    ])?;
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let run: Value = serde_json::from_slice(&output.stdout)?;
    assert_eq!(run["status"], "partial_failure");
    assert_eq!(run["failed_files"], 1);
    assert_eq!(run["files"][0]["error_stage"], "engine_execution");
    assert!(run["files"][0]["actual_engine"].is_null());
    assert_eq!(
        run["files"][0]["route_diagnostic"]["kind"],
        "no_eligible_backend"
    );
    assert_eq!(pytools_core::file_commit::sha256_file(&source)?, before);
    assert!(!fs::read_dir(temp.path())?.any(|entry| {
        entry.ok().is_some_and(|entry| {
            entry
                .file_name()
                .to_string_lossy()
                .starts_with(".pytools-institution-normalize-")
        })
    }));

    Ok(())
}

#[test]
fn duplicate_mapping_keys_use_last_value_without_changing_first_match_order() -> Result<()> {
    let root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let fixture = root.join("tests/fixtures/institution_normalize.xlsx");
    let temp = TempDirectory::new("pytools-pt-7-2-duplicate-mapping")?;
    let config = temp.path().join("duplicate-institution-config.xlsx");
    let source = temp.path().join("source.xlsx");
    let plan_dir = temp.path().join("plans");
    let report_path = temp.path().join("run.json");
    write_duplicate_institution_config(&config)?;
    fs::copy(&fixture, &source)?;
    let source_hash = pytools_core::file_commit::sha256_file(&source)?;

    let output = run_cli([
        "institution-normalize",
        "--config",
        config.to_str().context("temporary path is UTF-8")?,
        "--plan-dir",
        plan_dir.to_str().context("temporary path is UTF-8")?,
        "--report-output",
        report_path.to_str().context("temporary path is UTF-8")?,
        "--view",
        "developer",
        source.to_str().context("temporary path is UTF-8")?,
    ])?;
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );

    let plan_path = plan_dir.join("0001_institution_normalize.json");
    let plan: Value = serde_json::from_slice(&fs::read(plan_path)?)?;
    let operations = plan["operations"]
        .as_array()
        .context("plan operations missing")?;
    let updates = operations
        .iter()
        .map(|operation| {
            (
                operation["cell"].as_str().unwrap_or_default(),
                operation["value"]["value"].as_str().unwrap_or_default(),
            )
        })
        .collect::<Vec<_>>();
    assert!(updates.contains(&("A2", "Exact Standard")));
    assert!(updates.contains(&("A3", "Last AB Standard")));
    assert!(!updates
        .iter()
        .any(|(_, value)| *value == "First AB Standard"));
    assert_eq!(
        pytools_core::file_commit::sha256_file(&source)?,
        source_hash
    );
    Ok(())
}

#[test]
fn floating_mapping_key_with_python_scientific_notation_is_planned_by_cli() -> Result<()> {
    let temp = TempDirectory::new("pytools-pt-7-2-float-string")?;
    let config = temp.path().join("institution-config.xlsx");
    let source = temp.path().join("numeric-source.xlsx");
    let plan_dir = temp.path().join("plans");
    let report_path = temp.path().join("run.json");
    write_numeric_institution_config(&config)?;

    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("示例分机构")?;
    sheet.write_string(0, 0, "机构")?;
    sheet.write_number(1, 0, 0.00001)?;
    workbook.save(&source)?;
    let source_hash = pytools_core::file_commit::sha256_file(&source)?;
    let config_hash = pytools_core::file_commit::sha256_file(&config)?;

    let output = run_cli([
        "institution-normalize",
        "--config",
        config.to_str().context("temporary path is UTF-8")?,
        "--plan-dir",
        plan_dir.to_str().context("temporary path is UTF-8")?,
        "--report-output",
        report_path.to_str().context("temporary path is UTF-8")?,
        "--view",
        "developer",
        source.to_str().context("temporary path is UTF-8")?,
    ])?;
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );

    let report: Value = serde_json::from_slice(&fs::read(&report_path)?)?;
    assert_eq!(report["files"][0]["mapped_rows"], 1);
    let plan_path = plan_dir.join("0001_institution_normalize.json");
    let plan: Value = serde_json::from_slice(&fs::read(&plan_path)?)?;
    let operations = plan["operations"]
        .as_array()
        .context("plan operations missing")?;
    assert_eq!(operations.len(), 1);
    assert_eq!(operations[0]["sheet"], "示例分机构");
    assert_eq!(operations[0]["cell"], "A2");
    assert_eq!(operations[0]["value"]["value"], "Float Standard");
    assert_eq!(
        pytools_core::file_commit::sha256_file(&source)?,
        source_hash
    );
    assert_eq!(
        pytools_core::file_commit::sha256_file(&config)?,
        config_hash
    );

    if let Some(path) = std::env::var_os("PYTOOLS_INSTITUTION_FLOAT_STRING_REPORT") {
        let report = json!({
            "schema_version": 1,
            "suite": "pytools.PT-7-2.python_float_string_compatibility",
            "status": "passed_synthetic_rust_cli_regression",
            "baseline": "PyCode/core/feature_1_2_normalize_institution.py::_process_sheet",
            "rust_entry": "pytools institution-normalize --view developer",
            "case": {
                "input_kind": "synthetic XLSX numeric cell",
                "mapping_match": "Python float scientific-notation string",
                "expected_mapped_rows": 1
            },
            "comparison": {
                "mapped_rows": "exact",
                "planned_cell_update": "exact",
                "source_and_configuration_unchanged": true
            },
            "safety": {
                "source_unchanged": true,
                "configuration_unchanged": true,
                "outputs_temporary_only": true,
                "report_contains_input_filenames_or_cell_values": false
            },
            "limitations": [
                "Synthetic XLSX only; no authorized real institution-normalization workbook was available.",
                "This verifies planning semantics, not ExcelModifier execution or package preservation.",
                "Production capability and target Linux/deb acceptance remain unverified."
            ]
        });
        fs::write(path, serde_json::to_vec_pretty(&report)?)?;
    }

    Ok(())
}

#[test]
fn empty_mapping_still_emits_a_completed_task_summary() -> Result<()> {
    let root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let fixture = root.join("tests/fixtures/institution_normalize.xlsx");
    let temp = TempDirectory::new("pytools-pt-7-2-empty-mapping")?;
    let config = temp.path().join("empty-institution-config.xlsx");
    let source = temp.path().join("source.xlsx");
    let report_path = temp.path().join("empty-mapping-run.json");
    write_empty_institution_config(&config)?;
    fs::copy(&fixture, &source)?;
    let source_hash = pytools_core::file_commit::sha256_file(&source)?;

    let output = run_cli([
        "institution-normalize",
        "--config",
        config.to_str().context("temporary path is UTF-8")?,
        "--plan-dir",
        temp.path().to_str().context("temporary path is UTF-8")?,
        "--report-output",
        report_path.to_str().context("temporary path is UTF-8")?,
        source.to_str().context("temporary path is UTF-8")?,
    ])?;
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let user_view: Value = serde_json::from_slice(&output.stdout)?;
    assert_eq!(user_view["status"], "已完成");
    assert_eq!(user_view["summary"]["mapping_empty"], true);
    assert_eq!(user_view["summary"]["completed_files"], 0);
    assert_eq!(user_view["summary"]["failed_files"], 0);
    assert!(user_view["events"]
        .as_array()
        .unwrap()
        .iter()
        .any(|event| { event["phase"] == "任务汇总" && event["status"] == "已完成" }));
    assert_eq!(
        pytools_core::file_commit::sha256_file(&source)?,
        source_hash
    );
    Ok(())
}

#[test]
fn configuration_failure_is_saved_in_user_and_developer_history() -> Result<()> {
    let root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let fixture = root.join("tests/fixtures/institution_normalize.xlsx");
    let temp = TempDirectory::new("pytools-pt-7-2-config-failure")?;
    let missing_config = temp.path().join("missing-config.xlsx");
    let source = temp.path().join("source.xlsx");
    let report_path = temp.path().join("failed-run.json");
    fs::copy(&fixture, &source)?;
    let source_hash = pytools_core::file_commit::sha256_file(&source)?;

    let output = run_cli([
        "institution-normalize",
        "--config",
        missing_config.to_str().context("temporary path is UTF-8")?,
        "--plan-dir",
        temp.path().to_str().context("temporary path is UTF-8")?,
        "--report-output",
        report_path.to_str().context("temporary path is UTF-8")?,
        "--view",
        "user",
        source.to_str().context("temporary path is UTF-8")?,
    ])?;
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let user_view: Value = serde_json::from_slice(&output.stdout)?;
    assert_eq!(user_view["status"], "失败");
    assert_eq!(user_view["failure"]["stage"], "读取映射配置");
    assert_eq!(
        user_view["failure"]["message"],
        "机构映射配置读取失败，未处理所选工作簿。"
    );
    assert!(!user_view.to_string().contains("No such file"));

    let developer_history: Value = serde_json::from_slice(&fs::read(&report_path)?)?;
    assert_eq!(
        developer_history["failure"]["error_type"],
        "configuration_read_failed"
    );
    assert!(developer_history["failure"]["diagnostic"]
        .as_str()
        .unwrap()
        .contains("missing-config.xlsx"));
    assert!(developer_history["events"]
        .as_array()
        .unwrap()
        .iter()
        .any(|event| { event["phase"] == "configuration_read" && event["status"] == "failed" }));
    assert!(developer_history["events"]
        .as_array()
        .unwrap()
        .iter()
        .any(|event| { event["phase"] == "task" && event["status"] == "failed" }));
    assert_eq!(
        pytools_core::file_commit::sha256_file(&source)?,
        source_hash
    );
    Ok(())
}

fn run_cli<const N: usize>(arguments: [&str; N]) -> Result<Output> {
    Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(arguments)
        .output()
        .context("run pytools CLI")
}

fn write_institution_config(path: &Path) -> Result<()> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("机构映射表")?;
    sheet.write_string(0, 0, "原始机构")?;
    sheet.write_string(0, 1, "标准机构")?;
    for (row, (source, target)) in [
        ("AB", "Short Standard"),
        ("ABC Bank", "Exact Standard"),
        ("branch", "Branch Standard"),
    ]
    .iter()
    .enumerate()
    {
        sheet.write_string((row + 1) as u32, 0, *source)?;
        sheet.write_string((row + 1) as u32, 1, *target)?;
    }
    workbook.save(path)?;
    Ok(())
}

fn write_empty_institution_config(path: &Path) -> Result<()> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("机构映射表")?;
    sheet.write_string(0, 0, "原始机构")?;
    sheet.write_string(0, 1, "标准机构")?;
    workbook.save(path)?;
    Ok(())
}

fn write_numeric_institution_config(path: &Path) -> Result<()> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("机构映射表")?;
    sheet.write_string(0, 0, "原始机构")?;
    sheet.write_string(0, 1, "标准机构")?;
    sheet.write_string(1, 0, "1e-05")?;
    sheet.write_string(1, 1, "Float Standard")?;
    workbook.save(path)?;
    Ok(())
}

fn write_duplicate_institution_config(path: &Path) -> Result<()> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("机构映射表")?;
    sheet.write_string(0, 0, "原始机构")?;
    sheet.write_string(0, 1, "标准机构")?;
    for (row, (source, target)) in [
        ("AB", "First AB Standard"),
        ("ABC Bank", "Exact Standard"),
        ("AB", "Last AB Standard"),
    ]
    .iter()
    .enumerate()
    {
        sheet.write_string((row + 1) as u32, 0, *source)?;
        sheet.write_string((row + 1) as u32, 1, *target)?;
    }
    workbook.save(path)?;
    Ok(())
}
