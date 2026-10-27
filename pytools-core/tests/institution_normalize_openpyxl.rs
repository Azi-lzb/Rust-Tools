#![cfg(unix)]

use anyhow::{Context, Result};
use excel_editor::{
    BackendCapability, CapabilityEvidence, CapabilityStatus, EditOperation, EngineId, ExcelEditPlan,
};
use rust_xlsxwriter::Workbook;
use serde_json::{json, Value};
use std::{
    fs,
    path::{Path, PathBuf},
    process::Command,
};

const PLATFORM: &str = "linux-x86_64-pt-7-2-openpyxl-test";

struct TestDirectory(PathBuf);

impl TestDirectory {
    fn new() -> Result<Self> {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)?
            .as_nanos();
        let path = std::env::temp_dir().join(format!(
            "pytools-pt-7-2-openpyxl-{}-{nonce}",
            std::process::id()
        ));
        fs::create_dir_all(&path)?;
        Ok(Self(path))
    }

    fn path(&self, name: &str) -> PathBuf {
        self.0.join(name)
    }
}

impl Drop for TestDirectory {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

struct StagedProductRuntime(PathBuf);

impl StagedProductRuntime {
    fn install(python: &Path) -> Result<Self> {
        let runtime_root = python
            .parent()
            .context("isolated Python has no bin directory")?
            .parent()
            .context("isolated Python has no runtime root")?
            .canonicalize()?;
        let product_binary = PathBuf::from(env!("CARGO_BIN_EXE_pytools"));
        let installation_root = product_binary
            .parent()
            .context("pytools binary has no bin directory")?
            .parent()
            .context("pytools binary has no installation root")?;
        let package_dir = installation_root
            .join("libexec")
            .join("pytools")
            .join("openpyxl");
        anyhow::ensure!(
            !package_dir.exists(),
            "refusing to replace an existing openpyxl package layout: {}",
            package_dir.display()
        );
        fs::create_dir_all(package_dir.join("worker"))?;
        std::os::unix::fs::symlink(&runtime_root, package_dir.join("runtime"))?;
        fs::copy(
            Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../crates/excel-editor/python/openpyxl_worker.py"),
            package_dir.join("worker/openpyxl_worker.py"),
        )?;
        Ok(Self(package_dir))
    }
}

impl Drop for StagedProductRuntime {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

fn write_mapping_config(path: &Path) -> Result<()> {
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

fn capability_for(plan: &ExcelEditPlan) -> BackendCapability {
    BackendCapability {
        engine: EngineId::OpenpyxlWorker,
        status: CapabilityStatus::Passed,
        format: plan.format,
        operations: plan.operations.iter().map(EditOperation::kind).collect(),
        supported_features: plan.workbook.features_present.clone(),
        preserved: plan.preserve.clone(),
        platform: PLATFORM.to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: "synthetic-pt-7-2-institution-normalize".to_owned(),
            report_path:
                "pytools-core/tests/institution_normalize_openpyxl.rs (temporary test only)"
                    .to_owned(),
            adapter_version: "openpyxl-worker-1".to_owned(),
            engine_version: "openpyxl 3.1.5".to_owned(),
            platform: PLATFORM.to_owned(),
        }),
    }
}

#[test]
#[ignore = "requires RUST_WORKSPACE_OPENPYXL_PYTHON isolated runtime; run explicitly with --ignored"]
fn pt_7_2_batch_apply_uses_openpyxl_worker_and_preserves_fixture_objects() -> Result<()> {
    let python = PathBuf::from(
        std::env::var_os("RUST_WORKSPACE_OPENPYXL_PYTHON")
            .context("set RUST_WORKSPACE_OPENPYXL_PYTHON to the isolated test runtime")?,
    );
    let _runtime = StagedProductRuntime::install(&python)?;
    let version = Command::new(&python)
        .args(["-I", "-c", "import openpyxl; print(openpyxl.__version__)"])
        .output()?;
    anyhow::ensure!(version.status.success(), "isolated openpyxl import failed");
    let openpyxl_version = String::from_utf8(version.stdout)?.trim().to_owned();
    anyhow::ensure!(
        openpyxl_version == "3.1.5",
        "expected test openpyxl 3.1.5, got {openpyxl_version}"
    );

    let root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let fixture = root.join("tests/fixtures/institution_normalize.xlsx");
    let baseline: Value = serde_json::from_slice(&fs::read(
        root.join("probe/reports/institution_normalize_baseline.json"),
    )?)?;
    let mapping = baseline["mapping"]
        .as_array()
        .context("Python baseline mapping is missing")?
        .iter()
        .map(|entry| {
            Ok((
                entry[0]
                    .as_str()
                    .context("mapping source must be text")?
                    .to_owned(),
                entry[1]
                    .as_str()
                    .context("mapping target must be text")?
                    .to_owned(),
            ))
        })
        .collect::<Result<Vec<_>>>()?;

    let temp = TestDirectory::new()?;
    let source = temp.path("source.xlsx");
    let before = temp.path("before.xlsx");
    let result_dir =
        temp.path(pytools_core::institution_normalize::INSTITUTION_NORMALIZE_OUTPUT_DIRECTORY);
    let config = temp.path("institution-config.xlsx");
    let capability_path = temp.path("temporary-test-capability.json");
    let history_path = temp.path("run-history.json");
    fs::copy(&fixture, &source)?;
    fs::copy(&fixture, &before)?;
    write_mapping_config(&config)?;
    let source_hash_before = pytools_core::file_commit::sha256_file(&source)?;

    let profile = excel_editor::inspect_xlsx_profile(&source)?;
    anyhow::ensure!(
        profile.inventory_complete,
        "synthetic fixture inventory is incomplete"
    );
    let plan_result = pytools_core::institution_normalize::build_institution_normalize_plan(
        &source,
        temp.path("capability-plan-output.xlsx"),
        profile,
        "pt-7-2-openpyxl-test-capability",
        &mapping,
    )?;
    let plan = plan_result
        .plan
        .context("Python baseline expects normalization edits")?;
    anyhow::ensure!(
        plan.operations.len() == 3,
        "expected three cell updates in fixture"
    );
    let capability = capability_for(&plan);
    fs::write(&capability_path, serde_json::to_vec_pretty(&capability)?)?;

    let execution = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args([
            "institution-normalize",
            "--config",
            config.to_str().context("temporary path is not UTF-8")?,
            "--apply",
            "--capabilities",
            capability_path
                .to_str()
                .context("temporary path is not UTF-8")?,
            "--platform",
            PLATFORM,
            "--report-output",
            history_path
                .to_str()
                .context("temporary path is not UTF-8")?,
            "--view",
            "developer",
        ])
        .arg(&source)
        .output()?;
    anyhow::ensure!(
        execution.status.success(),
        "PT-7-2 failed: {}",
        String::from_utf8_lossy(&execution.stderr)
    );
    let report: Value = serde_json::from_slice(&fs::read(&history_path)?)?;
    anyhow::ensure!(
        report["status"] == "completed",
        "run did not complete: {report}"
    );
    anyhow::ensure!(
        report["modified_files"] == 1,
        "file was not committed: {report}"
    );
    anyhow::ensure!(
        report["files"][0]["mapped_rows"] == 3,
        "mapped rows differ: {report}"
    );
    let output_path = PathBuf::from(
        report["files"][0]["output_path"]
            .as_str()
            .context("normalized copy path missing from report")?,
    );
    anyhow::ensure!(
        output_path.starts_with(&result_dir),
        "normalized copy was not saved in the source-adjacent result directory: {}",
        output_path.display()
    );
    anyhow::ensure!(
        report["files"][0]["commit"]["operation"] == "created",
        "normalized copy publication was not recorded: {report}"
    );
    anyhow::ensure!(
        report["files"][0]["actual_engine"] == "openpyxl_worker",
        "actual engine was not recorded: {report}"
    );
    anyhow::ensure!(
        report["files"][0]["execution"]["engine"] == "openpyxl_worker",
        "ExcelModifier execution report is missing: {report}"
    );
    anyhow::ensure!(
        report["files"][0]["execution"]["route_decisions"]
            .as_array()
            .is_some_and(|decisions| {
                decisions.len() == 1
                    && decisions[0]["engine"] == "openpyxl_worker"
                    && decisions[0]["eligible"] == true
            }),
        "the run did not select the temporary openpyxl capability: {report}"
    );
    anyhow::ensure!(
        report["events"].as_array().unwrap().iter().any(|event| {
            event["phase"] == "engine_execution"
                && event["status"] == "completed"
                && event["components"]
                    .as_array()
                    .unwrap()
                    .contains(&json!("Python"))
                && event["components"]
                    .as_array()
                    .unwrap()
                    .contains(&json!("openpyxl_worker"))
        }),
        "developer history does not identify the executed Python/openpyxl worker"
    );
    anyhow::ensure!(
        pytools_core::file_commit::sha256_file(&source)? == source_hash_before,
        "source workbook changed despite copy-only normalization"
    );
    anyhow::ensure!(
        pytools_core::file_commit::sha256_file(&before)? == source_hash_before,
        "pre-execution snapshot changed"
    );

    let output = excel_reader::read_workbook(&output_path)?;
    let expected_sheets = baseline["sheets"]
        .as_array()
        .context("Python baseline sheets are missing")?;
    anyhow::ensure!(output.sheets.len() == expected_sheets.len());
    for (actual, expected) in output.sheets.iter().zip(expected_sheets) {
        anyhow::ensure!(
            actual.name == expected["name"],
            "worksheet order/name changed"
        );
        anyhow::ensure!(
            serde_json::to_value(&actual.values)? == expected["values"],
            "cell values differ from Python baseline on {}",
            actual.name
        );
        anyhow::ensure!(
            serde_json::to_value(&actual.formulas)? == expected["formulas"],
            "formula text differs from Python baseline on {}",
            actual.name
        );
        anyhow::ensure!(
            serde_json::to_value(&actual.comments)? == expected["comments"],
            "comments differ from Python baseline on {}",
            actual.name
        );
        anyhow::ensure!(
            serde_json::to_value(&actual.merges)? == expected["merges"],
            "merged ranges differ from Python baseline on {}",
            actual.name
        );
    }

    let preservation_script = root.join("probe/institution_normalize_preservation_check.py");
    let preservation = Command::new(&python)
        .arg("-I")
        .arg(&preservation_script)
        .arg(&before)
        .arg(&output_path)
        .output()?;
    anyhow::ensure!(
        preservation.status.success(),
        "workbook object preservation check failed: {}",
        String::from_utf8_lossy(&preservation.stdout)
    );
    let preservation_report: Value = serde_json::from_slice(&preservation.stdout)?;
    anyhow::ensure!(preservation_report["status"] == "保留");

    if let Some(path) = std::env::var_os("PYTOOLS_INSTITUTION_OPENPYXL_REPORT") {
        let report_path = PathBuf::from(path);
        if let Some(parent) = report_path.parent() {
            fs::create_dir_all(parent)?;
        }
        let candidate_report = json!({
            "schema_version": 1,
            "suite": "pytools.PT-7-2.openpyxl-worker",
            "status": "synthetic_passed_not_production_acceptance",
            "python_version": String::from_utf8_lossy(&Command::new(&python).args(["-I", "-c", "import sys; print(sys.version.split()[0])"]).output()?.stdout).trim(),
            "openpyxl_version": openpyxl_version,
            "executed_engine": "openpyxl_worker",
            "fixture": "tests/fixtures/institution_normalize.xlsx (synthetic)",
            "python_baseline": "probe/reports/institution_normalize_baseline.json",
            "stats": {
                "branch_sheets": report["branch_sheets"],
                "mapped_sheets": report["mapped_sheets"],
                "mapped_rows": report["mapped_rows"],
                "modified_files": report["modified_files"],
                "failed_files": report["failed_files"]
            },
            "execution": report,
            "preservation": preservation_report,
            "source_sha256_before": source_hash_before,
            "source_sha256_after": pytools_core::file_commit::sha256_file(&source)?,
            "temporary_test_capability": true,
            "python_worker_is_product_engine": true,
            "python_openpyxl_is_also_used_for_independent_fixture_metadata_checks": true,
            "production_capability_created": false,
            "default_route_changed": false,
            "limitations": [
                "synthetic fixture only; no authorized real business sample",
                "temporary test capability covers this fixture and test platform only",
                "PT-7-2 remains unverified for production",
                "Windows-native worker, target Linux glibc 2.28, and deb packaging are not verified"
            ]
        });
        let mut file = fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&report_path)?;
        use std::io::Write;
        file.write_all(&serde_json::to_vec_pretty(&candidate_report)?)?;
        file.sync_all()?;
    }
    Ok(())
}
