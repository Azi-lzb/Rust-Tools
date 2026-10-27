#![cfg(unix)]

use anyhow::{Context, Result};
use excel_editor::{
    BackendCapability, CapabilityEvidence, CapabilityStatus, EditOperation, EngineId, ExcelEditPlan,
};
use rust_xlsxwriter::{Format, Note, Workbook};
use serde_json::Value;
use std::{
    fs,
    path::{Path, PathBuf},
    process::{Command, Output},
};

const PLATFORM: &str = "linux-x86_64-openpyxl-test";

struct TestDir(PathBuf);

impl TestDir {
    fn new() -> Self {
        let path = std::env::temp_dir().join(format!(
            "pytools-dedup-comment-openpyxl-{}-{}",
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

struct StagedProductRuntime(PathBuf);

impl StagedProductRuntime {
    fn install() -> Result<Self> {
        let python = PathBuf::from(
            std::env::var_os("RUST_WORKSPACE_OPENPYXL_PYTHON")
                .context("set RUST_WORKSPACE_OPENPYXL_PYTHON to the isolated test runtime")?,
        );
        let runtime_root = python
            .parent()
            .context("isolated python has no bin directory")?
            .parent()
            .context("isolated python has no runtime root")?
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

fn make_source(path: &Path) -> Result<()> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Commented")?;
    sheet.write_string(0, 0, "源数据")?;
    sheet.write_string(0, 1, "标识列")?;
    sheet.write_string(0, 2, "值")?;
    sheet.insert_note(0, 0, &Note::new("源数据表"))?;
    sheet.insert_note(0, 1, &Note::new("标识列"))?;
    sheet.write_string(1, 0, "first-duplicate")?;
    sheet.write_string(1, 1, "same")?;
    sheet.write_string_with_format(1, 2, "first", &Format::new().set_bold())?;
    sheet.write_string(2, 0, "second-duplicate")?;
    sheet.write_string(2, 1, "same")?;
    sheet.write_string(2, 2, "duplicate")?;
    sheet.write_string(3, 0, "blank-key")?;
    sheet.write_string(3, 2, "blank keys are ignored")?;
    sheet.write_string(4, 0, "first-other")?;
    sheet.write_string(4, 1, "other")?;
    sheet.write_string_with_format(4, 2, "first", &Format::new().set_bold())?;
    sheet.write_string(5, 0, "second-other")?;
    sheet.write_string(5, 1, "other")?;
    sheet.write_string(5, 2, "duplicate")?;
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
        engine: EngineId::OpenpyxlWorker,
        status: CapabilityStatus::Passed,
        format: plan.format,
        operations: plan.operations.iter().map(EditOperation::kind).collect(),
        supported_features: plan.workbook.features_present.clone(),
        preserved: plan.preserve.clone(),
        platform: PLATFORM.to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: "synthetic-pt-2-1-comment-mark".to_owned(),
            report_path: "pytools-core/tests/dedup_comment_openpyxl.rs".to_owned(),
            adapter_version: "openpyxl-worker-1".to_owned(),
            engine_version: "test-runtime".to_owned(),
            platform: PLATFORM.to_owned(),
        }),
    }
}

#[test]
#[ignore = "requires an isolated openpyxl runtime; run explicitly with --ignored"]
fn pt_2_1_comment_duplicate_mark_executes_through_pytools_cli_and_excel_modifier() -> Result<()> {
    let _runtime = StagedProductRuntime::install()?;
    let dir = TestDir::new();
    let input = dir.path("source.xlsx");
    let output_path = dir.path("candidate.xlsx");
    let plan_path = dir.path("edit-plan.json");
    let capability_path = dir.path("test-capability.json");
    make_source(&input)?;
    let source_bytes = fs::read(&input)?;

    let plan_only = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-comment-plan", "--action", "check", "--input"])
        .arg(&input)
        .arg("--output")
        .arg(&output_path)
        .arg("--plan-output")
        .arg(&plan_path)
        .output()?;
    let report = assert_success(&plan_only)?;
    assert_eq!(report["execution"], "plan_only");
    assert_eq!(report["duplicate_rows"], 2);
    assert_eq!(
        fs::read(&input)?,
        source_bytes,
        "plan-only must be read-only"
    );
    let source_comments = excel_reader::read_workbook(&input)?
        .sheets
        .iter()
        .find(|sheet| sheet.name == "Commented")
        .context("Commented source sheet missing")?
        .comments
        .clone();

    let plan: ExcelEditPlan = serde_json::from_slice(&fs::read(&plan_path)?)?;
    assert_eq!(plan.operations.len(), 6);
    let capability = capability_for(&plan);
    fs::write(&capability_path, serde_json::to_vec_pretty(&capability)?)?;

    let executed = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-comment-plan", "--action", "check", "--input"])
        .arg(&input)
        .arg("--output")
        .arg(&output_path)
        .arg("--capabilities")
        .arg(&capability_path)
        .arg("--platform")
        .arg(PLATFORM)
        .output()?;
    let result = assert_success(&executed)?;
    assert_eq!(result["execution"]["engine"], "openpyxl_worker");
    assert_eq!(
        result["execution"]["output"],
        output_path.to_string_lossy().as_ref()
    );
    assert!(output_path.is_file());
    assert_eq!(
        fs::read(&input)?,
        source_bytes,
        "execution must not overwrite source"
    );

    let output = excel_editor::open_xlsx(&output_path)?;
    let sheet = output
        .sheet_by_name("Commented")
        .context("Commented sheet missing")?;
    for address in ["A3", "B3", "C3", "A6", "B6", "C6"] {
        let fill_color = sheet
            .get_cell(address)
            .and_then(|cell| cell.style().fill())
            .and_then(|fill| fill.pattern_fill())
            .and_then(|fill| fill.foreground_color())
            .map(|color| color.argb_str().to_ascii_uppercase());
        assert_eq!(
            fill_color.as_deref(),
            Some("FFFF0000"),
            "{address} red fill"
        );
    }
    let output_comments = excel_reader::read_workbook(&output_path)?
        .sheets
        .iter()
        .find(|candidate| candidate.name == "Commented")
        .context("Commented reader dump missing")?
        .comments
        .clone();
    assert_eq!(output_comments, source_comments);
    assert_eq!(
        fs::read(&input)?,
        source_bytes,
        "source remains byte-identical"
    );

    let capability_value: Value = serde_json::from_slice(&fs::read(&capability_path)?)?;
    assert_eq!(capability_value["engine"], "openpyxl_worker");
    Ok(())
}
