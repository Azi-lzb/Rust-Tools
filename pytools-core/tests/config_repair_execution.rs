#![cfg(unix)]

use anyhow::{Context, Result};
use excel_editor::{
    BackendCapability, CapabilityEvidence, CapabilityStatus, EditOperation, EngineId,
    ExcelEditPlan, PreservationRequirement,
};
use excel_reader::{read_workbook, SheetDump};
use rust_xlsxwriter::{Color, Format, Note, Workbook};
use serde_json::Value;
use std::{
    fs,
    path::{Path, PathBuf},
    process::{Command, Output},
};

const PLATFORM: &str = "synthetic-pt-5-1-config-repair-openpyxl-test";

struct TestDir(PathBuf);

impl TestDir {
    fn new() -> Self {
        let path = std::env::temp_dir().join(format!(
            "pytools-config-repair-{}-{}",
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
            .context("isolated Python executable has no bin directory")?
            .parent()
            .context("isolated Python executable has no runtime root")?
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
            "refusing to replace existing openpyxl package layout: {}",
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

fn make_incomplete_config(path: &Path) -> Result<()> {
    let mut workbook = Workbook::new();

    let global = workbook.add_worksheet();
    global.set_name("全局配置")?;
    for (column, value) in ["wrong-key", "wrong-value", "wrong-note"]
        .iter()
        .enumerate()
    {
        global.write_string(0, column as u16, *value)?;
    }
    global.insert_note(0, 0, &Note::new("old header note"))?;
    global.write_string(1, 0, "输入目录")?;
    global.write_string(1, 1, "custom-input")?;
    global.write_string(2, 0, "输出目录")?;
    global.write_string(3, 0, "错误策略")?;
    global.write_string(3, 1, "first-policy")?;
    global.write_string(4, 0, "错误策略")?;

    let rename = workbook.add_worksheet();
    rename.set_name("重命名配置")?;
    for column in 0..17 {
        rename.write_string(0, column, format!("wrong-{column}"))?;
    }
    rename.write_string(1, 0, "user alias")?;
    rename.insert_note(0, 11, &Note::new("obsolete header note"))?;
    rename.insert_note(1, 11, &Note::new("obsolete cell note"))?;

    let archive = workbook.add_worksheet();
    archive.set_name("归档类型配置")?;
    for (column, value) in [
        "wrong-1", "wrong-2", "wrong-3", "wrong-4", "wrong-5", "wrong-6",
    ]
    .iter()
    .enumerate()
    {
        archive.write_string(0, column as u16, *value)?;
    }

    let sort = workbook.add_worksheet();
    sort.set_name("文件排序配置")?;
    for (column, value) in ["是否启用", "排序前缀", "匹配关键词", "备注"]
        .iter()
        .enumerate()
    {
        sort.write_string(0, column as u16, *value)?;
    }
    for (column, value) in ["否", "777", "keep-me", "existing user data"]
        .iter()
        .enumerate()
    {
        sort.write_string(1, column as u16, *value)?;
    }

    for obsolete in ["校验规则生成配置", "校验规则映射表"] {
        let sheet = workbook.add_worksheet();
        sheet.set_name(obsolete)?;
        sheet.write_string(1, 0, "remove-me")?;
    }

    let custom = workbook.add_worksheet();
    custom.set_name("Custom user sheet")?;
    custom.write_string_with_format(
        0,
        0,
        "keep custom data",
        &Format::new()
            .set_background_color(Color::RGB(0x00_FF00))
            .set_bold(),
    )?;
    custom.insert_note(2, 1, &Note::new("keep custom comment"))?;
    custom.merge_range(0, 3, 0, 4, "keep merged", &Format::new())?;
    custom.set_row_height(1, 29.0)?;
    custom.set_column_width(5, 22.0)?;

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
            sample_id: "synthetic-pt-5-1-config-repair".to_owned(),
            report_path: "pytools-core/tests/config_repair_execution.rs".to_owned(),
            adapter_version: "openpyxl-worker-1".to_owned(),
            engine_version: "openpyxl 3.1.5".to_owned(),
            platform: PLATFORM.to_owned(),
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

#[test]
#[ignore = "requires isolated openpyxl runtime; run explicitly with --ignored"]
fn pt_5_1_config_repair_executes_as_one_capability_gated_openpyxl_batch() -> Result<()> {
    let _runtime = StagedProductRuntime::install()?;
    let dir = TestDir::new();
    let input = dir.path("incomplete.xlsx");
    let candidate = dir.path("candidate.xlsx");
    let plan_path = dir.path("repair-plan.json");
    let capability_path = dir.path("test-capability.json");
    make_incomplete_config(&input)?;
    let input_before = fs::read(&input)?;

    let plan_only = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["config-repair-plan", "--input"])
        .arg(&input)
        .args(["--output"])
        .arg(&candidate)
        .args(["--plan-output"])
        .arg(&plan_path)
        .output()?;
    let plan_json = assert_success(&plan_only)?;
    let plan: ExcelEditPlan = serde_json::from_slice(&fs::read(&plan_path)?)?;
    assert_eq!(plan_json["input"], input.to_string_lossy().as_ref());
    assert_eq!(plan.operations.len(), 256);
    assert!(!plan
        .preserve
        .contains(&PreservationRequirement::UnmodifiedPackageParts));
    assert!(!candidate.exists(), "plan-only must not write a candidate");
    assert_eq!(fs::read(&input)?, input_before);

    fs::write(
        &capability_path,
        serde_json::to_vec_pretty(&capability_for(&plan))?,
    )?;
    let executed = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["config-repair-plan", "--input"])
        .arg(&input)
        .args(["--output"])
        .arg(&candidate)
        .args(["--capabilities"])
        .arg(&capability_path)
        .args(["--platform", PLATFORM])
        .output()?;
    let response = assert_success(&executed)?;
    assert_eq!(response["execution"]["engine"], "openpyxl_worker");
    assert_eq!(response["execution"]["attempts"][0]["outcome"], "succeeded");
    assert!(candidate.is_file());
    assert_eq!(
        fs::read(&input)?,
        input_before,
        "source config must remain unchanged"
    );

    let output = read_workbook(&candidate)?;
    let sheet_names = output
        .sheets
        .iter()
        .map(|candidate| candidate.name.as_str())
        .collect::<Vec<_>>();
    assert_eq!(
        &sheet_names[..4],
        ["全局配置", "重命名配置", "归档类型配置", "文件排序配置"]
    );
    assert_eq!(sheet_names.len(), 14);
    let obsolete_generate = sheet(&output, "校验规则生成配置")?;
    let obsolete_mapping = sheet(&output, "校验规则映射表")?;
    assert_eq!(text_at(obsolete_generate, "A2"), Some("remove-me"));
    assert_eq!(text_at(obsolete_mapping, "A2"), Some("remove-me"));

    let global = sheet(&output, "全局配置")?;
    assert_eq!(text_at(global, "A1"), Some("键"));
    assert_eq!(text_at(global, "B1"), Some("值"));
    assert_eq!(text_at(global, "C1"), Some("备注"));
    assert_eq!(text_at(global, "B2"), Some("custom-input"));
    assert_eq!(text_at(global, "B3"), Some("output"));
    assert_eq!(text_at(global, "B4"), Some("first-policy"));
    assert_eq!(text_at(global, "B5"), Some("continue"));
    assert!(global
        .comments
        .get("A1")
        .is_some_and(|text| text.contains("关键字段")));

    let rename = sheet(&output, "重命名配置")?;
    assert_eq!(text_at(rename, "A2"), Some("user alias"));
    assert_eq!(text_at(rename, "O2"), Some("Y"));
    assert_eq!(text_at(rename, "P2"), Some("sheet1@A1;sheet2@A2"));
    assert_eq!(rename.comments.get("L1"), None);
    assert_eq!(rename.comments.get("L2"), None);

    let archive = sheet(&output, "归档类型配置")?;
    assert_eq!(text_at(archive, "A2"), Some("是"));
    let sort = sheet(&output, "文件排序配置")?;
    assert_eq!(text_at(sort, "B2"), Some("777"));
    assert_eq!(text_at(sort, "C2"), Some("keep-me"));
    assert_eq!(text_at(sort, "D2"), Some("existing user data"));
    assert_eq!(text_at(sort, "A3"), None, "existing rows suppress examples");

    let custom = sheet(&output, "Custom user sheet")?;
    assert_eq!(text_at(custom, "A1"), Some("keep custom data"));
    assert!(custom
        .comments
        .get("B3")
        .is_some_and(|text| text.contains("keep custom comment")));
    assert_eq!(custom.merges, ["D1:E1"]);

    let source_umya = excel_editor::open_xlsx(&input)?;
    let output_umya = excel_editor::open_xlsx(&candidate)?;
    assert_eq!(
        style_fill(&output_umya, "Custom user sheet", "A1")?,
        style_fill(&source_umya, "Custom user sheet", "A1")?
    );
    let source_custom = source_umya
        .sheet_by_name("Custom user sheet")
        .context("source custom sheet missing")?;
    let output_custom = output_umya
        .sheet_by_name("Custom user sheet")
        .context("output custom sheet missing")?;
    assert_eq!(
        output_custom.row_dimension(2).map(|row| row.height()),
        source_custom.row_dimension(2).map(|row| row.height())
    );
    assert_eq!(
        output_custom
            .column_dimension_by_number(6)
            .map(|column| column.width()),
        source_custom
            .column_dimension_by_number(6)
            .map(|column| column.width())
    );

    Ok(())
}
