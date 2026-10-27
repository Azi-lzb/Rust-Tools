use anyhow::{Context, Result};
use excel_editor::{
    BackendCapability, CapabilityEvidence, CapabilityStatus, EditOperation, EngineId, ExcelEditPlan,
};
use rust_xlsxwriter::{Color, Format, Note, Workbook};
use serde_json::Value;
use std::{
    fs,
    path::{Path, PathBuf},
    process::{Command, Output},
};

const PLATFORM: &str = "synthetic-pt-2-2-umya-test";

struct TestDir(PathBuf);

impl TestDir {
    fn new() -> Self {
        let path = std::env::temp_dir().join(format!(
            "pytools-dedup-comment-delete-{}-{}",
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
    {
        let sheet = workbook.add_worksheet();
        sheet.set_name("Commented")?;
        sheet.write_string(0, 0, "源数据")?;
        sheet.write_string(0, 1, "Key")?;
        sheet.write_string(0, 2, "Value")?;
        sheet.insert_note(0, 0, &Note::new("本表为源数据"))?;
        sheet.insert_note(0, 1, &Note::new("标识列"))?;

        sheet.write_string(1, 0, "first")?;
        sheet.write_string(1, 1, "dup")?;
        sheet.write_number(1, 2, 10.0)?;
        sheet.write_string(2, 0, "duplicate")?;
        sheet.write_string(2, 1, "dup")?;
        sheet.write_number(2, 2, 20.0)?;
        sheet.write_string(3, 0, "keep")?;
        sheet.write_string(3, 1, "keep")?;
        sheet.write_number_with_format(
            3,
            2,
            30.0,
            &Format::new().set_background_color(Color::RGB(0x00_FF_00)),
        )?;
        sheet.write_string(4, 0, "tail")?;
        sheet.write_string(4, 1, "tail")?;
        sheet.write_number(4, 2, 40.0)?;

        sheet.merge_range(1, 4, 4, 5, "merged-marker", &Format::new())?;
        sheet.insert_note(2, 6, &Note::new("deleted-row-comment"))?;
        sheet.insert_note(3, 6, &Note::new("survivor-comment"))?;
        sheet.insert_note(4, 6, &Note::new("tail-comment"))?;
        sheet.set_row_height(3, 24.0)?;
        sheet.autofilter(0, 0, 4, 2)?;
        sheet.write_formula(0, 7, "=SUM(C2:C5)")?;
        sheet.write_formula(1, 7, "=SUM($C$2:$C$5)")?;
    }
    {
        let refs = workbook.add_worksheet();
        refs.set_name("Refs")?;
        refs.write_formula(0, 0, "=Commented!C4")?;
        refs.write_formula(1, 0, "=SUM(Commented!C2:C5)")?;
        refs.write_formula(2, 0, "=A1+Commented!C5")?;
        refs.write_formula(3, 0, "=SUM(Commented!$C$2:$C$5)")?;
        refs.write_formula(4, 0, "=SUM(A1:A3)")?;
    }
    workbook.define_name("DataValues", "Commented!$C$2:$C$5")?;
    workbook.save(path)?;
    Ok(())
}

fn make_config(path: &Path, source: &Path) -> Result<()> {
    let mut workbook = Workbook::new();
    let global = workbook.add_worksheet();
    global.set_name("全局配置")?;
    global.write_string(0, 0, "配置项")?;
    global.write_string(0, 1, "值")?;
    global.write_string(1, 0, "输入目录")?;
    global.write_string(1, 1, ".")?;
    // absolute per-test output/log dirs: parallel CLI tests must not share
    // the crate-root relative "output"/"logs" directories
    let output_dir = path
        .parent()
        .context("configuration parent missing")?
        .join("output");
    let log_dir = path
        .parent()
        .context("configuration parent missing")?
        .join("logs");
    global.write_string(2, 0, "输出目录")?;
    global.write_string(2, 1, output_dir.to_string_lossy().into_owned())?;
    global.write_string(3, 0, "日志目录")?;
    global.write_string(3, 1, log_dir.to_string_lossy().into_owned())?;

    let sheet = workbook.add_worksheet();
    sheet.set_name("去重追加数据配置")?;
    for (column, header) in [
        "是否启用",
        "任务名",
        "源数据工作簿",
        "源数据工作表",
        "标识列序号",
    ]
    .iter()
    .enumerate()
    {
        sheet.write_string(0, column as u16, *header)?;
    }
    sheet.write_string(1, 0, "是")?;
    sheet.write_string(1, 1, "Delete-duplicates")?;
    sheet.write_string(1, 2, source.to_string_lossy().as_ref())?;
    sheet.write_string(1, 3, "Commented")?;
    sheet.write_string(1, 4, "2")?;
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

fn capability_for(plan: &ExcelEditPlan, sample_id: &str) -> BackendCapability {
    BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Passed,
        format: plan.format,
        operations: plan.operations.iter().map(EditOperation::kind).collect(),
        supported_features: plan.workbook.features_present.clone(),
        preserved: plan.preserve.clone(),
        platform: PLATFORM.to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: sample_id.to_owned(),
            report_path: "pytools-core/tests/dedup_comment_delete_umya.rs".to_owned(),
            adapter_version: "umya-spreadsheet-3.1.1".to_owned(),
            engine_version: "umya-spreadsheet-3.1.1".to_owned(),
            platform: PLATFORM.to_owned(),
        }),
    }
}

fn sheet<'a>(
    book: &'a excel_reader::WorkbookDump,
    name: &str,
) -> Result<&'a excel_reader::SheetDump> {
    book.sheets
        .iter()
        .find(|candidate| candidate.name == name)
        .with_context(|| format!("worksheet {name:?} missing"))
}

fn cell_text<'a>(sheet: &'a excel_reader::SheetDump, address: &str) -> Option<&'a str> {
    sheet
        .values
        .get(address)
        .and_then(|value| value.get("value"))
        .and_then(Value::as_str)
}

fn comment_body(value: &str) -> &str {
    value.split_once('\n').map_or(value, |(_, body)| body)
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
fn pt_2_2_comment_duplicate_delete_tracks_excel_row_objects_and_references() -> Result<()> {
    let dir = TestDir::new();
    let input = dir.path("source.xlsx");
    let output_path = dir.path("candidate.xlsx");
    let plan_path = dir.path("edit-plan.json");
    let capability_path = dir.path("test-capability.json");
    make_source(&input)?;
    let source_bytes = fs::read(&input)?;

    let plan_only = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-comment-plan", "--input"])
        .arg(&input)
        .args(["--output"])
        .arg(&output_path)
        .args(["--action", "delete", "--plan-output"])
        .arg(&plan_path)
        .output()?;
    let report = assert_success(&plan_only)?;
    assert_eq!(report["execution"], "plan_only");
    assert_eq!(report["duplicate_rows"], 1);
    assert_eq!(fs::read(&input)?, source_bytes, "plan-only is read-only");

    let plan: ExcelEditPlan = serde_json::from_slice(&fs::read(&plan_path)?)?;
    assert_eq!(plan.operations.len(), 1);
    assert!(matches!(
        &plan.operations[0],
        EditOperation::DeleteRows {
            sheet,
            first_row: 3,
            count: 1
        } if sheet == "Commented"
    ));
    assert!(!plan
        .preserve
        .contains(&excel_editor::PreservationRequirement::UnmodifiedPackageParts));
    fs::write(
        &capability_path,
        serde_json::to_vec_pretty(&capability_for(
            &plan,
            "synthetic-pt-2-2-comment-row-delete",
        ))?,
    )?;

    let executed = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-comment-plan", "--input"])
        .arg(&input)
        .args(["--output"])
        .arg(&output_path)
        .args(["--action", "delete", "--capabilities"])
        .arg(&capability_path)
        .args(["--platform", PLATFORM])
        .output()?;
    let execution_report = assert_success(&executed)?;
    assert_eq!(execution_report["execution"]["engine"], "umya_spreadsheet");
    assert_eq!(
        execution_report["execution"]["attempts"][0]["outcome"],
        "succeeded"
    );
    assert!(output_path.is_file());
    assert_eq!(
        fs::read(&input)?,
        source_bytes,
        "execution leaves source intact"
    );

    let output = excel_reader::read_workbook(&output_path)?;
    let commented = sheet(&output, "Commented")?;
    let refs = sheet(&output, "Refs")?;
    assert_eq!(cell_text(commented, "A2"), Some("first"));
    assert_eq!(cell_text(commented, "A3"), Some("keep"));
    assert_eq!(cell_text(commented, "A4"), Some("tail"));
    assert_eq!(
        commented.formulas.get("H1").map(String::as_str),
        Some("SUM(C2:C4)")
    );
    assert_eq!(
        refs.formulas.get("A1").map(String::as_str),
        Some("Commented!C3")
    );
    assert_eq!(
        refs.formulas.get("A2").map(String::as_str),
        Some("SUM(Commented!C2:C4)")
    );
    assert_eq!(
        refs.formulas.get("A3").map(String::as_str),
        Some("A1+Commented!C4"),
        "formula remains on its non-target sheet and follows the deleted row reference"
    );
    assert_eq!(
        refs.formulas.get("A4").map(String::as_str),
        Some("SUM(Commented!$C$2:$C$4)"),
        "absolute row references are adjusted by structural deletion"
    );
    assert_eq!(
        refs.formulas.get("A5").map(String::as_str),
        Some("SUM(A1:A3)"),
        "unqualified references on a non-target sheet remain unchanged"
    );
    assert_eq!(
        commented
            .comments
            .get("G3")
            .map(|value| comment_body(value)),
        Some("survivor-comment")
    );
    assert_eq!(
        commented
            .comments
            .get("G4")
            .map(|value| comment_body(value)),
        Some("tail-comment")
    );
    assert!(!commented.comments.contains_key("G5"));
    assert!(commented.comments.contains_key("A1"));
    assert!(commented.comments.contains_key("B1"));
    assert_eq!(commented.merges, ["E2:F4"]);

    let output_umya = excel_editor::open_xlsx(&output_path)?;
    let source_umya = excel_editor::open_xlsx(&input)?;
    let output_names = output_umya
        .defined_names()
        .iter()
        .chain(
            output_umya
                .sheet_collection()
                .iter()
                .flat_map(|sheet| sheet.defined_names()),
        )
        .map(|name| format!("{}={}", name.name(), name.address()))
        .collect::<Vec<_>>();
    assert_eq!(
        output_umya
            .sheet_by_name("Commented")?
            .cell("H2")
            .context("H2 absolute formula missing")?
            .cell_value()
            .formula(),
        "SUM($C$2:$C$4)",
        "absolute formula row references are adjusted without losing their locks"
    );
    assert_eq!(
        output_umya
            .defined_names()
            .iter()
            .chain(
                output_umya
                    .sheet_collection()
                    .iter()
                    .flat_map(|sheet| sheet.defined_names()),
            )
            .find(|name| name.name() == "DataValues")
            .with_context(|| format!("DataValues defined name missing; found {output_names:?}"))?
            .address(),
        "Commented!$C$2:$C$4",
        "absolute defined-name row references are adjusted without losing their locks"
    );
    assert_eq!(
        style_fill(&output_umya, "Commented", "C3")?,
        style_fill(&source_umya, "Commented", "C4")?,
        "style follows the surviving row"
    );
    let output_sheet = output_umya
        .sheet_by_name("Commented")
        .context("Commented sheet missing from Umya readback")?;
    let source_sheet = source_umya
        .sheet_by_name("Commented")
        .context("Commented source sheet missing from Umya readback")?;
    assert_eq!(
        output_sheet
            .row_dimension(3)
            .context("row 3 dimension missing")?
            .height(),
        source_sheet
            .row_dimension(4)
            .context("row 4 dimension missing")?
            .height(),
        "row height follows the surviving row"
    );
    Ok(())
}

#[test]
fn pt_2_4_config_duplicate_delete_executes_as_one_umya_workbook_batch() -> Result<()> {
    let dir = TestDir::new();
    let input = dir.path("source.xlsx");
    let config = dir.path("config.xlsx");
    let output_dir = dir.path("output");
    let capability_path = dir.path("test-capability.json");
    make_source(&input)?;
    make_config(&config, &input)?;
    let source_bytes = fs::read(&input)?;

    let plan_only = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-config-plan", "--config"])
        .arg(&config)
        .args(["--output-dir"])
        .arg(&output_dir)
        .args(["--action", "delete"])
        .output()?;
    let report = assert_success(&plan_only)?;
    assert_eq!(report[0]["execution"], "plan_only");
    assert_eq!(report[0]["report"]["duplicate_rows"], 1);
    let plan_path = PathBuf::from(
        report[0]["plan_file"]
            .as_str()
            .context("plan file path missing")?,
    );
    let plan: ExcelEditPlan = serde_json::from_slice(&fs::read(plan_path)?)?;
    assert_eq!(plan.operations.len(), 1);
    let capability = capability_for(&plan, "synthetic-pt-2-4-config-delete");
    fs::write(&capability_path, serde_json::to_vec_pretty(&capability)?)?;

    let executed = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-config-plan", "--config"])
        .arg(&config)
        .args(["--output-dir"])
        .arg(&output_dir)
        .args(["--action", "delete", "--capabilities"])
        .arg(&capability_path)
        .args(["--platform", PLATFORM])
        .output()?;
    let execution = assert_success(&executed)?;
    assert_eq!(execution[0]["execution"]["engine"], "umya_spreadsheet");
    assert_eq!(
        execution[0]["execution"]["attempts"][0]["outcome"],
        "succeeded"
    );
    let candidate = PathBuf::from(plan.output);
    assert!(candidate.is_file());
    assert_eq!(fs::read(&input)?, source_bytes, "source remains unchanged");

    let output = excel_reader::read_workbook(&candidate)?;
    let commented = sheet(&output, "Commented")?;
    let refs = sheet(&output, "Refs")?;
    assert_eq!(cell_text(commented, "A3"), Some("keep"));
    assert_eq!(cell_text(commented, "A4"), Some("tail"));
    assert_eq!(
        commented.formulas.get("H1").map(String::as_str),
        Some("SUM(C2:C4)")
    );
    assert_eq!(
        refs.formulas.get("A1").map(String::as_str),
        Some("Commented!C3")
    );
    assert_eq!(
        refs.formulas.get("A3").map(String::as_str),
        Some("A1+Commented!C4")
    );
    assert_eq!(
        refs.formulas.get("A4").map(String::as_str),
        Some("SUM(Commented!$C$2:$C$4)")
    );
    assert_eq!(commented.merges, ["E2:F4"]);
    assert_eq!(
        commented
            .comments
            .get("G3")
            .map(|value| comment_body(value)),
        Some("survivor-comment")
    );
    let output_umya = excel_editor::open_xlsx(&candidate)?;
    let output_names = output_umya
        .defined_names()
        .iter()
        .chain(
            output_umya
                .sheet_collection()
                .iter()
                .flat_map(|sheet| sheet.defined_names()),
        )
        .map(|name| format!("{}={}", name.name(), name.address()))
        .collect::<Vec<_>>();
    assert_eq!(
        output_umya
            .defined_names()
            .iter()
            .chain(
                output_umya
                    .sheet_collection()
                    .iter()
                    .flat_map(|sheet| sheet.defined_names()),
            )
            .find(|name| name.name() == "DataValues")
            .with_context(|| format!("DataValues defined name missing; found {output_names:?}"))?
            .address(),
        "Commented!$C$2:$C$4"
    );
    Ok(())
}
