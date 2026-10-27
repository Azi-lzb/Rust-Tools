use anyhow::{Context, Result};
use excel_editor::{
    BackendCapability, CapabilityEvidence, CapabilityStatus, EditOperation, EngineId, ExcelEditPlan,
};
use rust_xlsxwriter::{Note, Workbook};
use serde_json::Value;
use std::{
    fs,
    path::{Path, PathBuf},
    process::{Command, Output},
};

const PLATFORM: &str = "synthetic-pt-7-3-umya-test";

struct TestDir(PathBuf);

impl TestDir {
    fn new() -> Self {
        let path = std::env::temp_dir().join(format!(
            "pytools-foreign-bank-{}-{}",
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

fn make_config(path: &Path) -> Result<()> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("机构映射表")?;
    for (column, header) in ["序号", "机构名称", "是否为外资行"].iter().enumerate() {
        sheet.write_string(0, column as u16, *header)?;
    }
    sheet.write_number(1, 0, 1.0)?;
    sheet.write_string(1, 1, "外资甲")?;
    sheet.write_number(1, 2, 1.0)?;
    sheet.write_number(2, 0, 2.0)?;
    sheet.write_string(2, 1, "外资乙")?;
    sheet.write_number(2, 2, 1.0)?;
    sheet.write_number(3, 0, 3.0)?;
    sheet.write_string(3, 1, "普通行")?;
    sheet.write_number(3, 2, 0.0)?;
    workbook.save(path)?;
    Ok(())
}

fn make_source(
    path: &Path,
    include_foreign: bool,
    external_formula: Option<&str>,
    defined_name: Option<&str>,
) -> Result<()> {
    let mut workbook = Workbook::new();
    {
        let sheet = workbook.add_worksheet();
        sheet.set_name("甲分机构本外币")?;
        sheet.write_string(0, 0, "机构")?;
        sheet.write_string(0, 1, "余额")?;
        for (row, name, amount) in [
            (1, "普通甲", 10.0),
            (
                2,
                if include_foreign {
                    "外资甲"
                } else {
                    "普通己"
                },
                20.0,
            ),
            (3, "普通乙", 30.0),
            (
                4,
                if include_foreign {
                    "外资乙"
                } else {
                    "普通庚"
                },
                40.0,
            ),
            (5, "普通丙", 50.0),
            (6, "尾行", 60.0),
        ] {
            sheet.write_string(row, 0, name)?;
            sheet.write_number(row, 1, amount)?;
        }
        sheet.write_formula(0, 7, "=SUM(B2:B7)")?;
        sheet.insert_note(2, 2, &Note::new("删除行批注"))?;
        sheet.insert_note(3, 2, &Note::new("保留行批注一"))?;
        sheet.insert_note(6, 2, &Note::new("保留行批注二"))?;
        sheet.set_row_height(6, 24.0)?;
        sheet.merge_range(8, 4, 9, 5, "合并区域", &rust_xlsxwriter::Format::new())?;
    }
    {
        let sheet = workbook.add_worksheet();
        sheet.set_name("乙分机构人民币")?;
        sheet.write_string(0, 0, "机构")?;
        sheet.write_string(
            1,
            0,
            if include_foreign {
                "外资乙"
            } else {
                "普通戊"
            },
        )?;
        sheet.write_string(2, 0, "普通丁")?;
    }
    {
        let sheet = workbook.add_worksheet();
        sheet.set_name("非目标工作表")?;
        sheet.write_string(0, 0, "机构")?;
        sheet.write_string(1, 0, "外资甲")?;
    }
    {
        let sheet = workbook.add_worksheet();
        sheet.set_name("公式引用")?;
        if let Some(formula) = external_formula {
            sheet.write_formula(0, 0, formula)?;
        } else {
            sheet.write_string(0, 0, "无外部公式")?;
        }
    }
    if let Some(formula) = defined_name {
        workbook.define_name("BranchValues", formula)?;
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
        platform: PLATFORM.to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: "synthetic-pt-7-3-foreign-row-removal".to_owned(),
            report_path: "pytools-core/tests/foreign_bank_batch_execution.rs".to_owned(),
            adapter_version: "umya-spreadsheet-3.1.1".to_owned(),
            engine_version: "umya-spreadsheet-3.1.1".to_owned(),
            platform: PLATFORM.to_owned(),
        }),
    }
}

fn cell_text<'a>(sheet: &'a excel_reader::SheetDump, address: &str) -> Option<&'a str> {
    sheet
        .values
        .get(address)
        .and_then(|value| value.get("value"))
        .and_then(Value::as_str)
}

#[test]
fn pt_7_3_batch_plan_and_execution_copy_without_overwriting_inputs_or_targets() -> Result<()> {
    let dir = TestDir::new();
    let config = dir.path("config.xlsx");
    let source = dir.path("foreign.xlsx");
    let no_change_source = dir.path("no-change.xlsx");
    let capability_path = dir.path("test-capability.json");
    let foreign_output = dir.path("foreign(金融局).xlsx");
    let no_change_output = dir.path("no-change(金融局).xlsx");
    make_config(&config)?;
    make_source(
        &source,
        true,
        Some("='甲分机构本外币'!B7"),
        Some("'甲分机构本外币'!$B$2:$B$7"),
    )?;
    make_source(&no_change_source, false, None, None)?;
    let source_bytes = fs::read(&source)?;
    let no_change_bytes = fs::read(&no_change_source)?;

    let planned = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["remove-foreign-bank", "--config"])
        .arg(&config)
        .arg(&source)
        .arg(&no_change_source)
        .output()?;
    let plan_report = assert_success(&planned)?;
    assert_eq!(plan_report["mode"], "plan_only");
    assert_eq!(plan_report["files"], 2);
    assert_eq!(plan_report["planned_files"], 2);
    assert_eq!(plan_report["planned_branch_sheets"], 4);
    assert_eq!(plan_report["planned_deleted_rows"], 3);
    assert_eq!(plan_report["failed_files"], 0);
    assert!(plan_report["items"][0]["plan"].is_object());
    assert!(plan_report["items"][1]["plan"].is_null());
    assert!(!foreign_output.exists());
    assert!(!no_change_output.exists());
    assert_eq!(fs::read(&source)?, source_bytes, "planning is read-only");
    assert_eq!(
        fs::read(&no_change_source)?,
        no_change_bytes,
        "planning is read-only"
    );

    let plan: ExcelEditPlan = serde_json::from_value(plan_report["items"][0]["plan"].clone())?;
    fs::write(
        &capability_path,
        serde_json::to_vec_pretty(&capability_for(&plan))?,
    )?;
    let executed = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["remove-foreign-bank", "--config"])
        .arg(&config)
        .args(["--capabilities"])
        .arg(&capability_path)
        .args(["--platform", PLATFORM])
        .arg(&source)
        .arg(&no_change_source)
        .output()?;
    let execution_report = assert_success(&executed)?;
    assert_eq!(execution_report["mode"], "execute");
    assert_eq!(execution_report["ok_files"], 2);
    assert_eq!(execution_report["failed_files"], 0);
    assert_eq!(execution_report["branch_sheets"], 4);
    assert_eq!(execution_report["deleted_rows"], 3);
    assert_eq!(
        execution_report["items"][0]["execution"]["engine"],
        "umya_spreadsheet"
    );
    assert_eq!(execution_report["items"][0]["status"], "written");
    assert_eq!(
        execution_report["items"][1]["execution"]["operation"],
        "copied"
    );
    assert_eq!(fs::read(&source)?, source_bytes, "source remains unchanged");
    assert_eq!(
        fs::read(&no_change_source)?,
        no_change_bytes,
        "source remains unchanged"
    );
    assert_eq!(
        fs::read(&no_change_output)?,
        no_change_bytes,
        "no-op output is a byte-identical copy"
    );

    let output = excel_reader::read_workbook(&foreign_output)?;
    let main_sheet = output
        .sheets
        .iter()
        .find(|sheet| sheet.name == "甲分机构本外币")
        .context("main branch sheet missing")?;
    let second_sheet = output
        .sheets
        .iter()
        .find(|sheet| sheet.name == "乙分机构人民币")
        .context("second branch sheet missing")?;
    let non_target = output
        .sheets
        .iter()
        .find(|sheet| sheet.name == "非目标工作表")
        .context("non-target sheet missing")?;
    assert_eq!(cell_text(main_sheet, "A2"), Some("普通甲"));
    assert_eq!(cell_text(main_sheet, "A3"), Some("普通乙"));
    assert_eq!(cell_text(main_sheet, "A4"), Some("普通丙"));
    assert_eq!(cell_text(main_sheet, "A5"), Some("尾行"));
    assert_eq!(cell_text(second_sheet, "A2"), Some("普通丁"));
    assert_eq!(cell_text(non_target, "A2"), Some("外资甲"));
    assert_eq!(
        main_sheet.formulas.get("H1").map(String::as_str),
        Some("SUM(B2:B5)")
    );
    assert_eq!(
        output.sheets[3].formulas.get("A1").map(String::as_str),
        Some("'甲分机构本外币'!B5"),
        "cross-sheet formula follows both deleted rows"
    );
    let output_umya = excel_editor::open_xlsx(&foreign_output)?;
    let branch_values = output_umya
        .defined_names()
        .iter()
        .chain(
            output_umya
                .sheet_collection()
                .iter()
                .flat_map(|sheet| sheet.defined_names()),
        )
        .find(|name| name.name() == "BranchValues")
        .context("BranchValues defined name missing")?;
    assert_eq!(
        branch_values.address(),
        "'甲分机构本外币'!$B$2:$B$5",
        "defined name follows both deleted rows while retaining absolute references"
    );
    assert!(!main_sheet.comments.contains_key("C2"));
    assert!(main_sheet
        .comments
        .get("C3")
        .is_some_and(|text| text.contains("保留行批注一")));
    assert!(main_sheet
        .comments
        .get("C5")
        .is_some_and(|text| text.contains("保留行批注二")));
    assert_eq!(main_sheet.merges, ["E7:F8"]);

    let foreign_output_bytes = fs::read(&foreign_output)?;
    let no_change_output_bytes = fs::read(&no_change_output)?;

    let collision_run = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["remove-foreign-bank", "--config"])
        .arg(&config)
        .args(["--capabilities"])
        .arg(&capability_path)
        .args(["--platform", PLATFORM])
        .arg(&source)
        .arg(&no_change_source)
        .output()?;
    let collision_report = assert_success(&collision_run)?;
    assert_eq!(collision_report["ok_files"], 0);
    assert_eq!(collision_report["failed_files"], 2);
    assert_eq!(
        fs::read(&source)?,
        source_bytes,
        "collision leaves source intact"
    );
    assert_eq!(
        fs::read(&foreign_output)?,
        foreign_output_bytes,
        "collision does not replace the existing output"
    );
    assert_eq!(fs::read(&no_change_output)?, no_change_output_bytes);
    Ok(())
}

#[test]
fn pt_7_3_refuses_dynamic_or_complex_cross_sheet_references_before_writing() -> Result<()> {
    for (case, external_formula, defined_name, expected_error) in [
        (
            "dynamic-formula",
            Some("=INDIRECT(\"'甲分机构本外币'!B7\")"),
            None,
            "dynamic INDIRECT",
        ),
        (
            "complex-defined-name",
            None,
            Some("=OFFSET('甲分机构本外币'!$B$2,0,0,COUNTA('甲分机构本外币'!$B:$B))"),
            "cannot be safely adjusted",
        ),
    ] {
        let dir = TestDir::new();
        let config = dir.path("config.xlsx");
        let source = dir.path(&format!("{case}.xlsx"));
        let capability_path = dir.path("test-capability.json");
        let output = dir.path(&format!("{case}(金融局).xlsx"));
        make_config(&config)?;
        make_source(&source, true, external_formula, defined_name)?;
        let source_bytes = fs::read(&source)?;

        let planned = Command::new(env!("CARGO_BIN_EXE_pytools"))
            .args(["remove-foreign-bank", "--config"])
            .arg(&config)
            .arg(&source)
            .output()?;
        let plan_report = assert_success(&planned)?;
        let plan: ExcelEditPlan = serde_json::from_value(plan_report["items"][0]["plan"].clone())?;
        fs::write(
            &capability_path,
            serde_json::to_vec_pretty(&capability_for(&plan))?,
        )?;

        let executed = Command::new(env!("CARGO_BIN_EXE_pytools"))
            .args(["remove-foreign-bank", "--config"])
            .arg(&config)
            .args(["--capabilities"])
            .arg(&capability_path)
            .args(["--platform", PLATFORM])
            .arg(&source)
            .output()?;
        let report = assert_success(&executed)?;
        assert_eq!(report["ok_files"], 0, "{case}");
        assert_eq!(report["failed_files"], 1, "{case}");
        assert!(report["items"][0]["error"]
            .as_str()
            .is_some_and(|message| message.contains(expected_error)));
        assert!(!output.exists(), "{case}: no candidate may be published");
        assert_eq!(fs::read(&source)?, source_bytes, "{case}: source unchanged");
    }
    Ok(())
}
