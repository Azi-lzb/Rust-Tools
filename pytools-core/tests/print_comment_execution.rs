use anyhow::{Context, Result};
use excel_editor::{
    BackendCapability, CapabilityEvidence, CapabilityStatus, EditOperation, EngineId, ExcelEditPlan,
};
use excel_reader::read_workbook;
use rust_xlsxwriter::{Format, Note, Workbook};
use serde_json::Value;
use std::{
    fs,
    path::{Path, PathBuf},
    process::{Command, Output},
};

const PLATFORM: &str = "synthetic-pt-3-1-3-2-comment-print-test";

struct TestDir(PathBuf);

impl TestDir {
    fn new() -> Self {
        let path = std::env::temp_dir().join(format!(
            "pytools-print-comment-{}-{}",
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
    sheet.set_name("Marked")?;
    sheet.write_string(0, 0, "Report")?;
    sheet.insert_note(0, 0, &Note::new("打印输出 横向"))?;

    sheet.write_number_with_format(1, 1, 0.0, &Format::new().set_bold())?;
    sheet.write_string(1, 2, "A")?;
    sheet.write_string(1, 3, "B")?;
    sheet.write_number(2, 1, 2.0)?;
    sheet.merge_range(2, 2, 2, 3, "merged", &Format::new())?;
    sheet.write_number(3, 1, 3.0)?;
    sheet.write_string(3, 2, "E")?;
    sheet.write_string(3, 3, "F")?;
    sheet.insert_note(1, 1, &Note::new("打印区域 1"))?;
    sheet.insert_note(3, 3, &Note::new("打印区域#1"))?;

    sheet.write_string(1, 5, "second")?;
    sheet.write_string(2, 6, "block")?;
    sheet.insert_note(1, 5, &Note::new("打印区域 2"))?;
    sheet.insert_note(2, 6, &Note::new("打印区域#2"))?;
    sheet.set_row_height(1, 22.0)?;
    sheet.set_column_width(1, 14.0)?;
    workbook.add_worksheet().set_name("Plain")?;
    workbook.save(path)?;
    Ok(())
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
            sample_id: "synthetic-pt-3-1-3-2-comment-print".to_owned(),
            report_path: "pytools-core/tests/print_comment_execution.rs".to_owned(),
            adapter_version: "umya-spreadsheet-3.1.1".to_owned(),
            engine_version: "umya-spreadsheet-3.1.1".to_owned(),
            platform: PLATFORM.to_owned(),
        }),
    }
}

fn success_json(output: &Output) -> Result<Value> {
    anyhow::ensure!(
        output.status.success(),
        "pytools failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    Ok(serde_json::from_slice(&output.stdout)?)
}

fn run_print_candidate(
    source: &Path,
    output: &Path,
    capability_path: &Path,
    drop_comments: bool,
    hide_zero: bool,
) -> Result<Value> {
    let args = [
        "print-comment-plan",
        "--input",
        source.to_str().context("source path is not UTF-8")?,
        "--output",
        output.to_str().context("output path is not UTF-8")?,
        "--drop-comments",
        if drop_comments { "true" } else { "false" },
        "--hide-zero",
        if hide_zero { "true" } else { "false" },
        "--capabilities",
        capability_path
            .to_str()
            .context("capability path is not UTF-8")?,
        "--platform",
        PLATFORM,
    ];
    let result = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(args)
        .output()?;
    success_json(&result)
}

fn value_number(book: &excel_reader::WorkbookDump, sheet: &str, cell: &str) -> Option<f64> {
    book.sheets
        .iter()
        .find(|candidate| candidate.name == sheet)?
        .values
        .get(cell)?
        .get("value")?
        .as_f64()
}

fn style_bold(book: &excel_editor::Workbook, sheet: &str, cell: &str) -> Result<bool> {
    Ok(book
        .sheet_by_name(sheet)
        .with_context(|| format!("worksheet {sheet:?} missing"))?
        .cell(cell)
        .and_then(|value| value.style().font())
        .is_some_and(|font| font.bold()))
}

#[test]
fn pt_3_1_and_pt_3_2_candidates_execute_through_excel_modifier() -> Result<()> {
    let dir = TestDir::new();
    let source = dir.path("source.xlsx");
    make_source(&source)?;
    let source_before = fs::read(&source)?;

    for (name, drop_comments, hide_zero) in [("keep", false, false), ("fast", true, true)] {
        let output = dir.path(&format!("{name}.xlsx"));
        let plan_output = dir.path(&format!("{name}.plan.json"));
        let capability_path = dir.path(&format!("{name}.capability.json"));
        let planned = Command::new(env!("CARGO_BIN_EXE_pytools"))
            .args(["print-comment-plan", "--input"])
            .arg(&source)
            .args(["--output"])
            .arg(&output)
            .args([
                "--drop-comments",
                if drop_comments { "true" } else { "false" },
            ])
            .args(["--hide-zero", if hide_zero { "true" } else { "false" }])
            .args(["--plan-output"])
            .arg(&plan_output)
            .output()?;
        let planned_json = success_json(&planned)?;
        assert_eq!(planned_json["execution"], "plan_only");
        assert_eq!(planned_json["matching_sheets"], 1);
        assert_eq!(planned_json["marked_ranges"], 2);
        assert!(!output.exists(), "plan-only must not create {name} output");

        let plan: ExcelEditPlan = serde_json::from_slice(&fs::read(&plan_output)?)?;
        fs::write(
            &capability_path,
            serde_json::to_vec_pretty(&capability_for(&plan))?,
        )?;
        let executed =
            run_print_candidate(&source, &output, &capability_path, drop_comments, hide_zero)?;
        assert_eq!(executed["execution"]["engine"], "umya_spreadsheet");
        assert!(output.is_file(), "{name} output was not created");

        let dump = read_workbook(&output)?;
        assert_eq!(
            dump.sheets
                .iter()
                .map(|sheet| sheet.name.as_str())
                .collect::<Vec<_>>(),
            ["Marked_1", "Marked_2"]
        );
        assert_eq!(dump.sheets[0].merges, ["B2:C2"]);
        assert!(dump.sheets[1].merges.is_empty());
        if drop_comments {
            assert!(dump.sheets.iter().all(|sheet| sheet.comments.is_empty()));
            assert!(dump.sheets[0]
                .values
                .get("A1")
                .is_none_or(|value| value["value"].is_null()));
        } else {
            assert!(dump.sheets[0]
                .comments
                .get("A1")
                .is_some_and(|comment| comment.contains("打印区域 1")));
            assert_eq!(value_number(&dump, "Marked_1", "A1"), Some(0.0));
        }
        assert_eq!(value_number(&dump, "Marked_1", "A2"), Some(2.0));
        assert_eq!(dump.sheets[1].values["A1"]["value"], "second");

        let output_book = excel_editor::open_xlsx(&output)?;
        assert!(style_bold(&output_book, "Marked_1", "A1")?);
        for (index, (sheet_name, expected_area)) in [
            ("Marked_1", "'Marked_1'!A1:C3"),
            ("Marked_2", "'Marked_2'!A1:B2"),
        ]
        .into_iter()
        .enumerate()
        {
            let sheet = output_book.sheet_by_name(sheet_name).with_context(|| {
                format!("worksheet {sheet_name:?} missing for print-area check")
            })?;
            let print_area = sheet
                .defined_names()
                .iter()
                .find(|name| name.name().eq_ignore_ascii_case("_xlnm.Print_Area"))
                .with_context(|| format!("print area missing on {sheet_name:?}"))?;
            assert_eq!(print_area.local_sheet_id(), index as u32);
            assert_eq!(print_area.address(), expected_area);
        }
        let first = output_book
            .sheet_by_name("Marked_1")
            .context("Marked_1 missing after Umya readback")?;
        assert_eq!(first.page_setup().fit_to_width(), 1);
        assert_eq!(first.page_setup().fit_to_height(), 0);
        assert!(first
            .header_footer()
            .odd_footer()
            .value()
            .contains("第 1 页 / 共 2 页"));
        if !drop_comments {
            assert!(first
                .row_dimension(1)
                .is_some_and(|row| row.height() > 21.0));
            assert!(first
                .column_dimension_by_number(1)
                .is_some_and(|column| column.width() > 0.0));
        }
        assert_eq!(fs::read(&source)?, source_before);
    }
    Ok(())
}
