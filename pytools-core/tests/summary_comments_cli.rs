use anyhow::{Context, Result};
use excel_reader::read_workbook;
use serde_json::Value;
use std::{
    fs,
    path::{Path, PathBuf},
    process::{Command, Output},
    time::{SystemTime, UNIX_EPOCH},
};

#[test]
fn summary_comments_cli_logs_actual_stages_and_preserves_output_across_views() -> Result<()> {
    let template = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/summary_comments/comment_template.xlsx");
    let template_before = fs::read(&template)?;
    let root = unique_test_dir();
    fs::create_dir_all(&root)?;
    let source = root.join("comment_source.xlsx");
    create_source_workbook(&source)?;
    let source_before = fs::read(&source)?;

    let user = run_cli(&template, &[&source], &root.join("user"), None)?;
    assert!(
        user.status.success(),
        "user-view CLI failed: {}",
        String::from_utf8_lossy(&user.stderr)
    );
    let user_report = parse_stdout(&user)?;
    assert_eq!(user_report["saved"], true);
    assert_eq!(user_report["rows"], 3);
    assert_eq!(user_report["files"], 1);
    assert_eq!(user_report["execution"]["reader_engine"], "calamine");
    assert_eq!(user_report["execution"]["writer_engine"], "rust_xlsxwriter");
    assert!(user_report["execution"]["read_elapsed_ms"].is_number());
    assert!(user_report["execution"]["write_elapsed_ms"].is_number());
    assert!(user_report["execution"]["total_elapsed_ms"].is_number());
    let user_stderr = String::from_utf8_lossy(&user.stderr);
    assert!(user_stderr.contains("正在读取批注汇总模板"));
    assert!(user_stderr.contains("批注汇总工作簿已保存"));
    assert!(!user_stderr.contains("calamine"));
    assert!(!user_stderr.contains("rust_xlsxwriter"));
    let user_history = read_history(&user_report)?;
    assert_eq!(user_history["product"], "pytools");
    assert_eq!(user_history["feature_id"], "PT-6-2");
    assert_eq!(user_history["function"], "summary-comments");
    assert_eq!(user_history["status"], "succeeded");
    assert_eq!(user_history["completed_files"], 1);
    assert_eq!(user_history["failed_files"], 0);
    let user_history_text = fs::read_to_string(
        user_report["run_log_path"]
            .as_str()
            .context("user run log path missing")?,
    )?;
    for value in ["华东", "甲", "乙"] {
        assert!(
            !user_history_text.contains(value),
            "run history must not contain raw cell value {value}"
        );
    }

    let unsupported = root.join("ignored.xls");
    fs::write(&unsupported, b"unsupported input skipped by PyCode")?;
    let developer = run_cli(
        &template,
        &[&source, &unsupported],
        &root.join("developer"),
        Some("developer"),
    )?;
    assert!(
        developer.status.success(),
        "developer-view CLI failed: {}",
        String::from_utf8_lossy(&developer.stderr)
    );
    let developer_report = parse_stdout(&developer)?;
    let developer_history = read_history(&developer_report)?;
    let developer_events = String::from_utf8_lossy(&developer.stderr)
        .lines()
        .map(serde_json::from_str::<Value>)
        .collect::<std::result::Result<Vec<_>, _>>()
        .context("parse developer progress events")?;
    assert!(developer_events.len() >= 11);
    assert_eq!(developer_events.first().unwrap()["phase"], "run");
    assert_eq!(developer_events.first().unwrap()["status"], "started");
    assert_eq!(developer_events.last().unwrap()["status"], "succeeded");
    assert_eq!(
        developer_history["events"],
        serde_json::json!(developer_events)
    );
    for (index, event) in developer_events.iter().enumerate() {
        assert_eq!(event["sequence"], index + 1);
        assert!(event["total_elapsed_ms"].is_number());
    }
    for phase in ["template_read", "source_read", "row_extract", "output_save"] {
        assert!(
            developer_events.iter().any(|event| {
                event["phase"] == phase
                    && event["status"] == "succeeded"
                    && event["stage_elapsed_ms"].is_number()
            }),
            "completed {phase} event with elapsed time is missing"
        );
    }
    assert!(developer_events.iter().any(|event| {
        event["phase"] == "template_read"
            && event["actual_engine"] == "calamine + namespace-aware OOXML supplement"
            && event["components"].as_array().is_some_and(|components| {
                components
                    .iter()
                    .any(|component| component == "namespace-aware OOXML supplement")
                    && components
                        .iter()
                        .any(|component| component == "Calamine ISO date projection")
            })
    }));
    assert!(developer_events.iter().any(|event| {
        event["phase"] == "source_read"
            && event["actual_engine"] == "calamine + namespace-aware OOXML supplement"
    }));
    assert!(developer_events.iter().any(|event| {
        event["phase"] == "input_filter"
            && event["status"] == "skipped"
            && event["file_name"] == "ignored.xls"
            && event["actual_engine"].is_null()
    }));
    let save_event = developer_events
        .iter()
        .find(|event| {
            event["phase"] == "output_save"
                && event["status"] == "succeeded"
                && event["actual_engine"] == "rust_xlsxwriter"
        })
        .context("rust_xlsxwriter save event is missing")?;
    assert!(save_event["stage_elapsed_ms"].is_number());
    assert_eq!(
        developer_events.last().unwrap()["progress"]["processed_rows"],
        3
    );
    assert_eq!(
        developer_events.last().unwrap()["progress"]["total_rows"],
        3
    );
    assert_eq!(
        developer_events.last().unwrap()["progress"]["completed_files"],
        1
    );

    let off = run_cli(&template, &[&source], &root.join("off"), Some("off"))?;
    assert!(
        off.status.success(),
        "off-view CLI failed: {}",
        String::from_utf8_lossy(&off.stderr)
    );
    assert!(off.stderr.is_empty(), "off view must not emit logs");
    let off_report = parse_stdout(&off)?;
    assert!(off_report.get("run_id").is_none());
    assert!(off_report.get("run_log_path").is_none());
    assert!(!root.join("off/logs").exists());

    let user_output = output_workbook(&user_report)?;
    let developer_output = output_workbook(&developer_report)?;
    let off_output = output_workbook(&off_report)?;
    let expected_sheets = serde_json::to_value(&read_workbook(&user_output)?.sheets)?;
    assert_eq!(
        expected_sheets,
        serde_json::to_value(&read_workbook(&developer_output)?.sheets)?
    );
    assert_eq!(
        expected_sheets,
        serde_json::to_value(&read_workbook(&off_output)?.sheets)?
    );
    let summary_sheet = read_workbook(&user_output)?;
    let summary_sheet = summary_sheet
        .sheets
        .iter()
        .find(|sheet| sheet.name == "模板")
        .context("comment summary output sheet missing")?;
    assert_eq!(
        summary_sheet.values.get("A2"),
        Some(&serde_json::json!({"type":"string", "value":"comment_source"}))
    );
    assert_eq!(
        summary_sheet.values.get("B2"),
        Some(&serde_json::json!({"type":"string", "value":"Branch_A"}))
    );
    assert_eq!(
        summary_sheet.values.get("D2"),
        Some(&serde_json::json!({"type":"number", "value":10.5}))
    );
    assert_eq!(fs::read(&template)?, template_before, "template changed");
    assert_eq!(fs::read(&source)?, source_before, "source workbook changed");

    fs::remove_dir_all(root)?;
    Ok(())
}

#[test]
fn summary_comments_cli_uses_merge_anchor_for_non_anchor_cells() -> Result<()> {
    let template = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/summary_comments/comment_template.xlsx");
    let source = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/summary_comments/merged_children_source.xlsx");
    let source_before = fs::read(&source)?;
    let source_dump = read_workbook(&source)?;
    let source_sheet = source_dump
        .sheets
        .iter()
        .find(|sheet| sheet.name == "Branch_A")
        .context("merged source sheet missing")?;
    assert_eq!(
        source_sheet.values.get("B2"),
        Some(&serde_json::json!({
            "type": "string",
            "value": "raw merged child residue"
        })),
        "fixture must reproduce a raw reader value in a merged child cell"
    );

    let root = unique_test_dir();
    fs::create_dir_all(&root)?;
    let output = run_cli(&template, &[&source], &root.join("run"), None)?;
    assert!(
        output.status.success(),
        "summary-comments CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report = parse_stdout(&output)?;
    assert_eq!(report["rows"], 3);
    let summary = read_workbook(&output_workbook(&report)?)?;
    let summary_sheet = summary
        .sheets
        .iter()
        .find(|sheet| sheet.name == "模板")
        .context("comment summary output sheet missing")?;
    // PyCode's _gcv maps every coordinate in a merge to its top-left anchor.
    for address in ["D2", "E2", "D3", "E3"] {
        assert_eq!(
            summary_sheet.values.get(address),
            Some(&serde_json::json!({
                "type": "string",
                "value": "merged anchor"
            })),
            "merged source cell {address} must use the anchor value"
        );
    }
    assert_eq!(fs::read(&source)?, source_before, "source workbook changed");

    fs::remove_dir_all(root)?;
    Ok(())
}

#[test]
fn summary_comments_cli_preserves_formula_and_blank_template_headers() -> Result<()> {
    let root = unique_test_dir();
    fs::create_dir_all(&root)?;
    let template = root.join("formula_header_template.xlsx");
    create_formula_header_template(&template)?;
    let template_dump = read_workbook(&template)?;
    anyhow::ensure!(
        template_dump.sheets[0].formulas.contains_key("B1"),
        "formula fixture was not read: {:?}",
        template_dump.sheets[0].formulas
    );
    let source = root.join("source.xlsx");
    create_source_workbook(&source)?;
    let source_before = fs::read(&source)?;

    let output = run_cli(&template, &[&source], &root.join("run"), Some("off"))?;
    assert!(
        output.status.success(),
        "summary-comments CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report = parse_stdout(&output)?;
    assert_eq!(report["saved"], true);
    assert_eq!(report["rows"], 3);

    let summary = read_workbook(output_workbook(&report)?)?;
    let sheet = summary
        .sheets
        .first()
        .context("formula-header summary sheet missing")?;
    let formula = sheet.formulas.get("C1").with_context(|| {
        format!(
            "formula header was not written: formulas={:?}, C1={:?}, D1={:?}",
            sheet.formulas,
            sheet.values.get("C1"),
            sheet.values.get("D1")
        )
    })?;
    assert!(
        formula.trim_start_matches('=').contains("余额"),
        "unexpected formula header: {formula}"
    );
    assert_eq!(
        sheet.values.get("D1").and_then(|value| value.get("value")),
        Some(&serde_json::json!("C")),
        "blank template header should fall back to its column letter"
    );
    assert_eq!(
        sheet.values.get("E1").and_then(|value| value.get("value")),
        Some(&serde_json::json!("True")),
        "boolean template header should use Python's string capitalization"
    );
    assert_eq!(
        sheet.values.get("F1").and_then(|value| value.get("value")),
        Some(&serde_json::json!("10")),
        "numeric template header should use Python's string representation"
    );
    assert_eq!(
        sheet.values.get("G1").and_then(|value| value.get("value")),
        Some(&serde_json::json!("2026-06-30 00:00:00")),
        "date template header should use Python's datetime string representation"
    );
    for (address, expected) in [("H1", "1e-05"), ("I1", "1e+20"), ("J1", "0.0001")] {
        assert_eq!(
            sheet
                .values
                .get(address)
                .and_then(|value| value.get("value")),
            Some(&serde_json::json!(expected)),
            "numeric template header at {address} should use Python's float string representation"
        );
    }
    assert_eq!(fs::read(&source)?, source_before, "source workbook changed");

    fs::remove_dir_all(root)?;
    Ok(())
}

#[test]
fn summary_comments_cli_ignores_empty_set_marker_like_pycode() -> Result<()> {
    let root = unique_test_dir();
    fs::create_dir_all(&root)?;
    let template = root.join("empty_set_template.xlsx");
    create_empty_set_template(&template)?;
    let template_before = fs::read(&template)?;
    let source = root.join("source.xlsx");
    create_source_workbook(&source)?;
    let source_before = fs::read(&source)?;

    let output = run_cli(&template, &[&source], &root.join("run"), None)?;
    assert!(
        output.status.success(),
        "summary-comments CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report = parse_stdout(&output)?;
    assert_eq!(report["saved"], true);
    assert_eq!(report["rows"], 2);

    let summary = read_workbook(&output_workbook(&report)?)?;
    let sheet = summary
        .sheets
        .iter()
        .find(|sheet| sheet.name == "模板")
        .context("comment summary output sheet missing")?;
    for (address, value) in [
        ("A1", "工作簿"),
        ("B1", "工作表"),
        ("C1", "余额"),
        ("D1", "余额_2"),
        ("E1", "行号"),
    ] {
        assert_eq!(
            sheet.values.get(address),
            Some(&serde_json::json!({"type":"string", "value":value})),
            "unexpected header at {address}"
        );
    }
    assert!(
        !sheet.values.contains_key("F1"),
        "empty set() marker must not add an output column"
    );
    assert_eq!(
        sheet.values.get("C2"),
        Some(&serde_json::json!({"type":"number", "value":10.5}))
    );
    assert_eq!(
        sheet.values.get("D2"),
        Some(&serde_json::json!({"type":"string", "value":"甲"}))
    );
    assert_eq!(fs::read(&template)?, template_before, "template changed");
    assert_eq!(fs::read(&source)?, source_before, "source workbook changed");

    fs::remove_dir_all(root)?;
    Ok(())
}

#[test]
fn summary_comments_cli_parses_set_marker_after_unicode_case_mapping() -> Result<()> {
    let root = unique_test_dir();
    fs::create_dir_all(&root)?;
    let template = root.join("unicode_set_template.xlsx");
    create_unicode_set_template(&template)?;
    let template_before = fs::read(&template)?;
    let source = root.join("source.xlsx");
    create_source_workbook(&source)?;
    let source_before = fs::read(&source)?;

    let output = run_cli(&template, &[&source], &root.join("run"), Some("off"))?;
    assert!(
        output.status.success(),
        "summary-comments CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report = parse_stdout(&output)?;
    assert_eq!(report["saved"], true);
    assert_eq!(report["rows"], 2);

    let summary = read_workbook(&output_workbook(&report)?)?;
    let sheet = summary
        .sheets
        .iter()
        .find(|sheet| sheet.name == "模板")
        .context("Unicode set marker summary sheet missing")?;
    assert_eq!(
        sheet.values.get("C1"),
        Some(&serde_json::json!({"type":"string", "value":"区域"}))
    );
    assert_eq!(
        sheet.values.get("C2"),
        Some(&serde_json::json!({"type":"string", "value":"华东"}))
    );
    assert_eq!(
        sheet.values.get("C3"),
        Some(&serde_json::json!({"type":"string", "value":"华东"}))
    );
    assert_eq!(fs::read(&template)?, template_before, "template changed");
    assert_eq!(fs::read(&source)?, source_before, "source workbook changed");

    fs::remove_dir_all(root)?;
    Ok(())
}

#[test]
fn summary_comments_cli_preserves_excel_errors_and_textualizes_unknown_error_tokens() -> Result<()>
{
    let fixtures = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/summary_comments");
    let template = fixtures.join("error_tokens_template.xlsx");
    let source = fixtures.join("error_tokens_source.xlsx");
    let template_before = fs::read(&template)?;
    let source_before = fs::read(&source)?;
    let root = unique_test_dir();
    fs::create_dir_all(&root)?;

    let output = run_cli(&template, &[&source], &root, Some("developer"))?;
    assert!(
        output.status.success(),
        "summary-comments CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report = parse_stdout(&output)?;
    assert_eq!(report["saved"], true);
    assert_eq!(report["rows"], 6);
    assert_eq!(
        report["execution"]["reader_engine"],
        "calamine + OOXML error-token normalization"
    );
    assert_eq!(
        report["execution"]["writer_engine"],
        "rust_xlsxwriter + OOXML Patch"
    );
    assert_eq!(
        report["execution"]["unknown_error_token_cells_normalized"],
        2
    );
    assert_eq!(report["execution"]["ooxml_error_cells_patched"], 2);

    let output_path = output_workbook(&report)?;
    let workbook = read_workbook(&output_path)?;
    for sheet_name in ["模板", "Branch_B"] {
        let sheet = workbook
            .sheets
            .iter()
            .find(|sheet| sheet.name == sheet_name)
            .with_context(|| format!("summary output worksheet {sheet_name:?} is missing"))?;
        assert_eq!(
            sheet.values.get("D2"),
            Some(&serde_json::json!({"type":"error", "value":"Div0"})),
            "known Excel errors must remain error-typed on {sheet_name}"
        );
        assert_eq!(
            sheet.values.get("E2"),
            Some(&serde_json::json!({"type":"string", "value":"#DATA!"})),
            "unsupported error tokens must match PyCode text behavior on {sheet_name}"
        );
    }

    let history = read_history(&report)?;
    assert!(history["events"].as_array().unwrap().iter().any(|event| {
        event["phase"] == "source_read_compatibility"
            && event["status"] == "succeeded"
            && event["actual_engine"] == "calamine + OOXML error-token normalization"
    }));
    assert!(history["events"].as_array().unwrap().iter().any(|event| {
        event["phase"] == "ooxml_error_patch"
            && event["status"] == "succeeded"
            && event["actual_engine"] == "OOXML Patch"
    }));
    let history_text = fs::read_to_string(
        report["run_log_path"]
            .as_str()
            .context("run log path missing")?,
    )?;
    assert!(!history_text.contains("#DATA!"));
    assert!(!history_text.contains("#DIV/0!"));
    assert_eq!(fs::read(&template)?, template_before, "template changed");
    assert_eq!(fs::read(&source)?, source_before, "source workbook changed");

    fs::remove_dir_all(root)?;
    Ok(())
}

#[test]
fn summary_comments_cli_records_template_validation_failure() -> Result<()> {
    let root = unique_test_dir();
    fs::create_dir_all(&root)?;
    let invalid_template = root.join("template.txt");
    fs::write(&invalid_template, b"not a workbook")?;
    let source = root.join("source.xlsx");
    create_source_workbook(&source)?;

    let output = run_cli(
        &invalid_template,
        &[&source],
        &root.join("failure"),
        Some("developer"),
    )?;
    assert!(!output.status.success());
    let events = String::from_utf8_lossy(&output.stderr)
        .lines()
        .filter_map(|line| serde_json::from_str::<Value>(line).ok())
        .collect::<Vec<_>>();
    anyhow::ensure!(
        !events.is_empty(),
        "failed-run developer events are missing"
    );
    assert_eq!(events.last().unwrap()["phase"], "template_validation");
    assert_eq!(events.last().unwrap()["status"], "failed");
    assert_eq!(
        events.last().unwrap()["error_type"],
        "template_format_error"
    );
    let logs = fs::read_dir(root.join("failure/logs"))?
        .map(|entry| entry.map(|entry| entry.path()))
        .collect::<std::io::Result<Vec<_>>>()?;
    assert_eq!(logs.len(), 1);
    let history: Value = serde_json::from_slice(&fs::read(&logs[0])?)?;
    assert_eq!(history["feature_id"], "PT-6-2");
    assert_eq!(history["status"], "failed");
    assert_eq!(history["failure"]["stage"], "template_validation");
    assert_eq!(history["failure"]["error_type"], "template_format_error");

    fs::remove_dir_all(root)?;
    Ok(())
}

fn run_cli(
    template: &Path,
    sources: &[&Path],
    root: &Path,
    log_view: Option<&str>,
) -> Result<Output> {
    let mut command = Command::new(env!("CARGO_BIN_EXE_pytools"));
    command
        .arg("summary-comments")
        .arg("--template")
        .arg(template)
        .arg("--output-dir")
        .arg(root.join("output"))
        .arg("--log-dir")
        .arg(root.join("logs"));
    if let Some(log_view) = log_view {
        command.arg("--log-view").arg(log_view);
    }
    Ok(command
        .args(sources)
        .output()
        .context("run summary-comments CLI")?)
}

fn parse_stdout(output: &Output) -> Result<Value> {
    serde_json::from_slice(&output.stdout).context("parse summary-comments CLI JSON")
}

fn read_history(report: &Value) -> Result<Value> {
    let path = report["run_log_path"]
        .as_str()
        .context("run log path missing from CLI output")?;
    serde_json::from_slice(&fs::read(path)?).context("parse PT-6-2 run history")
}

fn output_workbook(report: &Value) -> Result<PathBuf> {
    Ok(PathBuf::from(
        report["path"]
            .as_str()
            .context("summary output path missing from CLI output")?,
    ))
}

fn create_source_workbook(path: &Path) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    {
        let worksheet = workbook.add_worksheet();
        worksheet.set_name("Branch_A")?;
        worksheet.write_string(0, 1, "余额")?;
        worksheet.write_string(0, 2, "余额")?;
        worksheet.write_string(0, 3, "区域")?;
        worksheet.write_number(1, 1, 10.5)?;
        worksheet.write_string(1, 2, "甲")?;
        worksheet.write_number(2, 1, 20.0)?;
        worksheet.write_string(2, 2, "乙")?;
        worksheet.write_number(3, 1, 0.0)?;
        worksheet.write_string(3, 2, "丙")?;
        worksheet.write_string(0, 3, "华东")?;
    }
    workbook
        .save(path)
        .context("save synthetic source workbook")
}

fn create_empty_set_template(path: &Path) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    {
        let worksheet = workbook.add_worksheet();
        worksheet.set_name("模板")?;
        worksheet.write_string(1, 0, "行开始")?;
        worksheet.write_string(2, 0, "行结束")?;
        worksheet.write_string(0, 1, "余额")?;
        worksheet.write_string(0, 2, "余额")?;
        worksheet.write_string(0, 3, "空集合")?;
        worksheet.insert_note(1, 0, &rust_xlsxwriter::Note::new("行区域1"))?;
        worksheet.insert_note(2, 0, &rust_xlsxwriter::Note::new("行区域#1"))?;
        worksheet.insert_note(0, 1, &rust_xlsxwriter::Note::new("列区域1"))?;
        worksheet.insert_note(0, 2, &rust_xlsxwriter::Note::new("列区域#1"))?;
        worksheet.insert_note(0, 3, &rust_xlsxwriter::Note::new("set()"))?;
    }
    workbook
        .save(path)
        .context("save empty-set comment template")
}

fn create_unicode_set_template(path: &Path) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    {
        let worksheet = workbook.add_worksheet();
        worksheet.set_name("模板")?;
        worksheet.write_string(1, 0, "行开始")?;
        worksheet.write_string(2, 0, "行结束")?;
        worksheet.write_string(0, 1, "余额")?;
        worksheet.write_string(0, 2, "余额")?;
        worksheet.write_string(0, 3, "区域")?;
        worksheet.insert_note(1, 0, &rust_xlsxwriter::Note::new("行区域1"))?;
        worksheet.insert_note(2, 0, &rust_xlsxwriter::Note::new("行区域#1"))?;
        worksheet.insert_note(0, 1, &rust_xlsxwriter::Note::new("列区域1"))?;
        worksheet.insert_note(0, 2, &rust_xlsxwriter::Note::new("列区域#1"))?;
        worksheet.insert_note(0, 3, &rust_xlsxwriter::Note::new("İİset(区域)"))?;
    }
    workbook
        .save(path)
        .context("save Unicode set-marker comment template")
}

fn create_formula_header_template(path: &Path) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    {
        let worksheet = workbook.add_worksheet();
        worksheet.set_name("Branch_A")?;
        worksheet.write_string(1, 0, "行开始")?;
        worksheet.write_string(3, 0, "行结束")?;
        worksheet.write_formula(0, 1, "=\"余额\"")?;
        worksheet.write_boolean(0, 3, true)?;
        worksheet.write_number(0, 4, 10.0)?;
        let date = rust_xlsxwriter::ExcelDateTime::from_ymd(2026, 6, 30)?;
        let date_format = rust_xlsxwriter::Format::new().set_num_format("yyyy-mm-dd");
        worksheet.write_datetime_with_format(0, 5, &date, &date_format)?;
        worksheet.write_number(0, 6, 1e-5)?;
        worksheet.write_number(0, 7, 1e20)?;
        worksheet.write_number(0, 8, 1e-4)?;
        worksheet.insert_note(1, 0, &rust_xlsxwriter::Note::new("行区域1"))?;
        worksheet.insert_note(3, 0, &rust_xlsxwriter::Note::new("行区域#1"))?;
        worksheet.insert_note(0, 1, &rust_xlsxwriter::Note::new("列区域1"))?;
        worksheet.insert_note(0, 2, &rust_xlsxwriter::Note::new("列区域#1"))?;
        worksheet.insert_note(0, 3, &rust_xlsxwriter::Note::new("列区域2"))?;
        worksheet.insert_note(1, 3, &rust_xlsxwriter::Note::new("列区域#2"))?;
        worksheet.insert_note(0, 4, &rust_xlsxwriter::Note::new("列区域3"))?;
        worksheet.insert_note(1, 4, &rust_xlsxwriter::Note::new("列区域#3"))?;
        worksheet.insert_note(0, 5, &rust_xlsxwriter::Note::new("列区域4"))?;
        worksheet.insert_note(1, 5, &rust_xlsxwriter::Note::new("列区域#4"))?;
        for (column, marker) in [(6, 5), (7, 6), (8, 7)] {
            worksheet.insert_note(
                0,
                column,
                &rust_xlsxwriter::Note::new(format!("列区域{marker}")),
            )?;
            worksheet.insert_note(
                1,
                column,
                &rust_xlsxwriter::Note::new(format!("列区域#{marker}")),
            )?;
        }
    }
    workbook
        .save(path)
        .context("save formula-header comment template")
}

fn unique_test_dir() -> PathBuf {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system clock after Unix epoch")
        .as_nanos();
    std::env::temp_dir().join(format!(
        "pytools-summary-comments-{}-{nonce}",
        std::process::id()
    ))
}
