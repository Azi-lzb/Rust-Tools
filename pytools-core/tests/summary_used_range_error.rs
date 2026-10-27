use anyhow::{Context, Result};
use serde_json::Value;
use std::{
    fs,
    io::Read,
    path::{Path, PathBuf},
    process::Command,
    sync::atomic::{AtomicU64, Ordering},
    time::{SystemTime, UNIX_EPOCH},
};

static NEXT_TEST_DIRECTORY_ID: AtomicU64 = AtomicU64::new(0);

#[test]
fn summary_used_range_cli_preserves_excel_error_cells() -> Result<()> {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/summary_used_range/error_value.xlsx");
    let original = fs::read(&fixture)?;
    let output_dir = unique_test_dir();
    let unsupported_dir = output_dir.join("unsupported_inputs");
    fs::create_dir_all(&unsupported_dir)?;
    let unsupported_xls = unsupported_dir.join("ignored.xls");
    let unsupported_xlsb = unsupported_dir.join("ignored.xlsb");
    fs::write(
        &unsupported_xls,
        b"not an Excel workbook; PyCode skips .xls",
    )?;
    fs::write(
        &unsupported_xlsb,
        b"not an Excel workbook; PyCode skips .xlsb",
    )?;
    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("summary-used-range")
        .arg("--output-dir")
        .arg(output_dir.join("output"))
        .arg("--log-dir")
        .arg(output_dir.join("logs"))
        .arg(&fixture)
        .arg(&unsupported_xls)
        .arg(&unsupported_xlsb)
        .output()
        .context("run pytools summary-used-range")?;
    assert!(
        output.status.success(),
        "CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report: Value = serde_json::from_slice(&output.stdout).context("parse CLI JSON")?;
    assert_eq!(report["saved"], true);
    assert_eq!(report["rows"], 2);
    assert_eq!(report["files"], 1);
    assert!(report["run_log_path"].is_string());
    assert_eq!(report["execution"]["reader_engine"], "calamine");
    assert_eq!(report["execution"]["writer_engine"], "rust_xlsxwriter");
    assert_eq!(report["execution"]["ooxml_error_cells_patched"], 7);
    assert!(report["execution"]["total_elapsed_ms"].is_number());
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("正在扫描工作簿范围"));
    assert!(stderr.contains("工作簿读取完成"));
    assert!(stderr.contains("汇总工作簿生成完成"));
    assert!(stderr.contains("忽略不支持的输入"));
    assert!(!stderr.contains("calamine"));
    assert!(!stderr.contains("rust_xlsxwriter"));

    let run_log_path = PathBuf::from(
        report["run_log_path"]
            .as_str()
            .context("CLI run log path must be a string")?,
    );
    let run_log_bytes = fs::read(&run_log_path)?;
    let run_log_text = String::from_utf8_lossy(&run_log_bytes);
    let run_log: Value = serde_json::from_slice(&run_log_bytes).context("parse PT-6-1 run log")?;
    assert_eq!(run_log["product"], "pytools");
    assert_eq!(run_log["feature_id"], "PT-6-1");
    assert_eq!(run_log["status"], "succeeded");
    assert_eq!(run_log["total_inputs"], 3);
    assert_eq!(run_log["events"][0]["progress"]["total_files"], 1);
    assert!(run_log["events"]
        .as_array()
        .is_some_and(|events| !events.is_empty()));
    for extension in ["xls", "xlsb"] {
        assert!(run_log["events"].as_array().unwrap().iter().any(|event| {
            event["phase"] == "input_filter"
                && event["status"] == "skipped"
                && event["file_name"] == format!("ignored.{extension}")
                && event["actual_engine"].is_null()
        }));
    }
    assert!(run_log_text.contains("calamine"));
    assert!(run_log_text.contains("rust_xlsxwriter"));
    assert!(!run_log_text.contains("#DATA!"));

    let developer_output_dir = output_dir.join("developer");
    let developer = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("summary-used-range")
        .arg("--output-dir")
        .arg(developer_output_dir.join("output"))
        .arg("--log-dir")
        .arg(developer_output_dir.join("logs"))
        .arg("--log-view")
        .arg("developer")
        .arg(&fixture)
        .arg(&unsupported_xls)
        .arg(&unsupported_xlsb)
        .output()
        .context("run pytools summary-used-range in developer log view")?;
    assert!(
        developer.status.success(),
        "developer CLI failed: {}",
        String::from_utf8_lossy(&developer.stderr)
    );
    let developer_report: Value =
        serde_json::from_slice(&developer.stdout).context("parse developer CLI JSON")?;
    let developer_stderr = String::from_utf8_lossy(&developer.stderr);
    let developer_events = developer_stderr
        .lines()
        .map(serde_json::from_str::<Value>)
        .collect::<std::result::Result<Vec<_>, _>>()
        .context("parse developer progress events")?;
    assert!(developer_events.len() >= 8);
    assert_eq!(developer_events.first().unwrap()["phase"], "run");
    assert_eq!(developer_events.first().unwrap()["status"], "started");
    assert_eq!(developer_events.last().unwrap()["status"], "succeeded");
    assert!(developer_events.iter().any(|event| {
        event["phase"] == "value_read"
            && event["actual_engine"] == "calamine"
            && event["stage_elapsed_ms"].is_number()
    }));
    assert!(developer_events.iter().any(|event| {
        event["phase"] == "output_save"
            && event["actual_engine"] == "rust_xlsxwriter"
            && event["stage_elapsed_ms"].is_number()
    }));
    let workbook_write_event = developer_events
        .iter()
        .find(|event| {
            event["phase"] == "output_workbook"
                && event["status"] == "succeeded"
                && event["actual_engine"] == "rust_xlsxwriter"
                && event["stage_elapsed_ms"].is_number()
        })
        .context("completed output workbook event missing")?;
    assert_eq!(
        developer_report["execution"]["write_elapsed_ms"],
        workbook_write_event["stage_elapsed_ms"]
    );
    assert!(developer_events
        .iter()
        .filter(|event| { event["phase"] == "input_filter" && event["status"] == "skipped" })
        .all(|event| event["actual_engine"].is_null()));
    for (index, event) in developer_events.iter().enumerate() {
        assert_eq!(event["sequence"], index + 1);
    }
    let developer_log_path = PathBuf::from(
        developer_report["run_log_path"]
            .as_str()
            .context("developer CLI run log path must be a string")?,
    );
    let developer_log: Value = serde_json::from_slice(&fs::read(developer_log_path)?)
        .context("parse developer PT-6-1 run log")?;
    assert_eq!(
        developer_log["events"].as_array().unwrap().len(),
        developer_events.len()
    );

    let quiet_output_dir = output_dir.join("quiet");
    let quiet = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("summary-used-range")
        .arg("--output-dir")
        .arg(quiet_output_dir.join("output"))
        .arg("--log-dir")
        .arg(quiet_output_dir.join("logs"))
        .arg("--log-view")
        .arg("off")
        .arg(&fixture)
        .arg(&unsupported_xls)
        .arg(&unsupported_xlsb)
        .output()
        .context("run pytools summary-used-range with logging disabled")?;
    assert!(
        quiet.status.success(),
        "quiet CLI failed: {}",
        String::from_utf8_lossy(&quiet.stderr)
    );
    assert!(
        quiet.stderr.is_empty(),
        "off view must not emit progress logs"
    );
    let quiet_report: Value =
        serde_json::from_slice(&quiet.stdout).context("parse quiet CLI JSON")?;
    assert!(quiet_report.get("run_id").is_none());
    assert!(quiet_report.get("run_log_path").is_none());
    assert!(!quiet_output_dir.join("logs").exists());

    let path = PathBuf::from(
        report["path"]
            .as_str()
            .context("CLI report path must be a string")?,
    );
    let workbook = excel_reader::read_workbook(&path)?;
    let sheet = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "汇总")
        .context("summary sheet missing")?;
    let expected_errors = [
        ("D2", "Div0"),
        ("E2", "NA"),
        ("F2", "Name"),
        ("G2", "Null"),
        ("H2", "Num"),
        ("I2", "Ref"),
        ("J2", "Value"),
    ];
    for (address, error) in expected_errors {
        assert_eq!(
            sheet.values.get(address),
            Some(&serde_json::json!({"type":"error", "value":error}))
        );
    }
    assert_eq!(
        sheet.values.get("A3"),
        Some(&serde_json::json!({"type":"string", "value":"error_value"}))
    );
    assert_eq!(
        sheet.values.get("B3"),
        Some(&serde_json::json!({"type":"string", "value":"Empty"}))
    );
    assert!(
        sheet.formulas.is_empty(),
        "summary must export values, not formulas"
    );
    let user_projection = serde_json::to_value(&workbook.sheets)?;
    let developer_path = PathBuf::from(
        developer_report["path"]
            .as_str()
            .context("developer summary path must be a string")?,
    );
    let developer_workbook = excel_reader::read_workbook(&developer_path)?;
    assert_eq!(
        user_projection,
        serde_json::to_value(&developer_workbook.sheets)?
    );
    let quiet_path = PathBuf::from(
        quiet_report["path"]
            .as_str()
            .context("quiet summary path must be a string")?,
    );
    let quiet_workbook = excel_reader::read_workbook(&quiet_path)?;
    assert_eq!(
        user_projection,
        serde_json::to_value(&quiet_workbook.sheets)?
    );
    assert_eq!(fs::read(&fixture)?, original, "source workbook changed");

    fs::remove_dir_all(output_dir)?;
    Ok(())
}

#[test]
fn summary_used_range_cli_preserves_common_and_textual_error_token_matrix() -> Result<()> {
    const TOKENS: [&str; 15] = [
        "#DIV/0!",
        "#N/A",
        "#NAME?",
        "#NULL!",
        "#NUM!",
        "#REF!",
        "#VALUE!",
        "#DATA!",
        "#SPILL!",
        "#CALC!",
        "#FIELD!",
        "#BLOCKED!",
        "#UNKNOWN!",
        "#CONNECT!",
        "#BUSY!",
    ];
    const INPUT_CELLS: [&str; 15] = [
        "A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1", "I1", "J1", "K1", "L1", "M1", "N1", "O1",
    ];
    const OUTPUT_CELLS: [&str; 15] = [
        "D2", "E2", "F2", "G2", "H2", "I2", "J2", "K2", "L2", "M2", "N2", "O2", "P2", "Q2", "R2",
    ];

    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/summary_used_range/error_token_matrix.xlsx");
    let original = fs::read(&fixture)?;
    let output_dir = unique_test_dir();
    let (output, report) = run_unknown_error_cli(&fixture, &output_dir, Some("developer"))?;
    assert!(
        output.status.success(),
        "matrix CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert_eq!(report["saved"], true);
    assert_eq!(report["rows"], 1);
    assert_eq!(report["files"], 1);
    assert_eq!(
        report["execution"]["reader_engine"],
        "calamine + OOXML error-token normalization"
    );
    assert_eq!(
        report["execution"]["unknown_error_token_cells_normalized"],
        8
    );
    assert_eq!(report["execution"]["ooxml_error_cells_patched"], 7);

    for (coordinate, token) in INPUT_CELLS.iter().zip(TOKENS) {
        let (cell_type, cell_value) = worksheet_cell_type_and_value(&fixture, coordinate)?;
        assert_eq!(cell_type, "e", "fixture {coordinate} must be an error cell");
        assert_eq!(cell_value.as_deref(), Some(token));
    }

    let output_path = PathBuf::from(
        report["path"]
            .as_str()
            .context("matrix summary output path must be a string")?,
    );
    let workbook = excel_reader::read_workbook(&output_path)?;
    let summary = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "汇总")
        .context("matrix summary sheet missing")?;
    for (index, (coordinate, token)) in OUTPUT_CELLS.iter().zip(TOKENS).enumerate() {
        let cell_value = summary
            .values
            .get(*coordinate)
            .with_context(|| format!("matrix summary cell {coordinate} missing"))?;
        let cell_value = excel_reader::CellValue::from_probe_json(cell_value)?;
        let (cell_type, raw_value) = worksheet_cell_type_and_value(&output_path, coordinate)?;
        if index < 7 {
            assert!(
                matches!(&cell_value, excel_reader::CellValue::Error(_)),
                "{coordinate} must remain an Excel error: {cell_value:?}"
            );
            assert_eq!(cell_type, "e", "{coordinate} must retain error type");
            assert_eq!(raw_value.as_deref(), Some(token));
        } else {
            assert_eq!(
                cell_value,
                excel_reader::CellValue::Text(token.to_string()),
                "{coordinate} must be text after unknown-token compatibility read"
            );
            assert_ne!(cell_type, "e", "{coordinate} must be written as text");
        }
    }

    let developer_log_path = PathBuf::from(
        report["run_log_path"]
            .as_str()
            .context("matrix developer run log path must be a string")?,
    );
    let developer_log_bytes = fs::read(&developer_log_path)?;
    let developer_log_text = String::from_utf8_lossy(&developer_log_bytes);
    let developer_log: Value = serde_json::from_slice(&developer_log_bytes)
        .context("parse error-token matrix PT-6-1 run log")?;
    assert_eq!(developer_log["status"], "succeeded");
    for token in TOKENS {
        assert!(
            !developer_log_text.contains(token),
            "developer history must not expose the {token} cell value"
        );
    }
    assert_eq!(fs::read(&fixture)?, original, "source fixture changed");

    fs::remove_dir_all(output_dir)?;
    Ok(())
}

#[test]
fn summary_used_range_cli_reads_vba_enabled_xlsm() -> Result<()> {
    // This XlsxWriter test fixture is pinned at commit a880ecc15c1573449dba051e5c21025f00c7c9fc.
    // It contains a real xl/vbaProject.bin; this test reads the workbook and never executes VBA.
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/summary_used_range/vba_enabled_xlsxwriter_macro01.xlsm");
    let original = fs::read(&fixture)?;
    let mut source_package =
        zip::ZipArchive::new(fs::File::open(&fixture)?).context("open VBA-enabled XLSM fixture")?;
    let vba_project_size = source_package
        .by_name("xl/vbaProject.bin")
        .context("VBA project part missing from XLSM fixture")?
        .size();
    assert!(vba_project_size > 0);

    let output_dir = unique_test_dir();
    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("summary-used-range")
        .arg("--output-dir")
        .arg(output_dir.join("output"))
        .arg("--log-dir")
        .arg(output_dir.join("logs"))
        .arg("--log-view")
        .arg("developer")
        .arg(&fixture)
        .output()
        .context("run pytools summary-used-range with VBA-enabled XLSM")?;
    assert!(
        output.status.success(),
        "CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report: Value = serde_json::from_slice(&output.stdout).context("parse CLI JSON")?;
    assert_eq!(report["saved"], true);
    assert_eq!(report["files"], 1);
    assert_eq!(report["rows"], 1);
    assert_eq!(report["execution"]["reader_engine"], "calamine");
    assert_eq!(report["execution"]["writer_engine"], "rust_xlsxwriter");

    let run_log_path = PathBuf::from(
        report["run_log_path"]
            .as_str()
            .context("CLI run log path must be a string")?,
    );
    let run_log: Value =
        serde_json::from_slice(&fs::read(run_log_path)?).context("parse PT-6-1 run log")?;
    assert_eq!(run_log["status"], "succeeded");
    let write_event = run_log["events"]
        .as_array()
        .context("run log events must be an array")?
        .iter()
        .find(|event| {
            event["phase"] == "output_workbook"
                && event["status"] == "succeeded"
                && event["actual_engine"] == "rust_xlsxwriter"
        })
        .context("completed output workbook event missing")?;
    assert_eq!(
        report["execution"]["write_elapsed_ms"],
        write_event["stage_elapsed_ms"]
    );

    let output_path = PathBuf::from(
        report["path"]
            .as_str()
            .context("summary output path must be a string")?,
    );
    assert_eq!(
        output_path.extension().and_then(|ext| ext.to_str()),
        Some("xlsx")
    );
    let mut output_package = zip::ZipArchive::new(fs::File::open(&output_path)?)
        .context("open generated summary workbook")?;
    assert!(
        output_package.by_name("xl/vbaProject.bin").is_err(),
        "new summary workbook must not copy source VBA"
    );
    let workbook = excel_reader::read_workbook(&output_path)?;
    let summary = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "汇总")
        .context("summary sheet missing")?;
    assert!(!summary.values.is_empty());
    assert!(summary.formulas.is_empty());
    assert_eq!(fs::read(&fixture)?, original, "source workbook changed");

    fs::remove_dir_all(output_dir)?;
    Ok(())
}

#[test]
fn summary_used_range_cli_preserves_unknown_xlsx_error_value_as_text() -> Result<()> {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/summary_used_range/getting_data_error.xlsx");
    let original = fs::read(&fixture)?;
    let output_dir = unique_test_dir();
    let user_dir = output_dir.join("user");
    let (user_output, user_report) = run_unknown_error_cli(&fixture, &user_dir, None)?;
    assert!(
        user_output.status.success(),
        "user CLI failed: {}",
        String::from_utf8_lossy(&user_output.stderr)
    );
    assert_eq!(user_report["saved"], true);
    assert_eq!(user_report["rows"], 1);
    assert_eq!(
        user_report["execution"]["reader_engine"],
        "calamine + OOXML error-token normalization"
    );
    assert_eq!(
        user_report["execution"]["unknown_error_token_cells_normalized"],
        1
    );
    assert_eq!(user_report["execution"]["ooxml_error_cells_patched"], 0);
    let user_stderr = String::from_utf8_lossy(&user_output.stderr);
    assert!(!user_stderr.contains("#DATA!"));
    assert!(!user_stderr.contains("calamine"));
    assert!(!user_stderr.contains("OOXML"));
    let (user_projection, user_error_value) = summary_output_projection(&user_report)?;
    assert_eq!(user_error_value["type"], "string");
    assert_eq!(user_error_value["value"], "#DATA!");

    let developer_dir = output_dir.join("developer");
    let (developer_output, developer_report) =
        run_unknown_error_cli(&fixture, &developer_dir, Some("developer"))?;
    assert!(
        developer_output.status.success(),
        "developer CLI failed: {}",
        String::from_utf8_lossy(&developer_output.stderr)
    );
    let developer_events = String::from_utf8_lossy(&developer_output.stderr)
        .lines()
        .map(serde_json::from_str::<Value>)
        .collect::<std::result::Result<Vec<_>, _>>()
        .context("parse developer compatibility events")?;
    let normalization_event = developer_events
        .iter()
        .find(|event| {
            event["phase"] == "value_read_compatibility" && event["status"] == "succeeded"
        })
        .context("unknown error-token compatibility event missing")?;
    assert_eq!(
        normalization_event["actual_engine"],
        "calamine + OOXML error-token normalization"
    );
    assert!(normalization_event["stage_elapsed_ms"].is_number());
    assert_eq!(
        normalization_event["user_message"],
        "已按文本兼容读取 1 个未知错误值"
    );
    assert!(normalization_event["components"]
        .as_array()
        .unwrap()
        .iter()
        .any(|component| component == "OOXML error-token normalization"));
    let developer_log_path = PathBuf::from(
        developer_report["run_log_path"]
            .as_str()
            .context("developer run log path must be a string")?,
    );
    let developer_log_bytes = fs::read(developer_log_path)?;
    let developer_log_text = String::from_utf8_lossy(&developer_log_bytes);
    let developer_log: Value =
        serde_json::from_slice(&developer_log_bytes).context("parse developer PT-6-1 log")?;
    assert_eq!(developer_log["status"], "succeeded");
    assert!(!developer_log_text.contains("#DATA!"));
    assert_eq!(
        developer_log["events"].as_array().unwrap().len(),
        developer_events.len()
    );
    let (developer_projection, _) = summary_output_projection(&developer_report)?;
    assert_eq!(user_projection, developer_projection);

    let quiet_dir = output_dir.join("quiet");
    let (quiet_output, quiet_report) = run_unknown_error_cli(&fixture, &quiet_dir, Some("off"))?;
    assert!(
        quiet_output.status.success(),
        "off CLI failed: {}",
        String::from_utf8_lossy(&quiet_output.stderr)
    );
    assert!(quiet_output.stderr.is_empty());
    assert!(quiet_report.get("run_id").is_none());
    assert!(quiet_report.get("run_log_path").is_none());
    assert!(!quiet_dir.join("logs").exists());
    let (quiet_projection, _) = summary_output_projection(&quiet_report)?;
    assert_eq!(user_projection, quiet_projection);

    assert_eq!(fs::read(&fixture)?, original, "source workbook changed");

    fs::remove_dir_all(output_dir)?;
    Ok(())
}

fn run_unknown_error_cli(
    fixture: &Path,
    run_dir: &Path,
    log_view: Option<&str>,
) -> Result<(std::process::Output, Value)> {
    let mut command = Command::new(env!("CARGO_BIN_EXE_pytools"));
    command
        .arg("summary-used-range")
        .arg("--output-dir")
        .arg(run_dir.join("output"))
        .arg("--log-dir")
        .arg(run_dir.join("logs"));
    if let Some(log_view) = log_view {
        command.arg("--log-view").arg(log_view);
    }
    let output = command
        .arg(fixture)
        .output()
        .context("run pytools summary-used-range with an unknown XLSX error")?;
    let report = serde_json::from_slice(&output.stdout).context("parse unknown error CLI JSON")?;
    Ok((output, report))
}

fn summary_output_projection(report: &Value) -> Result<(Value, Value)> {
    let path = PathBuf::from(
        report["path"]
            .as_str()
            .context("summary output path must be a string")?,
    );
    let workbook = excel_reader::read_workbook(&path)?;
    let summary = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "汇总")
        .context("summary sheet missing")?;
    let error_value = summary
        .values
        .get("D2")
        .context("normalized #DATA! output cell missing")?
        .clone();
    Ok((serde_json::to_value(&workbook.sheets)?, error_value))
}

fn worksheet_cell_type_and_value(
    workbook_path: &Path,
    coordinate: &str,
) -> Result<(String, Option<String>)> {
    let mut archive = zip::ZipArchive::new(fs::File::open(workbook_path)?)
        .with_context(|| format!("open workbook package {}", workbook_path.display()))?;
    let mut worksheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut worksheet_xml)
        .context("read first worksheet XML")?;
    let document = roxmltree::Document::parse(&worksheet_xml)
        .with_context(|| format!("parse first worksheet XML in {}", workbook_path.display()))?;
    let cell = document
        .descendants()
        .find(|node| {
            node.is_element()
                && node.tag_name().name() == "c"
                && node.attribute("r") == Some(coordinate)
        })
        .with_context(|| format!("worksheet cell {coordinate} missing"))?;
    let cell_type = cell.attribute("t").unwrap_or_default().to_string();
    let raw_value = cell
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == "v")
        .and_then(|node| node.text())
        .map(str::to_string);
    Ok((cell_type, raw_value))
}

#[test]
fn summary_used_range_cli_reports_no_reader_when_no_workbook_is_read() -> Result<()> {
    let output_dir = unique_test_dir();
    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("summary-used-range")
        .arg("--output-dir")
        .arg(output_dir.join("output"))
        .arg("--log-dir")
        .arg(output_dir.join("logs"))
        .arg(output_dir.join("missing.xlsx"))
        .output()
        .context("run pytools summary-used-range with a missing source")?;
    assert!(
        output.status.success(),
        "CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report: Value = serde_json::from_slice(&output.stdout).context("parse CLI JSON")?;
    assert_eq!(report["saved"], false);
    assert_eq!(report["files"], 0);
    assert!(report["execution"]["reader_engine"].is_null());
    assert!(report["execution"]["writer_engine"].is_null());
    let run_log_path = PathBuf::from(
        report["run_log_path"]
            .as_str()
            .context("skipped CLI run log path must be a string")?,
    );
    let run_log: Value =
        serde_json::from_slice(&fs::read(run_log_path)?).context("parse skipped PT-6-1 run log")?;
    assert_eq!(run_log["status"], "skipped");
    assert_eq!(
        run_log["events"].as_array().unwrap().last().unwrap()["status"],
        "skipped"
    );
    assert!(run_log["events"]
        .as_array()
        .unwrap()
        .iter()
        .all(|event| { event["actual_engine"].is_null() }));

    fs::remove_dir_all(output_dir)?;
    Ok(())
}

fn unique_test_dir() -> PathBuf {
    let temp_dir = std::env::temp_dir();
    loop {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("system clock after Unix epoch")
            .as_nanos();
        let sequence = NEXT_TEST_DIRECTORY_ID.fetch_add(1, Ordering::Relaxed);
        let candidate = temp_dir.join(format!(
            "pytools-summary-error-{}-{nonce}-{sequence}",
            std::process::id()
        ));
        match fs::create_dir(&candidate) {
            Ok(()) => return candidate,
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => panic!("create unique test directory: {error}"),
        }
    }
}
