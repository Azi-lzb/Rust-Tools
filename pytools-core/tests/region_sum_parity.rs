use anyhow::{Context, Result};
use pytools_core::config::load_global;
use pytools_core::region_sum::run_region_sum_check_with_timestamp;
use roxmltree::{Document, Node};
use serde_json::{json, Map, Value};
use std::fs::{self, File};
use std::io::{Read, Write};
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::sync::atomic::{AtomicU64, Ordering};
use zip::ZipArchive;

const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
static NEXT_TEST_DIRECTORY_ID: AtomicU64 = AtomicU64::new(0);

#[test]
fn region_sum_workflow_matches_python_values_stats_and_output_presentation() -> Result<()> {
    let cases: Value = serde_json::from_str(include_str!("fixtures/region_sum/cases.json"))?;
    let baseline: Value = serde_json::from_str(include_str!(
        "../probe/reports/region_sum_python_baseline.json"
    ))?;
    let fixture_dir = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/region_sum");
    let input_paths = cases["run_files"]
        .as_array()
        .context("run_files must be an array")?
        .iter()
        .map(|name| Ok(fixture_dir.join(name.as_str().context("fixture names must be strings")?)))
        .collect::<Result<Vec<_>>>()?;
    let input_bytes = input_paths
        .iter()
        .map(|path| Ok((path.clone(), fs::read(path)?)))
        .collect::<Result<Vec<_>>>()?;

    let test_dir = unique_test_dir();
    let result = run_region_sum_check_with_timestamp(
        &input_paths,
        test_dir.join("output"),
        test_dir.join("logs"),
        "20261023_080000",
        "2026-10-23 08:00:00",
    )?;

    for field in ["files", "sheets", "errors", "failed_files"] {
        assert_eq!(
            serde_json::to_value(match field {
                "files" => result.files,
                "sheets" => result.sheets,
                "errors" => result.errors,
                "failed_files" => result.failed_files,
                _ => unreachable!(),
            })?,
            baseline["statistics"][field],
            "statistics field {field}"
        );
    }
    for field in ["files", "sheets", "errors", "failed_files"] {
        assert_eq!(
            cases["expected_stats"][field], baseline["statistics"][field],
            "checked-in fixture expectation {field}"
        );
    }

    let report_path = result.path.context("report output path")?;
    assert!(report_path.is_file());
    assert_eq!(
        report_path.file_name().and_then(|name| name.to_str()),
        Some("地区数值核对结果_20261023_080000.xlsx")
    );
    assert!(result.log_path.is_file());
    let log = fs::read_to_string(&result.log_path)?;
    assert!(log.contains("balanced.xlsx: 校验 1 个分地区 sheet，错误 0"));
    assert!(log.contains("结果已保存:"));

    let workbook = excel_reader::read_workbook(&report_path)?;
    assert_eq!(workbook.sheets.len(), 1);
    let sheet = &workbook.sheets[0];
    assert_eq!(sheet.name, baseline["output_workbook"]["sheet_name"]);
    let actual_cells = normalize_cell_map(&serde_json::to_value(&sheet.values)?)?;
    let expected_cells = normalize_cell_map(&baseline["output_workbook"]["cells"])?;
    assert_eq!(actual_cells, expected_cells);
    assert_eq!(
        serde_json::to_value(&sheet.formulas)?,
        normalize_openpyxl_formula_map(&baseline["output_workbook"]["formulas"])?
    );

    let layout = inspect_layout(&report_path)?;
    let expected = &baseline["output_workbook"];
    assert_eq!(layout["max_row"], expected["max_row"]);
    assert_eq!(layout["max_column"], expected["max_column"]);
    assert_eq!(layout["styles"], expected["styles"]);
    let expected_widths = expected["column_widths"]
        .as_object()
        .context("column widths")?;
    let actual_widths = layout["column_widths"]
        .as_object()
        .context("actual column widths")?;
    assert_eq!(expected_widths.len(), actual_widths.len());
    for (column, expected_width) in expected_widths {
        let expected_width = expected_width
            .as_f64()
            .context("expected width must be numeric")?;
        let actual_width = actual_widths[column]
            .as_f64()
            .context("actual width must be numeric")?;
        assert!(
            (expected_width - actual_width).abs() <= 0.01,
            "column {column}: Python width {expected_width}, Rust width {actual_width}"
        );
    }

    for (path, original_bytes) in &input_bytes {
        let actual_bytes = fs::read(path)?;
        assert_eq!(
            actual_bytes.as_slice(),
            original_bytes.as_slice(),
            "input changed: {}",
            path.display()
        );
    }
    fs::remove_dir_all(test_dir)?;
    Ok(())
}

#[test]
fn region_sum_cli_runs_python_baseline_workflow() -> Result<()> {
    let cases: Value = serde_json::from_str(include_str!("fixtures/region_sum/cases.json"))?;
    let baseline: Value = serde_json::from_str(include_str!(
        "../probe/reports/region_sum_python_baseline.json"
    ))?;
    let fixture_dir = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/region_sum");
    let input_paths = cases["run_files"]
        .as_array()
        .context("run_files must be an array")?
        .iter()
        .map(|name| Ok(fixture_dir.join(name.as_str().context("fixture names must be strings")?)))
        .collect::<Result<Vec<_>>>()?;
    let input_bytes = input_paths
        .iter()
        .map(|path| Ok((path.clone(), fs::read(path)?)))
        .collect::<Result<Vec<_>>>()?;
    let test_dir = unique_test_dir();
    let output_dir = test_dir.join("output");
    let log_dir = test_dir.join("logs");
    let config_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/config_loader.xlsx");
    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("region-sum-check")
        .arg("--config")
        .arg(config_path)
        .arg("--output-dir")
        .arg(&output_dir)
        .arg("--log-dir")
        .arg(&log_dir)
        .args(&input_paths)
        .output()
        .context("run pytools region-sum-check")?;
    assert!(
        output.status.success(),
        "CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert!(
        output.stderr.is_empty(),
        "legacy CLI output changed without --log-view: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let result: Value = serde_json::from_slice(&output.stdout).context("parse CLI JSON")?;
    for field in ["files", "sheets", "errors", "failed_files"] {
        assert_eq!(
            result[field], baseline["statistics"][field],
            "field {field}"
        );
        assert_eq!(
            result[field], cases["expected_stats"][field],
            "fixture field {field}"
        );
    }

    let report_path = PathBuf::from(result["path"].as_str().context("CLI report path")?);
    let log_path = PathBuf::from(result["log_path"].as_str().context("CLI log path")?);
    assert!(report_path.is_file());
    assert!(log_path.is_file());
    let workbook = excel_reader::read_workbook(&report_path)?;
    let sheet = workbook.sheets.first().context("result worksheet")?;
    assert_eq!(sheet.name, baseline["output_workbook"]["sheet_name"]);
    assert_eq!(
        normalize_cell_map(&serde_json::to_value(&sheet.values)?)?,
        normalize_cell_map(&baseline["output_workbook"]["cells"])?
    );
    assert_eq!(
        serde_json::to_value(&sheet.formulas)?,
        normalize_openpyxl_formula_map(&baseline["output_workbook"]["formulas"])?
    );
    for (path, original_bytes) in input_bytes {
        assert_eq!(
            fs::read(&path)?,
            original_bytes,
            "input changed: {}",
            path.display()
        );
    }
    fs::remove_dir_all(test_dir)?;
    Ok(())
}

#[test]
fn region_sum_cli_projects_cached_formula_totals_like_python() -> Result<()> {
    let cases: Value = serde_json::from_str(include_str!(
        "fixtures/region_sum/cached_formula/cases.json"
    ))?;
    let baseline: Value = serde_json::from_str(include_str!(
        "../probe/reports/region_sum_cached_formula_python_baseline.json"
    ))?;
    let fixture_dir =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/region_sum/cached_formula");
    let input_paths = cases["run_files"]
        .as_array()
        .context("cached formula run_files must be an array")?
        .iter()
        .map(|name| Ok(fixture_dir.join(name.as_str().context("fixture names must be strings")?)))
        .collect::<Result<Vec<_>>>()?;
    let input_bytes = input_paths
        .iter()
        .map(|path| Ok((path.clone(), fs::read(path)?)))
        .collect::<Result<Vec<_>>>()?;

    let test_dir = unique_test_dir();
    let output_dir = test_dir.join("output");
    let log_dir = test_dir.join("logs");
    let config_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/config_loader.xlsx");
    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("region-sum-check")
        .arg("--config")
        .arg(config_path)
        .arg("--output-dir")
        .arg(&output_dir)
        .arg("--log-dir")
        .arg(&log_dir)
        .args(&input_paths)
        .output()
        .context("run PT-7-5 CLI with cached formula totals")?;
    assert!(
        output.status.success(),
        "CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert!(
        output.stderr.is_empty(),
        "legacy CLI output changed without --log-view: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let result: Value = serde_json::from_slice(&output.stdout).context("parse CLI JSON")?;
    for field in ["files", "sheets", "errors", "failed_files"] {
        assert_eq!(
            result[field], baseline["statistics"][field],
            "cached formula statistics field {field}"
        );
        assert_eq!(
            result[field], cases["expected_stats"][field],
            "cached formula fixture field {field}"
        );
    }

    let report_path = PathBuf::from(result["path"].as_str().context("CLI report path")?);
    let log_path = PathBuf::from(result["log_path"].as_str().context("CLI log path")?);
    assert!(report_path.is_file());
    assert!(log_path.is_file());
    let workbook = excel_reader::read_workbook(&report_path)?;
    let sheet = workbook.sheets.first().context("result worksheet")?;
    let expected = &baseline["output_workbook"];
    assert_eq!(sheet.name, expected["sheet_name"]);
    assert_eq!(
        normalize_cell_map(&serde_json::to_value(&sheet.values)?)?,
        normalize_cell_map(&expected["cells"])?
    );
    assert_eq!(
        serde_json::to_value(&sheet.formulas)?,
        normalize_openpyxl_formula_map(&expected["formulas"])?
    );
    let layout = inspect_layout(&report_path)?;
    assert_eq!(layout["max_row"], expected["max_row"]);
    assert_eq!(layout["max_column"], expected["max_column"]);
    assert_eq!(layout["styles"], expected["styles"]);
    let expected_widths = expected["column_widths"]
        .as_object()
        .context("cached formula Python baseline widths")?;
    let actual_widths = layout["column_widths"]
        .as_object()
        .context("cached formula Rust output widths")?;
    assert_eq!(actual_widths.len(), expected_widths.len());
    for (column, expected_width) in expected_widths {
        let expected_width = expected_width.as_f64().context("expected width")?;
        let actual_width = actual_widths[column].as_f64().context("actual width")?;
        assert!(
            (expected_width - actual_width).abs() <= 0.01,
            "cached formula column {column}: Python width {expected_width}, Rust width {actual_width}"
        );
    }
    for (path, original) in input_bytes {
        assert_eq!(
            fs::read(&path)?,
            original,
            "input changed: {}",
            path.display()
        );
    }

    fs::remove_dir_all(test_dir)?;
    Ok(())
}

#[test]
fn interactive_region_sum_menu_matches_python_baseline_workbook() -> Result<()> {
    let cases: Value = serde_json::from_str(include_str!("fixtures/region_sum/cases.json"))?;
    let baseline: Value = serde_json::from_str(include_str!(
        "../probe/reports/region_sum_python_baseline.json"
    ))?;
    let fixture_dir = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/region_sum");
    let input_paths = cases["run_files"]
        .as_array()
        .context("run_files must be an array")?
        .iter()
        .map(|name| Ok(fixture_dir.join(name.as_str().context("fixture names must be strings")?)))
        .collect::<Result<Vec<_>>>()?;
    let input_bytes = input_paths
        .iter()
        .map(|path| Ok((path.clone(), fs::read(path)?)))
        .collect::<Result<Vec<_>>>()?;

    let test_dir = unique_test_dir();
    let config_dir = test_dir.join("PyCode");
    fs::create_dir_all(&config_dir)?;
    let config_path = config_dir.join("config.xlsx");
    fs::copy(
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/config_loader.xlsx"),
        &config_path,
    )?;
    let global = load_global(&config_path)?;
    let paths = input_paths
        .iter()
        .map(|path| path.to_string_lossy())
        .collect::<Vec<_>>()
        .join(";");
    let input = format!("5\n5\n1\n{paths}\n0\n0\n");
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("menu")
        .arg("--config")
        .arg(&config_path)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .context("start PyTools interactive menu")?;
    child
        .stdin
        .take()
        .context("menu stdin missing")?
        .write_all(input.as_bytes())?;
    let output = child
        .wait_with_output()
        .context("wait for PyTools interactive menu")?;
    assert!(
        output.status.success(),
        "interactive PT-7-5 menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    for (field, expected) in [
        ("files", &baseline["statistics"]["files"]),
        ("sheets", &baseline["statistics"]["sheets"]),
        ("errors", &baseline["statistics"]["errors"]),
        ("failed_files", &baseline["statistics"]["failed_files"]),
    ] {
        assert!(
            stdout.contains(&format!("\"{field}\": {expected}")),
            "menu result did not report Python baseline {field}={expected}: {stdout}"
        );
    }

    let report_path = fs::read_dir(&global.output_dir)?
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .find(|path| {
            path.file_name()
                .and_then(|name| name.to_str())
                .is_some_and(|name| {
                    name.starts_with("地区数值核对结果_") && name.ends_with(".xlsx")
                })
        })
        .context("interactive PT-7-5 result workbook missing")?;
    let workbook = excel_reader::read_workbook(&report_path)?;
    assert_eq!(workbook.sheets.len(), 1);
    let sheet = &workbook.sheets[0];
    assert_eq!(sheet.name, baseline["output_workbook"]["sheet_name"]);
    assert_eq!(
        normalize_cell_map(&serde_json::to_value(&sheet.values)?)?,
        normalize_cell_map(&baseline["output_workbook"]["cells"])?
    );
    assert_eq!(
        serde_json::to_value(&sheet.formulas)?,
        normalize_openpyxl_formula_map(&baseline["output_workbook"]["formulas"])?
    );
    let layout = inspect_layout(&report_path)?;
    let expected = &baseline["output_workbook"];
    assert_eq!(layout["max_row"], expected["max_row"]);
    assert_eq!(layout["max_column"], expected["max_column"]);
    assert_eq!(layout["styles"], expected["styles"]);
    let expected_widths = expected["column_widths"]
        .as_object()
        .context("Python baseline column widths missing")?;
    let actual_widths = layout["column_widths"]
        .as_object()
        .context("interactive result column widths missing")?;
    assert_eq!(actual_widths.len(), expected_widths.len());
    for (column, expected_width) in expected_widths {
        let expected_width = expected_width
            .as_f64()
            .context("expected width must be numeric")?;
        let actual_width = actual_widths[column]
            .as_f64()
            .context("actual width must be numeric")?;
        assert!(
            (expected_width - actual_width).abs() <= 0.01,
            "interactive column {column}: Python width {expected_width}, Rust width {actual_width}"
        );
    }
    for (path, original) in input_bytes {
        assert_eq!(
            fs::read(&path)?,
            original,
            "input changed: {}",
            path.display()
        );
    }

    fs::remove_dir_all(test_dir)?;
    Ok(())
}

#[test]
fn region_sum_cli_uses_configured_output_and_log_directories() -> Result<()> {
    let test_dir = unique_test_dir();
    let pycode_dir = test_dir.join("PyCode");
    fs::create_dir_all(&pycode_dir)?;
    let source_config =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/config_loader.xlsx");
    let config_path = pycode_dir.join("config.xlsx");
    fs::copy(source_config, &config_path)?;
    let global = load_global(&config_path)?;
    let source =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/region_sum/balanced.xlsx");
    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("region-sum-check")
        .arg("--config")
        .arg(&config_path)
        .arg(&source)
        .output()
        .context("run pytools region-sum-check with configured output paths")?;
    assert!(
        output.status.success(),
        "CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let result: Value = serde_json::from_slice(&output.stdout).context("parse CLI JSON")?;
    let report_path = PathBuf::from(result["path"].as_str().context("CLI report path")?);
    let log_path = PathBuf::from(result["log_path"].as_str().context("CLI log path")?);
    assert!(report_path.starts_with(&global.output_dir));
    assert!(log_path.starts_with(&global.log_dir));
    assert_eq!(result["files"], 1);
    assert_eq!(result["sheets"], 1);
    assert_eq!(result["errors"], 0);
    assert_eq!(result["failed_files"], 0);
    fs::remove_dir_all(test_dir)?;
    Ok(())
}

#[test]
fn output_name_collisions_preserve_existing_files_and_input_workbook() -> Result<()> {
    let fixture_dir = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/region_sum");
    let input = fixture_dir.join("balanced.xlsx");
    let input_before = fs::read(&input)?;
    let test_dir = unique_test_dir();
    let output_dir = test_dir.join("output");
    let log_dir = test_dir.join("logs");
    fs::create_dir_all(&output_dir)?;
    fs::create_dir_all(&log_dir)?;

    let existing_output = output_dir.join("地区数值核对结果_20261023_080001.xlsx");
    let existing_log = log_dir.join("feature_1_5_region_sum_check_20261023_080001.log");
    let output_sentinel = b"existing output must remain untouched";
    let log_sentinel = b"existing log must remain untouched";
    fs::write(&existing_output, output_sentinel)?;
    fs::write(&existing_log, log_sentinel)?;

    let result = run_region_sum_check_with_timestamp(
        &[input.clone()],
        &output_dir,
        &log_dir,
        "20261023_080001",
        "2026-10-23 08:00:01",
    )?;

    assert_eq!(
        result.path.as_deref().and_then(Path::file_name),
        Some(std::ffi::OsStr::new(
            "地区数值核对结果_20261023_080001_1.xlsx"
        ))
    );
    assert_eq!(
        result.log_path.file_name(),
        Some(std::ffi::OsStr::new(
            "feature_1_5_region_sum_check_20261023_080001_1.log"
        ))
    );
    let output_workbook =
        excel_reader::read_workbook(result.path.as_ref().context("output path")?)?;
    let output_sheet = output_workbook.sheets.first().context("result worksheet")?;
    assert!(output_sheet.formulas.contains_key("A3"));
    assert!(!output_sheet.values.contains_key("A3"));
    assert_eq!(fs::read(existing_output)?, output_sentinel);
    assert_eq!(fs::read(existing_log)?, log_sentinel);
    assert_eq!(fs::read(input)?, input_before);
    fs::remove_dir_all(test_dir)?;
    Ok(())
}

#[test]
fn output_directory_failure_does_not_modify_input_or_blocking_file() -> Result<()> {
    let fixture_dir = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/region_sum");
    let input = fixture_dir.join("balanced.xlsx");
    let input_before = fs::read(&input)?;
    let test_dir = unique_test_dir();
    fs::create_dir_all(&test_dir)?;
    let output_path = test_dir.join("output-is-a-file");
    let blocking_file = b"output path blocker";
    fs::write(&output_path, blocking_file)?;

    let error = run_region_sum_check_with_timestamp(
        &[input.clone()],
        &output_path,
        test_dir.join("logs"),
        "20261023_080002",
        "2026-10-23 08:00:02",
    )
    .expect_err("a file cannot be used as the output directory");
    assert!(error.to_string().contains("create output directory"));
    assert_eq!(fs::read(&output_path)?, blocking_file);
    assert_eq!(fs::read(input)?, input_before);
    fs::remove_dir_all(test_dir)?;
    Ok(())
}

fn unique_test_dir() -> PathBuf {
    let nonce = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .expect("clock after epoch")
        .as_nanos();
    let serial = NEXT_TEST_DIRECTORY_ID.fetch_add(1, Ordering::Relaxed);
    std::env::temp_dir().join(format!(
        "pytools_region_sum_{}_{}_{}",
        std::process::id(),
        nonce,
        serial
    ))
}

fn normalize_cell_map(value: &Value) -> Result<Value> {
    let cells = value.as_object().context("cells must be an object")?;
    let mut normalized = Map::new();
    for (address, cell) in cells {
        let kind = cell
            .get("type")
            .and_then(Value::as_str)
            .context("cell type")?;
        if matches!(kind, "integer" | "number") {
            let number = cell
                .get("value")
                .and_then(Value::as_f64)
                .context("numeric cell value")?;
            normalized.insert(address.clone(), json!({"type":"number", "value":number}));
        } else {
            normalized.insert(address.clone(), cell.clone());
        }
    }
    Ok(Value::Object(normalized))
}

fn normalize_openpyxl_formula_map(value: &Value) -> Result<Value> {
    let formulas = value.as_object().context("formulas must be an object")?;
    let mut normalized = Map::new();
    for (address, formula) in formulas {
        let formula = formula.as_str().context("formula must be a string")?;
        normalized.insert(
            address.clone(),
            Value::String(formula.strip_prefix('=').unwrap_or(formula).to_owned()),
        );
    }
    Ok(Value::Object(normalized))
}

fn inspect_layout(path: &Path) -> Result<Value> {
    let file =
        File::open(path).with_context(|| format!("open generated report {}", path.display()))?;
    let mut archive = ZipArchive::new(file).context("open generated xlsx package")?;
    let sheet_xml = read_part(&mut archive, "xl/worksheets/sheet1.xml")?;
    let styles_xml = read_part(&mut archive, "xl/styles.xml")?;
    let sheet_doc = Document::parse(&sheet_xml).context("parse worksheet XML")?;
    let styles_doc = Document::parse(&styles_xml).context("parse styles XML")?;

    let fonts = read_fonts(styles_doc.root_element());
    let fills = read_fills(styles_doc.root_element());
    let formats = read_cell_formats(styles_doc.root_element());

    let dimension = sheet_doc
        .root_element()
        .descendants()
        .find(|node| is_main(*node, "dimension"))
        .and_then(|node| node.attribute("ref"))
        .context("worksheet dimension")?;
    let end_address = dimension.split(':').last().unwrap_or(dimension);
    let (max_row, max_column) = parse_cell_address(end_address).context("parse sheet dimension")?;

    let mut column_widths = Map::new();
    for column_node in sheet_doc
        .root_element()
        .descendants()
        .filter(|node| is_main(*node, "col"))
    {
        let start = column_node
            .attribute("min")
            .and_then(|value| value.parse::<usize>().ok())
            .context("column min")?;
        let end = column_node
            .attribute("max")
            .and_then(|value| value.parse::<usize>().ok())
            .context("column max")?;
        let width = column_node
            .attribute("width")
            .and_then(|value| value.parse::<f64>().ok())
            .context("column width")?;
        for column in start..=end {
            column_widths.insert(column_letters(column), json!(width));
        }
    }

    let mut styles = Map::new();
    let sheet_data = sheet_doc
        .root_element()
        .descendants()
        .find(|node| is_main(*node, "sheetData"))
        .context("worksheet sheetData")?;
    for cell in sheet_data.descendants().filter(|node| is_main(*node, "c")) {
        let Some(address) = cell.attribute("r") else {
            continue;
        };
        let has_formula = cell.children().any(|node| is_main(node, "f"));
        let has_cached_value = cell
            .children()
            .find(|node| is_main(*node, "v"))
            .and_then(|node| node.text())
            .is_some_and(|value| !value.trim().is_empty());
        // The Python baseline inspects styles through data_only=True and omits
        // formula cells without a cached result from its style map.
        if has_formula && !has_cached_value {
            continue;
        }
        let style_index = cell
            .attribute("s")
            .and_then(|value| value.parse::<usize>().ok())
            .unwrap_or(0);
        let (font_index, fill_index, horizontal_alignment) =
            formats.get(style_index).cloned().unwrap_or_default();
        let font = fonts.get(font_index).cloned().unwrap_or_default();
        let fill = fills.get(fill_index).cloned().unwrap_or_default();
        styles.insert(
            address.to_owned(),
            json!({
                "bold": font.0,
                "font_color": font.1,
                "fill_color": fill,
                "horizontal_alignment": horizontal_alignment,
            }),
        );
    }

    Ok(json!({
        "max_row": max_row,
        "max_column": max_column,
        "column_widths": column_widths,
        "styles": styles,
    }))
}

fn read_part(archive: &mut ZipArchive<File>, name: &str) -> Result<String> {
    let mut part = archive
        .by_name(name)
        .with_context(|| format!("read OOXML part {name}"))?;
    let mut xml = String::new();
    part.read_to_string(&mut xml)?;
    Ok(xml)
}

fn is_main(node: Node<'_, '_>, name: &str) -> bool {
    node.tag_name().namespace() == Some(MAIN_NS) && node.tag_name().name() == name
}

fn read_fonts(root: Node<'_, '_>) -> Vec<(bool, Option<String>)> {
    let Some(container) = root.descendants().find(|node| is_main(*node, "fonts")) else {
        return Vec::new();
    };
    container
        .children()
        .filter(|node| is_main(*node, "font"))
        .map(|font| {
            let bold = font
                .children()
                .any(|node| is_main(node, "b") && node.attribute("val") != Some("0"));
            let color = font
                .children()
                .find(|node| is_main(*node, "color"))
                .and_then(|node| node.attribute("rgb"))
                .map(normalize_rgb);
            (bold, color)
        })
        .collect()
}

fn read_fills(root: Node<'_, '_>) -> Vec<Option<String>> {
    let Some(container) = root.descendants().find(|node| is_main(*node, "fills")) else {
        return Vec::new();
    };
    container
        .children()
        .filter(|node| is_main(*node, "fill"))
        .map(|fill| {
            let pattern = fill.children().find(|node| is_main(*node, "patternFill"))?;
            if pattern.attribute("patternType") != Some("solid") {
                return None;
            }
            pattern
                .children()
                .find(|node| is_main(*node, "fgColor"))
                .and_then(|node| node.attribute("rgb"))
                .map(normalize_rgb)
        })
        .collect()
}

fn read_cell_formats(root: Node<'_, '_>) -> Vec<(usize, usize, Option<String>)> {
    let Some(container) = root.descendants().find(|node| is_main(*node, "cellXfs")) else {
        return Vec::new();
    };
    container
        .children()
        .filter(|node| is_main(*node, "xf"))
        .map(|format| {
            let font = format
                .attribute("fontId")
                .and_then(|value| value.parse::<usize>().ok())
                .unwrap_or(0);
            let fill = format
                .attribute("fillId")
                .and_then(|value| value.parse::<usize>().ok())
                .unwrap_or(0);
            let alignment = format
                .children()
                .find(|node| is_main(*node, "alignment"))
                .and_then(|node| node.attribute("horizontal"))
                .map(str::to_owned);
            (font, fill, alignment)
        })
        .collect()
}

fn normalize_rgb(value: &str) -> String {
    value
        .chars()
        .rev()
        .take(6)
        .collect::<String>()
        .chars()
        .rev()
        .collect()
}

fn parse_cell_address(address: &str) -> Option<(usize, usize)> {
    let letters = address.bytes().take_while(u8::is_ascii_alphabetic).count();
    if letters == 0 || letters == address.len() {
        return None;
    }
    let mut column = 0usize;
    for byte in &address.as_bytes()[..letters] {
        column = column
            .checked_mul(26)?
            .checked_add(usize::from(byte.to_ascii_uppercase() - b'A' + 1))?;
    }
    let row = address[letters..].parse::<usize>().ok()?;
    (row > 0).then_some((row, column))
}

fn column_letters(mut column: usize) -> String {
    let mut letters = String::new();
    while column > 0 {
        column -= 1;
        letters.insert(0, char::from(b'A' + (column % 26) as u8));
        column /= 26;
    }
    letters
}
