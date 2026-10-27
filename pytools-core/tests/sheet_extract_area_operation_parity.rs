use anyhow::Context;
use excel_editor::{
    inspect_xlsx_profile, open_xlsx, BackendCapability, CapabilityEvidence, CapabilityStatus,
    EditOperationKind, EngineId, PreservationRequirement, WorkbookFormat,
};
use excel_reader::{
    read_ooxml_comments, read_value_columns_with_iso_dates, read_workbook,
    read_xlsx_cell_number_formats, read_xlsx_worksheet_extents, CellNumberFormatRequest, CellValue,
};
use pytools_core::config::ExtractTask;
use pytools_core::sheet_extract::run_extract_sheets_area_candidate;
use serde_json::{json, Value};
use std::collections::BTreeSet;
use std::fs;
use std::io::{Read, Write};
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::time::{SystemTime, UNIX_EPOCH};

fn cases() -> Value {
    serde_json::from_str(include_str!("fixtures/sheet_extract_area/cases.json"))
        .expect("valid worksheet extraction operation cases")
}

fn baseline() -> Value {
    serde_json::from_str(include_str!(
        "../probe/reports/sheet_extract_area_python_baseline.json"
    ))
    .expect("valid Python worksheet extraction baseline")
}

fn full_sheet_baseline() -> Value {
    serde_json::from_str(include_str!(
        "../probe/reports/sheet_extract_full_menu_python_baseline.json"
    ))
    .expect("valid Python full-sheet extraction baseline")
}

fn load_tasks(cases: &Value) -> Vec<ExtractTask> {
    cases["tasks"]
        .as_array()
        .unwrap()
        .iter()
        .map(|task| ExtractTask {
            row_no: task["row_no"].as_u64().unwrap() as usize,
            workbook_keywords: task["workbook_keywords"].as_str().unwrap().to_owned(),
            sheet_keywords: task["sheet_keywords"].as_str().unwrap().to_owned(),
            exclude_keywords: task["exclude_keywords"].as_str().unwrap().to_owned(),
            area_spec: task["area_spec"].as_str().unwrap().to_owned(),
            stitch_mode: task["stitch_mode"].as_str().unwrap().to_owned(),
            output_folder: task["output_folder"].as_str().unwrap().to_owned(),
            output_name: task["output_name"].as_str().unwrap().to_owned(),
            target_sheet: task["target_sheet"].as_str().unwrap_or_default().to_owned(),
            remark: task["remark"].as_str().unwrap().to_owned(),
        })
        .collect()
}

fn fixture_path(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/sheet_extract_area")
        .join(name)
}

fn full_sheet_fixture_path(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/sheet_extract_full_menu")
        .join(name)
}

fn temporary_directory(label: &str) -> PathBuf {
    let unique = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap()
        .as_nanos();
    let path =
        std::env::temp_dir().join(format!("pytools-{label}-{}-{unique}", std::process::id()));
    fs::create_dir_all(&path).unwrap();
    path
}

fn create_menu_config(root: &Path, tasks: &[ExtractTask]) -> PathBuf {
    let app = root.join("app");
    let config = app.join("config/config.xlsx");
    fs::create_dir_all(config.parent().unwrap()).unwrap();
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let global = workbook.add_worksheet();
    global.set_name("全局配置").unwrap();
    global.write_string(0, 0, "配置项").unwrap();
    global.write_string(0, 1, "值").unwrap();
    for (row, (key, value)) in [
        ("输入目录", app.join("input")),
        ("输出目录", app.join("output")),
        ("日志目录", app.join("logs")),
    ]
    .iter()
    .enumerate()
    {
        global.write_string((row + 1) as u32, 0, *key).unwrap();
        global
            .write_string((row + 1) as u32, 1, value.to_string_lossy())
            .unwrap();
    }

    let extraction = workbook.add_worksheet();
    extraction.set_name("工作表提取").unwrap();
    let headers = [
        "是否启用",
        "工作簿关键字",
        "工作表关键字",
        "排除关键字",
        "提取区域",
        "多区域拼接方式",
        "输出文件夹",
        "文件名",
        "目标工作表",
        "备注",
    ];
    for (column, header) in headers.iter().enumerate() {
        extraction.write_string(0, column as u16, *header).unwrap();
    }
    for (index, task) in tasks.iter().enumerate() {
        let row = (index + 1) as u32;
        for (column, value) in [
            "是",
            task.workbook_keywords.as_str(),
            task.sheet_keywords.as_str(),
            task.exclude_keywords.as_str(),
            task.area_spec.as_str(),
            task.stitch_mode.as_str(),
            task.output_folder.as_str(),
            task.output_name.as_str(),
            task.target_sheet.as_str(),
            task.remark.as_str(),
        ]
        .iter()
        .enumerate()
        {
            extraction.write_string(row, column as u16, *value).unwrap();
        }
    }
    workbook.save(&config).unwrap();
    config
}

fn run_menu(config: &Path, input: &str) -> std::process::Output {
    run_menu_with_view(config, input, "developer")
}

fn run_menu_with_view(config: &Path, input: &str, view: &str) -> std::process::Output {
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("menu")
        .arg("--config")
        .arg(config)
        .arg("--log-view")
        .arg(view)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .unwrap();
    child
        .stdin
        .take()
        .unwrap()
        .write_all(input.as_bytes())
        .unwrap();
    child.wait_with_output().unwrap()
}

fn run_menu_with_capability(
    config: &Path,
    input: &str,
    view: &str,
    capability: &Path,
    platform: &str,
) -> std::process::Output {
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("menu")
        .arg("--config")
        .arg(config)
        .arg("--log-view")
        .arg(view)
        .arg("--capabilities")
        .arg(capability)
        .arg("--platform")
        .arg(platform)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .unwrap();
    child
        .stdin
        .take()
        .unwrap()
        .write_all(input.as_bytes())
        .unwrap();
    child.wait_with_output().unwrap()
}

fn write_full_sheet_test_capability(
    root: &Path,
    source: &Path,
    platform: &str,
) -> anyhow::Result<PathBuf> {
    let skeleton_path = root.join("full-sheet-capability-skeleton.xlsx");
    let mut skeleton = rust_xlsxwriter::Workbook::new();
    skeleton.add_worksheet().set_name("报表")?;
    skeleton.save(&skeleton_path)?;

    let mut supported_features = inspect_xlsx_profile(source)?.features_present;
    supported_features.extend(inspect_xlsx_profile(&skeleton_path)?.features_present);
    let capability = BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Passed,
        format: WorkbookFormat::Xlsx,
        operations: [EditOperationKind::ImportRange].into_iter().collect(),
        supported_features,
        preserved: [
            PreservationRequirement::CellValuesAndTypes,
            PreservationRequirement::Styles,
            PreservationRequirement::MergedRanges,
            PreservationRequirement::RowAndColumnDimensions,
            PreservationRequirement::UnmodifiedPackageParts,
            PreservationRequirement::SheetIdentityAndOrder,
        ]
        .into_iter()
        .collect::<BTreeSet<_>>(),
        platform: platform.to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: "synthetic-pt-7-6-full-sheet-menu".to_owned(),
            report_path: "pytools-core/tests/sheet_extract_area_operation_parity.rs".to_owned(),
            adapter_version: "pytools-core fixture-scoped test capability".to_owned(),
            engine_version: "umya-spreadsheet 3.1.1".to_owned(),
            platform: platform.to_owned(),
        }),
    };
    let capability_path = root.join("temporary-full-sheet-capability.json");
    fs::write(&capability_path, serde_json::to_vec_pretty(&capability)?)?;
    Ok(capability_path)
}

fn write_append_test_capability(
    root: &Path,
    existing: &Path,
    incoming_source: &Path,
    platform: &str,
) -> anyhow::Result<PathBuf> {
    let mut supported_features = inspect_xlsx_profile(existing)?.features_present;
    supported_features.extend(inspect_xlsx_profile(incoming_source)?.features_present);
    let skeleton_path = root.join("append-capability-skeleton.xlsx");
    let mut skeleton = rust_xlsxwriter::Workbook::new();
    for sheet in read_value_columns_with_iso_dates(incoming_source, 16_384)? {
        skeleton.add_worksheet().set_name(&sheet.name)?;
    }
    skeleton.save(&skeleton_path)?;
    supported_features.extend(inspect_xlsx_profile(&skeleton_path)?.features_present);
    let capability = BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Passed,
        format: WorkbookFormat::Xlsx,
        operations: [
            EditOperationKind::CreateSheet,
            EditOperationKind::ImportRange,
        ]
        .into_iter()
        .collect(),
        supported_features,
        preserved: [
            PreservationRequirement::CellValuesAndTypes,
            PreservationRequirement::FormulaText,
            PreservationRequirement::CommentsAndPositions,
            PreservationRequirement::Styles,
            PreservationRequirement::MergedRanges,
            PreservationRequirement::ConditionalFormatting,
            PreservationRequirement::DataValidations,
            PreservationRequirement::DefinedNames,
            PreservationRequirement::RowAndColumnDimensions,
            PreservationRequirement::PageSettings,
            PreservationRequirement::HeaderFooterText,
            PreservationRequirement::Relationships,
            PreservationRequirement::UnmodifiedPackageParts,
            PreservationRequirement::SheetIdentityAndOrder,
            PreservationRequirement::FreezePanes,
            PreservationRequirement::AutoFilter,
        ]
        .into_iter()
        .collect(),
        platform: platform.to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: "synthetic-pt-7-6-append-cli".to_owned(),
            report_path: "pytools-core/tests/sheet_extract_area_operation_parity.rs".to_owned(),
            adapter_version: "pytools-core fixture-scoped test capability".to_owned(),
            engine_version: "umya-spreadsheet 3.1.1".to_owned(),
            platform: platform.to_owned(),
        }),
    };
    let capability_path = root.join("temporary-append-test-capability.json");
    fs::write(&capability_path, serde_json::to_vec_pretty(&capability)?)?;
    Ok(capability_path)
}

fn write_append_source(path: &Path, sheet_name: &str) -> anyhow::Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name(sheet_name)?;
    sheet.write_string(0, 0, "R1C1")?;
    sheet.write_string(0, 1, "R1C2")?;
    workbook.save(path)?;
    Ok(())
}

fn write_append_full_sheet_source(path: &Path, sheet_name: &str) -> anyhow::Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name(sheet_name)?;
    let title_format = rust_xlsxwriter::Format::new().set_bold();
    sheet.merge_range(0, 0, 0, 1, "标题", &title_format)?;
    sheet.write_number(1, 0, 2)?;
    sheet.write_number(1, 1, 3)?;
    sheet.write_formula(1, 2, "=A2+B2")?;
    sheet.set_row_height(0, 30.0)?;
    sheet.set_column_width(0, 20.0)?;
    workbook.save(path)?;
    Ok(())
}

fn write_existing_full_sheet_output(path: &Path) -> anyhow::Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("第一张表")?;
    let retained_format = rust_xlsxwriter::Format::new().set_bold();
    sheet.write_string_with_format(0, 0, "保留值", &retained_format)?;
    sheet.write_formula(0, 1, "=1+1")?;
    sheet.insert_note(0, 0, &rust_xlsxwriter::Note::new("保留批注"))?;
    sheet.merge_range(1, 0, 1, 1, "保留合并", &retained_format)?;
    sheet.set_row_height(0, 24.0)?;
    sheet.set_column_width(0, 18.0)?;
    workbook.save(path)?;
    Ok(())
}

fn run_extract_cli(
    config: &Path,
    output_dir: &Path,
    log_dir: &Path,
    append_existing: bool,
    capability: Option<&Path>,
    platform: Option<&str>,
    source: &Path,
) -> std::process::Output {
    let mut command = Command::new(env!("CARGO_BIN_EXE_pytools"));
    command
        .arg("extract-sheets")
        .arg("--config")
        .arg(config)
        .arg("--output-dir")
        .arg(output_dir)
        .arg("--log-dir")
        .arg(log_dir)
        .arg("--log-view")
        .arg("developer");
    if append_existing {
        command.arg("--append-existing");
    }
    if let Some(capability) = capability {
        command.arg("--capabilities").arg(capability);
    }
    if let Some(platform) = platform {
        command.arg("--platform").arg(platform);
    }
    command.arg(source).output().unwrap()
}

fn worksheet_xml(path: &Path) -> anyhow::Result<String> {
    worksheet_xml_at(path, 1)
}

fn worksheet_xml_at(path: &Path, sheet_index: usize) -> anyhow::Result<String> {
    let file = fs::File::open(path)?;
    let mut archive = zip::ZipArchive::new(file)?;
    let part_name = format!("xl/worksheets/sheet{sheet_index}.xml");
    let mut part = archive.by_name(&part_name)?;
    let mut xml = String::new();
    part.read_to_string(&mut xml)?;
    Ok(xml)
}

fn style_bold(book: &excel_editor::Workbook, sheet: &str, address: &str) -> anyhow::Result<bool> {
    Ok(book
        .sheet_by_name(sheet)?
        .cell(address)
        .and_then(|cell| cell.style().font())
        .is_some_and(|font| font.bold()))
}

fn style_fill(
    book: &excel_editor::Workbook,
    sheet: &str,
    address: &str,
) -> anyhow::Result<Option<String>> {
    Ok(book
        .sheet_by_name(sheet)?
        .cell(address)
        .and_then(|cell| cell.style().fill())
        .and_then(|fill| fill.pattern_fill())
        .and_then(|fill| fill.foreground_color())
        .map(|color| color.argb_str().to_ascii_uppercase()))
}

fn cell_number_format(
    book: &excel_editor::Workbook,
    sheet: &str,
    address: &str,
) -> anyhow::Result<Option<String>> {
    Ok(book
        .sheet_by_name(sheet)?
        .cell(address)
        .and_then(|cell| cell.style().numbering_format())
        .map(|format| format.format_code().to_owned()))
}

fn excel_cell(row: usize, column: usize) -> String {
    let mut column = column;
    let mut letters = String::new();
    while column > 0 {
        column -= 1;
        letters.insert(0, (b'A' + (column % 26) as u8) as char);
        column /= 26;
    }
    format!("{letters}{row}")
}

fn close_enough(left: f64, right: f64) -> bool {
    if left == right {
        return true;
    }
    let scale = left.abs().max(right.abs());
    scale > 0.0 && (left - right).abs() <= 1e-15 * scale
}

#[test]
fn floating_point_comparison_uses_relative_tolerance() {
    let value = 0.125_f64;
    assert!(close_enough(value, f64::from_bits(value.to_bits() + 1)));
    assert!(!close_enough(value, f64::from_bits(value.to_bits() + 8)));
}

fn assert_cell(expected: &Value, actual: Option<&CellValue>, address: &str, sheet: &str) {
    let expected_type = expected["type"].as_str();
    match (expected_type, actual) {
        (None, None) => {}
        (None, Some(CellValue::Empty)) => {}
        (Some("string"), None) if expected["value"].as_str() == Some("") => {}
        (Some("string"), Some(CellValue::Text(actual))) => {
            assert_eq!(
                actual,
                expected["value"].as_str().unwrap(),
                "{sheet}!{address}"
            );
        }
        (Some("integer" | "number"), Some(CellValue::Integer(actual))) => {
            let expected = expected["value"].as_f64().unwrap();
            assert!(
                close_enough(expected, *actual as f64),
                "{sheet}!{address}: expected {expected}, got {actual}"
            );
        }
        (Some("integer" | "number"), Some(CellValue::Number(actual))) => {
            let expected = expected["value"].as_f64().unwrap();
            assert!(
                close_enough(expected, *actual),
                "{sheet}!{address}: expected {expected}, got {actual}"
            );
        }
        (Some("boolean"), Some(CellValue::Boolean(actual))) => {
            assert_eq!(
                *actual,
                expected["value"].as_bool().unwrap(),
                "{sheet}!{address}"
            );
        }
        (Some("string"), None) if expected["value"].as_str().unwrap().trim().is_empty() => {}
        (kind, value) => panic!("{sheet}!{address}: unexpected {kind:?} / {value:?}"),
    }
}

#[test]
fn area_extraction_matches_python_baseline_and_preserves_inputs() {
    let cases = cases();
    let baseline = baseline();
    assert_eq!(baseline["status"], "passed");
    let parity_report_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("probe/reports/sheet_extract_area_parity.json");
    if parity_report_path.exists() {
        fs::remove_file(&parity_report_path).unwrap();
    }
    let source_paths = cases["input_files"]
        .as_array()
        .unwrap()
        .iter()
        .map(|name| fixture_path(name.as_str().unwrap()))
        .collect::<Vec<_>>();
    let input_bytes = source_paths
        .iter()
        .map(|path| fs::read(path).unwrap())
        .collect::<Vec<_>>();
    let root = temporary_directory("sheet-extract-area");
    let output_dir = root.join("output");
    let result = run_extract_sheets_area_candidate(
        &source_paths,
        &load_tasks(&cases),
        &output_dir,
        false,
        Some("unused-default.xlsx"),
    )
    .unwrap();

    for key in ["files", "extracted_sheets", "output_files", "failed_files"] {
        assert_eq!(
            result_value(&result, key),
            cases["expected"][key].as_u64().unwrap() as usize
        );
        assert_eq!(
            result_value(&result, key),
            baseline["stats"][key].as_u64().unwrap() as usize
        );
    }
    assert_eq!(result.paths.len(), 1);
    assert_eq!(
        Path::new(&result.paths[0])
            .file_name()
            .unwrap()
            .to_string_lossy(),
        cases["expected"]["output_file"].as_str().unwrap()
    );

    let output_sheets = read_value_columns_with_iso_dates(&result.paths[0], 16_384).unwrap();
    let expected_sheets = baseline["worksheets"].as_array().unwrap();
    assert_eq!(output_sheets.len(), expected_sheets.len());
    for (actual, expected) in output_sheets.iter().zip(expected_sheets) {
        let expected_name = expected["name"].as_str().unwrap();
        assert_eq!(actual.name, expected_name);
        assert_eq!(
            actual.max_row,
            expected["max_row"].as_u64().unwrap() as usize
        );
        assert_eq!(
            actual.max_column,
            expected["max_column"].as_u64().unwrap() as usize
        );
        for (row_index, row) in expected["values"].as_array().unwrap().iter().enumerate() {
            for (column_index, expected_cell) in row.as_array().unwrap().iter().enumerate() {
                let address = excel_cell(row_index + 1, column_index + 1);
                assert_cell(
                    expected_cell,
                    actual.values.get(&address),
                    &address,
                    expected_name,
                );
            }
        }
    }

    for (path, bytes) in source_paths.iter().zip(input_bytes) {
        assert_eq!(
            fs::read(path).unwrap(),
            bytes,
            "input changed: {}",
            path.display()
        );
    }
    let compared_cells = expected_sheets
        .iter()
        .map(|sheet| {
            sheet["values"]
                .as_array()
                .unwrap()
                .iter()
                .map(|row| row.as_array().unwrap().len())
                .sum::<usize>()
        })
        .sum::<usize>();
    let report = json!({
        "schema_version": 1,
        "status": "passed",
        "suite": "pytools.sheet_extract_area_parity",
        "baseline": "probe/reports/sheet_extract_area_python_baseline.json",
        "rust_operation": "run_extract_sheets_area_candidate",
        "stats": {
            "files": result.files,
            "extracted_sheets": result.extracted_sheets,
            "output_files": result.output_files,
            "failed_files": result.failed_files,
        },
        "worksheet_count": output_sheets.len(),
        "compared_cells": compared_cells,
        "numeric_comparison": "relative tolerance <= 1e-15",
        "checks": {
            "statistics": "preserved",
            "worksheet_names": "preserved",
            "worksheet_dimensions": "preserved",
            "cell_values_and_types": "preserved",
            "input_bytes": "preserved",
        }
    });
    let report_bytes = serde_json::to_vec_pretty(&report).unwrap();
    let mut report_with_newline = report_bytes;
    report_with_newline.push(b'\n');
    fs::write(&parity_report_path, report_with_newline).unwrap();
    fs::remove_dir_all(root).unwrap();
}

#[test]
fn interactive_menu_warns_and_does_not_publish_for_invalid_area_configuration() {
    let cases = cases();
    let root = temporary_directory("sheet-extract-invalid-area-menu");
    let mut tasks = load_tasks(&cases);
    tasks.truncate(1);
    tasks[0].sheet_keywords.clear();
    tasks[0].area_spec = "NOT_A_RANGE".to_owned();
    let config = create_menu_config(&root, &tasks);
    let source = fixture_path("省级月度报表_A.xlsx");
    let source_before = fs::read(&source).unwrap();

    let input = format!("5\n6\nn\n{}\n", source.display());
    let output = run_menu_with_view(&config, &input, "user");
    assert!(
        output.status.success(),
        "menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(
        stderr.contains("提取区域") && stderr.contains("NOT_A_RANGE"),
        "invalid area warning not shown; stdout={}; stderr={}",
        String::from_utf8_lossy(&output.stdout),
        stderr
    );

    let output_path = root.join("app/output/区域提取/area-result.xlsx");
    assert!(
        !output_path.exists(),
        "invalid area configuration must not publish a blank workbook"
    );
    assert_eq!(fs::read(&source).unwrap(), source_before);

    let log_dir = root.join("app/logs");
    let history_path = fs::read_dir(&log_dir)
        .unwrap()
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .find(|path| {
            path.file_name()
                .and_then(|name| name.to_str())
                .is_some_and(|name| name.starts_with("PT-7-6_"))
        })
        .unwrap();
    let history: Value = serde_json::from_slice(&fs::read(history_path).unwrap()).unwrap();
    assert_eq!(history["result"]["output_files"], 0);
    assert!(history["events"].as_array().unwrap().iter().any(|event| {
        event["phase"] == "area_parse"
            && event["status"] == "warning"
            && event["error_type"] == "invalid_area_spec"
            && event["user_message"]
                .as_str()
                .unwrap()
                .contains("NOT_A_RANGE")
    }));
    fs::remove_dir_all(root).unwrap();
}

fn result_value(result: &pytools_core::sheet_extract::ExtractSheetsAreaResult, key: &str) -> usize {
    match key {
        "files" => result.files,
        "extracted_sheets" => result.extracted_sheets,
        "output_files" => result.output_files,
        "failed_files" => result.failed_files,
        _ => panic!("unknown result field {key}"),
    }
}

#[test]
fn cli_loads_config_and_runs_the_area_extraction_candidate() {
    let cases = cases();
    let root = temporary_directory("sheet-extract-cli");
    let output_dir = root.join("output");
    let mut command = Command::new(env!("CARGO_BIN_EXE_pytools"));
    command
        .arg("extract-sheets-area")
        .arg("--config")
        .arg(fixture_path("extract_config.xlsx"))
        .arg("--output-dir")
        .arg(&output_dir);
    for name in cases["input_files"].as_array().unwrap() {
        command.arg(fixture_path(name.as_str().unwrap()));
    }
    let output = command.output().unwrap();
    assert!(
        output.status.success(),
        "CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report: Value = serde_json::from_slice(&output.stdout).unwrap();
    assert_eq!(report["files"], cases["expected"]["files"]);
    assert_eq!(
        report["extracted_sheets"],
        cases["expected"]["extracted_sheets"]
    );
    assert_eq!(report["output_files"], cases["expected"]["output_files"]);
    assert_eq!(report["failed_files"], cases["expected"]["failed_files"]);
    assert!(output_dir.join("区域提取/area-result.xlsx").exists());
    fs::remove_dir_all(root).unwrap();
}

#[test]
fn interactive_menu_routes_area_extraction_matches_python_and_records_engines() {
    let cases = cases();
    let root = temporary_directory("sheet-extract-menu");
    let tasks = load_tasks(&cases);
    let config = create_menu_config(&root, &tasks);
    let mut sources = cases["input_files"]
        .as_array()
        .unwrap()
        .iter()
        .map(|name| fixture_path(name.as_str().unwrap()))
        .collect::<Vec<_>>();
    let unsupported = root.join("not-validated.xls");
    fs::write(&unsupported, b"synthetic legacy workbook placeholder").unwrap();
    sources.push(unsupported.clone());
    let original_inputs = sources
        .iter()
        .map(|path| (path.clone(), fs::read(path).unwrap()))
        .collect::<Vec<_>>();
    let input = format!(
        "5\n6\nn\n{}\n",
        sources
            .iter()
            .map(|path| path.to_string_lossy())
            .collect::<Vec<_>>()
            .join(";")
    );

    let output = run_menu(&config, &input);
    assert!(
        output.status.success(),
        "menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stdout.contains("6  工作表数据提取（普通无批注版）"));
    assert!(stdout.contains("PT-7-6 当前仅验收 .xlsx"));
    assert!(stdout.contains("\"files\": 2"));
    assert!(stdout.contains("\"extracted_sheets\": 4"));
    assert!(stdout.contains("\"output_files\": 1"));
    assert!(stdout.contains("\"failed_files\": 1"));

    let output_path = root.join("app/output/区域提取/area-result.xlsx");
    assert!(output_path.is_file());
    let actual_sheets = read_value_columns_with_iso_dates(&output_path, 16_384).unwrap();
    let expected_sheets = baseline()["worksheets"].as_array().unwrap().clone();
    assert_eq!(actual_sheets.len(), expected_sheets.len());
    for (actual, expected) in actual_sheets.iter().zip(&expected_sheets) {
        let expected_name = expected["name"].as_str().unwrap();
        assert_eq!(actual.name, expected_name);
        assert_eq!(
            actual.max_row,
            expected["max_row"].as_u64().unwrap() as usize
        );
        assert_eq!(
            actual.max_column,
            expected["max_column"].as_u64().unwrap() as usize
        );
        for (row_index, row) in expected["values"].as_array().unwrap().iter().enumerate() {
            for (column_index, expected_cell) in row.as_array().unwrap().iter().enumerate() {
                let address = excel_cell(row_index + 1, column_index + 1);
                assert_cell(
                    expected_cell,
                    actual.values.get(&address),
                    &address,
                    expected_name,
                );
            }
        }
    }
    for (path, original) in original_inputs {
        assert_eq!(
            fs::read(&path).unwrap(),
            original,
            "input changed: {}",
            path.display()
        );
    }

    let events = stderr
        .lines()
        .filter(|line| line.starts_with('{'))
        .filter_map(|line| serde_json::from_str::<Value>(line).ok())
        .collect::<Vec<_>>();
    for phase in [
        "source_read",
        "output_write",
        "output_validation",
        "output_commit",
    ] {
        assert!(
            events.iter().any(|event| event["phase"] == phase),
            "missing {phase} event"
        );
    }
    assert!(events
        .iter()
        .any(|event| event["actual_engine"] == "calamine"));
    assert!(events
        .iter()
        .any(|event| event["actual_engine"] == "rust_xlsxwriter"));
    assert!(events
        .iter()
        .any(|event| { event["phase"] == "source_read" && event["status"] == "failed" }));
    assert!(events
        .iter()
        .any(|event| { event["phase"] == "source_read" && event["stage_elapsed_ms"].is_number() }));

    let log_dir = root.join("app/logs");
    let history_path = fs::read_dir(&log_dir)
        .unwrap()
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .find(|path| {
            path.file_name()
                .unwrap()
                .to_string_lossy()
                .starts_with("PT-7-6_")
        })
        .unwrap();
    let history_bytes = fs::read(&history_path).unwrap();
    let history_text = String::from_utf8_lossy(&history_bytes);
    let history: Value = serde_json::from_slice(&history_bytes).unwrap();
    assert_eq!(history["feature_id"], "PT-7-6");
    assert_eq!(history["status"], "partial");
    assert_eq!(history["total_inputs"], 3);
    assert_eq!(history["result"]["extracted_sheets"], 4);
    let actual_engines = history["actual_engines"].as_array().unwrap();
    for engine in [
        "calamine",
        "rust_xlsxwriter",
        "hard_link create-only publication",
    ] {
        assert!(
            actual_engines.iter().any(|actual| actual == engine),
            "run history is missing actual engine {engine:?}: {actual_engines:?}"
        );
    }
    let sample_cell_value = expected_sheets
        .iter()
        .flat_map(|sheet| sheet["values"].as_array().unwrap())
        .flat_map(|row| row.as_array().unwrap())
        .find_map(|cell| {
            (cell["type"] == "string")
                .then(|| cell["value"].as_str())
                .flatten()
                .filter(|value| !value.is_empty())
        })
        .unwrap();
    assert!(
        !history_text.contains(sample_cell_value),
        "run history must not contain raw cell values"
    );

    let compared_cells = expected_sheets
        .iter()
        .map(|sheet| {
            sheet["values"]
                .as_array()
                .unwrap()
                .iter()
                .map(|row| row.as_array().unwrap().len())
                .sum::<usize>()
        })
        .sum::<usize>();
    let report = json!({
        "schema_version": 1,
        "status": "passed",
        "suite": "pytools.sheet_extract_area_menu_parity",
        "entry_point": "pytools menu -> 7 -> 6",
        "baseline": "probe/reports/sheet_extract_area_python_baseline.json",
        "inputs": {
            "xlsx": 3,
            "filtered_unverified_formats": 1,
            "damaged_xlsx_expected_failures": 1,
        },
        "stats": {
            "files": 2,
            "extracted_sheets": 4,
            "output_files": 1,
            "failed_files": 1,
        },
        "worksheet_count": actual_sheets.len(),
        "compared_cells": compared_cells,
        "actual_engines": ["calamine", "rust_xlsxwriter"],
        "checks": {
            "menu_task_loading_and_file_selection": "passed",
            "unsupported_format_filtered_before_reader": "passed",
            "damaged_xlsx_reported_as_partial_failure": "passed",
            "worksheet_names_dimensions_values_and_types": "preserved",
            "source_input_bytes": "preserved",
            "run_history_excludes_cell_values": "passed",
            "production_capability_created": false,
        }
    });
    let report_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("probe/reports/sheet_extract_menu_parity.json");
    fs::write(report_path, serde_json::to_vec_pretty(&report).unwrap()).unwrap();
    fs::remove_dir_all(root).unwrap();
}

#[test]
fn interactive_menu_writes_area_extraction_to_configured_target_sheet() -> anyhow::Result<()> {
    let root = temporary_directory("sheet-extract-configured-target-sheet");
    let source = root.join("app/input/target-routing.xlsx");
    fs::create_dir_all(source.parent().unwrap())?;
    let mut source_workbook = rust_xlsxwriter::Workbook::new();
    let source_sheet = source_workbook.add_worksheet();
    source_sheet.set_name("Report")?;
    source_sheet.write_string(0, 0, "机构")?;
    source_sheet.write_string(0, 1, "金额")?;
    source_sheet.write_string(1, 0, "东区")?;
    source_sheet.write_number(1, 1, 25.0)?;
    source_workbook.save(&source)?;
    let source_before = fs::read(&source)?;

    let task = ExtractTask {
        row_no: 2,
        workbook_keywords: "target-routing".to_owned(),
        sheet_keywords: "Report".to_owned(),
        exclude_keywords: String::new(),
        area_spec: "A1:B2".to_owned(),
        stitch_mode: String::new(),
        output_folder: "mapped".to_owned(),
        output_name: "target-workbook.xlsx".to_owned(),
        target_sheet: "人工指定结果".to_owned(),
        remark: String::new(),
    };
    let config = create_menu_config(&root, &[task]);
    let input = format!("5\n6\nn\n{}\n", source.display());
    let output = run_menu(&config, &input);
    assert!(
        output.status.success(),
        "PT-7-6 menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let target = root.join("app/output/mapped/target-workbook.xlsx");
    let workbook = read_workbook(&target)?;
    assert_eq!(
        workbook
            .sheets
            .iter()
            .map(|sheet| sheet.name.as_str())
            .collect::<Vec<_>>(),
        ["人工指定结果"]
    );
    assert_eq!(workbook.sheets[0].values["A1"]["value"], "机构");
    assert_eq!(workbook.sheets[0].values["A2"]["value"], "东区");
    assert_eq!(
        workbook.sheets[0].values["B2"]["value"].as_f64(),
        Some(25.0)
    );
    assert_eq!(fs::read(&source)?, source_before);
    fs::remove_dir_all(root)?;
    Ok(())
}

#[test]
fn interactive_menu_routes_full_sheet_import_against_python_helper_baseline() -> anyhow::Result<()>
{
    const PLATFORM: &str = "synthetic-pt-7-6-full-sheet-menu";
    let baseline = full_sheet_baseline();
    assert_eq!(baseline["status"], "partial");
    assert!(baseline["worksheet"]["style_checks"]["python_output"].is_null());
    assert!(
        baseline["worksheet"]["style_checks"]["python_output_read_error"]
            .as_str()
            .is_some_and(|error| error.contains("IndexError"))
    );

    let source = full_sheet_fixture_path("source.xlsx");
    let original_source = fs::read(&source)?;
    let root = temporary_directory("sheet-extract-full-menu");
    let mut tasks = load_tasks(&cases());
    tasks.truncate(1);
    tasks[0].workbook_keywords = "source".to_owned();
    tasks[0].sheet_keywords = "报表".to_owned();
    tasks[0].exclude_keywords.clear();
    tasks[0].area_spec.clear();
    tasks[0].stitch_mode = "horizontal".to_owned();
    tasks[0].output_folder = "整表导入".to_owned();
    tasks[0].output_name = "full-sheet-output.xlsx".to_owned();
    let config = create_menu_config(&root, &tasks);
    let capability = write_full_sheet_test_capability(&root, &source, PLATFORM)?;
    let output = run_menu_with_capability(
        &config,
        &format!("5\n6\nn\n{}\n", source.display()),
        "developer",
        &capability,
        PLATFORM,
    );
    assert!(
        output.status.success(),
        "full-sheet menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("6  工作表数据提取（普通无批注版）"));
    assert!(stdout.contains("\"extracted_sheets\": 1"));
    assert!(stdout.contains("\"output_files\": 1"));
    assert!(stdout.contains("\"failed_files\": 0"));

    let output_path = root.join("app/output/整表导入/full-sheet-output.xlsx");
    assert!(output_path.is_file());
    let expected_sheet = &baseline["worksheet"];
    let expected_name = expected_sheet["name"].as_str().unwrap();
    let expected_values = expected_sheet["values"].as_array().unwrap();
    let actual_values = read_value_columns_with_iso_dates(&output_path, 16_384)?;
    assert_eq!(actual_values.len(), 1);
    assert_eq!(actual_values[0].name, expected_name);
    for (row_index, row) in expected_values.iter().enumerate() {
        for (column_index, expected_cell) in row.as_array().unwrap().iter().enumerate() {
            let address = excel_cell(row_index + 1, column_index + 1);
            assert_cell(
                expected_cell,
                actual_values[0].values.get(&address),
                &address,
                expected_name,
            );
        }
    }
    let actual_extents = read_xlsx_worksheet_extents(&output_path)?;
    assert_eq!(actual_extents.len(), 1);
    assert_eq!(actual_extents[0].name, expected_name);
    assert_eq!(
        actual_extents[0].max_row,
        expected_sheet["max_row"].as_u64().unwrap() as usize
    );
    assert_eq!(
        actual_extents[0].max_column,
        expected_sheet["max_column"].as_u64().unwrap() as usize
    );

    let source_workbook = open_xlsx(&source)?;
    let output_workbook = open_xlsx(&output_path)?;
    let source_sheet = source_workbook
        .sheet_by_name(expected_name)
        .context("source worksheet missing")?;
    let output_sheet = output_workbook
        .sheet_by_name(expected_name)
        .context("full-sheet output missing its source worksheet")?;
    let expected_merges = expected_sheet["merged_ranges"]
        .as_array()
        .unwrap()
        .iter()
        .map(|value| value.as_str().unwrap().to_owned())
        .collect::<Vec<_>>();
    assert_eq!(
        output_sheet
            .merge_cells()
            .iter()
            .map(|range| range.range())
            .collect::<Vec<_>>(),
        expected_merges
    );
    let source_styles = &expected_sheet["style_checks"]["source_workbook_expected"];
    assert_eq!(
        style_bold(&source_workbook, expected_name, "A1")?,
        source_styles["A1_bold"].as_bool().unwrap()
    );
    assert_eq!(
        style_bold(&output_workbook, expected_name, "A1")?,
        source_styles["A1_bold"].as_bool().unwrap()
    );
    for (address, expected_key) in [("A1", "A1_fill"), ("E8", "E8_fill")] {
        let expected_fill = source_styles[expected_key].as_str();
        assert_eq!(
            style_fill(&source_workbook, expected_name, address)?.as_deref(),
            expected_fill,
            "source fill at {expected_name}!{address}"
        );
        assert_eq!(
            style_fill(&output_workbook, expected_name, address)?.as_deref(),
            expected_fill,
            "copied fill at {expected_name}!{address}"
        );
    }
    for (address, expected_key) in [("B3", "B3_number_format"), ("C3", "C3_number_format")] {
        let expected_format = source_styles[expected_key].as_str();
        assert_eq!(
            cell_number_format(&source_workbook, expected_name, address)?.as_deref(),
            expected_format,
            "source number format at {expected_name}!{address}"
        );
        assert_eq!(
            cell_number_format(&output_workbook, expected_name, address)?.as_deref(),
            expected_format,
            "copied number format at {expected_name}!{address}"
        );
    }

    let source_row = source_sheet
        .row_dimension(12)
        .context("source extended row dimension missing")?;
    let output_row = output_sheet
        .row_dimension(12)
        .context("copied extended row dimension missing")?;
    assert_eq!(output_row.height(), source_row.height());
    assert_eq!(output_row.hidden(), source_row.hidden());
    assert_eq!(
        output_row.height(),
        expected_sheet["extended_row_dimension"]["height"]
            .as_f64()
            .unwrap()
    );
    assert_eq!(
        output_row.hidden(),
        expected_sheet["extended_row_dimension"]["hidden"]
            .as_bool()
            .unwrap()
    );
    let source_column = source_sheet
        .column_dimension("G")
        .context("source extended column dimension missing")?;
    let output_column = output_sheet
        .column_dimension("G")
        .context("copied extended column dimension missing")?;
    assert_eq!(output_column.width(), source_column.width());
    assert_eq!(output_column.hidden(), source_column.hidden());
    assert_eq!(
        output_column.width(),
        expected_sheet["extended_column_dimension"]["width"]
            .as_f64()
            .unwrap()
    );
    assert_eq!(
        output_column.hidden(),
        expected_sheet["extended_column_dimension"]["hidden"]
            .as_bool()
            .unwrap()
    );

    let source_comments = read_ooxml_comments(&source)?;
    assert!(source_comments
        .get(expected_name)
        .is_some_and(|comments| comments.contains_key("D5")));
    assert_eq!(expected_sheet["comment_copied"], false);
    let output_comments = read_ooxml_comments(&output_path)?;
    assert!(!output_comments
        .get(expected_name)
        .is_some_and(|comments| comments.contains_key("D5")));

    let output_xml = worksheet_xml(&output_path)?;
    let document = roxmltree::Document::parse(&output_xml)?;
    let row = document
        .descendants()
        .find(|node| node.tag_name().name() == "row" && node.attribute("r") == Some("12"))
        .context("serialized row 12 not found")?;
    assert_eq!(row.attribute("ht"), Some("26"));
    assert_eq!(row.attribute("hidden"), Some("1"));
    assert_eq!(row.attribute("outlineLevel"), Some("1"));
    let column = document
        .descendants()
        .find(|node| {
            node.tag_name().name() == "col"
                && node.attribute("min") == Some("7")
                && node.attribute("max") == Some("7")
        })
        .context("serialized column G definition not found")?;
    assert_eq!(column.attribute("width"), Some("20"));
    assert_eq!(column.attribute("hidden"), Some("1"));
    assert_eq!(column.attribute("outlineLevel"), Some("2"));

    assert_eq!(fs::read(&source)?, original_source);
    let stderr = String::from_utf8_lossy(&output.stderr);
    let events = stderr
        .lines()
        .filter(|line| line.starts_with('{'))
        .filter_map(|line| serde_json::from_str::<Value>(line).ok())
        .collect::<Vec<_>>();
    for phase in [
        "source_read",
        "worksheet_extent_read",
        "skeleton_write",
        "excel_modifier",
        "full_sheet_validation",
        "output_validation",
        "output_commit",
    ] {
        assert!(
            events.iter().any(|event| event["phase"] == phase),
            "missing {phase} event"
        );
    }
    for engine in ["calamine", "rust_xlsxwriter", "umya_spreadsheet"] {
        assert!(
            events.iter().any(|event| event["actual_engine"] == engine),
            "missing actual engine {engine}"
        );
    }

    let log_dir = root.join("app/logs");
    let history_path = fs::read_dir(&log_dir)?
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .find(|path| {
            path.file_name()
                .unwrap()
                .to_string_lossy()
                .starts_with("PT-7-6_")
        })
        .context("full-sheet menu run history was not saved")?;
    let history_bytes = fs::read(&history_path)?;
    let history_text = String::from_utf8_lossy(&history_bytes);
    let history: Value = serde_json::from_slice(&history_bytes)?;
    assert_eq!(history["feature_id"], "PT-7-6");
    assert_eq!(history["status"], "succeeded");
    assert_eq!(history["total_inputs"], 1);
    assert_eq!(history["result"]["extracted_sheets"], 1);
    assert_eq!(history["result"]["output_files"], 1);
    let actual_engines = history["actual_engines"].as_array().unwrap();
    for engine in ["calamine", "rust_xlsxwriter", "umya_spreadsheet"] {
        assert!(actual_engines.iter().any(|actual| actual == engine));
    }
    let sample_cell_value = expected_values
        .iter()
        .flat_map(|row| row.as_array().unwrap())
        .find_map(|cell| cell["value"].as_str())
        .context("full-sheet Python baseline has no text cell")?;
    assert!(!history_text.contains(sample_cell_value));

    let compared_cells = expected_values
        .iter()
        .map(|row| row.as_array().unwrap().len())
        .sum::<usize>();
    let report = json!({
        "schema_version": 1,
        "status": "partial",
        "suite": "pytools.sheet_extract.full_sheet_menu_parity",
        "entry_point": "pytools menu -> 7 -> 6",
        "baseline": "PyCode/core/feature_1_6_extract_sheets.py::_copy_full_sheet_with_style",
        "baseline_style_limitation": baseline["limitations"],
        "input_format": "xlsx",
        "temporary_test_capability": {
            "engine": "umya_spreadsheet",
            "platform": PLATFORM,
            "production_capability_created": false
        },
        "result": {
            "files": 1,
            "extracted_sheets": 1,
            "output_files": 1,
            "failed_files": 0,
            "sheet": expected_name,
            "cell_extent": [expected_sheet["max_row"], expected_sheet["max_column"]],
            "compared_value_cells": compared_cells,
            "merged_ranges": expected_merges,
            "extended_row_dimension": expected_sheet["extended_row_dimension"],
            "extended_column_dimension": expected_sheet["extended_column_dimension"],
            "copied_comments": false,
            "actual_engines": ["calamine", "rust_xlsxwriter", "umya_spreadsheet"]
        },
        "checks": {
            "menu_task_loading_and_source_selection": "passed",
            "python_helper_values_types_extent_merges_and_dimensions": "matched",
            "rust_output_styles_against_source_properties": "matched",
            "ordinary_extraction_drops_comments": "passed",
            "source_bytes": "preserved",
            "developer_events_and_history": "passed",
            "history_excludes_cell_values": "passed",
            "python_helper_output_style_reopen": baseline["worksheet"]["style_checks"]["python_output_read_error"],
            "production_capability_created": false
        }
    });
    let report_path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("probe/reports/sheet_extract_full_menu_parity.json");
    fs::write(report_path, serde_json::to_vec_pretty(&report)?)?;
    fs::remove_dir_all(root)?;
    Ok(())
}

#[test]
fn interactive_menu_stops_when_extract_tasks_are_missing_or_full_sheet_capability_is_unavailable() {
    let cases = cases();
    let root = temporary_directory("sheet-extract-menu-gates");
    let no_tasks_config = create_menu_config(&root.join("no-tasks"), &[]);
    let no_tasks = run_menu(&no_tasks_config, "5\n6\n");
    assert!(no_tasks.status.success());
    let no_tasks_stdout = String::from_utf8_lossy(&no_tasks.stdout);
    assert!(no_tasks_stdout.contains("配置无启用任务"));
    assert!(!no_tasks_stdout.contains("文件路径>"));
    assert!(!root.join("no-tasks/app/output").exists());

    let mut full_sheet_tasks = load_tasks(&cases);
    full_sheet_tasks[0].area_spec.clear();
    let full_sheet_config = create_menu_config(&root.join("full-sheet"), &full_sheet_tasks);
    let full_sheet = run_menu(&full_sheet_config, "5\n6\n");
    assert!(full_sheet.status.success());
    let full_sheet_stdout = String::from_utf8_lossy(&full_sheet.stdout);
    assert!(full_sheet_stdout.contains("整表提取任务"));
    assert!(full_sheet_stdout.contains("没有匹配的 ExcelModifier 能力文件与平台"));
    assert!(!full_sheet_stdout.contains("文件路径>"));
    assert!(!root.join("full-sheet/app/output").exists());
    fs::remove_dir_all(root).unwrap();
}

#[test]
fn interactive_menu_appends_only_after_confirmation_and_with_matching_capability(
) -> anyhow::Result<()> {
    const PLATFORM: &str = "synthetic-pt-7-6-append-menu";
    let root = temporary_directory("sheet-extract-append-menu");
    let source = root.join("menu-source.xlsx");
    write_append_source(&source, "新增工作表")?;
    let output_dir = root.join("app/output");
    fs::create_dir_all(&output_dir)?;
    let output_path = output_dir.join("合并结果.xlsx");
    write_existing_full_sheet_output(&output_path)?;
    let original_output = fs::read(&output_path)?;
    let original_source = fs::read(&source)?;

    let task = ExtractTask {
        row_no: 2,
        workbook_keywords: "menu-source".to_owned(),
        sheet_keywords: String::new(),
        exclude_keywords: String::new(),
        area_spec: "A1:B1".to_owned(),
        stitch_mode: String::new(),
        output_folder: String::new(),
        output_name: "合并结果.xlsx".to_owned(),
        target_sheet: String::new(),
        remark: String::new(),
    };
    let config = create_menu_config(&root, &[task]);

    let refused = run_menu(&config, "5\n6\nyes\n");
    assert!(
        refused.status.success(),
        "menu should report the missing capability without launching extraction: {}",
        String::from_utf8_lossy(&refused.stderr)
    );
    let refused_stdout = String::from_utf8_lossy(&refused.stdout);
    assert!(refused_stdout.contains("追加到已存在的输出工作簿需要 --capabilities 与 --platform"));
    assert!(!refused_stdout.contains("文件路径>"));
    assert_eq!(fs::read(&output_path)?, original_output);

    let capability = write_append_test_capability(&root, &output_path, &source, PLATFORM)?;
    let appended = run_menu_with_capability(
        &config,
        &format!("5\n6\nyes\n{}\n", source.display()),
        "developer",
        &capability,
        PLATFORM,
    );
    assert!(
        appended.status.success(),
        "menu append failed: {}",
        String::from_utf8_lossy(&appended.stderr)
    );
    let stdout = String::from_utf8_lossy(&appended.stdout);
    let stderr = String::from_utf8_lossy(&appended.stderr);
    assert!(stdout.contains("[追加模式]"));
    assert!(stdout.contains("\"output_files\": 1"));
    assert!(stderr.contains("\"phase\":\"excel_modifier\""));
    assert!(stderr.contains("\"actual_engine\":\"umya_spreadsheet\""));

    let result = read_workbook(&output_path)?;
    assert_eq!(
        result
            .sheets
            .iter()
            .map(|sheet| sheet.name.as_str())
            .collect::<Vec<_>>(),
        ["第一张表", "新增工作表"]
    );
    assert_eq!(result.sheets[0].values["A1"]["value"], "保留值");
    assert!(result.sheets[0].formulas["B1"].contains("1+1"));
    assert_eq!(result.sheets[0].merges, ["A2:B2"]);
    assert!(result.sheets[0].comments["A1"].contains("保留批注"));
    assert_eq!(result.sheets[1].values["A1"]["value"], "R1C1");
    assert_eq!(result.sheets[1].values["B1"]["value"], "R1C2");
    assert_eq!(fs::read(&source)?, original_source);

    let history_path = fs::read_dir(root.join("app/logs"))?
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .find(|path| {
            path.file_name()
                .and_then(|name| name.to_str())
                .is_some_and(|name| name.starts_with("PT-7-6_"))
        })
        .context("PT-7-6 append run history was not written")?;
    let history = fs::read_to_string(history_path)?;
    assert!(history.contains("umya_spreadsheet"));
    assert!(!history.contains("R1C1"));

    fs::remove_dir_all(root)?;
    Ok(())
}

#[test]
fn interactive_menu_supports_user_and_off_log_views() {
    let cases = cases();
    let tasks = load_tasks(&cases);
    let input_path = fixture_path("省级月度报表_A.xlsx");

    let user_root = temporary_directory("sheet-extract-menu-user-view");
    let user_config = create_menu_config(&user_root, &tasks);
    let user_output = run_menu_with_view(
        &user_config,
        &format!("5\n6\nn\n{}\n", input_path.display()),
        "user",
    );
    assert!(user_output.status.success());
    let user_stdout = String::from_utf8_lossy(&user_output.stdout);
    let user_stderr = String::from_utf8_lossy(&user_output.stderr);
    assert!(user_stdout.contains("[完成] 源文件=1 提取工作表=2 输出文件=1 失败=0"));
    assert!(user_stderr.contains("正在读取源工作簿"));
    assert!(user_stderr.contains("工作表提取运行记录已保存"));
    assert!(user_root.join("app/logs").is_dir());
    fs::remove_dir_all(user_root).unwrap();

    let off_root = temporary_directory("sheet-extract-menu-off-view");
    let off_config = create_menu_config(&off_root, &tasks);
    let off_output = run_menu_with_view(
        &off_config,
        &format!("5\n6\nn\n{}\n", input_path.display()),
        "off",
    );
    assert!(off_output.status.success());
    let off_stdout = String::from_utf8_lossy(&off_output.stdout);
    assert!(off_stdout.contains("\"files\": 1"));
    assert!(String::from_utf8_lossy(&off_output.stderr).is_empty());
    assert!(!off_root.join("app/logs").exists());
    fs::remove_dir_all(off_root).unwrap();
}

#[test]
fn append_cli_keeps_previous_python_extraction_sheets() -> anyhow::Result<()> {
    const PLATFORM: &str = "synthetic-pt-7-6-append-cli";
    let root = temporary_directory("sheet-extract-append-cli");
    let source_first = root.join("2026第一批.xlsx");
    let source_second = root.join("2026第二批.xlsx");
    write_append_source(&source_first, "第一张表")?;
    write_append_source(&source_second, "第二张表")?;

    let task = ExtractTask {
        row_no: 2,
        workbook_keywords: "2026".to_owned(),
        sheet_keywords: String::new(),
        exclude_keywords: String::new(),
        area_spec: "A1:B1".to_owned(),
        stitch_mode: String::new(),
        output_folder: String::new(),
        output_name: "合并结果".to_owned(),
        target_sheet: String::new(),
        remark: String::new(),
    };
    let config = create_menu_config(&root, &[task]);
    let output_dir = root.join("app/output");
    let log_dir = root.join("app/logs");
    let output_path = output_dir.join("合并结果.xlsx");

    let first_run = run_extract_cli(
        &config,
        &output_dir,
        &log_dir,
        false,
        None,
        None,
        &source_first,
    );
    assert!(
        first_run.status.success(),
        "first extraction failed: {}",
        String::from_utf8_lossy(&first_run.stderr)
    );
    let first_report: Value = serde_json::from_slice(&first_run.stdout)?;
    assert_eq!(first_report["extracted_sheets"], 1);
    let original_target = fs::read(&output_path)?;
    let original_sheets = read_value_columns_with_iso_dates(&output_path, 16_384)?;
    assert_eq!(original_sheets.len(), 1);
    assert_eq!(original_sheets[0].name, "第一张表");
    assert_eq!(
        original_sheets[0].values.get("A1"),
        Some(&CellValue::Text("R1C1".to_owned()))
    );

    let without_capability = run_extract_cli(
        &config,
        &output_dir,
        &log_dir,
        true,
        None,
        None,
        &source_second,
    );
    assert!(!without_capability.status.success());
    assert!(String::from_utf8_lossy(&without_capability.stderr)
        .contains("requires --capabilities and --platform"));
    assert_eq!(fs::read(&output_path)?, original_target);

    let capability = write_append_test_capability(&root, &output_path, &source_second, PLATFORM)?;
    let mismatched_platform = run_extract_cli(
        &config,
        &output_dir,
        &log_dir,
        true,
        Some(&capability),
        Some("different-synthetic-platform"),
        &source_second,
    );
    assert!(!mismatched_platform.status.success());
    assert_eq!(fs::read(&output_path)?, original_target);

    let append_run = run_extract_cli(
        &config,
        &output_dir,
        &log_dir,
        true,
        Some(&capability),
        Some(PLATFORM),
        &source_second,
    );
    assert!(
        append_run.status.success(),
        "append extraction failed: {}",
        String::from_utf8_lossy(&append_run.stderr)
    );
    let append_report: Value = serde_json::from_slice(&append_run.stdout)?;
    assert_eq!(append_report["extracted_sheets"], 1);
    assert_eq!(append_report["output_files"], 1);
    let events = String::from_utf8_lossy(&append_run.stderr);
    assert!(events.contains("\"actual_engine\":\"umya_spreadsheet\""));
    assert!(events.contains("\"phase\":\"output_commit\""));

    let history_path = fs::read_dir(&log_dir)?
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .find(|path| {
            fs::read_to_string(path)
                .map(|history| {
                    history.contains("umya_spreadsheet") && history.contains("output_commit")
                })
                .unwrap_or(false)
        })
        .context("successful append run history was not saved")?;
    let history_bytes = fs::read(&history_path)?;
    let history_text = String::from_utf8_lossy(&history_bytes);
    let history: Value = serde_json::from_slice(&history_bytes)?;
    assert_eq!(history["feature_id"], "PT-7-6");
    assert_eq!(history["status"], "succeeded");
    assert_eq!(history["result"]["extracted_sheets"], 1);
    assert!(history["actual_engines"]
        .as_array()
        .is_some_and(|engines| engines.iter().any(|engine| engine == "umya_spreadsheet")));
    assert!(history["events"]
        .as_array()
        .is_some_and(|events| events.iter().any(|event| event["phase"] == "output_commit")));
    assert!(
        !history_text.contains("R1C1"),
        "run history must not contain raw cell values"
    );

    let appended_sheets = read_value_columns_with_iso_dates(&output_path, 16_384)?;
    assert_eq!(
        appended_sheets
            .iter()
            .map(|sheet| sheet.name.as_str())
            .collect::<Vec<_>>(),
        ["第一张表", "第二张表"]
    );
    for sheet in &appended_sheets {
        assert_eq!(
            sheet.values.get("A1"),
            Some(&CellValue::Text("R1C1".to_owned()))
        );
        assert_eq!(
            sheet.values.get("B1"),
            Some(&CellValue::Text("R1C2".to_owned()))
        );
    }
    fs::remove_dir_all(root)?;
    Ok(())
}

#[test]
fn append_cli_preserves_existing_workbook_properties_and_imports_full_sheet_layout(
) -> anyhow::Result<()> {
    const PLATFORM: &str = "synthetic-pt-7-6-full-sheet-append";
    let root = temporary_directory("sheet-extract-full-sheet-append");
    let source = root.join("2026第二批.xlsx");
    write_append_full_sheet_source(&source, "第二张表")?;
    let output_dir = root.join("app/output");
    fs::create_dir_all(&output_dir)?;
    let output_path = output_dir.join("合并结果.xlsx");
    write_existing_full_sheet_output(&output_path)?;

    let task = ExtractTask {
        row_no: 2,
        workbook_keywords: "2026".to_owned(),
        sheet_keywords: String::new(),
        exclude_keywords: String::new(),
        area_spec: "整表".to_owned(),
        stitch_mode: String::new(),
        output_folder: String::new(),
        output_name: "合并结果".to_owned(),
        target_sheet: String::new(),
        remark: String::new(),
    };
    let config = create_menu_config(&root, &[task]);
    let before = read_workbook(&output_path)?;
    assert_eq!(before.sheets[0].name, "第一张表");
    assert_eq!(before.sheets[0].values["A1"]["value"], "保留值");
    assert!(before.sheets[0].formulas["B1"].contains("1+1"));
    assert_eq!(before.sheets[0].merges, ["A2:B2"]);
    assert!(before.sheets[0].comments["A1"].contains("保留批注"));

    let capability = write_append_test_capability(&root, &output_path, &source, PLATFORM)?;
    let append_run = run_extract_cli(
        &config,
        &output_dir,
        &root.join("app/logs"),
        true,
        Some(&capability),
        Some(PLATFORM),
        &source,
    );
    assert!(
        append_run.status.success(),
        "full-sheet append failed: {}",
        String::from_utf8_lossy(&append_run.stderr)
    );
    let append_report: Value = serde_json::from_slice(&append_run.stdout)?;
    assert_eq!(append_report["extracted_sheets"], 1);
    let events = String::from_utf8_lossy(&append_run.stderr);
    assert!(events.contains("\"actual_engine\":\"umya_spreadsheet\""));
    assert!(events.contains("\"actual_engine\":\"rust_xlsxwriter\""));

    let after = read_workbook(&output_path)?;
    assert_eq!(
        after
            .sheets
            .iter()
            .map(|sheet| sheet.name.as_str())
            .collect::<Vec<_>>(),
        ["第一张表", "第二张表"]
    );
    assert_eq!(after.sheets[0].values["A1"]["value"], "保留值");
    assert!(after.sheets[0].formulas["B1"].contains("1+1"));
    assert_eq!(after.sheets[0].merges, ["A2:B2"]);
    assert!(after.sheets[0].comments["A1"].contains("保留批注"));
    assert_eq!(after.sheets[1].merges, ["A1:B1"]);
    assert!(after.sheets[1].formulas.is_empty());
    assert!(after.sheets[1].comments.is_empty());

    let workbook = open_xlsx(&output_path)?;
    assert!(style_bold(&workbook, "第一张表", "A1")?);
    assert!(style_bold(&workbook, "第二张表", "A1")?);
    let first_sheet_xml = worksheet_xml_at(&output_path, 1)?;
    let second_sheet_xml = worksheet_xml_at(&output_path, 2)?;
    assert!(first_sheet_xml.contains("ht=\"24\""));
    assert!(first_sheet_xml.contains("width=\"18.7109375\""));
    assert!(second_sheet_xml.contains("ht=\"30\""));
    assert!(second_sheet_xml.contains("width=\"20.7109375\""));

    fs::remove_dir_all(root)?;
    Ok(())
}

#[test]
fn append_mode_creates_a_missing_output_and_unsupported_operations_are_rejected_safely() {
    let cases = cases();
    let source = fixture_path("省级月度报表_A.xlsx");
    let root = temporary_directory("sheet-extract-reject");
    let output_dir = root.join("output");
    let mut loaded_tasks = load_tasks(&cases);

    let append_new_output_dir = root.join("append-new-output");
    let append_result = run_extract_sheets_area_candidate(
        std::slice::from_ref(&source),
        &loaded_tasks,
        &append_new_output_dir,
        true,
        Some("unused.xlsx"),
    )
    .unwrap();
    assert_eq!(append_result.output_files, 1);
    assert!(append_new_output_dir
        .join("区域提取/area-result.xlsx")
        .is_file());

    loaded_tasks[0].area_spec.clear();
    let full_sheet_error = run_extract_sheets_area_candidate(
        std::slice::from_ref(&source),
        &loaded_tasks,
        &output_dir,
        false,
        Some("unused.xlsx"),
    )
    .unwrap_err();
    assert!(full_sheet_error
        .to_string()
        .contains("full-sheet extraction"));
    loaded_tasks = load_tasks(&cases);

    let date_output = root.join("date-output");
    let source_bytes_before = fs::read(&source).unwrap();
    let mut date_task = loaded_tasks[0].clone();
    date_task.workbook_keywords.clear();
    date_task.sheet_keywords.clear();
    date_task.exclude_keywords.clear();
    date_task.area_spec = "E1".to_owned();
    date_task.output_name = "date-output.xlsx".to_owned();
    let date_result = run_extract_sheets_area_candidate(
        std::slice::from_ref(&source),
        &[date_task],
        &date_output,
        false,
        Some("unused.xlsx"),
    )
    .unwrap();
    assert_eq!(date_result.files, 1);
    assert_eq!(date_result.extracted_sheets, 1);
    assert_eq!(date_result.output_files, 1);
    let date_path = date_output.join("区域提取/date-output.xlsx");
    assert!(date_path.is_file());
    let date_sheets = read_value_columns_with_iso_dates(&date_path, 16_384).unwrap();
    assert_eq!(date_sheets.len(), 1);
    assert_eq!(
        date_sheets[0].values.get("A1"),
        Some(&CellValue::DateTimeIso(
            "2024-12-31T09:30:00.000".to_owned()
        ))
    );
    assert_eq!(fs::read(&source).unwrap(), source_bytes_before);

    let collision_dir = root.join("collision");
    let collision_path = collision_dir.join("区域提取/area-result.xlsx");
    fs::create_dir_all(collision_path.parent().unwrap()).unwrap();
    fs::write(&collision_path, b"existing output must survive").unwrap();
    let collision_error = run_extract_sheets_area_candidate(
        std::slice::from_ref(&source),
        &loaded_tasks,
        &collision_dir,
        false,
        Some("unused.xlsx"),
    )
    .unwrap_err();
    assert!(collision_error.to_string().contains("refuse to overwrite"));
    assert_eq!(
        fs::read(collision_path).unwrap(),
        b"existing output must survive"
    );

    fs::remove_dir_all(root).unwrap();
}

#[test]
fn interactive_menu_extracts_date_cells_as_excel_datetimes() {
    let cases = cases();
    let root = temporary_directory("sheet-extract-date-menu");
    let source = fixture_path("省级月度报表_A.xlsx");
    let source_bytes = fs::read(&source).unwrap();
    let mut task = load_tasks(&cases)[0].clone();
    task.workbook_keywords.clear();
    task.sheet_keywords = "月度资金报表".to_owned();
    task.exclude_keywords.clear();
    task.area_spec = "E1".to_owned();
    task.output_folder = "区域提取".to_owned();
    task.output_name = "date-menu-output.xlsx".to_owned();
    let config = create_menu_config(&root, &[task]);
    let input = format!("5\n6\nn\n{}\n", source.display());
    let output = run_menu(&config, &input);
    assert!(
        output.status.success(),
        "PT-7-6 menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let output_path = root.join("app/output/区域提取/date-menu-output.xlsx");
    let sheets = read_value_columns_with_iso_dates(&output_path, 16_384).unwrap();
    assert_eq!(sheets.len(), 1);
    assert_eq!(sheets[0].name, "月度资金报表");
    assert_eq!(
        sheets[0].values.get("A1"),
        Some(&CellValue::DateTimeIso(
            "2024-12-31T09:30:00.000".to_owned()
        ))
    );
    assert_eq!(fs::read(&source).unwrap(), source_bytes);
    fs::remove_dir_all(root).unwrap();
}

#[test]
fn interactive_menu_extracts_macro_enabled_xlsm_area_to_new_xlsx() {
    let root = temporary_directory("sheet-extract-xlsm-menu");
    let source =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/vba_config/VbaV2.4.2.xlsm");
    let source_bytes = fs::read(&source).unwrap();
    let source_sheets = read_value_columns(&source, 16_384).unwrap();
    let (source_sheet_name, source_cell, expected_text) = source_sheets
        .iter()
        .find_map(|sheet| {
            sheet
                .values
                .iter()
                .find_map(|(address, value)| match value {
                    CellValue::Text(text) if !text.trim().is_empty() => {
                        Some((sheet.name.clone(), address.clone(), text.clone()))
                    }
                    _ => None,
                })
        })
        .expect("macro-enabled test workbook contains a non-empty text cell");

    let mut task = load_tasks(&cases())[0].clone();
    task.workbook_keywords = "VbaV2.4.2".to_owned();
    task.sheet_keywords = source_sheet_name;
    task.exclude_keywords.clear();
    task.area_spec = source_cell;
    task.output_folder = "xlsm-area".to_owned();
    task.output_name = "xlsm-area.xlsx".to_owned();
    let config = create_menu_config(&root, &[task]);
    let output = run_menu(&config, &format!("5\n6\nn\n{}\n", source.display()));
    assert!(
        output.status.success(),
        "XLSM area extraction menu failed: stdout={} stderr={}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );

    let output_path = root.join("app/output/xlsm-area/xlsm-area.xlsx");
    assert!(output_path.is_file());
    assert_eq!(
        output_path
            .extension()
            .and_then(|extension| extension.to_str()),
        Some("xlsx")
    );
    let output_sheets = read_value_columns(&output_path, 16_384).unwrap();
    assert_eq!(output_sheets.len(), 1);
    assert_eq!(output_sheets[0].values.len(), 1);
    assert!(matches!(
        output_sheets[0].values.get("A1"),
        Some(CellValue::Text(value)) if value == &expected_text
    ));
    assert_eq!(fs::read(&source).unwrap(), source_bytes);
    fs::remove_dir_all(root).unwrap();
}

#[test]
fn interactive_menu_preserves_excel_error_cells_during_area_extraction() {
    let root = temporary_directory("sheet-extract-error-cells-menu");
    let source = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/summary_used_range/error_value.xlsx");
    let source_bytes = fs::read(&source).unwrap();
    let source_sheets = read_value_columns(&source, 16_384).unwrap();
    let error_sheet = source_sheets
        .iter()
        .find(|sheet| sheet.name == "Errors")
        .unwrap();
    let addresses = ["A1", "B1", "C1", "D1", "E1", "F1", "G1"];
    for address in addresses {
        assert!(matches!(
            error_sheet.values.get(address),
            Some(CellValue::Error(_))
        ));
    }

    let mut task = load_tasks(&cases())[0].clone();
    task.workbook_keywords.clear();
    task.sheet_keywords = "Errors".to_owned();
    task.exclude_keywords.clear();
    task.area_spec = "A1:G1".to_owned();
    task.output_folder = "错误值提取".to_owned();
    task.output_name = "error-values-output.xlsx".to_owned();
    let config = create_menu_config(&root, &[task]);
    let output = run_menu(&config, &format!("5\n6\nn\n{}\n", source.display()));
    assert!(
        output.status.success(),
        "error-cell extraction menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let output_path = root.join("app/output/错误值提取/error-values-output.xlsx");
    assert!(output_path.is_file());
    let output_sheets = read_value_columns(&output_path, 16_384).unwrap();
    assert_eq!(output_sheets.len(), 1);
    assert_eq!(output_sheets[0].name, error_sheet.name);
    for address in addresses {
        assert_eq!(
            output_sheets[0].values.get(address),
            error_sheet.values.get(address),
            "area extraction changed {address}"
        );
        assert!(matches!(
            output_sheets[0].values.get(address),
            Some(CellValue::Error(_))
        ));
    }
    assert_eq!(fs::read(&source).unwrap(), source_bytes);
    fs::remove_dir_all(root).unwrap();
}

#[test]
fn interactive_menu_extracts_elapsed_duration_values() -> anyhow::Result<()> {
    let root = temporary_directory("sheet-extract-duration-menu");
    let source = root.join("duration-source.xlsx");
    let source_bytes;
    {
        let mut workbook = rust_xlsxwriter::Workbook::new();
        let worksheet = workbook.add_worksheet();
        worksheet.set_name("时长数据")?;
        let elapsed_format = rust_xlsxwriter::Format::new().set_num_format("[h]:mm:ss.000");
        worksheet.write_number_with_format(0, 0, 1.5, &elapsed_format)?;
        workbook.save(&source)?;
        source_bytes = fs::read(&source)?;
    }
    let source_sheets = read_value_columns_with_iso_dates(&source, 16_384)?;
    assert_eq!(
        source_sheets[0].values.get("A1"),
        Some(&CellValue::DateTimeSerial(1.5))
    );

    let mut task = load_tasks(&cases())[0].clone();
    task.workbook_keywords.clear();
    task.sheet_keywords = "时长数据".to_owned();
    task.exclude_keywords.clear();
    task.area_spec = "A1".to_owned();
    task.output_folder = "区域提取".to_owned();
    task.output_name = "duration-menu-output.xlsx".to_owned();
    let config = create_menu_config(&root, &[task]);
    let input = format!("5\n6\nn\n{}\n", source.display());
    let output = run_menu(&config, &input);
    assert!(
        output.status.success(),
        "PT-7-6 elapsed-duration menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let output_path = root.join("app/output/区域提取/duration-menu-output.xlsx");
    let sheets = read_value_columns_with_iso_dates(&output_path, 16_384)?;
    assert_eq!(sheets.len(), 1);
    assert_eq!(sheets[0].name, "时长数据");
    assert_eq!(
        sheets[0].values.get("A1"),
        Some(&CellValue::DateTimeSerial(1.5))
    );
    let number_formats = read_xlsx_cell_number_formats(
        &output_path,
        &[CellNumberFormatRequest {
            sheet_name: "时长数据".to_owned(),
            cell_reference: "A1".to_owned(),
        }],
    )?;
    assert_eq!(number_formats[0].format_code.as_deref(), Some("[hh]:mm:ss"));
    assert_eq!(fs::read(&source)?, source_bytes);
    fs::remove_dir_all(root)?;
    Ok(())
}
