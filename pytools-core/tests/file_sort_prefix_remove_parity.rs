use pytools_core::file_sort_prefix::{
    run_remove_file_sort_prefix, run_remove_file_sort_prefix_with_progress,
};
use serde_json::{json, Value};
use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::time::{SystemTime, UNIX_EPOCH};

fn cases() -> Value {
    serde_json::from_str(include_str!("fixtures/file_sort_prefix_cases.json"))
        .expect("valid operation cases")
}

fn baseline() -> Value {
    serde_json::from_str(include_str!(
        "../probe/reports/file_sort_prefix_remove_python_baseline.json"
    ))
    .expect("valid Python operation 4-10 baseline")
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

fn write_config(path: &Path, cases: &Value, include_rules: bool) {
    let mut workbook = excel_writer::Workbook::new();
    {
        let sheet = workbook.add_worksheet();
        sheet.set_name("重命名配置").unwrap();
        sheet.write_string(0, 14, "内容重命名启用").unwrap();
        sheet.write_string(0, 15, "内容重命名取值规则").unwrap();
        sheet.write_string(0, 16, "备注").unwrap();
        sheet.write_string(1, 14, "Y").unwrap();
        sheet
            .write_string(
                1,
                15,
                cases["operation410"]["content_spec"].as_str().unwrap(),
            )
            .unwrap();
        sheet.write_string(1, 16, "内容前缀对拍").unwrap();
    }
    {
        let sheet = workbook.add_worksheet();
        sheet.set_name("文件排序配置").unwrap();
        let headers = ["是否启用", "排序前缀", "匹配关键词", "备注"];
        for (column, header) in headers.iter().enumerate() {
            sheet.write_string(0, column as u16, *header).unwrap();
        }
        if include_rules {
            for (index, row) in cases["config_rows"].as_array().unwrap().iter().enumerate() {
                for (column, key) in ["enabled", "sort_prefix", "keywords", "remark"]
                    .iter()
                    .enumerate()
                {
                    let value = row[*key].as_str().unwrap_or_default();
                    sheet
                        .write_string(index as u32 + 1, column as u16, value)
                        .unwrap();
                }
            }
        }
    }
    workbook.save(path).unwrap();
}

fn build_inputs(input_dir: &Path, operation: &Value) -> Vec<PathBuf> {
    for name in string_set(operation, "existing_names") {
        fs::write(input_dir.join(name), b"occupied target").unwrap();
    }
    let missing = string_set(operation, "missing_names");
    let directories = string_set(operation, "directory_names");
    let corrupt = string_set(operation, "corrupt_names");
    let workbook_values = operation["workbook_values"].as_object().unwrap();

    operation["input_names"]
        .as_array()
        .unwrap()
        .iter()
        .map(|value| {
            let name = value.as_str().unwrap();
            let path = input_dir.join(name);
            if missing.contains(name) {
                return path;
            }
            if directories.contains(name) {
                fs::create_dir(&path).unwrap();
                fs::write(path.join("marker.txt"), b"directory").unwrap();
            } else if corrupt.contains(name) {
                fs::write(&path, b"not an xlsx package").unwrap();
            } else if let Some(cell_value) = workbook_values.get(name) {
                let mut workbook = excel_writer::Workbook::new();
                let sheet = workbook.add_worksheet();
                sheet.set_name("sheet1").unwrap();
                sheet
                    .write_string(0, 0, cell_value.as_str().unwrap())
                    .unwrap();
                workbook.save(&path).unwrap();
            } else {
                fs::write(&path, b"fixture payload").unwrap();
            }
            path
        })
        .collect()
}

fn string_set(case: &Value, key: &str) -> BTreeSet<String> {
    case[key]
        .as_array()
        .unwrap()
        .iter()
        .map(|value| value.as_str().unwrap().to_owned())
        .collect()
}

fn report_cell(sheet: &excel_reader::SheetDump, address: &str) -> String {
    let Some(value) = sheet.values.get(address) else {
        return String::new();
    };
    match value["type"].as_str().unwrap_or_default() {
        "string" => value["value"].as_str().unwrap_or_default().to_owned(),
        "integer" => value["value"].to_string(),
        "number" => value["value"]
            .as_f64()
            .map(|number| {
                if number.fract() == 0.0 {
                    format!("{number:.0}")
                } else {
                    number.to_string()
                }
            })
            .unwrap_or_default(),
        other => panic!("unexpected report cell type {other} at {address}"),
    }
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

#[test]
fn operation_410_progress_reports_output_directory_failure() {
    let cases = cases();
    let root = temporary_directory("sort-prefix-remove-output-failure");
    let config = root.join("config.xlsx");
    write_config(&config, &cases, true);
    let output_file = root.join("output-is-a-file");
    fs::write(&output_file, b"occupied").unwrap();

    let mut progress = Vec::new();
    let error = run_remove_file_sort_prefix_with_progress(&config, &[], &output_file, |event| {
        progress.push(event.clone())
    })
    .unwrap_err();
    assert!(error.to_string().contains("create output directory"));
    assert!(progress.iter().any(|event| {
        event.phase == "output_directory_prepare"
            && event.status == "failed"
            && event.actual_engine.as_deref() == Some("RustFilesystem")
            && event.error_type.as_deref() == Some("output_directory_create_failed")
    }));
}

#[test]
fn operation_410_matches_python_prefix_removal_and_report_baseline() {
    let cases = cases();
    let expected = baseline();
    let operation = &cases["operation410"];
    let root = temporary_directory("sort-prefix-remove");
    let input_dir = root.join("input");
    let output_dir = root.join("output");
    fs::create_dir_all(&input_dir).unwrap();
    let config = root.join("config.xlsx");
    write_config(&config, &cases, true);
    let files = build_inputs(&input_dir, operation);
    let original_contents = files
        .iter()
        .filter_map(|path| fs::read(path).ok().map(|bytes| (path.clone(), bytes)))
        .collect::<BTreeMap<_, _>>();

    let mut progress = Vec::new();
    let result = run_remove_file_sort_prefix_with_progress(&config, &files, &output_dir, |event| {
        progress.push(event.clone())
    })
    .unwrap();
    assert!(progress.iter().any(|event| {
        event.phase == "file_processing"
            && event.actual_engine.as_deref() == Some("Calamine + RustFilesystem")
    }));
    assert_eq!(
        result.rules,
        expected["stats"]["rules"].as_u64().unwrap() as usize
    );
    assert_eq!(
        result.files,
        expected["stats"]["files"].as_u64().unwrap() as usize
    );
    assert_eq!(
        result.renamed,
        expected["stats"]["renamed"].as_u64().unwrap() as usize
    );
    assert_eq!(
        result.skipped,
        expected["stats"]["skipped"].as_u64().unwrap() as usize
    );
    assert_eq!(
        result.failed,
        expected["stats"]["failed"].as_u64().unwrap() as usize
    );

    let rows = result
        .rows
        .iter()
        .map(|row| {
            json!({
                "source_file": Path::new(&row.source_path).file_name().unwrap().to_string_lossy(),
                "original_name": row.original_name,
                "new_name": row.new_name,
                "status": row.status,
                "remark": row.remark,
            })
        })
        .collect::<Vec<_>>();
    assert_eq!(json!(rows), expected["rows"]);

    let final_files = fs::read_dir(&input_dir)
        .unwrap()
        .filter_map(Result::ok)
        .filter_map(|entry| {
            entry
                .file_type()
                .ok()
                .filter(|kind| kind.is_file())
                .map(|_| entry.file_name().to_string_lossy().into_owned())
        })
        .collect::<BTreeSet<_>>()
        .into_iter()
        .collect::<Vec<_>>();
    assert_eq!(json!(final_files), expected["final_files"]);

    for row in &result.rows {
        if row.new_name.is_empty() {
            continue;
        }
        let original = input_dir.join(&row.original_name);
        let target = input_dir.join(&row.new_name);
        let before = original_contents.get(&original).unwrap();
        assert_eq!(&fs::read(target).unwrap(), before, "{}", row.original_name);
    }

    let dump = excel_reader::read_workbook(&result.output).unwrap();
    let sheet = dump
        .sheets
        .iter()
        .find(|sheet| sheet.name == "取消文件排序前缀结果")
        .unwrap();
    let headers = ["源文件路径", "原文件名", "新文件名", "状态", "备注"];
    for (index, header) in headers.iter().enumerate() {
        assert_eq!(report_cell(sheet, &excel_cell(1, index + 1)), *header);
    }
    for (index, row) in expected["rows"].as_array().unwrap().iter().enumerate() {
        let excel_row = index + 2;
        let source_path = report_cell(sheet, &excel_cell(excel_row, 1));
        assert_eq!(
            Path::new(&source_path)
                .file_name()
                .unwrap()
                .to_string_lossy(),
            row["source_file"].as_str().unwrap()
        );
        assert_eq!(
            report_cell(sheet, &excel_cell(excel_row, 2)),
            row["original_name"]
        );
        assert_eq!(
            report_cell(sheet, &excel_cell(excel_row, 3)),
            row["new_name"]
        );
        assert_eq!(report_cell(sheet, &excel_cell(excel_row, 4)), row["status"]);
        assert_eq!(report_cell(sheet, &excel_cell(excel_row, 5)), row["remark"]);
    }
    fs::remove_dir_all(root).unwrap();
}

#[test]
fn operation_410_cli_executes_rust_workflow_against_python_baseline() {
    let cases = cases();
    let expected = baseline();
    let operation = &cases["operation410"];
    let root = temporary_directory("sort-prefix-remove-cli");
    let input_dir = root.join("input");
    let output_dir = root.join("output");
    let log_dir = root.join("logs");
    fs::create_dir_all(&input_dir).unwrap();
    let config = root.join("config.xlsx");
    write_config(&config, &cases, true);
    let files = build_inputs(&input_dir, operation);
    let original_contents = files
        .iter()
        .filter_map(|path| fs::read(path).ok().map(|bytes| (path.clone(), bytes)))
        .collect::<BTreeMap<_, _>>();

    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("remove-file-sort-prefix")
        .arg("--config")
        .arg(&config)
        .arg("--output-dir")
        .arg(&output_dir)
        .arg("--log-dir")
        .arg(&log_dir)
        .args(&files)
        .output()
        .unwrap();
    assert!(
        output.status.success(),
        "CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let result: Value = serde_json::from_slice(&output.stdout).unwrap();
    for field in ["rules", "files", "renamed", "skipped", "failed"] {
        assert_eq!(result[field], expected["stats"][field], "stat {field}");
    }

    let report_path = PathBuf::from(result["output"].as_str().unwrap());
    assert!(report_path.is_file());
    assert!(report_path.starts_with(&output_dir));
    let dump = excel_reader::read_workbook(&report_path).unwrap();
    let report_sheet = dump
        .sheets
        .iter()
        .find(|sheet| sheet.name == "取消文件排序前缀结果")
        .unwrap();
    let headers = ["源文件路径", "原文件名", "新文件名", "状态", "备注"];
    for (index, header) in headers.iter().enumerate() {
        assert_eq!(
            report_cell(report_sheet, &excel_cell(1, index + 1)),
            *header
        );
    }
    for (index, expected_row) in expected["rows"].as_array().unwrap().iter().enumerate() {
        let excel_row = index + 2;
        let source_path = report_cell(report_sheet, &excel_cell(excel_row, 1));
        assert_eq!(
            Path::new(&source_path)
                .file_name()
                .unwrap()
                .to_string_lossy(),
            expected_row["source_file"].as_str().unwrap()
        );
        assert_eq!(
            report_cell(report_sheet, &excel_cell(excel_row, 2)),
            expected_row["original_name"]
        );
        assert_eq!(
            report_cell(report_sheet, &excel_cell(excel_row, 3)),
            expected_row["new_name"]
        );
        assert_eq!(
            report_cell(report_sheet, &excel_cell(excel_row, 4)),
            expected_row["status"]
        );
        assert_eq!(
            report_cell(report_sheet, &excel_cell(excel_row, 5)),
            expected_row["remark"]
        );
    }

    let log_path = PathBuf::from(result["log_path"].as_str().unwrap());
    assert!(log_path.is_file());
    assert!(log_path.starts_with(&log_dir));
    let log_contents = fs::read_to_string(log_path).unwrap();
    assert_eq!(log_contents.lines().count(), 5);
    assert!(log_contents.contains("[INFO] ok old="));

    let final_files = fs::read_dir(&input_dir)
        .unwrap()
        .filter_map(Result::ok)
        .filter_map(|entry| {
            entry
                .file_type()
                .ok()
                .filter(|kind| kind.is_file())
                .map(|_| entry.file_name().to_string_lossy().into_owned())
        })
        .collect::<BTreeSet<_>>()
        .into_iter()
        .collect::<Vec<_>>();
    assert_eq!(json!(final_files), expected["final_files"]);
    for row in result["rows"].as_array().unwrap() {
        let original_name = row["original_name"].as_str().unwrap();
        let new_name = row["new_name"].as_str().unwrap();
        if new_name.is_empty() {
            continue;
        }
        let original = input_dir.join(original_name);
        let Some(bytes) = original_contents.get(&original) else {
            continue;
        };
        assert_eq!(fs::read(input_dir.join(new_name)).unwrap(), *bytes);
    }
    fs::remove_dir_all(root).unwrap();
}

#[test]
fn operation_410_cli_accepts_empty_input_and_writes_blank_report() {
    let cases = cases();
    let root = temporary_directory("sort-prefix-remove-cli-empty");
    let config = root.join("config.xlsx");
    let output_dir = root.join("output");
    let log_dir = root.join("logs");
    write_config(&config, &cases, true);

    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("remove-file-sort-prefix")
        .arg("--config")
        .arg(&config)
        .arg("--output-dir")
        .arg(&output_dir)
        .arg("--log-dir")
        .arg(&log_dir)
        .output()
        .unwrap();
    assert!(
        output.status.success(),
        "CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let result: Value = serde_json::from_slice(&output.stdout).unwrap();
    assert_eq!(result["rules"], 4);
    assert_eq!(result["files"], 0);
    assert_eq!(result["renamed"], 0);
    assert_eq!(result["skipped"], 0);
    assert_eq!(result["failed"], 0);

    let report_path = PathBuf::from(result["output"].as_str().unwrap());
    let dump = excel_reader::read_workbook(report_path).unwrap();
    let report_sheet = dump
        .sheets
        .iter()
        .find(|sheet| sheet.name == "取消文件排序前缀结果")
        .unwrap();
    assert!(report_sheet.values.is_empty());
    let log_path = PathBuf::from(result["log_path"].as_str().unwrap());
    assert_eq!(fs::read_to_string(log_path).unwrap(), "");
    fs::remove_dir_all(root).unwrap();
}

#[test]
fn operation_410_empty_input_matches_python_blank_report_sheet() {
    let cases = cases();
    let expected = baseline();
    let root = temporary_directory("sort-prefix-remove-empty");
    let config = root.join("config.xlsx");
    let output_dir = root.join("output");
    write_config(&config, &cases, true);

    let result = run_remove_file_sort_prefix(&config, &[], &output_dir).unwrap();
    assert_eq!(result.rules, expected["empty_input"]["stats"]["rules"]);
    assert_eq!(result.files, 0);
    assert_eq!(result.renamed, 0);
    assert_eq!(result.skipped, 0);
    assert_eq!(result.failed, 0);

    let dump = excel_reader::read_workbook(&result.output).unwrap();
    let sheet = dump
        .sheets
        .iter()
        .find(|sheet| sheet.name == "取消文件排序前缀结果")
        .unwrap();
    assert!(sheet.values.is_empty());
    assert_eq!(expected["empty_input"]["report"]["cells"], json!([]));
    fs::remove_dir_all(root).unwrap();
}

#[test]
fn operation_410_requires_at_least_one_valid_sort_rule() {
    let cases = cases();
    let root = temporary_directory("sort-prefix-remove-no-rules");
    let config = root.join("config.xlsx");
    write_config(&config, &cases, false);
    let error = run_remove_file_sort_prefix(&config, &[], root.join("output")).unwrap_err();
    assert!(format!("{error:#}").contains("文件排序配置 没有启用且有效的排序规则。"));
    fs::remove_dir_all(root).unwrap();
}
