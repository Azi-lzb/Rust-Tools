use pytools_core::file_sort_prefix::run_file_sort_prefix;
use serde_json::{json, Value};
use std::collections::HashSet;
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
        "../probe/reports/file_sort_prefix_operation_python_baseline.json"
    ))
    .expect("valid Python operation baseline")
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

fn strings(case: &Value, key: &str) -> HashSet<String> {
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
fn operation_49_matches_python_rename_and_report_baseline() {
    let cases = cases();
    let baseline = baseline();
    let operation = &cases["operation49"];
    let root = temporary_directory("sort-prefix-operation");
    let input_dir = root.join("input");
    let output_dir = root.join("output");
    fs::create_dir_all(&input_dir).unwrap();

    for name in strings(operation, "existing_names") {
        fs::write(input_dir.join(name), b"occupied target").unwrap();
    }
    let missing = strings(operation, "missing_names");
    let directories = strings(operation, "directory_names");
    let files = operation["input_names"]
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
            } else {
                fs::write(&path, b"fixture payload").unwrap();
            }
            path
        })
        .collect::<Vec<_>>();

    let config =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/file_sort_prefix_config.xlsx");
    let result = run_file_sort_prefix(&config, &files, &output_dir).unwrap();
    let expected_stats = &baseline["stats"];
    assert_eq!(
        result.rules,
        expected_stats["rules"].as_u64().unwrap() as usize
    );
    assert_eq!(
        result.files,
        expected_stats["files"].as_u64().unwrap() as usize
    );
    assert_eq!(
        result.renamed,
        expected_stats["renamed"].as_u64().unwrap() as usize
    );
    assert_eq!(
        result.unmatched,
        expected_stats["unmatched"].as_u64().unwrap() as usize
    );
    assert_eq!(
        result.skipped,
        expected_stats["skipped"].as_u64().unwrap() as usize
    );
    assert_eq!(
        result.failed,
        expected_stats["failed"].as_u64().unwrap() as usize
    );

    let normalized_rows = result
        .rows
        .iter()
        .map(|row| {
            json!({
                "source_file": Path::new(&row.source_path)
                    .file_name()
                    .unwrap()
                    .to_string_lossy(),
                "original_name": row.original_name,
                "new_name": row.new_name,
                "config_row": row.config_row.map_or(json!(""), |value| json!(value)),
                "sort_prefix": row.sort_prefix,
                "matched_keywords": row.matched_keywords,
                "status": row.status,
                "remark": row.remark,
            })
        })
        .collect::<Vec<_>>();
    assert_eq!(json!(normalized_rows), baseline["rows"]);

    let final_files = fs::read_dir(&input_dir)
        .unwrap()
        .filter_map(Result::ok)
        .filter_map(|entry| {
            entry
                .file_type()
                .ok()
                .filter(|ty| ty.is_file())
                .map(|_| entry)
        })
        .map(|entry| entry.file_name().to_string_lossy().into_owned())
        .collect::<std::collections::BTreeSet<_>>()
        .into_iter()
        .collect::<Vec<_>>();
    assert_eq!(json!(final_files), baseline["final_files"]);

    let dump = excel_reader::read_workbook(&result.output).unwrap();
    let sheet = dump
        .sheets
        .iter()
        .find(|sheet| sheet.name == "文件排序前缀结果")
        .unwrap();
    let expected_headers = [
        "源文件路径",
        "原文件名",
        "新文件名",
        "命中配置行",
        "排序前缀",
        "匹配关键词",
        "状态",
        "备注",
    ];
    for (index, header) in expected_headers.iter().enumerate() {
        assert_eq!(report_cell(sheet, &excel_cell(1, index + 1)), *header);
    }
    for (index, expected) in baseline["rows"].as_array().unwrap().iter().enumerate() {
        let row = index + 2;
        let actual_source = report_cell(sheet, &excel_cell(row, 1));
        assert_eq!(
            Path::new(&actual_source)
                .file_name()
                .unwrap()
                .to_string_lossy()
                .as_ref(),
            expected["source_file"].as_str().unwrap()
        );
        assert_eq!(
            report_cell(sheet, &excel_cell(row, 2)),
            expected["original_name"]
        );
        assert_eq!(
            report_cell(sheet, &excel_cell(row, 3)),
            expected["new_name"]
        );
        assert_eq!(
            report_cell(sheet, &excel_cell(row, 4)),
            expected["config_row"]
                .as_str()
                .map_or_else(|| expected["config_row"].to_string(), str::to_owned)
        );
        assert_eq!(
            report_cell(sheet, &excel_cell(row, 5)),
            expected["sort_prefix"]
        );
        assert_eq!(
            report_cell(sheet, &excel_cell(row, 6)),
            expected["matched_keywords"]
        );
        assert_eq!(report_cell(sheet, &excel_cell(row, 7)), expected["status"]);
        assert_eq!(report_cell(sheet, &excel_cell(row, 8)), expected["remark"]);
    }

    assert_eq!(
        fs::read(input_dir.join("001_202506_国开行_重复_1.xlsx")).unwrap(),
        b"fixture payload"
    );
    assert_eq!(
        fs::read(input_dir.join("001_202506_国开行_重复.xlsx")).unwrap(),
        b"occupied target"
    );
    fs::remove_dir_all(root).unwrap();
}

#[test]
fn operation_49_cli_executes_rust_workflow_against_python_baseline() {
    let cases = cases();
    let baseline = baseline();
    let operation = &cases["operation49"];
    let root = temporary_directory("sort-prefix-cli");
    let input_dir = root.join("input");
    let output_dir = root.join("output");
    let log_dir = root.join("logs");
    fs::create_dir_all(&input_dir).unwrap();

    for name in strings(operation, "existing_names") {
        fs::write(input_dir.join(name), b"occupied target").unwrap();
    }
    let missing = strings(operation, "missing_names");
    let directories = strings(operation, "directory_names");
    let files = operation["input_names"]
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
            } else {
                fs::write(&path, b"fixture payload").unwrap();
            }
            path
        })
        .collect::<Vec<_>>();
    let config =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/file_sort_prefix_config.xlsx");

    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("file-sort-prefix")
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
    for field in [
        "rules",
        "files",
        "renamed",
        "unmatched",
        "skipped",
        "failed",
    ] {
        assert_eq!(result[field], baseline["stats"][field], "stat {field}");
    }

    let report_path = PathBuf::from(result["output"].as_str().unwrap());
    assert!(report_path.is_file());
    assert!(report_path.starts_with(&output_dir));
    let report = excel_reader::read_workbook(&report_path).unwrap();
    assert!(report
        .sheets
        .iter()
        .any(|sheet| sheet.name == "文件排序前缀结果"));
    let log_path = PathBuf::from(result["log_path"].as_str().unwrap());
    assert!(log_path.is_file());
    assert!(log_path.starts_with(&log_dir));
    let log_contents = fs::read_to_string(log_path).unwrap();
    let log_lines = log_contents.lines().collect::<Vec<_>>();
    assert_eq!(log_lines.len(), 5);
    assert!(log_contents.contains("[INFO] ok old="));
    assert!(log_contents.contains("row=3 prefix=002"));

    let final_files = fs::read_dir(&input_dir)
        .unwrap()
        .filter_map(Result::ok)
        .filter_map(|entry| {
            entry
                .file_type()
                .ok()
                .filter(|file_type| file_type.is_file())
                .map(|_| entry)
        })
        .map(|entry| entry.file_name().to_string_lossy().into_owned())
        .collect::<std::collections::BTreeSet<_>>()
        .into_iter()
        .collect::<Vec<_>>();
    assert_eq!(json!(final_files), baseline["final_files"]);
    for name in &final_files {
        let expected_bytes = if name == "001_202506_国开行_重复.xlsx" {
            b"occupied target".as_slice()
        } else {
            b"fixture payload".as_slice()
        };
        assert_eq!(fs::read(input_dir.join(name)).unwrap(), expected_bytes);
    }
    fs::remove_dir_all(root).unwrap();
}

#[test]
fn output_directory_failure_leaves_source_files_untouched() {
    let root = temporary_directory("sort-prefix-failure");
    let source = root.join("202506_建设银行_存款表.xlsx");
    fs::write(&source, b"must remain").unwrap();
    let output_blocker = root.join("output-is-a-file");
    fs::write(&output_blocker, b"not a directory").unwrap();
    let config =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/file_sort_prefix_config.xlsx");

    let error =
        run_file_sort_prefix(&config, std::slice::from_ref(&source), &output_blocker).unwrap_err();
    assert!(error.to_string().contains("create output directory"));
    assert_eq!(fs::read(&source).unwrap(), b"must remain");
    fs::remove_dir_all(root).unwrap();
}
