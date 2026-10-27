use pytools_core::file_sort_prefix::run_content_prefix_unrename_xlsx_candidate;
use serde_json::{json, Value};
use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::time::{SystemTime, UNIX_EPOCH};

fn cases() -> Value {
    serde_json::from_str(include_str!("fixtures/content_unrename_cases.json"))
        .expect("valid content unrename cases")
}

fn baseline() -> Value {
    serde_json::from_str(include_str!(
        "../probe/reports/content_unrename_47_python_baseline.json"
    ))
    .expect("valid Python operation 4-7 baseline")
}

fn temporary_directory(label: &str) -> PathBuf {
    let unique = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap()
        .as_nanos();
    let path = std::env::temp_dir().join(format!(
        "pytools-content-unrename-{label}-{}-{unique}",
        std::process::id()
    ));
    fs::create_dir_all(&path).unwrap();
    path
}

fn write_config(path: &Path, case: &Value, include_sort_sheet: bool) {
    let mut workbook = excel_writer::Workbook::new();
    {
        let sheet = workbook.add_worksheet();
        sheet.set_name("重命名配置").unwrap();
        sheet.write_string(0, 14, "内容重命名启用").unwrap();
        sheet.write_string(0, 15, "内容重命名取值规则").unwrap();
        sheet.write_string(0, 16, "备注").unwrap();
        sheet.write_string(1, 14, "Y").unwrap();
        sheet
            .write_string(1, 15, case["content_spec"].as_str().unwrap())
            .unwrap();
        sheet.write_string(1, 16, "4-7 对拍").unwrap();
    }
    if include_sort_sheet {
        let sheet = workbook.add_worksheet();
        sheet.set_name("文件排序配置").unwrap();
        let headers = ["是否启用", "排序前缀", "匹配关键词", "备注"];
        for (column, header) in headers.iter().enumerate() {
            sheet.write_string(0, column as u16, *header).unwrap();
        }
        for (index, row) in case["sort_rows"].as_array().unwrap().iter().enumerate() {
            for (column, key) in ["enabled", "sort_prefix", "keywords", "remark"]
                .iter()
                .enumerate()
            {
                sheet
                    .write_string(
                        index as u32 + 1,
                        column as u16,
                        row[*key].as_str().unwrap_or_default(),
                    )
                    .unwrap();
            }
        }
    }
    workbook.save(path).unwrap();
}

fn build_inputs(input_dir: &Path, case: &Value) -> Vec<PathBuf> {
    for name in string_set(case, "existing_names") {
        fs::write(input_dir.join(name), b"occupied target").unwrap();
    }
    let missing = string_set(case, "missing_names");
    let directories = string_set(case, "directory_names");
    let corrupt = string_set(case, "corrupt_names");
    let workbooks = case["workbooks"].as_object().unwrap();

    case["input_names"]
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
                fs::write(&path, b"not a workbook package").unwrap();
            } else if let Some(sheet_data) = workbooks.get(name) {
                let mut workbook = excel_writer::Workbook::new();
                for (sheet_index, (sheet_name, cells)) in
                    sheet_data.as_object().unwrap().iter().enumerate()
                {
                    let sheet = workbook.add_worksheet();
                    sheet.set_name(sheet_name).unwrap();
                    for (address, value) in cells.as_object().unwrap() {
                        sheet
                            .write_string(
                                address_to_row(address),
                                address_to_column(address),
                                value.as_str().unwrap_or_default(),
                            )
                            .unwrap();
                    }
                    let _ = sheet_index;
                }
                workbook.save(&path).unwrap();
            } else {
                fs::write(&path, b"not a workbook package").unwrap();
            }
            path
        })
        .collect()
}

fn address_to_row(address: &str) -> u32 {
    let digits = address
        .find(|character: char| character.is_ascii_digit())
        .unwrap();
    address[digits..].parse::<u32>().unwrap() - 1
}

fn address_to_column(address: &str) -> u16 {
    let digits = address
        .find(|character: char| character.is_ascii_digit())
        .unwrap();
    address[..digits].chars().fold(0_u16, |column, character| {
        column * 26 + (character.to_ascii_uppercase() as u16 - b'A' as u16 + 1)
    }) - 1
}

fn string_set(case: &Value, key: &str) -> BTreeSet<String> {
    case.get(key)
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
        .map(|value| value.as_str().unwrap().to_owned())
        .collect()
}

fn run_case(
    case: &Value,
    expected_final_entries: &Value,
    root: &Path,
    include_sort_sheet: bool,
) -> Value {
    let input_dir = root.join("input");
    fs::create_dir_all(&input_dir).unwrap();
    let config = root.join("config.xlsx");
    write_config(&config, case, include_sort_sheet);
    let files = build_inputs(&input_dir, case);
    let original_bytes = files
        .iter()
        .filter_map(|path| fs::read(path).ok().map(|bytes| (path.clone(), bytes)))
        .collect::<BTreeMap<_, _>>();

    let result = run_content_prefix_unrename_xlsx_candidate(&config, &files).unwrap();
    let stats = json!({
        "ok": result.ok,
        "skip": result.skip,
        "spec": result.spec,
        "config_row": result.config_row,
        "config_remark": result.config_remark,
    });
    let final_entries = fs::read_dir(&input_dir)
        .unwrap()
        .filter_map(Result::ok)
        .map(|entry| entry.file_name().to_string_lossy().into_owned())
        .collect::<BTreeSet<_>>();
    let expected_entries = expected_final_entries
        .as_array()
        .unwrap()
        .iter()
        .map(|value| value.as_str().unwrap().to_owned())
        .collect::<BTreeSet<_>>();
    assert_eq!(final_entries, expected_entries);

    for row in &result.rows {
        if row.new_name.is_empty() {
            continue;
        }
        let original = input_dir.join(&row.original_name);
        let target = input_dir.join(&row.new_name);
        assert_eq!(fs::read(&target).unwrap(), original_bytes[&original]);
    }
    json!({
        "stats": stats,
        "final_entries": final_entries,
        "rows": result.rows,
    })
}

#[test]
fn operation_47_matches_python_for_xlsx_xlsm_prefixes_collisions_and_failures() {
    let cases = cases();
    let expected = baseline();
    let case = &cases["operation47"];
    let root = temporary_directory("main");
    let actual = run_case(case, &expected["operation47"]["final_entries"], &root, true);

    assert_eq!(actual["stats"], expected["operation47"]["stats"]);
    assert_eq!(
        actual["final_entries"],
        expected["operation47"]["final_entries"]
    );
    assert_eq!(actual["stats"]["ok"], 8);
    assert_eq!(actual["stats"]["skip"], 5);
    fs::remove_dir_all(root).unwrap();
}

#[test]
fn operation_47_uses_no_sort_prefixes_when_legacy_config_sheet_is_absent() {
    let cases = cases();
    let expected = baseline();
    let case = &cases["legacy_without_sort_sheet"];
    let root = temporary_directory("legacy");
    let actual = run_case(
        case,
        &expected["legacy_without_sort_sheet"]["final_entries"],
        &root,
        false,
    );

    assert_eq!(
        actual["stats"],
        expected["legacy_without_sort_sheet"]["stats"]
    );
    assert_eq!(actual["stats"]["ok"], 1);
    assert_eq!(actual["stats"]["skip"], 1);
    fs::remove_dir_all(root).unwrap();
}

fn run_cli_case(case: &Value, expected: &Value, root: &Path, include_sort_sheet: bool) -> Value {
    let input_dir = root.join("input");
    fs::create_dir_all(&input_dir).unwrap();
    let config = root.join("config.xlsx");
    write_config(&config, case, include_sort_sheet);
    let files = build_inputs(&input_dir, case);
    let original_bytes = files
        .iter()
        .filter_map(|path| fs::read(path).ok().map(|bytes| (path.clone(), bytes)))
        .collect::<BTreeMap<_, _>>();

    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("unrename-by-content")
        .arg("--config")
        .arg(&config)
        .args(&files)
        .output()
        .unwrap();
    assert!(
        output.status.success(),
        "CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let result: Value = serde_json::from_slice(&output.stdout).unwrap();
    assert_eq!(result["ok"], expected["stats"]["ok"]);
    assert_eq!(result["skip"], expected["stats"]["skip"]);
    assert_eq!(result["spec"], expected["stats"]["spec"]);
    assert_eq!(result["config_row"], expected["stats"]["config_row"]);
    assert_eq!(result["config_remark"], expected["stats"]["config_remark"]);

    let final_entries = fs::read_dir(&input_dir)
        .unwrap()
        .filter_map(Result::ok)
        .map(|entry| entry.file_name().to_string_lossy().into_owned())
        .collect::<BTreeSet<_>>();
    let expected_entries = expected["final_entries"]
        .as_array()
        .unwrap()
        .iter()
        .map(|value| value.as_str().unwrap().to_owned())
        .collect::<BTreeSet<_>>();
    assert_eq!(final_entries, expected_entries);

    for row in result["rows"].as_array().unwrap() {
        let Some(new_name) = row["new_name"].as_str().filter(|name| !name.is_empty()) else {
            continue;
        };
        let original_name = row["original_name"].as_str().unwrap();
        let original = input_dir.join(original_name);
        let target = input_dir.join(new_name);
        assert_eq!(fs::read(&target).unwrap(), original_bytes[&original]);
    }
    result
}

#[test]
fn operation_47_cli_executes_rust_workflow_against_python_baseline() {
    let cases = cases();
    let expected = baseline();
    let case = &cases["operation47"];
    let root = temporary_directory("cli-main");
    let result = run_cli_case(case, &expected["operation47"], &root, true);
    assert_eq!(result["ok"], 8);
    assert_eq!(result["skip"], 5);
    fs::remove_dir_all(root).unwrap();
}

#[test]
fn operation_47_cli_keeps_legacy_behavior_without_sort_config_sheet() {
    let cases = cases();
    let expected = baseline();
    let case = &cases["legacy_without_sort_sheet"];
    let root = temporary_directory("cli-legacy");
    let result = run_cli_case(case, &expected["legacy_without_sort_sheet"], &root, false);
    assert_eq!(result["ok"], 1);
    assert_eq!(result["skip"], 1);
    fs::remove_dir_all(root).unwrap();
}
