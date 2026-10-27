use pytools_core::file_sort_prefix::run_content_prefix_rename_xlsx_candidate;
use serde_json::{json, Value};
use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::{
    sync::atomic::{AtomicU64, Ordering},
    time::{SystemTime, UNIX_EPOCH},
};

static NEXT_TEMPORARY_DIRECTORY_ID: AtomicU64 = AtomicU64::new(0);

fn cases() -> Value {
    serde_json::from_str(include_str!("fixtures/content_rename_46_cases.json"))
        .expect("valid content rename cases")
}

fn baseline() -> Value {
    serde_json::from_str(include_str!(
        "../probe/reports/content_rename_46_python_baseline.json"
    ))
    .expect("valid Python operation 4-6 baseline")
}

fn temporary_directory() -> PathBuf {
    let unique = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap()
        .as_nanos();
    let serial = NEXT_TEMPORARY_DIRECTORY_ID.fetch_add(1, Ordering::Relaxed);
    let path = std::env::temp_dir().join(format!(
        "pytools-content-rename-46-{}-{unique}-{serial}",
        std::process::id(),
    ));
    fs::create_dir_all(&path).unwrap();
    path
}

fn write_config(path: &Path, case: &Value) {
    let mut workbook = excel_writer::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("重命名配置").unwrap();
    sheet.write_string(0, 14, "内容重命名启用").unwrap();
    sheet.write_string(0, 15, "内容重命名取值规则").unwrap();
    sheet.write_string(0, 16, "备注").unwrap();
    sheet.write_string(1, 14, "Y").unwrap();
    sheet
        .write_string(1, 15, case["content_spec"].as_str().unwrap())
        .unwrap();
    sheet
        .write_string(1, 16, case["remark"].as_str().unwrap())
        .unwrap();
    workbook.save(path).unwrap();
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
                for (sheet_name, cells) in sheet_data.as_object().unwrap() {
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
                }
                workbook.save(&path).unwrap();
            } else {
                fs::write(&path, b"not a workbook package").unwrap();
            }
            path
        })
        .collect()
}

fn run_case(case: &Value, expected_final_entries: &Value, root: &Path) -> Value {
    let input_dir = root.join("input");
    fs::create_dir_all(&input_dir).unwrap();
    let config = root.join("config.xlsx");
    write_config(&config, case);
    let files = build_inputs(&input_dir, case);
    let original_bytes = files
        .iter()
        .filter_map(|path| fs::read(path).ok().map(|bytes| (path.clone(), bytes)))
        .collect::<BTreeMap<_, _>>();

    let result = run_content_prefix_rename_xlsx_candidate(&config, &files).unwrap();
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
        let original = files
            .iter()
            .find(|path| path.file_name().unwrap().to_string_lossy() == row.original_name);
        let Some(original) = original else {
            continue;
        };
        let Some(bytes) = original_bytes.get(original) else {
            continue;
        };
        let output_name = if row.new_name.is_empty() {
            &row.original_name
        } else {
            &row.new_name
        };
        assert_eq!(fs::read(input_dir.join(output_name)).unwrap(), *bytes);
    }
    assert_eq!(
        fs::read(input_dir.join("A_B_conflict.xlsx")).unwrap(),
        b"occupied target"
    );

    json!({
        "stats": stats,
        "final_entries": final_entries,
        "renamed_file_bytes_preserved": true,
    })
}

#[test]
fn operation_46_matches_python_for_prefixes_conflicts_and_skips() {
    let cases = cases();
    let expected = baseline()["operation46"].clone();
    let case = &cases["operation46"];
    let root = temporary_directory();
    let actual = run_case(case, &expected["final_entries"], &root);

    assert_eq!(actual["stats"], expected["stats"]);
    assert_eq!(actual["final_entries"], expected["final_entries"]);
    assert_eq!(actual["stats"]["ok"], 5);
    assert_eq!(actual["stats"]["skip"], 7);
    fs::remove_dir_all(root).unwrap();
}

#[test]
fn operation_46_cli_executes_rust_workflow_against_python_baseline() {
    let cases = cases();
    let expected = baseline()["operation46"].clone();
    let case = &cases["operation46"];
    let root = temporary_directory();
    let input_dir = root.join("input");
    fs::create_dir_all(&input_dir).unwrap();
    let config = root.join("config.xlsx");
    write_config(&config, case);
    let files = build_inputs(&input_dir, case);
    let original_bytes = files
        .iter()
        .filter_map(|path| fs::read(path).ok().map(|bytes| (path.clone(), bytes)))
        .collect::<BTreeMap<_, _>>();

    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("rename-by-content")
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
        let original_name = row["original_name"].as_str().unwrap();
        let original = files
            .iter()
            .find(|path| path.file_name().unwrap().to_string_lossy() == original_name)
            .unwrap();
        let Some(bytes) = original_bytes.get(original) else {
            continue;
        };
        let output_name = row["new_name"]
            .as_str()
            .filter(|name| !name.is_empty())
            .unwrap_or(original_name);
        assert_eq!(fs::read(input_dir.join(output_name)).unwrap(), *bytes);
    }
    assert_eq!(
        fs::read(input_dir.join("A_B_conflict.xlsx")).unwrap(),
        b"occupied target"
    );
    fs::remove_dir_all(root).unwrap();
}

#[test]
fn operation_46_xls_integer_values_match_legacy_com_float_text() {
    let root = temporary_directory();
    let input_dir = root.join("input");
    fs::create_dir_all(&input_dir).unwrap();
    let mut case = cases()["operation46"].clone();
    case["content_spec"] = json!("Data@A1;Data@B2");
    case["remark"] = json!("synthetic legacy XLS COM numeric parity");
    let config = root.join("config.xlsx");
    write_config(&config, &case);

    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("timeline")
        .join("wide_summary_legacy.xls");
    let source = input_dir.join("legacy.xls");
    fs::copy(&fixture, &source).unwrap();
    let original_bytes = fs::read(&source).unwrap();

    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("rename-by-content")
        .arg("--config")
        .arg(&config)
        .arg(&source)
        .output()
        .unwrap();
    assert!(
        output.status.success(),
        "CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let result: Value = serde_json::from_slice(&output.stdout).unwrap();
    assert_eq!(result["ok"], 1);
    assert_eq!(
        result["rows"][0]["new_name"],
        "Installed smoke_42.0_legacy.xls"
    );
    assert!(!source.exists());
    assert_eq!(
        fs::read(input_dir.join("Installed smoke_42.0_legacy.xls")).unwrap(),
        original_bytes
    );
    fs::remove_dir_all(root).unwrap();
}
