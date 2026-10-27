use pytools_core::local_filename_map::run_local_filename_map_candidate;
use serde_json::Value;
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
    serde_json::from_str(include_str!("fixtures/local_filename_map_48_cases.json"))
        .expect("valid operation 4-8 cases")
}

fn baseline() -> Value {
    serde_json::from_str(include_str!(
        "../probe/reports/local_filename_map_48_python_baseline.json"
    ))
    .expect("valid operation 4-8 Python baseline")
}

fn temporary_directory() -> PathBuf {
    let unique = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap()
        .as_nanos();
    let serial = NEXT_TEMPORARY_DIRECTORY_ID.fetch_add(1, Ordering::Relaxed);
    let path = std::env::temp_dir().join(format!(
        "pytools-local-map-48-{}-{unique}-{serial}",
        std::process::id(),
    ));
    fs::create_dir_all(&path).unwrap();
    path
}

fn mappings(case: &Value) -> Vec<(String, String)> {
    case["mapping"]
        .as_array()
        .unwrap()
        .iter()
        .map(|rule| {
            let source = rule["source"].as_str().unwrap().to_owned();
            let target = rule
                .get("target")
                .and_then(Value::as_str)
                .map(str::to_owned)
                .unwrap_or_else(|| "x".repeat(rule["target_repeat"].as_u64().unwrap() as usize));
            (source, target)
        })
        .collect()
}

fn write_config(path: &Path, case: &Value) {
    let mut workbook = excel_writer::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("重命名配置").unwrap();
    sheet.write_string(0, 12, "原文件名片段").unwrap();
    sheet.write_string(0, 13, "新文件名片段").unwrap();
    for (index, (source, target)) in mappings(case).iter().enumerate() {
        sheet.write_string(index as u32 + 1, 12, source).unwrap();
        sheet.write_string(index as u32 + 1, 13, target).unwrap();
    }
    sheet.write_string(0, 14, "备注").unwrap();
    sheet
        .write_string(1, 14, case["remark"].as_str().unwrap())
        .unwrap();
    workbook.save(path).unwrap();
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
            } else {
                fs::write(&path, format!("content for {name}")).unwrap();
            }
            path
        })
        .collect()
}

#[test]
fn operation_48_matches_python_order_collision_and_skip_behavior() {
    let cases = cases();
    let case = &cases["operation48"];
    let expected = baseline();
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

    let result = run_local_filename_map_candidate(&config, &files).unwrap();
    let final_entries = fs::read_dir(&input_dir)
        .unwrap()
        .filter_map(Result::ok)
        .map(|entry| entry.file_name().to_string_lossy().into_owned())
        .collect::<BTreeSet<_>>();
    let expected_entries = expected["operation48"]["final_entries"]
        .as_array()
        .unwrap()
        .iter()
        .map(|value| value.as_str().unwrap().to_owned())
        .collect::<BTreeSet<_>>();
    assert_eq!(final_entries, expected_entries);
    assert_eq!(result.ok, 5);
    assert_eq!(result.skip, 5);
    assert_eq!(result.rules, 6);
    assert_eq!(result.ok, expected["operation48"]["stats"]["ok"]);
    assert_eq!(result.skip, expected["operation48"]["stats"]["skip"]);

    let renamed_names = result
        .rows
        .iter()
        .filter(|row| !row.new_name.is_empty())
        .map(|row| row.new_name.clone())
        .collect::<BTreeSet<_>>();
    let expected_renamed_names = string_set(case, "expected_renamed_names");
    assert_eq!(renamed_names, expected_renamed_names);
    let skipped_names = result
        .rows
        .iter()
        .filter(|row| row.new_name.is_empty())
        .map(|row| row.original_name.clone())
        .collect::<BTreeSet<_>>();
    let expected_skipped_names = string_set(case, "expected_skipped_names");
    assert_eq!(skipped_names, expected_skipped_names);

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
        fs::read(input_dir.join("same_cc.xlsx")).unwrap(),
        b"occupied target"
    );
    assert_eq!(
        fs::read(input_dir.join("same_cc_1.xlsx")).unwrap(),
        b"content for same_aa.xlsx"
    );
    assert_eq!(
        fs::read(input_dir.join("same_cc_2.xlsx")).unwrap(),
        b"content for same_bb.xlsx"
    );
    fs::remove_dir_all(root).unwrap();
}

#[test]
fn operation_48_cli_executes_rust_mapping_against_python_baseline() {
    let cases = cases();
    let case = &cases["operation48"];
    let expected = baseline();
    let root = temporary_directory();
    let input_dir = root.join("cli-input");
    fs::create_dir_all(&input_dir).unwrap();
    let config = root.join("config.xlsx");
    write_config(&config, case);
    let files = build_inputs(&input_dir, case);
    let original_bytes = files
        .iter()
        .filter_map(|path| fs::read(path).ok().map(|bytes| (path.clone(), bytes)))
        .collect::<BTreeMap<_, _>>();

    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("rename-local-map")
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
    let result: serde_json::Value = serde_json::from_slice(&output.stdout).unwrap();
    let expected_operation = &expected["operation48"];
    assert_eq!(result["ok"], expected_operation["stats"]["ok"]);
    assert_eq!(result["skip"], expected_operation["stats"]["skip"]);
    assert_eq!(result["rules"], expected_operation["stats"]["rules"]);
    let final_entries = fs::read_dir(&input_dir)
        .unwrap()
        .filter_map(Result::ok)
        .map(|entry| entry.file_name().to_string_lossy().into_owned())
        .collect::<BTreeSet<_>>();
    let expected_entries = expected_operation["final_entries"]
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
            .find(|path| path.file_name().unwrap().to_string_lossy() == original_name);
        let Some(original) = original else {
            continue;
        };
        let Some(bytes) = original_bytes.get(original) else {
            continue;
        };
        let new_name = row["new_name"].as_str().unwrap();
        let output_name = if new_name.is_empty() {
            original_name
        } else {
            new_name
        };
        assert_eq!(fs::read(input_dir.join(output_name)).unwrap(), *bytes);
    }
    fs::remove_dir_all(root).unwrap();
}

#[test]
fn operation_48_rejects_an_empty_mapping() {
    let root = temporary_directory();
    let config = root.join("config.xlsx");
    let mut workbook = excel_writer::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("重命名配置").unwrap();
    sheet.write_string(0, 12, "原文件名片段").unwrap();
    sheet.write_string(0, 13, "新文件名片段").unwrap();
    workbook.save(&config).unwrap();
    let error = run_local_filename_map_candidate(&config, &[]).unwrap_err();
    assert!(error.to_string().contains("缺少 M/N 列"));
    fs::remove_dir_all(root).unwrap();
}

#[test]
fn operation_48_matches_pandas_nan_stringification_for_missing_config_cells() {
    let expected = baseline()["pandas_empty_config_cells"].clone();
    let root = temporary_directory();
    let input_dir = root.join("input");
    fs::create_dir_all(&input_dir).unwrap();
    let config = root.join("config.xlsx");
    let mut workbook = excel_writer::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet
        .set_name("\u{91cd}\u{547d}\u{540d}\u{914d}\u{7f6e}")
        .unwrap();
    sheet
        .write_string(0, 12, "\u{539f}\u{6587}\u{4ef6}\u{540d}\u{7247}\u{6bb5}")
        .unwrap();
    sheet
        .write_string(0, 13, "\u{65b0}\u{6587}\u{4ef6}\u{540d}\u{7247}\u{6bb5}")
        .unwrap();
    sheet.write_string(1, 0, "keep this pandas row").unwrap();
    sheet.write_string(2, 13, "mapped").unwrap();
    sheet.write_string(3, 12, "remove").unwrap();
    sheet
        .write_string(4, 0, "keep the trailing pandas row")
        .unwrap();
    workbook.save(&config).unwrap();

    let inputs = [
        ("nan.txt", "payload from nan"),
        ("remove.txt", "payload from remove"),
        ("keep.txt", "payload from keep"),
    ];
    let files = inputs
        .iter()
        .map(|(name, content)| {
            let path = input_dir.join(name);
            fs::write(&path, content.as_bytes()).unwrap();
            path
        })
        .collect::<Vec<_>>();

    let result = run_local_filename_map_candidate(&config, &files).unwrap();
    assert_eq!(
        result.rules,
        expected["stats"]["rules"].as_u64().unwrap() as usize
    );
    assert_eq!(
        result.ok,
        expected["stats"]["ok"].as_u64().unwrap() as usize
    );
    assert_eq!(
        result.skip,
        expected["stats"]["skip"].as_u64().unwrap() as usize
    );

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

    for (output, content) in [
        ("mapped.txt", "payload from nan"),
        ("nan.txt", "payload from remove"),
        ("keep.txt", "payload from keep"),
    ] {
        assert_eq!(
            fs::read(input_dir.join(output)).unwrap(),
            content.as_bytes()
        );
    }
    fs::remove_dir_all(root).unwrap();
}
