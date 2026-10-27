use pytools_core::batch_rename::{run_batch_rename, run_batch_rename_with_progress};
use rust_xlsxwriter::{ExcelDateTime, Format, Workbook};
use serde_json::Value;
use std::{
    fs,
    path::PathBuf,
    process::Command,
    time::{SystemTime, UNIX_EPOCH},
};

struct TestDirectory(PathBuf);

impl TestDirectory {
    fn new(label: &str) -> Self {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock must be after Unix epoch")
            .as_nanos();
        let path = std::env::temp_dir().join(format!(
            "pytools-batch-rename-{label}-{}-{nonce}",
            std::process::id()
        ));
        fs::create_dir_all(&path).expect("create isolated test directory");
        Self(path)
    }
}

impl Drop for TestDirectory {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

fn cases() -> Value {
    serde_json::from_str(include_str!("fixtures/batch_rename_cases.json"))
        .expect("batch rename cases must be valid JSON")
}

fn setup(label: &str) -> (TestDirectory, PathBuf, Vec<PathBuf>, Vec<String>) {
    let cases = cases();
    let root = TestDirectory::new(label);
    let config_path = root.0.join("config.xlsx");
    let fixture =
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/batch_rename_config.xlsx");
    fs::copy(fixture, &config_path).expect("copy synthetic rename config");

    let mut paths = Vec::new();
    for input in cases["inputs"].as_array().unwrap() {
        let path = root.0.join(input["name"].as_str().unwrap());
        paths.push(path.clone());
        if !input["exists"].as_bool().unwrap() {
            continue;
        }
        match input["kind"].as_str().unwrap() {
            "file" => {
                fs::write(&path, input["content"].as_str().unwrap()).expect("write synthetic input")
            }
            "directory" => {
                fs::create_dir_all(&path).expect("create synthetic input directory");
                fs::write(path.join("marker.txt"), input["content"].as_str().unwrap())
                    .expect("write directory marker");
            }
            other => panic!("unknown input kind {other:?}"),
        }
    }
    for input in cases["reserved"].as_array().unwrap() {
        let path = root.0.join(input["name"].as_str().unwrap());
        fs::write(&path, input["content"].as_str().unwrap()).expect("write reserved target");
    }

    let baseline_entries = include_str!("../probe/reports/batch_rename_python_baseline.json");
    let baseline: Value =
        serde_json::from_str(baseline_entries).expect("Python baseline report must be valid JSON");
    let expected = baseline["result"]["entries"]
        .as_array()
        .unwrap()
        .iter()
        .map(|entry| entry["name"].as_str().unwrap().to_owned())
        .collect();
    (root, config_path, paths, expected)
}

fn top_level_names(root: &std::path::Path, config: &std::path::Path) -> Vec<String> {
    let mut names = fs::read_dir(root)
        .expect("list test directory")
        .map(|entry| entry.expect("read directory entry").path())
        .filter(|path| path != config)
        .filter(|path| path.file_name().is_none_or(|name| name != "logs"))
        .filter_map(|path| {
            path.file_name()
                .map(|name| name.to_string_lossy().into_owned())
        })
        .collect::<Vec<_>>();
    names.sort();
    names
}

#[test]
fn rust_operation_matches_python_baseline_and_keeps_contents_and_collisions() {
    let (root, config_path, paths, expected_entries) = setup("api");
    let original_config = fs::read(&config_path).unwrap();
    let result = run_batch_rename(&config_path, &paths).unwrap();
    let baseline: Value = serde_json::from_str(include_str!(
        "../probe/reports/batch_rename_python_baseline.json"
    ))
    .unwrap();

    assert_eq!(
        result.ok,
        baseline["result"]["ok"].as_u64().unwrap() as usize
    );
    assert_eq!(
        result.skip,
        baseline["result"]["skip"].as_u64().unwrap() as usize
    );
    assert_eq!(result.ok, 3);
    assert_eq!(result.skip, 4);
    let skip_reasons = result
        .rows
        .iter()
        .filter(|row| row.status == "skipped")
        .map(|row| row.reason.split_whitespace().next().unwrap().to_owned())
        .collect::<std::collections::BTreeSet<_>>();
    assert_eq!(
        skip_reasons,
        [
            "full_name_no_code",
            "name_same",
            "not_exists",
            "short_name_not_matched"
        ]
        .into_iter()
        .map(str::to_owned)
        .collect()
    );
    assert_eq!(top_level_names(&root.0, &config_path), expected_entries);
    assert_eq!(
        fs::read_to_string(root.0.join("2026-01-31 0008 短名优先银行 存款表_2.xlsx")).unwrap(),
        "source-one"
    );
    assert_eq!(
        fs::read_to_string(root.0.join("2026-01-31 0008 短名优先银行 存款表_1.xlsx")).unwrap(),
        "reserved-one"
    );
    assert_eq!(
        fs::read_to_string(root.0.join("2026-01-31 0008 短名优先银行 存款表.xlsx")).unwrap(),
        "existing-target"
    );
    assert_eq!(
        fs::read_to_string(root.0.join("2026-01-31 0003 丙银行 存款表/marker.txt")).unwrap(),
        baseline["result"]["renamed_directory_marker"]
            .as_str()
            .unwrap()
    );
    assert_eq!(
        fs::read_to_string(root.0.join("2026-01-31 A01 英文银行 存款表.xlsm")).unwrap(),
        "case-insensitive"
    );
    assert_eq!(fs::read(&config_path).unwrap(), original_config);
}

#[test]
fn rust_cli_runs_the_same_operation_and_writes_a_log() {
    let (root, config_path, paths, expected_entries) = setup("cli");
    let log_dir = root.0.join("logs");
    let mut command = Command::new(env!("CARGO_BIN_EXE_pytools"));
    command
        .arg("rename-batch")
        .arg("--config")
        .arg(&config_path)
        .arg("--log-dir")
        .arg(&log_dir)
        .args(&paths);
    let output = command.output().expect("run pytools CLI");
    assert!(
        output.status.success(),
        "CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report: Value = serde_json::from_slice(&output.stdout).expect("CLI emits JSON");
    assert_eq!(report["ok"], 3);
    assert_eq!(report["skip"], 4);
    let log_path = PathBuf::from(report["log_path"].as_str().unwrap());
    let log = fs::read_to_string(&log_path).expect("CLI writes UTF-8 operation log");
    assert!(log.contains("[INFO] ok old="));
    assert!(log.contains("full_name_no_code"));
    assert_eq!(top_level_names(&root.0, &config_path), expected_entries);
}

#[test]
fn missing_rename_sheet_is_rejected_before_touching_inputs() {
    let root = TestDirectory::new("invalid-config");
    let config_path =
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/config_loader.xlsx");
    let input = root.0.join("建行.xlsx");
    fs::write(&input, b"unchanged").unwrap();

    let mut events = Vec::new();
    let error = run_batch_rename_with_progress(&config_path, &[input.clone()], |event| {
        events.push(event.clone());
    })
    .unwrap_err();
    assert!(error.to_string().contains("重命名配置"));
    assert_eq!(fs::read(input).unwrap(), b"unchanged");
    assert_eq!(events.len(), 2);
    assert_eq!(events[0].phase, "configuration_read");
    assert_eq!(events[0].status, "started");
    assert_eq!(events[1].phase, "configuration_read");
    assert_eq!(events[1].status, "failed");
    assert_eq!(events[1].actual_engine.as_deref(), Some("Calamine"));
    assert_eq!(events[1].error_type.as_deref(), Some("configuration_error"));
}

#[test]
fn date_serial_configuration_uses_safe_filename_and_preserves_contents() -> anyhow::Result<()> {
    let root = TestDirectory::new("datetime-setting");
    let config_path = root.0.join("config.xlsx");
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("重命名配置")?;
    for (column, header) in [
        (0, "简称"),
        (1, "全称"),
        (3, "代码"),
        (4, "全称(代码映射)"),
        (6, "键"),
        (7, "值"),
    ] {
        sheet.write_string(0, column, header)?;
    }
    sheet.write_string(1, 0, "测试")?;
    sheet.write_string(1, 1, "测试银行")?;
    sheet.write_string(1, 3, "0008")?;
    sheet.write_string(1, 4, "测试银行")?;
    sheet.write_string(1, 6, "数据日期")?;
    let date = ExcelDateTime::from_ymd(2026, 1, 31)?;
    let date_format = Format::new().set_num_format("yyyy-mm-dd");
    sheet.write_datetime_with_format(1, 7, &date, &date_format)?;
    sheet.write_string(2, 6, "报表名称")?;
    sheet.write_string(2, 7, "存款表")?;
    workbook.save(&config_path)?;

    let input_path = root.0.join("测试.xlsx");
    fs::write(&input_path, b"preserve file payload")?;
    let original_payload = fs::read(&input_path)?;
    let result = run_batch_rename(&config_path, &[input_path.clone()])?;

    let expected_name = "2026-01-31 0008 测试银行 存款表.xlsx";
    assert_eq!(result.ok, 1);
    assert_eq!(result.skip, 0);
    assert_eq!(result.rows[0].new_name, expected_name);
    assert!(!input_path.exists());
    assert_eq!(fs::read(root.0.join(expected_name))?, original_payload);
    Ok(())
}
