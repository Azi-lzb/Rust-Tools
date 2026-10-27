use anyhow::{Context, Result};
use serde_json::{json, Value};
use std::{
    collections::BTreeMap,
    fs::{self, File},
    io::{Read, Write},
    path::{Path, PathBuf},
    process::{Command, Output, Stdio},
    sync::atomic::{AtomicU64, Ordering},
    time::{SystemTime, UNIX_EPOCH},
};
use zip::ZipArchive;

static NEXT_TEST_DIRECTORY_ID: AtomicU64 = AtomicU64::new(0);

struct TestDirectory(PathBuf);

impl TestDirectory {
    fn new() -> Result<Self> {
        let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let serial = NEXT_TEST_DIRECTORY_ID.fetch_add(1, Ordering::Relaxed);
        let path = std::env::temp_dir().join(format!(
            "pytools-submission-check-{}-{nonce}-{serial}",
            std::process::id()
        ));
        fs::create_dir_all(&path)?;
        Ok(Self(path))
    }
}

impl Drop for TestDirectory {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

#[test]
fn submission_check_cli_matches_python_workbook_and_statistics() -> Result<()> {
    let test_directory = TestDirectory::new()?;
    let root = &test_directory.0;
    let config = create_config(&root)?;
    let sources = root.join("sources");
    fs::create_dir_all(&sources)?;
    let files = [
        sources.join("2026_Area_Bank_base.txt"),
        sources.join("2026_Area_Bank_multi.txt"),
        sources.join("misc_submission.txt"),
    ];
    for (index, path) in files.iter().enumerate() {
        fs::write(path, format!("synthetic payload {index}"))?;
    }
    let original_config = fs::read(&config)?;
    let original_files = files
        .iter()
        .map(|path| fs::read(path).map(|bytes| (path.clone(), bytes)))
        .collect::<std::io::Result<Vec<_>>>()?;

    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("submission-check")
        .arg("--config")
        .arg(&config)
        .arg("--output-dir")
        .arg(root.join("output"))
        .args(&files)
        .output()
        .context("run real submission-check CLI")?;
    assert!(
        output.status.success(),
        "submission-check failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report: Value = serde_json::from_slice(&output.stdout)?;
    let expected: Value = serde_json::from_str(include_str!(
        "../probe/reports/submission_check_python_baseline.json"
    ))?;
    for field in [
        "rules",
        "files",
        "submitted",
        "missing",
        "unmatched",
        "multi",
    ] {
        assert_eq!(
            report[field], expected["statistics"][field],
            "statistic {field}"
        );
    }
    let report_path = PathBuf::from(report["output"].as_str().context("report path missing")?);
    assert!(report_path.is_file());

    let actual = excel_reader::read_workbook(&report_path)?;
    let expected_sheets = expected["output_workbook"]["sheets"]
        .as_array()
        .context("baseline sheets must be an array")?;
    assert_eq!(
        actual
            .sheets
            .iter()
            .map(|sheet| sheet.name.as_str())
            .collect::<Vec<_>>(),
        expected_sheets
            .iter()
            .map(|sheet| sheet["name"]
                .as_str()
                .context("sheet name must be a string"))
            .collect::<Result<Vec<_>>>()?
    );
    for (sheet_index, (actual_sheet, expected_sheet)) in
        actual.sheets.iter().zip(expected_sheets).enumerate()
    {
        assert!(
            actual_sheet.formulas.is_empty(),
            "unexpected formulas in {}",
            actual_sheet.name
        );
        assert!(
            actual_sheet.merges.is_empty(),
            "unexpected merged cells in {}",
            actual_sheet.name
        );
        assert!(
            actual_sheet.comments.is_empty(),
            "unexpected comments in {}",
            actual_sheet.name
        );
        assert!(actual_sheet.conditional_formatting.is_empty());
        assert!(actual_sheet.data_validations.is_empty());
        let actual_cells = actual_sheet
            .values
            .iter()
            .filter_map(|(address, value)| {
                normalize_cell(value, &root).map(|normalized| (address.clone(), normalized))
            })
            .collect::<BTreeMap<_, _>>();
        let expected_cells =
            serde_json::from_value::<BTreeMap<String, Value>>(expected_sheet["cells"].clone())?;
        assert_eq!(
            actual_cells, expected_cells,
            "worksheet {} cells",
            actual_sheet.name
        );

        let (max_row, max_column) = actual_sheet
            .values
            .keys()
            .filter_map(|address| parse_cell_address(address))
            .fold((0usize, 0usize), |(max_row, max_column), (row, column)| {
                (max_row.max(row), max_column.max(column))
            });
        assert_eq!(
            max_row,
            expected_sheet["max_row"].as_u64().unwrap() as usize
        );
        assert_eq!(
            max_column,
            expected_sheet["max_column"].as_u64().unwrap() as usize
        );
        let expected_bold =
            serde_json::from_value::<Vec<bool>>(expected_sheet["header_bold"].clone())?;
        assert_eq!(
            header_bold_flags(&report_path, sheet_index, expected_bold.len())?,
            expected_bold
        );
    }
    assert_eq!(
        actual.sheets[0].values.get("B5"),
        Some(&json!({
            "type": "string",
            "value": "2026-06-30 01:02:03.123000"
        })),
        "Excel date config values must match Python's string form"
    );

    let recursive_sources = root.join("recursive-sources");
    for (relative_path, (source_path, bytes)) in [
        "division-a/2026_Area_Bank_base.txt",
        "division-a/nested/2026_Area_Bank_multi.txt",
        "division-b/misc_submission.txt",
    ]
    .into_iter()
    .zip(original_files.iter())
    {
        let destination = recursive_sources.join(relative_path);
        fs::create_dir_all(
            destination
                .parent()
                .context("recursive input parent missing")?,
        )?;
        assert_eq!(fs::read(source_path)?, *bytes);
        fs::write(destination, bytes)?;
    }
    let recursive_output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("submission-check")
        .arg("--config")
        .arg(&config)
        .arg("--output-dir")
        .arg(root.join("recursive-output"))
        .arg(&recursive_sources)
        .output()
        .context("run submission-check CLI with a recursively selected directory")?;
    assert!(
        recursive_output.status.success(),
        "recursive submission-check failed: {}",
        String::from_utf8_lossy(&recursive_output.stderr)
    );
    let recursive_report: Value = serde_json::from_slice(&recursive_output.stdout)?;
    for field in [
        "rules",
        "files",
        "submitted",
        "missing",
        "unmatched",
        "multi",
    ] {
        assert_eq!(
            recursive_report[field], report[field],
            "recursive directory statistic {field} must match explicit file list"
        );
    }
    assert!(PathBuf::from(
        recursive_report["output"]
            .as_str()
            .context("recursive report path missing")?
    )
    .is_file());

    assert_eq!(fs::read(&config)?, original_config);
    for (path, bytes) in original_files {
        assert_eq!(fs::read(&path)?, bytes, "input changed: {}", path.display());
    }
    Ok(())
}

#[test]
fn submission_check_uses_and_between_columns_and_or_for_institution_aliases() -> Result<()> {
    let test_directory = TestDirectory::new()?;
    let root = &test_directory.0;
    let config = create_and_or_submission_config(root)?;
    let sources = root.join("sources");
    fs::create_dir_all(&sources)?;
    let names = [
        "惠城区_中国工商银行股份有限公司_单家机构.xlsx",
        "惠城区_工行_单家机构.xlsx",
        "惠城区_未知机构_单家机构.xlsx",
        "惠城区_工行_贷款报表.xlsx",
        "龙门县_工行_单家机构.xlsx",
    ];
    let files = names
        .iter()
        .enumerate()
        .map(|(index, name)| {
            let path = sources.join(name);
            fs::write(&path, format!("synthetic payload {index}"))?;
            Ok(path)
        })
        .collect::<std::io::Result<Vec<_>>>()?;
    let original_config = fs::read(&config)?;
    let original_files = files
        .iter()
        .map(|path| fs::read(path).map(|bytes| (path.clone(), bytes)))
        .collect::<std::io::Result<Vec<_>>>()?;

    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("submission-check")
        .arg("--config")
        .arg(&config)
        .arg("--output-dir")
        .arg(root.join("output"))
        .arg("--log-dir")
        .arg(root.join("logs"))
        .args(&files)
        .output()
        .context("run submission-check CLI for Chinese AND/OR keywords")?;
    assert!(
        output.status.success(),
        "submission-check failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report: Value = serde_json::from_slice(&output.stdout)?;
    assert_eq!(report["rules"], 1);
    assert_eq!(report["files"], 5);
    assert_eq!(report["submitted"], 1);
    assert_eq!(report["missing"], 0);
    assert_eq!(report["unmatched"], 3);

    let report_path = PathBuf::from(report["output"].as_str().context("report path missing")?);
    let workbook = excel_reader::read_workbook(&report_path)?;
    let result_sheet = &workbook.sheets[0];
    assert_eq!(result_sheet.name, "检查结果");
    assert_eq!(
        result_sheet.values["F2"]["value"].as_f64(),
        Some(2.0),
        "both a full name and its short alias satisfy the institution group"
    );
    assert_eq!(
        result_sheet.values["G2"]["value"].as_str(),
        Some("惠城区_中国工商银行股份有限公司_单家机构.xlsx；惠城区_工行_单家机构.xlsx")
    );
    let unmatched_sheet = &workbook.sheets[2];
    assert_eq!(unmatched_sheet.name, "未识别文件");
    assert_eq!(unmatched_sheet.values["A2"]["value"], json!(names[2]));
    assert_eq!(unmatched_sheet.values["A3"]["value"], json!(names[3]));
    assert_eq!(unmatched_sheet.values["A4"]["value"], json!(names[4]));

    assert_eq!(fs::read(&config)?, original_config);
    for (path, bytes) in original_files {
        assert_eq!(fs::read(&path)?, bytes, "input changed: {}", path.display());
    }
    Ok(())
}

#[test]
fn submission_check_menu_uses_and_between_columns_and_or_for_institution_aliases() -> Result<()> {
    let test_directory = TestDirectory::new()?;
    let root = &test_directory.0;
    let config = create_and_or_submission_config(root)?;
    let sources = root.join("sources");
    fs::create_dir_all(&sources)?;
    fs::create_dir_all(root.join("output"))?;
    fs::create_dir_all(root.join("logs"))?;
    let names = [
        "惠城区_中国工商银行股份有限公司_单家机构.xlsx",
        "惠城区_工行_单家机构.xlsx",
        "惠城区_未知机构_单家机构.xlsx",
        "惠城区_工行_贷款报表.xlsx",
        "龙门县_工行_单家机构.xlsx",
    ];
    let files = names
        .iter()
        .enumerate()
        .map(|(index, name)| {
            let path = sources.join(name);
            fs::write(&path, format!("synthetic payload {index}"))?;
            Ok(path)
        })
        .collect::<std::io::Result<Vec<_>>>()?;
    let original_config = fs::read(&config)?;
    let original_files = files
        .iter()
        .map(|path| fs::read(path).map(|bytes| (path.clone(), bytes)))
        .collect::<std::io::Result<Vec<_>>>()?;

    let input = format!("4\n6\n2\n{}\n0\n", sources.display());
    let output = run_menu(&config, &input)?;
    assert!(
        output.status.success(),
        "submission-check menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let start = output
        .stdout
        .iter()
        .position(|byte| *byte == b'{')
        .context("submission-check menu JSON output missing")?;
    let report: Value = serde_json::Deserializer::from_slice(&output.stdout[start..])
        .into_iter::<Value>()
        .next()
        .context("submission-check menu report missing")??;
    assert_eq!(report["rules"], 1);
    assert_eq!(report["files"], 5);
    assert_eq!(report["submitted"], 1);
    assert_eq!(report["missing"], 0);
    assert_eq!(report["unmatched"], 3);

    let report_path = PathBuf::from(report["output"].as_str().context("report path missing")?);
    let workbook = excel_reader::read_workbook(&report_path)?;
    let result_sheet = &workbook.sheets[0];
    assert_eq!(result_sheet.name, "检查结果");
    assert_eq!(
        result_sheet.values["F2"]["value"].as_f64(),
        Some(2.0),
        "both a full name and its short alias satisfy the institution group"
    );
    assert_eq!(
        {
            let mut matched = result_sheet.values["G2"]["value"]
                .as_str()
                .context("matched-file list missing")?
                .split('；')
                .collect::<Vec<_>>();
            matched.sort_unstable();
            matched
        },
        [
            "惠城区_中国工商银行股份有限公司_单家机构.xlsx",
            "惠城区_工行_单家机构.xlsx",
        ]
    );
    let unmatched_sheet = &workbook.sheets[2];
    assert_eq!(unmatched_sheet.name, "未识别文件");
    let mut unmatched = ["A2", "A3", "A4"]
        .map(|address| {
            unmatched_sheet.values[address]["value"]
                .as_str()
                .map(str::to_owned)
                .context("unmatched-file name missing")
        })
        .into_iter()
        .collect::<Result<Vec<_>>>()?;
    unmatched.sort_unstable();
    let mut expected_unmatched = names[2..]
        .iter()
        .map(|name| (*name).to_owned())
        .collect::<Vec<_>>();
    expected_unmatched.sort_unstable();
    assert_eq!(unmatched, expected_unmatched);

    assert_eq!(fs::read(&config)?, original_config);
    for (path, bytes) in original_files {
        assert_eq!(fs::read(&path)?, bytes, "input changed: {}", path.display());
    }
    Ok(())
}

#[test]
fn submission_check_menu_matches_python_date_keyword_semantics() -> Result<()> {
    let test_directory = TestDirectory::new()?;
    let root = &test_directory.0;
    let config = create_config(root)?;
    let sources = root.join("sources");
    fs::create_dir_all(&sources)?;
    fs::create_dir_all(root.join("output"))?;
    fs::create_dir_all(root.join("logs"))?;
    for (index, name) in [
        "2026_Area_Bank_base.txt",
        "2026_Area_Bank_multi.txt",
        "misc_submission.txt",
    ]
    .iter()
    .enumerate()
    {
        fs::write(sources.join(name), format!("synthetic payload {index}"))?;
    }
    let original_config = fs::read(&config)?;
    let original_files = fs::read_dir(&sources)?
        .map(|entry| {
            let path = entry?.path();
            fs::read(&path).map(|bytes| (path, bytes))
        })
        .collect::<std::io::Result<Vec<_>>>()?;

    let input = format!("4\n6\n2\n{}\n0\n", sources.display());
    let output = run_menu(&config, &input)?;
    assert!(
        output.status.success(),
        "submission-check menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let start = output
        .stdout
        .iter()
        .position(|byte| *byte == b'{')
        .context("submission-check menu JSON output missing")?;
    let report: Value = serde_json::Deserializer::from_slice(&output.stdout[start..])
        .into_iter::<Value>()
        .next()
        .context("submission-check menu report missing")??;
    let expected: Value = serde_json::from_str(include_str!(
        "../probe/reports/submission_check_python_baseline.json"
    ))?;
    for field in [
        "rules",
        "files",
        "submitted",
        "missing",
        "unmatched",
        "multi",
    ] {
        assert_eq!(
            report[field], expected["statistics"][field],
            "menu statistic {field}"
        );
    }

    let report_path = PathBuf::from(report["output"].as_str().context("report path missing")?);
    let actual = excel_reader::read_workbook(&report_path)?;
    assert_eq!(
        actual.sheets[0].values.get("B5"),
        Some(&json!({
            "type": "string",
            "value": "2026-06-30 01:02:03.123000"
        })),
        "menu output must retain Python's date-keyword string form"
    );

    assert_eq!(fs::read(&config)?, original_config);
    for (path, bytes) in original_files {
        assert_eq!(fs::read(&path)?, bytes, "input changed: {}", path.display());
    }
    Ok(())
}

#[test]
fn submission_check_menu_logs_user_developer_and_off_views() -> Result<()> {
    let test_directory = TestDirectory::new()?;
    let root = &test_directory.0;
    let config = create_config(root)?;
    let sources = root.join("sources");
    fs::create_dir_all(&sources)?;
    fs::create_dir_all(root.join("output"))?;
    fs::create_dir_all(root.join("logs"))?;
    for (index, name) in [
        "2026_Area_Bank_base.txt",
        "2026_Area_Bank_multi.txt",
        "misc_submission.txt",
    ]
    .iter()
    .enumerate()
    {
        fs::write(sources.join(name), format!("synthetic payload {index}"))?;
    }
    let input = format!("4\n6\n2\n{}\n0\n", sources.display());

    for log_view in ["user", "off", "developer"] {
        let output = run_menu_with_log_view(&config, &input, log_view)?;
        assert!(
            output.status.success(),
            "{log_view} menu failed: {}",
            String::from_utf8_lossy(&output.stderr)
        );
        let report_start = output
            .stdout
            .iter()
            .position(|byte| *byte == b'{')
            .context("submission-check menu JSON output missing")?;
        let report: Value = serde_json::Deserializer::from_slice(&output.stdout[report_start..])
            .into_iter::<Value>()
            .next()
            .context("submission-check menu report missing")??;
        assert_eq!(report["files"], 3);
        assert_eq!(report["submitted"], 2);

        let stderr = String::from_utf8_lossy(&output.stderr);
        match log_view {
            "user" => {
                assert!(stderr.contains("正在读取提交检查配置"));
                assert!(stderr.contains("阶段用时"));
                assert!(stderr.contains("总用时"));
                assert!(!stderr.contains(&sources.to_string_lossy().to_string()));
                assert!(!stderr.contains("2026_Area_Bank_base.txt"));
            }
            "off" => assert!(stderr.trim().is_empty(), "off view emitted: {stderr}"),
            "developer" => {
                let events = stderr
                    .lines()
                    .map(serde_json::from_str::<Value>)
                    .collect::<std::result::Result<Vec<_>, _>>()?;
                let phases = events
                    .iter()
                    .filter_map(|event| event["phase"].as_str())
                    .collect::<Vec<_>>();
                for phase in [
                    "configuration_load",
                    "input_scan",
                    "filename_match",
                    "report_write",
                ] {
                    assert!(phases.contains(&phase), "missing developer phase {phase}");
                }
                assert!(events.iter().any(|event| {
                    event["phase"] == "filename_match"
                        && event["status"] == "progress"
                        && event["progress"]["processed_rules"].is_number()
                }));
                assert!(events.iter().any(|event| {
                    event["status"] == "completed"
                        && event["stage_elapsed_ms"].is_number()
                        && event["total_elapsed_ms"].is_number()
                }));
                assert!(events.iter().any(|event| {
                    event["actual_engine"] == "calamine" && event["status"] == "completed"
                }));
                assert!(events.iter().any(|event| {
                    event["actual_engine"] == "rust_xlsxwriter" && event["status"] == "completed"
                }));
                assert!(!stderr.contains(&sources.to_string_lossy().to_string()));
                assert!(!stderr.contains("2026_Area_Bank_base.txt"));

                let history_path = PathBuf::from(
                    report["run_history_path"]
                        .as_str()
                        .context("developer run history path missing")?,
                );
                let history_text = fs::read_to_string(&history_path)?;
                let history: Value = serde_json::from_str(&history_text)?;
                assert_eq!(history["feature_id"], "PT-4-8");
                assert_eq!(history["status"], "succeeded");
                assert!(history["actual_engines"]
                    .as_array()
                    .is_some_and(|engines| engines.iter().any(|engine| engine == "calamine")));
                assert!(history["actual_engines"].as_array().is_some_and(|engines| {
                    engines.iter().any(|engine| engine == "rust_xlsxwriter")
                }));
                assert!(!history_text.contains(&sources.to_string_lossy().to_string()));
                assert!(!history_text.contains("2026_Area_Bank_base.txt"));
            }
            _ => unreachable!(),
        }
    }

    let history_count = fs::read_dir(root.join("logs"))?
        .filter_map(|entry| entry.ok())
        .filter(|entry| {
            entry
                .path()
                .extension()
                .is_some_and(|extension| extension == "json")
        })
        .count();
    assert_eq!(history_count, 1, "only developer view writes history");
    Ok(())
}

#[test]
fn submission_check_developer_history_records_report_write_failure() -> Result<()> {
    let test_directory = TestDirectory::new()?;
    let root = &test_directory.0;
    let config = create_config(root)?;
    let sources = root.join("sources");
    fs::create_dir_all(&sources)?;
    let input = sources.join("2026_Area_Bank_base.txt");
    fs::write(&input, b"synthetic payload")?;
    let blocked_output = root.join("output-is-a-file");
    fs::write(&blocked_output, b"must remain unchanged")?;
    let logs = root.join("logs");
    fs::create_dir_all(&logs)?;
    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("submission-check")
        .arg("--config")
        .arg(&config)
        .arg("--output-dir")
        .arg(&blocked_output)
        .arg("--log-dir")
        .arg(&logs)
        .arg("--log-view")
        .arg("developer")
        .arg(&input)
        .output()
        .context("run submission-check with an unwritable output directory")?;
    assert!(!output.status.success());
    assert_eq!(fs::read(&blocked_output)?, b"must remain unchanged");
    assert_eq!(fs::read(&input)?, b"synthetic payload");

    let stderr = String::from_utf8_lossy(&output.stderr);
    let events = stderr
        .lines()
        .filter_map(|line| serde_json::from_str::<Value>(line).ok())
        .collect::<Vec<_>>();
    let failed_write = events
        .iter()
        .find(|event| event["phase"] == "report_write" && event["status"] == "failed")
        .context("missing developer report-write failure event")?;
    assert_eq!(failed_write["error_type"], "report_write_failed");

    let histories = fs::read_dir(&logs)?
        .map(|entry| entry.map(|entry| entry.path()))
        .collect::<std::io::Result<Vec<_>>>()?;
    assert_eq!(histories.len(), 1);
    assert!(stderr.contains(&histories[0].to_string_lossy().to_string()));
    let history_text = fs::read_to_string(&histories[0])?;
    let history: Value = serde_json::from_str(&history_text)?;
    assert_eq!(history["status"], "failed");
    assert!(history["summary"].is_null());
    assert!(!history_text.contains(&sources.to_string_lossy().to_string()));
    assert!(!history_text.contains("2026_Area_Bank_base.txt"));
    Ok(())
}

fn run_menu(config: &Path, input: &str) -> Result<Output> {
    run_menu_with_log_view(config, input, "user")
}

fn run_menu_with_log_view(config: &Path, input: &str, log_view: &str) -> Result<Output> {
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("menu")
        .arg("--config")
        .arg(config)
        .arg("--log-view")
        .arg(log_view)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .context("start interactive PyTools menu")?;
    child
        .stdin
        .take()
        .context("interactive stdin missing")?
        .write_all(input.as_bytes())
        .context("send interactive menu selections")?;
    child
        .wait_with_output()
        .context("wait for interactive PyTools menu")
}

fn create_config(root: &Path) -> Result<PathBuf> {
    let config = root.join("config.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    {
        let sheet = workbook.add_worksheet();
        sheet.set_name("全局配置")?;
        let output = root.join("output");
        let logs = root.join("logs");
        for (row, (key, value)) in [
            ("输入目录", root.to_string_lossy().into_owned()),
            ("输出目录", output.to_string_lossy().into_owned()),
            ("日志目录", logs.to_string_lossy().into_owned()),
        ]
        .iter()
        .enumerate()
        {
            sheet.write_string(row as u32, 0, *key)?;
            sheet.write_string(row as u32, 1, value)?;
        }
    }
    let sheet = workbook.add_worksheet();
    sheet.set_name("提交检查配置")?;
    for (column, header) in [
        "是否启用",
        "日期关键词",
        "县区关键词",
        "机构关键词",
        "备用关键词",
        "备注",
    ]
    .iter()
    .enumerate()
    {
        sheet.write_string(0, column as u16, *header)?;
    }
    let rules = [
        ["是", "2026", "Area", "Bank", "", "base rule"],
        ["是", "2026", "Area", "Bank", "multi", "second rule"],
        ["是", "2026", "Missing", "Bank", "", "missing rule"],
        ["否", "disabled", "Area", "Bank", "", "ignored rule"],
    ];
    for (row, values) in rules.iter().enumerate() {
        for (column, value) in values.iter().enumerate() {
            sheet.write_string((row + 1) as u32, column as u16, *value)?;
        }
    }
    sheet.write_string(5, 0, "是")?;
    let date =
        rust_xlsxwriter::ExcelDateTime::from_ymd(2026, 6, 30)?.and_hms_milli(1, 2, 3, 123)?;
    let date_format = rust_xlsxwriter::Format::new().set_num_format("yyyy-mm-dd");
    sheet.write_datetime_with_format(5, 1, &date, &date_format)?;
    sheet.write_string(5, 2, "Area")?;
    sheet.write_string(5, 3, "Bank")?;
    sheet.write_string(5, 5, "date keyword")?;
    workbook.save(&config)?;
    Ok(config)
}

fn create_and_or_submission_config(root: &Path) -> Result<PathBuf> {
    let config = root.join("submission-and-or-config.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let global = workbook.add_worksheet();
    global.set_name("全局配置")?;
    for (row, (key, value)) in [
        ("输入目录", root.to_string_lossy().into_owned()),
        (
            "输出目录",
            root.join("output").to_string_lossy().into_owned(),
        ),
        ("日志目录", root.join("logs").to_string_lossy().into_owned()),
    ]
    .iter()
    .enumerate()
    {
        global.write_string(row as u32, 0, *key)?;
        global.write_string(row as u32, 1, value)?;
    }
    let sheet = workbook.add_worksheet();
    sheet.set_name("提交检查配置")?;
    for (column, header) in [
        "是否启用",
        "日期关键词",
        "县区关键词",
        "机构关键词",
        "备用关键词",
        "备注",
    ]
    .iter()
    .enumerate()
    {
        sheet.write_string(0, column as u16, *header)?;
    }
    sheet.write_string(1, 0, "是")?;
    sheet.write_string(1, 2, "惠城区")?;
    sheet.write_string(
        1,
        3,
        "中国工商银行股份有限公司;中国工商银行;中国工商银行;工行",
    )?;
    sheet.write_string(1, 4, "单家机构")?;
    sheet.write_string(1, 5, "A07 institution alias contract")?;
    workbook.save(&config)?;
    Ok(config)
}

fn normalize_cell(value: &Value, root: &Path) -> Option<Value> {
    match value["type"].as_str()? {
        "empty" => None,
        "string" => {
            let text = value["value"].as_str()?;
            if text.is_empty() {
                return None;
            }
            Some(json!({
                "type": "string",
                "value": text.replace(root.to_string_lossy().as_ref(), "<ROOT>"),
            }))
        }
        "integer" => Some(json!({ "type": "number", "value": value["value"] })),
        "number" => {
            let number = value["value"].as_f64()?;
            let normalized = if number.is_finite() && number.fract() == 0.0 {
                json!(number as i64)
            } else {
                json!(number)
            };
            Some(json!({ "type": "number", "value": normalized }))
        }
        "boolean" => Some(json!({ "type": "boolean", "value": value["value"] })),
        kind => Some(json!({ "type": kind, "value": value["value"] })),
    }
}

fn parse_cell_address(address: &str) -> Option<(usize, usize)> {
    let split = address.find(|character: char| character.is_ascii_digit())?;
    let (letters, digits) = address.split_at(split);
    if letters.is_empty() || digits.is_empty() {
        return None;
    }
    let column = letters.chars().try_fold(0usize, |value, character| {
        let digit = character.to_ascii_uppercase().to_digit(36)? as usize - 9;
        value.checked_mul(26)?.checked_add(digit)
    })?;
    let row = digits.parse::<usize>().ok()?;
    Some((row, column))
}

fn header_bold_flags(path: &Path, sheet_index: usize, columns: usize) -> Result<Vec<bool>> {
    let file = File::open(path)?;
    let mut archive = ZipArchive::new(file)?;
    let styles_xml = read_zip_entry(&mut archive, "xl/styles.xml")?;
    let styles = roxmltree::Document::parse(&styles_xml)?;
    let fonts = styles
        .root_element()
        .descendants()
        .find(|node| node.tag_name().name() == "fonts")
        .context("styles.xml has no fonts")?;
    let font_bold = fonts
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "font")
        .map(|font| {
            font.children().any(|node| {
                node.is_element()
                    && node.tag_name().name() == "b"
                    && node.attribute("val") != Some("0")
            })
        })
        .collect::<Vec<_>>();
    let cell_xfs = styles
        .root_element()
        .descendants()
        .find(|node| node.tag_name().name() == "cellXfs")
        .context("styles.xml has no cellXfs")?;
    let style_bold = cell_xfs
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "xf")
        .map(|xf| {
            xf.attribute("fontId")
                .and_then(|value| value.parse::<usize>().ok())
                .and_then(|font_id| font_bold.get(font_id).copied())
                .unwrap_or(false)
        })
        .collect::<Vec<_>>();
    let sheet_xml = read_zip_entry(
        &mut archive,
        &format!("xl/worksheets/sheet{}.xml", sheet_index + 1),
    )?;
    let sheet = roxmltree::Document::parse(&sheet_xml)?;
    let mut header_styles = BTreeMap::new();
    for cell in sheet
        .root_element()
        .descendants()
        .filter(|node| node.tag_name().name() == "c")
    {
        let Some(address) = cell.attribute("r") else {
            continue;
        };
        if parse_cell_address(address).is_some_and(|(row, _)| row == 1) {
            let style = cell
                .attribute("s")
                .and_then(|value| value.parse::<usize>().ok())
                .unwrap_or(0);
            header_styles.insert(
                address.to_owned(),
                style_bold.get(style).copied().unwrap_or(false),
            );
        }
    }
    (1..=columns)
        .map(|column| {
            header_styles
                .get(&format!("{}1", column_letters(column)))
                .copied()
                .context("header cell missing from worksheet XML")
        })
        .collect()
}

fn read_zip_entry(archive: &mut ZipArchive<File>, name: &str) -> Result<String> {
    let mut entry = archive
        .by_name(name)
        .with_context(|| format!("missing {name}"))?;
    let mut xml = String::new();
    entry.read_to_string(&mut xml)?;
    Ok(xml)
}

fn column_letters(mut column: usize) -> String {
    let mut letters = Vec::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        letters.push((b'A' + remainder as u8) as char);
        column = (column - 1) / 26;
    }
    letters.into_iter().rev().collect()
}
