use anyhow::{Context, Result};
use excel_reader::read_workbook;
use serde_json::Value;
use std::{
    fs,
    io::Write,
    path::{Path, PathBuf},
    process::{Command, Output, Stdio},
    sync::atomic::{AtomicU64, Ordering},
    time::{SystemTime, UNIX_EPOCH},
};

static NEXT_TEST_DIRECTORY_ID: AtomicU64 = AtomicU64::new(0);

struct TestDirectory(PathBuf);

impl TestDirectory {
    fn new() -> Result<Self> {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .context("system clock before Unix epoch")?
            .as_nanos();
        let serial = NEXT_TEST_DIRECTORY_ID.fetch_add(1, Ordering::Relaxed);
        let path = std::env::temp_dir().join(format!(
            "pytools-timeline-menu-{}-{nonce}-{serial}",
            std::process::id()
        ));
        fs::create_dir_all(&path)?;
        Ok(Self(path))
    }

    fn path(&self) -> &Path {
        &self.0
    }
}

impl Drop for TestDirectory {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

#[test]
fn interactive_v2_data_menu_routes_compare_wide_summary_and_precheck() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config(root.path())?;
    let template = root.path().join("template.xlsx");
    let compare_source = root.path().join("sample compare source.xlsx");
    let wide_source_dir = root.path().join("z sample wide sources");
    let wide_source = wide_source_dir.join("sample wide source.xlsx");
    let nested_source_dir = wide_source_dir.join("nested");
    fs::create_dir_all(&nested_source_dir)?;
    let nested_wide_source = nested_source_dir.join("sample nested wide source.xlsx");
    let earlier_source_dir = root.path().join("a sample wide sources");
    fs::create_dir_all(&earlier_source_dir)?;
    let earlier_wide_source = earlier_source_dir.join("sample earliest source.xlsx");
    create_wide_source(&template, "模板甲", 10.0)?;
    create_wide_source(&compare_source, "来源甲", 12.0)?;
    create_wide_source(&wide_source, "宽表甲", 17.5)?;
    create_wide_source(&nested_wide_source, "宽表乙", 27.5)?;
    create_wide_source(&earlier_wide_source, "宽表丙", 37.5)?;
    let wide_source_before = fs::read(&wide_source)?;
    let nested_wide_source_before = fs::read(&nested_wide_source)?;
    let earlier_wide_source_before = fs::read(&earlier_wide_source)?;

    let compare = run_menu(
        &config,
        &format!(
            "1
4
{}
{}
0
",
            template.display(),
            compare_source.display()
        ),
    )?;
    assert!(
        compare.status.success(),
        "timeline compare menu failed: {}",
        String::from_utf8_lossy(&compare.stderr)
    );
    let compare_stdout = String::from_utf8_lossy(&compare.stdout);
    assert!(compare_stdout.contains("1  数据提取与汇总"));
    assert!(compare_stdout.contains("4  表头路径比对与映射"));
    assert!(compare_stdout.contains("表头路径比对模板"));
    assert!(compare_stdout.contains(r#""output""#));
    let output_dir = root.path().join("app").join("output");
    assert!(find_output(&output_dir, "compare_3_9_3_")?.is_file());

    let no_suffix = run_menu(
        &config,
        &format!(
            "1
1
3
{};{}
0
",
            wide_source_dir.display(),
            earlier_source_dir.display()
        ),
    )?;
    assert!(
        no_suffix.status.success(),
        "no-suffix wide summary menu failed: {}",
        String::from_utf8_lossy(&no_suffix.stderr)
    );
    assert!(String::from_utf8_lossy(&no_suffix.stdout).contains("1  时序宽表汇总"));
    let no_suffix_stdout = String::from_utf8_lossy(&no_suffix.stdout);
    assert!(no_suffix_stdout.contains("递归共识别文件 3 个"));
    assert!(no_suffix_stdout.contains("sample wide source.xlsx"));
    assert!(no_suffix_stdout.contains("sample nested wide source.xlsx"));
    assert!(no_suffix_stdout.contains("sample earliest source.xlsx"));
    let earlier_position = no_suffix_stdout
        .find(earlier_wide_source.to_string_lossy().as_ref())
        .context("earlier selected workbook is missing from menu output")?;
    let nested_position = no_suffix_stdout
        .find(nested_wide_source.to_string_lossy().as_ref())
        .context("nested selected workbook is missing from menu output")?;
    let top_level_position = no_suffix_stdout
        .find(wide_source.to_string_lossy().as_ref())
        .context("top-level selected workbook is missing from menu output")?;
    assert!(earlier_position < nested_position && nested_position < top_level_position);
    let no_suffix_output = find_output(&output_dir, "行头不加后缀")?;
    let no_suffix_dump = read_workbook(&no_suffix_output)?;
    let wide_sheet = no_suffix_dump
        .sheets
        .iter()
        .find(|sheet| sheet.name == "WideRule")
        .context("wide summary sheet missing")?;
    let wide_values = wide_sheet
        .values
        .values()
        .filter_map(|cell| cell.get("value").and_then(Value::as_str))
        .collect::<Vec<_>>();
    assert!(wide_values.contains(&"宽表甲"));
    assert!(wide_values.contains(&"宽表乙"));
    assert!(wide_values.contains(&"宽表丙"));
    assert_eq!(fs::read(&wide_source)?, wide_source_before);
    assert_eq!(fs::read(&nested_wide_source)?, nested_wide_source_before);
    assert_eq!(fs::read(&earlier_wide_source)?, earlier_wide_source_before);

    let precheck = run_menu(
        &config, "1
5
0
",
    )?;
    assert!(
        precheck.status.success(),
        "timeline precheck menu failed: {}",
        String::from_utf8_lossy(&precheck.stderr)
    );
    let precheck_stdout = String::from_utf8_lossy(&precheck.stdout);
    assert!(precheck_stdout.contains("5  时序配置预校验"));
    assert!(precheck_stdout.contains("=== 预校验完成: 错误 0 项，警告 0 项 ==="));
    assert!(precheck_stdout.contains("[完成] 时序配置预校验通过。"));
    assert!(precheck_stdout.contains("再见。"));
    assert!(!precheck_stdout.contains(r#""errors""#));
    Ok(())
}

#[test]
fn wide_summary_menu_current_folder_mode_excludes_nested_sources() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config(root.path())?;
    let source_dir = root.path().join("sample current-level sources");
    let nested_dir = source_dir.join("nested");
    fs::create_dir_all(&nested_dir)?;
    let current_source = source_dir.join("sample current-level source.xlsx");
    let nested_source = nested_dir.join("sample nested source.xlsx");
    create_wide_source(&current_source, "仅当前层", 17.5)?;
    create_wide_source(&nested_source, "不应进入", 27.5)?;
    let current_source_before = fs::read(&current_source)?;
    let nested_source_before = fs::read(&nested_source)?;

    let output = run_menu(&config, &format!("1\n1\n2\n{}\n0\n", source_dir.display()))?;
    assert!(
        output.status.success(),
        "current-folder wide summary menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("当前层共识别文件 1 个"));
    assert!(stdout.contains("sample current-level source.xlsx"));
    assert!(!stdout.contains("sample nested source.xlsx"));

    let output_dir = root.path().join("app").join("output");
    let output_path = find_output(&output_dir, "行头不加后缀")?;
    let dump = read_workbook(&output_path)?;
    let wide_sheet = dump
        .sheets
        .iter()
        .find(|sheet| sheet.name == "WideRule")
        .context("current-folder summary sheet missing")?;
    let values = wide_sheet
        .values
        .values()
        .filter_map(|cell| cell.get("value").and_then(Value::as_str))
        .collect::<Vec<_>>();
    assert!(values.contains(&"仅当前层"));
    assert!(!values.contains(&"不应进入"));
    assert_eq!(fs::read(&current_source)?, current_source_before);
    assert_eq!(fs::read(&nested_source)?, nested_source_before);
    Ok(())
}

#[test]
fn wide_summary_cli_filters_missing_and_unconfigured_extensions_before_reading() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config(root.path())?;
    let source = root.path().join("sample valid wide.xlsx");
    let unsupported = root.path().join("sample unsupported.xlsb");
    let missing = root.path().join("sample missing.xlsx");
    create_wide_source(&source, "宽表甲", 17.5)?;
    fs::write(&unsupported, b"not a supported workbook")?;
    let source_before = fs::read(&source)?;
    let unsupported_before = fs::read(&unsupported)?;

    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("wide-summary")
        .arg("--config")
        .arg(&config)
        .arg("--log-view")
        .arg("developer")
        .arg(&source)
        .arg(&unsupported)
        .arg(&missing)
        .output()
        .context("run wide-summary CLI with filtered inputs")?;
    assert!(
        output.status.success(),
        "wide-summary CLI did not filter unsupported inputs: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let report: Value = serde_json::from_slice(&output.stdout)?;
    let summary_path = report["output"]
        .as_str()
        .map(PathBuf::from)
        .context("wide-summary CLI output path missing")?;
    assert!(summary_path.is_file());
    let run_log_path = report["run_log_path"]
        .as_str()
        .map(PathBuf::from)
        .context("wide-summary run log path missing")?;
    let history: Value = serde_json::from_slice(&fs::read(run_log_path)?)?;
    assert_eq!(history["total_inputs"], 1);
    assert_eq!(history["completed_inputs"], 1);
    assert_eq!(history["failed_inputs"], 0);
    let events = history["events"]
        .as_array()
        .context("wide-summary events missing")?;
    assert!(events
        .iter()
        .all(|event| event["progress"]["total_inputs"] == 1));
    assert!(events.iter().any(|event| {
        event["phase"] == "source_filter"
            && event["status"] == "skipped"
            && event["workbook_name"] == "sample unsupported.xlsb"
    }));
    assert!(events.iter().any(|event| {
        event["phase"] == "source_filter"
            && event["status"] == "skipped"
            && event["workbook_name"] == "sample missing.xlsx"
    }));
    assert_eq!(fs::read(&source)?, source_before);
    assert_eq!(fs::read(&unsupported)?, unsupported_before);
    Ok(())
}

#[test]
fn wide_summary_cli_projects_elapsed_duration_as_numeric_for_both_row_modes() -> Result<()> {
    for command in ["wide-summary", "wide-summary-row-suffix"] {
        let root = TestDirectory::new()?;
        let config = create_timeline_config(root.path())?;
        let source = root.path().join("sample elapsed duration.xlsx");
        create_wide_source_with_elapsed_duration(&source, "累计时长", 1.123_456_789)?;

        let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
            .arg(command)
            .arg("--config")
            .arg(&config)
            .arg("--log-view")
            .arg("developer")
            .arg(&source)
            .output()
            .context("run elapsed-duration wide-summary CLI")?;
        assert!(
            output.status.success(),
            "{command} did not process elapsed duration: {}",
            String::from_utf8_lossy(&output.stderr)
        );
        let report: Value = serde_json::from_slice(&output.stdout)?;
        let summary_path = report["output"]
            .as_str()
            .map(PathBuf::from)
            .context("elapsed-duration summary output path missing")?;
        let dump = read_workbook(&summary_path)?;
        let summary = dump
            .sheets
            .iter()
            .find(|sheet| sheet.name == "WideRule")
            .context("elapsed-duration summary sheet missing")?;
        let cell = &summary.values["G2"];
        assert_eq!(cell["type"], "number");
        let actual = cell["value"]
            .as_f64()
            .context("duration serial is not numeric")?;
        assert!((actual - 1.123_456_793_981_481).abs() < 1e-15);

        let history_path = report["run_log_path"]
            .as_str()
            .map(PathBuf::from)
            .context("elapsed-duration run history missing")?;
        let history: Value = serde_json::from_slice(&fs::read(history_path)?)?;
        assert!(history["events"].as_array().is_some_and(|events| {
            events.iter().any(|event| {
                event["phase"] == "elapsed_duration_normalization"
                    && event["status"] == "completed"
                    && event["actual_engine"] == "pytools-core"
                    && event["stage_elapsed_ms"].as_u64().is_some()
            })
        }));
    }
    Ok(())
}

#[test]
fn wide_summary_cli_preserves_excel_1900_leap_boundary_dates_in_both_row_modes() -> Result<()> {
    for command in ["wide-summary", "wide-summary-row-suffix"] {
        for serial in [59.0, 60.0] {
            let root = TestDirectory::new()?;
            let config = create_timeline_config(root.path())?;
            let source = root
                .path()
                .join(format!("sample_1900_boundary_{serial}.xlsx"));
            create_wide_source_with_date_serial(&source, "1900序列边界", serial)?;
            let source_before = fs::read(&source)?;

            let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
                .arg(command)
                .arg("--config")
                .arg(&config)
                .arg("--log-view")
                .arg("developer")
                .arg(&source)
                .output()
                .context("run wide-summary CLI with an Excel 1900 boundary date")?;
            assert!(
                output.status.success(),
                "{command} did not process Excel serial {serial}: {}",
                String::from_utf8_lossy(&output.stderr)
            );

            let report: Value = serde_json::from_slice(&output.stdout)?;
            let summary_path = report["output"]
                .as_str()
                .map(PathBuf::from)
                .context("1900-boundary summary output path missing")?;
            let dump = read_workbook(&summary_path)?;
            let summary = dump
                .sheets
                .iter()
                .find(|sheet| sheet.name == "WideRule")
                .context("1900-boundary summary sheet missing")?;
            assert_eq!(summary.values["G2"]["type"], "datetime_serial");
            assert_eq!(
                summary.values["G2"]["value"].as_f64(),
                Some(serial),
                "Excel 1900 serial {serial} changed in {command}"
            );
            assert_eq!(fs::read(&source)?, source_before);
        }
    }
    Ok(())
}

#[test]
fn wide_summary_cli_projects_1900_zero_serial_as_midnight_time() -> Result<()> {
    for command in ["wide-summary", "wide-summary-row-suffix"] {
        let root = TestDirectory::new()?;
        let config = create_timeline_config(root.path())?;
        let source = root.path().join("sample_1900_zero_serial.xlsx");
        create_wide_source_with_date_serial(&source, "1900零序列", 0.0)?;
        let source_before = fs::read(&source)?;

        let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
            .arg(command)
            .arg("--config")
            .arg(&config)
            .arg("--log-view")
            .arg("developer")
            .arg(&source)
            .output()
            .context("run wide-summary CLI with Excel 1900 serial zero")?;
        assert!(
            output.status.success(),
            "{command} did not process Excel serial zero: {}",
            String::from_utf8_lossy(&output.stderr)
        );

        let report: Value = serde_json::from_slice(&output.stdout)?;
        let summary_path = report["output"]
            .as_str()
            .map(PathBuf::from)
            .context("1900-zero summary output path missing")?;
        let dump = read_workbook(&summary_path)?;
        let summary = dump
            .sheets
            .iter()
            .find(|sheet| sheet.name == "WideRule")
            .context("1900-zero summary sheet missing")?;
        assert_eq!(summary.values["G2"]["type"], "string");
        assert_eq!(summary.values["G2"]["value"], "00:00:00");
        assert_eq!(fs::read(&source)?, source_before);
    }
    Ok(())
}

#[test]
fn wide_summary_cli_preserves_time_format_serials_at_and_above_one_day() -> Result<()> {
    for command in ["wide-summary", "wide-summary-row-suffix"] {
        for serial in [1.0, 1.5] {
            let root = TestDirectory::new()?;
            let config = create_timeline_config(root.path())?;
            let source = root
                .path()
                .join(format!("sample_time_serial_{serial}.xlsx"));
            create_wide_source_with_number_format_serial(&source, "跨日时分秒", serial, "h:mm:ss")?;
            let source_before = fs::read(&source)?;

            let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
                .arg(command)
                .arg("--config")
                .arg(&config)
                .arg("--log-view")
                .arg("developer")
                .arg(&source)
                .output()
                .context("run wide-summary CLI with a full-day time-format serial")?;
            assert!(
                output.status.success(),
                "{command} did not process serial {serial}: {}",
                String::from_utf8_lossy(&output.stderr)
            );

            let report: Value = serde_json::from_slice(&output.stdout)?;
            let summary_path = report["output"]
                .as_str()
                .map(PathBuf::from)
                .context("full-day time summary output path missing")?;
            let dump = read_workbook(&summary_path)?;
            let summary = dump
                .sheets
                .iter()
                .find(|sheet| sheet.name == "WideRule")
                .context("full-day time summary sheet missing")?;
            assert_eq!(summary.values["G2"]["type"], "datetime_serial");
            assert_eq!(
                summary.values["G2"]["value"].as_f64(),
                Some(serial),
                "time-format serial {serial} changed in {command}"
            );
            assert_eq!(fs::read(&source)?, source_before);
        }
    }
    Ok(())
}

#[test]
fn wide_summary_cli_projects_clock_formats_across_midnight_rounding() -> Result<()> {
    for command in ["wide-summary", "wide-summary-row-suffix"] {
        for (format_index, (number_format, serial, expected_type)) in [
            ("h:mm", 0.5, "string"),
            ("h:mm", 0.999_999_999, "datetime_serial"),
            ("hh:mm AM/PM", 0.5, "string"),
            ("hh:mm AM/PM", 0.999_999_999, "datetime_serial"),
            ("h:mm:ss.000", 0.5, "string"),
            ("h:mm:ss.000", 0.999_999_999, "datetime_serial"),
            ("mm:ss", 0.5, "string"),
            ("mm:ss", 0.999_999_999, "datetime_serial"),
        ]
        .into_iter()
        .enumerate()
        {
            let root = TestDirectory::new()?;
            let config = create_timeline_config(root.path())?;
            let source = root
                .path()
                .join(format!("sample_clock_format_{format_index}_{serial}.xlsx"));
            create_wide_source_with_number_format_serial(
                &source,
                "时钟格式边界",
                serial,
                number_format,
            )?;
            let source_before = fs::read(&source)?;

            let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
                .arg(command)
                .arg("--config")
                .arg(&config)
                .arg("--log-view")
                .arg("developer")
                .arg(&source)
                .output()
                .context("run wide-summary CLI with a clock-only number format")?;
            assert!(
                output.status.success(),
                "{command} did not process {number_format} serial {serial}: {}",
                String::from_utf8_lossy(&output.stderr)
            );

            let report: Value = serde_json::from_slice(&output.stdout)?;
            let summary_path = report["output"]
                .as_str()
                .map(PathBuf::from)
                .context("clock-format summary output path missing")?;
            let dump = read_workbook(&summary_path)?;
            let summary = dump
                .sheets
                .iter()
                .find(|sheet| sheet.name == "WideRule")
                .context("clock-format summary sheet missing")?;
            let value = &summary.values["G2"];
            assert_eq!(
                value["type"], expected_type,
                "{number_format} serial {serial} type changed in {command}"
            );
            if expected_type == "string" {
                assert_eq!(value["value"], "12:00:00");
            } else {
                assert_eq!(value["value"].as_f64(), Some(1.0));
            }
            assert_eq!(fs::read(&source)?, source_before);
        }
    }
    Ok(())
}

#[test]
fn wide_summary_cli_processes_legacy_xls_without_ooxml_metadata() -> Result<()> {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("timeline")
        .join("wide_summary_legacy.xls");

    for command in ["wide-summary", "wide-summary-row-suffix"] {
        let root = TestDirectory::new()?;
        let config = create_legacy_xls_timeline_config(root.path())?;
        let source = root.path().join("app").join("input").join("input.xls");
        fs::copy(&fixture, &source)?;
        let source_before = fs::read(&source)?;

        let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
            .arg(command)
            .arg("--config")
            .arg(&config)
            .arg("--log-view")
            .arg("developer")
            .arg(&source)
            .output()
            .context("run wide-summary CLI with legacy XLS input")?;
        assert!(
            output.status.success(),
            "{command} did not process legacy XLS: {}",
            String::from_utf8_lossy(&output.stderr)
        );

        let report: Value = serde_json::from_slice(&output.stdout)?;
        assert!(report["output"].is_null());

        let run_log_path = report["run_log_path"]
            .as_str()
            .map(PathBuf::from)
            .context("legacy XLS run history missing")?;
        let history: Value = serde_json::from_slice(&fs::read(run_log_path)?)?;
        assert_eq!(history["status"], "skipped");
        assert_eq!(history["total_inputs"], 1);
        assert_eq!(history["completed_inputs"], 1);
        assert_eq!(history["failed_inputs"], 0);
        assert_eq!(history["matched_rule_sheets"], 1);
        assert_eq!(history["extracted_cells"], 0);
        assert!(history["events"].as_array().is_some_and(|events| {
            events.iter().any(|event| {
                event["phase"] == "elapsed_duration_metadata"
                    && event["status"] == "skipped"
                    && event["user_message"]
                        .as_str()
                        .is_some_and(|message| message.contains("格式识别未验证"))
            })
        }));
        assert_eq!(fs::read(&source)?, source_before);
    }
    Ok(())
}

#[test]
fn wide_summary_cli_summarizes_nonempty_legacy_xls_elapsed_values() -> Result<()> {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("timeline")
        .join("wide_summary_elapsed_duration.xls");

    for command in ["wide-summary", "wide-summary-row-suffix"] {
        let root = TestDirectory::new()?;
        let config = create_timeline_config(root.path())?;
        let source = root.path().join("sample_2024-06-30.xls");
        fs::copy(&fixture, &source)?;
        let source_before = fs::read(&source)?;

        let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
            .arg(command)
            .arg("--config")
            .arg(&config)
            .arg("--log-view")
            .arg("developer")
            .arg(&source)
            .output()
            .context("run wide-summary CLI with nonempty BIFF XLS input")?;
        assert!(
            output.status.success(),
            "{command} did not summarize nonempty BIFF XLS input: {}",
            String::from_utf8_lossy(&output.stderr)
        );

        let report: Value = serde_json::from_slice(&output.stdout)?;
        let summary_path = report["output"]
            .as_str()
            .map(PathBuf::from)
            .context("nonempty legacy XLS summary output path missing")?;
        let dump = read_workbook(&summary_path)?;
        let summary = dump
            .sheets
            .iter()
            .find(|sheet| sheet.name == "WideRule")
            .context("nonempty legacy XLS summary sheet missing")?;
        assert_eq!(summary.values["G2"]["type"], "datetime_serial");
        let elapsed = summary.values["G2"]["value"]
            .as_f64()
            .context("BIFF elapsed date serial missing")?;
        assert!((elapsed - 1.123_456_789).abs() < 1e-15);
        assert_eq!(summary.values["G3"]["type"], "string");
        assert_eq!(summary.values["G3"]["value"], "12:00:00");

        let run_log_path = report["run_log_path"]
            .as_str()
            .map(PathBuf::from)
            .context("nonempty legacy XLS run history missing")?;
        let history: Value = serde_json::from_slice(&fs::read(run_log_path)?)?;
        assert_eq!(history["completed_inputs"], 1);
        assert_eq!(history["failed_inputs"], 0);
        assert_eq!(fs::read(&source)?, source_before);
    }
    Ok(())
}

#[test]
fn wide_summary_cli_discovers_default_sources_for_both_row_modes() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config(root.path())?;
    let input_dir = root.path().join("app").join("input");
    let nested_dir = input_dir.join("nested");
    fs::create_dir_all(&nested_dir)?;

    create_wide_source(
        &input_dir.join("sample root source.xlsx"),
        "默认发现根目录",
        21.0,
    )?;
    create_wide_source(
        &nested_dir.join("sample nested source.XLSX"),
        "默认发现子目录",
        31.0,
    )?;
    create_wide_source(
        &input_dir.join("~$sample locked source.xlsx"),
        "锁文件必须跳过",
        41.0,
    )?;
    fs::write(
        input_dir.join("sample unsupported source.xlsb"),
        b"not xlsb",
    )?;

    for command in ["wide-summary", "wide-summary-row-suffix"] {
        let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
            .arg(command)
            .arg("--config")
            .arg(&config)
            .arg("--log-view")
            .arg("developer")
            .output()
            .with_context(|| format!("run {command} with configured source discovery"))?;
        assert!(
            output.status.success(),
            "{command} default source discovery failed: {}",
            String::from_utf8_lossy(&output.stderr)
        );

        let report: Value = serde_json::from_slice(&output.stdout)?;
        let summary_path = report["output"]
            .as_str()
            .map(PathBuf::from)
            .with_context(|| format!("{command} output path missing"))?;
        let summary = read_workbook(&summary_path)?;
        assert_eq!(summary.sheets.len(), 1);
        assert!(summary.sheets[0].values.len() > 10);

        let history: Value = serde_json::from_slice(&fs::read(
            report["run_log_path"]
                .as_str()
                .with_context(|| format!("{command} run history path missing"))?,
        )?)?;
        assert_eq!(history["total_inputs"], 2);
        assert_eq!(history["completed_inputs"], 2);
        assert_eq!(history["failed_inputs"], 0);
        let events = history["events"]
            .as_array()
            .context("wide-summary events missing")?;
        assert!(events.iter().any(|event| {
            event["phase"] == "source_discovery"
                && event["status"] == "completed"
                && event["progress"]["total_inputs"] == 2
        }));
        assert!(events
            .iter()
            .enumerate()
            .all(|(index, event)| { event["sequence"].as_u64() == Some(index as u64 + 1) }));
    }
    Ok(())
}

#[test]
fn wide_summary_cli_resolves_relative_configured_directories_for_discovery() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config_with_relative_dirs(root.path())?;
    let input_dir = root.path().join("app").join("input");
    let nested_dir = input_dir.join("nested");
    fs::create_dir_all(&nested_dir)?;
    create_wide_source(
        &input_dir.join("sample relative root source.xlsx"),
        "相对目录根",
        21.0,
    )?;
    create_wide_source(
        &nested_dir.join("sample relative nested source.xlsx"),
        "相对目录子",
        31.0,
    )?;
    create_wide_source(
        &input_dir.join("~$sample relative locked source.xlsx"),
        "锁文件必须跳过",
        41.0,
    )?;
    fs::write(
        input_dir.join("sample relative unsupported source.xlsb"),
        b"not xlsb",
    )?;

    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("wide-summary")
        .arg("--config")
        .arg(&config)
        .arg("--log-view")
        .arg("developer")
        .output()
        .context("run wide-summary with relative configured directories")?;
    assert!(
        output.status.success(),
        "relative directory wide-summary failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let report: Value = serde_json::from_slice(&output.stdout)?;
    let summary_path = report["output"]
        .as_str()
        .map(PathBuf::from)
        .with_context(|| "relative wide-summary output path missing")?;
    let output_dir = root.path().join("app").join("output");
    assert!(
        summary_path.starts_with(&output_dir),
        "summary must be written under the config-relative output directory: {summary_path:?}"
    );
    let summary = read_workbook(&summary_path)?;
    assert_eq!(summary.sheets.len(), 1);
    assert!(summary.sheets[0].values.len() > 10);

    let run_log_path = report["run_log_path"]
        .as_str()
        .map(PathBuf::from)
        .with_context(|| "relative wide-summary run log path missing")?;
    let log_dir = root.path().join("app").join("logs");
    assert!(
        run_log_path.starts_with(&log_dir),
        "run history must be written under the config-relative log directory: {run_log_path:?}"
    );
    let history: Value = serde_json::from_slice(&fs::read(&run_log_path)?)?;
    assert_eq!(history["total_inputs"], 2);
    assert_eq!(history["completed_inputs"], 2);
    assert_eq!(history["failed_inputs"], 0);

    Ok(())
}

#[test]
fn timeline_extract_cli_discovers_default_sources_recursively_and_skips_lock_files() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config(root.path())?;
    let input_dir = root.path().join("app").join("input");
    let nested_dir = input_dir.join("nested");
    fs::create_dir_all(&nested_dir)?;

    let root_source = input_dir.join("sample root source.xlsx");
    let nested_source = nested_dir.join("sample nested source.XLSX");
    let lock_source = input_dir.join("~$sample locked source.xlsx");
    let unsupported_source = input_dir.join("sample unsupported source.xlsb");
    create_wide_source(&root_source, "自动发现根目录", 21.0)?;
    create_wide_source(&nested_source, "自动发现子目录", 31.0)?;
    create_wide_source(&lock_source, "锁文件必须跳过", 41.0)?;
    fs::write(&unsupported_source, b"unsupported workbook format")?;

    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("timeline-extract")
        .arg("--config")
        .arg(&config)
        .arg("--log-view")
        .arg("developer")
        .output()
        .context("run timeline-extract with configured input discovery")?;
    assert!(
        output.status.success(),
        "timeline-extract default discovery failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let report: Value = serde_json::from_slice(&output.stdout)?;
    assert_eq!(report["total_files"], 2);
    assert_eq!(report["processed_files"], 2);
    assert_eq!(report["failed_files"], 0);
    assert!(report["rows"].as_u64().is_some_and(|rows| rows > 0));
    let output_path = report["path"]
        .as_str()
        .map(PathBuf::from)
        .context("timeline extraction output path missing")?;
    let workbook = read_workbook(&output_path)?;
    let output_values = &workbook.sheets[0].values;
    for expected_source in ["sample root source.xlsx", "sample nested source.XLSX"] {
        assert!(
            output_values
                .values()
                .any(|cell| { cell.get("value").and_then(Value::as_str) == Some(expected_source) }),
            "default discovery output missing source {expected_source}"
        );
    }
    assert!(!output_values.values().any(|cell| {
        cell.get("value").and_then(Value::as_str) == Some("~$sample locked source.xlsx")
    }));

    let history: Value = serde_json::from_slice(&fs::read(
        report["run_log_path"]
            .as_str()
            .context("timeline extraction history path missing")?,
    )?)?;
    assert_eq!(history["progress"]["total_inputs"], 2);
    assert!(history["events"].as_array().is_some_and(|events| {
        events.iter().any(|event| {
            event["phase"] == "source_discovery"
                && event["status"] == "completed"
                && event["progress"]["total_inputs"] == 2
        })
    }));
    Ok(())
}

#[test]
fn timeline_extract_cli_with_empty_default_input_directory_skips_output() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config(root.path())?;
    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("timeline-extract")
        .arg("--config")
        .arg(&config)
        .arg("--log-view")
        .arg("developer")
        .output()
        .context("run timeline-extract with no configured source files")?;
    assert!(
        output.status.success(),
        "timeline-extract empty discovery failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let report: Value = serde_json::from_slice(&output.stdout)?;
    assert_eq!(report["total_files"], 0);
    assert_eq!(report["rows"], 0);
    assert!(report["path"].is_null());
    assert!(!root.path().join("app").join("output").exists());
    Ok(())
}

#[test]
fn header_path_compare_menu_supports_user_developer_and_off_views() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config(root.path())?;
    let template = root.path().join("compare template.xlsx");
    let source = root.path().join("sample compare source.xlsx");
    let broken = root.path().join("broken compare source.xlsx");
    create_wide_source(&template, "模板甲", 10.0)?;
    create_wide_source(&source, "来源甲", 12.0)?;
    fs::write(&broken, b"not an xlsx workbook")?;
    let config_before = fs::read(&config)?;
    let source_before = fs::read(&source)?;
    let log_dir = root.path().join("app").join("logs");

    let user = run_menu_with_view(
        &config,
        "user",
        &format!("1\n4\n{}\n{}\n0\n", template.display(), source.display()),
    )?;
    assert!(
        user.status.success(),
        "user-view compare failed: {}",
        String::from_utf8_lossy(&user.stderr)
    );
    let user_stderr = String::from_utf8_lossy(&user.stderr);
    assert!(user_stderr.contains("正在读取表头比对配置"));
    assert!(user_stderr.contains("表头比对模板读取完成"));
    assert!(user_stderr.contains("[人工确认]"));
    assert!(user_stderr.contains("路径标准化映射"));
    assert!(user_stderr.contains("已处理 1/1 个源工作簿"));
    assert!(user_stderr.contains("表头路径比对结束"));
    assert!(user_stderr.contains("阶段用时"));
    assert!(!user_stderr.contains(&source.display().to_string()));
    assert!(!user_stderr.contains("来源甲"));
    let user_report = latest_timeline_compare_report(&user)?;
    assert!(user_report.get("output").is_some());
    assert!(user_report.get("run_history_path").is_none());
    assert_eq!(fs::read(&config)?, config_before);
    assert_eq!(fs::read(&source)?, source_before);
    assert!(!log_dir.exists() || run_history_paths(&log_dir)?.is_empty());

    let developer_sources = format!(
        "{};{};{}",
        source.display(),
        broken.display(),
        template.display()
    );
    let developer = run_menu_with_view(
        &config,
        "developer",
        &format!("1\n4\n{}\n{}\n0\n", template.display(), developer_sources),
    )?;
    assert!(
        developer.status.success(),
        "developer-view compare failed: {}",
        String::from_utf8_lossy(&developer.stderr)
    );
    let developer_report = latest_timeline_compare_report(&developer)?;
    let history_path = PathBuf::from(
        developer_report["run_history_path"]
            .as_str()
            .context("developer run-history path missing")?,
    );
    let history_bytes = fs::read(&history_path)?;
    let history_text = String::from_utf8_lossy(&history_bytes);
    assert!(!history_text.contains(&source.display().to_string()));
    assert!(!history_text.contains(&broken.display().to_string()));
    assert!(!history_text.contains("来源甲"));
    assert!(!history_text.contains("not an xlsx workbook"));
    let history: Value = serde_json::from_slice(&history_bytes)?;
    assert_eq!(history["feature_id"], "PT-1-1");
    assert_eq!(history["status"], "partial");
    assert_eq!(history["progress"]["total_inputs"], 3);
    assert_eq!(history["progress"]["completed_inputs"], 1);
    assert_eq!(history["progress"]["failed_inputs"], 1);
    assert_eq!(history["progress"]["skipped_inputs"], 1);
    for phase in [
        "configuration_read",
        "template_read",
        "source_read",
        "source_compare",
        "report_write",
        "task",
    ] {
        assert!(
            history["events"]
                .as_array()
                .is_some_and(|events| events.iter().any(|event| event["phase"] == phase)),
            "missing {phase} event"
        );
    }
    assert!(history["actual_engines"].as_array().is_some_and(|engines| {
        engines.iter().any(|engine| engine == "calamine")
            && engines.iter().any(|engine| engine == "rust_xlsxwriter")
    }));
    let developer_events = String::from_utf8_lossy(&developer.stderr)
        .lines()
        .filter_map(|line| serde_json::from_str::<Value>(line).ok())
        .collect::<Vec<_>>();
    assert!(developer_events.len() >= 10);
    assert!(developer_events.iter().all(|event| {
        event["stage_elapsed_ms"].is_number() && event["total_elapsed_ms"].is_number()
    }));

    let history_count_before_off = run_history_paths(&log_dir)?.len();
    let off = run_menu_with_view(
        &config,
        "off",
        &format!("1\n4\n{}\n{}\n0\n", template.display(), source.display()),
    )?;
    assert!(
        off.status.success(),
        "off-view compare failed: {}",
        String::from_utf8_lossy(&off.stderr)
    );
    assert!(String::from_utf8_lossy(&off.stderr).is_empty());
    assert!(latest_timeline_compare_report(&off)?
        .get("run_history_path")
        .is_none());
    assert_eq!(run_history_paths(&log_dir)?.len(), history_count_before_off);
    Ok(())
}

#[test]
fn header_path_compare_formats_date_headers_as_python_datetime_text() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config(root.path())?;
    let template = root.path().join("sample template.xlsx");
    let source = root.path().join("sample source.xlsx");
    create_date_header_source(&template, 2024, 1, 31)?;
    create_date_header_source(&source, 2024, 2, 29)?;

    let input = format!("1\n4\n{}\n{}\n0\n", template.display(), source.display());
    let output = run_menu_with_view(&config, "off", &input)?;
    assert!(
        output.status.success(),
        "date-header compare failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report = latest_timeline_compare_report(&output)?;
    let result_path = PathBuf::from(
        report["output"]
            .as_str()
            .context("date-header comparison output path missing")?,
    );
    let workbook = read_workbook(result_path)?;
    let differences = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "表头比对结果")
        .context("date-header difference sheet missing")?;
    assert_eq!(
        differences.values["H2"]["value"],
        "存款_2024-02-29 00:00:00"
    );
    assert_eq!(
        differences.values["I2"]["value"],
        "存款_2024-01-31 00:00:00"
    );
    assert_eq!(
        differences.values["H3"]["value"],
        "存款_2024-02-29 00:00:00"
    );
    assert_eq!(
        differences.values["I4"]["value"],
        "存款_2024-01-31 00:00:00"
    );
    Ok(())
}

#[test]
fn header_path_compare_displays_merged_multilevel_paths_like_wide_summary() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config_with_header_rows(root.path(), "2,3,4")?;
    let template = root.path().join("multi-level compare template.xlsx");
    let source = root.path().join("sample multi-level compare source.xlsx");
    create_multi_level_header_compare_source(&template, "外币")?;
    create_multi_level_header_compare_source(&source, "本币")?;
    let template_before = fs::read(&template)?;
    let source_before = fs::read(&source)?;

    let input = format!("1\n4\n{}\n{}\n0\n", template.display(), source.display());
    let output = run_menu_with_view(&config, "off", &input)?;
    assert!(
        output.status.success(),
        "multilevel header compare failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let report = latest_timeline_compare_report(&output)?;
    let result_path = PathBuf::from(
        report["output"]
            .as_str()
            .context("multilevel comparison output path missing")?,
    );
    let workbook = read_workbook(result_path)?;
    let differences = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "表头比对结果")
        .context("multilevel difference sheet missing")?;
    assert_eq!(differences.values["H2"]["value"], "资产_本币_余额");
    assert_eq!(differences.values["I2"]["value"], "资产_外币_余额");
    assert_eq!(differences.values["H3"]["value"], "资产_本币_余额");
    assert_eq!(differences.values["I4"]["value"], "资产_外币_余额");
    assert!(differences
        .values
        .values()
        .any(|value| value["value"] == "资产_本币_余额"));
    assert!(differences
        .values
        .values()
        .any(|value| value["value"] == "资产_外币_余额"));
    assert_eq!(fs::read(&template)?, template_before);
    assert_eq!(fs::read(&source)?, source_before);
    Ok(())
}

#[test]
fn header_path_compare_reports_unmatched_rules_and_worksheets_in_logs() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config(root.path())?;
    let template = root.path().join("template.xlsx");
    let unmatched_workbook = root.path().join("other-workbook.xlsx");
    let unmatched_worksheet = root.path().join("sample-no-report-sheet.xlsx");
    let matched = root.path().join("sample-matched.xlsx");
    create_wide_source(&template, "机构", 10.0)?;
    create_wide_source(&unmatched_workbook, "机构", 10.0)?;
    create_header_compare_source_with_sheet(&unmatched_worksheet, "Other")?;
    create_wide_source(&matched, "机构", 10.0)?;

    let selected_sources = format!(
        "{};{};{}",
        unmatched_workbook.display(),
        unmatched_worksheet.display(),
        matched.display()
    );
    let user = run_menu_with_view(
        &config,
        "user",
        &format!("1\n4\n{}\n{}\n0\n", template.display(), selected_sources),
    )?;
    assert!(
        user.status.success(),
        "user-view compare failed: {}",
        String::from_utf8_lossy(&user.stderr)
    );
    let user_log = String::from_utf8_lossy(&user.stderr);
    assert!(user_log.contains("未匹配规则：源工作簿 other-workbook.xlsx"));
    assert!(user_log.contains("工作簿关键字未命中任何启用规则"));
    assert!(user_log.contains("规则「WideRule」匹配了工作簿，但工作表关键字「Report」未命中"));
    assert!(user_log.contains("工作簿 sample-matched.xlsx 命中规则「WideRule」"));
    assert!(user_log.contains("规则「WideRule」已匹配工作表：源「Report」→ 模板「Report」"));
    assert!(user_log.contains("未匹配 2 个"));

    let report = latest_timeline_compare_report(&user)?;
    let result_path = PathBuf::from(
        report["output"]
            .as_str()
            .context("unmatched-rule comparison output path missing")?,
    );
    let workbook = read_workbook(result_path)?;
    let summary = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "汇总")
        .context("comparison summary sheet missing")?;
    assert_eq!(summary.values["B2"]["value"], "未匹配规则");
    assert!(summary.values["E2"]["value"]
        .as_str()
        .unwrap_or_default()
        .contains("工作簿关键字未命中"));
    assert_eq!(summary.values["B3"]["value"], "未匹配工作表");
    assert!(summary.values["E3"]["value"]
        .as_str()
        .unwrap_or_default()
        .contains("工作表关键字「Report」未命中"));
    assert_eq!(summary.values["B4"]["value"], "通过");
    assert_eq!(summary.values["C4"]["value"], "0");
    assert_eq!(summary.values["D4"]["value"], "0");

    let developer = run_menu_with_view(
        &config,
        "developer",
        &format!(
            "1\n4\n{}\n{}\n0\n",
            template.display(),
            unmatched_workbook.display()
        ),
    )?;
    assert!(developer.status.success());
    let developer_report = latest_timeline_compare_report(&developer)?;
    let history_path = PathBuf::from(
        developer_report["run_history_path"]
            .as_str()
            .context("developer run-history path missing")?,
    );
    let history: Value = serde_json::from_slice(&fs::read(history_path)?)?;
    assert_eq!(history["status"], "partial");
    assert_eq!(history["progress"]["matched_rules"], 0);
    assert_eq!(history["progress"]["unmatched_inputs"], 1);
    assert!(history["events"].as_array().is_some_and(|events| {
        events.iter().any(|event| {
            event["phase"] == "rule_match"
                && event["user_message"]
                    .as_str()
                    .is_some_and(|message| message.contains("工作簿关键字未命中"))
        })
    }));
    Ok(())
}

#[test]
fn header_path_compare_menu_records_terminal_failure_summary() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config(root.path())?;
    let template = root.path().join("broken template.xlsx");
    let source = root.path().join("valid source.xlsx");
    fs::write(&template, b"not an xlsx workbook")?;
    create_wide_source(&source, "来源甲", 12.0)?;

    let user = run_menu_with_view(
        &config,
        "user",
        &format!("1\n4\n{}\n{}\n0\n", template.display(), source.display()),
    )?;
    assert!(!user.status.success());
    let user_stderr = String::from_utf8_lossy(&user.stderr);
    assert!(user_stderr.contains("表头比对模板读取失败"));
    assert!(user_stderr.contains("表头路径比对任务失败"));
    assert!(!user_stderr.contains(&template.display().to_string()));
    assert!(!user_stderr.contains("not an xlsx workbook"));
    assert!(!user_stderr.contains("calamine"));

    let log_dir = root.path().join("app").join("logs");
    assert!(!log_dir.exists() || run_history_paths(&log_dir)?.is_empty());

    let developer = run_menu_with_view(
        &config,
        "developer",
        &format!("1\n4\n{}\n{}\n0\n", template.display(), source.display()),
    )?;
    assert!(!developer.status.success());
    let history = read_single_run_history(&log_dir, "PT-1-1")?;
    assert_eq!(history["status"], "failed");
    assert_eq!(history["error_type"], "template_read_error");
    let history_text = serde_json::to_string(&history)?;
    assert!(!history_text.contains(&template.display().to_string()));
    assert!(!history_text.contains("not an xlsx workbook"));
    assert!(history["events"].as_array().is_some_and(|events| {
        events
            .iter()
            .any(|event| event["phase"] == "template_read" && event["status"] == "failed")
            && events.iter().any(|event| {
                event["phase"] == "task"
                    && event["status"] == "failed"
                    && event["error_type"] == "template_read_error"
            })
    }));
    Ok(())
}

#[test]
fn interactive_wide_summary_writes_the_configured_intermediate_for_later_dedup() -> Result<()> {
    let root = TestDirectory::new()?;
    let target = root.path().join("configured target.xlsx");
    let config = create_timeline_config_with_target(root.path(), Some(&target))?;
    let config_before = fs::read(&config)?;
    let source = root.path().join("sample wide source.xlsx");
    create_wide_source(&source, "交互汇总甲", 23.5)?;
    let source_before = fs::read(&source)?;

    let completed = run_menu(&config, &format!("1\n1\n1\n{}\n0\n", source.display()))?;
    assert!(
        completed.status.success(),
        "intermediate wide summary failed: {}",
        String::from_utf8_lossy(&completed.stderr)
    );
    let stderr = String::from_utf8_lossy(&completed.stderr);
    assert!(stderr.contains("中间工作簿 configured target.xlsx 已写入 1 行"));
    assert!(target.is_file(), "configured intermediate was not created");

    let report = latest_menu_json_report(&completed)?;
    assert_eq!(report["target_write_status"], "succeeded");
    assert_eq!(report["target_writes"][0]["status"], "succeeded");
    assert!(report["total_elapsed_ms"].is_u64());
    assert_eq!(report["target_writes"][0]["appended_rows"], 1);
    assert_eq!(
        report["target_writes"][0]["target_name"],
        "configured target.xlsx"
    );
    let summary_path = PathBuf::from(
        report["output"]
            .as_str()
            .context("intermediate workbook path missing")?,
    );
    assert!(
        summary_path.is_file(),
        "intermediate workbook was not created"
    );
    let summary = read_workbook(&summary_path)?;
    assert!(summary.sheets.iter().any(|sheet| {
        sheet
            .values
            .values()
            .any(|cell| cell.get("value").and_then(Value::as_str) == Some("交互汇总甲"))
    }));
    let intermediate = read_workbook(&target)?;
    assert!(intermediate.sheets.iter().any(|sheet| {
        sheet.name == "Append"
            && sheet
                .values
                .values()
                .any(|cell| cell.get("value").and_then(Value::as_str) == Some("交互汇总甲"))
            && sheet
                .values
                .values()
                .any(|cell| cell.get("value").and_then(Value::as_f64) == Some(23.5))
    }));
    assert_eq!(fs::read(&config)?, config_before);
    assert_eq!(fs::read(&source)?, source_before);

    let history = read_single_run_history(&root.path().join("app").join("logs"), "1-1")?;
    assert_eq!(history["feature_id"], "1-1");
    assert_eq!(history["target_write_status"], "succeeded");
    assert!(history["events"].as_array().is_some_and(|events| {
        events
            .iter()
            .any(|event| event["phase"] == "target_append_plan")
            && events
                .iter()
                .any(|event| event["phase"] == "target_append_write")
    }));
    Ok(())
}

#[test]
fn interactive_wide_summary_does_not_modify_existing_intermediate_without_capability() -> Result<()>
{
    let root = TestDirectory::new()?;
    let target = root.path().join("configured target.xlsx");
    let config = create_timeline_config_with_target(root.path(), Some(&target))?;
    let first_source = root.path().join("first sample wide source.xlsx");
    let next_source = root.path().join("next sample wide source.xlsx");
    create_wide_source(&first_source, "中间表已有记录", 10.0)?;
    create_wide_source(&next_source, "中间表新增记录", 99.0)?;

    let first = run_menu(
        &config,
        &format!("1\n1\n1\n{}\n0\n", first_source.display()),
    )?;
    assert!(
        first.status.success(),
        "initial intermediate creation failed"
    );
    let target_before = fs::read(&target)?;

    let second = run_menu(&config, &format!("1\n1\n1\n{}\n0\n", next_source.display()))?;
    assert!(
        second.status.success(),
        "blocked target append should preserve summary output"
    );
    let report = latest_menu_json_report(&second)?;
    assert_eq!(report["target_write_status"], "partial");
    assert_eq!(report["target_writes"][0]["status"], "blocked");
    assert_eq!(
        report["target_writes"][0]["reason"],
        "existing_intermediate_requires_excel_modifier_capability"
    );
    assert_eq!(fs::read(&target)?, target_before);
    let stderr = String::from_utf8_lossy(&second.stderr);
    assert!(stderr.contains("本次未提供匹配的 ExcelModifier 能力"));
    Ok(())
}

#[test]
fn interactive_wide_summary_intermediate_preserves_negative_date_serial() -> Result<()> {
    let root = TestDirectory::new()?;
    let target = root.path().join("configured target.xlsx");
    let config = create_timeline_config_with_target(root.path(), Some(&target))?;
    let config_before = fs::read(&config)?;
    let source = root.path().join("sample negative date source.xlsx");
    create_wide_source_with_negative_date(&source, "负序列分行")?;
    let source_before = fs::read(&source)?;

    let completed = run_menu(&config, &format!("1\n1\n1\n{}\n0\n", source.display()))?;
    assert!(
        completed.status.success(),
        "wide-summary intermediate failed: {}",
        String::from_utf8_lossy(&completed.stderr)
    );
    assert!(target.is_file(), "configured intermediate was not created");
    assert_eq!(fs::read(&source)?, source_before);
    assert_eq!(fs::read(&config)?, config_before);

    let report = latest_menu_json_report(&completed)?;
    let summary_path = PathBuf::from(
        report["output"]
            .as_str()
            .context("intermediate workbook path missing")?,
    );
    let summary = read_workbook(&summary_path)?;
    let negative_serial_exists = summary.sheets.iter().any(|sheet| {
        sheet.values.values().any(|cell| {
            cell.get("type").and_then(Value::as_str) == Some("datetime_serial")
                && cell
                    .get("value")
                    .and_then(Value::as_f64)
                    .is_some_and(|serial| (serial + 0.5).abs() < 1e-9)
        })
    });
    assert!(
        negative_serial_exists,
        "negative cached Excel serial was not preserved in the intermediate workbook"
    );
    let intermediate = read_workbook(&target)?;
    assert!(intermediate.sheets.iter().any(|sheet| {
        sheet.values.values().any(|cell| {
            cell.get("type").and_then(Value::as_str) == Some("datetime_serial")
                && cell
                    .get("value")
                    .and_then(Value::as_f64)
                    .is_some_and(|serial| (serial + 0.5).abs() < 1e-9)
        })
    }));
    assert_eq!(
        read_single_run_history(&root.path().join("app").join("logs"), "1-1")?["status"],
        "succeeded"
    );
    Ok(())
}
#[test]
fn wide_summary_menu_logging_views_record_progress_and_respect_off() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config(root.path())?;
    let source = root.path().join("sample wide source.xlsx");
    create_wide_source(&source, "宽表甲", 17.5)?;

    let user = run_menu_with_view(
        &config,
        "user",
        &format!("1\n1\n1\n{}\n0\n", source.display()),
    )?;
    assert!(
        user.status.success(),
        "user-view wide summary failed: {}",
        String::from_utf8_lossy(&user.stderr)
    );
    let user_stderr = String::from_utf8_lossy(&user.stderr);
    assert!(user_stderr.contains("正在读取工作簿 sample wide source.xlsx（第 1/1 个）"));
    assert!(!user_stderr.contains("\"phase\""));

    let log_dir = root.path().join("app").join("logs");
    let user_history = read_single_run_history(&log_dir, "1-1")?;
    assert_eq!(user_history["status"], "succeeded");
    assert_eq!(user_history["target_write_status"], "not_configured");
    assert!(user_history["total_elapsed_ms"].is_u64());
    assert_eq!(user_history["completed_inputs"], 1);
    assert_eq!(user_history["failed_inputs"], 0);
    assert!(user_history["actual_engines"]
        .as_array()
        .is_some_and(|engines| engines.iter().any(|engine| engine == "calamine")
            && engines.iter().any(|engine| engine == "rust_xlsxwriter")));
    assert_history_has_phase(&user_history, "source_read", "completed", "calamine")?;
    assert_history_has_phase(&user_history, "rule_aggregation", "completed", "")?;
    assert_history_has_phase(
        &user_history,
        "summary_write",
        "completed",
        "rust_xlsxwriter",
    )?;
    let events = user_history["events"]
        .as_array()
        .context("run history events must be an array")?;
    assert!(events
        .iter()
        .any(|event| event["phase"] == "target_append_plan" && event["status"] == "skipped"));
    assert!(events.iter().any(|event| {
        event["phase"] == "source_read"
            && event["workbook_name"] == "sample wide source.xlsx"
            && event["stage_elapsed_ms"].is_u64()
    }));
    assert!(events.iter().any(|event| {
        event["phase"] == "worksheet_aggregation"
            && event["worksheet_name"] == "Report"
            && event["stage_elapsed_ms"].is_u64()
    }));
    let history_text = serde_json::to_string(&user_history)?;
    assert!(!history_text.contains(&source.display().to_string()));
    assert!(!history_text.contains("宽表甲"));

    let developer_root = TestDirectory::new()?;
    let developer_config = create_timeline_config(developer_root.path())?;
    let developer_source = developer_root.path().join("sample wide source.xlsx");
    create_wide_source(&developer_source, "宽表甲", 17.5)?;
    let developer = run_menu_with_view(
        &developer_config,
        "developer",
        &format!("1\n1\n1\n{}\n0\n", developer_source.display()),
    )?;
    assert!(
        developer.status.success(),
        "developer-view wide summary failed: {}",
        String::from_utf8_lossy(&developer.stderr)
    );
    let developer_events = String::from_utf8_lossy(&developer.stderr)
        .lines()
        .filter_map(|line| serde_json::from_str::<Value>(line).ok())
        .collect::<Vec<_>>();
    assert!(developer_events.iter().any(|event| {
        event["phase"] == "summary_write"
            && event["status"] == "completed"
            && event["actual_engine"] == "rust_xlsxwriter"
            && event["components"]
                .as_array()
                .is_some_and(|components| components.iter().any(|value| value == "Rust"))
    }));
    assert!(developer_events.iter().all(|event| {
        let text = event.to_string();
        !text.contains(&developer_source.display().to_string()) && !text.contains("宽表甲")
    }));
    assert!(developer_events.iter().any(|event| {
        event["workbook_name"] == "sample wide source.xlsx" && event["worksheet_name"] == "Report"
    }));
    assert!(developer_events.windows(2).all(|pair| {
        pair[0]["total_elapsed_ms"]
            .as_u64()
            .zip(pair[1]["total_elapsed_ms"].as_u64())
            .is_some_and(|(previous, next)| previous <= next)
    }));
    let developer_log_dir = developer_root.path().join("app").join("logs");
    let developer_history = read_single_run_history(&developer_log_dir, "1-1")?;
    assert_eq!(developer_history["status"], "succeeded");
    assert_history_has_phase(&developer_history, "source_read", "completed", "calamine")?;

    let history_count_before_off = run_history_paths(&log_dir)?.len();
    let off = run_menu_with_view(
        &config,
        "off",
        &format!("1\n1\n1\n{}\n0\n", source.display()),
    )?;
    assert!(
        off.status.success(),
        "off-view wide summary failed: {}",
        String::from_utf8_lossy(&off.stderr)
    );
    let off_stdout = String::from_utf8_lossy(&off.stdout);
    let off_stderr = String::from_utf8_lossy(&off.stderr);
    assert!(!off_stdout.contains("run_log_path"));
    assert!(!off_stdout.contains("run_id"));
    assert!(!off_stderr.contains("正在读取第"));
    assert_eq!(run_history_paths(&log_dir)?.len(), history_count_before_off);
    Ok(())
}

#[test]
fn wide_summary_menu_failure_history_identifies_read_stage_without_source_data() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config(root.path())?;
    let source = root.path().join("broken source.xlsx");
    fs::write(&source, b"not an xlsx workbook")?;

    let failure = run_menu_with_view(
        &config,
        "user",
        &format!("1\n1\n1\n{}\n0\n", source.display()),
    )?;
    assert!(!failure.status.success());
    let user_stderr = String::from_utf8_lossy(&failure.stderr);
    assert!(user_stderr.contains("读取失败"));
    assert!(user_stderr.contains("宽表汇总失败"));
    assert!(!user_stderr.contains(&source.display().to_string()));
    assert!(!user_stderr.contains("calamine"));

    let log_dir = root.path().join("app").join("logs");
    let history = read_single_run_history(&log_dir, "1-1")?;
    assert_eq!(history["status"], "failed");
    assert_eq!(history["failure"]["stage"], "source_read");
    assert_eq!(history["failure"]["input_index"], 1);
    assert_eq!(history["failure"]["error_type"], "workbook_read_error");
    assert_history_has_phase(&history, "source_read", "failed", "calamine")?;
    let history_text = serde_json::to_string(&history)?;
    assert!(!history_text.contains(&source.display().to_string()));
    assert!(!history_text.contains("not an xlsx workbook"));
    Ok(())
}

#[test]
fn wide_summary_continue_policy_explains_read_failure_without_leaking_input() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config_with_error_policy(root.path(), "continue")?;
    let source = root.path().join("broken source.xlsx");
    fs::write(&source, b"not an xlsx workbook")?;
    let source_before = fs::read(&source)?;

    let completed = run_menu_with_view(
        &config,
        "user",
        &format!("1\n1\n1\n{}\n0\n", source.display()),
    )?;
    assert!(
        completed.status.success(),
        "continue policy should return to the menu: {}",
        String::from_utf8_lossy(&completed.stderr)
    );
    let user_output = String::from_utf8_lossy(&completed.stderr);
    assert!(user_output.contains("broken source.xlsx") && user_output.contains("读取失败"));
    assert!(user_output.contains("请确认文件存在、可读取且为受支持的工作簿格式"));
    assert!(user_output.contains("已按继续策略跳过"));
    assert!(user_output.contains("没有匹配的宽表数据，未生成汇总工作簿"));
    assert!(!user_output.contains(&source.display().to_string()));
    assert_eq!(fs::read(&source)?, source_before);
    assert!(!root.path().join("app/output").exists());

    let log_dir = root.path().join("app").join("logs");
    let history = read_single_run_history(&log_dir, "1-1")?;
    assert_eq!(history["status"], "partial");
    assert_eq!(history["failed_inputs"], 1);
    assert!(history["output"].is_null());
    let history_text = serde_json::to_string(&history)?;
    assert!(history_text.contains("请确认文件存在、可读取且为受支持的工作簿格式"));
    assert!(!history_text.contains(&source.display().to_string()));
    assert!(!history_text.contains("not an xlsx workbook"));
    Ok(())
}

fn assert_history_has_phase(
    history: &Value,
    phase: &str,
    status: &str,
    actual_engine: &str,
) -> Result<()> {
    let events = history["events"]
        .as_array()
        .context("run history events must be an array")?;
    assert!(
        events.iter().any(|event| {
            event["phase"] == phase
                && event["status"] == status
                && (actual_engine.is_empty() || event["actual_engine"] == actual_engine)
        }),
        "missing {phase}/{status} history event"
    );
    Ok(())
}

fn run_history_paths(log_dir: &Path) -> Result<Vec<PathBuf>> {
    let paths = fs::read_dir(log_dir)?
        .map(|entry| entry.map(|entry| entry.path()))
        .collect::<std::io::Result<Vec<_>>>()?;
    Ok(paths
        .into_iter()
        .filter(|path| {
            path.extension()
                .is_some_and(|extension| extension == "json")
        })
        .collect())
}

fn read_single_run_history(log_dir: &Path, feature_id: &str) -> Result<Value> {
    let prefix = format!("{feature_id}_");
    let paths = run_history_paths(log_dir)?
        .into_iter()
        .filter(|path| {
            path.file_name()
                .is_some_and(|name| name.to_string_lossy().starts_with(&prefix))
        })
        .collect::<Vec<_>>();
    anyhow::ensure!(
        paths.len() == 1,
        "expected one {feature_id} history file, found {}",
        paths.len()
    );
    serde_json::from_slice(&fs::read(&paths[0])?).context("parse wide summary run history")
}

fn create_timeline_config(root: &Path) -> Result<PathBuf> {
    create_timeline_config_inner(root, None, false)
}

fn create_timeline_config_with_header_rows(root: &Path, header_rows: &str) -> Result<PathBuf> {
    create_timeline_config_inner_with_header_rows(root, None, false, "fail_fast", header_rows)
}

fn create_timeline_config_with_relative_dirs(root: &Path) -> Result<PathBuf> {
    create_timeline_config_inner(root, None, true)
}

fn create_timeline_config_with_target(root: &Path, target: Option<&Path>) -> Result<PathBuf> {
    create_timeline_config_inner(root, target, false)
}

fn create_timeline_config_inner(
    root: &Path,
    target: Option<&Path>,
    relative_dirs: bool,
) -> Result<PathBuf> {
    create_timeline_config_inner_with_error_policy(root, target, relative_dirs, "fail_fast")
}

fn create_timeline_config_with_error_policy(root: &Path, error_policy: &str) -> Result<PathBuf> {
    create_timeline_config_inner_with_error_policy(root, None, false, error_policy)
}

fn create_timeline_config_inner_with_error_policy(
    root: &Path,
    target: Option<&Path>,
    relative_dirs: bool,
    error_policy: &str,
) -> Result<PathBuf> {
    create_timeline_config_inner_with_header_rows(root, target, relative_dirs, error_policy, "3,4")
}

fn create_timeline_config_inner_with_header_rows(
    root: &Path,
    target: Option<&Path>,
    relative_dirs: bool,
    error_policy: &str,
    header_rows: &str,
) -> Result<PathBuf> {
    let config = root.join("app").join("config").join("config.xlsx");
    fs::create_dir_all(config.parent().context("config parent missing")?)?;
    let app = root.join("app");
    fs::create_dir_all(app.join("input"))?;
    let directory_values = if relative_dirs {
        ["input".to_owned(), "output".to_owned(), "logs".to_owned()]
    } else {
        [
            app.join("input").to_string_lossy().into_owned(),
            app.join("output").to_string_lossy().into_owned(),
            app.join("logs").to_string_lossy().into_owned(),
        ]
    };

    let mut workbook = rust_xlsxwriter::Workbook::new();
    {
        let global = workbook.add_worksheet();
        global.set_name("全局配置")?;
        global.write_string(0, 0, "配置项")?;
        global.write_string(0, 1, "值")?;
        let settings = [
            ("输入目录", directory_values[0].clone()),
            ("输出目录", directory_values[1].clone()),
            ("日志目录", directory_values[2].clone()),
            ("错误策略", error_policy.to_owned()),
            ("时序规则命中策略", "all_match".to_owned()),
        ];
        for (index, (key, value)) in settings.iter().enumerate() {
            global.write_string((index + 1) as u32, 0, *key)?;
            global.write_string((index + 1) as u32, 1, value)?;
        }
    }
    {
        let timeline = workbook.add_worksheet();
        timeline.set_name("时序提取规则")?;
        let headers = [
            "是否启用",
            "规则名称",
            "工作簿关键字",
            "工作表关键字",
            "行头列",
            "列表头行",
            "必含列头",
            "必含行头",
            "数据起始行",
            "数据结束行",
            "数据起始列",
            "数据结束列",
            "跳过关键字",
            "目标工作簿路径",
            "目标工作表",
            "启用目标写入",
            "set区域",
            "目标去重列",
            "是否输出行序号",
        ];
        for (column, header) in headers.iter().enumerate() {
            timeline.write_string(0, column as u16, *header)?;
        }
        let target_workbook = target
            .map(|path| path.to_string_lossy().into_owned())
            .unwrap_or_default();
        let target_sheet = if target.is_some() { "Append" } else { "" };
        let target_enabled = if target.is_some() { "是" } else { "否" };
        let values = [
            "是",
            "WideRule",
            "sample",
            "Report",
            "A",
            header_rows,
            "",
            "",
            "5",
            "",
            "B",
            "D",
            "",
            target_workbook.as_str(),
            target_sheet,
            target_enabled,
            "集团@B2",
            "",
            "是",
        ];
        for (column, value) in values.iter().enumerate() {
            timeline.write_string(1, column as u16, *value)?;
        }
    }
    {
        let path_maps = workbook.add_worksheet();
        path_maps.set_name("路径标准化映射")?;
        for (column, header) in [
            "是否启用",
            "映射名称",
            "适用规则名",
            "工作簿关键字",
            "工作表关键字",
            "作用对象",
            "匹配方式",
            "原始路径",
            "标准路径",
            "备注",
        ]
        .iter()
        .enumerate()
        {
            path_maps.write_string(0, column as u16, *header)?;
        }
    }
    workbook
        .save(&config)
        .context("save timeline menu configuration")?;
    Ok(config)
}

fn create_legacy_xls_timeline_config(root: &Path) -> Result<PathBuf> {
    let app = root.join("app");
    let input_dir = app.join("input");
    let output_dir = app.join("output");
    let log_dir = app.join("logs");
    fs::create_dir_all(&input_dir)?;
    fs::create_dir_all(&output_dir)?;
    fs::create_dir_all(&log_dir)?;

    let config = app.join("config").join("config.xlsx");
    fs::create_dir_all(config.parent().context("config parent missing")?)?;

    let mut workbook = rust_xlsxwriter::Workbook::new();
    {
        let global = workbook.add_worksheet();
        global.set_name("全局配置")?;
        global.write_string(0, 0, "配置项")?;
        global.write_string(0, 1, "值")?;
        let settings = [
            ("输入目录", input_dir.to_string_lossy().into_owned()),
            ("输出目录", output_dir.to_string_lossy().into_owned()),
            ("日志目录", log_dir.to_string_lossy().into_owned()),
            ("错误策略", "fail_fast".to_owned()),
            ("源文件扩展名", ".xlsx".to_owned()),
            ("时序规则命中策略", "all_match".to_owned()),
        ];
        for (index, (key, value)) in settings.iter().enumerate() {
            global.write_string((index + 1) as u32, 0, *key)?;
            global.write_string((index + 1) as u32, 1, value)?;
        }
    }
    {
        let timeline = workbook.add_worksheet();
        timeline.set_name("时序提取规则")?;
        let headers = [
            "是否启用",
            "规则名称",
            "工作簿关键字",
            "工作表关键字",
            "行头列",
            "列表头行",
            "必含列头",
            "必含行头",
            "数据起始行",
            "数据结束行",
            "数据起始列",
            "数据结束列",
            "跳过关键字",
            "目标工作簿路径",
            "目标工作表",
            "启用目标写入",
            "set区域",
            "目标去重列",
            "是否输出行序号",
        ];
        for (column, header) in headers.iter().enumerate() {
            timeline.write_string(0, column as u16, *header)?;
        }
        let values = [
            "是", "WideRule", "input", "Data", "A", "1", "", "", "2", "", "B", "B", "", "", "",
            "否", "", "", "是",
        ];
        for (column, value) in values.iter().enumerate() {
            timeline.write_string(1, column as u16, *value)?;
        }
    }
    {
        let path_maps = workbook.add_worksheet();
        path_maps.set_name("路径标准化映射")?;
        for (column, header) in [
            "是否启用",
            "映射名称",
            "适用规则名",
            "工作簿关键字",
            "工作表关键字",
            "作用对象",
            "匹配方式",
            "原始路径",
            "标准路径",
            "备注",
        ]
        .iter()
        .enumerate()
        {
            path_maps.write_string(0, column as u16, *header)?;
        }
    }
    workbook
        .save(&config)
        .context("save legacy XLS timeline configuration")?;
    Ok(config)
}

fn latest_menu_json_report(output: &Output) -> Result<Value> {
    String::from_utf8_lossy(&output.stdout)
        .lines()
        .rev()
        .find_map(|line| {
            serde_json::from_str::<Value>(line)
                .ok()
                .filter(|value| value.get("output").is_some())
        })
        .context("interactive menu did not print a wide-summary report")
}

fn latest_timeline_compare_report(output: &Output) -> Result<Value> {
    let start = output
        .stdout
        .iter()
        .position(|byte| *byte == b'{')
        .context("interactive menu did not print a timeline comparison report")?;
    serde_json::Deserializer::from_slice(&output.stdout[start..])
        .into_iter::<Value>()
        .next()
        .context("timeline comparison JSON missing")?
        .context("decode timeline comparison JSON")
}

fn create_multi_level_header_compare_source(path: &Path, third_column_header: &str) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Report")?;
    sheet.merge_range(1, 1, 1, 2, "资产", &rust_xlsxwriter::Format::new())?;
    sheet.write_string(2, 0, "地区")?;
    sheet.write_string(2, 1, "人民币")?;
    sheet.write_string(2, 2, third_column_header)?;
    sheet.write_string(3, 0, "机构")?;
    sheet.write_string(3, 1, "余额")?;
    sheet.write_string(3, 2, "余额")?;
    sheet.write_string(4, 0, "甲机构")?;
    sheet.write_number(4, 1, 10.0)?;
    sheet.write_number(4, 2, 20.0)?;
    workbook
        .save(path)
        .context("save multilevel header compare workbook")
}

fn create_wide_source(path: &Path, organization: &str, amount: f64) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Report")?;
    sheet.write_string(0, 0, "报表")?;
    sheet.merge_range(1, 0, 1, 1, "集团A", &rust_xlsxwriter::Format::new())?;
    sheet.write_string(2, 0, "地区")?;
    sheet.merge_range(2, 1, 2, 2, "存款", &rust_xlsxwriter::Format::new())?;
    sheet.write_string(2, 3, "贷款")?;
    sheet.write_string(3, 0, "名称")?;
    sheet.write_string(3, 1, "余额")?;
    sheet.write_string(3, 2, "余额")?;
    sheet.write_string(3, 3, "本息")?;
    sheet.write_string(4, 0, organization)?;
    sheet.write_number(4, 1, amount)?;
    sheet.write_number(4, 2, amount + 1.0)?;
    sheet.write_number(4, 3, amount + 2.0)?;
    workbook.save(path).context("save timeline source workbook")
}

fn create_header_compare_source_with_sheet(path: &Path, sheet_name: &str) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name(sheet_name)?;
    sheet.write_string(0, 0, "test")?;
    workbook
        .save(path)
        .context("save header comparison workbook with unmatched worksheet")
}

fn create_wide_source_with_elapsed_duration(
    path: &Path,
    organization: &str,
    serial: f64,
) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Report")?;
    sheet.write_string(0, 0, "报表")?;
    sheet.merge_range(1, 0, 1, 1, "集团A", &rust_xlsxwriter::Format::new())?;
    sheet.write_string(2, 0, "地区")?;
    sheet.merge_range(2, 1, 2, 2, "存款", &rust_xlsxwriter::Format::new())?;
    sheet.write_string(2, 3, "贷款")?;
    sheet.write_string(3, 0, "名称")?;
    sheet.write_string(3, 1, "余额")?;
    sheet.write_string(3, 2, "余额")?;
    sheet.write_string(3, 3, "本息")?;
    sheet.write_string(4, 0, organization)?;
    let elapsed_format = rust_xlsxwriter::Format::new().set_num_format("[h]:mm:ss.000");
    sheet.write_number_with_format(4, 1, serial, &elapsed_format)?;
    sheet.write_number(4, 2, 12.0)?;
    sheet.write_number(4, 3, 23.0)?;
    workbook
        .save(path)
        .context("save elapsed-duration timeline source workbook")
}

fn create_wide_source_with_date_serial(path: &Path, organization: &str, serial: f64) -> Result<()> {
    create_wide_source_with_number_format_serial(path, organization, serial, "yyyy/m/d")
}

fn create_wide_source_with_number_format_serial(
    path: &Path,
    organization: &str,
    serial: f64,
    number_format: &str,
) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Report")?;
    sheet.write_string(0, 0, "报表")?;
    sheet.merge_range(1, 0, 1, 1, "集团A", &rust_xlsxwriter::Format::new())?;
    sheet.write_string(2, 0, "地区")?;
    sheet.merge_range(2, 1, 2, 2, "存款", &rust_xlsxwriter::Format::new())?;
    sheet.write_string(2, 3, "贷款")?;
    sheet.write_string(3, 0, "名称")?;
    sheet.write_string(3, 1, "余额")?;
    sheet.write_string(3, 2, "余额")?;
    sheet.write_string(3, 3, "本息")?;
    sheet.write_string(4, 0, organization)?;
    let format = rust_xlsxwriter::Format::new().set_num_format(number_format);
    sheet.write_number_with_format(4, 1, serial, &format)?;
    sheet.write_number(4, 2, 12.0)?;
    sheet.write_number(4, 3, 23.0)?;
    workbook
        .save(path)
        .context("save Excel 1900-boundary timeline source workbook")
}

fn create_wide_source_with_negative_date(path: &Path, organization: &str) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Report")?;
    sheet.write_string(0, 0, "报表")?;
    sheet.merge_range(1, 0, 1, 1, "集团A", &rust_xlsxwriter::Format::new())?;
    sheet.write_string(2, 0, "地区")?;
    sheet.merge_range(2, 1, 2, 2, "存款", &rust_xlsxwriter::Format::new())?;
    sheet.write_string(2, 3, "贷款")?;
    sheet.write_string(3, 0, "名称")?;
    sheet.write_string(3, 1, "余额")?;
    sheet.write_string(3, 2, "余额")?;
    sheet.write_string(3, 3, "本息")?;
    sheet.write_string(4, 0, organization)?;
    let date_format = rust_xlsxwriter::Format::new().set_num_format("yyyy-mm-dd");
    sheet.write_number_with_format(4, 1, -0.5, &date_format)?;
    sheet.write_number(4, 2, 24.5)?;
    sheet.write_number(4, 3, 25.5)?;
    workbook
        .save(path)
        .context("save negative-date timeline source workbook")
}

fn create_date_header_source(path: &Path, year: u16, month: u8, day: u8) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Report")?;
    let date = rust_xlsxwriter::ExcelDateTime::from_ymd(year, month, day)?;
    let date_format = rust_xlsxwriter::Format::new().set_num_format("yyyy-mm-dd");
    sheet.write_string(2, 0, "地区")?;
    sheet.write_string(2, 1, "存款")?;
    sheet.write_string(3, 0, "名称")?;
    sheet.write_datetime_with_format(3, 1, &date, &date_format)?;
    sheet.write_string(4, 0, "甲支行")?;
    sheet.write_number(4, 1, 1.0)?;
    workbook
        .save(path)
        .context("save date-header timeline source")
}

fn find_output(output_dir: &Path, stem_contains: &str) -> Result<PathBuf> {
    fs::read_dir(output_dir)?
        .map(|entry| entry.map(|entry| entry.path()))
        .collect::<std::io::Result<Vec<_>>>()?
        .into_iter()
        .find(|path| {
            path.extension()
                .is_some_and(|extension| extension == "xlsx")
                && path
                    .file_name()
                    .is_some_and(|name| name.to_string_lossy().contains(stem_contains))
        })
        .with_context(|| {
            format!(
                "no XLSX output containing {stem_contains:?} in {}",
                output_dir.display()
            )
        })
}

fn run_menu(config: &Path, input: &str) -> Result<Output> {
    run_menu_with_view(config, "user", input)
}

fn run_default_menu(input: &str) -> Result<Output> {
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .context("start pytools directly without arguments")?;
    child
        .stdin
        .take()
        .context("default menu stdin missing")?
        .write_all(input.as_bytes())
        .context("send default menu selection")?;
    child
        .wait_with_output()
        .context("wait for default interactive menu")
}

#[test]
fn launching_without_arguments_opens_the_numbered_main_menu() -> Result<()> {
    let output = run_default_menu("0\n")?;
    assert!(
        output.status.success(),
        "direct menu launch failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    for expected in [
        "======== PyTools 菜单 ========",
        "1  数据提取与汇总",
        "2  数据去重与追加",
        "3  报表打印",
        "4  文件处理",
        "5  专用报表处理",
        "6  归档整理",
        "7  配置与快捷操作",
        "0  退出",
        "再见。",
    ] {
        assert!(
            stdout.contains(expected),
            "missing {expected:?} in menu output: {stdout}"
        );
    }
    Ok(())
}

#[test]
fn excluded_pt_1_2_is_not_exposed_by_the_v2_data_menu() -> Result<()> {
    let directory = TestDirectory::new()?;
    let config = create_timeline_config(directory.path())?;
    let output = run_menu(&config, "1\n0\n0\n")?;
    assert!(
        output.status.success(),
        "V2 data menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(
        !stdout.contains("时序提取（PT-1-2"),
        "the excluded PT-1-2 action must not be listed: {stdout}"
    );
    assert!(!stdout.contains("选择时序提取源文件"));
    assert!(!stdout.contains("PT-1-2 开始"));
    assert!(stdout.contains("0  退出"));
    Ok(())
}

#[test]
fn timeline_precheck_menu_prints_readable_success_and_returns_to_main() -> Result<()> {
    let directory = TestDirectory::new()?;
    let config = create_timeline_config(directory.path())?;
    let output = run_menu(&config, "1\n5\n0\n")?;
    assert!(
        output.status.success(),
        "successful precheck menu route failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("=== 预校验报告 ==="));
    assert!(stdout.contains("=== 预校验完成: 错误 0 项，警告 0 项 ==="));
    assert!(stdout.contains("[完成] 时序配置预校验通过。"));
    assert!(stdout.contains("======== PyTools 菜单 ========"));
    assert!(stdout.contains("再见。"));
    assert!(!stdout.contains("\"errors\""));
    Ok(())
}

#[test]
fn timeline_precheck_menu_reports_errors_without_terminating_menu() -> Result<()> {
    let directory = TestDirectory::new()?;
    let config = create_timeline_config(directory.path())?;
    let input_directory = directory.path().join("app").join("input");
    fs::remove_dir_all(&input_directory)?;

    let output = run_menu(&config, "1\n5\n0\n")?;
    assert!(
        output.status.success(),
        "failed precheck must return to the menu: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("[错误] 全局配置/输入目录: 输入目录不存在:"));
    assert!(stdout.contains("=== 预校验完成: 错误 1 项，警告 0 项 ==="));
    assert!(stdout.contains("[失败] 发现 1 项错误"));
    assert!(stdout.contains("======== PyTools 菜单 ========"));
    assert!(stdout.contains("再见。"));
    assert!(!stdout.contains("\"errors\""));
    Ok(())
}

fn run_menu_with_view(config: &Path, log_view: &str, input: &str) -> Result<Output> {
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
        .context("start interactive pytools menu")?;
    child
        .stdin
        .take()
        .context("interactive stdin missing")?
        .write_all(input.as_bytes())
        .context("send interactive menu selections")?;
    child
        .wait_with_output()
        .context("wait for interactive menu")
}

fn run_menu_with_capabilities(
    config: &Path,
    log_view: &str,
    capabilities: &Path,
    platform: &str,
    input: &str,
) -> Result<Output> {
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("menu")
        .arg("--config")
        .arg(config)
        .arg("--log-view")
        .arg(log_view)
        .arg("--capabilities")
        .arg(capabilities)
        .arg("--platform")
        .arg(platform)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .context("start interactive pytools menu with edit capability")?;
    child
        .stdin
        .take()
        .context("interactive stdin missing")?
        .write_all(input.as_bytes())
        .context("send interactive menu selections")?;
    child
        .wait_with_output()
        .context("wait for interactive menu")
}
