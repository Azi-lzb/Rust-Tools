use anyhow::{Context, Result};
use excel_editor::{
    BackendCapability, CapabilityEvidence, CapabilityStatus, EditOperation, EngineId,
};
use excel_reader::read_workbook;
use pytools_core::sheet_rename::build_sheet_rename_plan;
use serde::Deserialize;
use serde_json::Value;
use std::{
    env, fs,
    io::Write,
    path::{Path, PathBuf},
    process::{Command, Output, Stdio},
    sync::atomic::{AtomicU64, Ordering},
    time::{SystemTime, UNIX_EPOCH},
};

static NEXT_TEST_DIRECTORY_ID: AtomicU64 = AtomicU64::new(0);
const SHEET_RENAME_PLATFORM: &str = "pt-4-4-menu-synthetic";

struct TestDirectory(PathBuf);

impl TestDirectory {
    fn new() -> Result<Self> {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .context("system clock before Unix epoch")?
            .as_nanos();
        let serial = NEXT_TEST_DIRECTORY_ID.fetch_add(1, Ordering::Relaxed);
        let path = std::env::temp_dir().join(format!(
            "pytools-conversion-menu-{}-{nonce}-{serial}",
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
fn v2_file_menu_routes_supported_workflows() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let sources = root.path().join("sources");
    fs::create_dir_all(&sources)?;

    let rename_folder = sources.join("rename-folder");
    fs::create_dir_all(&rename_folder)?;
    let rename_source = rename_folder.join("Bank-report.txt");
    fs::write(&rename_source, b"rename payload")?;
    let renamed = rename_folder.join("2026-10-24 B01 Example Bank Monthly.txt");
    let rename = run_action(&config, 1, 3, &rename_folder)?;
    assert!(
        rename.status.success(),
        "rename failed: {}",
        stderr(&rename)
    );
    assert!(renamed.is_file(), "folder input must expand to its files");
    assert_eq!(fs::read(&renamed)?, b"rename payload");

    let legacy_excel_source = sources.join("needs-engine.xls");
    fs::write(&legacy_excel_source, b"synthetic xls payload")?;
    let excel_before = fs::read(&legacy_excel_source)?;
    let excel_engine_failure = run_action_with_engines_disabled_and_log_view(
        &config,
        2,
        1,
        &legacy_excel_source,
        "developer",
    )?;
    assert!(
        excel_engine_failure.status.success(),
        "a handled conversion failure should return to the menu: {}",
        stderr(&excel_engine_failure)
    );
    assert!(
        String::from_utf8_lossy(&excel_engine_failure.stdout).contains("文件处理菜单仍可继续操作")
    );
    assert!(
        String::from_utf8_lossy(&excel_engine_failure.stdout).contains("------ 文件处理 ------")
    );
    assert!(stderr(&excel_engine_failure).contains("needs-engine.xls"));
    assert!(stderr(&excel_engine_failure).contains("需要可用的转换引擎"));
    let excel_failure_report = output_json(&excel_engine_failure)?;
    assert_eq!(excel_failure_report["failed"], 1);
    assert_eq!(excel_failure_report["items"][0]["status"], "failed");
    assert!(excel_failure_report["items"][0]["detail"]
        .as_str()
        .unwrap_or_default()
        .contains("该功能需要外部引擎"));
    assert_eq!(fs::read(&legacy_excel_source)?, excel_before);
    assert!(excel_failure_report["items"][0]["output"].is_null());

    let legacy_word_source = sources.join("needs-engine.doc");
    fs::write(&legacy_word_source, b"synthetic doc payload")?;
    let word_before = fs::read(&legacy_word_source)?;
    let word_engine_failure = run_action_with_engines_disabled_and_log_view(
        &config,
        3,
        1,
        &legacy_word_source,
        "developer",
    )?;
    assert!(
        word_engine_failure.status.success(),
        "a handled conversion failure should return to the menu: {}",
        stderr(&word_engine_failure)
    );
    assert!(
        String::from_utf8_lossy(&word_engine_failure.stdout).contains("文件处理菜单仍可继续操作")
    );
    assert!(stderr(&word_engine_failure).contains("needs-engine.doc"));
    assert!(stderr(&word_engine_failure).contains("需要可用的转换引擎"));
    let word_failure_report = output_json(&word_engine_failure)?;
    assert_eq!(word_failure_report["failed"], 1);
    assert!(word_failure_report["items"][0]["detail"]
        .as_str()
        .unwrap_or_default()
        .contains("该功能需要外部引擎"));
    assert_eq!(fs::read(&legacy_word_source)?, word_before);
    assert!(word_failure_report["items"][0]["output"].is_null());

    let sheet_source = sources.join("sheet-rename.xlsx");
    create_source_workbook(&sheet_source, "OldSheet")?;
    let sheet_rename = run_action(&config, 4, 1, &sheet_source)?;
    assert!(
        sheet_rename.status.success(),
        "sheet rename capability check failed: {}",
        stderr(&sheet_rename)
    );
    assert!(String::from_utf8_lossy(&sheet_rename.stdout)
        .contains("[无法执行] 缺少已验收的 ExcelModifier capability/platform"));
    assert!(sheet_source.is_file());
    let plan_files = find_files_with_extension(&root.path().join("output"), "json")?;
    assert!(plan_files.is_empty());

    let content_source = sources.join("add-content.xlsx");
    create_source_workbook(&content_source, "Plain")?;
    let content_before = fs::read(&content_source)?;
    let add_content = run_action(&config, 5, 1, &content_source)?;
    assert!(
        add_content.status.success(),
        "content prefix route failed: {}",
        stderr(&add_content)
    );
    assert!(!content_source.exists());
    assert_eq!(
        fs::read(sources.join("Bank_add-content.xlsx"))?,
        content_before
    );

    let legacy_xls_fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("timeline")
        .join("wide_summary_legacy.xls");
    let xls_content_source = sources.join("legacy-content.xls");
    fs::copy(&legacy_xls_fixture, &xls_content_source)?;
    let xls_content_before = fs::read(&xls_content_source)?;
    let xls_content = run_action_with_log_view(&config, 5, 1, &xls_content_source, "developer")?;
    assert!(
        xls_content.status.success(),
        "content-name menu route rejected XLS: {}",
        stderr(&xls_content)
    );
    assert_eq!(output_json(&xls_content)?["ok"], 1);
    let xls_content_target = sources.join("Installed smoke_legacy-content.xls");
    assert!(!xls_content_source.exists());
    assert_eq!(fs::read(&xls_content_target)?, xls_content_before);

    let submission_dir = sources.join("submission-files");
    fs::create_dir_all(&submission_dir)?;
    for (name, bytes) in [
        ("2026_Area_Bank_base.txt", b"base".as_slice()),
        ("2026_Area_Bank_multi.txt", b"multi".as_slice()),
        ("misc_submission.txt", b"unmatched".as_slice()),
    ] {
        fs::write(submission_dir.join(name), bytes)?;
    }
    let submission = run_action_with_log_view(&config, 8, 2, &submission_dir, "developer")?;
    assert!(
        submission.status.success(),
        "submission check failed: {}",
        stderr(&submission)
    );
    let submission_report = output_json(&submission)?;
    assert_eq!(submission_report["files"], 3);
    assert_eq!(submission_report["submitted"], 2);
    assert_eq!(submission_report["missing"], 1);
    assert_eq!(submission_report["unmatched"], 1);
    assert_eq!(submission_report["multi"], 1);
    let submission_path = PathBuf::from(
        submission_report["output"]
            .as_str()
            .context("submission report path missing")?,
    );
    assert!(submission_path.is_file());
    let submission_workbook = read_workbook(&submission_path)?;
    assert_eq!(
        submission_workbook
            .sheets
            .iter()
            .map(|sheet| sheet.name.as_str())
            .collect::<Vec<_>>(),
        vec!["检查结果", "缺失清单", "未识别文件", "多重命中文件"]
    );
    assert_eq!(
        submission_workbook.sheets[0].values["F2"]["value"].as_f64(),
        Some(2.0),
        "the published workbook must contain the matched rule count"
    );
    assert_eq!(
        submission_workbook.sheets[2].values["A2"]["value"],
        "misc_submission.txt"
    );
    assert!(!fs::read_dir(
        submission_path
            .parent()
            .context("submission output directory missing")?
    )?
    .filter_map(std::result::Result::ok)
    .any(|entry| entry
        .file_name()
        .to_string_lossy()
        .starts_with(".pytools-submission-check-")));

    let sort_source = sources.join("sort-payload.txt");
    fs::write(&sort_source, b"sort")?;
    let sort = run_action(&config, 9, 1, &sort_source)?;
    assert!(
        sort.status.success(),
        "sort prefix failed: {}",
        stderr(&sort)
    );
    assert!(sources.join("1_sort-payload.txt").is_file());

    let remove_sort_source = sources.join("1_remove-sort.txt");
    fs::write(&remove_sort_source, b"remove sort")?;
    let remove_sort = run_action(&config, 10, 1, &remove_sort_source)?;
    assert!(
        remove_sort.status.success(),
        "remove sort prefix failed: {}",
        stderr(&remove_sort)
    );
    assert!(sources.join("remove-sort.txt").is_file());

    Ok(())
}

#[test]
fn v2_file_menu_hides_native_value_conversion() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let output = run_menu(&config, "4\n0\n0\n")?;
    assert!(output.status.success(), "menu failed: {}", stderr(&output));
    let menu = String::from_utf8_lossy(&output.stdout);
    assert!(menu.contains("Office 格式转换（XLS / ET → XLSX，DOC → DOCX）"));
    assert!(!menu.contains("批量原生值转换"));
    Ok(())
}

#[test]
fn native_conversion_help_lists_legacy_epoch_formats() -> Result<()> {
    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("--help")
        .output()?;
    assert!(
        output.status.success(),
        "pytools --help failed: {}",
        stderr(&output)
    );
    let help = String::from_utf8_lossy(&output.stderr);
    assert!(
        help.contains("检测到的 1904 epoch XLS/XLSM/XLSB 会拒绝转换"),
        "help output did not describe the 1904 date-system guard: {help}"
    );
    assert!(
        help.contains(
            "convert-excel --config <config.xlsx> [--to xlsx] [--engine auto|lo|com|native]"
        ),
        "Excel conversion help must list the supported COM engine: {help}"
    );
    assert!(
        help.contains("convert-word --config <config.xlsx> [--to docx] [--engine auto|lo|com]"),
        "Word conversion help must list the supported COM engine: {help}"
    );
    assert!(
        help.contains("convert-file --input <file> --to <extension> [--output <file>] [--engine auto|lo|com|native]"),
        "single-file conversion help must list the supported COM engine: {help}"
    );
    Ok(())
}

#[test]
fn conversion_cli_rejects_xlsx_payload_mislabeled_as_et() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let sources = root.path().join("sources");
    fs::create_dir_all(&sources)?;
    let xlsx = sources.join("source.xlsx");
    create_source_workbook(&xlsx, "Other")?;
    let mislabeled_et = sources.join("mislabeled.et");
    fs::copy(&xlsx, &mislabeled_et)?;
    let source_before = fs::read(&mislabeled_et)?;
    let output_dir = root.path().join("output");

    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("convert-excel")
        .arg("--config")
        .arg(&config)
        .arg("--to")
        .arg("xlsx")
        .arg("--engine")
        .arg("com")
        .arg("--output-dir")
        .arg(&output_dir)
        .arg("--log-view")
        .arg("developer")
        .arg(&mislabeled_et)
        .output()?;

    let report = output_json(&output)?;
    assert_eq!(report["converted"], 0);
    assert_eq!(report["failed"], 1);
    assert_eq!(report["items"][0]["status"], "failed");
    assert!(report["items"][0]["detail"]
        .as_str()
        .unwrap_or_default()
        .contains("refusing extension-only conversion"));
    assert!(!report["items"][0]["detail"]
        .as_str()
        .unwrap_or_default()
        .contains("engine is unavailable"));
    assert!(!output_dir.join("mislabeled.xlsx").exists());
    assert_eq!(fs::read(&mislabeled_et)?, source_before);
    Ok(())
}

#[test]
#[ignore = "requires Windows Office/WPS COM; run explicitly for PT-4-4 source-format regression"]
fn conversion_cli_does_not_publish_xls_payload_mislabeled_as_et() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let sources = root.path().join("sources");
    fs::create_dir_all(&sources)?;
    let xls_fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/timeline/wide_summary_elapsed_duration.xls");
    let mislabeled_et = sources.join("mislabeled.et");
    fs::copy(&xls_fixture, &mislabeled_et)?;
    let source_before = fs::read(&mislabeled_et)?;
    let output_dir = root.path().join("output");

    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("convert-excel")
        .arg("--config")
        .arg(&config)
        .arg("--to")
        .arg("xlsx")
        .arg("--engine")
        .arg("com")
        .arg("--output-dir")
        .arg(&output_dir)
        .arg("--log-view")
        .arg("off")
        .arg(&mislabeled_et)
        .output()?;

    let report = output_json(&output)?;
    assert_eq!(report["converted"], 0);
    assert_eq!(report["failed"], 1);
    let detail = report["items"][0]["detail"].as_str().unwrap_or_default();
    assert!(
        detail.contains("native ET")
            || detail.contains("原生 ET")
            || detail.contains("FileFormat")
            || detail.contains("65521"),
        "failure must identify the unverified/mismatched ET source format: {detail}"
    );
    assert!(!output_dir.join("mislabeled.xlsx").exists());
    assert_eq!(fs::read(&mislabeled_et)?, source_before);
    assert_eq!(fs::read_dir(&output_dir)?.count(), 0);
    Ok(())
}

#[test]
#[ignore = "requires Windows Excel COM; converts temporary copies through the V2 menu"]
fn conversion_menu_batch_uses_rust_com_and_preserves_sources() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let sources = root.path().join("sources");
    fs::create_dir_all(&sources)?;
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/timeline/wide_summary_elapsed_duration.xls");
    let first = sources.join("first.xls");
    let second = sources.join("second.xls");
    fs::copy(&fixture, &first)?;
    fs::copy(&fixture, &second)?;
    let source_bytes = [fs::read(&first)?, fs::read(&second)?];

    let input = format!(
        "4\n4\n1\n1\n{};{}\n0\n0\n",
        first.display(),
        second.display()
    );
    let mut child = menu_command(&config);
    child.arg("--log-view").arg("developer");
    let output = finish_menu(child, &input)?;
    assert!(
        output.status.success(),
        "conversion menu failed: stdout={} stderr={}",
        String::from_utf8_lossy(&output.stdout),
        stderr(&output)
    );

    let report = output_json(&output)?;
    assert_eq!(report["converted"], 2, "{report}");
    assert_eq!(report["failed"], 0, "{report}");
    let items = report["items"].as_array().context("conversion items missing")?;
    assert_eq!(items.len(), 2);
    for (index, item) in items.iter().enumerate() {
        assert_eq!(item["status"], "converted");
        let detail: Value = serde_json::from_str(
            item["detail"].as_str().context("conversion detail missing")?,
        )?;
        assert_eq!(detail["com_bridge"], "rust");
        assert_eq!(detail["com_application"], "Excel.Application");
        let output_path = PathBuf::from(
            item["output"]
                .as_str()
                .context("converted output path missing")?,
        );
        assert!(output_path.is_file(), "{}", output_path.display());
        assert_eq!(fs::read([&first, &second][index])?, source_bytes[index]);
    }

    let history_path = PathBuf::from(
        report["run_history_path"]
            .as_str()
            .context("conversion run history path missing")?,
    );
    let history: Value = serde_json::from_slice(&fs::read(history_path)?)?;
    assert!(history["actual_engines"]
        .as_array()
        .is_some_and(|engines| engines.iter().any(|engine| engine == "Microsoft Excel COM (Rust)")));
    Ok(())
}

#[test]
#[ignore = "requires Windows Office/WPS COM and a verified native ET fixture; set PYTOOLS_NATIVE_ET_FIXTURE"]
fn conversion_cli_converts_verified_native_et_and_preserves_source() -> Result<()> {
    let fixture = env::var_os("PYTOOLS_NATIVE_ET_FIXTURE")
        .context("set PYTOOLS_NATIVE_ET_FIXTURE to a verified native ET workbook")?;
    let fixture = PathBuf::from(fixture);
    let fixture_before = fs::read(&fixture)?;
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let sources = root.path().join("sources");
    fs::create_dir_all(&sources)?;
    let native_et = sources.join("native.et");
    fs::copy(&fixture, &native_et)?;
    let source_before = fs::read(&native_et)?;
    let output_dir = root.path().join("output");

    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("convert-excel")
        .arg("--config")
        .arg(&config)
        .arg("--to")
        .arg("xlsx")
        .arg("--engine")
        .arg("com")
        .arg("--output-dir")
        .arg(&output_dir)
        .arg("--log-view")
        .arg("off")
        .arg(&native_et)
        .output()?;

    assert!(
        output.status.success(),
        "native ET CLI conversion failed: {}",
        stderr(&output)
    );
    let report = output_json(&output)?;
    assert_eq!(report["converted"], 1);
    assert_eq!(report["failed"], 0);
    assert_eq!(report["items"][0]["status"], "converted");
    let detail = report["items"][0]["detail"]
        .as_str()
        .context("successful ET conversion omitted its sidecar report")?;
    let sidecar_report: Value = serde_json::from_str(detail)?;
    assert_eq!(sidecar_report["status"], "converted");
    assert_eq!(sidecar_report["engine"], "com");
    assert_eq!(sidecar_report["validation"]["source_format_verified"], true);
    assert_eq!(sidecar_report["validation"]["source_format"], "et");
    assert_eq!(sidecar_report["validation"]["source_file_format"], 65521);
    assert_eq!(
        sidecar_report["validation"]["level"],
        "worksheet_order_names_values_and_formula_text"
    );

    let converted = output_dir.join("native.xlsx");
    let reported_output = PathBuf::from(
        report["items"][0]["output"]
            .as_str()
            .context("successful ET conversion omitted its output path")?,
    );
    assert_eq!(reported_output, converted);
    assert!(
        converted.is_file(),
        "validated XLSX output was not published"
    );
    let workbook = read_workbook(&converted)?;
    assert!(
        !workbook.sheets.is_empty(),
        "converted workbook has no sheets"
    );
    assert_eq!(fs::read(&native_et)?, source_before);
    assert_eq!(fs::read(&fixture)?, fixture_before);
    assert_eq!(fs::read_dir(&output_dir)?.count(), 1);
    Ok(())
}

#[test]
fn conversion_menu_and_cli_emit_redacted_dual_view_progress() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let sources = root.path().join("sources");
    let log_dir = root.path().join("logs");
    fs::create_dir_all(&sources)?;

    let user_source = sources.join("user-private-report.csv");
    fs::write(&user_source, "Region,Value\nNorth,17\n")?;
    let user_output_dir = root.path().join("user-output");
    let user = run_convert_excel_cli(&config, &user_source, &user_output_dir, "native", "user")?;
    assert!(
        user.status.success(),
        "user conversion failed: {}",
        stderr(&user)
    );
    assert!(user_output_dir.join("user-private-report.xlsx").is_file());
    let user_stderr = stderr(&user);
    assert!(user_stderr.contains("已发现 1 个待处理文件"));
    assert!(user_stderr.contains("正在处理第 1/1 个文件"));
    assert!(user_stderr.contains("转换结束：成功 1，失败 0"));
    assert!(!user_stderr.contains("user-private-report"));
    let user_report = output_json(&user)?;
    let unified_path = PathBuf::from(user_report["run_summary_path"].as_str().context("unified summary path missing")?);
    let unified: Value = serde_json::from_slice(&fs::read(unified_path)?)?;
    assert_eq!(unified["run_id"], user_report["run_id"]);
    assert_eq!(unified["overall_status"], "SUCCESS");
    assert_eq!(unified["input_stats"]["selected"], 1);
    assert_eq!(unified["input_stats"]["succeeded"], 1);
    let events_path = log_dir.join(unified["diagnostic_log_path"].as_str().context("diagnostic path missing")?);
    let structured = fs::read_to_string(events_path)?;
    let lines = structured.lines().map(serde_json::from_str::<Value>).collect::<std::result::Result<Vec<_>, _>>()?;
    assert!(lines.iter().all(|event| event["run_id"] == unified["run_id"]));
    assert!(lines.iter().any(|event| event["file_name"] == "user-private-report.csv"));

    let native_source = sources.join("native-private.csv");
    fs::write(&native_source, "Region,Value\nNorth,17\n")?;
    let native_output_dir = root.path().join("native-output");
    let native = run_convert_excel_cli(
        &config,
        &native_source,
        &native_output_dir,
        "native",
        "developer",
    )?;
    assert!(
        native.status.success(),
        "native conversion failed: {}",
        stderr(&native)
    );
    let native_report = output_json(&native)?;
    assert_eq!(native_report["converted"], 1);
    let native_history_path = PathBuf::from(
        native_report["run_history_path"]
            .as_str()
            .context("native conversion history path missing")?,
    );
    assert!(native_output_dir.join("native-private.xlsx").is_file());
    let native_history_bytes = fs::read(&native_history_path)?;
    let native_history_text = String::from_utf8_lossy(&native_history_bytes);
    assert!(!native_history_text.contains(&native_source.display().to_string()));
    assert!(!native_history_text.contains("native-private.csv"));
    let native_history: Value = serde_json::from_slice(&native_history_bytes)?;
    assert_eq!(native_history["feature_id"], "PT-4-2");
    assert!(native_history["actual_engines"]
        .as_array()
        .context("actual engine list missing")?
        .iter()
        .any(|engine| engine == "csv + rust_xlsxwriter"));
    let native_events = native_history["events"]
        .as_array()
        .context("events missing")?;
    assert!(native_events.iter().any(|event| {
        event["phase"] == "convert_file"
            && event["status"] == "completed"
            && event["actual_engine"] == "csv + rust_xlsxwriter"
            && event["components"].as_array().is_some_and(|components| {
                components
                    .iter()
                    .any(|component| component == "rust_xlsxwriter")
            })
            && event["stage_elapsed_ms"].is_number()
            && event["total_elapsed_ms"].is_number()
    }));
    assert!(stderr(&native).contains("\"feature_id\":\"PT-4-2\""));

    let off_source = sources.join("off-private.csv");
    fs::write(&off_source, "Region,Value\nSouth,19\n")?;
    let off_output_dir = root.path().join("off-output");
    let history_count = fs::read_dir(&log_dir)?.count();
    let off = run_convert_excel_cli(&config, &off_source, &off_output_dir, "native", "off")?;
    assert!(
        off.status.success(),
        "off conversion failed: {}",
        stderr(&off)
    );
    assert!(off_output_dir.join("off-private.xlsx").is_file());
    assert!(stderr(&off).is_empty());
    assert_eq!(fs::read_dir(&log_dir)?.count(), history_count);

    let failed_source = sources.join("private-needs-engine.xls");
    fs::write(&failed_source, b"synthetic legacy workbook")?;
    let before = fs::read(&failed_source)?;
    let failed = run_action_with_engines_disabled_and_log_view(
        &config,
        2,
        1,
        &failed_source,
        "developer",
    )?;
    assert!(
        failed.status.success(),
        "interactive conversion failure should not close the menu: {}",
        stderr(&failed)
    );
    assert!(stderr(&failed).contains("private-needs-engine.xls"));
    assert_eq!(fs::read(&failed_source)?, before);
    let failed_report = output_json(&failed)?;
    assert_eq!(failed_report["failed"], 1);
    let failed_history_path = PathBuf::from(
        failed_report["run_history_path"]
            .as_str()
            .context("failed conversion history path missing")?,
    );
    let failed_history_bytes = fs::read(&failed_history_path)?;
    let failed_history_text = String::from_utf8_lossy(&failed_history_bytes);
    assert!(!failed_history_text.contains(&failed_source.display().to_string()));
    assert!(!failed_history_text.contains("private-needs-sidecar.xls"));
    let failed_history: Value = serde_json::from_slice(&failed_history_bytes)?;
    assert_eq!(failed_history["status"], "failed");
    assert!(failed_history["events"].as_array().is_some_and(|events| {
        events.iter().any(|event| {
            event["status"] == "failed"
                && event["error_type"] == "external_engine_required"
                && event["actual_engine"].is_null()
                && event["input_index"] == 1
        })
    }));

    Ok(())
}

#[test]
fn sheet_rename_menu_executes_only_after_capability_and_confirmation() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let input = root.path().join("rename-confirmed.xlsx");
    create_source_workbook(&input, "OldSheet")?;
    let capability = root.path().join("passed-sheet-rename.json");
    write_sheet_rename_capability(&config, &input, &capability)?;
    let original = fs::read(&input)?;

    let result = run_sheet_rename_menu(&config, &capability, &input, "y")?;
    assert!(
        result.status.success(),
        "menu execution failed: {}",
        stderr(&result)
    );
    assert!(String::from_utf8_lossy(&result.stdout)
        .contains("副本保存到源文件旁的 PyTools_Sheet改名结果"));
    let report = menu_object_json(&result)?;
    assert_eq!(report["rename_ok"], 1);
    assert_eq!(report["files"][0]["status"], "renamed");
    assert_eq!(
        report["files"][0]["execution"]["engine"],
        "umya_spreadsheet"
    );

    // 4-5 改为就地修改：源工作簿被改写，不再生成结果目录副本。
    assert_ne!(
        fs::read(&input)?,
        original,
        "in-place sheet rename must rewrite the source workbook"
    );
    let output_path = PathBuf::from(
        report["files"][0]["output"]
            .as_str()
            .context("renamed workbook path missing")?,
    );
    assert!(output_path.is_file());
    assert_eq!(
        fs::canonicalize(&output_path)?,
        fs::canonicalize(&input)?,
        "output must be the source workbook itself"
    );
    let workbook = read_workbook(&output_path)?;
    let names = workbook
        .sheets
        .iter()
        .map(|sheet| sheet.name.as_str())
        .collect::<Vec<_>>();
    assert_eq!(names, ["Data", "NewSheet"]);
    assert_eq!(
        workbook.sheets[0]
            .values
            .get("A1")
            .and_then(|value| value.get("value"))
            .and_then(Value::as_str),
        Some("Bank")
    );
    Ok(())
}

#[test]
fn sheet_rename_menu_decline_keeps_the_source_and_writes_no_plan_or_copy() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let input = root.path().join("rename-declined.xlsx");
    create_source_workbook(&input, "OldSheet")?;
    let capability = root.path().join("passed-sheet-rename.json");
    write_sheet_rename_capability(&config, &input, &capability)?;
    let original = fs::read(&input)?;

    let result = run_sheet_rename_menu(&config, &capability, &input, "n")?;
    assert!(
        result.status.success(),
        "menu cancellation failed: {}",
        stderr(&result)
    );
    let stdout = String::from_utf8_lossy(&result.stdout);
    assert!(stdout.contains("[已取消] 未执行改名，也未生成计划或输出文件。"));
    assert_eq!(fs::read(&input)?, original);
    let plans = find_files_with_extension(&root.path().join("output"), "json")?;
    assert!(plans.is_empty());
    assert!(!root.path().join("PyTools_Sheet改名结果").exists());
    Ok(())
}

#[test]
fn sheet_rename_menu_without_capability_reports_blocker_and_writes_no_plans() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let input = root.path().join("rename-no-capability.xlsx");
    create_source_workbook(&input, "OldSheet")?;
    let original = fs::read(&input)?;

    let result = run_menu(&config, &format!("4\n5\n1\n{}\n0\n0\n", input.display()))?;
    assert!(result.status.success(), "menu failed: {}", stderr(&result));
    let stdout = String::from_utf8_lossy(&result.stdout);
    assert!(stdout.contains("[无法执行] 缺少已验收的 ExcelModifier capability/platform"));
    assert!(!stdout.contains("Sheet 改名计划="));
    assert_eq!(fs::read(&input)?, original);
    assert!(find_files_with_extension(&root.path().join("output"), "json")?.is_empty());
    assert!(!root.path().join("PyTools_Sheet改名结果").exists());
    Ok(())
}

#[test]
fn file_sort_menu_supports_user_developer_and_off_progress_views() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let sources = root.path().join("sources");
    fs::create_dir_all(&sources)?;
    let log_dir = root.path().join("logs");

    let user_source = sources.join("sort-user-payload.txt");
    fs::write(&user_source, b"user view")?;
    let user = run_action_with_log_view(&config, 9, 1, &user_source, "user")?;
    assert!(user.status.success(), "user view failed: {}", stderr(&user));
    assert!(sources.join("1_sort-user-payload.txt").is_file());
    let user_stderr = stderr(&user);
    let user_stderr = stderr(&user);
    assert!(user_stderr.contains("正在读取文件排序配置"));
    assert!(user_stderr.contains("已处理 1/1 个文件"));
    assert!(user_stderr.contains("阶段用时"));
    assert!(!user_stderr.contains(&user_source.display().to_string()));
    assert!(!user_stderr.contains("sort-user-payload"));
    assert!(String::from_utf8_lossy(&user.stdout).contains("重命名 1 个"));
    // user 视图同样保存统一运行记录（1 个 .run.json + 1 个 .jsonl 事件日志）。
    assert_eq!(fs::read_dir(&log_dir)?.count(), 2);

    let developer_source = sources.join("1_sort-developer-payload.txt");
    fs::write(&developer_source, b"developer view")?;
    let developer = run_action_with_log_view(&config, 10, 1, &developer_source, "developer")?;
    assert!(
        developer.status.success(),
        "developer view failed: {}",
        stderr(&developer)
    );
    assert!(sources.join("sort-developer-payload.txt").is_file());
    let developer_report = output_json(&developer)?;
    let history_path = PathBuf::from(
        developer_report["run_history_path"]
            .as_str()
            .context("developer run-history path missing")?,
    );
    let history_bytes = fs::read(&history_path)?;
    let history_text = String::from_utf8_lossy(&history_bytes);
    assert!(!history_text.contains(&developer_source.display().to_string()));
    assert!(!history_text.contains("sort-developer-payload"));
    let history: Value = serde_json::from_slice(&history_bytes)?;
    let events = history["events"].as_array().context("events missing")?;
    assert!(history["total_elapsed_ms"].is_number());
    assert!(events
        .iter()
        .any(|event| event["phase"] == "configuration_read"));
    assert!(events
        .iter()
        .any(|event| event["phase"] == "file_processing"));
    assert!(events.iter().any(|event| event["phase"] == "report_write"));
    assert!(events.iter().any(|event| event["phase"] == "task"));
    for engine in ["Calamine", "RustFilesystem", "rust_xlsxwriter"] {
        assert!(
            events.iter().any(|event| event["actual_engine"] == engine),
            "missing engine event for {engine}"
        );
    }
    assert!(events.iter().all(|event| event["elapsed_ms"].is_number()));
    let missing_stage_timing = events
        .iter()
        .filter(|event| event["status"] == "completed" || event["status"] == "failed")
        .filter(|event| !event["stage_elapsed_ms"].is_number())
        .collect::<Vec<_>>();
    assert!(
        missing_stage_timing.is_empty(),
        "events missing stage timings: {}",
        serde_json::to_string(&missing_stage_timing)?
    );
    assert!(stderr(&developer).contains("\"feature_id\":\"PT-4-10\""));

    let logs_before_off = fs::read_dir(&log_dir)?.count();
    let off_source = sources.join("sort-off-payload.txt");
    fs::write(&off_source, b"off view")?;
    let off = run_action_with_log_view(&config, 9, 1, &off_source, "off")?;
    assert!(off.status.success(), "off view failed: {}", stderr(&off));
    assert!(sources.join("1_sort-off-payload.txt").is_file());
    assert!(stderr(&off).is_empty());
    assert_eq!(fs::read_dir(&log_dir)?.count(), logs_before_off);

    Ok(())
}

#[test]
fn workbook_content_naming_menu_logs_user_developer_and_off_views() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let sources = root.path().join("sources");
    fs::create_dir_all(&sources)?;
    let log_dir = root.path().join("logs");

    let user_source = sources.join("content-user.xlsx");
    create_source_workbook(&user_source, "Plain")?;
    let user_source_bytes = fs::read(&user_source)?;
    let user = run_action_with_log_view(&config, 5, 1, &user_source, "user")?;
    assert!(user.status.success(), "user view failed: {}", stderr(&user));
    let user_target = sources.join("Bank_content-user.xlsx");
    assert_eq!(fs::read(&user_target)?, user_source_bytes);
    let user_stdout = String::from_utf8_lossy(&user.stdout);
    // user 视图：stdout 只输出人类可读汇总行，不输出 JSON。
    assert!(user_stdout.contains("处理结束：成功 1 个，跳过 0 个，失败 0 个"), "{user_stdout}");
    assert!(!user_stdout.contains("\"ok\""), "user view must not print JSON: {user_stdout}");
    assert!(!user_stdout.contains("run_history_path"), "user view must not reference a run history path: {user_stdout}");
    let user_stderr = stderr(&user);
    assert!(user_stderr.contains("正在读取内容前缀配置"));
    assert!(user_stderr.contains("已处理 1/1 个文件"));
    assert!(user_stderr.contains("阶段用时"));
    assert!(!user_stderr.contains(&user_source.display().to_string()));
    assert!(!user_stderr.contains("content-user.xlsx"));
    // user 视图同样保存统一运行记录（1 个 .run.json）。
    assert_eq!(find_files_with_extension(&log_dir, "json")?.len(), 1);

    let developer_source = sources.join("content-developer.xlsx");
    create_source_workbook(&developer_source, "Plain")?;
    let developer_source_bytes = fs::read(&developer_source)?;
    let developer = run_action_with_log_view(&config, 5, 1, &developer_source, "developer")?;
    assert!(
        developer.status.success(),
        "developer view failed: {}",
        stderr(&developer)
    );
    let developer_target = sources.join("Bank_content-developer.xlsx");
    assert_eq!(fs::read(&developer_target)?, developer_source_bytes);
    let developer_report = output_json(&developer)?;
    assert_eq!(developer_report["ok"], 1);
    let history_path = PathBuf::from(
        developer_report["run_history_path"]
            .as_str()
            .context("content-prefix history path missing")?,
    );
    let history_bytes = fs::read(&history_path)?;
    let history_text = String::from_utf8_lossy(&history_bytes);
    assert!(!history_text.contains(&developer_source.display().to_string()));
    assert!(!history_text.contains("Bank_content-developer.xlsx"));
    assert!(!history_text.contains("content-developer.xlsx"));
    let history: Value = serde_json::from_slice(&history_bytes)?;
    assert_eq!(history["feature_id"], "PT-4-6");
    assert_eq!(history["function"], "rename-by-content");
    assert_eq!(history["status"], "succeeded");
    assert!(history["total_elapsed_ms"].is_number());
    let events = history["events"].as_array().context("events missing")?;
    assert!(events
        .iter()
        .any(|event| event["phase"] == "configuration_read"));
    assert!(events
        .iter()
        .any(|event| event["phase"] == "file_processing"));
    assert!(events.iter().any(|event| event["phase"] == "task"));
    assert_eq!(events.last().unwrap()["completed_files"], 1);
    for engine in ["Calamine", "Calamine + RustFilesystem"] {
        assert!(
            history["actual_engines"]
                .as_array()
                .context("actual_engines missing")?
                .iter()
                .any(|actual| actual == engine),
            "missing engine {engine}"
        );
    }
    for phase in ["configuration_read", "file_processing", "task"] {
        assert!(events.iter().any(|event| {
            event["phase"] == phase
                && event["status"] == "completed"
                && event["stage_elapsed_ms"].is_number()
        }));
    }
    assert!(stderr(&developer).contains("\"feature_id\":\"PT-4-6\""));

    let history_count = fs::read_dir(&log_dir)?.count();
    let off_source = sources.join("content-off.xlsx");
    create_source_workbook(&off_source, "Plain")?;
    let off = run_action_with_log_view(&config, 5, 1, &off_source, "off")?;
    assert!(off.status.success(), "off view failed: {}", stderr(&off));
    assert!(sources.join("Bank_content-off.xlsx").is_file());
    assert!(stderr(&off).is_empty());
    assert_eq!(fs::read_dir(&log_dir)?.count(), history_count);
    assert!(output_json(&off)?.get("run_history_path").is_none());

    Ok(())
}

#[test]
fn v2_file_menu_hides_cancelled_local_mapping() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let output = run_menu(&config, "4\n0\n0\n")?;
    assert!(output.status.success(), "menu failed: {}", stderr(&output));
    let menu = String::from_utf8_lossy(&output.stdout);
    assert!(menu.contains("机构提交完整性检查"));
    assert!(!menu.contains("局部文件名映射"));
    assert!(!menu.contains("取消内容前缀"));
    Ok(())
}

#[test]
fn batch_rename_menu_supports_user_developer_and_off_progress_views() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let sources = root.path().join("sources");
    fs::create_dir_all(&sources)?;
    let log_dir = root.path().join("logs");

    let user_dir = sources.join("user");
    fs::create_dir_all(&user_dir)?;
    let user_source = user_dir.join("Bank-user-report.txt");
    fs::write(&user_source, b"user rename")?;
    let user = run_action_with_log_view(&config, 1, 1, &user_source, "user")?;
    assert!(user.status.success(), "user view failed: {}", stderr(&user));
    assert!(user_dir
        .join("2026-10-24 B01 Example Bank Monthly.txt")
        .is_file());
    let user_stderr = stderr(&user);
    assert!(user_stderr.contains("正在读取批量重命名配置"));
    assert!(user_stderr.contains("已处理 1/1 个文件或目录"));
    assert!(user_stderr.contains("阶段用时"));
    assert!(!user_stderr.contains(&user_source.display().to_string()));
    assert!(!user_stderr.contains("Bank-user-report"));
    let user_history_name = user_stderr.lines().find_map(|line| line.split_once("运行记录：").map(|(_, path)| path.trim())).context("unified rename summary missing")?;
    // user 视图仅显示可移植文件名；读取时需按配置的日志目录解析。
    let user_history_path = PathBuf::from(user_history_name);
    let user_history_path = if user_history_path.is_absolute() {
        user_history_path
    } else {
        log_dir.join(user_history_path)
    };
    let user_report: Value = serde_json::from_slice(&fs::read(&user_history_path)?)?;
    assert_eq!(user_report["input_stats"]["succeeded"], 1);
    assert_eq!(user_report["overall_status"], "SUCCESS");
    assert!(find_files_with_extension(&log_dir, "log")?.len() == 1);
    assert_eq!(find_files_with_extension(&log_dir, "json")?.len(), 1);

    let developer_dir = sources.join("developer");
    fs::create_dir_all(&developer_dir)?;
    let developer_source = developer_dir.join("Bank-developer-report.txt");
    fs::write(&developer_source, b"developer rename")?;
    let developer = run_action_with_log_view(&config, 1, 1, &developer_source, "developer")?;
    assert!(
        developer.status.success(),
        "developer view failed: {}",
        stderr(&developer)
    );
    let developer_report = output_json(&developer)?;
    assert_eq!(developer_report["ok"], 1);
    let history_path = PathBuf::from(
        developer_report["run_history_path"]
            .as_str()
            .context("developer run-history path missing")?,
    );
    let history_bytes = fs::read(&history_path)?;
    let history_text = String::from_utf8_lossy(&history_bytes);
    assert!(!history_text.contains(&developer_source.display().to_string()));
    assert!(!history_text.contains("Bank-developer-report"));
    let history: Value = serde_json::from_slice(&history_bytes)?;
    assert_eq!(history["feature_id"], "PT-4-1");
    assert!(history["total_elapsed_ms"].is_number());
    let events = history["events"].as_array().context("events missing")?;
    assert!(events
        .iter()
        .any(|event| event["phase"] == "configuration_read"));
    assert!(events
        .iter()
        .any(|event| event["phase"] == "file_processing"));
    assert!(events.iter().any(|event| event["phase"] == "task"));
    for engine in ["Calamine", "RustFilesystem"] {
        assert!(
            events.iter().any(|event| event["actual_engine"] == engine),
            "missing engine event for {engine}"
        );
    }
    assert!(events.iter().all(|event| {
        event["stage_elapsed_ms"].is_number() && event["total_elapsed_ms"].is_number()
    }));
    assert!(stderr(&developer).contains("\"feature_id\":\"PT-4-1\""));

    let json_count_before_off = find_files_with_extension(&log_dir, "json")?.len();
    let off_dir = sources.join("off");
    fs::create_dir_all(&off_dir)?;
    let off_source = off_dir.join("Bank-off-report.txt");
    fs::write(&off_source, b"off rename")?;
    let off = run_action_with_log_view(&config, 1, 1, &off_source, "off")?;
    assert!(off.status.success(), "off view failed: {}", stderr(&off));
    assert!(off_dir
        .join("2026-10-24 B01 Example Bank Monthly.txt")
        .is_file());
    assert!(stderr(&off).is_empty());
    assert_eq!(find_files_with_extension(&log_dir, "json")?.len(), json_count_before_off);

    Ok(())
}

#[test]
fn batch_rename_menu_persists_configuration_failure_log_without_touching_input() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config_without_rename_date(root.path())?;
    let source = root.path().join("Bank-preserve-report.txt");
    let source_bytes = b"keep source unchanged";
    fs::write(&source, source_bytes)?;

    let output = run_action_with_log_view(&config, 1, 1, &source, "user")?;
    assert!(!output.status.success());
    let user_message = stderr(&output);
    assert!(user_message.contains("批量重命名配置读取失败"));
    assert!(user_message.contains("缺少必要配置"));
    assert!(source.is_file());
    assert_eq!(fs::read(&source)?, source_bytes);

    let log_dir = root.path().join("logs");
    let logs = find_files_with_extension(&log_dir, "log")?;
    assert_eq!(logs.len(), 1);
    let log = fs::read_to_string(&logs[0])?;
    assert!(log.contains("[ERROR]"));
    assert!(log.contains("phase=configuration_read"));
    assert!(log.contains("error_type=configuration_error"));
    assert!(!log.contains(&source.display().to_string()));
    assert!(!log.contains("Bank-preserve-report"));
    Ok(())
}

fn create_config(root: &Path) -> Result<PathBuf> {
    create_config_with_rename_date(root, Some("2026-10-24"))
}

fn create_config_without_rename_date(root: &Path) -> Result<PathBuf> {
    create_config_with_rename_date(root, None)
}

fn create_config_with_rename_date(root: &Path, rename_date: Option<&str>) -> Result<PathBuf> {
    let config = root.join("config.xlsx");
    let output_dir = root.join("output");
    let log_dir = root.join("logs");
    fs::create_dir_all(&output_dir)?;
    fs::create_dir_all(&log_dir)?;
    let mut workbook = rust_xlsxwriter::Workbook::new();
    {
        let global = workbook.add_worksheet();
        global.set_name("全局配置")?;
        for (row, (key, value)) in [
            ("输入目录", root.to_string_lossy().into_owned()),
            ("输出目录", output_dir.to_string_lossy().into_owned()),
            ("日志目录", log_dir.to_string_lossy().into_owned()),
            ("excel目的格式", "xlsx".to_owned()),
            ("word目的格式", "docx".to_owned()),
        ]
        .iter()
        .enumerate()
        {
            global.write_string(row as u32, 0, *key)?;
            global.write_string(row as u32, 1, value)?;
        }
    }
    {
        let rename = workbook.add_worksheet();
        rename.set_name("重命名配置")?;
        for (column, header) in [
            (0, "简称"),
            (1, "全称"),
            (3, "代码"),
            (4, "全称(代码映射)"),
            (6, "键"),
            (7, "值"),
            (9, "原表名"),
            (10, "新表名"),
            (12, "原文件名片段"),
            (13, "新文件名片段"),
            (14, "内容重命名启用"),
            (15, "内容重命名取值规则"),
            (16, "备注"),
        ] {
            rename.write_string(0, column, header)?;
        }
        for (column, value) in [
            (0, "Bank"),
            (1, "Example Bank"),
            (3, "B01"),
            (4, "Example Bank"),
            (6, "数据日期"),
            (9, "OldSheet"),
            (10, "NewSheet"),
            (12, "before"),
            (13, "after"),
            (14, "是"),
            (15, "Data@A1"),
        ] {
            rename.write_string(1, column, value)?;
        }
        if let Some(rename_date) = rename_date {
            rename.write_string(1, 7, rename_date)?;
        }
        rename.write_string(2, 6, "报表名称")?;
        rename.write_string(2, 7, "Monthly")?;
    }
    {
        let submission = workbook.add_worksheet();
        submission.set_name("提交检查配置")?;
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
            submission.write_string(0, column as u16, *header)?;
        }
        let rules = [
            ["是", "2026", "Area", "Bank", "", "base rule"],
            ["是", "2026", "Area", "Bank", "multi", "second rule"],
            ["是", "2026", "Missing", "Bank", "", "missing rule"],
            ["否", "disabled", "Area", "Bank", "", "ignored rule"],
        ];
        for (row, values) in rules.iter().enumerate() {
            for (column, value) in values.iter().enumerate() {
                submission.write_string((row + 1) as u32, column as u16, *value)?;
            }
        }
    }
    {
        let sort = workbook.add_worksheet();
        sort.set_name("文件排序配置")?;
        for (column, header) in ["是否启用", "排序前缀", "匹配关键词", "备注"]
            .iter()
            .enumerate()
        {
            sort.write_string(0, column as u16, *header)?;
        }
        for (column, value) in ["是", "1", "sort", "synthetic"].iter().enumerate() {
            sort.write_string(1, column as u16, *value)?;
        }
    }
    workbook.save(&config)?;
    Ok(config)
}

fn create_source_workbook(path: &Path, second_sheet: &str) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let data = workbook.add_worksheet();
    data.set_name("Data")?;
    data.write_string(0, 0, "Bank")?;
    workbook.add_worksheet().set_name(second_sheet)?;
    workbook.save(path)?;
    Ok(())
}

fn write_sheet_rename_capability(config: &Path, input: &Path, path: &Path) -> Result<()> {
    let plan_output = input.with_file_name("sheet-rename-plan-output.xlsx");
    let plan = build_sheet_rename_plan(config, input, &plan_output)?
        .plan
        .context("synthetic workbook should produce a sheet rename plan")?;
    let capability = BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Passed,
        format: plan.format,
        operations: plan.operations.iter().map(EditOperation::kind).collect(),
        supported_features: plan.workbook.features_present.clone(),
        preserved: plan.preserve.clone(),
        platform: SHEET_RENAME_PLATFORM.to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: "synthetic-pt-4-4-menu".to_owned(),
            report_path: "pytools-core/tests/conversion_menu_cli.rs".to_owned(),
            adapter_version: "umya-spreadsheet-3.1.1".to_owned(),
            engine_version: "umya-spreadsheet-3.1.1".to_owned(),
            platform: SHEET_RENAME_PLATFORM.to_owned(),
        }),
    };
    fs::write(path, serde_json::to_vec_pretty(&capability)?)?;
    Ok(())
}

fn run_sheet_rename_menu(
    config: &Path,
    capability: &Path,
    input: &Path,
    confirmation: &str,
) -> Result<Output> {
    let mut child = menu_command(config);
    child
        .arg("--capabilities")
        .arg(capability)
        .arg("--platform")
        .arg(SHEET_RENAME_PLATFORM)
        .arg("--log-view")
        .arg("off");
    let menu_input = format!("4\n5\n1\n{}\n{confirmation}\n0\n0\n", input.display());
    finish_menu(child, &menu_input)
}

fn menu_input_for_action(action: u8, mode: u8, selected: &Path) -> String {
    let path = selected.display();
    match action {
        1 => format!("4\n1\n{mode}\n{path}\n0\n0\n"),
        2 => format!("4\n4\n1\n1\n{path}\n0\n0\n"),
        3 => format!("4\n4\n2\n1\n{path}\n0\n0\n"),
        4 => format!("4\n5\n1\n{path}\n0\n0\n"),
        5 => format!("4\n3\n{mode}\n{path}\n0\n0\n"),
        8 => format!("4\n6\n{mode}\n{path}\n0\n0\n"),
        9 => format!("4\n2\n1\n{mode}\n{path}\n0\n0\n"),
        10 => format!("4\n2\n2\n{mode}\n{path}\n0\n0\n"),
        _ => panic!("action {action} is not exposed by the V2 file menu"),
    }
}

fn run_action(config: &Path, action: u8, mode: u8, selected: &Path) -> Result<Output> {
    run_menu(config, &menu_input_for_action(action, mode, selected))
}

fn run_action_with_log_view(
    config: &Path,
    action: u8,
    mode: u8,
    selected: &Path,
    log_view: &str,
) -> Result<Output> {
    let input = menu_input_for_action(action, mode, selected);
    let mut child = menu_command(config);
    child.arg("--log-view").arg(log_view);
    finish_menu(child, &input)
}

fn run_action_with_engines_disabled(
    config: &Path,
    action: u8,
    mode: u8,
    selected: &Path,
) -> Result<Output> {
    let input = menu_input_for_action(action, mode, selected);
    let mut child = menu_command(config);
    child.env("PYTOOLS_CONVERT_ENGINES_DISABLED", "1");
    finish_menu(child, &input)
}

fn run_action_with_engines_disabled_and_log_view(
    config: &Path,
    action: u8,
    mode: u8,
    selected: &Path,
    log_view: &str,
) -> Result<Output> {
    let input = menu_input_for_action(action, mode, selected);
    let mut child = menu_command(config);
    child
        .env("PYTOOLS_CONVERT_ENGINES_DISABLED", "1")
        .arg("--log-view")
        .arg(log_view);
    finish_menu(child, &input)
}

fn run_convert_excel_cli(
    config: &Path,
    input: &Path,
    output_dir: &Path,
    engine: &str,
    log_view: &str,
) -> Result<Output> {
    Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("convert-excel")
        .arg("--config")
        .arg(config)
        .arg("--to")
        .arg("xlsx")
        .arg("--engine")
        .arg(engine)
        .arg("--output-dir")
        .arg(output_dir)
        .arg("--log-view")
        .arg(log_view)
        .arg(input)
        .output()
        .context("run direct Excel conversion CLI")
}

fn run_menu(config: &Path, input: &str) -> Result<Output> {
    finish_menu(menu_command(config), input)
}

fn menu_command(config: &Path) -> Command {
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"));
    child
        .arg("menu")
        .arg("--config")
        .arg(config)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped());
    child
}

fn finish_menu(mut child: Command, input: &str) -> Result<Output> {
    let mut child = child.spawn().context("start interactive pytools menu")?;
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

fn output_json(output: &Output) -> Result<Value> {
    let start = output
        .stdout
        .iter()
        .enumerate()
        .filter(|(_, byte)| matches!(byte, b'{' | b'['))
        .map(|(index, _)| index)
        .min()
        .context("CLI JSON output missing")?;
    let mut deserializer = serde_json::Deserializer::from_slice(&output.stdout[start..]);
    Value::deserialize(&mut deserializer).context("decode CLI JSON output")
}

fn menu_object_json(output: &Output) -> Result<Value> {
    let start = output
        .stdout
        .iter()
        .position(|byte| *byte == b'{')
        .context("menu JSON object missing")?;
    let mut deserializer = serde_json::Deserializer::from_slice(&output.stdout[start..]);
    Value::deserialize(&mut deserializer).context("decode menu JSON object")
}

fn find_files_with_extension(root: &Path, extension: &str) -> Result<Vec<PathBuf>> {
    let mut files = Vec::new();
    let mut pending = vec![root.to_path_buf()];
    while let Some(directory) = pending.pop() {
        for entry in fs::read_dir(directory)? {
            let path = entry?.path();
            if path.is_dir() {
                pending.push(path);
            } else if path.extension().and_then(|value| value.to_str()) == Some(extension) {
                files.push(path);
            }
        }
    }
    Ok(files)
}

fn stderr(output: &Output) -> String {
    String::from_utf8_lossy(&output.stderr).into_owned()
}
