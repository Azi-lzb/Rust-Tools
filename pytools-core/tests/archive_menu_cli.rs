use anyhow::{Context, Result};
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
            "pytools-archive-menu-{}-{nonce}-{serial}",
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
fn interactive_archive_menu_routes_all_five_workflows() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_app_config(root.path())?;
    let log_dir = root.path().join("app").join("logs");

    let date_source = root
        .path()
        .join("date-preview")
        .join("20261025_存款表.xlsx");
    write_source(&date_source, b"date preview source")?;
    let date_target = root.path().join("date-preview-target");
    fs::create_dir_all(&date_target)?;
    let date_input = format!(
        "6\n1\n1\n{}\n{}\n0\n",
        date_source.display(),
        date_target.display()
    );
    let date_preview = run_menu(&config, "user", &date_input)?;
    assert_menu_success(&date_preview, "PT-8-1")?;
    assert!(String::from_utf8_lossy(&date_preview.stdout).contains("6  归档整理"));
    assert!(String::from_utf8_lossy(&date_preview.stdout).contains("按日期归档预览"));
    assert!(String::from_utf8_lossy(&date_preview.stdout).contains("用时="));
    assert_eq!(fs::read(&date_source)?, b"date preview source");
    let date_history = read_history(&log_dir, "PT-8-1_")?;
    assert_eq!(date_history["status"], "succeeded");
    assert_eq!(date_history["result"]["files"], 1);
    assert_eq!(date_history["result"]["copied"], 0);
    let unified_path = log_dir.join("run-unified").join(format!("PT-8-1_{}.json", date_history["run_id"].as_str().context("archive run ID missing")?));
    let unified: Value = serde_json::from_slice(&fs::read(unified_path)?)?;
    assert_eq!(unified["run_id"], date_history["run_id"]);
    assert_eq!(unified["input_stats"]["selected"], 1);
    assert_eq!(unified["input_stats"]["succeeded"], 1);
    let lines = fs::read_to_string(log_dir.join(unified["diagnostic_log_path"].as_str().context("archive diagnostics missing")?))?;
    let events = lines.lines().map(serde_json::from_str::<Value>).collect::<std::result::Result<Vec<_>, _>>()?;
    assert!(events.iter().any(|event| event["event_name"] == "file_result" && event["run_id"] == unified["run_id"]));
    assert!(date_history["total_elapsed_ms"].is_number());
    assert!(date_history["events"][1]["stage_elapsed_ms"].is_number());
    assert_history_sanitized(&date_history, &date_source);

    let type_folder = root.path().join("type-preview-folder");
    let nested_type_source = type_folder.join("nested").join("20261025_贷款表.xlsx");
    write_source(&nested_type_source, b"type preview source")?;
    let type_target = root.path().join("type-preview-target");
    fs::create_dir_all(&type_target)?;
    let type_input = format!(
        "6\n2\n3\n{}\n{}\n0\n",
        type_folder.display(),
        type_target.display()
    );
    let type_preview = run_menu(&config, "developer", &type_input)?;
    assert_menu_success(&type_preview, "PT-8-2")?;
    let type_history = read_history(&log_dir, "PT-8-2_")?;
    assert_eq!(type_history["status"], "succeeded");
    assert_eq!(type_history["result"]["files"], 1);
    assert_eq!(type_history["result"]["copied"], 0);
    assert!(type_history["actual_engines"]
        .as_array()
        .is_some_and(|engines| engines
            .iter()
            .any(|engine| engine == "excel-reader/Calamine")));
    assert!(String::from_utf8_lossy(&type_preview.stderr)
        .lines()
        .filter_map(|line| serde_json::from_str::<Value>(line).ok())
        .any(|event| event["phase"] == "archive_operation" && event["status"] == "succeeded"));
    assert_history_sanitized(&type_history, &nested_type_source);
    assert!(!type_target.join("贷款").exists());

    let plan_source = root
        .path()
        .join("plan-source")
        .join("20261025_计划输入.xlsx");
    write_source(&plan_source, b"plan execute source")?;
    let plan_target = root.path().join("plan-target");
    fs::create_dir_all(&plan_target)?;
    let plan_output = root.path().join("plan-output");
    let plan = pytools_core::archive_ops::archive_by_date_preview(
        std::slice::from_ref(&plan_source),
        &plan_target,
        &plan_output,
        Some(&config),
    )?
    .output;
    let histories_before_off = fs::read_dir(&log_dir)?.count();
    let plan_input = format!("6\n3\n{}\n0\n", plan.display());
    let plan_execute = run_menu(&config, "off", &plan_input)?;
    assert_menu_success(&plan_execute, "PT-8-3")?;
    assert_eq!(
        fs::read(plan_target.join("202610").join("20261025_计划输入.xlsx"))?,
        b"plan execute source"
    );
    assert_eq!(fs::read_dir(&log_dir)?.count(), histories_before_off);
    assert!(String::from_utf8_lossy(&plan_execute.stdout).contains("执行归档计划"));

    let date_copy_source = root.path().join("date-copy").join("20261025_直接复制.xlsx");
    write_source(&date_copy_source, b"date copy source")?;
    let date_copy_target = root.path().join("date-copy-target");
    fs::create_dir_all(&date_copy_target)?;
    let date_copy_input = format!(
        "6\n4\n1\n{}\n{}\n0\n",
        date_copy_source.display(),
        date_copy_target.display()
    );
    let date_copy = run_menu(&config, "developer", &date_copy_input)?;
    assert_menu_success(&date_copy, "PT-8-4")?;
    assert_eq!(
        fs::read(
            date_copy_target
                .join("202610")
                .join("20261025_直接复制.xlsx")
        )?,
        b"date copy source"
    );
    let date_copy_history = read_history(&log_dir, "PT-8-4_")?;
    assert_eq!(date_copy_history["result"]["copied"], 1);
    assert!(date_copy_history["actual_engines"]
        .as_array()
        .is_some_and(|engines| { engines.iter().any(|engine| engine == "Rust std::fs::copy") }));

    let type_copy_source = root.path().join("type-copy").join("20261025_贷款表.xlsx");
    write_source(&type_copy_source, b"type copy source")?;
    let type_copy_target = root.path().join("type-copy-target");
    fs::create_dir_all(&type_copy_target)?;
    let type_copy_input = format!(
        "6\n5\n1\n{}\n{}\n0\n",
        type_copy_source.display(),
        type_copy_target.display()
    );
    let type_copy = run_menu(&config, "user", &type_copy_input)?;
    assert_menu_success(&type_copy, "PT-8-5")?;
    assert_eq!(
        fs::read(type_copy_target.join("贷款").join("20261025_贷款表.xlsx"))?,
        b"type copy source"
    );
    let type_copy_history = read_history(&log_dir, "PT-8-5_")?;
    assert_eq!(type_copy_history["result"]["copied"], 1);
    assert_history_sanitized(&type_copy_history, &type_copy_source);

    Ok(())
}

#[test]
fn archive_copy_menu_supports_current_level_and_recursive_folder_selection() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_app_config(root.path())?;
    let source_folder = root.path().join("selection-source");
    let top_level = source_folder.join("20261025_贷款根表.xlsx");
    let nested = source_folder.join("nested").join("20261025_贷款子表.xlsx");
    write_source(&top_level, b"top-level source")?;
    write_source(&nested, b"nested source")?;

    let date_target = root.path().join("date-current-target");
    fs::create_dir_all(&date_target)?;
    let date_input = format!(
        "6\n4\n2\n{}\n{}\n0\n",
        source_folder.display(),
        date_target.display()
    );
    let date_copy = run_menu(&config, "user", &date_input)?;
    assert_menu_success(&date_copy, "PT-8-4")?;
    let stdout = String::from_utf8_lossy(&date_copy.stdout);
    assert!(stdout.contains("2  选择文件夹（仅当前层）"));
    assert!(stdout.contains("3  选择文件夹（递归子目录）"));
    assert!(date_target
        .join("202610")
        .join("20261025_贷款根表.xlsx")
        .is_file());
    assert!(!date_target
        .join("202610")
        .join("20261025_贷款子表.xlsx")
        .exists());

    let type_target = root.path().join("type-recursive-target");
    fs::create_dir_all(&type_target)?;
    let type_input = format!(
        "6\n5\n3\n{}\n{}\n0\n",
        source_folder.display(),
        type_target.display()
    );
    let type_copy = run_menu(&config, "user", &type_input)?;
    assert_menu_success(&type_copy, "PT-8-5")?;
    assert!(type_target
        .join("贷款")
        .join("20261025_贷款根表.xlsx")
        .is_file());
    assert!(type_target
        .join("贷款")
        .join("20261025_贷款子表.xlsx")
        .is_file());
    assert_eq!(fs::read(&top_level)?, b"top-level source");
    assert_eq!(fs::read(&nested)?, b"nested source");
    Ok(())
}

fn create_app_config(root: &Path) -> Result<PathBuf> {
    let app_dir = root.join("app");
    let config_dir = app_dir.join("config");
    fs::create_dir_all(&config_dir)?;
    fs::create_dir_all(app_dir.join("output"))?;
    fs::create_dir_all(app_dir.join("logs"))?;
    let config = config_dir.join("config.xlsx");

    let mut workbook = rust_xlsxwriter::Workbook::new();
    let global = workbook.add_worksheet();
    // 加载器契约（config::load_global）：键在 A 列、值在 B 列；此前用
    // 序号/配置项/配置值三列布局写出的键值从未被读到，断言只因旧的
    // "logs" 默认值与夹具路径巧合一致而通过。
    global.set_name("全局配置")?;
    global.write_string(0, 0, "键")?;
    global.write_string(0, 1, "值")?;
    global.write_string(1, 0, "输出目录")?;
    global.write_string(1, 1, app_dir.join("output").to_string_lossy())?;
    global.write_string(2, 0, "日志目录")?;
    global.write_string(2, 1, app_dir.join("logs").to_string_lossy())?;

    let archive = workbook.add_worksheet();
    archive.set_name("归档类型配置")?;
    for (column, value) in ["是否启用", "匹配关键词", "归档文件夹", "备注"]
        .iter()
        .enumerate()
    {
        archive.write_string(0, column as u16, *value)?;
    }
    archive.write_string(0, 4, "包括后缀")?;
    archive.write_string(0, 5, "排除后缀")?;
    archive.write_string(1, 0, "是")?;
    archive.write_string(1, 1, "贷款")?;
    archive.write_string(1, 2, "贷款")?;
    archive.write_string(1, 4, "xlsx")?;
    archive.write_string(1, 5, "docx")?;
    workbook.save(&config).context("save archive menu config")?;
    Ok(config)
}

fn write_source(path: &Path, contents: &[u8]) -> Result<()> {
    fs::create_dir_all(path.parent().context("source parent directory missing")?)?;
    fs::write(path, contents)?;
    Ok(())
}

fn run_menu(config: &Path, log_view: &str, input: &str) -> Result<Output> {
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
        .context("menu stdin was not piped")?
        .write_all(input.as_bytes())
        .context("write archive menu input")?;
    child.wait_with_output().context("wait for archive menu")
}

fn assert_menu_success(output: &Output, feature_id: &str) -> Result<()> {
    assert!(
        output.status.success(),
        "{feature_id} menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stdout.contains("6  归档整理"));
    assert!(
        stdout.contains(feature_id) || stderr.contains(feature_id) || feature_id == "PT-8-3",
        "{feature_id} run identifier missing from menu output"
    );
    Ok(())
}

fn read_history(log_dir: &Path, prefix: &str) -> Result<Value> {
    let path = fs::read_dir(log_dir)?
        .filter_map(std::result::Result::ok)
        .map(|entry| entry.path())
        .find(|path| {
            path.file_name()
                .is_some_and(|name| name.to_string_lossy().starts_with(prefix))
        })
        .with_context(|| format!("{prefix} archive history was not written"))?;
    let bytes = fs::read(path)?;
    serde_json::from_slice(&bytes).context("parse archive history")
}

fn assert_history_sanitized(history: &Value, source: &Path) {
    let text = serde_json::to_string(history).unwrap();
    assert!(!text.contains(&source.to_string_lossy().to_string()));
    assert!(!text.contains(
        source
            .file_name()
            .unwrap_or_default()
            .to_string_lossy()
            .as_ref()
    ));
}
