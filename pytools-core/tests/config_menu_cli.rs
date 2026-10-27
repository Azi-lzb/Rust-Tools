use anyhow::{Context, Result};
use excel_editor::{EditOperation, ExcelEditPlan};
use excel_reader::read_workbook;
use pytools_core::config::load_global;
use rust_xlsxwriter::Workbook;
use serde_json::Value;
use std::{
    fs,
    io::{BufReader, Read, Write},
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
            "pytools-config-menu-{}-{nonce}-{serial}",
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
fn missing_configuration_bootstraps_through_menu_and_can_be_viewed() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = root.path().join("app/config/config.xlsx");
    let input = "2\n7\n2\n1\n1\n2\n0\n0\n";

    let output = run_menu(&config, input, "user")?;
    anyhow::ensure!(
        output.status.success(),
        "menu bootstrap failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("7  配置与快捷操作"));
    assert!(stdout.contains("当前配置不可用，该工具暂不能运行"));
    assert!(stdout.contains("请先选择 7-1 初始化/修复配置表"));
    assert!(stdout.contains("已新建配置表"));
    assert!(stdout.contains("重新加载并查看配置状态"));
    assert!(!stdout.contains("实际读取组件"));

    let workbook = read_workbook(&config)?;
    assert_eq!(workbook.sheets.len(), 11);
    assert_eq!(workbook.sheets[0].name, "全局配置");
    assert!(stdout.contains("源文件扩展名"));
    Ok(())
}

#[test]
fn reloading_configuration_updates_output_directory_for_the_running_menu() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = root.path().join("app/config/config.xlsx");
    let initial_output = root.path().join("initial-output");
    let reloaded_output = root.path().join("reloaded-output");
    write_global_config(&config, &initial_output)?;
    assert_eq!(load_global(&config)?.output_dir, initial_output);

    let source = root.path().join("summary-source.xlsx");
    let mut source_workbook = Workbook::new();
    let source_sheet = source_workbook.add_worksheet();
    source_sheet.set_name("Data")?;
    source_sheet.write_string(0, 0, "名称")?;
    source_sheet.write_string(0, 1, "金额")?;
    source_sheet.write_string(1, 0, "样本")?;
    source_sheet.write_number(1, 1, 12.5)?;
    source_workbook.save(&source)?;
    let source_before = fs::read(&source)?;

    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("menu")
        .arg("--config")
        .arg(&config)
        .arg("--log-view")
        .arg("off")
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::null())
        .spawn()
        .context("start PyTools menu for configuration reload test")?;
    let mut stdout = BufReader::new(child.stdout.take().context("menu stdout missing")?);
    let mut transcript = Vec::new();
    let mut transcript_cursor = 0;
    read_menu_until(
        &mut stdout,
        &mut transcript,
        &mut transcript_cursor,
        "主菜单选择> ",
    )?;
    send_menu_input(&mut child, "7\n")?;
    read_menu_until(
        &mut stdout,
        &mut transcript,
        &mut transcript_cursor,
        "子菜单选择> ",
    )?;

    // The process has already loaded the initial config before the main prompt.
    // Replace the workbook while it is running, then exercise the real reload menu.
    write_global_config(&config, &reloaded_output)?;
    assert_eq!(load_global(&config)?.output_dir, reloaded_output);
    send_menu_input(&mut child, "2\n")?;
    read_menu_until(
        &mut stdout,
        &mut transcript,
        &mut transcript_cursor,
        "子菜单选择> ",
    )?;
    send_menu_input(&mut child, "0\n")?;
    read_menu_until(
        &mut stdout,
        &mut transcript,
        &mut transcript_cursor,
        "主菜单选择> ",
    )?;
    fs::write(&config, b"not an xlsx workbook")?;
    send_menu_input(&mut child, "7\n")?;
    read_menu_until(
        &mut stdout,
        &mut transcript,
        &mut transcript_cursor,
        "子菜单选择> ",
    )?;
    send_menu_input(&mut child, "2\n")?;
    read_menu_until(
        &mut stdout,
        &mut transcript,
        &mut transcript_cursor,
        "子菜单选择> ",
    )?;
    send_menu_input(&mut child, "0\n")?;
    read_menu_until(
        &mut stdout,
        &mut transcript,
        &mut transcript_cursor,
        "主菜单选择> ",
    )?;
    let malformed_view_follow_up_start = transcript.len();
    send_menu_input(&mut child, "3\n0\n")?;
    read_menu_until(
        &mut stdout,
        &mut transcript,
        &mut transcript_cursor,
        "主菜单选择> ",
    )?;
    assert!(
        String::from_utf8_lossy(&transcript[malformed_view_follow_up_start..])
            .contains("------ 报表打印 ------")
    );

    write_config_without_global_sheet(&config)?;
    send_menu_input(&mut child, "7\n")?;
    read_menu_until(
        &mut stdout,
        &mut transcript,
        &mut transcript_cursor,
        "子菜单选择> ",
    )?;
    send_menu_input(&mut child, "2\n")?;
    read_menu_until(
        &mut stdout,
        &mut transcript,
        &mut transcript_cursor,
        "子菜单选择> ",
    )?;
    send_menu_input(&mut child, "0\n")?;
    read_menu_until(
        &mut stdout,
        &mut transcript,
        &mut transcript_cursor,
        "主菜单选择> ",
    )?;
    let invalid_global_follow_up_start = transcript.len();
    send_menu_input(&mut child, "3\n0\n")?;
    read_menu_until(
        &mut stdout,
        &mut transcript,
        &mut transcript_cursor,
        "主菜单选择> ",
    )?;
    assert!(
        String::from_utf8_lossy(&transcript[invalid_global_follow_up_start..])
            .contains("------ 报表打印 ------")
    );

    send_menu_input(&mut child, "1\n")?;
    read_menu_until(
        &mut stdout,
        &mut transcript,
        &mut transcript_cursor,
        "子菜单选择> ",
    )?;
    send_menu_input(&mut child, "2\n")?;
    read_menu_until(
        &mut stdout,
        &mut transcript,
        &mut transcript_cursor,
        "文件路径> ",
    )?;
    send_menu_input(&mut child, &format!("{}\n", source.display()))?;
    read_menu_until(
        &mut stdout,
        &mut transcript,
        &mut transcript_cursor,
        "主菜单选择> ",
    )?;
    send_menu_input(&mut child, "0\n")?;
    drop(child.stdin.take());
    stdout.read_to_end(&mut transcript)?;
    let status = child.wait().context("wait for reloaded PyTools menu")?;

    anyhow::ensure!(
        status.success(),
        "reloaded menu failed: {}\n{}",
        status,
        String::from_utf8_lossy(&transcript)
    );
    let workbook_count = |directory: &Path| -> Result<usize> {
        if !directory.exists() {
            return Ok(0);
        }
        Ok(fs::read_dir(directory)?
            .filter_map(|entry| entry.ok())
            .filter(|entry| {
                entry
                    .path()
                    .extension()
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
            })
            .count())
    };
    assert_eq!(workbook_count(&initial_output)?, 0);
    assert_eq!(
        workbook_count(&reloaded_output)?,
        1,
        "menu output: {}",
        String::from_utf8_lossy(&transcript)
    );
    assert_eq!(fs::read(&source)?, source_before);
    assert!(
        String::from_utf8_lossy(&transcript).contains(reloaded_output.to_string_lossy().as_ref())
    );
    assert!(String::from_utf8_lossy(&transcript).contains("[读取失败]"));
    assert!(String::from_utf8_lossy(&transcript).contains("继续使用上次成功加载的配置"));
    assert!(!String::from_utf8_lossy(&transcript).contains("[配置不可用]"));
    Ok(())
}

#[test]
fn config_initialize_or_repair_requires_explicit_second_confirmation() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = root.path().join("app/config/config.xlsx");
    let output = run_menu(&config, "7\n1\n0\n0\n0\n", "off")?;
    anyhow::ensure!(
        output.status.success(),
        "configuration cancellation failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("[二次确认] 7-1"));
    assert!(stdout.contains("[已取消] 配置文件未修改。"));
    assert!(!config.exists());
    assert!(!root.path().join("app/output").exists());
    assert!(!root.path().join("app/logs").exists());
    Ok(())
}

#[test]
fn default_menu_configuration_is_created_beside_executable_not_working_directory() -> Result<()> {
    let root = TestDirectory::new()?;
    let application_dir = root.path().join("installed-app");
    let launch_dir = root.path().join("launch-from-here");
    fs::create_dir_all(&application_dir)?;
    fs::create_dir_all(&launch_dir)?;

    let source_executable = Path::new(env!("CARGO_BIN_EXE_pytools"));
    let executable_name = source_executable
        .file_name()
        .context("PyTools executable has no file name")?;
    let copied_executable = application_dir.join(executable_name);
    fs::copy(source_executable, &copied_executable)
        .context("copy PyTools executable into temporary install layout")?;

    let mut child = Command::new(&copied_executable)
        .args(["menu", "--log-view", "off"])
        .current_dir(&launch_dir)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .context("start menu without an explicit configuration path")?;
    child
        .stdin
        .take()
        .context("default-path menu stdin missing")?
        .write_all(b"7\n1\n1\n0\n0\n")
        .context("initialize the default menu configuration")?;
    let output = child
        .wait_with_output()
        .context("wait for default-path menu")?;
    anyhow::ensure!(
        output.status.success(),
        "default-path menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let expected_config = application_dir.join("config/config.xlsx");
    assert!(expected_config.is_file());
    assert!(!launch_dir.join("config/config.xlsx").exists());
    let workbook = read_workbook(&expected_config)?;
    assert_eq!(workbook.sheets.len(), 11);
    assert_eq!(workbook.sheets[0].name, "全局配置");
    Ok(())
}

#[test]
fn existing_configuration_repair_menu_saves_a_plan_without_modifying_source() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = root.path().join("app/config/config.xlsx");
    write_incomplete_config(&config)?;
    let source_before = fs::read(&config)?;

    let output = run_menu(&config, "7\n1\n1\n0\n0\n", "developer")?;
    anyhow::ensure!(
        output.status.success(),
        "configuration repair menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("[计划已生成]"));
    assert!(stdout.contains("源配置未修改，也未生成候选"));
    assert!(stdout.contains("实际读取组件: excel-reader / Calamine"));
    assert!(stdout.contains("修改引擎: 未执行"));

    let output_root = root.path().join("app/output");
    let action_dir = fs::read_dir(&output_root)?
        .map(|entry| entry.map(|entry| entry.path()))
        .collect::<std::io::Result<Vec<_>>>()?
        .into_iter()
        .find(|path| {
            path.file_name()
                .is_some_and(|name| name.to_string_lossy().starts_with("pytools-config-repair-"))
        })
        .context("configuration repair output directory missing")?;
    let plan_path = action_dir.join("repair-plan.json");
    let plan: ExcelEditPlan = serde_json::from_slice(&fs::read(&plan_path)?)?;
    assert!(!plan.operations.is_empty());
    assert!(plan
        .operations
        .iter()
        .all(|operation| !matches!(operation, EditOperation::DeleteSheet { .. })));
    assert!(!action_dir.join("config-candidate.xlsx").exists());
    assert_eq!(fs::read(&config)?, source_before);

    let history_path = fs::read_dir(root.path().join("app/logs"))?
        .filter_map(|entry| entry.ok())
        .map(|entry| entry.path())
        .find(|path| {
            path.file_name()
                .and_then(|name| name.to_str())
                .is_some_and(|name| name.starts_with("pytools-config-initialize_or_repair-"))
        })
        .context("configuration repair run history missing")?;
    let history_text = fs::read_to_string(&history_path)?;
    assert!(!history_text.contains(&config.to_string_lossy().to_string()));
    let history: Value = serde_json::from_str(&history_text)?;
    assert_eq!(history["feature_id"], "PT-5-1");
    assert_eq!(history["status"], "succeeded");
    assert_eq!(history["metrics"]["operation"], "repair_plan");
    assert!(
        history["metrics"]["planned_operations"]
            .as_u64()
            .unwrap_or_default()
            > 0
    );
    assert_eq!(history["plan_only"], true);
    assert_eq!(history["excel_modifier_invoked"], false);
    assert!(history["actual_engines"]
        .as_array()
        .is_some_and(|engines| engines.iter().any(|engine| engine == "calamine")));
    Ok(())
}

#[test]
fn ambiguous_configuration_headers_fail_closed_without_creating_a_repair_plan() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = root.path().join("app/config/config.xlsx");
    write_ambiguous_config(&config)?;
    let source_before = fs::read(&config)?;

    let output = run_menu(&config, "7\n1\n1\n0\n0\n", "developer")?;
    anyhow::ensure!(
        output.status.success(),
        "ambiguous configuration repair menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(
        stdout.contains("configuration header layout is ambiguous"),
        "{stdout}"
    );
    assert_eq!(fs::read(&config)?, source_before);
    assert!(!root.path().join("app/output").exists());

    let history_path = fs::read_dir(root.path().join("app/logs"))?
        .filter_map(|entry| entry.ok())
        .map(|entry| entry.path())
        .find(|path| {
            path.file_name()
                .and_then(|name| name.to_str())
                .is_some_and(|name| name.starts_with("pytools-config-initialize_or_repair-"))
        })
        .context("ambiguous configuration repair history missing")?;
    let history: Value = serde_json::from_slice(&fs::read(history_path)?)?;
    assert_eq!(history["feature_id"], "PT-5-1");
    assert_eq!(history["status"], "failed");
    assert_eq!(history["error_type"], "config_initialize_or_repair_failed");
    Ok(())
}

#[test]
fn config_menu_history_records_writer_and_view_without_config_contents() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = root.path().join("app/config/config.xlsx");

    let initialized = run_menu(&config, "7\n1\n1\n0\n0\n", "developer")?;
    anyhow::ensure!(
        initialized.status.success(),
        "developer config initialization failed: {}",
        String::from_utf8_lossy(&initialized.stderr)
    );
    let initialized_stdout = String::from_utf8_lossy(&initialized.stdout);
    assert!(initialized_stdout.contains("PT-5 完成：PT-5-1"));
    let log_dir = root.path().join("app/logs");
    let init_history_path = fs::read_dir(&log_dir)?
        .filter_map(|entry| entry.ok())
        .map(|entry| entry.path())
        .find(|path| {
            path.file_name()
                .and_then(|name| name.to_str())
                .is_some_and(|name| name.starts_with("pytools-config-initialize_or_repair-"))
        })
        .context("config initialization history missing")?;
    let init_history_text = fs::read_to_string(&init_history_path)?;
    assert!(!init_history_text.contains(&config.to_string_lossy().to_string()));
    let init_history: Value = serde_json::from_str(&init_history_text)?;
    assert_eq!(init_history["metrics"]["operation"], "initialized");
    assert_eq!(init_history["metrics"]["sheets"], 11);
    assert_eq!(init_history["actual_engines"][0], "rust_xlsxwriter");
    assert_eq!(init_history["excel_writer_invoked"], true);
    assert_eq!(init_history["excel_modifier_invoked"], false);

    let viewed = run_menu(&config, "7\n2\n0\n0\n", "developer")?;
    anyhow::ensure!(
        viewed.status.success(),
        "developer config view failed: {}",
        String::from_utf8_lossy(&viewed.stderr)
    );
    let view_history_path = fs::read_dir(&log_dir)?
        .filter_map(|entry| entry.ok())
        .map(|entry| entry.path())
        .find(|path| {
            path.file_name()
                .and_then(|name| name.to_str())
                .is_some_and(|name| name.starts_with("pytools-config-view-"))
        })
        .context("config view history missing")?;
    let view_history_text = fs::read_to_string(view_history_path)?;
    assert!(!view_history_text.contains(&config.to_string_lossy().to_string()));
    assert!(!view_history_text.contains("原始配置视图"));
    let view_history: Value = serde_json::from_str(&view_history_text)?;
    assert_eq!(view_history["feature_id"], "PT-5-2");
    assert_eq!(view_history["status"], "succeeded");
    assert_eq!(view_history["metrics"]["sheet_count"], 11);
    assert!(view_history["actual_engines"]
        .as_array()
        .is_some_and(|engines| engines.iter().any(|engine| engine == "calamine")));
    assert!(view_history["elapsed_ms"].as_u64().is_some());
    Ok(())
}

#[test]
fn config_view_off_keeps_requested_view_but_suppresses_progress_and_history() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = root.path().join("app/config/config.xlsx");
    write_incomplete_config(&config)?;

    let output = run_menu(&config, "7\n2\n0\n0\n", "off")?;
    anyhow::ensure!(
        output.status.success(),
        "off config view failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("配置文件:"));
    assert!(stdout.contains("Sheet 列表"));
    assert!(!stdout.contains("PT-5 开始："));
    assert!(!stdout.contains("PT-5 完成："));
    assert!(!stdout.contains("开始处理："));
    assert!(!stdout.contains("处理完成："));
    assert!(!stdout.contains("用时="));
    assert!(!root.path().join("app/logs").exists());
    Ok(())
}

#[test]
fn malformed_configuration_repair_logs_failure_without_config_contents() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = root.path().join("app/config/config.xlsx");
    fs::create_dir_all(config.parent().context("configuration parent missing")?)?;
    fs::write(&config, b"not a workbook")?;

    let output = run_menu(&config, "7\n1\n1\n0\n0\n", "developer")?;
    anyhow::ensure!(
        output.status.success(),
        "developer config repair failure path exited unexpectedly: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let log_dir = root.path().join("app/logs");
    let history_path = fs::read_dir(&log_dir)?
        .filter_map(|entry| entry.ok())
        .map(|entry| entry.path())
        .find(|path| {
            path.file_name()
                .and_then(|name| name.to_str())
                .is_some_and(|name| name.starts_with("pytools-config-initialize_or_repair-"))
        })
        .context("failed configuration repair history missing")?;
    let history_text = fs::read_to_string(history_path)?;
    assert!(!history_text.contains(&config.to_string_lossy().to_string()));
    assert!(!history_text.contains("not a workbook"));
    let history: Value = serde_json::from_str(&history_text)?;
    assert_eq!(history["feature_id"], "PT-5-1");
    assert_eq!(history["status"], "failed");
    assert_eq!(history["error_type"], "config_initialize_or_repair_failed");
    assert!(history["actual_components"]
        .as_array()
        .is_some_and(|components| components
            .iter()
            .any(|component| component == "pytools_core::config::build_config_repair_plan")));
    Ok(())
}

fn write_incomplete_config(path: &Path) -> Result<()> {
    fs::create_dir_all(path.parent().context("configuration parent missing")?)?;
    let mut workbook = Workbook::new();
    let global = workbook.add_worksheet();
    global.set_name("全局配置")?;
    global.write_string(0, 0, "旧键头")?;
    global.write_string(0, 1, "旧值头")?;
    for (row, (key, value)) in [
        ("输入目录", "input"),
        ("输出目录", "output"),
        ("日志目录", "logs"),
    ]
    .iter()
    .enumerate()
    {
        global.write_string((row + 1) as u32, 0, *key)?;
        global.write_string((row + 1) as u32, 1, *value)?;
    }
    for obsolete in ["校验规则生成配置", "校验规则映射表"] {
        let sheet = workbook.add_worksheet();
        sheet.set_name(obsolete)?;
        sheet.write_string(1, 0, "existing user detail")?;
    }
    workbook.save(path)?;
    Ok(())
}

fn write_global_config(path: &Path, output_dir: &Path) -> Result<()> {
    fs::create_dir_all(path.parent().context("configuration parent missing")?)?;
    let mut workbook = Workbook::new();
    let global = workbook.add_worksheet();
    global.set_name("全局配置")?;
    global.write_string(0, 0, "键")?;
    global.write_string(0, 1, "值")?;
    global.write_string(0, 2, "备注")?;
    global.write_string(1, 0, "输出目录")?;
    global.write_string(1, 1, output_dir.to_string_lossy())?;
    global.write_string(2, 0, "日志目录")?;
    global.write_string(
        2,
        1,
        path.parent()
            .context("configuration parent missing")?
            .join("logs")
            .to_string_lossy(),
    )?;
    workbook.save(path)?;
    Ok(())
}

fn write_config_without_global_sheet(path: &Path) -> Result<()> {
    fs::create_dir_all(path.parent().context("configuration parent missing")?)?;
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("其他配置")?;
    sheet.write_string(0, 0, "键")?;
    sheet.write_string(0, 1, "值")?;
    workbook.save(path)?;
    Ok(())
}

fn send_menu_input(child: &mut std::process::Child, input: &str) -> Result<()> {
    let stdin = child.stdin.as_mut().context("interactive stdin missing")?;
    stdin.write_all(input.as_bytes())?;
    stdin.flush()?;
    Ok(())
}

fn read_menu_until<R: Read>(
    reader: &mut R,
    transcript: &mut Vec<u8>,
    cursor: &mut usize,
    marker: &str,
) -> Result<()> {
    let marker = marker.as_bytes();
    loop {
        if transcript[*cursor..]
            .windows(marker.len())
            .any(|window| window == marker)
        {
            *cursor = transcript.len();
            return Ok(());
        }
        let mut byte = [0_u8; 1];
        anyhow::ensure!(
            reader.read(&mut byte)? != 0,
            "interactive menu exited before `{}`; output: {}",
            String::from_utf8_lossy(marker),
            String::from_utf8_lossy(transcript)
        );
        transcript.push(byte[0]);
    }
}

fn write_ambiguous_config(path: &Path) -> Result<()> {
    fs::create_dir_all(path.parent().context("configuration parent missing")?)?;
    let mut workbook = Workbook::new();
    let global = workbook.add_worksheet();
    global.set_name("全局配置")?;
    global.write_string(0, 0, "值")?;
    global.write_string(0, 1, "键")?;
    global.write_string(0, 2, "备注")?;
    global.write_string(1, 0, "输入目录")?;
    global.write_string(1, 1, "existing path")?;
    workbook.save(path)?;
    Ok(())
}

fn run_menu(config: &Path, input: &str, log_view: &str) -> Result<Output> {
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
        .context("start interactive configuration menu")?;
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
