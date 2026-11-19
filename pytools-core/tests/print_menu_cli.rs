//! 报表打印菜单端到端测试，覆盖 3-1 的两种引擎和 3-3 预校验。

#[path = "common/logging_evidence.rs"]
mod logging_evidence;

use anyhow::{Context, Result};
use excel_reader::read_workbook;
use serde::Deserialize;
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
            "pytools-print-menu-{}-{nonce}-{serial}",
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
fn interactive_print_menu_uses_openpyxl_and_precheck() -> Result<()> {
    let root = TestDirectory::new()?;
    let source = root.path().join("source.xlsx");
    let target = root.path().join("printed.xlsx");
    let config =
        create_print_config_with_settings(root.path(), &source, &target, "1", "B2:D4", "是", 1)?;
    create_print_source(&source)?;
    let source_before = fs::read(&source)?;

    let menu = run_menu(&config, "3\n0\n0\n")?;
    assert!(menu.status.success());
    let menu_text = String::from_utf8_lossy(&menu.stdout);
    assert!(
        menu_text.contains("1  按打印配置生成（选择执行模式）"),
        "{menu_text}"
    );
    assert!(!menu_text.contains("2  按打印配置生成"), "{menu_text}");
    assert!(menu_text.contains("3  打印配置预校验"), "{menu_text}");
    assert!(!target.exists());
    assert!(!menu_text.contains("PDF"), "PDF entry must stay removed");
    assert!(
        !menu_text.contains("按批注"),
        "comment print entries are removed"
    );

    // 3-1 保留原菜单编号，实际使用 openpyxl，直接执行且无 y 确认。
    let print_run = run_menu(&config, "3\n1\n1\n0\n0\n")?;
    assert!(
        print_run.status.success(),
        "openpyxl print failed: {}",
        String::from_utf8_lossy(&print_run.stderr)
    );
    let print_stdout = String::from_utf8_lossy(&print_run.stdout);
    assert!(!print_stdout.contains("输入 y/yes/是"), "{print_stdout}");
    assert!(
        print_stdout.contains("按配置生成静态打印文件：执行成功")
            && print_stdout.contains("生成：工作簿 1 个"),
        "{print_stdout}"
    );
    assert!(!print_stdout.contains("开始按打印配置生成静态打印文件"));
    let done_line = print_stdout
        .lines()
        .find(|line| line.contains("按配置生成静态打印文件：执行成功"))
        .context("done line missing")?;
    assert!(starts_with_timestamp(done_line), "{done_line}");
    let history_path = fs::read_dir(root.path().join("logs"))?
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .find(|path| {
            path.file_name().is_some_and(|name| {
                let name = name.to_string_lossy();
                name.starts_with("3-1_") && name.ends_with(".run.json")
            })
        })
        .context("print run summary missing")?;
    let history: Value = serde_json::from_slice(&fs::read(history_path)?)?;
    assert_eq!(history["overall_status"], "SUCCESS");
    assert_eq!(history["task_stats"]["succeeded"], 1);
    assert_eq!(history["task_stats"]["failed"], 0);
    assert_eq!(history["input_stats_available"], false);
    let events_path = root.path().join("logs").join(
        history["diagnostic_log_path"]
            .as_str()
            .context("print diagnostic log missing")?,
    );
    let events = fs::read_to_string(events_path)?;
    let events = events
        .lines()
        .map(serde_json::from_str::<Value>)
        .collect::<std::result::Result<Vec<_>, _>>()?;
    assert!(events
        .iter()
        .all(|event| event["run_id"] == history["run_id"]));
    assert!(events
        .iter()
        .any(|event| event["event_name"] == "run_finished"));
    let printed = target.clone();
    assert!(printed.is_file(), "openpyxl print output missing");
    assert_eq!(
        fs::read(&source)?,
        source_before,
        "source must stay unchanged"
    );

    // 3-3：配置预校验（机器可读 JSON 仅 developer 视图输出）。
    let precheck = run_menu_with_view(&config, "3\n3\n0\n0\n", "developer")?;
    assert!(precheck.status.success());
    let precheck_response = output_json(&precheck)?;
    assert_eq!(
        precheck_response["failed"], 1,
        "existing print target must fail precheck"
    );
    assert_eq!(precheck_response["ok"], 0);
    Ok(())
}

#[test]
fn configured_static_print_failure_shows_reason_and_returns_to_menu() -> Result<()> {
    let root = TestDirectory::new()?;
    let source = root.path().join("source.xlsx");
    // 目标扩展名不受支持，引擎必须明确失败且不产出文件。
    let target = root.path().join("printed.xls");
    let config = create_print_config(root.path(), &source, &target)?;
    create_print_source(&source)?;
    let source_before = fs::read(&source)?;

    let output = run_menu(&config, "3\n1\n1\n0\n0\n")?;
    assert!(
        output.status.success(),
        "menu must stay usable after a failure"
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    // 分组摘要：状态行「…：执行失败」+ 需要关注里的错误明细（保留中断原因）。
    let failure_line = stdout
        .lines()
        .find(|line| line.contains("按配置生成静态打印文件：执行失败"))
        .context("failure status header missing")?;
    assert!(
        stdout.contains("打印任务中断"),
        "error detail must name the interruption: {stdout}"
    );
    assert!(
        starts_with_timestamp(failure_line),
        "failure line must carry a timestamp: {failure_line}"
    );
    assert!(stdout.contains(".xlsx"), "{stdout}");
    assert!(
        stdout.contains("======== rustTools 菜单 ========"),
        "{stdout}"
    );
    assert_eq!(fs::read(&source)?, source_before);
    assert!(!target.exists(), "failed print must not create a target");
    Ok(())
}

#[test]
fn interactive_static_print_executes_without_confirmation() -> Result<()> {
    let root = TestDirectory::new()?;
    let source = root.path().join("source.xlsx");
    let target = root.path().join("printed.xlsx");
    let config =
        create_print_config_with_settings(root.path(), &source, &target, "1", "B2:D4", "是", 1)?;
    create_print_source(&source)?;
    let source_before = fs::read(&source)?;

    let execution = run_menu(&config, "3\n1\n1\n0\n0\n")?;
    assert!(
        execution.status.success(),
        "static print failed: {}",
        String::from_utf8_lossy(&execution.stderr)
    );
    let stdout = String::from_utf8_lossy(&execution.stdout);
    assert!(!stdout.contains("开始按打印配置生成静态打印文件"));
    assert!(
        !stdout.contains("输入 y/yes/是"),
        "no confirmation expected"
    );
    // 同一输出文件只显示一次。
    let output_line_count = stdout
        .lines()
        .filter(|line| line.contains("输出："))
        .count();
    assert_eq!(output_line_count, 1, "{stdout}");

    let printed_path = target.clone();
    assert!(printed_path.is_file(), "print target was not created");
    let printed = read_workbook(&printed_path)?;
    assert_eq!(printed.sheets.len(), 1);
    let sheet = &printed.sheets[0];
    assert_eq!(sheet.name, "Printed");
    // 数值 0 被隐藏（零值不输出=是），文本 "0" 保留，0.0001 保持数值。
    assert!(
        !sheet.values.contains_key("A1"),
        "numeric zero must be blank: {:?}",
        sheet.values.get("A1")
    );
    assert_eq!(sheet.values["A2"]["value"], 2.0);
    assert_eq!(
        sheet.values["A3"]["value"], 3.0,
        "formula must become its cached value"
    );
    assert_eq!(sheet.values["B3"]["type"], "number");
    assert_eq!(sheet.values["B3"]["value"], 0.0001);
    assert_eq!(sheet.values["C3"]["type"], "string");
    assert_eq!(sheet.values["C3"]["value"], "0");
    assert!(
        sheet.formulas.is_empty(),
        "print output must be static values"
    );
    assert_eq!(
        sheet.comments.get("A1").map(String::as_str),
        Some("Author:\n打印区域 1"),
        "PyCode keeps comments inside printed ranges by default"
    );
    assert_eq!(
        sheet.comments.get("C3").map(String::as_str),
        Some("Author:\n打印区域#1"),
        "comments should follow their cells into the output range"
    );
    assert_eq!(
        sheet.comments.len(),
        2,
        "comments outside the print range are omitted"
    );
    assert_eq!(
        sheet.merges,
        vec!["B2:C2".to_owned()],
        "merged range must be preserved and shifted"
    );
    assert_eq!(
        fs::read(&source)?,
        source_before,
        "source must stay unchanged"
    );
    Ok(())
}

#[test]
fn print_precheck_reports_history_and_off_suppresses_progress() -> Result<()> {
    let root = TestDirectory::new()?;
    let source = root.path().join("source.xlsx");
    let target = root.path().join("printed.xlsx");
    let config = create_print_config(root.path(), &source, &target)?;
    create_print_source(&source)?;
    let log_dir = root.path().join("logs");

    let developer = run_menu_with_view(&config, "3\n3\n0\n0\n", "developer")?;
    assert!(
        developer.status.success(),
        "precheck failed: {}",
        String::from_utf8_lossy(&developer.stderr)
    );
    let developer_stdout = String::from_utf8_lossy(&developer.stdout);
    assert!(
        developer_stdout.contains("PT-3 开始：PT-3-4 config_precheck"),
        "{developer_stdout}"
    );
    assert!(
        developer_stdout.contains("PT-3 完成：PT-3-4"),
        "{developer_stdout}"
    );
    let _ = output_json(&developer)?;

    let histories = fs::read_dir(&log_dir)?
        .filter_map(|entry| entry.ok())
        .map(|entry| entry.path())
        .filter(|path| {
            path.file_name()
                .and_then(|name| name.to_str())
                .is_some_and(|name| name.starts_with("pytools-print-config_precheck-"))
        })
        .collect::<Vec<_>>();
    assert_eq!(histories.len(), 1, "precheck must write one run history");
    let history_text = fs::read_to_string(&histories[0])?;
    assert!(!history_text.contains(&source.to_string_lossy().to_string()));
    assert!(!history_text.contains(&config.to_string_lossy().to_string()));
    let history: Value = serde_json::from_str(&history_text)?;
    assert_eq!(history["feature_id"], "PT-3-4");
    assert_eq!(history["action"], "config_precheck");
    assert_eq!(history["status"], "succeeded");

    let user = run_menu_with_view(&config, "3\n3\n0\n0\n", "user")?;
    assert!(user.status.success());
    let user_stdout = String::from_utf8_lossy(&user.stdout);
    assert!(!user_stdout.contains("开始：按配置预校验打印任务"));
    assert!(user_stdout.contains("通过配置数：1"), "{user_stdout}");
    let history_count_before_off = fs::read_dir(&log_dir)?.count();

    let off = run_menu_with_view(&config, "3\n3\n0\n0\n", "off")?;
    assert!(off.status.success());
    let off_stdout = String::from_utf8_lossy(&off.stdout);
    assert!(!off_stdout.contains("PT-3 开始："));
    assert!(!off_stdout.contains("开始处理："));
    assert_eq!(
        fs::read_dir(&log_dir)?.count(),
        history_count_before_off,
        "off view must not write history"
    );
    Ok(())
}

#[test]
fn rust_print_entry_writes_configured_output_without_overwriting() -> Result<()> {
    let root = TestDirectory::new()?;
    let source = root.path().join("source.xlsx");
    let target = root.path().join("printed.xlsx");
    let config =
        create_print_config_with_settings(root.path(), &source, &target, "1", "B2:D4", "是", 2)?;
    create_print_source(&source)?;
    let source_before = fs::read(&source)?;

    let rust_target = root.path().join("printed_Rust.xlsx");
    let removed_entry = run_menu(&config, "3\n2\n0\n0\n")?;
    assert!(removed_entry.status.success());
    let removed_stdout = String::from_utf8_lossy(&removed_entry.stdout);
    assert!(
        removed_stdout.contains("仅接受 1、3 或 0"),
        "{removed_stdout}"
    );
    assert!(!target.exists() && !rust_target.exists());
    let execution = pytools_core::static_print_engines::run_configured_static_print(
        &config,
        pytools_core::static_print_engines::StaticPrintEngine::RustWriter,
    )?;
    assert_eq!(execution.task_ok, 2);
    assert_eq!((execution.task_failed, execution.task_skipped), (0, 0));
    assert_eq!(execution.outputs, vec![target.clone()]);
    assert!(execution.events.iter().any(|event| event.engine == "rust_xlsxwriter"));
    assert!(target.is_file());
    assert!(!root.path().join("printed_Rust.xlsx").exists());
    assert_eq!(
        fs::read(&source)?,
        source_before,
        "source must stay unchanged"
    );

    let printed = read_workbook(&target)?;
    assert_eq!(printed.sheets.len(), 1);
    let sheet = &printed.sheets[0];
    assert_eq!(sheet.name, "Printed");
    // 数值/缓存值/零值隐藏语义与 Openpyxl 入口一致。
    assert_eq!(sheet.values["A3"]["value"], 3.0, "formula cache value");
    assert!(!sheet.values.contains_key("A1"), "numeric zero hidden");
    assert_eq!(sheet.values["A2"]["value"], 2.0);
    assert_eq!(sheet.values["B3"]["value"], 0.0001);
    assert_eq!(sheet.values["C3"]["value"], "0");
    // 样式：锚点字体（OOXML 读取映射为 Format）与合并区域平移。
    assert_eq!(sheet.merges, vec!["B2:C2".to_owned()], "merged range");
    assert_eq!(
        fs::read(&source)?,
        source_before,
        "source must stay unchanged after close"
    );
    let output_before = fs::read(&target)?;
    let repeated = run_menu(&config, "3\n1\n1\n0\n0\n")?;
    assert!(repeated.status.success());
    let stdout = String::from_utf8_lossy(&repeated.stdout);
    assert!(stdout.contains("目标路径已存在，拒绝写入"), "{stdout}");
    assert_eq!(fs::read(&target)?, output_before);
    Ok(())
}

#[test]
#[cfg(windows)]
fn com_print_menu_rejects_xlsx_source_with_visible_reason() -> Result<()> {
    let root = TestDirectory::new()?;
    let source = root.path().join("source.xlsx");
    let target = root.path().join("printed.xlsx");
    let config = create_print_config(root.path(), &source, &target)?;
    create_print_source(&source)?;
    let source_before = fs::read(&source)?;

    for view in ["user", "developer"] {
        let output = run_menu_with_view(&config, "3\n1\n3\n0\n0\n", view)?;
        let stdout = String::from_utf8_lossy(&output.stdout);
        assert!(output.status.success(), "{output:?}");
        assert!(stdout.contains("COM 版式仅支持 XLS 源 → 新 XLSX 目标"), "{stdout}");
        assert!(!target.exists());
    }
    assert_eq!(fs::read(&source)?, source_before);
    Ok(())
}

#[test]
fn rust_print_entry_refuses_xlsm_target_with_clear_message() -> Result<()> {
    let root = TestDirectory::new()?;
    let source = root.path().join("source.xlsx");
    let target = root.path().join("printed.xlsm");
    let config = create_print_config(root.path(), &source, &target)?;
    create_print_source(&source)?;

    let error = pytools_core::static_print_engines::run_configured_static_print(
        &config,
        pytools_core::static_print_engines::StaticPrintEngine::RustWriter,
    )
    .expect_err("Rust must reject XLSM targets");
    assert!(error.to_string().contains("Rust 引擎暂不支持 XLSM 输出"), "{error}");
    assert!(!target.exists(), "failed print must not create a target");
    assert!(!root.path().join("printed_Rust.xlsm").exists());
    Ok(())
}

#[test]
fn v2_print_menu_hides_removed_pdf_entry() -> Result<()> {
    let root = TestDirectory::new()?;
    let source = root.path().join("unused-source.xlsx");
    let target = root.path().join("must-not-be-created.xlsx");
    let config = create_print_config(root.path(), &source, &target)?;

    let output = run_menu(&config, "3\n0\n0\n")?;
    assert!(
        output.status.success(),
        "print menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("打印配置预校验"));
    assert!(!stdout.contains("PDF"));
    assert!(!stdout.contains("按批注"));
    assert!(!target.exists(), "menu listing must not create output");
    Ok(())
}

fn create_print_source(path: &Path) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Marked")?;
    sheet.write_string(0, 0, "Report")?;
    sheet.insert_note(0, 0, &rust_xlsxwriter::Note::new("打印输出 横向"))?;

    sheet.write_number_with_format(1, 1, 0.0, &rust_xlsxwriter::Format::new().set_bold())?;
    sheet.write_string(1, 2, "A")?;
    sheet.write_string(1, 3, "B")?;
    sheet.write_number(2, 1, 2.0)?;
    sheet.merge_range(2, 2, 2, 3, "merged", &rust_xlsxwriter::Format::new())?;
    sheet.write_formula(3, 1, rust_xlsxwriter::Formula::new("1+2").set_result("3"))?;
    sheet.write_number(3, 2, 0.0001)?;
    sheet.write_string(3, 3, "0")?;
    sheet.insert_note(1, 1, &rust_xlsxwriter::Note::new("打印区域 1"))?;
    sheet.insert_note(3, 3, &rust_xlsxwriter::Note::new("打印区域#1"))?;

    sheet.write_string(1, 5, "second")?;
    sheet.write_string(2, 6, "block")?;
    sheet.insert_note(1, 5, &rust_xlsxwriter::Note::new("打印区域 2"))?;
    sheet.insert_note(2, 6, &rust_xlsxwriter::Note::new("打印区域#2"))?;
    sheet.set_row_height(1, 22.0)?;
    sheet.set_column_width(1, 14.0)?;
    workbook.add_worksheet().set_name("Plain")?;
    workbook.save(path)?;
    Ok(())
}

fn create_print_config(root: &Path, source: &Path, target: &Path) -> Result<PathBuf> {
    create_print_config_with_settings(root, source, target, "1", "B2:D4", "", 1)
}

fn create_print_config_with_settings(
    root: &Path,
    source: &Path,
    target: &Path,
    mode: &str,
    print_range: &str,
    hide_zero: &str,
    task_count: u32,
) -> Result<PathBuf> {
    let config = root.join("config.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    {
        let global = workbook.add_worksheet();
        global.set_name("全局配置")?;
        global.write_string(0, 0, "配置项")?;
        global.write_string(0, 1, "值")?;
        for (row, (key, value)) in [
            ("输入目录", root.to_string_lossy().into_owned()),
            (
                "输出目录",
                root.join("output").to_string_lossy().into_owned(),
            ),
            ("日志目录", root.join("logs").to_string_lossy().into_owned()),
            ("错误策略", "fail_fast".to_owned()),
        ]
        .iter()
        .enumerate()
        {
            global.write_string((row + 1) as u32, 0, *key)?;
            global.write_string((row + 1) as u32, 1, value)?;
        }
    }
    {
        let settings = workbook.add_worksheet();
        settings.set_name("打印配置")?;
        let headers = [
            "是否启用",
            "打印模式",
            "源工作簿",
            "源工作表",
            "源工作表打印区域",
            "目标工作簿",
            "目标工作表",
            "FitToPagesWide",
            "FitToPagesTall",
            "打印方向",
            "水平居中",
            "垂直居中",
            "不输出批注",
            "零值不输出",
            "备注",
        ];
        for (column, header) in headers.iter().enumerate() {
            settings.write_string(0, column as u16, *header)?;
        }
        let values = [
            "是".to_owned(),
            mode.to_owned(),
            source.to_string_lossy().into_owned(),
            "Marked".to_owned(),
            print_range.to_owned(),
            target.to_string_lossy().into_owned(),
            "Printed".to_owned(),
            "1".to_owned(),
            "1".to_owned(),
            "横向".to_owned(),
            "是".to_owned(),
            "是".to_owned(),
            "".to_owned(),
            hide_zero.to_owned(),
            "".to_owned(),
        ];
        for row in 1..=task_count {
            for (column, value) in values.iter().enumerate() {
                settings.write_string(row, column as u16, value.as_str())?;
            }
        }
    }
    workbook.save(&config)?;
    Ok(config)
}

fn run_menu(config: &Path, input: &str) -> Result<Output> {
    run_menu_with_view(config, input, "user")
}

fn run_menu_with_view(config: &Path, input: &str, log_view: &str) -> Result<Output> {
    let mut command = Command::new(env!("CARGO_BIN_EXE_RustTools"));
    command.env("PYTOOLS_PRINT_STAGE_TIMING", "1");
    // 3-1 打印只认随包 openpyxl 运行时（不做 PATH 回退）。测试二进制旁边
    // 没有 python/，因此显式指向发行包的 dist/python 与仓库内的 worker。
    if let Some((python, worker)) = bundled_openpyxl_runtime() {
        command
            .env("PYTOOLS_OPENPYXL_PYTHON", python)
            .env("PYTOOLS_OPENPYXL_WORKER", worker);
    }
    let mut child = command
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
    let output = child
        .wait_with_output()
        .context("wait for interactive menu")?;
    logging_evidence::capture(log_view, &output)?;
    Ok(output)
}

/// 发行包随附的 openpyxl 运行时：(解释器, worker)。
fn bundled_openpyxl_runtime() -> Option<(PathBuf, PathBuf)> {
    let workspace = Path::new(env!("CARGO_MANIFEST_DIR")).parent()?;
    let worker = workspace.join("crates/excel-editor/python/openpyxl_worker.py");
    let python_names: &[&str] = if cfg!(windows) {
        &[
            "dist/python/python.exe",
            "dist/python/runtime/bin/python.exe",
        ]
    } else {
        &["dist/python/python3", "dist/python/runtime/bin/python3"]
    };
    let python = python_names
        .iter()
        .map(|relative| workspace.join(relative))
        .find(|path| path.is_file())?;
    worker.is_file().then_some((python, worker))
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

/// 统一日志时间戳：行内出现 `[YYYY-MM-DD HH:MM:SS] ` 前缀（沿用旧版 PyTools 时间格式）。
/// 菜单提示符（如 `子菜单选择> `）可能出现在时间戳之前。
fn starts_with_timestamp(line: &str) -> bool {
    let line = match line.find('[') {
        Some(index) => &line[index..],
        None => return false,
    };
    let bytes = line.as_bytes();
    bytes.len() >= 22
        && bytes[0] == b'['
        && bytes[5] == b'-'
        && bytes[8] == b'-'
        && bytes[11] == b' '
        && bytes[14] == b':'
        && bytes[17] == b':'
        && bytes[20] == b']'
        && bytes[4].is_ascii_digit()
        && bytes[21] == b' '
}
