use anyhow::{Context, Result};
use std::{
    fs,
    io::Write,
    path::{Path, PathBuf},
    process::{Command, Stdio},
    sync::atomic::{AtomicU64, Ordering},
    time::{SystemTime, UNIX_EPOCH},
};

static NEXT_TEMP_DIRECTORY: AtomicU64 = AtomicU64::new(0);

struct TestDirectory(PathBuf);

impl TestDirectory {
    fn new() -> Result<Self> {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .context("system clock before Unix epoch")?
            .as_nanos();
        let serial = NEXT_TEMP_DIRECTORY.fetch_add(1, Ordering::Relaxed);
        let path = std::env::temp_dir().join(format!(
            "pytools-quick-menu-{}-{nonce}-{serial}",
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
fn v2_configuration_menu_routes_all_three_quick_actions_without_starting_a_desktop_app(
) -> Result<()> {
    let root = TestDirectory::new()?;
    let blocked_parent = root.path().join("blocked-parent");
    fs::write(&blocked_parent, b"not a directory")?;
    let config = blocked_parent.join("config").join("config.xlsx");

    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args([
            "menu",
            "--config",
            config.to_string_lossy().as_ref(),
            "--log-view",
            "off",
        ])
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()?;
    child
        .stdin
        .take()
        .context("menu stdin was not piped")?
        .write_all(b"7\n3\n4\n5\n0\n0\n")?;
    let result = child.wait_with_output()?;
    assert!(
        result.status.success(),
        "quick menu exited unsuccessfully: {}",
        String::from_utf8_lossy(&result.stderr)
    );

    // 失败/提示信息统一走终端日志行（stderr），不再以 [失败]/[提示] 前缀打印到 stdout。
    let terminal = format!(
        "{}{}",
        String::from_utf8_lossy(&result.stdout),
        String::from_utf8_lossy(&result.stderr)
    );
    let stdout = String::from_utf8_lossy(&result.stdout);
    assert!(
        stdout.contains("  7  配置与快捷操作"),
        "main menu missing configuration/quick-actions option: {stdout}"
    );
    assert!(
        stdout.contains("------ 配置与快捷操作 ------"),
        "configuration and quick-actions menu missing: {stdout}"
    );
    assert!(
        terminal.contains("打开配置文件失败；请检查文件是否存在以及系统是否允许打开"),
        "config route missing: {terminal}"
    );
    assert!(
        terminal.contains("输出目录中没有可打开的文件"),
        "latest-output route missing: {terminal}"
    );
    assert!(
        terminal.contains("打开输出目录失败；请检查文件是否存在以及系统是否允许打开"),
        "output-folder route missing: {terminal}"
    );
    assert!(
        stdout.contains("再见。"),
        "quick action did not return to the main menu: {stdout}"
    );
    Ok(())
}

#[test]
fn launching_without_arguments_opens_the_legacy_numbered_main_menu() -> Result<()> {
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()?;
    child
        .stdin
        .take()
        .context("menu stdin was not piped")?
        .write_all(b"0\n")?;
    let result = child.wait_with_output()?;
    assert!(
        result.status.success(),
        "direct menu launch failed: {}",
        String::from_utf8_lossy(&result.stderr)
    );

    let stdout = String::from_utf8_lossy(&result.stdout);
    for legacy_entry in [
        "======== PyTools 菜单 ========",
        "  1  数据提取与汇总",
        "  2  数据去重与追加",
        "  3  报表打印",
        "  4  文件处理",
        "  5  专用报表处理",
        "  6  归档整理",
        "  7  配置与快捷操作",
        "  0  退出",
        "再见。",
    ] {
        assert!(
            stdout.contains(legacy_entry),
            "direct launch is missing {legacy_entry:?}: {stdout}"
        );
    }
    Ok(())
}

#[test]
fn launching_with_global_menu_options_without_command_opens_the_menu() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = root.path().join("selected config.xlsx");
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("--config")
        .arg(&config)
        .args(["--log-view", "off"])
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()?;
    child
        .stdin
        .take()
        .context("menu stdin was not piped")?
        .write_all(b"0\n")?;
    let result = child.wait_with_output()?;
    assert!(
        result.status.success(),
        "menu options without an explicit command failed: {}",
        String::from_utf8_lossy(&result.stderr)
    );

    let stdout = String::from_utf8_lossy(&result.stdout);
    assert!(stdout.contains("  1  数据提取与汇总"));
    assert!(stdout.contains(&config.to_string_lossy().to_string()));
    assert!(stdout.contains("再见。"));
    Ok(())
}

#[test]
fn launching_without_arguments_honors_the_legacy_config_path_override() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = root.path().join("selected config.xlsx");
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .env("PYTOOLS_CONFIG_PATH", &config)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()?;
    child
        .stdin
        .take()
        .context("menu stdin was not piped")?
        .write_all(b"0\n")?;
    let result = child.wait_with_output()?;
    assert!(
        result.status.success(),
        "direct menu launch failed: {}",
        String::from_utf8_lossy(&result.stderr)
    );

    let stdout = String::from_utf8_lossy(&result.stdout);
    assert!(
        stdout.contains(&config.to_string_lossy().to_string()),
        "menu did not use PYTOOLS_CONFIG_PATH: {stdout}"
    );
    assert!(stdout.contains("再见。"));
    Ok(())
}

#[test]
fn cli_help_lists_quick_tools_menu_category() -> Result<()> {
    let result = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("--help")
        .output()?;
    assert!(
        result.status.success(),
        "pytools --help failed: {}",
        String::from_utf8_lossy(&result.stderr)
    );

    let usage = String::from_utf8_lossy(&result.stderr);
    assert!(
        usage.contains("7 (configuration and quick actions 1-5)"),
        "usage should list the quick tools menu category: {usage}"
    );
    Ok(())
}
