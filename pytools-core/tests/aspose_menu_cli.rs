use anyhow::{Context, Result};
use excel_reader::read_workbook;
use std::{
    fs,
    io::Write,
    path::{Path, PathBuf},
    process::{Command, Output, Stdio},
    time::{SystemTime, UNIX_EPOCH},
};

fn root(label: &str) -> Result<PathBuf> {
    let root = std::env::temp_dir().join(format!(
        "rusttools-aspose-main-{label}-{}-{}",
        std::process::id(),
        SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos()
    ));
    fs::create_dir_all(&root)?;
    Ok(root)
}
fn config(root: &Path, engine: &str) -> Result<PathBuf> {
    let path = root.join("config.xlsx");
    let mut book = rust_xlsxwriter::Workbook::new();
    let sheet = book.add_worksheet();
    sheet.set_name("全局配置")?;
    for (i, (k, v)) in [
        ("键", "值"),
        ("输出目录", "output"),
        ("日志目录", "logs"),
        ("校验计算引擎", engine),
    ]
    .into_iter()
    .enumerate()
    {
        sheet.write_string(i as u32, 0, k)?;
        sheet.write_string(i as u32, 1, v)?;
    }
    book.save(&path)?;
    Ok(path)
}
fn fixtures(root: &Path, count: usize) -> Result<(PathBuf, Vec<PathBuf>)> {
    let template = root.join("template.xlsx");
    let mut book = rust_xlsxwriter::Workbook::new();
    let sheet = book.add_worksheet();
    sheet.set_name("数据")?;
    sheet.write_string(0, 0, "项目")?;
    sheet.write_string(0, 1, "金额")?;
    sheet.write_formula(
        1,
        3,
        rust_xlsxwriter::Formula::new(
            "=IF(SUM('数据'!B2:B3)='参数'!A1,\"通过\",\"校验失败：金额不平\")",
        )
        .set_result("STALE"),
    )?;
    book.add_worksheet()
        .set_name("参数")?
        .write_number(0, 0, 30.0)?;
    book.define_name("表结构区域", "'数据'!$A$1:$B$1")?;
    book.define_name("校验区域", "'数据'!$D$2")?;
    book.save(&template)?;
    let mut sources = Vec::new();
    for i in 0..count {
        let path = root.join(format!("source-{i:03}.xlsx"));
        let mut book = rust_xlsxwriter::Workbook::new();
        let sheet = book.add_worksheet();
        sheet.set_name("数据")?;
        sheet.write_string(0, 0, "项目")?;
        sheet.write_string(0, 1, "金额")?;
        sheet.write_string(1, 0, "甲")?;
        sheet.write_string(2, 0, "乙")?;
        sheet.write_number(1, 1, 10.0)?;
        sheet.write_number(2, 1, if i == 0 { 20.0 } else { 40.0 })?;
        book.save(&path)?;
        sources.push(path);
    }
    Ok((template, sources))
}
fn command(config: &Path) -> Command {
    let mut c = Command::new(env!("CARGO_BIN_EXE_RustTools"));
    c.env("PYTOOLS_CONFIG_PATH", config);
    c
}
fn menu(root: &Path, config: &Path, input: &str) -> Result<Output> {
    let mut child = command(config)
        .arg("menu")
        .arg("--config")
        .arg(config)
        .arg("--log-dir")
        .arg(root.join("logs"))
        .env(
            "RUSTTOOLS_ASPOSE_WORKER_EXE",
            root.join("not-installed.exe"),
        )
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()?;
    child.stdin.take().unwrap().write_all(input.as_bytes())?;
    Ok(child.wait_with_output()?)
}
fn run(
    root: &Path,
    config: &Path,
    mode: &str,
    template: &Path,
    sources: &[PathBuf],
    missing_worker: bool,
) -> Result<Output> {
    let mut cmd = command(config);
    cmd.arg("summary-named-validation")
        .args(["--calc-engine", mode, "--template"])
        .arg(template)
        .arg("--output-dir")
        .arg(root.join("output"))
        .arg("--log-dir")
        .arg(root.join("logs"))
        .args(["--log-view", "developer"]);
    if missing_worker {
        cmd.env(
            "RUSTTOOLS_ASPOSE_WORKER_EXE",
            root.join("not-installed.exe"),
        );
    }
    cmd.args(sources);
    Ok(cmd.output()?)
}
fn assert_ok(output: &Output) {
    assert!(
        output.status.success(),
        "{}\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
}
fn history(root: &Path) -> Result<serde_json::Value> {
    let path = fs::read_dir(root.join("logs"))?
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .find(|p| {
            p.file_name()
                .unwrap()
                .to_string_lossy()
                .starts_with("PT-6-3_")
                && p.extension().is_some_and(|e| e == "json")
        })
        .context("missing run report")?;
    Ok(serde_json::from_slice(&fs::read(path)?)?)
}
#[test]
fn calc_choice_is_before_template_and_cancel_keeps_configuration() -> Result<()> {
    let r = root("cancel")?;
    let c = config(&r, "betteroffice")?;
    let before = fs::read(&c)?;
    let out = menu(&r, &c, "1\n5\n0\n0\n")?;
    assert_ok(&out);
    let text = String::from_utf8_lossy(&out.stdout);
    assert!(text.contains("选择本次计算模式"), "{text}");
    assert!(!text.contains("选择按模板校验区域提取校验的模板工作簿"), "{text}");
    assert!(!r.join("output").exists());
    assert_eq!(before, fs::read(c)?);
    Ok(())
}
#[test]
fn calculation_menu_uses_configured_automatic_engine() -> Result<()> {
    let r = root("missing-menu")?;
    let c = config(&r, "betteroffice")?;
    let (template, sources) = fixtures(&r, 1)?;
    let out = menu(
        &r,
        &c,
        &format!("1\n5\n1\n{}\n1\n{}\n0\n0\n", template.display(), sources[0].display()),
    )?;
    assert_ok(&out);
    let text = String::from_utf8_lossy(&out.stdout);
    assert!(text.contains("默认 BetterOffice"), "{text}");
    assert!(!text.contains("  4  Aspose"), "{text}");
    assert!(history(&r).is_ok(), "{text}");
    Ok(())
}

#[cfg(target_os = "linux")]
#[test]
fn auto_calculation_requires_explicit_libreoffice_fallback_in_selected_config() -> Result<()> {
    let r = root("linux-opt-in")?;
    let (template, sources) = fixtures(&r, 1)?;
    let before = fs::read(&sources[0])?;
    for fallback in ["auto", "libreoffice"] {
        let work = r.join(fallback);
        fs::create_dir(&work)?;
        let config = work.join("selected.xlsx");
        let mut book = rust_xlsxwriter::Workbook::new();
        let sheet = book.add_worksheet().set_name("全局配置")?;
        for (row, (key, value)) in [
            ("键", "值"),
            ("校验计算引擎", "com"),
            ("校验计算预算上限", "1"),
            ("校验计算回退策略", fallback),
        ].into_iter().enumerate() {
            sheet.write_string(row as u32, 0, key)?;
            sheet.write_string(row as u32, 1, value)?;
        }
        book.save(&config)?;
        let missing_worker = work.join("missing-lo-worker.py");
        let output = Command::new(env!("CARGO_BIN_EXE_RustTools"))
            .current_dir(&work)
            .env_remove("PYTOOLS_CONFIG_PATH")
            .env("PYTOOLS_LO_CALC_WORKER", &missing_worker)
            .arg("summary-named-validation")
            .arg("--config").arg(&config)
            .arg("--template").arg(&template)
            .arg("--log-dir").arg(work.join("logs"))
            .arg("--output-dir").arg(work.join("output"))
            .args(["--log-view", "developer"])
            .args(&sources)
            .output()?;
        assert!(!output.status.success(), "{output:?}");
        let stderr = String::from_utf8_lossy(&output.stderr);
        if fallback == "auto" {
            assert!(stderr.contains("禁止回退"), "{stderr}");
            assert!(!stderr.contains("missing-lo-worker"), "{stderr}");
        } else {
            assert!(stderr.contains("missing-lo-worker"), "{stderr}");
        }
        assert_eq!(fs::read(&sources[0])?, before);
    }
    Ok(())
}
#[test]
fn retired_aspose_calculation_never_falls_back_or_changes_source() -> Result<()> {
    let r = root("missing-cli")?;
    let c = config(&r, "betteroffice")?;
    let (t, s) = fixtures(&r, 1)?;
    let before = fs::read(&s[0])?;
    let out = run(&r, &c, "aspose", &t, &s, true)?;
    assert!(!out.status.success());
    assert!(String::from_utf8_lossy(&out.stderr).contains("仅保留用于 XLS→XLSX"));
    assert!(!r.join("output").exists());
    assert_eq!(before, fs::read(&s[0])?);
    Ok(())
}
#[test]
fn invalid_mode_is_rejected_without_side_effects() -> Result<()> {
    let r = root("invalid")?;
    let c = config(&r, "betteroffice")?;
    let (t, s) = fixtures(&r, 1)?;
    let out = run(&r, &c, "typo", &t, &s, true)?;
    assert!(!out.status.success());
    assert!(String::from_utf8_lossy(&out.stderr).contains("calc-engine"));
    assert!(!r.join("output").exists());
    Ok(())
}
#[test]
fn forced_betteroffice_overrides_com_config_without_using_aspose() -> Result<()> {
    let r = root("bo-only")?;
    let c = config(&r, "com")?;
    let (t, s) = fixtures(&r, 2)?;
    let out = run(&r, &c, "betteroffice", &t, &s, true)?;
    assert_ok(&out);
    let h = history(&r)?;
    assert_eq!(h["requested_calc_engine"], "betteroffice");
    assert_eq!(h["input_stats"]["succeeded"], 2);
    assert_eq!(h["input_stats"]["failed"], 0);
    assert!(String::from_utf8_lossy(&out.stderr).contains("betteroffice (pure Rust)"));
    Ok(())
}
#[test]
fn auto_keeps_existing_betteroffice_route_when_aspose_is_absent() -> Result<()> {
    let r = root("auto")?;
    let c = config(&r, "betteroffice")?;
    let (t, s) = fixtures(&r, 2)?;
    let out = run(&r, &c, "auto", &t, &s, true)?;
    assert_ok(&out);
    let h = history(&r)?;
    assert_eq!(h["requested_calc_engine"], "auto");
    assert_eq!(h["input_stats"]["succeeded"], 2);
    assert_eq!(h["input_stats"]["failed"], 0);
    Ok(())
}

#[test]
fn forced_betteroffice_budget_refusal_is_locatable_and_never_falls_back() -> Result<()> {
    let r = root("bo-refusal")?;
    let c = r.join("config.xlsx");
    let mut book = rust_xlsxwriter::Workbook::new();
    let sheet = book.add_worksheet();
    sheet.set_name("全局配置")?;
    for (i, (k, v)) in [
        ("键", "值"),
        ("校验计算引擎", "com"),
        ("校验计算预算上限", "1"),
    ]
    .into_iter()
    .enumerate()
    {
        sheet.write_string(i as u32, 0, k)?;
        sheet.write_string(i as u32, 1, v)?;
    }
    book.save(&c)?;
    let (t, sources) = fixtures(&r, 1)?;
    let before = fs::read(&sources[0])?;
    let out = run(&r, &c, "betteroffice", &t, &sources, true)?;
    assert!(!out.status.success());
    let h = history(&r)?;
    assert_eq!(h["input_stats"]["failed"], 1);
    assert_eq!(h["input_stats"]["succeeded"], 0);
    assert_eq!(h["error_summary"]["file_name"], "source-000.xlsx");
    assert_eq!(before, fs::read(&sources[0])?);
    Ok(())
}

#[cfg(any(windows, target_os = "linux"))]
fn real_batch(count: usize) -> Result<()> {
    let r = root(&format!("real-{count}"))?;
    let c = config(&r, "betteroffice")?;
    let (t, s) = fixtures(&r, count)?;
    let originals = s
        .iter()
        .map(fs::read)
        .collect::<std::io::Result<Vec<_>>>()?;
    let template_before = fs::read(&t)?;
    let out = run(&r, &c, "aspose", &t, &s, false)?;
    fs::write(r.join("terminal.stdout.txt"), &out.stdout)?;
    fs::write(r.join("terminal.stderr.txt"), &out.stderr)?;
    assert_ok(&out);
    let h = history(&r)?;
    assert_eq!(h["requested_calc_engine"], "aspose");
    assert_eq!(h["input_stats"]["succeeded"], count);
    assert_eq!(h["input_stats"]["failed"], 0);
    let result: serde_json::Value = serde_json::from_slice(&out.stdout)?;
    assert_eq!(result["result"]["files"], count);
    assert_eq!(result["result"]["rows"], count - 1);
    let host_log = fs::read_dir(r.join("logs/aspose-worker"))?
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .find(|p| {
            p.file_name()
                .unwrap()
                .to_string_lossy()
                .starts_with("host-calculations-")
        })
        .context("missing per-run Aspose host log")?;
    let logs = fs::read_to_string(host_log)?;
    let rows = logs
        .lines()
        .map(serde_json::from_str::<serde_json::Value>)
        .collect::<serde_json::Result<Vec<_>>>()?;
    assert_eq!(rows.len(), count);
    let mut pids = std::collections::BTreeSet::new();
    for (i, row) in rows.iter().enumerate() {
        assert_eq!(row["run_id"], h["run_id"]);
        assert!(row["outcome"]["evaluation_output"].as_bool().unwrap());
        pids.insert(row["outcome"]["worker_pid"].as_u64().unwrap());
        let copy = r.join("output/副本").join(format!("source-{i:03}.xlsx"));
        let book = read_workbook(copy)?;
        assert!(book.sheets.iter().any(|s| s.name == "Evaluation Warning"));
        let sheet = book.sheets.iter().find(|s| s.name == "数据").unwrap();
        assert_eq!(
            sheet.values["D2"]["value"],
            if i == 0 {
                "通过"
            } else {
                "校验失败：金额不平"
            }
        );
        assert_eq!(originals[i], fs::read(&s[i])?);
    }
    assert_eq!(pids.len(), if count > 99 { 2 } else { 1 });
    if count > 99 {
        assert_eq!(rows[98]["outcome"]["worker_attempts"], 99);
        assert_eq!(rows[99]["outcome"]["worker_attempts"], 1);
    }
    assert_eq!(template_before, fs::read(&t)?);
    println!("REAL_ASPOSE_REPORT={}", r.display());
    Ok(())
}
#[cfg(any(windows, target_os = "linux"))]
#[test]
#[ignore = "Requires explicitly configured external Aspose worker and .NET runtime"]
fn real_aspose_two_files_reuses_worker_and_preserves_sources() -> Result<()> {
    real_batch(2)
}
#[cfg(any(windows, target_os = "linux"))]
#[test]
#[ignore = "Requires explicitly configured external Aspose worker and .NET runtime"]
fn real_aspose_main_batch_rotates_after_99_successful_files() -> Result<()> {
    real_batch(100)
}

#[cfg(any(windows, target_os = "linux"))]
#[test]
#[ignore = "Requires explicitly configured external Aspose worker and .NET runtime"]
fn real_aspose_unknown_function_cannot_be_reported_as_validation_pass() -> Result<()> {
    let r = root("aspose-unsupported")?;
    let c = config(&r, "betteroffice")?;
    let (t, sources) = fixtures(&r, 1)?;
    let mut book = rust_xlsxwriter::Workbook::new();
    let sheet = book.add_worksheet();
    sheet.set_name("数据")?;
    sheet.write_string(0, 0, "项目")?;
    sheet.write_string(0, 1, "金额")?;
    sheet.write_formula(
        1,
        3,
        rust_xlsxwriter::Formula::new("=RUSTTOOLS_UNSUPPORTED_FUNCTION(B2)").set_result("STALE"),
    )?;
    book.define_name("表结构区域", "'数据'!$A$1:$B$1")?;
    book.define_name("校验区域", "'数据'!$D$2")?;
    book.save(&t)?;
    let before = fs::read(&sources[0])?;
    let out = run(&r, &c, "aspose", &t, &sources, false)?;
    fs::write(r.join("terminal.stdout.txt"), &out.stdout)?;
    fs::write(r.join("terminal.stderr.txt"), &out.stderr)?;
    assert!(!out.status.success());
    let h = history(&r)?;
    assert_eq!(h["input_stats"]["failed"], 1);
    assert_eq!(h["input_stats"]["succeeded"], 0);
    assert_eq!(h["error_summary"]["file_name"], "source-000.xlsx");
    assert_eq!(before, fs::read(&sources[0])?);
    println!("REAL_ASPOSE_REPORT={}", r.display());
    Ok(())
}
