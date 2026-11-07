//! 2-1 / 2-2 / 5-3 / 5-4 的 `.xls` Excel/WPS COM 路由回归测试。
//!
//! 夹具用 rust_xlsxwriter 生成后经随包 Python + Excel COM 另存为真实
//! `.xls`；运行环境缺少随包运行时或 Excel 时整体跳过，保持 CI 绿灯。

use anyhow::{ensure, Result};
use excel_reader::read_workbook;
use serde_json::Value;
use std::{
    fs,
    io::Write,
    path::{Path, PathBuf},
    process::{Command, Stdio},
    sync::atomic::{AtomicU64, Ordering},
};

static NEXT_ID: AtomicU64 = AtomicU64::new(0);

fn bundled_python() -> Option<PathBuf> {
    if let Ok(python) = std::env::var("PYTOOLS_COM_PYTHON") {
        if Path::new(&python).is_file() {
            return Some(PathBuf::from(python));
        }
    }
    let exe = std::env::current_exe().ok()?;
    let mut cursor = exe.parent().map(Path::to_path_buf);
    while let Some(dir) = cursor {
        let candidate = dir.join("python").join("python.exe");
        if candidate.is_file() {
            return Some(candidate);
        }
        cursor = dir.parent().map(Path::to_path_buf);
    }
    None
}

fn runtime_available() -> bool {
    bundled_python().is_some()
}

fn next_id() -> u64 {
    NEXT_ID.fetch_add(1, Ordering::Relaxed)
}

fn run_com_python(python: &Path, script_text: &str, args: &[&Path]) -> Result<String> {
    let nonce = next_id();
    let script_path = args
        .first()
        .and_then(|first| first.parent().map(Path::to_path_buf))
        .unwrap_or_else(|| PathBuf::from("."))
        .join(format!("com-setup-{nonce}.py"));
    fs::write(&script_path, script_text)?;
    let output = Command::new(python)
        .arg("-I")
        .arg(&script_path)
        .args(args.iter().map(|p| p.as_os_str()))
        .env("PYTHONIOENCODING", "utf-8")
        .output()?;
    let _ = fs::remove_file(&script_path);
    ensure!(
        output.status.success(),
        "COM python step failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    Ok(String::from_utf8_lossy(&output.stdout).into_owned())
}

/// 生成 .xlsx 内容后经 Excel COM 另存为真实 .xls（FileFormat 56）。
fn make_xls_via_com(python: &Path, path: &Path, kind: &str) -> Result<()> {
    let xlsx = path.with_extension("xls.tmp.xlsx");
    if kind == "fx" {
        let mut book = rust_xlsxwriter::Workbook::new();
        let sheet = book.add_worksheet();
        sheet.set_name("(6)惠州市外汇信贷收支表")?;
        sheet.write_string(0, 0, "汇总")?;
        let plain = book.add_worksheet();
        plain.set_name("普通表")?;
        plain.write_string(0, 0, "非外汇")?;
        book.save(&xlsx)?;
        let script = r#"import win32com.client, sys, pathlib
src = pathlib.Path(sys.argv[1])
dst = pathlib.Path(sys.argv[2])
app = win32com.client.DispatchEx("Excel.Application")
app.Visible = False
app.DisplayAlerts = False
try:
    book = app.Workbooks.Open(str(src))
    for sheet in book.Worksheets:
        if "外汇" in sheet.Name:
            sheet.PageSetup.RightHeader = "单位：万元"
        else:
            sheet.PageSetup.RightHeader = "单位：万元"
    book.SaveAs(str(dst), FileFormat=56)
    book.Close(False)
finally:
    app.Quit()
"#;
        run_com_python(python, script, &[&xlsx, path])?;
        let _ = fs::remove_file(&xlsx);
        return Ok(());
    }
    let mut book = rust_xlsxwriter::Workbook::new();
    match kind {
        "duplicates" => {
            let sheet = book.add_worksheet();
            sheet.set_name("Commented")?;
            for (row, (a, b)) in [
                ("机构", "代码"),
                ("模拟甲", "001"),
                ("模拟乙", "002"),
                ("模拟甲", "001"),
                ("模拟丙", "003"),
                ("模拟甲", "001"),
            ]
            .iter()
            .enumerate()
            {
                sheet.write_string(row as u32, 0, *a)?;
                sheet.write_string(row as u32, 1, *b)?;
            }
        }
        "foreign" => {
            let sheet = book.add_worksheet();
            sheet.set_name("(2)惠州市本外币存贷款分机构表")?;
            for (row, name) in [
                "本地商业银行",
                "外资甲银行",
                "本地农商行",
                "外资乙银行",
                "本地建行",
            ]
            .iter()
            .enumerate()
            {
                sheet.write_string(row as u32, 0, *name)?;
                sheet.write_number(row as u32, 1, (row + 1) as f64 * 100.0)?;
            }
        }
        other => anyhow::bail!("unknown fixture kind {other}"),
    }
    book.save(&xlsx)?;
    let script = r#"import win32com.client, sys, pathlib
src = pathlib.Path(sys.argv[1])
dst = pathlib.Path(sys.argv[2])
app = win32com.client.DispatchEx("Excel.Application")
app.Visible = False
app.DisplayAlerts = False
try:
    book = app.Workbooks.Open(str(src))
    book.SaveAs(str(dst), FileFormat=56)
    book.Close(False)
finally:
    app.Quit()
"#;
    run_com_python(python, script, &[&xlsx, path])?;
    let _ = fs::remove_file(&xlsx);
    Ok(())
}

fn build_dedup_config(root: &Path, source: &Path) -> Result<PathBuf> {
    let config = root.join("config.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    {
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
    }
    {
        let tasks = workbook.add_worksheet();
        tasks.set_name("去重追加数据配置")?;
        for (column, header) in [
            "是否启用",
            "任务名",
            "源数据工作簿",
            "源数据工作表",
            "标识列序号",
            "目标工作簿",
            "目标工作表",
            "执行模式",
        ]
        .iter()
        .enumerate()
        {
            tasks.write_string(0, column as u16, *header)?;
        }
        tasks.write_string(1, 0, "是")?;
        tasks.write_string(1, 1, "DedupXls")?;
        tasks.write_string(1, 2, source.to_string_lossy().into_owned())?;
        tasks.write_string(1, 3, "Commented")?;
        tasks.write_string(1, 4, "2")?;
    }
    workbook.save(&config)?;
    Ok(config)
}

fn run_menu(config: &Path, input: &str) -> Result<String> {
    let mut child = Command::new(env!("CARGO_BIN_EXE_RustTools"))
        .arg("menu")
        .arg("--config")
        .arg(config)
        .arg("--log-view")
        .arg("off")
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()?;
    child
        .stdin
        .as_mut()
        .expect("menu stdin")
        .write_all(input.as_bytes())?;
    let output = child.wait_with_output()?;
    ensure!(
        output.status.success(),
        "menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    Ok(String::from_utf8_lossy(&output.stdout).into_owned())
}

fn run_cli(argv: &[String]) -> Result<Value> {
    let output = Command::new(env!("CARGO_BIN_EXE_RustTools"))
        .args(argv)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .output()?;
    ensure!(
        output.status.success(),
        "CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    Ok(serde_json::from_slice(&output.stdout)?)
}

fn max_row_of(dump: &excel_reader::WorkbookDump, sheet_name: &str) -> usize {
    dump.sheets
        .iter()
        .find(|sheet| sheet.name == sheet_name)
        .map(|sheet| {
            sheet
                .values
                .keys()
                .filter_map(|address| {
                    address
                        .rfind(|c: char| c.is_ascii_digit())
                        .and_then(|index| address[index..].parse::<usize>().ok())
                })
                .max()
                .unwrap_or(0)
        })
        .unwrap_or(0)
}

fn column_values(
    dump: &excel_reader::WorkbookDump,
    sheet_name: &str,
    column: usize,
) -> Vec<String> {
    let max_row = max_row_of(dump, sheet_name);
    let letter = ["A", "B", "C", "D"]
        .get(column.saturating_sub(1))
        .copied()
        .unwrap_or("A");
    (1..=max_row)
        .map(|row| {
            dump.sheets
                .iter()
                .find(|sheet| sheet.name == sheet_name)
                .and_then(|sheet| sheet.values.get(&format!("{letter}{row}")))
                .and_then(|value| value.get("value"))
                .and_then(|value| value.as_str())
                .unwrap_or_default()
                .to_owned()
        })
        .collect()
}

#[test]
#[cfg(windows)]
fn dedup_xls_com_routes_mark_and_delete() -> Result<()> {
    if !runtime_available() {
        eprintln!("skipping: no bundled python runtime");
        return Ok(());
    }
    let python = bundled_python().expect("runtime checked");
    let nonce = next_id();
    let root = std::env::temp_dir().join(format!("pytools-dedup-xls-{nonce}"));
    fs::create_dir_all(&root)?;
    let outcome = (|| -> Result<()> {
        let source = root.join("dup.xls");
        make_xls_via_com(&python, &source, "duplicates")?;
        let before_rows = max_row_of(&read_workbook(&source)?, "Commented");
        let config = build_dedup_config(&root, &source)?;

        let menu = run_menu(&config, "2\n1\n0\n0\n")?;
        ensure!(
            !menu.contains("执行失败"),
            "2-1 menu reported failure: {menu}"
        );
        let after_mark = read_workbook(&source)?;
        ensure!(
            max_row_of(&after_mark, "Commented") == before_rows,
            "2-1 must keep all rows when marking"
        );

        let menu = run_menu(&config, "2\n2\n0\n0\n")?;
        ensure!(
            !menu.contains("执行失败"),
            "2-2 menu reported failure: {menu}"
        );
        let values = column_values(&read_workbook(&source)?, "Commented", 1);
        ensure!(
            values == ["机构", "模拟甲", "模拟乙", "模拟丙"],
            "2-2 unexpected surviving rows: {values:?}"
        );
        Ok(())
    })();
    let _ = fs::remove_dir_all(&root);
    outcome
}

#[test]
#[cfg(windows)]
fn remove_foreign_bank_xls_com_creates_financial_bureau_copy() -> Result<()> {
    if !runtime_available() {
        eprintln!("skipping: no bundled python runtime");
        return Ok(());
    }
    let python = bundled_python().expect("runtime checked");
    let nonce = next_id();
    let root = std::env::temp_dir().join(format!("pytools-foreign-xls-{nonce}"));
    fs::create_dir_all(&root)?;
    let outcome = (|| -> Result<()> {
        let source = root.join("fb.xls");
        make_xls_via_com(&python, &source, "foreign")?;
        let config = root.join("config.xlsx");
        let mut workbook = rust_xlsxwriter::Workbook::new();
        {
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
        }
        {
            let mapping = workbook.add_worksheet();
            mapping.set_name("机构映射表")?;
            mapping.write_string(0, 0, "占位")?;
            mapping.write_string(0, 1, "机构名称")?;
            mapping.write_string(0, 2, "外资行标志")?;
            mapping.write_string(1, 1, "外资甲银行")?;
            mapping.write_string(1, 2, "1")?;
            mapping.write_string(2, 1, "外资乙银行")?;
            mapping.write_string(2, 2, "1")?;
        }
        workbook.save(&config)?;
        let argv = [
            "remove-foreign-bank".to_owned(),
            "--config".to_owned(),
            config.to_string_lossy().into_owned(),
            "--create-copy".to_owned(),
            "--log-view".to_owned(),
            "off".to_owned(),
            source.to_string_lossy().into_owned(),
        ];
        run_cli(&argv)?;
        let copy = root.join("fb(金融局).xls");
        ensure!(copy.is_file(), "(金融局) copy missing: {}", copy.display());
        let sheet_name = "(2)惠州市本外币存贷款分机构表";
        let copy_rows = column_values(&read_workbook(&copy)?, sheet_name, 1);
        ensure!(
            copy_rows == ["本地商业银行", "本地农商行", "本地建行"],
            "copy rows unexpected: {copy_rows:?}"
        );
        let source_rows = column_values(&read_workbook(&source)?, sheet_name, 1);
        ensure!(
            source_rows.contains(&"外资甲银行".to_owned()),
            "source must keep foreign rows untouched"
        );
        Ok(())
    })();
    let _ = fs::remove_dir_all(&root);
    outcome
}

#[test]
#[cfg(windows)]
fn fx_header_xls_com_replaces_unit_in_place() -> Result<()> {
    if !runtime_available() {
        eprintln!("skipping: no bundled python runtime");
        return Ok(());
    }
    let python = bundled_python().expect("runtime checked");
    let nonce = next_id();
    let root = std::env::temp_dir().join(format!("pytools-fx-xls-{nonce}"));
    fs::create_dir_all(&root)?;
    let outcome = (|| -> Result<()> {
        let source = root.join("fx.xls");
        make_xls_via_com(&python, &source, "fx")?;
        let argv = [
            "fx-header".to_owned(),
            "--in-place".to_owned(),
            "openpyxl".to_owned(),
            "--view".to_owned(),
            "off".to_owned(),
            "--log-dir".to_owned(),
            root.join("logs").to_string_lossy().into_owned(),
            source.to_string_lossy().into_owned(),
        ];
        run_cli(&argv)?;
        let script = r#"import sys
import win32com.client
sys.stdout.reconfigure(encoding="utf-8")
app = win32com.client.DispatchEx("Excel.Application")
app.Visible = False
app.DisplayAlerts = False
try:
    book = app.Workbooks.Open(sys.argv[1], ReadOnly=True)
    fx = plain = None
    for sheet in book.Worksheets:
        if "外汇" in sheet.Name:
            fx = sheet
        else:
            plain = sheet
    print(fx.PageSetup.RightHeader)
    print(plain.PageSetup.RightHeader)
    book.Close(False)
finally:
    app.Quit()
"#;
        let output = run_com_python(&python, script, &[&source])?;
        let mut lines = output.lines();
        let fx_header = lines.next().unwrap_or_default();
        let plain_header = lines.next().unwrap_or_default();
        ensure!(
            fx_header.contains("万美元"),
            "fx sheet header not converted: {fx_header:?}"
        );
        ensure!(
            !fx_header.contains("单位：万元"),
            "fx sheet header keeps old unit: {fx_header:?}"
        );
        ensure!(
            plain_header.contains("单位：万元") && !plain_header.contains("万美元"),
            "non-fx header must stay untouched: {plain_header:?}"
        );
        Ok(())
    })();
    let _ = fs::remove_dir_all(&root);
    outcome
}
