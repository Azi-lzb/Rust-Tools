//! `输出附带计划文件` global toggle: plan JSON files are development-time
//! review artifacts; actual use may disable them and keep only the core
//! candidate/target outputs.

use anyhow::Result;
use std::{fs, path::Path, process::Command};

struct Dir(std::path::PathBuf);
impl Dir {
    fn new(tag: &str) -> Result<Self> {
        let base = std::env::temp_dir().join(format!(
            "pytools-plan-toggle-{}-{}",
            tag,
            std::process::id()
        ));
        let _ = fs::remove_dir_all(&base);
        fs::create_dir_all(&base)?;
        Ok(Dir(base))
    }
    fn path(&self, relative: &str) -> std::path::PathBuf {
        self.0.join(relative)
    }
}
impl Drop for Dir {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

fn make_workbooks(dir: &Dir, toggle: Option<&str>) -> Result<()> {
    use rust_xlsxwriter::Workbook;
    let mut source = Workbook::new();
    let sheet = source.add_worksheet();
    sheet.set_name("数据")?;
    sheet.write_string(0, 0, "机构")?;
    sheet.write_string(0, 1, "金额")?;
    sheet.write_string(1, 0, "甲")?;
    sheet.write_number(1, 1, 1)?;
    sheet.write_string(2, 0, "甲")?;
    sheet.write_number(2, 1, 2)?;
    source.save(dir.path("source.xlsx"))?;

    let mut target = Workbook::new();
    let sheet = target.add_worksheet();
    sheet.set_name("汇总")?;
    sheet.write_string(0, 0, "机构")?;
    sheet.write_string(0, 1, "金额")?;
    sheet.write_string(1, 0, "乙")?;
    sheet.write_number(1, 1, 3)?;
    target.save(dir.path("target.xlsx"))?;

    let mut config = Workbook::new();
    let global = config.add_worksheet();
    global.set_name("全局配置")?;
    global.write_string(0, 0, "配置项")?;
    global.write_string(0, 1, "值")?;
    global.write_string(1, 0, "输入目录")?;
    global.write_string(1, 1, ".")?;
    global.write_string(2, 0, "输出目录")?;
    global.write_string(2, 1, dir.path("output").to_string_lossy().into_owned())?;
    if let Some(value) = toggle {
        global.write_string(3, 0, "输出附带计划文件")?;
        global.write_string(3, 1, value)?;
    }
    let tasks = config.add_worksheet();
    tasks.set_name("去重追加数据配置")?;
    let headers = [
        "任务名",
        "源工作簿",
        "源工作表",
        "键列",
        "目标工作簿",
        "目标工作表",
        "写入模式",
    ];
    for (column, header) in headers.iter().enumerate() {
        tasks.write_string(0, column as u16, *header)?;
    }
    tasks.write_string(1, 0, "追加")?;
    tasks.write_string(1, 1, dir.path("source.xlsx").to_string_lossy().into_owned())?;
    tasks.write_string(1, 2, "数据")?;
    tasks.write_string(1, 3, "A")?;
    tasks.write_string(1, 4, dir.path("target.xlsx").to_string_lossy().into_owned())?;
    tasks.write_string(1, 5, "汇总")?;
    tasks.write_string(1, 6, "追加")?;
    config.save(dir.path("config.xlsx"))?;
    Ok(())
}

fn run_plan_cli(config: &Path, output_dir: &Path) -> Result<serde_json::Value> {
    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-append-plan", "--config"])
        .arg(config)
        .args(["--output-dir"])
        .arg(output_dir)
        .output()?;
    assert!(output.status.success());
    let text = String::from_utf8_lossy(&output.stdout);
    Ok(serde_json::from_str(text.trim())?)
}

fn json_files_under(root: &Path) -> usize {
    fn walk(dir: &Path, count: &mut usize) {
        let Ok(entries) = fs::read_dir(dir) else {
            return;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                walk(&path, count);
            } else if path.extension().is_some_and(|ext| ext == "json") {
                *count += 1;
            }
        }
    }
    let mut count = 0;
    walk(root, &mut count);
    count
}

#[test]
fn plan_files_toggle_controls_plan_artifacts_only() -> Result<()> {
    // default (key absent): plan JSON emitted and reported
    let on_dir = Dir::new("on")?;
    make_workbooks(&on_dir, None)?;
    let report = run_plan_cli(&on_dir.path("config.xlsx"), &on_dir.path("out_on"))?;
    let entries = report.as_array().context("array report")?;
    assert_eq!(entries.len(), 1);
    assert!(
        entries[0]["plan_file"].is_string(),
        "default must keep plan files: {}",
        entries[0]
    );
    assert!(json_files_under(&on_dir.path("out_on")) >= 1);

    // disabled: no plan JSON anywhere, plan_file null, candidate still produced
    let off_dir = Dir::new("off")?;
    make_workbooks(&off_dir, Some("否"))?;
    let report = run_plan_cli(&off_dir.path("config.xlsx"), &off_dir.path("out_off"))?;
    let entries = report.as_array().context("array report")?;
    assert_eq!(entries.len(), 1);
    assert!(
        entries[0]["plan_file"].is_null(),
        "disabled must not report a plan file: {}",
        entries[0]
    );
    assert_eq!(
        json_files_under(&off_dir.path("out_off")),
        0,
        "disabled must not write plan JSON files"
    );
    assert!(
        entries[0]["candidate"].is_string(),
        "candidate copy stays a core output"
    );
    Ok(())
}

use anyhow::Context as _;
