use anyhow::{Context, Result};
use pytools_core::static_print_engines::{run_configured_static_print, StaticPrintEngine};
use rust_xlsxwriter::Workbook;
use std::{fs, path::PathBuf, process::Command};

struct Scratch(PathBuf);
impl Scratch {
    fn new() -> Result<Self> {
        let dir = std::env::temp_dir().join(format!(
            "rust-tools-current-workflows-{}-{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)?
                .as_nanos()
        ));
        fs::create_dir_all(&dir)?;
        Ok(Self(dir))
    }
}
impl Drop for Scratch {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

#[test]
fn print_uses_configured_target_groups_sheets_and_refuses_existing_file() -> Result<()> {
    let temp = Scratch::new()?;
    let source = temp.0.join("source.xlsx");
    let target = temp.0.join("configured.xlsx");
    let config = temp.0.join("config.xlsx");
    let mut book = Workbook::new();
    book.add_worksheet()
        .set_name("One")?
        .write_string(0, 0, "first")?;
    book.add_worksheet()
        .set_name("Two")?
        .write_string(0, 0, "second")?;
    book.save(&source)?;
    let source_before = fs::read(&source)?;
    let mut settings = Workbook::new();
    let sheet = settings.add_worksheet().set_name("打印配置")?;
    for (column, header) in [
        "是否启用",
        "打印模式",
        "源工作簿",
        "源工作表",
        "源工作表打印区域",
        "目标工作簿",
        "目标工作表",
    ]
    .iter()
    .enumerate()
    {
        sheet.write_string(0, column as u16, *header)?;
    }
    for (row, name) in [(1, "One"), (2, "Two")] {
        for (col, value) in [
            (0, "Y".to_owned()),
            (1, "1".to_owned()),
            (2, source.to_string_lossy().into_owned()),
            (3, name.to_owned()),
            (4, "A1:A1".to_owned()),
            (5, target.to_string_lossy().into_owned()),
            (6, name.to_owned()),
        ] {
            sheet.write_string(row, col, value)?;
        }
    }
    settings.save(&config)?;
    let run = run_configured_static_print(&config, StaticPrintEngine::Umya)?;
    assert_eq!((run.task_ok, run.task_failed, run.outputs.len()), (2, 0, 1));
    assert_eq!(run.outputs[0], target);
    let output = excel_reader::read_workbook(&target)?;
    assert_eq!(output.sheets.len(), 2);
    assert_eq!(output.sheets[0].values["A1"]["value"], "first");
    assert_eq!(output.sheets[1].values["A1"]["value"], "second");
    let before = fs::read(&target)?;
    for engine in [StaticPrintEngine::Umya, StaticPrintEngine::Openpyxl] {
        let rejected = run_configured_static_print(&config, engine)?;
        assert_eq!((rejected.task_ok, rejected.task_failed), (0, 2));
        assert!(rejected.outputs.is_empty());
        assert_eq!(fs::read(&target)?, before);
    }
    assert_eq!(fs::read(&source)?, source_before);
    assert!(!temp.0.join("3-4-umya").exists());
    assert!(!temp.0.join("3-5-openpyxl").exists());
    Ok(())
}

#[test]
fn foreign_bank_in_place_deletes_only_matching_rows_without_copy_or_plan() -> Result<()> {
    let temp = Scratch::new()?;
    let source = temp.0.join("source.xlsx");
    let config = temp.0.join("config.xlsx");
    let mut book = Workbook::new();
    let sheet = book.add_worksheet().set_name("分机构")?;
    sheet.write_string(0, 0, "机构")?;
    sheet.write_string(1, 0, "外资甲")?;
    sheet.write_number(1, 1, 10.0)?;
    sheet.write_string(2, 0, "普通乙")?;
    sheet.write_number(2, 1, 20.0)?;
    book.add_worksheet()
        .set_name("其他")?
        .write_string(0, 0, "外资甲")?;
    book.save(&source)?;
    let mut settings = Workbook::new();
    let sheet = settings.add_worksheet().set_name("机构映射表")?;
    for (col, value) in [(0, "序号"), (1, "机构名称"), (2, "是否为外资行")] {
        sheet.write_string(0, col, value)?;
    }
    sheet.write_number(1, 0, 1.0)?;
    sheet.write_string(1, 1, "外资甲")?;
    sheet.write_number(1, 2, 1.0)?;
    settings.save(&config)?;
    let output = Command::new(env!("CARGO_BIN_EXE_RustTools"))
        .args(["remove-foreign-bank", "--in-place", "umya", "--config"])
        .arg(&config)
        .arg(&source)
        .output()?;
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report: serde_json::Value = serde_json::from_slice(&output.stdout)?;
    assert_eq!(report["mode"], "in_place");
    assert_eq!(report["failed_files"], 0);
    assert_eq!(report["deleted_rows"], 1);
    let book = excel_reader::read_workbook(&source)?;
    let branch = book
        .sheets
        .iter()
        .find(|s| s.name == "分机构")
        .context("branch missing")?;
    assert_eq!(branch.values["A2"]["value"], "普通乙");
    assert_eq!(branch.values["B2"]["value"], 20.0);
    let other = book
        .sheets
        .iter()
        .find(|s| s.name == "其他")
        .context("other missing")?;
    assert_eq!(other.values["A1"]["value"], "外资甲");
    assert!(!temp.0.join("source(金融局).xlsx").exists());
    assert!(!temp.0.join("PyTools_专用报表处理结果").exists());
    Ok(())
}
