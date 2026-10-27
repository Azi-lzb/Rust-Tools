use anyhow::{ensure, Context, Result};
use roxmltree::Document;
use rust_xlsxwriter::Workbook;
use std::{
    env, fs,
    io::Read,
    path::PathBuf,
    process::Command,
    time::{SystemTime, UNIX_EPOCH},
};
use zip::ZipArchive;

#[test]
fn cli_recalculates_formula_and_publishes_validated_xlsx() -> Result<()> {
    if !libreoffice_available() {
        if env::var_os("LO_CONVERT_REQUIRE_LO_TEST").is_some() {
            anyhow::bail!("LO_CONVERT_REQUIRE_LO_TEST is set but soffice is unavailable");
        }
        eprintln!("skipping LibreOffice CLI integration: soffice is unavailable");
        return Ok(());
    }

    let root = temp_dir()?;
    fs::create_dir_all(&root)?;
    let input = root.join("formula-input.xlsx");
    let output = root.join("formula-output.xlsx");
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.write_number(0, 0, 2)?;
    sheet.write_number(0, 1, 3)?;
    sheet.write_formula(0, 2, "=SUM(A1:B1)")?;
    workbook.save(&input)?;

    let result = Command::new(env!("CARGO_BIN_EXE_lo-convert"))
        .args([
            "recalculate",
            "--input",
            input.to_str().context("input path is not UTF-8")?,
            "--output",
            output.to_str().context("output path is not UTF-8")?,
            "--engine",
            "lo",
        ])
        .output()?;
    ensure!(
        result.status.success(),
        "recalculate CLI failed: {}",
        String::from_utf8_lossy(&result.stderr)
    );
    let report: serde_json::Value = serde_json::from_slice(&result.stdout)?;
    ensure!(report["status"] == "recalculated");
    ensure!(report["engine"] == "LibreOffice");
    ensure!(report["formula_cache_values_cleared_before_engine"] == 1);
    ensure!(report["validation"]["formula_text_preserved"] == true);
    ensure!(report["validation"]["cache_nodes_complete"] == true);
    ensure!(output.is_file());

    let mut zip = ZipArchive::new(fs::File::open(&output)?)?;
    let mut worksheet_xml = String::new();
    zip.by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut worksheet_xml)?;
    let document = Document::parse(&worksheet_xml)?;
    let formula_cell = document
        .descendants()
        .find(|node| {
            node.is_element() && node.tag_name().name() == "c" && node.attribute("r") == Some("C1")
        })
        .context("recalculated formula cell C1 missing")?;
    let cached_value = formula_cell
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == "v")
        .and_then(|node| node.text())
        .context("recalculated formula cell C1 has no cache value")?
        .parse::<f64>()?;
    ensure!(
        (cached_value - 5.0).abs() < f64::EPSILON,
        "expected the formula cache to equal 5, got {cached_value}; fixture retained at {}",
        root.display()
    );

    fs::remove_dir_all(root)?;
    Ok(())
}

fn libreoffice_available() -> bool {
    Command::new("soffice")
        .arg("--version")
        .output()
        .is_ok_and(|result| result.status.success())
}

fn temp_dir() -> Result<PathBuf> {
    let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
    Ok(env::temp_dir().join(format!(
        "lo_convert_integration_{}_{}",
        std::process::id(),
        nonce
    )))
}
