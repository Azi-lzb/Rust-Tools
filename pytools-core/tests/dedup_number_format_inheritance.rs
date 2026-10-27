use anyhow::Result;
use pytools_core::{
    config::DedupTask,
    dedup_ops::{
        build_config_duplicate_plan, write_config_duplicate_delete_candidate, DuplicateAction,
    },
};
use rust_xlsxwriter::{Format, Workbook};
use std::{
    collections::BTreeMap,
    fs,
    io::Write,
    path::{Path, PathBuf},
    time::{SystemTime, UNIX_EPOCH},
};
use zip::{write::SimpleFileOptions, ZipWriter};

struct TestDir(PathBuf);

impl TestDir {
    fn new() -> Self {
        let path = std::env::temp_dir().join(format!(
            "pytools-number-format-repro-{}-{}",
            std::process::id(),
            SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        fs::create_dir_all(&path).unwrap();
        Self(path)
    }

    fn path(&self, name: &str) -> PathBuf {
        self.0.join(name)
    }
}

impl Drop for TestDir {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

#[test]
fn configured_delete_preserves_numeric_zero_padding_format() -> Result<()> {
    let dir = TestDir::new();
    let input = dir.path("input.xlsx");
    let output = dir.path("output.xlsx");

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("Data")?;
    let code_format = Format::new().set_num_format("000000");
    worksheet.write_string(0, 0, "机构代码")?;
    worksheet.write_string(0, 1, "复合键")?;
    worksheet.write_number_with_format(1, 0, 123.0, &code_format)?;
    worksheet.write_string(1, 1, "同键")?;
    worksheet.write_number_with_format(2, 0, 456.0, &code_format)?;
    worksheet.write_string(2, 1, "同键")?;
    worksheet.write_number_with_format(3, 0, 789.0, &code_format)?;
    worksheet.write_string(3, 1, "唯一键")?;
    workbook.save(&input)?;
    let input_bytes = fs::read(&input)?;

    let read_format = |path: &PathBuf, address: &str| -> Result<Option<String>> {
        let book = excel_editor::open_xlsx(path)?;
        let format = book
            .sheet_by_name("Data")?
            .cell(address)
            .and_then(|cell| cell.style().number_format())
            .map(|format| format.format_code().to_owned());
        Ok(format)
    };
    assert_eq!(read_format(&input, "A2")?.as_deref(), Some("000000"));
    assert_eq!(read_format(&input, "A4")?.as_deref(), Some("000000"));

    let task = DedupTask {
        row_no: 2,
        enabled: true,
        task_name: "synthetic".into(),
        source_wb: Some(input.clone()),
        source_ws: "Data".into(),
        key_cols: vec![2],
        target_wb: None,
        target_ws: String::new(),
        write_mode: String::new(),
    };
    let outputs = BTreeMap::from([(input.clone(), output.clone())]);
    let reports = build_config_duplicate_plan(&[task], &outputs, DuplicateAction::Delete)?;
    let plan = reports[0].plan.as_ref();
    write_config_duplicate_delete_candidate(&input, &output, plan)?;

    assert_eq!(read_format(&output, "A2")?.as_deref(), Some("000000"));
    assert_eq!(read_format(&output, "A3")?.as_deref(), Some("000000"));
    assert_eq!(fs::read(&input)?, input_bytes);
    Ok(())
}

#[test]
fn configured_delete_preserves_inherited_number_formats() -> Result<()> {
    let dir = TestDir::new();

    for source in ["row", "column"] {
        let input = dir.path(&format!("input-{source}.xlsx"));
        let output = dir.path(&format!("output-{source}.xlsx"));
        create_inherited_style_xlsx(&input, source == "row")?;
        let input_bytes = fs::read(&input)?;

        let task = DedupTask {
            row_no: 2,
            enabled: true,
            task_name: format!("synthetic-{source}"),
            source_wb: Some(input.clone()),
            source_ws: "Data".into(),
            key_cols: vec![2],
            target_wb: None,
            target_ws: String::new(),
            write_mode: String::new(),
        };
        let outputs = BTreeMap::from([(input.clone(), output.clone())]);
        let reports = build_config_duplicate_plan(&[task], &outputs, DuplicateAction::Delete)?;
        write_config_duplicate_delete_candidate(&input, &output, reports[0].plan.as_ref())?;
        let read_format = |path: &PathBuf, address: &str| -> Result<Option<String>> {
            let book = excel_editor::open_xlsx(path)?;
            let format = book
                .sheet_by_name("Data")?
                .cell(address)
                .and_then(|cell| cell.style().number_format())
                .map(|format| format.format_code().to_owned());
            Ok(format)
        };
        assert_eq!(
            read_format(&output, "A2")?.as_deref(),
            Some("000000"),
            "{source} inherited style was not preserved for the first retained row"
        );
        assert_eq!(
            read_format(&output, "A3")?.as_deref(),
            Some("000000"),
            "{source} inherited style was not preserved for the row moved up after deletion"
        );
        assert_eq!(
            fs::read(&input)?,
            input_bytes,
            "{source} styled source changed"
        );
    }
    Ok(())
}

fn create_inherited_style_xlsx(path: &Path, inherited_on_row: bool) -> Result<()> {
    const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
    const PACKAGE_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
    const OFFICE_REL_NS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const SPREADSHEET_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#
    );
    let package_relationships = format!(
        r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId0" Type="{OFFICE_REL_NS}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
    );
    let workbook = format!(
        r#"<s:workbook xmlns:s="{SPREADSHEET_NS}" xmlns:r="{OFFICE_REL_NS}"><s:sheets><s:sheet name="Data" sheetId="1" r:id="rId1"/></s:sheets></s:workbook>"#
    );
    let workbook_relationships = format!(
        r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId1" Type="{OFFICE_REL_NS}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{OFFICE_REL_NS}/styles" Target="styles.xml"/></Relationships>"#
    );
    let styles = format!(
        r#"<s:styleSheet xmlns:s="{SPREADSHEET_NS}"><s:numFmts count="1"><s:numFmt numFmtId="165" formatCode="000000"/></s:numFmts><s:cellXfs count="2"><s:xf numFmtId="0"/><s:xf numFmtId="165"/></s:cellXfs></s:styleSheet>"#
    );
    let (column_style, row_style, final_row_style) = if inherited_on_row {
        (
            "",
            " s=\"1\" customFormat=\"1\"",
            " s=\"1\" customFormat=\"1\"",
        )
    } else {
        (
            "<s:cols><s:col min=\"1\" max=\"1\" style=\"1\"/></s:cols>",
            "",
            "",
        )
    };
    let worksheet = format!(
        r#"<s:worksheet xmlns:s="{SPREADSHEET_NS}">{column_style}<s:sheetData><s:row r="1"><s:c r="A1" t="inlineStr"><s:is><s:t>机构代码</s:t></s:is></s:c><s:c r="B1" t="inlineStr"><s:is><s:t>复合键</s:t></s:is></s:c></s:row><s:row r="2"{row_style}><s:c r="A2"><s:v>123</s:v></s:c><s:c r="B2" t="inlineStr"><s:is><s:t>同键</s:t></s:is></s:c></s:row><s:row r="3"><s:c r="A3"><s:v>456</s:v></s:c><s:c r="B3" t="inlineStr"><s:is><s:t>同键</s:t></s:is></s:c></s:row><s:row r="4"{final_row_style}><s:c r="A4"><s:v>789</s:v></s:c><s:c r="B4" t="inlineStr"><s:is><s:t>唯一键</s:t></s:is></s:c></s:row></s:sheetData></s:worksheet>"#
    );
    let entries = [
        ("[Content_Types].xml", content_types),
        ("_rels/.rels", package_relationships),
        ("xl/workbook.xml", workbook),
        ("xl/_rels/workbook.xml.rels", workbook_relationships),
        ("xl/styles.xml", styles),
        ("xl/worksheets/sheet1.xml", worksheet),
    ];
    let mut archive = ZipWriter::new(fs::File::create(path)?);
    for (name, data) in entries {
        archive.start_file(name, SimpleFileOptions::default())?;
        archive.write_all(data.as_bytes())?;
    }
    archive.finish()?;
    Ok(())
}
