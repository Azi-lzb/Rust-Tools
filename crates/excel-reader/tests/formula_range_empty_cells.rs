use excel_reader::read_workbook;
use std::{fs::File, io::Write};
use zip::{write::SimpleFileOptions, ZipWriter};

fn fixture() -> std::path::PathBuf {
    let nonce = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap()
        .as_nanos();
    let path = std::env::temp_dir().join(format!(
        "reader_formula_range_{}_{}.xlsx",
        std::process::id(),
        nonce
    ));
    let mut zip = ZipWriter::new(File::create(&path).expect("create xlsx fixture"));
    let entries = [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#,
        ),
        (
            "_rels/.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
        ),
        (
            "xl/workbook.xml",
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Formulas" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
        ),
        (
            "xl/worksheets/sheet1.xml",
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>1+1</f><v>2</v></c><c r="D1"><f>2+2</f><v>4</v></c><c r="G1"><f>3+3</f><v>6</v></c><c r="J1"><f>4+4</f><v>8</v></c></row></sheetData></worksheet>"#,
        ),
    ];
    for (name, contents) in entries {
        zip.start_file(name, SimpleFileOptions::default())
            .expect("add xlsx part");
        zip.write_all(contents.as_bytes()).expect("write xlsx part");
    }
    zip.finish().expect("finish xlsx fixture");
    path
}

#[test]
fn formula_dump_omits_empty_cells_inside_calamine_formula_range() {
    let path = fixture();
    let dump = read_workbook(&path).expect("read xlsx fixture");
    let formulas = &dump.sheets[0].formulas;

    assert_eq!(formulas.len(), 4);
    assert_eq!(formulas.get("A1").map(String::as_str), Some("1+1"));
    assert_eq!(formulas.get("D1").map(String::as_str), Some("2+2"));
    assert_eq!(formulas.get("G1").map(String::as_str), Some("3+3"));
    assert_eq!(formulas.get("J1").map(String::as_str), Some("4+4"));
    assert!(formulas.values().all(|formula| !formula.trim().is_empty()));

    std::fs::remove_file(path).expect("remove temporary fixture");
}
