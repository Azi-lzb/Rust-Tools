use excel_reader::{read_workbook, read_workbook_with_unknown_error_tokens_as_text};
use serde_json::json;
use std::{fs::File, io::Write, path::PathBuf};
use zip::{write::SimpleFileOptions, ZipWriter};

fn fixture() -> PathBuf {
    let nonce = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap()
        .as_nanos();
    let path = std::env::temp_dir().join(format!(
        "reader_unknown_error_{}_{}.xlsx",
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
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Errors" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
        ),
        (
            "xl/worksheets/sheet1.xml",
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="e"><v>#DATA!</v></c></row></sheetData></worksheet>"#,
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
fn opt_in_reader_normalizes_only_unsupported_literal_error_tokens() {
    let path = fixture();
    let original = std::fs::read(&path).expect("read fixture bytes");

    let default_error = read_workbook(&path).expect_err("default reader remains fail-closed");
    assert!(format!("{default_error:#}").contains("Unsupported cell error value '#DATA!'"));

    let dump =
        read_workbook_with_unknown_error_tokens_as_text(&path).expect("opt-in compatibility read");
    assert!(dump.compatibility_normalization_used);
    assert_eq!(dump.normalized_cell_count, 1);
    let errors = dump
        .workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "Errors")
        .unwrap();
    assert_eq!(
        errors.values.get("A1"),
        Some(&json!({"type":"string", "value":"#DATA!"}))
    );
    assert_eq!(
        std::fs::read(&path).expect("read source bytes after call"),
        original
    );
    std::fs::remove_file(path).expect("remove temporary fixture");
}
