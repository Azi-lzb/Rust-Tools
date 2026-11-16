use excel_reader::read_workbook;

#[test]
fn biff8_xls_formula_references_preserve_sheet_flags_coordinates_and_comparisons() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/xls_formula_refs.xls");
    let before = std::fs::read(&path).expect("synthetic XLS fixture");
    let spec: serde_json::Value =
        serde_json::from_str(include_str!("fixtures/xls_formula_refs.json")).unwrap();
    let book = read_workbook(&path).expect("read synthetic BIFF8 workbook");
    let sheet = book.sheets.iter().find(|s| s.name == "Calc").unwrap();
    let expected = spec["formulas"].as_array().unwrap();
    assert_eq!(sheet.formulas.len(), expected.len());
    for (i, formula) in expected.iter().enumerate() {
        let address = format!("A{}", i + 1);
        assert_eq!(
            sheet.formulas.get(&address).map(String::as_str),
            formula.as_str(),
            "BIFF8 formula at {address}"
        );
    }
    assert_eq!(std::fs::read(path).unwrap(), before);
}
