use rust_xlsxwriter::{DataValidation, DataValidationRule, Workbook};
use std::{collections::BTreeMap, io::Read, path::Path};

fn validation_snapshot(path: &Path) -> Vec<BTreeMap<String, String>> {
    let mut zip = zip::ZipArchive::new(std::fs::File::open(path).unwrap()).unwrap();
    let mut xml = String::new();
    zip.by_name("xl/worksheets/sheet1.xml")
        .unwrap()
        .read_to_string(&mut xml)
        .unwrap();
    let doc = roxmltree::Document::parse(&xml).unwrap();
    doc.descendants()
        .filter(|n| n.has_tag_name("dataValidation"))
        .map(|n| {
            let mut values = BTreeMap::new();
            for key in [
                "sqref",
                "type",
                "operator",
                "promptTitle",
                "prompt",
                "errorTitle",
                "error",
            ] {
                if let Some(v) = n.attribute(key) {
                    values.insert(key.to_owned(), v.to_owned());
                }
            }
            for child in n.children().filter(|c| c.is_element()) {
                values.insert(
                    child.tag_name().name().to_owned(),
                    child.text().unwrap_or("").to_owned(),
                );
            }
            values
        })
        .collect()
}

fn roundtrip(
    rules: &[DataValidation],
) -> (Vec<BTreeMap<String, String>>, Vec<BTreeMap<String, String>>) {
    let dir = std::env::temp_dir().join(format!(
        "rusttools-dv-regression-{}-{}",
        std::process::id(),
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos()
    ));
    std::fs::create_dir(&dir).unwrap();
    let input = dir.join("input.xlsx");
    let output = dir.join("output.xlsx");
    let mut book = Workbook::new();
    let sheet = book.add_worksheet();
    sheet.set_name("Data").unwrap();
    sheet.write_string(0, 0, "synthetic regression").unwrap();
    for (i, rule) in rules.iter().enumerate() {
        sheet
            .add_data_validation(1, i as u16, 15, i as u16, rule)
            .unwrap();
    }
    book.save(&input).unwrap();
    let original_bytes = std::fs::read(&input).unwrap();
    let before = validation_snapshot(&input);
    let loaded = umya_spreadsheet::reader::xlsx::read(&input).unwrap();
    umya_spreadsheet::writer::xlsx::write(&loaded, &output).unwrap();
    let after = validation_snapshot(&output);
    assert_eq!(std::fs::read(&input).unwrap(), original_bytes);
    // Only the two owned temporary files are removed; never source fixtures.
    std::fs::remove_file(input).unwrap();
    std::fs::remove_file(output).unwrap();
    std::fs::remove_dir(dir).unwrap();
    (before, after)
}

#[test]
fn umya_preserves_both_zero_negative_decimal_and_single_validation_bounds() {
    let rules = [
        DataValidation::new().allow_whole_number(DataValidationRule::Between(0, 1000)),
        DataValidation::new().allow_whole_number(DataValidationRule::NotBetween(-9, 99)),
        DataValidation::new().allow_decimal_number(DataValidationRule::Between(-1.5, 2.75)),
        DataValidation::new().allow_whole_number(DataValidationRule::GreaterThan(0)),
        DataValidation::new()
            .allow_list_strings(&["A", "B"])
            .unwrap(),
        DataValidation::new().allow_custom("=AND(A1>0,A1<1000)".into()),
    ];
    let (before, after) = roundtrip(&rules);
    assert_eq!(before.len(), 6);
    assert_eq!(before[0].get("formula1").map(String::as_str), Some("0"));
    assert_eq!(before[0].get("formula2").map(String::as_str), Some("1000"));
    assert_eq!(
        after, before,
        "saved OOXML must preserve validation criteria"
    );
}

#[test]
fn umya_validation_text_attributes_decode_xml_entities_exactly_once() {
    let rule = DataValidation::new()
        .allow_whole_number(DataValidationRule::Between(0, 1000))
        .set_input_title("Title &amp; literal")
        .unwrap()
        .set_input_message("中文 & < > \" ' &#8211; literal")
        .unwrap()
        .set_error_title("Error & < >")
        .unwrap()
        .set_error_message("literal &amp; and &#123; stay literal")
        .unwrap();
    let (before, after) = roundtrip(&[rule]);
    assert_eq!(
        after, before,
        "entities must neither double escape nor double decode"
    );
}
