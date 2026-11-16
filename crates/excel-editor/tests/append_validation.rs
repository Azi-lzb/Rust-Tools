use anyhow::Result;
use excel_editor::{
    CellValue, EditOperation, ExcelEditPlan, OpenpyxlXlsxValidator, OutputValidator,
    PreservationRequirement as P, WorkbookFormat,
};
use rust_xlsxwriter::{Format, Formula, Workbook};
use std::{
    collections::BTreeMap,
    fs,
    io::{Read, Write},
    path::Path,
    sync::atomic::{AtomicU64, Ordering},
};

fn fixture(path: &Path, variant: &str) -> Result<()> {
    let mut book = Workbook::new();
    book.define_name(
        "Anchor",
        if variant == "name" {
            "Data!$B$1"
        } else {
            "Data!$A$1"
        },
    )?;
    let sheet = book.add_worksheet();
    sheet.set_name("Data")?;
    let mut font = Format::new()
        .set_font_name(if variant == "source" {
            "Calibri"
        } else {
            "Arial"
        })
        .set_font_size(11.0);
    if variant == "number-format" {
        font = font.set_num_format("0.00%");
    }
    if variant == "appearance" {
        font = font.set_bold().set_font_size(25.0).set_num_format("0.00%");
        sheet.set_row_height(0, 42.0)?;
        sheet.set_column_width(0, 32.0)?;
    }
    sheet.write_string_with_format(0, 0, "old header", &font)?;
    sheet.write_number(1, 1, if variant == "value" { 9 } else { 0 })?;
    sheet.write_number_with_format(
        1,
        3,
        46000.0,
        &Format::new().set_num_format(if variant == "appearance" {
            "General"
        } else {
            "yyyy-mm-dd"
        }),
    )?;
    sheet.write_formula(
        98,
        26,
        Formula::new(if variant == "formula" { "B2+2" } else { "B2+1" })
            .set_result(if variant == "cache" { "999" } else { "1" }),
    )?;
    if variant != "merge" {
        sheet.merge_range(4, 0, 4, 1, "merged title", &Format::new())?;
    }
    if variant != "source" {
        sheet.write_string(99, 0, "append")?;
    }
    book.save(path)?;
    Ok(())
}

fn rewrite_sheet(path: &Path, rewrite: impl FnOnce(String) -> String) -> Result<()> {
    let mut archive = zip::ZipArchive::new(fs::File::open(path)?)?;
    let mut parts = Vec::new();
    for i in 0..archive.len() {
        let mut entry = archive.by_index(i)?;
        let name = entry.name().to_owned();
        let mut data = Vec::new();
        entry.read_to_end(&mut data)?;
        parts.push((name, data));
    }
    drop(archive);
    let mut rewrite = Some(rewrite);
    let mut writer = zip::ZipWriter::new(fs::File::create(path)?);
    for (name, mut data) in parts {
        if name == "xl/worksheets/sheet1.xml" {
            data = rewrite.take().unwrap()(String::from_utf8(data)?).into_bytes();
        }
        writer.start_file(name, zip::write::SimpleFileOptions::default())?;
        writer.write_all(&data)?;
    }
    writer.finish()?;
    Ok(())
}

#[test]
fn content_validation_ignores_shared_and_array_results_but_checks_formula_meaning() -> Result<()> {
    let dir = std::env::temp_dir().join(format!(
        "rusttools-shared-cache-contract-{}",
        std::process::id()
    ));
    fs::create_dir_all(&dir)?;
    let source = dir.join("source.xlsx");
    let candidate = dir.join("candidate.xlsx");
    for (path, cached) in [(&source, "3"), (&candidate, "999")] {
        let mut book = Workbook::new();
        let sheet = book.add_worksheet();
        sheet.set_name("Data")?;
        sheet.write_number(0, 0, 7)?;
        sheet.write_formula(8, 3, Formula::new("B9+C9").set_result(cached))?;
        sheet.write_formula(9, 3, Formula::new("B10+C10").set_result(cached))?;
        sheet.write_array_formula(8, 4, 9, 4, Formula::new("B9:B10*2").set_result(cached))?;
        if path == &candidate {
            sheet.write_string(10, 0, "append")?;
        }
        book.save(path)?;
    }
    rewrite_sheet(&source, |xml| {
        assert!(xml.contains("<f>B9+C9</f>") && xml.contains("<f>B10+C10</f>"));
        xml.replace(
            "<f>B9+C9</f>",
            "<f t=\"shared\" si=\"0\" ref=\"D9:D10\">B9+C9</f>",
        )
        .replace("<f>B10+C10</f>", "<f t=\"shared\" si=\"0\"/>")
    })?;
    rewrite_sheet(&candidate, |xml| {
        assert!(xml.contains("<c r=\"E10\"><v>0</v></c>"));
        xml.replace("<c r=\"E10\"><v>0</v></c>", "<c r=\"E10\"><v>123</v></c>")
    })?;
    let plan = ExcelEditPlan {
        request_id: "shared-cache-contract".into(),
        input: source.clone(),
        output: dir.join("published.xlsx"),
        format: WorkbookFormat::Xlsx,
        workbook: excel_editor::inspect_xlsx_profile(&source)?,
        operations: vec![EditOperation::SetCellValue {
            sheet: "Data".into(),
            cell: "A11".into(),
            value: CellValue::Text("append".into()),
        }],
        preserve: [
            P::CellValuesAndTypes,
            P::FormulaText,
            P::SheetIdentityAndOrder,
        ]
        .into_iter()
        .collect(),
        drop_allowed: BTreeMap::new(),
    };
    let original = fs::read(&source)?;
    OpenpyxlXlsxValidator.validate(&plan, &candidate)?;
    rewrite_sheet(&candidate, |xml| {
        xml.replace("<f>B10+C10</f>", "<f>B10+C10+1</f>")
    })?;
    assert!(OpenpyxlXlsxValidator.validate(&plan, &candidate).is_err());
    rewrite_sheet(&candidate, |xml| {
        xml.replace("<f>B10+C10+1</f>", "<f>B10+C10</f>")
            .replace("ref=\"E9:E10\"", "ref=\"E9:E9\"")
    })?;
    assert!(OpenpyxlXlsxValidator.validate(&plan, &candidate).is_err());
    assert_eq!(fs::read(&source)?, original);
    assert!(!plan.output.exists());
    Ok(())
}

#[test]
fn content_validation_accepts_appearance_and_cache_drift_but_rejects_real_loss() -> Result<()> {
    let dir = std::env::temp_dir().join(format!(
        "rusttools-content-validation-{}",
        std::process::id()
    ));
    fs::create_dir_all(&dir)?;
    let source = dir.join("source.xlsx");
    fixture(&source, "source")?;
    let before = fs::read(&source)?;
    let plan = ExcelEditPlan {
        request_id: "content-contract".into(),
        input: source.clone(),
        output: dir.join("published.xlsx"),
        format: WorkbookFormat::Xlsx,
        workbook: excel_editor::inspect_xlsx_profile(&source)?,
        operations: vec![EditOperation::SetCellValue {
            sheet: "Data".into(),
            cell: "A100".into(),
            value: CellValue::Text("append".into()),
        }],
        preserve: [
            P::CellValuesAndTypes,
            P::FormulaText,
            P::MergedRanges,
            P::DefinedNames,
            P::SheetIdentityAndOrder,
        ]
        .into_iter()
        .collect(),
        drop_allowed: BTreeMap::new(),
    };
    for variant in [
        "appearance",
        "number-format",
        "cache",
        "value",
        "formula",
        "merge",
        "name",
    ] {
        let candidate = dir.join(format!("{variant}.xlsx"));
        fixture(&candidate, variant)?;
        let result = OpenpyxlXlsxValidator.validate(&plan, &candidate);
        assert_eq!(
            result.is_ok(),
            matches!(variant, "appearance" | "number-format" | "cache"),
            "{variant}: {result:?}"
        );
        assert_eq!(fs::read(&source)?, before);
        assert!(!plan.output.exists());
    }
    Ok(())
}

#[test]
fn append_candidate_accepts_cosmetics_and_rejects_content_or_structure_loss() -> Result<()> {
    static SEQUENCE: AtomicU64 = AtomicU64::new(0);
    let dir = std::env::temp_dir().join(format!(
        "rusttools-append-validation-{}-{}",
        std::process::id(),
        SEQUENCE.fetch_add(1, Ordering::Relaxed)
    ));
    fs::create_dir(&dir)?;
    let source = dir.join("source.xlsx");
    fixture(&source, "source")?;
    let original = fs::read(&source)?;
    let plan = ExcelEditPlan {
        request_id: "synthetic-append-validation".into(),
        input: source.clone(),
        output: dir.join("published.xlsx"),
        format: WorkbookFormat::Xlsx,
        workbook: excel_editor::inspect_xlsx_profile(&source)?,
        operations: vec![EditOperation::SetCellValue {
            sheet: "Data".into(),
            cell: "A100".into(),
            value: CellValue::Text("append".into()),
        }],
        preserve: [
            P::CellValuesAndTypes,
            P::FormulaText,
            P::Styles,
            P::MergedRanges,
            P::DefinedNames,
            P::SheetIdentityAndOrder,
        ]
        .into_iter()
        .collect(),
        drop_allowed: BTreeMap::new(),
    };
    for variant in [
        "cosmetic",
        "number-format",
        "value",
        "formula",
        "merge",
        "name",
    ] {
        let candidate = dir.join(format!("{variant}.xlsx"));
        fixture(&candidate, variant)?;
        let result = OpenpyxlXlsxValidator.validate_with_notes(&plan, &candidate);
        if variant == "cosmetic" {
            let notes = result?;
            assert!(notes
                .iter()
                .any(|n| n.contains("cosmetic normalization") && n.contains("A1")));
        } else {
            assert!(result.is_err(), "corruption accepted: {variant}");
        }
        assert_eq!(fs::read(&source)?, original);
        assert!(!plan.output.exists());
    }
    // Leave synthetic fixtures available for diagnosis in the configured TEMP.
    Ok(())
}
