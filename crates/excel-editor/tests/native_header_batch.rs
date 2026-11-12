use excel_editor::{
    EditOperation, ExcelEditPlan, NativeOoxmlValidator, OutputValidator, PreservationRequirement,
    RustNativeOoxmlBackend, WorkbookFormat,
};
use std::{
    collections::BTreeSet,
    fs,
    time::{SystemTime, UNIX_EPOCH},
};

#[test]
fn multiple_header_updates_validate_as_one_batch_and_reject_unplanned_cell_changes(
) -> anyhow::Result<()> {
    let root = std::env::temp_dir().join(format!(
        "native-header-batch-{}-{}",
        std::process::id(),
        SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos()
    ));
    fs::create_dir(&root)?;
    let input = root.join("source.xlsx");
    let output = root.join("candidate.xlsx");
    let mut book = rust_xlsxwriter::Workbook::new();
    for name in ["FX1", "FX2"] {
        let sheet = book.add_worksheet();
        sheet.set_name(name)?.set_header("&R单位：万元");
        sheet.write_string(0, 0, "preserved")?;
    }
    book.save(&input)?;
    let plan = ExcelEditPlan {
        request_id: "multi-header-regression".into(),
        input: input.clone(),
        output: output.clone(),
        format: WorkbookFormat::Xlsx,
        workbook: excel_editor::inspect_xlsx_profile(&input)?,
        operations: ["FX1", "FX2"]
            .into_iter()
            .map(|name| EditOperation::ReplaceHeaderText {
                sheet: name.into(),
                search: "万元".into(),
                replacement: "万美元".into(),
            })
            .collect(),
        preserve: BTreeSet::from([
            PreservationRequirement::CellValuesAndTypes,
            PreservationRequirement::HeaderFooterText,
            PreservationRequirement::UnmodifiedPackageParts,
        ]),
        drop_allowed: Default::default(),
    };
    let edited = RustNativeOoxmlBackend::apply_bytes(&fs::read(&input)?, &plan)?;
    fs::write(&output, &edited)?;
    NativeOoxmlValidator.validate(&plan, &output)?;
    let mut tamper = plan.clone();
    tamper.operations = vec![EditOperation::SetCellValue {
        sheet: "FX2".into(),
        cell: "A1".into(),
        value: excel_editor::CellValue::Text("unexpected".into()),
    }];
    fs::write(
        &output,
        RustNativeOoxmlBackend::apply_bytes(&edited, &tamper)?,
    )?;
    assert!(
        NativeOoxmlValidator.validate(&plan, &output).is_err(),
        "batch validation must still reject unrelated cell changes"
    );
    fs::remove_dir_all(&root)?;
    Ok(())
}
