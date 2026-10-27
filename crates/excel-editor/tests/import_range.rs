use anyhow::{Context, Result};
use excel_editor::{
    inspect_xlsx_profile, BackendCapability, CapabilityStatus, EditBackend, EditOperation,
    EditOperationKind, EngineId, RangeStyleMode, UmyaXlsxBackend, WorkbookFormat,
};
use std::{
    collections::BTreeSet,
    fs,
    path::{Path, PathBuf},
    time::{SystemTime, UNIX_EPOCH},
};

fn temp_path(label: &str) -> PathBuf {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("clock after epoch")
        .as_nanos();
    std::env::temp_dir().join(format!(
        "excel_editor_import_range_{}_{}_{}",
        std::process::id(),
        nonce,
        label
    ))
}

fn write_source(path: &Path) -> Result<()> {
    let mut workbook = umya_spreadsheet::new_file();
    let sheet = workbook
        .get_sheet_by_name_mut("Sheet1")
        .context("default worksheet missing")?;
    sheet.set_name("Source");
    sheet.cell_mut("A1").set_value_string("primary");
    sheet.get_style_mut("A1").get_font_mut().set_bold(true);
    sheet
        .cell_mut("C3")
        .set_formula("1+1")
        .set_formula_result_number(2.0);
    sheet.add_merge_cells("A1:B1");
    sheet.row_dimension_mut(8).set_height(22.0).set_hidden(true);
    sheet
        .column_dimension_mut("D")
        .set_width(24.0)
        .set_hidden(true);
    umya_spreadsheet::writer::xlsx::write(&workbook, path)
        .with_context(|| format!("write source fixture {}", path.display()))
}

fn write_target(path: &Path) -> Result<()> {
    let mut workbook = umya_spreadsheet::new_file();
    let sheet = workbook
        .get_sheet_by_name_mut("Sheet1")
        .context("default worksheet missing")?;
    sheet.set_name("Source");
    umya_spreadsheet::writer::xlsx::write(&workbook, path)
        .with_context(|| format!("write target fixture {}", path.display()))
}

#[test]
fn imports_full_sheet_properties_and_dimensions_beyond_cell_range() -> Result<()> {
    let artifact_dir =
        std::env::var_os("EXCEL_EDITOR_IMPORT_RANGE_ARTIFACT_DIR").map(PathBuf::from);
    let (source, target, output) = if let Some(directory) = artifact_dir.as_ref() {
        fs::create_dir_all(directory)?;
        (
            directory.join("source.xlsx"),
            directory.join("target.xlsx"),
            directory.join("rust_output.xlsx"),
        )
    } else {
        (
            temp_path("source.xlsx"),
            temp_path("target.xlsx"),
            temp_path("output.xlsx"),
        )
    };
    if !source.is_file() {
        write_source(&source)?;
    }
    if !target.is_file() {
        write_target(&target)?;
    }

    let capability = BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Unverified,
        format: WorkbookFormat::Xlsx,
        operations: [EditOperationKind::ImportRange].into_iter().collect(),
        supported_features: BTreeSet::new(),
        preserved: BTreeSet::new(),
        platform: "test-only-direct-backend".to_owned(),
        evidence: None,
    };
    let backend = UmyaXlsxBackend::new(capability)?;
    let plan = excel_editor::ExcelEditPlan {
        request_id: "test-import-range".to_owned(),
        input: target.clone(),
        output: output.clone(),
        format: WorkbookFormat::Xlsx,
        workbook: inspect_xlsx_profile(&target)?,
        operations: vec![EditOperation::ImportRange {
            source_workbook: source.clone(),
            source_sheet: "Source".to_owned(),
            source_range: "A1:C3".to_owned(),
            target_sheet: "Source".to_owned(),
            target_start: "A1".to_owned(),
            copy_values: true,
            copy_formulas: false,
            style_mode: RangeStyleMode::Full,
            copy_comments: false,
            copy_merges: true,
            copy_row_heights: true,
            copy_column_widths: true,
            copy_row_hidden: true,
            copy_column_hidden: true,
            copy_outline_levels: true,
            copy_all_row_dimensions: true,
            copy_all_column_dimensions: true,
            materialize_blank_cells: true,
            hide_zero: false,
        }],
        preserve: BTreeSet::new(),
        drop_allowed: Default::default(),
    };
    backend
        .apply(&plan)
        .map_err(|failure| anyhow::anyhow!(failure.message))?;

    let output_workbook = excel_editor::open_xlsx(&output)?;
    let output_sheet = output_workbook
        .get_sheet_by_name("Source")
        .context("extracted worksheet missing")?;
    let source_workbook = excel_editor::open_xlsx(&source)?;
    let source_sheet = source_workbook
        .get_sheet_by_name("Source")
        .context("source worksheet missing")?;

    assert_eq!(
        output_sheet.get_cell("A1").map(|cell| cell.value()),
        source_sheet.get_cell("A1").map(|cell| cell.value())
    );
    assert_eq!(
        output_sheet.get_cell("A1").map(|cell| cell.style()),
        source_sheet.get_cell("A1").map(|cell| cell.style())
    );
    let imported_formula_cache = output_sheet
        .get_cell("C3")
        .context("cached value missing")?;
    assert!(!imported_formula_cache.cell_value().is_formula());
    assert_eq!(imported_formula_cache.value_number(), Some(2.0));
    assert_eq!(
        output_sheet
            .merge_cells()
            .iter()
            .map(|range| range.range())
            .collect::<Vec<_>>(),
        vec!["A1:B1".to_owned()]
    );
    let source_row = source_sheet
        .row_dimension(8)
        .context("source row missing")?;
    let output_row = output_sheet
        .row_dimension(8)
        .context("copied row missing")?;
    assert_eq!(output_row.height(), source_row.height());
    assert_eq!(output_row.hidden(), source_row.hidden());
    let source_column = source_sheet
        .column_dimension("D")
        .context("source column missing")?;
    let output_column = output_sheet
        .column_dimension("D")
        .context("copied column missing")?;
    assert_eq!(output_column.width(), source_column.width());
    assert_eq!(output_column.hidden(), source_column.hidden());
    if artifact_dir.is_some() {
        assert!(!output_sheet
            .column_dimensions()
            .iter()
            .any(|dimension| dimension.col_num() == 1));
    }

    if artifact_dir.is_none() {
        fs::remove_file(source)?;
        fs::remove_file(target)?;
        fs::remove_file(output)?;
    }
    Ok(())
}
