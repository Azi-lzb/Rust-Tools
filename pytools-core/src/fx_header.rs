//! Candidate plan builder for PyCode's ordinary FX header unit correction.
//!
//! The module only describes the business operation. Existing-workbook edits
//! are executed by the shared ExcelModifier and an evidence-qualified backend.

use anyhow::{bail, Result};
use excel_editor::{
    EditOperation, ExcelEditPlan, PreservationRequirement, WorkbookFeature, WorkbookFormat,
    WorkbookProfile,
};
use excel_reader::read_header_inventory;
use std::{
    collections::BTreeSet,
    path::{Path, PathBuf},
};

const FX_SHEET_KEY: &str = "外汇";
const OLD_UNIT: &str = "万元";
const NEW_UNIT: &str = "万美元";

#[derive(Clone, Debug, PartialEq)]
pub struct FxHeaderInspection {
    pub fx_sheets: usize,
    pub modified_sheets: usize,
    pub plan: Option<ExcelEditPlan>,
}

/// Builds a modification plan for every worksheet whose name contains 外汇.
///
/// The caller must supply a complete workbook feature inventory. The
/// candidate currently accepts only XLSX; XLSM/VBA preservation has not been
/// validated and must remain on the source implementation.
pub fn build_fx_header_plan(
    input: impl AsRef<Path>,
    output: impl AsRef<Path>,
    workbook: WorkbookProfile,
    request_id: impl Into<String>,
) -> Result<FxHeaderInspection> {
    let input = input.as_ref();
    let output = output.as_ref();
    if !has_extension(input, "xlsx") || !has_extension(output, "xlsx") {
        bail!("FX header candidate currently accepts only .xlsx files");
    }
    if workbook.format != WorkbookFormat::Xlsx {
        bail!("FX header candidate requires an XLSX workbook profile");
    }

    let inventory = read_header_inventory(input)?;
    let mut operations = Vec::new();
    let mut fx_sheets = 0;
    for sheet in inventory
        .iter()
        .filter(|sheet| sheet.sheet_name.contains(FX_SHEET_KEY))
    {
        fx_sheets += 1;
        if sheet
            .headers
            .iter()
            .any(|header| header.text.contains(OLD_UNIT))
        {
            operations.push(EditOperation::ReplaceHeaderText {
                sheet: sheet.sheet_name.clone(),
                search: OLD_UNIT.to_owned(),
                replacement: NEW_UNIT.to_owned(),
            });
        }
    }

    let modified_sheets = operations.len();
    let plan = if operations.is_empty() {
        None
    } else {
        let mut features_present = workbook.features_present;
        features_present.insert(WorkbookFeature::HeadersFooters);
        let preserve = BTreeSet::from([
            PreservationRequirement::CellValuesAndTypes,
            PreservationRequirement::FormulaText,
            PreservationRequirement::CommentsAndPositions,
            PreservationRequirement::Styles,
            PreservationRequirement::MergedRanges,
            PreservationRequirement::DefinedNames,
            PreservationRequirement::RowAndColumnDimensions,
            PreservationRequirement::PageSettings,
            PreservationRequirement::HeaderFooterText,
            PreservationRequirement::Relationships,
            PreservationRequirement::UnmodifiedPackageParts,
            PreservationRequirement::SheetIdentityAndOrder,
        ]);
        Some(ExcelEditPlan {
            request_id: request_id.into(),
            input: PathBuf::from(input),
            output: PathBuf::from(output),
            format: WorkbookFormat::Xlsx,
            workbook: WorkbookProfile {
                format: WorkbookFormat::Xlsx,
                inventory_complete: workbook.inventory_complete,
                features_present,
            },
            operations,
            preserve,
            drop_allowed: Default::default(),
        })
    };

    Ok(FxHeaderInspection {
        fx_sheets,
        modified_sheets,
        plan,
    })
}

fn has_extension(path: &Path, expected: &str) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| extension.eq_ignore_ascii_case(expected))
}
