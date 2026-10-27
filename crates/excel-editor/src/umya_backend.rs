//! XLSX edit candidate implemented with umya-spreadsheet.
//!
//! This adapter deliberately does not promote its own capabilities. Callers
//! must provide a capability record backed by matching preservation evidence;
//! the router rejects `Unverified` registrations.

use crate::{
    BackendCapability, BackendFailure, EditBackend, EditOperation, ExcelEditPlan, OutputValidator,
    PageOrientation, RangeStyleMode, Workbook, WorkbookFormat,
};
use anyhow::{bail, Context, Result};
use excel_reader::read_workbook;
use std::path::Path;
use umya_spreadsheet::Style;

pub struct UmyaXlsxBackend {
    capability: BackendCapability,
}

impl UmyaXlsxBackend {
    pub fn new(capability: BackendCapability) -> Result<Self> {
        anyhow::ensure!(
            capability.engine == crate::EngineId::UmyaSpreadsheet,
            "Umya backend requires an umya capability record"
        );
        anyhow::ensure!(
            capability.format == WorkbookFormat::Xlsx,
            "Umya candidate supports XLSX only"
        );
        Ok(Self { capability })
    }
}

impl EditBackend for UmyaXlsxBackend {
    fn capability(&self) -> &BackendCapability {
        &self.capability
    }

    fn apply(&self, plan: &ExcelEditPlan) -> std::result::Result<(), BackendFailure> {
        if plan.format != WorkbookFormat::Xlsx {
            return Err(BackendFailure::unsupported_before_write(
                "umya candidate supports XLSX only",
            ));
        }
        ooxml_patch::reject_unverified_parts(&plan.input).map_err(|error| {
            BackendFailure::unsupported_before_write(format!(
                "unverified workbook structure: {error:#}"
            ))
        })?;
        for operation in &plan.operations {
            if let EditOperation::ImportRange {
                source_workbook, ..
            } = operation
            {
                if !source_workbook.is_file() {
                    return Err(BackendFailure::unsupported_before_write(format!(
                        "range import source does not exist: {}",
                        source_workbook.display()
                    )));
                }
                if !source_workbook
                    .extension()
                    .and_then(|ext| ext.to_str())
                    .is_some_and(|ext| ext.eq_ignore_ascii_case("xlsx"))
                {
                    return Err(BackendFailure::unsupported_before_write(format!(
                        "range import currently accepts .xlsx sources only: {}",
                        source_workbook.display()
                    )));
                }
                ooxml_patch::reject_unverified_parts(source_workbook).map_err(|error| {
                    BackendFailure::unsupported_before_write(format!(
                        "unverified range import source structure: {error:#}"
                    ))
                })?;
            }
        }
        let created_sheets = plan
            .operations
            .iter()
            .filter_map(|operation| match operation {
                EditOperation::CreateSheet { name } => Some(name.as_str()),
                EditOperation::CreateSheetAt { name, .. } => Some(name.as_str()),
                _ => None,
            })
            .collect::<std::collections::BTreeSet<_>>();
        if let Some(sheet) = plan
            .operations
            .iter()
            .find_map(|operation| match operation {
                EditOperation::SetPageSettings {
                    sheet,
                    settings:
                        crate::PageSettings {
                            clear_scale: true, ..
                        },
                } if !created_sheets.contains(sheet.as_str()) => Some(sheet),
                _ => None,
            })
        {
            return Err(BackendFailure::unsupported_before_write(format!(
                "umya-spreadsheet cannot safely remove an existing pageSetup scale for worksheet {sheet:?}"
            )));
        }

        let mut workbook = crate::open_xlsx(&plan.input)
            .map_err(|error| BackendFailure::failed(format!("open XLSX: {error:#}")))?;
        validate_row_deletion_scope(plan, &workbook)
            .map_err(|error| BackendFailure::unsupported_before_write(format!("{error:#}")))?;
        validate_row_insertion_scope(plan, &workbook)
            .map_err(|error| BackendFailure::unsupported_before_write(format!("{error:#}")))?;
        for operation in &plan.operations {
            if let EditOperation::SetFreezePanes { sheet, cell } = operation {
                if is_freeze_clear_request(cell.as_deref())
                    && sheet_has_frozen_pane(&workbook, sheet)
                {
                    return Err(BackendFailure::unsupported_before_write(format!(
                        "umya-spreadsheet cannot clear an existing frozen pane on worksheet {sheet:?}; route freeze-pane clears to another verified engine"
                    )));
                }
            }
        }
        let outline_copies = plan
            .operations
            .iter()
            .filter_map(|operation| match operation {
                EditOperation::ImportRange {
                    source_workbook,
                    source_sheet,
                    source_range,
                    target_sheet,
                    target_start,
                    copy_outline_levels: true,
                    copy_all_row_dimensions,
                    copy_all_column_dimensions,
                    ..
                } => Some((
                    source_workbook.clone(),
                    source_sheet.clone(),
                    source_range.clone(),
                    target_sheet.clone(),
                    target_start.clone(),
                    *copy_all_row_dimensions,
                    *copy_all_column_dimensions,
                )),
                _ => None,
            })
            .collect::<Vec<_>>();
        for operation in &plan.operations {
            apply_operation(&mut workbook, operation).map_err(|error| {
                BackendFailure::failed(format!("apply {:?}: {error:#}", operation.kind()))
            })?;
        }
        apply_final_print_areas(&mut workbook, &plan.operations).map_err(|error| {
            BackendFailure::failed(format!("apply final worksheet print areas: {error:#}"))
        })?;
        crate::save_xlsx(&workbook, &plan.output)
            .map_err(|error| BackendFailure::failed(format!("save XLSX: {error:#}")))?;
        for (
            source_workbook,
            source_sheet,
            source_range,
            target_sheet,
            target_start,
            copy_all_row_dimensions,
            copy_all_column_dimensions,
        ) in outline_copies
        {
            let outline_range = if copy_all_row_dimensions || copy_all_column_dimensions {
                "A1:XFD1048576"
            } else {
                &source_range
            };
            ooxml_patch::copy_outline_levels(
                source_workbook,
                &plan.output,
                &source_sheet,
                outline_range,
                &target_sheet,
                &target_start,
            )
            .map_err(|error| BackendFailure::failed(format!("copy outline levels: {error:#}")))?;
        }
        if let Some((full_calc_on_load, force_full_calc)) =
            plan.operations.iter().find_map(|operation| {
                if let EditOperation::SetWorkbookCalculationProperties {
                    full_calc_on_load,
                    force_full_calc,
                } = operation
                {
                    Some((*full_calc_on_load, *force_full_calc))
                } else {
                    None
                }
            })
        {
            let patch_output = plan.output.with_file_name(format!(
                ".{}.calcpr-{}.xlsx",
                plan.output
                    .file_name()
                    .and_then(|name| name.to_str())
                    .unwrap_or("candidate"),
                std::process::id()
            ));
            let patch_result = ooxml_patch::patch_workbook_calculation_properties(
                &plan.output,
                &patch_output,
                ooxml_patch::WorkbookCalculationProperties {
                    full_calc_on_load,
                    force_full_calc,
                },
            );
            if let Err(error) = patch_result {
                return Err(BackendFailure::failed(format!(
                    "patch workbook calculation properties: {error:#}"
                )));
            }
            let copy_result = std::fs::copy(&patch_output, &plan.output).with_context(|| {
                "replace disposable XLSX candidate with calculation-property patched package"
                    .to_owned()
            });
            let _ = std::fs::remove_file(&patch_output);
            copy_result.map_err(|error| {
                BackendFailure::failed(format!("install calculation-property candidate: {error:#}"))
            })?;
        }
        Ok(())
    }
}

fn validate_row_deletion_scope(plan: &ExcelEditPlan, workbook: &Workbook) -> Result<()> {
    let deletions = plan
        .operations
        .iter()
        .filter_map(|operation| match operation {
            EditOperation::DeleteRows {
                sheet,
                first_row,
                count,
            } => Some((sheet.as_str(), *first_row, *count)),
            _ => None,
        })
        .collect::<Vec<_>>();
    if deletions.is_empty() {
        return Ok(());
    }

    for (sheet_name, first_row, count) in &deletions {
        anyhow::ensure!(
            *first_row > 0 && *count > 0,
            "row deletion requires positive one-based indexes"
        );
        let last_row = first_row
            .checked_add(*count - 1)
            .context("row deletion range overflows")?;
        anyhow::ensure!(
            last_row <= EXCEL_MAX_ROWS,
            "row deletion range exceeds the XLSX row limit"
        );
        anyhow::ensure!(
            workbook.sheet_by_name(sheet_name).is_ok(),
            "worksheet {sheet_name:?} not found"
        );
    }

    // Run the same reference transformations against cloned formula/name values
    // before any workbook mutation. This makes unsupported structures an
    // explicit pre-write refusal and leaves the original input untouched.
    for (formula_sheet, worksheet) in workbook.sheet_collection().iter().map(|s| (s.name(), s)) {
        let mut formulas = worksheet
            .cells()
            .into_iter()
            .filter(|cell| cell.is_formula())
            .map(|cell| (cell.coordinate().to_string(), cell.cell_value().clone()))
            .collect::<Vec<_>>();
        let mut formula_rows = formulas
            .iter()
            .map(|(address, _)| parse_cell(address).map(|(row, _)| row))
            .collect::<Result<Vec<_>>>()?;
        let mut removed = vec![false; formulas.len()];

        for (target_sheet, first_row, count) in &deletions {
            let last_row = first_row + count - 1;
            let mut applicable_indices = Vec::new();
            for (index, (_, value)) in formulas.iter().enumerate() {
                if removed[index] {
                    continue;
                }
                if formula_sheet == *target_sheet
                    && formula_rows[index] >= *first_row
                    && formula_rows[index] <= last_row
                {
                    removed[index] = true;
                    continue;
                }
                if formula_sheet != *target_sheet
                    && !formula_references_sheet(value.formula(), target_sheet)
                {
                    continue;
                }
                ensure_row_delete_formula_supported(value, formula_sheet, target_sheet)?;
                applicable_indices.push(index);
            }

            if !applicable_indices.is_empty() {
                let batch = applicable_indices
                    .iter()
                    .map(|index| formulas[*index].clone())
                    .collect::<Vec<_>>();
                let adjusted = adjust_formula_batch_for_row_deletion(
                    formula_sheet,
                    target_sheet,
                    *first_row,
                    *count,
                    &batch,
                )?;
                for (index, (_, value)) in applicable_indices.into_iter().zip(adjusted) {
                    formulas[index].1 = value;
                }
            }

            if formula_sheet == *target_sheet {
                for (index, row) in formula_rows.iter_mut().enumerate() {
                    if !removed[index] && *row > last_row {
                        *row -= count;
                    }
                }
            }
        }
    }

    let workbook_names = workbook
        .defined_names()
        .iter()
        .cloned()
        .map(|name| (None, name))
        .chain(workbook.sheet_collection().iter().flat_map(|sheet| {
            sheet
                .defined_names()
                .iter()
                .cloned()
                .map(move |name| (Some(sheet.name().to_owned()), name))
        }))
        .collect::<Vec<_>>();
    for (scope_sheet, mut name) in workbook_names {
        for (target_sheet, first_row, count) in &deletions {
            name = match adjust_defined_name_for_row_deletion(
                &name,
                scope_sheet.as_deref(),
                target_sheet,
                *first_row,
                *count,
            )? {
                Some(updated) => updated,
                None => break,
            };
        }
    }
    Ok(())
}

fn validate_row_insertion_scope(plan: &ExcelEditPlan, workbook: &Workbook) -> Result<()> {
    if !plan
        .operations
        .iter()
        .any(|operation| matches!(operation, EditOperation::InsertRows { .. }))
    {
        return Ok(());
    }
    let mut simulated = workbook.clone();
    for operation in &plan.operations {
        if let EditOperation::InsertRows {
            sheet,
            first_row,
            count,
        } = operation
        {
            insert_rows_with_excel_references(&mut simulated, sheet, *first_row, *count)?;
        }
    }
    Ok(())
}

const EXCEL_MAX_ROWS: u32 = 1_048_576;
const EXCEL_MAX_COLUMNS: u32 = 16_384;

fn delete_rows_with_excel_references(
    workbook: &mut Workbook,
    target_sheet: &str,
    first_row: u32,
    count: u32,
) -> Result<()> {
    let last_row = first_row
        .checked_add(
            count
                .checked_sub(1)
                .context("row deletion count must be positive")?,
        )
        .context("row deletion range overflows")?;
    anyhow::ensure!(
        first_row > 0 && last_row <= EXCEL_MAX_ROWS,
        "row deletion range is outside the XLSX grid"
    );
    let target_index = workbook
        .sheet_collection()
        .iter()
        .position(|sheet| sheet.name() == target_sheet)
        .with_context(|| format!("worksheet {target_sheet:?} not found"))?;

    let mut formula_updates = Vec::new();
    for sheet in workbook.sheet_collection() {
        let formula_sheet = sheet.name().to_owned();
        let formulas = sheet
            .cells()
            .into_iter()
            .filter(|cell| cell.is_formula())
            .map(|cell| (cell.coordinate().to_string(), cell.cell_value().clone()))
            .collect::<Vec<_>>();
        let applicable = formulas
            .into_iter()
            .filter(|(_, value)| {
                formula_sheet == target_sheet
                    || formula_references_sheet(value.formula(), target_sheet)
            })
            .collect::<Vec<_>>();
        if applicable.is_empty() {
            continue;
        }
        let adjusted = adjust_formula_batch_for_row_deletion(
            &formula_sheet,
            target_sheet,
            first_row,
            count,
            &applicable,
        )?;
        formula_updates.push((formula_sheet, adjusted));
    }

    let updated_global_names = workbook
        .defined_names()
        .iter()
        .map(|name| {
            adjust_defined_name_for_row_deletion(name, None, target_sheet, first_row, count)
        })
        .collect::<Result<Vec<_>>>()?
        .into_iter()
        .flatten()
        .collect::<Vec<_>>();
    let updated_local_names = workbook
        .sheet_collection()
        .iter()
        .map(|sheet| {
            let owner = sheet.name().to_owned();
            let names = sheet
                .defined_names()
                .iter()
                .map(|name| {
                    adjust_defined_name_for_row_deletion(
                        name,
                        Some(&owner),
                        target_sheet,
                        first_row,
                        count,
                    )
                })
                .collect::<Result<Vec<_>>>()?
                .into_iter()
                .flatten()
                .collect::<Vec<_>>();
            Ok::<_, anyhow::Error>((owner, names))
        })
        .collect::<Result<Vec<_>>>()?;

    *workbook.defined_names_mut() = updated_global_names;
    for (formula_sheet, formulas) in &formula_updates {
        if formula_sheet == target_sheet {
            continue;
        }
        let sheet = workbook
            .sheet_by_name_mut(formula_sheet)
            .with_context(|| format!("worksheet {formula_sheet:?} disappeared"))?;
        for (address, value) in formulas {
            sheet
                .cell_mut(address.as_str())
                .set_cell_value(value.clone());
        }
    }

    let target_formulas = formula_updates_for_sheet(&formula_updates, target_sheet);
    let target_names = updated_local_names
        .iter()
        .find(|(sheet, _)| sheet == target_sheet)
        .map(|(_, names)| names.clone())
        .unwrap_or_default();
    for (owner, names) in updated_local_names {
        if owner == target_sheet {
            continue;
        }
        *workbook
            .sheet_by_name_mut(&owner)
            .with_context(|| format!("worksheet {owner:?} disappeared"))?
            .defined_names_mut() = names;
    }

    let target = &mut workbook.sheet_collection_mut()[target_index];
    for (address, _) in &target_formulas {
        target.cell_value_mut(address.as_str()).remove_formula();
    }
    target.defined_names_mut().clear();
    target.remove_row(first_row, count);
    *target.defined_names_mut() = target_names;

    for (address, value) in target_formulas {
        let (old_row, column) = parse_cell(&address)?;
        if old_row >= first_row && old_row <= last_row {
            continue;
        }
        let new_row = if old_row > last_row {
            old_row - count
        } else {
            old_row
        };
        target
            .cell_mut(cell_address(new_row, column).as_str())
            .set_cell_value(value);
    }
    Ok(())
}

#[derive(Clone)]
struct InsertedFormula {
    inserted_address: String,
    value: umya_spreadsheet::CellValue,
}

fn insert_rows_with_excel_references(
    workbook: &mut Workbook,
    target_sheet: &str,
    first_row: u32,
    count: u32,
) -> Result<()> {
    let last_inserted_row = first_row
        .checked_add(
            count
                .checked_sub(1)
                .context("row insertion count must be positive")?,
        )
        .context("row insertion range overflows")?;
    anyhow::ensure!(
        first_row > 0 && last_inserted_row <= EXCEL_MAX_ROWS,
        "row insertion range is outside the XLSX grid"
    );
    let target_index = workbook
        .sheet_collection()
        .iter()
        .position(|sheet| sheet.name() == target_sheet)
        .with_context(|| format!("worksheet {target_sheet:?} not found"))?;
    let target = workbook
        .sheet_by_name(target_sheet)
        .with_context(|| format!("worksheet {target_sheet:?} not found"))?;
    for cell in target.cells() {
        let (row, _) = parse_cell(&cell.coordinate().to_string())?;
        if row >= first_row {
            anyhow::ensure!(
                row.checked_add(count)
                    .is_some_and(|new_row| new_row <= EXCEL_MAX_ROWS),
                "row insertion would move target cell {} beyond the XLSX grid",
                cell.coordinate()
            );
        }
    }

    let mut target_formulas = Vec::new();
    let mut external_formula_groups =
        Vec::<(String, Vec<(String, umya_spreadsheet::CellValue)>)>::new();
    for sheet in workbook.sheet_collection() {
        let sheet_name = sheet.name().to_owned();
        let formulas = sheet
            .cells()
            .into_iter()
            .filter(|cell| cell.is_formula())
            .map(|cell| (cell.coordinate().to_string(), cell.cell_value().clone()))
            .collect::<Vec<_>>();
        if sheet_name == target_sheet {
            for (address, value) in formulas {
                ensure_row_insert_formula_supported(&value, &sheet_name, target_sheet)?;
                target_formulas.push((address, value));
            }
        } else {
            let relevant = formulas
                .into_iter()
                .filter(|(_, value)| formula_references_sheet(value.formula(), target_sheet))
                .collect::<Vec<_>>();
            if !relevant.is_empty() {
                for (_, value) in &relevant {
                    ensure_row_insert_formula_supported(&value, &sheet_name, target_sheet)?;
                }
                external_formula_groups.push((sheet_name, relevant));
            }
        }
    }

    // Transform target-sheet formulas in an isolated worksheet so formula
    // cells are not adjusted twice when the real target sheet is shifted.
    let target_formula_updates =
        adjust_target_formulas_for_row_insertion(target_sheet, first_row, count, &target_formulas)?;

    // The umya workbook-level insertion currently shifts every worksheet.
    // Use the target worksheet API for local rows, and use isolated formula
    // sheets only to update references from other sheets to the target.
    let mut external_formula_updates = Vec::new();
    for (formula_sheet, formulas) in external_formula_groups {
        let updates = adjust_external_formulas_for_row_insertion(
            &formula_sheet,
            target_sheet,
            first_row,
            count,
            &formulas,
        )?;
        external_formula_updates.push((formula_sheet, updates));
    }

    let global_names = workbook
        .defined_names()
        .iter()
        .map(|name| {
            adjust_defined_name_for_row_insertion(name, None, target_sheet, first_row, count)
        })
        .collect::<Result<Vec<_>>>()?;
    let mut local_name_updates = Vec::new();
    for sheet in workbook.sheet_collection() {
        let owner = sheet.name().to_owned();
        if owner == target_sheet {
            continue;
        }
        let names = sheet
            .defined_names()
            .iter()
            .map(|name| {
                adjust_defined_name_for_row_insertion(
                    name,
                    Some(&owner),
                    target_sheet,
                    first_row,
                    count,
                )
            })
            .collect::<Result<Vec<_>>>()?;
        local_name_updates.push((owner, names));
    }

    *workbook.defined_names_mut() = global_names;
    for (owner, names) in local_name_updates {
        *workbook
            .sheet_by_name_mut(&owner)
            .with_context(|| format!("worksheet {owner:?} disappeared"))?
            .defined_names_mut() = names;
    }
    for (formula_sheet, formulas) in external_formula_updates {
        let worksheet = workbook
            .sheet_by_name_mut(&formula_sheet)
            .with_context(|| format!("worksheet {formula_sheet:?} disappeared"))?;
        for (address, value) in formulas {
            worksheet.cell_mut(address.as_str()).set_cell_value(value);
        }
    }

    // Preserve cached formula values while letting the worksheet-level API
    // move cells, dimensions, comments, merges and other row-bound objects.
    {
        let target = &mut workbook.sheet_collection_mut()[target_index];
        for (address, _) in &target_formulas {
            target.cell_value_mut(address.as_str()).remove_formula();
        }
        target.insert_new_row(first_row, count);
    }
    let target = &mut workbook.sheet_collection_mut()[target_index];
    for formula in target_formula_updates {
        target
            .cell_mut(formula.inserted_address.as_str())
            .set_cell_value(formula.value);
    }
    Ok(())
}

fn adjust_target_formulas_for_row_insertion(
    target_sheet: &str,
    first_row: u32,
    count: u32,
    formulas: &[(String, umya_spreadsheet::CellValue)],
) -> Result<Vec<InsertedFormula>> {
    if formulas.is_empty() {
        return Ok(Vec::new());
    }
    let mut helper = umya_spreadsheet::new_file();
    let worksheet = helper
        .sheet_by_name_mut("Sheet1")
        .context("create target formula adjustment worksheet")?;
    worksheet.set_name(target_sheet);
    let mut locks = Vec::with_capacity(formulas.len());
    for (address, original) in formulas {
        let (unlocked, row_locks) = unlock_absolute_formula_rows(original.formula());
        let mut value = original.clone();
        set_cell_value_formula(&mut value, unlocked)?;
        worksheet.cell_mut(address.as_str()).set_cell_value(value);
        locks.push((address.clone(), row_locks));
    }
    worksheet.insert_new_row(first_row, count);
    let mut result = Vec::with_capacity(formulas.len());
    for (original_address, row_locks) in locks {
        let (row, column) = parse_cell(&original_address)?;
        let inserted_address = cell_address(
            if row >= first_row {
                row.checked_add(count)
                    .context("formula row exceeds XLSX limit")?
            } else {
                row
            },
            column,
        );
        let mut value = worksheet
            .get_cell(inserted_address.as_str())
            .with_context(|| format!("adjusted formula {inserted_address} is missing"))?
            .cell_value()
            .clone();
        let locked = restore_absolute_formula_rows(value.formula(), &row_locks)?;
        set_cell_value_formula(&mut value, locked)?;
        result.push(InsertedFormula {
            inserted_address,
            value,
        });
    }
    Ok(result)
}

fn adjust_external_formulas_for_row_insertion(
    formula_sheet: &str,
    target_sheet: &str,
    first_row: u32,
    count: u32,
    formulas: &[(String, umya_spreadsheet::CellValue)],
) -> Result<Vec<(String, umya_spreadsheet::CellValue)>> {
    let parking = formula_insertion_parking_coordinates(first_row, count, formulas.len())?;
    let mut helper = umya_spreadsheet::new_file();
    let worksheet = helper
        .sheet_by_name_mut("Sheet1")
        .context("create external formula adjustment worksheet")?;
    worksheet.set_name(formula_sheet);
    let mut marker = Vec::with_capacity(formulas.len());
    for ((original_address, original), parked_address) in formulas.iter().zip(&parking) {
        let (unlocked, row_locks) = unlock_absolute_formula_rows(original.formula());
        let mut value = original.clone();
        set_cell_value_formula(&mut value, unlocked)?;
        worksheet
            .cell_mut(parked_address.as_str())
            .set_cell_value(value);
        marker.push((original_address.clone(), parked_address.clone(), row_locks));
    }
    worksheet.insert_new_row_from_other_sheet(target_sheet, first_row, count);
    let mut result = Vec::with_capacity(marker.len());
    for (original_address, parked_address, row_locks) in marker {
        let (parked_row, parked_column) = parse_cell(&parked_address)?;
        let shifted_address = cell_address(
            if parked_row >= first_row {
                parked_row
                    .checked_add(count)
                    .context("parked formula exceeds XLSX limit")?
            } else {
                parked_row
            },
            parked_column,
        );
        let mut value = worksheet
            .get_cell(shifted_address.as_str())
            .with_context(|| format!("external formula {shifted_address} disappeared"))?
            .cell_value()
            .clone();
        let locked = restore_absolute_formula_rows(value.formula(), &row_locks)?;
        set_cell_value_formula(&mut value, locked)?;
        result.push((original_address, value));
    }
    Ok(result)
}

fn ensure_row_insert_formula_supported(
    value: &umya_spreadsheet::CellValue,
    formula_sheet: &str,
    target_sheet: &str,
) -> Result<()> {
    anyhow::ensure!(
        value.is_formula(),
        "row insertion formula adjustment received a non-formula cell"
    );
    let formula = value.formula();
    anyhow::ensure!(
        !formula.is_empty(),
        "shared formula followers are unsupported for row insertion"
    );
    if let Some(formula_object) = value.formula_obj() {
        anyhow::ensure!(
            formula_object.formula_type() == &umya_spreadsheet::CellFormulaValues::Normal,
            "array, data-table and shared formulas are unsupported for row insertion"
        );
    }
    let upper = formula.to_ascii_uppercase();
    anyhow::ensure!(
        !upper.contains("INDIRECT(") && !formula.contains(['[', ']']),
        "dynamic INDIRECT, structured and external formula references are unsupported for row insertion"
    );
    anyhow::ensure!(
        !formula_has_3d_sheet_reference(formula, target_sheet),
        "3-D formula references are unsupported for row insertion"
    );
    anyhow::ensure!(
        !formula_has_whole_row_reference(formula),
        "whole-row references are unsupported for row insertion"
    );
    anyhow::ensure!(
        !formula_sheet.is_empty() && !target_sheet.is_empty(),
        "formula and target worksheet names must not be empty"
    );
    Ok(())
}

fn formula_insertion_parking_coordinates(
    first_row: u32,
    count: u32,
    needed: usize,
) -> Result<Vec<String>> {
    if needed == 0 {
        return Ok(Vec::new());
    }
    let last_inserted_row = first_row
        .checked_add(
            count
                .checked_sub(1)
                .context("row insertion count must be positive")?,
        )
        .context("row insertion range overflows")?;
    let rows_before = u64::from(first_row - 1);
    let rows_after =
        u64::from(EXCEL_MAX_ROWS.saturating_sub(last_inserted_row.saturating_add(count)));
    let columns = u64::from(EXCEL_MAX_COLUMNS);
    let capacity = (rows_before + rows_after) * columns;
    anyhow::ensure!(
        capacity >= needed as u64,
        "no safe temporary worksheet cells are available for cross-sheet formula adjustment"
    );

    let mut result = Vec::with_capacity(needed);
    let before_capacity = rows_before * columns;
    for index in 0..needed as u64 {
        let (row, column) = if index < before_capacity {
            (index / columns + 1, index % columns + 1)
        } else {
            let after_index = index - before_capacity;
            (
                u64::from(last_inserted_row) + 1 + after_index / columns,
                after_index % columns + 1,
            )
        };
        result.push(cell_address(row as u32, column as u32));
    }
    Ok(result)
}

fn adjust_defined_name_for_row_insertion(
    original: &umya_spreadsheet::DefinedName,
    scope_sheet: Option<&str>,
    target_sheet: &str,
    first_row: u32,
    count: u32,
) -> Result<umya_spreadsheet::DefinedName> {
    let address = original.address();
    let trimmed = address.trim_start_matches('=');
    if umya_spreadsheet::helper::address::is_address(trimmed) {
        let (referenced_sheet, _) = umya_spreadsheet::helper::address::split_address(trimmed);
        let local_to_target = referenced_sheet.is_empty()
            && scope_sheet.is_some_and(|scope| scope.eq_ignore_ascii_case(target_sheet));
        if !referenced_sheet.eq_ignore_ascii_case(target_sheet) && !local_to_target {
            anyhow::ensure!(
                !referenced_sheet.is_empty()
                    || scope_sheet.is_some()
                    || !formula_has_a1_reference(trimmed),
                "workbook-level relative defined name {:?} is ambiguous during row insertion",
                original.name()
            );
            return Ok(original.clone());
        }

        let (unlocked, row_locks) = unlock_absolute_formula_rows(trimmed);
        let helper_address = if local_to_target {
            format!("'{}'!{unlocked}", target_sheet.replace('\'', "''"))
        } else {
            unlocked
        };
        let name = defined_name_with_address(original, &helper_address);
        let mut helper = umya_spreadsheet::new_file();
        let worksheet = helper
            .sheet_by_name_mut("Sheet1")
            .context("create defined-name adjustment worksheet")?;
        worksheet.set_name(target_sheet);
        worksheet.defined_names_mut().push(name);
        worksheet.insert_new_row(first_row, count);
        let adjusted = worksheet
            .defined_names()
            .first()
            .context("defined name disappeared during row insertion")?;
        let adjusted_address = adjusted.address();
        let final_address = if local_to_target {
            let (_, adjusted_range) =
                umya_spreadsheet::helper::address::split_address(&adjusted_address);
            restore_absolute_formula_rows(adjusted_range, &row_locks)?
        } else {
            restore_absolute_formula_rows(&adjusted_address, &row_locks)?
        };
        return Ok(defined_name_with_address(original, &final_address));
    }

    let target_reference = formula_references_sheet(&address, target_sheet);
    let local_relative_reference = scope_sheet
        .is_some_and(|scope| scope.eq_ignore_ascii_case(target_sheet))
        && formula_has_a1_reference(&address);
    anyhow::ensure!(
        !target_reference && !local_relative_reference,
        "defined name {:?} uses an expression that cannot be safely adjusted for row insertion",
        original.name()
    );
    Ok(original.clone())
}

fn formula_updates_for_sheet(
    updates: &[(String, Vec<(String, umya_spreadsheet::CellValue)>)],
    sheet_name: &str,
) -> Vec<(String, umya_spreadsheet::CellValue)> {
    updates
        .iter()
        .find(|(sheet, _)| sheet == sheet_name)
        .map(|(_, formulas)| formulas.clone())
        .unwrap_or_default()
}

fn adjust_formula_batch_for_row_deletion(
    formula_sheet: &str,
    target_sheet: &str,
    first_row: u32,
    count: u32,
    formulas: &[(String, umya_spreadsheet::CellValue)],
) -> Result<Vec<(String, umya_spreadsheet::CellValue)>> {
    let parking = formula_parking_coordinates(first_row, count, formulas.len())?;
    let mut helper = umya_spreadsheet::new_file();
    let worksheet = helper
        .sheet_mut(0)
        .context("create formula adjustment worksheet")?;
    worksheet.set_name(formula_sheet);

    let mut marker = Vec::with_capacity(formulas.len());
    for ((original_address, original_value), parked_address) in formulas.iter().zip(&parking) {
        ensure_row_delete_formula_supported(original_value, formula_sheet, target_sheet)?;
        let (unlocked_formula, row_locks) = unlock_absolute_formula_rows(original_value.formula());
        let mut value = original_value.clone();
        set_cell_value_formula(&mut value, unlocked_formula)?;
        worksheet
            .cell_mut(parked_address.as_str())
            .set_cell_value(value);
        marker.push((original_address.clone(), parked_address.clone(), row_locks));
    }

    worksheet.remove_row_from_other_sheet(target_sheet, first_row, count);
    let last_row = first_row + count - 1;
    let mut result = Vec::with_capacity(marker.len());
    for (original_address, parked_address, row_locks) in marker {
        let (parked_row, parked_column) = parse_cell(&parked_address)?;
        let adjusted_row = if parked_row > last_row {
            parked_row - count
        } else {
            parked_row
        };
        let adjusted_address = cell_address(adjusted_row, parked_column);
        let mut value = worksheet
            .get_cell(adjusted_address.as_str())
            .with_context(|| format!("formula adjustment cell {adjusted_address} disappeared"))?
            .cell_value()
            .clone();
        let locked_formula = restore_absolute_formula_rows(value.formula(), &row_locks)?;
        set_cell_value_formula(&mut value, locked_formula)?;
        result.push((original_address, value));
    }
    Ok(result)
}

fn formula_parking_coordinates(first_row: u32, count: u32, needed: usize) -> Result<Vec<String>> {
    if needed == 0 {
        return Ok(Vec::new());
    }
    let last_row = first_row
        .checked_add(
            count
                .checked_sub(1)
                .context("row deletion count must be positive")?,
        )
        .context("row deletion range overflows")?;
    let rows_before = u64::from(first_row - 1);
    let rows_after = u64::from(EXCEL_MAX_ROWS - last_row);
    let capacity = (rows_before + rows_after) * u64::from(EXCEL_MAX_COLUMNS);
    anyhow::ensure!(
        capacity >= needed as u64,
        "no safe temporary worksheet cells are available for formula reference adjustment"
    );

    let mut result = Vec::with_capacity(needed);
    let columns = u64::from(EXCEL_MAX_COLUMNS);
    let before_capacity = rows_before * columns;
    for index in 0..needed as u64 {
        let (row, column) = if index < before_capacity {
            (index / columns + 1, index % columns + 1)
        } else {
            let after_index = index - before_capacity;
            (
                u64::from(last_row) + after_index / columns + 1,
                after_index % columns + 1,
            )
        };
        result.push(cell_address(row as u32, column as u32));
    }
    Ok(result)
}

fn ensure_row_delete_formula_supported(
    value: &umya_spreadsheet::CellValue,
    formula_sheet: &str,
    target_sheet: &str,
) -> Result<()> {
    anyhow::ensure!(
        value.is_formula(),
        "formula reference adjustment received a non-formula cell"
    );
    let formula = value.formula();
    anyhow::ensure!(
        !formula.is_empty(),
        "shared formula followers are not supported for row deletion"
    );
    if let Some(formula_object) = value.formula_obj() {
        anyhow::ensure!(
            formula_object.formula_type() == &umya_spreadsheet::CellFormulaValues::Normal,
            "array, data-table, and shared formulas are not yet supported for row deletion"
        );
    }
    let upper = formula.to_ascii_uppercase();
    anyhow::ensure!(
        !upper.contains("INDIRECT(") && !formula.contains(['[', ']']),
        "dynamic INDIRECT or structured/external formula references are not yet supported for row deletion"
    );
    anyhow::ensure!(
        !formula_has_3d_sheet_reference(formula, target_sheet),
        "3-D formula references are not yet supported for row deletion"
    );
    anyhow::ensure!(
        !formula_has_whole_row_reference(formula),
        "whole-row formula references are not yet supported for row deletion"
    );
    anyhow::ensure!(
        !formula_sheet.is_empty() && !target_sheet.is_empty(),
        "formula and target worksheet names must not be empty"
    );
    Ok(())
}

fn adjust_defined_name_for_row_deletion(
    original: &umya_spreadsheet::DefinedName,
    scope_sheet: Option<&str>,
    target_sheet: &str,
    first_row: u32,
    count: u32,
) -> Result<Option<umya_spreadsheet::DefinedName>> {
    let address = original.address();
    let trimmed = address.trim_start_matches('=');
    if umya_spreadsheet::helper::address::is_address(trimmed) {
        let (referenced_sheet, _range) = umya_spreadsheet::helper::address::split_address(trimmed);
        let local_to_target = referenced_sheet.is_empty()
            && scope_sheet.is_some_and(|scope| scope.eq_ignore_ascii_case(target_sheet));
        if !referenced_sheet.eq_ignore_ascii_case(target_sheet) && !local_to_target {
            anyhow::ensure!(
                !referenced_sheet.is_empty()
                    || scope_sheet.is_some()
                    || !formula_has_a1_reference(trimmed),
                "workbook-level relative defined name {:?} is ambiguous during row deletion",
                original.name()
            );
            return Ok(Some(original.clone()));
        }

        let (unlocked_address, row_locks) = unlock_absolute_formula_rows(trimmed);
        let helper_address = if local_to_target {
            format!("'{}'!{unlocked_address}", target_sheet.replace('\'', "''"))
        } else {
            unlocked_address
        };
        let name = defined_name_with_address(original, &helper_address);
        let mut helper = umya_spreadsheet::new_file();
        let worksheet = helper
            .sheet_mut(0)
            .context("create defined-name adjustment worksheet")?;
        worksheet.set_name(target_sheet);
        worksheet.defined_names_mut().push(name);
        worksheet.remove_row(first_row, count);
        let Some(adjusted) = worksheet.defined_names().first() else {
            return Ok(None);
        };
        let adjusted_address = adjusted.address();
        let final_address = if local_to_target {
            let (_, adjusted_range) =
                umya_spreadsheet::helper::address::split_address(&adjusted_address);
            restore_absolute_formula_rows(adjusted_range, &row_locks).with_context(|| {
                format!(
                    "adjust local defined name {:?} from {:?} to {:?}",
                    original.name(),
                    trimmed,
                    adjusted_address
                )
            })?
        } else {
            restore_absolute_formula_rows(&adjusted_address, &row_locks).with_context(|| {
                format!(
                    "adjust defined name {:?} from {:?} to {:?}",
                    original.name(),
                    trimmed,
                    adjusted_address
                )
            })?
        };
        return Ok(Some(defined_name_with_address(original, &final_address)));
    }

    let target_reference = formula_references_sheet(&address, target_sheet);
    let local_relative_reference = scope_sheet
        .is_some_and(|scope| scope.eq_ignore_ascii_case(target_sheet))
        && formula_has_a1_reference(&address);
    anyhow::ensure!(
        !target_reference && !local_relative_reference,
        "defined name {:?} uses an expression that cannot be safely adjusted for row deletion",
        original.name()
    );
    Ok(Some(original.clone()))
}

fn defined_name_with_address(
    original: &umya_spreadsheet::DefinedName,
    address: &str,
) -> umya_spreadsheet::DefinedName {
    let mut name = umya_spreadsheet::DefinedName::default();
    name.set_name(original.name());
    if original.has_local_sheet_id() {
        name.set_local_sheet_id(original.local_sheet_id());
    }
    name.set_hidden(original.hidden());
    name.set_address(address);
    name
}

fn set_cell_value_formula(value: &mut umya_spreadsheet::CellValue, formula: String) -> Result<()> {
    let mut formula_object = value
        .formula_obj()
        .cloned()
        .context("formula cell has no formula object")?;
    formula_object.set_text(formula);
    value.set_formula_obj(formula_object);
    Ok(())
}

fn formula_references_sheet(formula: &str, sheet_name: &str) -> bool {
    let encoded = sheet_name.replace('\'', "''").to_ascii_lowercase();
    let formula = formula.to_ascii_lowercase();
    formula.contains(&format!("{encoded}!"))
        || formula.contains(&format!("'{encoded}'!"))
        || formula.contains(&format!("{encoded}'!"))
}

fn formula_has_a1_reference(formula: &str) -> bool {
    scan_formula_references(formula)
        .map(|references| !references.is_empty())
        .unwrap_or(true)
}

fn formula_has_whole_row_reference(formula: &str) -> bool {
    let bytes = formula.as_bytes();
    let mut index = 0;
    let mut in_string = false;
    while index < bytes.len() {
        if bytes[index] == b'"' {
            if in_string && bytes.get(index + 1) == Some(&b'"') {
                index += 2;
                continue;
            }
            in_string = !in_string;
            index += 1;
            continue;
        }
        if in_string {
            index += 1;
            continue;
        }
        if !bytes[index].is_ascii_digit() || !is_formula_boundary(bytes, index, true) {
            index += 1;
            continue;
        }
        let start = index;
        while index < bytes.len() && bytes[index].is_ascii_digit() {
            index += 1;
        }
        if bytes.get(index) != Some(&b':') {
            continue;
        }
        index += 1;
        let second_start = index;
        while index < bytes.len() && bytes[index].is_ascii_digit() {
            index += 1;
        }
        if second_start < index && is_formula_boundary(bytes, index, false) {
            let _ = start;
            return true;
        }
    }
    false
}

fn formula_has_3d_sheet_reference(formula: &str, target_sheet: &str) -> bool {
    let lower = formula.to_ascii_lowercase();
    let target = target_sheet.to_ascii_lowercase();
    let mut search_from = 0;
    while let Some(relative) = lower[search_from..].find('!') {
        let bang = search_from + relative;
        let prefix = &lower[..bang];
        let boundary = prefix
            .rfind(|character: char| "=,+-*/^&<>(){},; \t\r\n".contains(character))
            .map_or(0, |index| index + 1);
        let qualifier = &prefix[boundary..];
        if qualifier.contains(&target) && qualifier.contains(':') {
            return true;
        }
        search_from = bang + 1;
    }
    false
}

#[derive(Clone, Copy)]
struct FormulaReferenceSpan {
    start: usize,
    end: usize,
    row_lock: Option<usize>,
    row_digits: usize,
}

fn scan_formula_references(formula: &str) -> Result<Vec<FormulaReferenceSpan>> {
    let bytes = formula.as_bytes();
    let mut references = Vec::new();
    let mut index = 0;
    let mut in_string = false;
    let mut in_sheet_name = false;
    let mut bracket_depth = 0u32;
    while index < bytes.len() {
        let byte = bytes[index];
        if byte == b'"' && !in_sheet_name {
            if in_string && bytes.get(index + 1) == Some(&b'"') {
                index += 2;
                continue;
            }
            in_string = !in_string;
            index += 1;
            continue;
        }
        if !in_string && byte == b'\'' {
            if in_sheet_name && bytes.get(index + 1) == Some(&b'\'') {
                index += 2;
                continue;
            }
            in_sheet_name = !in_sheet_name;
            index += 1;
            continue;
        }
        if in_string || in_sheet_name {
            index += 1;
            continue;
        }
        match byte {
            b'[' => {
                bracket_depth += 1;
                index += 1;
                continue;
            }
            b']' if bracket_depth > 0 => {
                bracket_depth -= 1;
                index += 1;
                continue;
            }
            _ => {}
        }
        if bracket_depth > 0 {
            index += 1;
            continue;
        }
        if let Some(span) = parse_formula_a1_reference(bytes, index) {
            references.push(span);
            index = span.end;
        } else {
            index += 1;
        }
    }
    Ok(references)
}

fn parse_formula_a1_reference(bytes: &[u8], start: usize) -> Option<FormulaReferenceSpan> {
    if !is_formula_boundary(bytes, start, true) {
        return None;
    }
    let mut index = start;
    if bytes.get(index) == Some(&b'$') {
        index += 1;
    }
    let column_start = index;
    while bytes.get(index).is_some_and(u8::is_ascii_alphabetic) {
        index += 1;
    }
    let column_length = index - column_start;
    if !(1..=3).contains(&column_length) {
        return None;
    }
    let column = bytes[column_start..index]
        .iter()
        .try_fold(0u32, |value, byte| {
            value
                .checked_mul(26)?
                .checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))
        })?;
    if column > EXCEL_MAX_COLUMNS {
        return None;
    }
    let row_lock = if bytes.get(index) == Some(&b'$') {
        let lock = index;
        index += 1;
        Some(lock)
    } else {
        None
    };
    let row_digits = index;
    while bytes.get(index).is_some_and(u8::is_ascii_digit) {
        index += 1;
    }
    if row_digits == index || !is_formula_boundary(bytes, index, false) {
        return None;
    }
    if bytes.get(index) == Some(&b'!') || bytes.get(index) == Some(&b'(') {
        return None;
    }
    Some(FormulaReferenceSpan {
        start,
        end: index,
        row_lock,
        row_digits,
    })
}

fn is_formula_boundary(bytes: &[u8], index: usize, before: bool) -> bool {
    let neighbor = if before {
        index.checked_sub(1).and_then(|i| bytes.get(i)).copied()
    } else {
        bytes.get(index).copied()
    };
    neighbor.is_none_or(|byte| !(byte.is_ascii_alphanumeric() || byte == b'_' || byte == b'.'))
}

fn unlock_absolute_formula_rows(formula: &str) -> (String, Vec<bool>) {
    let spans = scan_formula_references(formula).unwrap_or_default();
    let mut output = String::with_capacity(formula.len());
    let mut row_locks = Vec::with_capacity(spans.len());
    let mut copied_to = 0;
    for span in spans {
        row_locks.push(span.row_lock.is_some());
        output.push_str(&formula[copied_to..span.start]);
        if let Some(row_lock) = span.row_lock {
            output.push_str(&formula[span.start..row_lock]);
            output.push_str(&formula[row_lock + 1..span.end]);
        } else {
            output.push_str(&formula[span.start..span.end]);
        }
        copied_to = span.end;
    }
    output.push_str(&formula[copied_to..]);
    (output, row_locks)
}

fn restore_absolute_formula_rows(formula: &str, row_locks: &[bool]) -> Result<String> {
    let spans = scan_formula_references(formula)?;
    anyhow::ensure!(
        spans.len() == row_locks.len(),
        "formula reference structure changed while adjusting absolute row references: expected {}, found {} in {formula:?}",
        row_locks.len(),
        spans.len()
    );
    let mut output =
        String::with_capacity(formula.len() + row_locks.iter().filter(|locked| **locked).count());
    let mut copied_to = 0;
    for (span, locked) in spans.into_iter().zip(row_locks) {
        output.push_str(&formula[copied_to..span.row_digits]);
        if *locked && span.row_lock.is_none() {
            output.push('$');
        }
        output.push_str(&formula[span.row_digits..span.end]);
        copied_to = span.end;
    }
    output.push_str(&formula[copied_to..]);
    Ok(output)
}

/// Structural validator for the umya candidate. Operation-specific output
/// assertions remain the responsibility of the caller's parity probe.
pub struct UmyaXlsxValidator;

impl OutputValidator for UmyaXlsxValidator {
    fn validate(&self, plan: &ExcelEditPlan, candidate_output: &Path) -> Result<()> {
        let dump = read_workbook(candidate_output)
            .with_context(|| format!("reopen candidate {}", candidate_output.display()))?;
        anyhow::ensure!(
            !dump.sheets.is_empty(),
            "candidate has no readable worksheets"
        );
        let profile = crate::inspect_xlsx_profile(candidate_output)?;
        anyhow::ensure!(
            profile.inventory_complete,
            "candidate contains unknown OOXML package parts"
        );
        for operation in &plan.operations {
            if let EditOperation::SetWorkbookCalculationProperties {
                full_calc_on_load,
                force_full_calc,
            } = operation
            {
                let actual = ooxml_patch::read_workbook_calculation_properties(candidate_output)?;
                anyhow::ensure!(
                    full_calc_on_load
                        .is_none_or(|expected| actual.full_calc_on_load == Some(expected)),
                    "fullCalcOnLoad readback does not match edit plan"
                );
                anyhow::ensure!(
                    force_full_calc.is_none_or(|expected| actual.force_full_calc == Some(expected)),
                    "forceFullCalc readback does not match edit plan"
                );
            }
        }
        let moves = plan
            .operations
            .iter()
            .filter_map(|operation| match operation {
                EditOperation::MoveSheet { sheet, new_index } => Some((sheet.as_str(), *new_index)),
                _ => None,
            })
            .collect::<Vec<_>>();
        if !moves.is_empty() {
            let mut expected_order = excel_reader::read_sheet_names(&plan.input)
                .with_context(|| format!("reopen source {}", plan.input.display()))?;
            for (sheet, new_index) in &moves {
                let current = expected_order
                    .iter()
                    .position(|name| name == sheet)
                    .with_context(|| format!("worksheet {sheet:?} not found"))?;
                anyhow::ensure!(
                    *new_index < expected_order.len(),
                    "move target index {new_index} exceeds workbook sheet count {}",
                    expected_order.len()
                );
                let moved = expected_order.remove(current);
                expected_order.insert(*new_index, moved);
            }
            let actual_order = dump
                .sheets
                .iter()
                .map(|sheet| sheet.name.clone())
                .collect::<Vec<_>>();
            anyhow::ensure!(
                actual_order == expected_order,
                "candidate sheet order {actual_order:?} does not match the move-sheet plan expectation {expected_order:?}"
            );
            assert_scoped_defined_names_follow_sheets(&plan.input, candidate_output)?;
        }
        let freeze_edits = plan
            .operations
            .iter()
            .filter_map(|operation| match operation {
                EditOperation::SetFreezePanes { sheet, cell } => {
                    Some((sheet.as_str(), cell.as_deref()))
                }
                _ => None,
            })
            .collect::<Vec<_>>();
        let filter_edits = plan
            .operations
            .iter()
            .filter_map(|operation| match operation {
                EditOperation::SetAutoFilter { sheet, range } => {
                    Some((sheet.as_str(), range.as_deref()))
                }
                _ => None,
            })
            .collect::<Vec<_>>();
        if !freeze_edits.is_empty() || !filter_edits.is_empty() {
            let reopened = crate::open_xlsx(candidate_output)
                .with_context(|| format!("reopen candidate {}", candidate_output.display()))?;
            for (sheet, cell) in &freeze_edits {
                assert_freeze_pane_matches_plan(&reopened, &final_sheet_name(plan, sheet), *cell)?;
            }
            for (sheet, range) in &filter_edits {
                assert_auto_filter_matches_plan(&reopened, &final_sheet_name(plan, sheet), *range)?;
            }
            assert_filter_database_matches_plan(plan, candidate_output, &filter_edits)?;
        }
        Ok(())
    }
}

/// Sheet-scoped defined names must stay attached to the worksheet named in
/// their scope after a sheet move: the write path renumbers `localSheetId`,
/// so a stale id would silently re-bind print areas and print titles to a
/// different worksheet. Compare the resolved (name, scoped worksheet,
/// reference) records of source and candidate; records whose scope the source
/// file itself left unresolved are skipped, because the umya read path already
/// drops them before any write.
fn assert_scoped_defined_names_follow_sheets(source: &Path, candidate: &Path) -> Result<()> {
    let scoped_keys = |path: &Path| -> Result<Vec<(String, String, String)>> {
        let mut keys = Vec::new();
        for record in excel_reader::read_defined_names(path)? {
            let scoped_sheet = match &record.scope {
                excel_reader::DefinedNameScope::Workbook => String::new(),
                excel_reader::DefinedNameScope::Worksheet {
                    sheet_name: Some(name),
                    ..
                } => name.clone(),
                _ => continue,
            };
            keys.push((
                record.name,
                scoped_sheet,
                normalize_defined_name_reference(&record.raw_reference),
            ));
        }
        keys.sort();
        Ok(keys)
    };
    let expected = scoped_keys(source)
        .with_context(|| format!("read source defined names {}", source.display()))?;
    let actual = scoped_keys(candidate)
        .with_context(|| format!("read candidate defined names {}", candidate.display()))?;
    anyhow::ensure!(
        actual == expected,
        "candidate defined names do not follow their worksheets after the move: expected {expected:?}, found {actual:?}"
    );
    Ok(())
}

/// openpyxl freeze semantics: `None` and `A1` clear the freeze; otherwise the
/// cell is the first scrollable cell and the active pane follows the splits.
/// Returns the expected splits, active pane and top-left cell coordinates.
fn expected_freeze_pane(
    cell: Option<&str>,
) -> Result<Option<(f64, f64, umya_spreadsheet::PaneValues, u32, u32)>> {
    let Some(cell) = cell else {
        return Ok(None);
    };
    if cell.trim().eq_ignore_ascii_case("A1") {
        return Ok(None);
    }
    let (row, column) = parse_cell(cell)?;
    let x_split = if column > 1 { (column - 1) as f64 } else { 0.0 };
    let y_split = if row > 1 { (row - 1) as f64 } else { 0.0 };
    let active_pane = match (column > 1, row > 1) {
        (true, true) => umya_spreadsheet::PaneValues::BottomRight,
        (true, false) => umya_spreadsheet::PaneValues::TopRight,
        (false, true) => umya_spreadsheet::PaneValues::BottomLeft,
        (false, false) => bail!("freeze panes at A1 clears the freeze"),
    };
    Ok(Some((x_split, y_split, active_pane, row, column)))
}

fn assert_freeze_pane_matches_plan(
    workbook: &crate::Workbook,
    sheet: &str,
    cell: Option<&str>,
) -> Result<()> {
    let worksheet = workbook
        .sheet_by_name(sheet)
        .with_context(|| format!("output sheet {sheet:?} missing"))?;
    let pane = worksheet
        .sheets_views()
        .sheet_view_list()
        .iter()
        .find_map(|view| view.pane());
    match expected_freeze_pane(cell)? {
        None => anyhow::ensure!(
            pane.is_none(),
            "sheet {sheet:?} still carries a frozen pane after a freeze-panes clear"
        ),
        Some((x_split, y_split, active_pane, row, column)) => {
            let pane = pane.with_context(|| {
                format!("sheet {sheet:?} has no frozen pane after SetFreezePanes")
            })?;
            anyhow::ensure!(
                format!("{:?}", pane.state()) == "Frozen",
                "sheet {sheet:?} pane state is not frozen"
            );
            anyhow::ensure!(
                pane.horizontal_split() == x_split && pane.vertical_split() == y_split,
                "sheet {sheet:?} pane splits are {:?}/{:?} instead of {x_split}/{y_split}",
                pane.horizontal_split(),
                pane.vertical_split()
            );
            anyhow::ensure!(
                format!("{:?}", pane.active_pane()) == format!("{:?}", active_pane),
                "sheet {sheet:?} pane active pane does not match the plan"
            );
            anyhow::ensure!(
                pane.top_left_cell().row_num() == row && pane.top_left_cell().col_num() == column,
                "sheet {sheet:?} pane top-left cell does not match the plan"
            );
        }
    }
    Ok(())
}

fn assert_auto_filter_matches_plan(
    workbook: &crate::Workbook,
    sheet: &str,
    range: Option<&str>,
) -> Result<()> {
    let worksheet = workbook
        .sheet_by_name(sheet)
        .with_context(|| format!("output sheet {sheet:?} missing"))?;
    match range {
        Some(range) => {
            let actual = worksheet.auto_filter().with_context(|| {
                format!("sheet {sheet:?} has no AutoFilter after SetAutoFilter")
            })?;
            anyhow::ensure!(
                actual.range().range() == range,
                "sheet {sheet:?} AutoFilter range is {:?} instead of {range:?}",
                actual.range().range()
            );
        }
        None => anyhow::ensure!(
            worksheet.auto_filter().is_none(),
            "sheet {sheet:?} still carries an AutoFilter after a clear"
        ),
    }
    Ok(())
}

/// The hidden `_xlnm._FilterDatabase` name must follow the requested
/// AutoFilter range the way Excel and the openpyxl worker maintain it; the
/// raw reference comparison normalizes quotes because umya re-renders
/// non-ASCII sheet names quoted.
fn assert_filter_database_matches_plan(
    plan: &ExcelEditPlan,
    candidate_output: &Path,
    filter_edits: &[(&str, Option<&str>)],
) -> Result<()> {
    if filter_edits.is_empty() {
        return Ok(());
    }
    let records = excel_reader::read_defined_names(candidate_output)?;
    for (sheet, range) in filter_edits {
        let final_sheet = final_sheet_name(plan, sheet);
        let record = records.iter().find(|record| {
            record.name == "_xlnm._FilterDatabase"
                && matches!(
                    &record.scope,
                    excel_reader::DefinedNameScope::Worksheet {
                        sheet_name: Some(name),
                        ..
                    } if name == &final_sheet
                )
        });
        match range {
            Some(range) => {
                let record = record.with_context(|| {
                    format!("output has no _FilterDatabase defined name for sheet {final_sheet:?}")
                })?;
                let expected = format!("{final_sheet}!{}", absolute_range(range)?);
                anyhow::ensure!(
                    normalize_defined_name_reference(&record.raw_reference)
                        == normalize_defined_name_reference(&expected),
                    "sheet {final_sheet:?} _FilterDatabase is {:?} instead of {expected:?}",
                    record.raw_reference
                );
            }
            None => anyhow::ensure!(
                record.is_none(),
                "sheet {final_sheet:?} still carries a _FilterDatabase after the AutoFilter clear"
            ),
        }
    }
    Ok(())
}

fn apply_operation(workbook: &mut crate::Workbook, operation: &EditOperation) -> Result<()> {
    match operation {
        EditOperation::SetCellValue { sheet, cell, value } => {
            let target = workbook
                .get_sheet_by_name_mut(sheet)
                .with_context(|| format!("worksheet {sheet:?} not found"))?;
            let target = target.cell_mut(cell.as_str());
            match value {
                crate::CellValue::Text(value) => {
                    target.set_value_string(value);
                }
                crate::CellValue::Decimal(value) => {
                    let number = value.parse::<f64>().with_context(|| {
                        format!(
                            "decimal value {value:?} is not representable as XLSX numeric value"
                        )
                    })?;
                    anyhow::ensure!(number.is_finite(), "decimal value must be finite");
                    target.set_value_number(number);
                }
                crate::CellValue::DateTime {
                    serial,
                    number_format,
                } => {
                    let number = serial.parse::<f64>().context("parse date/time serial")?;
                    anyhow::ensure!(number.is_finite(), "date/time serial must be finite");
                    target.set_value_number(number);
                    target
                        .get_style_mut()
                        .get_number_format_mut()
                        .set_format_code(number_format);
                }
                crate::CellValue::Boolean(value) => {
                    target.set_value_bool(*value);
                }
                crate::CellValue::Error(value) => {
                    target.set_error(value);
                }
                crate::CellValue::Blank => {
                    target.set_cell_value(umya_spreadsheet::CellValue::default());
                }
            }
        }
        EditOperation::SetFormulaText {
            sheet,
            cell,
            formula,
        } => {
            workbook
                .get_sheet_by_name_mut(sheet)
                .with_context(|| format!("worksheet {sheet:?} not found"))?
                .cell_mut(cell.as_str())
                .set_formula(formula.trim_start_matches('='));
        }
        EditOperation::SetComment {
            sheet,
            cell,
            text,
            author,
        } => {
            let target = workbook
                .get_sheet_by_name_mut(sheet)
                .with_context(|| format!("worksheet {sheet:?} not found"))?;
            target
                .comments_mut()
                .retain(|comment| comment.coordinate().to_string().eq_ignore_ascii_case(cell));
            if let Some(text) = text {
                let mut comment = umya_spreadsheet::Comment::default();
                comment.new_comment(cell.as_str());
                comment.set_author(author.as_deref().unwrap_or("pytools"));
                comment.set_text_string(text);
                target.add_comments(comment);
            }
        }
        EditOperation::SetColumnWidth {
            sheet,
            column_index,
            width,
        } => {
            anyhow::ensure!(
                *column_index > 0 && width.is_finite() && *width >= 0.0,
                "invalid column width request"
            );
            workbook
                .get_sheet_by_name_mut(sheet)
                .with_context(|| format!("worksheet {sheet:?} not found"))?
                .column_dimension_by_number_mut(*column_index)
                .set_width(*width);
        }
        EditOperation::SetRowHeight {
            sheet,
            row_index,
            height,
        } => {
            anyhow::ensure!(
                *row_index > 0 && height.is_finite() && *height >= 0.0,
                "invalid row height request"
            );
            workbook
                .get_sheet_by_name_mut(sheet)
                .with_context(|| format!("worksheet {sheet:?} not found"))?
                .row_dimension_mut(*row_index)
                .set_height(*height);
        }
        EditOperation::DeleteRows {
            sheet,
            first_row,
            count,
        } => delete_rows_with_excel_references(workbook, sheet, *first_row, *count)?,
        EditOperation::InsertRows {
            sheet,
            first_row,
            count,
        } => {
            anyhow::ensure!(
                *first_row > 0 && *count > 0,
                "row insertion requires positive one-based indexes"
            );
            insert_rows_with_excel_references(workbook, sheet, *first_row, *count)?;
        }
        EditOperation::RenameSheet { old_name, new_name } => {
            validate_sheet_name(new_name)?;
            anyhow::ensure!(
                old_name != new_name,
                "source and destination sheet names are identical"
            );
            anyhow::ensure!(
                workbook
                    .get_sheet_collection()
                    .iter()
                    .all(|sheet| sheet.get_name() != new_name),
                "worksheet {new_name:?} already exists"
            );
            let worksheet = workbook
                .get_sheet_by_name_mut(old_name)
                .with_context(|| format!("worksheet {old_name:?} not found"))?;
            let original_defined_names = worksheet.defined_names().to_vec();
            worksheet.set_name(new_name);

            // Match the migrated openpyxl behavior: a worksheet rename updates
            // its print area, but leaves user-defined references untouched.
            // umya stores sheet-qualified workbook names on the worksheet and
            // updates every such reference from `set_name`, so restore those
            // names while retaining the newly renamed print area.
            let worksheet = workbook
                .get_sheet_by_name_mut(new_name)
                .with_context(|| format!("renamed worksheet {new_name:?} not found"))?;
            for (updated, original) in worksheet
                .defined_names_mut()
                .iter_mut()
                .zip(original_defined_names)
            {
                if updated.name().eq_ignore_ascii_case("_xlnm.Print_Area") {
                    let updated_address = updated.address();
                    let (_, area) =
                        umya_spreadsheet::helper::address::split_address(&updated_address);
                    let quoted_sheet = new_name.replace('\'', "''");
                    *updated =
                        defined_name_with_address(updated, &format!("'{quoted_sheet}'!{area}"));
                } else {
                    *updated = original;
                }
            }
        }
        EditOperation::CreateSheet { name } => {
            validate_sheet_name(name)?;
            anyhow::ensure!(
                workbook
                    .get_sheet_collection()
                    .iter()
                    .all(|sheet| sheet.get_name() != name),
                "worksheet {name:?} already exists"
            );
            workbook
                .new_sheet(name)
                .with_context(|| format!("create worksheet {name:?}"))?;
        }
        EditOperation::CreateSheetAt { name, index } => {
            validate_sheet_name(name)?;
            let current_count = workbook.sheet_count();
            anyhow::ensure!(
                *index <= current_count,
                "worksheet insertion index {index} exceeds workbook sheet count {current_count}"
            );
            anyhow::ensure!(
                workbook
                    .get_sheet_collection()
                    .iter()
                    .all(|sheet| sheet.get_name() != name),
                "worksheet {name:?} already exists"
            );
            let active_sheet = active_sheet_name(workbook);
            workbook
                .new_sheet(name)
                .with_context(|| format!("create worksheet {name:?}"))?;
            let collection = workbook.get_sheet_collection_mut();
            let appended_index = collection.len().saturating_sub(1);
            for destination in (*index..appended_index).rev() {
                collection.swap(destination, destination + 1);
            }
            remap_sheet_index_metadata(workbook, active_sheet.as_deref());
        }
        EditOperation::MoveSheet { sheet, new_index } => {
            validate_sheet_name(sheet)?;
            let current_count = workbook.sheet_count();
            anyhow::ensure!(
                *new_index < current_count,
                "move target index {new_index} exceeds workbook sheet count {current_count}"
            );
            let current_index = workbook
                .get_sheet_collection()
                .iter()
                .position(|candidate| candidate.get_name() == sheet.as_str())
                .with_context(|| format!("worksheet {sheet:?} not found"))?;
            let active_sheet = active_sheet_name(workbook);
            let collection = workbook.get_sheet_collection_mut();
            if *new_index < current_index {
                for slot in (*new_index..current_index).rev() {
                    collection.swap(slot, slot + 1);
                }
            } else if *new_index > current_index {
                for slot in current_index..*new_index {
                    collection.swap(slot, slot + 1);
                }
            }
            remap_sheet_index_metadata(workbook, active_sheet.as_deref());
        }
        EditOperation::DeleteSheet { name } => {
            anyhow::ensure!(
                workbook.sheet_count() > 1,
                "cannot delete the only worksheet"
            );
            let active_sheet = active_sheet_name(workbook).filter(|active| active != name);
            workbook
                .remove_sheet_by_name(name)
                .with_context(|| format!("delete worksheet {name:?}"))?;
            remap_sheet_index_metadata(workbook, active_sheet.as_deref());
        }
        EditOperation::SetCellFillColor { sheet, cell, argb } => {
            let normalized = normalize_argb(argb)?;
            let target = workbook
                .get_sheet_by_name_mut(sheet)
                .with_context(|| format!("worksheet {sheet:?} not found"))?;
            let mut color = umya_spreadsheet::Color::default();
            color.set_argb_str(&normalized);
            let pattern_fill = target
                .get_style_mut(cell.as_str())
                .get_fill_mut()
                .get_pattern_fill_mut();
            pattern_fill
                .set_pattern_type(umya_spreadsheet::PatternValues::Solid)
                .set_foreground_color(color);
        }
        EditOperation::CopyRange {
            source_sheet,
            source_range,
            target_sheet,
            target_start,
            copy_values,
            copy_formulas,
            copy_styles,
            copy_comments,
        } => {
            copy_range(
                workbook,
                source_sheet,
                source_range,
                target_sheet,
                target_start,
                *copy_values,
                *copy_formulas,
                *copy_styles,
                *copy_comments,
            )?;
        }
        EditOperation::ImportRange {
            source_workbook,
            source_sheet,
            source_range,
            target_sheet,
            target_start,
            copy_values,
            copy_formulas,
            style_mode,
            copy_comments,
            copy_merges,
            copy_row_heights,
            copy_column_widths,
            copy_row_hidden,
            copy_column_hidden,
            copy_outline_levels: _,
            copy_all_row_dimensions,
            copy_all_column_dimensions,
            materialize_blank_cells,
            hide_zero,
        } => {
            let source = crate::open_xlsx(source_workbook).with_context(|| {
                format!("open range-import source {}", source_workbook.display())
            })?;
            import_range(
                &source,
                workbook,
                source_sheet,
                source_range,
                target_sheet,
                target_start,
                *copy_values,
                *copy_formulas,
                *style_mode,
                *copy_comments,
                *copy_merges,
                *copy_row_heights,
                *copy_column_widths,
                *copy_row_hidden,
                *copy_column_hidden,
                *copy_all_row_dimensions,
                *copy_all_column_dimensions,
                *materialize_blank_cells,
                *hide_zero,
            )?;
        }
        EditOperation::SetWorkbookCalculationProperties { .. } => {
            // Applied after umya saves, through the namespace-aware workbook.xml patch.
        }
        EditOperation::SetPageSettings { sheet, settings } => {
            let target = workbook
                .sheet_by_name_mut(sheet)
                .with_context(|| format!("worksheet {sheet:?} not found"))?;
            let page = target.get_page_setup_mut();
            if let Some(orientation) = settings.orientation {
                page.set_orientation(match orientation {
                    PageOrientation::Portrait => umya_spreadsheet::OrientationValues::Portrait,
                    PageOrientation::Landscape => umya_spreadsheet::OrientationValues::Landscape,
                });
            }
            if let Some(width) = settings.fit_to_width {
                page.set_fit_to_width(width as u32);
            }
            if let Some(height) = settings.fit_to_height {
                page.set_fit_to_height(height as u32);
            }
            if let Some(value) = settings.fit_to_page {
                page.set_fit_to_page(value);
            }
            // umya 3.1.1 has no API to remove the optional `scale` field.
            // Writing 0 serializes invalid `scale="0"`; new sheets already
            // have no scale, while existing sheets are rejected in preflight.
            if let Some(paper_size) = settings.paper_size {
                page.set_paper_size(paper_size as u32);
            }
            if let Some(value) = settings.horizontal_centered {
                target
                    .get_print_options_mut()
                    .set_horizontal_centered(value);
            }
            if let Some(value) = settings.vertical_centered {
                target.get_print_options_mut().set_vertical_centered(value);
            }
            if let Some(footer) = &settings.footer {
                let footer = if footer.starts_with("&C") {
                    footer.clone()
                } else {
                    format!("&C{footer}")
                };
                target
                    .get_header_footer_mut()
                    .get_odd_footer_mut()
                    .set_value(footer);
            }
            // The edit plan may delete or reorder worksheets after this
            // operation. Print-area localSheetId values are applied once the
            // final sheet collection is known.
        }
        EditOperation::SetFreezePanes { sheet, cell } => {
            let target = workbook
                .get_sheet_by_name_mut(sheet)
                .with_context(|| format!("worksheet {sheet:?} not found"))?;
            if is_freeze_clear_request(cell.as_deref()) {
                // Clearing is a no-op here: a sheet that still carries a pane
                // is rejected before any write by the preflight check, and a
                // sheet without a pane is already clear.
                return Ok(());
            }
            apply_freeze_panes(
                target,
                cell.as_deref().context("freeze panes cell is missing")?,
            )?;
        }
        EditOperation::SetAutoFilter { sheet, range } => {
            {
                let target = workbook
                    .get_sheet_by_name_mut(sheet)
                    .with_context(|| format!("worksheet {sheet:?} not found"))?;
                match range {
                    Some(range) => target.set_auto_filter(range.as_str()),
                    None => target.remove_auto_filter(),
                }
            }
            if range.is_some() {
                update_sheet_filter_database(
                    workbook,
                    sheet,
                    range.as_deref().context("AutoFilter range is missing")?,
                )?;
            } else {
                remove_sheet_filter_database(workbook, sheet)?;
            }
        }
        EditOperation::ReplaceHeaderText { .. } => {
            bail!("header replacement must use the namespace-aware OOXML header backend")
        }
    }
    Ok(())
}

fn apply_final_print_areas(
    workbook: &mut crate::Workbook,
    operations: &[EditOperation],
) -> Result<()> {
    for operation in operations {
        let EditOperation::SetPageSettings {
            sheet,
            settings:
                crate::PageSettings {
                    print_area: Some(area),
                    ..
                },
        } = operation
        else {
            continue;
        };

        let sheet_index = workbook
            .sheet_collection()
            .iter()
            .position(|candidate| candidate.name() == sheet)
            .with_context(|| format!("worksheet {sheet:?} not found for print area"))?
            as u32;

        workbook
            .sheet_by_name_mut(sheet)
            .with_context(|| format!("worksheet {sheet:?} not found for print area"))?
            .defined_names_mut()
            .retain(|name| !name.name().eq_ignore_ascii_case("_xlnm.Print_Area"));
        workbook.defined_names_mut().retain(|name| {
            !(name.name().eq_ignore_ascii_case("_xlnm.Print_Area")
                && name.has_local_sheet_id()
                && name.local_sheet_id() == sheet_index)
        });

        let quoted_sheet = sheet.replace('\'', "''");
        let mut defined_name = umya_spreadsheet::DefinedName::default();
        defined_name
            .set_name("_xlnm.Print_Area")
            .set_address(format!("'{quoted_sheet}'!{area}"));
        defined_name.set_local_sheet_id(sheet_index);
        workbook.add_defined_names(defined_name);
    }
    Ok(())
}

fn copy_range(
    workbook: &mut crate::Workbook,
    source_sheet: &str,
    source_range: &str,
    target_sheet: &str,
    target_start: &str,
    copy_values: bool,
    copy_formulas: bool,
    copy_styles: bool,
    copy_comments: bool,
) -> Result<()> {
    let (min_row, min_col, max_row, max_col) = parse_range(source_range)?;
    let (target_row, target_col) = parse_cell(target_start)?;
    let snapshots = {
        let source = workbook
            .get_sheet_by_name(source_sheet)
            .with_context(|| format!("worksheet {source_sheet:?} not found"))?;
        let mut cells = Vec::new();
        for row in min_row..=max_row {
            for col in min_col..=max_col {
                let address = cell_address(row, col);
                let Some(cell) = source.get_cell(address.as_str()) else {
                    continue;
                };
                let cell_value = if cell.cell_value().is_formula() && copy_formulas {
                    Some(ImportCellValue::Typed(cell.cell_value().clone()))
                } else if copy_values && cell.cell_value().is_formula() {
                    Some(ImportCellValue::FormulaResult(cell.value().to_string()))
                } else if copy_values {
                    Some(ImportCellValue::Typed(cell.cell_value().clone()))
                } else {
                    None
                };
                let style = cell.style().clone();
                let comment = if copy_comments {
                    source
                        .comments()
                        .iter()
                        .find(|comment| {
                            comment
                                .coordinate()
                                .to_string()
                                .eq_ignore_ascii_case(&address)
                        })
                        .cloned()
                } else {
                    None
                };
                cells.push((row, col, cell_value, style, comment));
            }
        }
        cells
    };
    let target = workbook
        .get_sheet_by_name_mut(target_sheet)
        .with_context(|| format!("worksheet {target_sheet:?} not found"))?;
    for (row, col, cell_value, style, comment) in snapshots {
        let dest_row = target_row + row - min_row;
        let dest_col = target_col + col - min_col;
        let dest = cell_address(dest_row, dest_col);
        let cell = target.cell_mut(dest.as_str());
        if copy_styles {
            cell.set_style(style);
        }
        if let Some(value) = cell_value {
            match value {
                ImportCellValue::Typed(value) => {
                    cell.set_cell_value(value);
                }
                ImportCellValue::FormulaResult(value) => {
                    if value.is_empty() {
                        cell.set_blank();
                    } else {
                        cell.set_value_lazy(value);
                    }
                }
            }
        }
        if let Some(mut comment) = comment {
            comment
                .coordinate_mut()
                .set_col_num(dest_col)
                .set_row_num(dest_row);
            target.add_comments(comment);
        }
    }
    Ok(())
}

enum ImportCellValue {
    Typed(umya_spreadsheet::CellValue),
    FormulaResult(String),
}

#[allow(clippy::too_many_arguments)]
fn import_range(
    source_workbook: &crate::Workbook,
    target_workbook: &mut crate::Workbook,
    source_sheet: &str,
    source_range: &str,
    target_sheet: &str,
    target_start: &str,
    copy_values: bool,
    copy_formulas: bool,
    style_mode: RangeStyleMode,
    copy_comments: bool,
    copy_merges: bool,
    copy_row_heights: bool,
    copy_column_widths: bool,
    copy_row_hidden: bool,
    copy_column_hidden: bool,
    copy_all_row_dimensions: bool,
    copy_all_column_dimensions: bool,
    materialize_blank_cells: bool,
    hide_zero: bool,
) -> Result<()> {
    let (min_row, min_col, max_row, max_col) = parse_range(source_range)?;
    let (target_row, target_col) = parse_cell(target_start)?;
    anyhow::ensure!(
        (!copy_all_row_dimensions && !copy_all_column_dimensions)
            || (min_row == 1 && min_col == 1 && target_row == 1 && target_col == 1),
        "copying all worksheet dimensions requires source and target ranges rooted at A1"
    );
    let preexisting_target_columns = target_workbook
        .get_sheet_by_name(target_sheet)
        .with_context(|| format!("target worksheet {target_sheet:?} not found"))?
        .column_dimensions()
        .iter()
        .map(|dimension| dimension.col_num())
        .collect::<Vec<_>>();
    let snapshots = {
        let source = source_workbook
            .get_sheet_by_name(source_sheet)
            .with_context(|| format!("source worksheet {source_sheet:?} not found"))?;
        let mut cells = Vec::new();
        for row in min_row..=max_row {
            for col in min_col..=max_col {
                let address = cell_address(row, col);
                let Some(cell) = source.get_cell(address.as_str()) else {
                    if materialize_blank_cells {
                        cells.push((row, col, None, None, None));
                    }
                    continue;
                };
                let value = if cell.cell_value().is_formula() && copy_formulas {
                    Some(ImportCellValue::Typed(cell.cell_value().clone()))
                } else if copy_values && cell.cell_value().is_formula() {
                    let mut cached_value = cell.cell_value().clone();
                    cached_value.remove_formula();
                    Some(ImportCellValue::Typed(cached_value))
                } else if copy_values {
                    Some(ImportCellValue::Typed(cell.cell_value().clone()))
                } else {
                    None
                };
                let style = match style_mode {
                    RangeStyleMode::None => None,
                    RangeStyleMode::Full => Some(cell.style().clone()),
                    RangeStyleMode::Basic => {
                        let mut style = cell.style().clone();
                        style.remove_borders();
                        style.remove_protection();
                        Some(style)
                    }
                };
                let comment = if copy_comments {
                    source
                        .comments()
                        .iter()
                        .find(|comment| {
                            comment
                                .coordinate()
                                .to_string()
                                .eq_ignore_ascii_case(&address)
                        })
                        .cloned()
                } else {
                    None
                };
                cells.push((row, col, value, style, comment));
            }
        }
        let merge_ranges = if copy_merges {
            source
                .merge_cells()
                .iter()
                .filter_map(|merge| {
                    let (r1, c1, r2, c2) = parse_range(&merge.range()).ok()?;
                    (r1 >= min_row && c1 >= min_col && r2 <= max_row && c2 <= max_col)
                        .then_some((r1, c1, r2, c2))
                })
                .collect::<Vec<_>>()
        } else {
            Vec::new()
        };
        let row_heights = if copy_row_heights || copy_row_hidden {
            if copy_all_row_dimensions {
                source
                    .row_dimensions()
                    .into_iter()
                    .map(|dimension| {
                        (
                            dimension.row_num(),
                            copy_row_heights
                                .then(|| {
                                    (dimension.custom_height() || dimension.height() != 0.0)
                                        .then(|| dimension.height())
                                })
                                .flatten(),
                            copy_row_hidden.then(|| dimension.get_hidden()),
                        )
                    })
                    .collect::<Vec<_>>()
            } else {
                (min_row..=max_row)
                    .filter_map(|row| {
                        source.row_dimension(row).map(|dimension| {
                            (
                                row,
                                copy_row_heights
                                    .then(|| {
                                        (dimension.custom_height() || dimension.height() != 0.0)
                                            .then(|| dimension.height())
                                    })
                                    .flatten(),
                                copy_row_hidden.then(|| dimension.get_hidden()),
                            )
                        })
                    })
                    .collect::<Vec<_>>()
            }
        } else {
            Vec::new()
        };
        let column_widths = if copy_column_widths || copy_column_hidden {
            if copy_all_column_dimensions {
                source
                    .column_dimensions()
                    .iter()
                    .map(|dimension| {
                        (
                            dimension.col_num(),
                            copy_column_widths.then(|| dimension.width()),
                            copy_column_hidden.then(|| dimension.get_hidden()),
                        )
                    })
                    .collect::<Vec<_>>()
            } else {
                (min_col..=max_col)
                    .filter_map(|column| {
                        source.column_dimension_by_number(column).map(|dimension| {
                            (
                                column,
                                copy_column_widths.then(|| dimension.width()),
                                copy_column_hidden.then(|| dimension.get_hidden()),
                            )
                        })
                    })
                    .collect::<Vec<_>>()
            }
        } else {
            Vec::new()
        };
        (cells, merge_ranges, row_heights, column_widths)
    };
    let target = target_workbook
        .get_sheet_by_name_mut(target_sheet)
        .with_context(|| format!("target worksheet {target_sheet:?} not found"))?;
    for (row, col, value, style, comment) in snapshots.0 {
        let dest_row = target_row + row - min_row;
        let dest_col = target_col + col - min_col;
        let dest = cell_address(dest_row, dest_col);
        let cell = target.cell_mut(dest.as_str());
        if let Some(style) = style {
            cell.set_style(style);
        }
        if let Some(value) = value {
            match value {
                ImportCellValue::Typed(value) if value.is_formula() && copy_formulas => {
                    cell.set_cell_value(value);
                }
                ImportCellValue::Typed(value) => {
                    let zero = value.value_number().is_some_and(|number| number == 0.0);
                    if hide_zero && zero {
                        cell.set_blank();
                    } else {
                        cell.set_cell_value(value);
                    }
                }
                ImportCellValue::FormulaResult(value) => {
                    let is_zero = value
                        .trim()
                        .parse::<f64>()
                        .is_ok_and(|number| number == 0.0);
                    if value.is_empty() || (hide_zero && is_zero) {
                        cell.set_blank();
                    } else {
                        cell.set_value_lazy(value);
                    }
                }
            }
        }
        if let Some(mut comment) = comment {
            comment
                .coordinate_mut()
                .set_col_num(dest_col)
                .set_row_num(dest_row);
            target.add_comments(comment);
        }
    }
    for (r1, c1, r2, c2) in snapshots.1 {
        let (dest_r1, dest_c1, dest_r2, dest_c2) = (
            target_row + r1 - min_row,
            target_col + c1 - min_col,
            target_row + r2 - min_row,
            target_col + c2 - min_col,
        );
        let dest = format!(
            "{}{}:{}{}",
            column_letters(dest_c1),
            dest_r1,
            column_letters(dest_c2),
            dest_r2
        );
        let overlaps = target
            .merge_cells()
            .iter()
            .filter_map(|range| parse_range(&range.range()).ok())
            .any(|(r1, c1, r2, c2)| {
                r1 <= dest_r2 && dest_r1 <= r2 && c1 <= dest_c2 && dest_c1 <= c2
            });
        anyhow::ensure!(
            !overlaps,
            "target merge range overlaps an existing merge: {dest}"
        );
        target.add_merge_cells(dest);
    }
    for (row, height, hidden) in snapshots.2 {
        let dimension = target.row_dimension_mut(target_row + row - min_row);
        if let Some(height) = height {
            dimension.set_height(height);
        }
        if let Some(hidden) = hidden {
            dimension.set_hidden(hidden);
        }
    }
    let imported_target_columns = snapshots
        .3
        .iter()
        .map(|(column, _, _)| target_col + column - min_col)
        .collect::<Vec<_>>();
    for (column, width, hidden) in snapshots.3 {
        let dimension = target.column_dimension_by_number_mut(target_col + column - min_col);
        if let Some(width) = width {
            dimension.set_width(width);
        }
        if let Some(hidden) = hidden {
            dimension.set_hidden(hidden);
        }
    }
    // umya materializes its default-width column record while cells are
    // created. Do not turn that library default into an explicit worksheet
    // dimension; preserve dimensions already present on the target and those
    // explicitly imported from the source.
    target.column_dimensions_mut().retain(|dimension| {
        let column = dimension.col_num();
        preexisting_target_columns.contains(&column)
            || imported_target_columns.contains(&column)
            || dimension.width() != 8.38
            || dimension.hidden()
            || dimension.best_fit()
            || dimension.auto_width()
            || dimension.style() != &Style::default()
    });
    let _ = copy_values;
    Ok(())
}

/// Name of the currently active worksheet, when the stored tab index is in range.
fn active_sheet_name(workbook: &crate::Workbook) -> Option<String> {
    let active_tab = workbook.workbook_view().active_tab() as usize;
    workbook
        .get_sheet_collection()
        .get(active_tab)
        .map(|sheet| sheet.get_name().to_owned())
}

/// umya-spreadsheet keeps the `localSheetId` captured at read time on
/// sheet-scoped defined names (print areas, print titles, user names) and
/// re-emits that value unchanged, and stores `activeTab` as a raw index. After
/// the sheet collection order changes, both must be rewritten from the new
/// positions: scoped names keep pointing at their own worksheet and the
/// previously active sheet stays active. When the active sheet did not
/// survive the change, the stored index is clamped into range.
fn remap_sheet_index_metadata(workbook: &mut crate::Workbook, active_sheet: Option<&str>) {
    let collection = workbook.get_sheet_collection_mut();
    for (index, sheet) in collection.iter_mut().enumerate() {
        for defined_name in sheet.defined_names_mut() {
            if defined_name.has_local_sheet_id() {
                defined_name.set_local_sheet_id(index as u32);
            }
        }
    }
    let count = workbook.sheet_count();
    let new_active_tab = active_sheet
        .and_then(|name| {
            workbook
                .get_sheet_collection()
                .iter()
                .position(|sheet| sheet.get_name() == name)
        })
        .map(|position| position as u32)
        .unwrap_or_else(|| {
            workbook
                .workbook_view()
                .active_tab()
                .min(count.saturating_sub(1) as u32)
        });
    workbook.workbook_view_mut().set_active_tab(new_active_tab);
}

/// `None` and the degenerate `A1` both mean "no frozen panes" (openpyxl
/// semantics).
fn is_freeze_clear_request(cell: Option<&str>) -> bool {
    cell.map(|cell| cell.trim().eq_ignore_ascii_case("A1"))
        .unwrap_or(true)
}

/// Freeze panes with openpyxl semantics: the requested cell is the first
/// scrollable cell, splits follow its zero-based offsets and the active pane
/// follows which splits are non-zero. `set_pane` replaces any previous pane
/// element wholesale, so stale splits from an older freeze cannot survive.
fn apply_freeze_panes(worksheet: &mut umya_spreadsheet::Worksheet, cell: &str) -> Result<()> {
    let (row, column) = parse_cell(cell)?;
    let mut pane = umya_spreadsheet::Pane::default();
    if column > 1 {
        pane.set_horizontal_split((column - 1) as f64);
    }
    if row > 1 {
        pane.set_vertical_split((row - 1) as f64);
    }
    let mut top_left_cell = umya_spreadsheet::Coordinate::default();
    top_left_cell.set_coordinate(cell);
    pane.set_top_left_cell(top_left_cell);
    pane.set_active_pane(match (column > 1, row > 1) {
        (true, true) => umya_spreadsheet::PaneValues::BottomRight,
        (true, false) => umya_spreadsheet::PaneValues::TopRight,
        (false, true) => umya_spreadsheet::PaneValues::BottomLeft,
        (false, false) => bail!("freeze panes at A1 clears the freeze"),
    });
    pane.set_state(umya_spreadsheet::PaneStateValues::Frozen);
    let views = worksheet.sheet_views_mut();
    if views.sheet_view_list().is_empty() {
        views.add_sheet_view_list_mut(umya_spreadsheet::SheetView::default());
    }
    views
        .sheet_view_list_mut()
        .first_mut()
        .context("worksheet has no sheet view")?
        .set_pane(pane);
    Ok(())
}

/// True when any sheet view of the named worksheet currently carries a pane.
fn sheet_has_frozen_pane(workbook: &crate::Workbook, sheet: &str) -> bool {
    workbook
        .get_sheet_collection()
        .iter()
        .filter(|sheet_obj| sheet_obj.get_name() == sheet)
        .flat_map(|sheet_obj| sheet_obj.sheets_views().sheet_view_list().iter())
        .any(|view| view.pane().is_some())
}

/// Excel keeps the hidden sheet-scoped `_xlnm._FilterDatabase` name in step
/// with the worksheet AutoFilter range, and the openpyxl worker regenerates
/// that name on save; the umya candidate must do the same so AutoFilter plans
/// cannot leave a stale range behind after row operations.
fn update_sheet_filter_database(
    workbook: &mut crate::Workbook,
    sheet: &str,
    range: &str,
) -> Result<()> {
    let sheet_index = workbook
        .get_sheet_collection()
        .iter()
        .position(|sheet_obj| sheet_obj.get_name() == sheet)
        .with_context(|| format!("worksheet {sheet:?} not found"))?;
    let address = format!("{sheet}!{}", absolute_range(range)?);
    let target = &mut workbook.get_sheet_collection_mut()[sheet_index];
    // umya's DefinedName::set_address appends to the parsed address list and
    // offers no way to reset it, so the name is rebuilt wholesale. Excel and
    // openpyxl always serialize _FilterDatabase as hidden, including when the
    // AutoFilter is newly added.
    target
        .defined_names_mut()
        .retain(|name| name.name() != "_xlnm._FilterDatabase");
    let mut name = umya_spreadsheet::DefinedName::default();
    name.set_name("_xlnm._FilterDatabase");
    name.set_address(address);
    name.set_local_sheet_id(sheet_index as u32);
    name.set_hidden(true);
    target.add_defined_names(name);
    Ok(())
}

/// Excel removes the sheet-scoped `_xlnm._FilterDatabase` name when the
/// worksheet AutoFilter is cleared; mirror that here.
fn remove_sheet_filter_database(workbook: &mut crate::Workbook, sheet: &str) -> Result<()> {
    let target = workbook
        .get_sheet_collection_mut()
        .iter_mut()
        .find(|sheet_obj| sheet_obj.get_name() == sheet)
        .with_context(|| format!("worksheet {sheet:?} not found"))?;
    target
        .defined_names_mut()
        .retain(|name| name.name() != "_xlnm._FilterDatabase");
    Ok(())
}

/// Render a plan range as the absolute `$A$1:$M$10` form Excel uses inside
/// defined names.
fn absolute_range(range: &str) -> Result<String> {
    range
        .split(':')
        .map(|part| {
            let (row, column) = parse_cell(part)?;
            Ok(format!("${}${}", column_letters(column), row))
        })
        .collect::<Result<Vec<_>>>()
        .map(|parts| parts.join(":"))
}

/// umya re-renders defined-name addresses with quotes around non-ASCII sheet
/// names; both sides of a comparison are dequoted so semantically equal
/// references compare equal. Sheet names containing apostrophes compare by
/// their dequoted text.
fn normalize_defined_name_reference(reference: &str) -> String {
    reference.replace('\'', "")
}

/// Worksheet identity after the whole plan: later RenameSheet operations
/// re-label the sheet targeted by earlier operations (sequential apply
/// order). Mirrors the openpyxl validator's resolution.
fn final_sheet_name(plan: &ExcelEditPlan, source_name: &str) -> String {
    let mut name = source_name.to_owned();
    for operation in &plan.operations {
        if let EditOperation::RenameSheet { old_name, new_name } = operation {
            if old_name == &name {
                name.clone_from(new_name);
            }
        }
    }
    name
}

fn validate_sheet_name(name: &str) -> Result<()> {
    anyhow::ensure!(
        !name.trim().is_empty() && name.chars().count() <= 31,
        "worksheet name must contain 1-31 characters"
    );
    anyhow::ensure!(
        !name
            .chars()
            .any(|character| matches!(character, '[' | ']' | ':' | '*' | '?' | '/' | '\\')),
        "worksheet name contains an Excel-reserved character"
    );
    anyhow::ensure!(
        !name.starts_with('\'') && !name.ends_with('\''),
        "worksheet name cannot begin or end with an apostrophe"
    );
    Ok(())
}

fn normalize_argb(value: &str) -> Result<String> {
    let value = value.trim().trim_start_matches('#');
    let normalized = match value.len() {
        6 => format!("FF{value}"),
        8 => value.to_owned(),
        _ => bail!("fill color must be RRGGBB or AARRGGBB"),
    };
    anyhow::ensure!(
        normalized.bytes().all(|byte| byte.is_ascii_hexdigit()),
        "fill color contains non-hex digits"
    );
    Ok(normalized.to_ascii_uppercase())
}

fn parse_range(value: &str) -> Result<(u32, u32, u32, u32)> {
    let (start, end) = value.split_once(':').unwrap_or((value, value));
    let (start_row, start_col) = parse_cell(start)?;
    let (end_row, end_col) = parse_cell(end)?;
    anyhow::ensure!(
        end_row >= start_row && end_col >= start_col,
        "range end precedes range start"
    );
    Ok((start_row, start_col, end_row, end_col))
}

fn parse_cell(value: &str) -> Result<(u32, u32)> {
    let value = value.trim().trim_start_matches('$');
    let split = value
        .find(|character: char| character.is_ascii_digit())
        .context("cell address has no row number")?;
    let (column, row) = value.split_at(split);
    let col = column
        .bytes()
        .try_fold(0u32, |acc, byte| {
            if !byte.is_ascii_alphabetic() {
                return None;
            }
            acc.checked_mul(26)?
                .checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))
        })
        .context("invalid column in cell address")?;
    let row = row.parse::<u32>().context("invalid row in cell address")?;
    anyhow::ensure!(
        col > 0 && row > 0,
        "cell address must use positive one-based indexes"
    );
    Ok((row, col))
}

fn cell_address(row: u32, mut col: u32) -> String {
    let mut letters = String::new();
    while col > 0 {
        let remainder = (col - 1) % 26;
        letters.insert(0, char::from(b'A' + remainder as u8));
        col = (col - 1) / 26;
    }
    format!("{letters}{row}")
}

fn column_letters(mut col: u32) -> String {
    let mut letters = String::new();
    while col > 0 {
        let remainder = (col - 1) % 26;
        letters.insert(0, char::from(b'A' + remainder as u8));
        col = (col - 1) / 26;
    }
    letters
}
