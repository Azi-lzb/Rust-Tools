//! BetterOffice adapter for edit plans using its XLSX model and calculation engine.

use anyhow::{anyhow, bail, ensure, Context, Result};
use betteroffice_xlsx::{self as bo, Op, SheetId};
use excel_editor::{
    BackendCapability, BackendFailure, EditBackend, EditOperation, EditOperationKind,
    ExcelEditPlan, OutputValidator, WorkbookFormat,
};
use serde_json::{json, Value};
use std::{fs, io::Write, path::Path};

pub struct BetterOfficeAdapter {
    capability: BackendCapability,
}

enum PendingGroup {
    Model(Vec<Op>),
    Batch(Vec<Value>),
}

impl BetterOfficeAdapter {
    pub fn new(capability: BackendCapability) -> Result<Self> {
        ensure!(
            capability.engine == excel_editor::EngineId::BetterOffice,
            "BetterOffice adapter requires a betteroffice capability"
        );
        Ok(Self { capability })
    }

    pub fn apply_bytes(source: &[u8], plan: &ExcelEditPlan) -> Result<Vec<u8>> {
        ensure!(
            plan.format == WorkbookFormat::Xlsx,
            "BetterOffice adapter accepts XLSX only"
        );
        let mut workbook = bo::Workbook::open(source).context("BetterOffice open XLSX")?;
        let catalog = match workbook.read_cells(&bo::ReadRequest::default())? {
            Ok(catalog) => catalog,
            Err(refusal) => bail!(
                "BetterOffice could not read worksheet catalog: {}",
                refusal.failure.message
            ),
        };
        let sheet_ids = catalog
            .sheets
            .iter()
            .enumerate()
            .map(|(index, sheet)| (sheet.name.clone(), SheetId(index as u32)))
            .collect::<std::collections::BTreeMap<_, _>>();

        let mut pending: Option<PendingGroup> = None;
        for operation in &plan.operations {
            match operation {
                EditOperation::SetCellValue { sheet, cell, value } => {
                    let sheet_id = required_sheet(&sheet_ids, sheet)?;
                    let (value, number_format) = betteroffice_cell_value(value)?;
                    let target = json!({
                        "sheetId": format!("sheet:{}", sheet_id.0),
                        "range": {"kind": "a1", "a1": cell}
                    });
                    let mut steps = vec![json!({
                        "op": "setCellValues",
                        "target": target,
                        "values": [[value]],
                    })];
                    if let Some(number_format) = number_format {
                        steps.push(json!({
                            "op": "setNumberFormat",
                            "target": target.clone(),
                            "format": {"type": "custom", "pattern": number_format},
                        }));
                    }
                    push_batch(&mut workbook, &mut pending, steps)?;
                }
                EditOperation::SetFormulaText {
                    sheet,
                    cell,
                    formula,
                } => {
                    let sheet_id = required_sheet(&sheet_ids, sheet)?;
                    push_batch(
                        &mut workbook,
                        &mut pending,
                        vec![json!({
                            "op": "setFormulas",
                            "target": {
                                "sheetId": format!("sheet:{}", sheet_id.0),
                                "range": {"kind": "a1", "a1": cell}
                            },
                            "formulas": [[formula.trim_start_matches('=')]],
                        })],
                    )?;
                }
                EditOperation::InsertRows {
                    sheet,
                    first_row,
                    count,
                } => {
                    push_model(
                        &mut workbook,
                        &mut pending,
                        vec![Op::InsertRows {
                            sheet: *required_sheet(&sheet_ids, sheet)?,
                            at: first_row - 1,
                            count: *count,
                        }],
                    )?;
                }
                EditOperation::DeleteRows {
                    sheet,
                    first_row,
                    count,
                } => {
                    push_model(
                        &mut workbook,
                        &mut pending,
                        vec![Op::DeleteRows {
                            sheet: *required_sheet(&sheet_ids, sheet)?,
                            at: first_row - 1,
                            count: *count,
                        }],
                    )?;
                }
                EditOperation::InsertColumns {
                    sheet,
                    first_column,
                    count,
                } => {
                    push_model(
                        &mut workbook,
                        &mut pending,
                        vec![Op::InsertCols {
                            sheet: *required_sheet(&sheet_ids, sheet)?,
                            at: first_column - 1,
                            count: *count,
                        }],
                    )?;
                }
                EditOperation::DeleteColumns {
                    sheet,
                    first_column,
                    count,
                } => {
                    push_model(
                        &mut workbook,
                        &mut pending,
                        vec![Op::DeleteCols {
                            sheet: *required_sheet(&sheet_ids, sheet)?,
                            at: first_column - 1,
                            count: *count,
                        }],
                    )?;
                }
                EditOperation::SetColumnWidth {
                    sheet,
                    column_index,
                    width,
                } => {
                    let calibrated = calibrate_excel_column_width(*width)?;
                    push_model(
                        &mut workbook,
                        &mut pending,
                        vec![Op::SetColWidth {
                            sheet: *required_sheet(&sheet_ids, sheet)?,
                            col: column_index - 1,
                            width: Some(calibrated),
                        }],
                    )?;
                }
                EditOperation::SetRowHeight {
                    sheet,
                    row_index,
                    height,
                } => {
                    push_model(
                        &mut workbook,
                        &mut pending,
                        vec![Op::SetRowHeight {
                            sheet: *required_sheet(&sheet_ids, sheet)?,
                            row: row_index - 1,
                            height: Some(*height),
                        }],
                    )?;
                }
                other => bail!("BetterOffice adapter cannot execute {:?}", other.kind()),
            }
        }
        flush_pending(&mut workbook, pending)?;
        workbook.save().context("BetterOffice save edited XLSX")
    }
}

impl EditBackend for BetterOfficeAdapter {
    fn capability(&self) -> &BackendCapability {
        &self.capability
    }

    fn apply(&self, plan: &ExcelEditPlan) -> std::result::Result<(), BackendFailure> {
        if plan.format != WorkbookFormat::Xlsx {
            return Err(BackendFailure::unsupported_before_write(
                "BetterOffice adapter accepts XLSX only",
            ));
        }
        if plan.operations.iter().any(|operation| {
            matches!(
                operation,
                EditOperation::InsertRows { .. }
                    | EditOperation::DeleteRows { .. }
                    | EditOperation::InsertColumns { .. }
                    | EditOperation::DeleteColumns { .. }
            ) && (plan
                .workbook
                .features_present
                .contains(&excel_editor::WorkbookFeature::CalcChain)
                || plan
                    .workbook
                    .features_present
                    .contains(&excel_editor::WorkbookFeature::ExternalLinks))
        }) {
            return Err(BackendFailure::unsupported_before_write(
                "BetterOffice refuses structural edits when calcChain or external links are present",
            ));
        }
        if !plan
            .operations
            .iter()
            .all(|operation| supported_operation(operation.kind()))
        {
            return Err(BackendFailure::unsupported_before_write(
                "BetterOffice does not support the complete operation batch",
            ));
        }
        let source = fs::read(&plan.input)
            .map_err(|error| BackendFailure::failed(format!("read BetterOffice input: {error}")))?;
        let candidate = Self::apply_bytes(&source, plan).map_err(|error| {
            BackendFailure::unsupported_before_write(format!(
                "BetterOffice preflight/edit: {error:#}"
            ))
        })?;
        let mut output = fs::OpenOptions::new()
            .write(true)
            .truncate(true)
            .open(&plan.output)
            .map_err(|error| {
                BackendFailure::failed(format!("create BetterOffice candidate: {error}"))
            })?;
        output
            .write_all(&candidate)
            .and_then(|_| output.sync_all())
            .map_err(|error| {
                BackendFailure::failed(format!("write BetterOffice candidate: {error}"))
            })
    }

    fn validate_output(&self, plan: &ExcelEditPlan, candidate_output: &Path) -> Result<()> {
        excel_editor::OpenpyxlXlsxValidator.validate(plan, candidate_output)?;
        excel_editor::UmyaXlsxValidator.validate(plan, candidate_output)
    }
}

pub(crate) fn supported_operation(operation: EditOperationKind) -> bool {
    matches!(
        operation,
        EditOperationKind::SetCellValue
            | EditOperationKind::SetFormulaText
            | EditOperationKind::SetColumnWidth
            | EditOperationKind::SetRowHeight
            | EditOperationKind::DeleteRows
            | EditOperationKind::InsertRows
            | EditOperationKind::DeleteColumns
            | EditOperationKind::InsertColumns
    )
}

fn required_sheet<'a>(
    sheet_ids: &'a std::collections::BTreeMap<String, SheetId>,
    name: &str,
) -> Result<&'a SheetId> {
    sheet_ids
        .get(name)
        .with_context(|| format!("BetterOffice worksheet {name:?} was not found"))
}

fn push_model(
    workbook: &mut bo::Workbook,
    pending: &mut Option<PendingGroup>,
    ops: Vec<Op>,
) -> Result<()> {
    if matches!(pending.as_ref(), Some(PendingGroup::Batch(_))) {
        flush_pending(workbook, pending.take())?;
    }
    match pending {
        Some(PendingGroup::Model(existing)) => existing.extend(ops),
        None => *pending = Some(PendingGroup::Model(ops)),
        Some(PendingGroup::Batch(_)) => unreachable!("batch group was flushed"),
    }
    Ok(())
}

fn push_batch(
    workbook: &mut bo::Workbook,
    pending: &mut Option<PendingGroup>,
    steps: Vec<Value>,
) -> Result<()> {
    if matches!(pending.as_ref(), Some(PendingGroup::Model(_))) {
        flush_pending(workbook, pending.take())?;
    }
    match pending {
        Some(PendingGroup::Batch(existing)) => existing.extend(steps),
        None => *pending = Some(PendingGroup::Batch(steps)),
        Some(PendingGroup::Model(_)) => unreachable!("model group was flushed"),
    }
    Ok(())
}

fn flush_pending(workbook: &mut bo::Workbook, pending: Option<PendingGroup>) -> Result<()> {
    match pending {
        Some(PendingGroup::Model(operations)) => {
            workbook
                .apply_ops(operations, bo::CalculationOptions::default())
                .context("BetterOffice apply structural/dimension operations")?;
        }
        Some(PendingGroup::Batch(steps)) => {
            let request = json!({
                "expectVersion": workbook.version(),
                "steps": steps,
            });
            let response = workbook.apply_edits_json(&request.to_string())?;
            let response: Value =
                serde_json::from_str(&response).context("decode BetterOffice edit response")?;
            ensure!(
                response.get("ok").and_then(Value::as_bool) == Some(true),
                "BetterOffice refused edit batch: {response}"
            );
        }
        None => {}
    }
    Ok(())
}

fn betteroffice_cell_value(value: &excel_editor::CellValue) -> Result<(Value, Option<String>)> {
    let (value, number_format) = match value {
        excel_editor::CellValue::Text(value) => (json!({"kind": "text", "value": value}), None),
        excel_editor::CellValue::Decimal(value) => {
            let number = decimal_for_betteroffice(value)?;
            (json!({"kind": "number", "value": number}), None)
        }
        excel_editor::CellValue::DateTime {
            serial,
            number_format,
        } => {
            let number = serial
                .parse::<f64>()
                .context("parse BetterOffice date/time serial")?;
            ensure!(
                number.is_finite(),
                "BetterOffice date/time serial must be finite"
            );
            (
                json!({"kind": "number", "value": number}),
                Some(number_format.clone()),
            )
        }
        excel_editor::CellValue::Boolean(value) => (json!({"kind": "bool", "value": value}), None),
        excel_editor::CellValue::Error(value) => (json!({"kind": "error", "value": value}), None),
        excel_editor::CellValue::Blank => (json!({"kind": "empty"}), None),
    };
    Ok((value, number_format))
}

fn decimal_for_betteroffice(value: &str) -> Result<f64> {
    let number = value.parse::<f64>().context("parse BetterOffice decimal")?;
    ensure!(number.is_finite(), "BetterOffice decimal must be finite");
    let significant = value
        .chars()
        .filter(|character| character.is_ascii_digit())
        .skip_while(|character| *character == '0')
        .count();
    ensure!(
        significant <= 15,
        "BetterOffice numeric storage is limited to 15 significant decimal digits"
    );
    Ok(number)
}

/// BetterOffice stores OOXML column widths in the same Excel character-width
/// unit (the raw `<col width>` value), not pixels. The route layer deliberately
/// applies the identity conversion here and verifies the serialized width
/// through the candidate validator, avoiding a pixel round trip.
fn calibrate_excel_column_width(width: f64) -> Result<f64> {
    ensure!(
        width.is_finite() && (0.0..=255.0).contains(&width),
        "Excel column width must be in the 0..=255 character-width range"
    );
    Ok(width)
}

#[cfg(test)]
mod tests {
    use super::calibrate_excel_column_width;

    #[test]
    fn column_width_calibration_preserves_raw_excel_character_width() {
        for width in [0.0, 8.43, 13.0, 255.0] {
            assert_eq!(calibrate_excel_column_width(width).unwrap(), width);
        }
    }

    #[test]
    fn column_width_calibration_rejects_non_finite_or_out_of_range_values() {
        for width in [-0.01, 255.01, f64::NAN, f64::INFINITY] {
            assert!(calibrate_excel_column_width(width).is_err());
        }
    }
}
