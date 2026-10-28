//! Duplicate detection and edit-plan construction for menu 2.

use crate::config::DedupTask;
use anyhow::{bail, Context, Result};
use excel_editor::{
    inspect_xlsx_profile, CellValue, EditOperation, ExcelEditPlan, PreservationRequirement,
    WorkbookFeature, WorkbookFormat,
};
use excel_reader::{
    read_value_columns_with_iso_dates, read_workbook, read_xlsx_effective_cell_number_formats,
    uses_1904_date_system, CellNumberFormatRequest, CellValue as ReaderCellValue, SheetDump,
};
use serde::{Deserialize, Serialize};
use std::{
    collections::{BTreeMap, BTreeSet, HashSet},
    fs,
    path::{Path, PathBuf},
    sync::atomic::{AtomicU64, Ordering},
};

type DateIsoMap = std::collections::HashMap<(String, String), String>;
static DUPLICATE_RESULT_TEMP_SEQUENCE: AtomicU64 = AtomicU64::new(0);

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum DuplicateAction {
    Check,
    Mark,
    Delete,
}

#[derive(Clone, Debug, Default, Serialize)]
pub struct DuplicatePlanSummary {
    pub workbook: Option<PathBuf>,
    pub hit_sheets: usize,
    pub duplicate_rows: usize,
    pub merged_tasks: usize,
    pub skipped_tasks: usize,
    pub plan: Option<ExcelEditPlan>,
}

#[derive(Clone, Debug, Default, Serialize)]
pub struct DuplicatePrecheck {
    pub ok: usize,
    pub warn: usize,
    pub fail: usize,
    pub issues: Vec<String>,
}

/// A serialized 2-5 append plan. Existing target workbooks are edited through
/// ExcelEditPlan; a missing target is created as a new XLSX with the shared
/// writer. Both paths produce a candidate file and leave configured targets
/// untouched until the caller chooses to replace them.
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct AppendTargetPlan {
    pub target: PathBuf,
    pub candidate: PathBuf,
    pub target_existed: bool,
    /// SHA-256 of an existing target when the plan was created. Execution
    /// uses it as an optimistic concurrency check before replacing the file.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub target_sha256: Option<String>,
    pub task_ok: usize,
    pub task_skipped: usize,
    pub appended_rows: usize,
    pub duplicate_rows_filtered: usize,
    pub merged_tasks: usize,
    pub issues: Vec<String>,
    pub edit_plan: Option<ExcelEditPlan>,
    pub new_workbook: Option<NewAppendWorkbook>,
    /// Replacement sheets for PT-1-3/PT-1-4. The workbook-level edit plan
    /// imports these into an existing target while preserving other sheets.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub replacement_workbook: Option<NewAppendWorkbook>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct NewAppendWorkbook {
    pub sheets: Vec<NewAppendSheet>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct NewAppendSheet {
    pub name: String,
    /// Rows include the header at index zero.
    pub rows: Vec<Vec<AppendCell>>,
    /// Match the legacy wide-summary DataFrame header style on row zero.
    #[serde(default)]
    pub style_header_row: bool,
    /// Match pandas/openpyxl's default bold table header without adding the
    /// wide-summary borders and alignment.
    #[serde(default)]
    pub header_bold: bool,
    /// Apply the legacy data-date number format to this zero-based column.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_date_column: Option<usize>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(tag = "kind", content = "value", rename_all = "snake_case")]
pub enum AppendCell {
    Value(CellValue),
    Formula(String),
    Number {
        serial: String,
        number_format: String,
    },
    DateTime {
        serial: String,
        number_format: String,
    },
}

/// Build a plan for the two comment-based duplicate actions.
pub fn build_comment_duplicate_plan(
    input: &Path,
    output: &Path,
    action: DuplicateAction,
) -> Result<DuplicatePlanSummary> {
    ensure_xlsx(input, output)?;
    let dump = read_workbook(input).with_context(|| format!("read {}", input.display()))?;
    let date_values = date_iso_map(input, &dump.sheets)?;
    let profile = inspect_xlsx_profile(input)?;
    let mut operations = Vec::new();
    let mut hit_sheets = 0;
    let mut duplicate_rows = 0;
    for sheet in &dump.sheets {
        let a1 = sheet
            .comments
            .get("A1")
            .map(|value| normalize(value))
            .unwrap_or_default();
        if !a1.contains("源数据") {
            continue;
        }
        hit_sheets += 1;
        let key_cols = comment_key_columns(sheet);
        let duplicate = duplicate_rows_for_sheet(sheet, &key_cols, &date_values);
        duplicate_rows += duplicate.len();
        append_duplicate_operations(&mut operations, sheet, duplicate, action);
    }
    let mut plan = make_plan(input, output, profile, operations, "comment-duplicate")?;
    if let Some(plan) = plan.as_mut() {
        plan.preserve = duplicate_preservation(action);
    }
    Ok(DuplicatePlanSummary {
        workbook: Some(input.to_path_buf()),
        hit_sheets,
        duplicate_rows,
        merged_tasks: 0,
        skipped_tasks: 0,
        plan,
    })
}

/// Build a combined plan for config-driven check/delete tasks. Duplicate
/// definitions with the same source sheet and key columns are applied once.
pub fn build_config_duplicate_plan(
    tasks: &[DedupTask],
    output_by_input: &BTreeMap<PathBuf, PathBuf>,
    action: DuplicateAction,
) -> Result<Vec<DuplicatePlanSummary>> {
    let mut groups = BTreeMap::<PathBuf, Vec<&DedupTask>>::new();
    let mut seen = HashSet::new();
    let mut merged = 0usize;
    let mut skipped = 0usize;
    for task in tasks.iter().filter(|task| task.enabled) {
        let Some(input) = task.source_wb.as_ref() else {
            skipped += 1;
            continue;
        };
        if !input.is_file()
            || !task.source_wb.as_ref().is_some_and(|p| is_xlsx(p))
            || task.source_ws.trim().is_empty()
        {
            skipped += 1;
            continue;
        }
        let signature = (input.clone(), task.source_ws.clone(), task.key_cols.clone());
        if !seen.insert(signature) {
            merged += 1;
            continue;
        }
        groups.entry(input.clone()).or_default().push(task);
    }
    let mut summaries = Vec::new();
    for (input, grouped) in groups {
        let output = output_by_input
            .get(&input)
            .with_context(|| format!("missing output mapping for {}", input.display()))?;
        let dump = read_workbook(&input).with_context(|| format!("read {}", input.display()))?;
        let date_values = date_iso_map(&input, &dump.sheets)?;
        let profile = inspect_xlsx_profile(&input)?;
        let mut operations = Vec::new();
        let mut hit_sheets = 0;
        let mut duplicate_count = 0;
        let mut surviving_rows_by_sheet = BTreeMap::<String, Vec<usize>>::new();
        for task in grouped {
            let Some(sheet) = dump
                .sheets
                .iter()
                .find(|sheet| sheet.name == task.source_ws)
            else {
                skipped += 1;
                continue;
            };
            hit_sheets += 1;
            let key_cols = if task.key_cols.is_empty() {
                (1..=used_column_count(sheet)).collect()
            } else {
                task.key_cols.clone()
            };
            // Python applies config rows sequentially to the already-mutated
            // worksheet. For delete mode, later tasks therefore see row
            // numbers after earlier deletions. Keep a virtual row map so each
            // planned DeleteRows operation uses the same intermediate
            // coordinates that Python would see.
            let duplicate = if action == DuplicateAction::Delete {
                let last_row = used_row_count(sheet);
                let surviving = surviving_rows_by_sheet
                    .entry(sheet.name.clone())
                    .or_insert_with(|| (1..=last_row).collect());
                duplicate_rows_for_surviving_rows(sheet, &key_cols, surviving, &date_values)
            } else {
                duplicate_rows_for_sheet(sheet, &key_cols, &date_values)
            };
            duplicate_count += duplicate.len();
            append_duplicate_operations(&mut operations, sheet, duplicate.clone(), action);
            if action == DuplicateAction::Delete && !duplicate.is_empty() {
                if let Some(surviving) = surviving_rows_by_sheet.get_mut(&sheet.name) {
                    let duplicate_positions = duplicate.into_iter().collect::<BTreeSet<_>>();
                    let mut current_row = 1usize;
                    surviving.retain(|_| {
                        let keep = !duplicate_positions.contains(&current_row);
                        current_row += 1;
                        keep
                    });
                }
            }
        }
        let mut plan = make_plan(&input, output, profile, operations, "config-duplicate")?;
        if let Some(plan) = plan.as_mut() {
            plan.preserve = duplicate_preservation(action);
        }
        summaries.push(DuplicatePlanSummary {
            workbook: Some(input),
            hit_sheets,
            duplicate_rows: duplicate_count,
            merged_tasks: merged,
            skipped_tasks: skipped,
            plan,
        });
    }
    Ok(summaries)
}

/// Create a new data-only duplicate result workbook.
/// The source is never opened for writing; unsupported workbook features fail closed.
pub fn write_config_duplicate_check_candidate(
    input: &Path,
    output: &Path,
    plan: Option<&ExcelEditPlan>,
) -> Result<PathBuf> {
    write_config_duplicate_result_candidate(input, output, plan, DuplicateAction::Check)
}

/// Create a new, data-only workbook with configured duplicate rows removed.
/// The source is never opened for writing; unsupported workbook features fail closed.
pub fn write_config_duplicate_delete_candidate(
    input: &Path,
    output: &Path,
    plan: Option<&ExcelEditPlan>,
) -> Result<PathBuf> {
    write_config_duplicate_result_candidate(input, output, plan, DuplicateAction::Delete)
}

fn write_config_duplicate_result_candidate(
    input: &Path,
    output: &Path,
    plan: Option<&ExcelEditPlan>,
    action: DuplicateAction,
) -> Result<PathBuf> {
    anyhow::ensure!(
        matches!(action, DuplicateAction::Check | DuplicateAction::Delete),
        "unsupported configured duplicate output action"
    );
    ensure_xlsx(input, output)?;
    anyhow::ensure!(
        !output.exists(),
        "refusing to overwrite {}",
        output.display()
    );
    if let Some(plan) = plan {
        anyhow::ensure!(plan.input == input, "duplicate plan input changed");
        anyhow::ensure!(plan.output == output, "duplicate plan output changed");
    }

    let source = read_workbook(input).with_context(|| format!("read {}", input.display()))?;
    let input_number_formats = read_numeric_cell_formats(input, &source)?;
    let mut duplicate_rows = BTreeMap::<String, BTreeSet<usize>>::new();
    let mut current_rows_by_sheet = BTreeMap::<String, Vec<usize>>::new();
    if action == DuplicateAction::Delete {
        for sheet in &source.sheets {
            current_rows_by_sheet.insert(sheet.name.clone(), (1..=used_row_count(sheet)).collect());
        }
    }
    if let Some(plan) = plan {
        for operation in &plan.operations {
            match (action, operation) {
                (DuplicateAction::Check, EditOperation::SetCellFillColor { sheet, cell, argb }) => {
                    anyhow::ensure!(
                        argb.eq_ignore_ascii_case("FFFF0000"),
                        "unexpected duplicate check fill color {argb:?}"
                    );
                    let (row, _) = parse_a1(cell)
                        .with_context(|| format!("invalid duplicate mark cell {cell:?}"))?;
                    anyhow::ensure!(
                        source
                            .sheets
                            .iter()
                            .any(|candidate| candidate.name == *sheet),
                        "duplicate mark targets unknown worksheet {sheet:?}"
                    );
                    duplicate_rows.entry(sheet.clone()).or_default().insert(row);
                }
                (
                    DuplicateAction::Delete,
                    EditOperation::DeleteRows {
                        sheet,
                        first_row,
                        count,
                    },
                ) => {
                    let rows = current_rows_by_sheet.get_mut(sheet).with_context(|| {
                        format!("delete plan targets unknown worksheet {sheet:?}")
                    })?;
                    let start = usize::try_from(*first_row)
                        .context("duplicate deletion start row exceeds usize")?
                        .checked_sub(1)
                        .context("duplicate deletion row must be 1-based")?;
                    let end = start
                        .checked_add(usize::try_from(*count).context("row count exceeds usize")?)
                        .context("duplicate deletion range overflows")?;
                    anyhow::ensure!(
                        end <= rows.len(),
                        "duplicate deletion range exceeds current worksheet rows"
                    );
                    rows.drain(start..end);
                }
                _ => anyhow::bail!(
                    "duplicate plan contains an operation incompatible with {action:?}"
                ),
            }
        }
    }
    let output_rows_by_sheet = if action == DuplicateAction::Delete {
        current_rows_by_sheet
            .into_iter()
            .map(|(sheet, source_rows)| {
                let output_rows = source_rows
                    .into_iter()
                    .enumerate()
                    .map(|(output_index, source_row)| (source_row, output_index + 1))
                    .collect::<BTreeMap<_, _>>();
                (sheet, output_rows)
            })
            .collect::<BTreeMap<_, _>>()
    } else {
        BTreeMap::new()
    };

    let red_fill =
        excel_writer::Format::new().set_background_color(excel_writer::Color::RGB(0xFF_00_00));
    let mut number_formats = BTreeMap::<String, excel_writer::Format>::new();
    let mut red_number_formats = BTreeMap::<String, excel_writer::Format>::new();
    let mut expected_number_formats = BTreeMap::<String, BTreeMap<String, String>>::new();
    let mut workbook = excel_writer::Workbook::new();
    for sheet in &source.sheets {
        anyhow::ensure!(
            sheet.formula_status == "read_by_calamine",
            "formula metadata unavailable for duplicate output worksheet {:?}",
            sheet.name
        );
        anyhow::ensure!(
            sheet.formulas.is_empty(),
            "pure-data duplicate output does not support formula cells in worksheet {:?}",
            sheet.name
        );
        anyhow::ensure!(
            sheet.comments_status == "read_by_namespace_aware_ooxml_supplement"
                && sheet.comments.is_empty(),
            "pure-data duplicate output cannot preserve comments in worksheet {:?}",
            sheet.name
        );
        anyhow::ensure!(
            sheet.merge_status == "read_by_calamine" && sheet.merges.is_empty(),
            "pure-data duplicate output cannot preserve merged cells in worksheet {:?}",
            sheet.name
        );
        anyhow::ensure!(
            sheet.conditional_formatting_status == "read_by_namespace_aware_ooxml_supplement"
                && sheet.conditional_formatting.is_empty(),
            "pure-data duplicate output cannot preserve conditional formatting in worksheet {:?}",
            sheet.name
        );
        anyhow::ensure!(
            sheet.data_validations_status == "read_by_namespace_aware_ooxml_supplement"
                && sheet.data_validations.is_empty(),
            "pure-data duplicate output cannot preserve data validations in worksheet {:?}",
            sheet.name
        );
        anyhow::ensure!(
            !sheet.values.values().any(|value| {
                value.get("type").and_then(serde_json::Value::as_str) == Some("error")
            }),
            "typed Excel error cells cannot be written by the duplicate output engine in worksheet {:?}",
            sheet.name
        );

        let marked_rows = duplicate_rows.get(&sheet.name);
        let max_row = used_row_count(sheet);
        let max_column = used_column_count(sheet);
        let worksheet = workbook.add_worksheet();
        worksheet.set_name(&sheet.name)?;
        for (address, value) in &sheet.values {
            let Some((row, column)) = parse_a1(address) else {
                anyhow::bail!("invalid source cell address {address:?}");
            };
            let output_row = if action == DuplicateAction::Delete {
                output_rows_by_sheet
                    .get(&sheet.name)
                    .and_then(|rows| rows.get(&row))
                    .copied()
            } else {
                Some(row)
            };
            let Some(output_row) = output_row else {
                continue;
            };
            let row_index = u32::try_from(output_row.saturating_sub(1))
                .context("worksheet row exceeds XLSX limit")?;
            let column_index = u16::try_from(column.saturating_sub(1))
                .context("worksheet column exceeds XLSX limit")?;
            let marked = action == DuplicateAction::Check
                && marked_rows.is_some_and(|rows| rows.contains(&row));
            let value = ReaderCellValue::from_probe_json(value)
                .with_context(|| format!("decode {}/{}", sheet.name, address))?;
            let number_format = match &value {
                ReaderCellValue::Number(_)
                | ReaderCellValue::Integer(_)
                | ReaderCellValue::DateTimeSerial(_) => Some(
                    input_number_formats
                        .get(&sheet.name)
                        .and_then(|formats| formats.get(address))
                        .map(String::as_str)
                        .with_context(|| {
                            format!("number format metadata missing at {}!{address}", sheet.name)
                        })?,
                ),
                _ => None,
            };
            if let Some(number_format) = number_format {
                expected_number_formats
                    .entry(sheet.name.clone())
                    .or_default()
                    .insert(cell_address(output_row, column), number_format.to_owned());
            }
            match value {
                ReaderCellValue::Empty => {
                    if marked {
                        worksheet.write_blank(row_index, column_index, &red_fill)?;
                    }
                }
                ReaderCellValue::Text(value) => {
                    if marked {
                        worksheet.write_string_with_format(
                            row_index,
                            column_index,
                            &value,
                            &red_fill,
                        )?;
                    } else {
                        worksheet.write_string(row_index, column_index, &value)?;
                    }
                }
                ReaderCellValue::Number(value) => {
                    anyhow::ensure!(
                        value.is_finite(),
                        "non-finite value at {}/{}",
                        sheet.name,
                        address
                    );
                    let number_format = number_format.expect("numeric value has a number format");
                    if marked {
                        let format = red_number_formats
                            .entry(number_format.to_owned())
                            .or_insert_with(|| {
                                excel_writer::Format::new()
                                    .set_background_color(excel_writer::Color::RGB(0xFF_00_00))
                                    .set_num_format(number_format)
                            });
                        worksheet.write_number_with_format(
                            row_index,
                            column_index,
                            value,
                            format,
                        )?;
                    } else {
                        let format = number_formats
                            .entry(number_format.to_owned())
                            .or_insert_with(|| {
                                excel_writer::Format::new().set_num_format(number_format)
                            });
                        worksheet.write_number_with_format(
                            row_index,
                            column_index,
                            value,
                            format,
                        )?;
                    }
                }
                ReaderCellValue::Integer(value) => {
                    let number = value as f64;
                    anyhow::ensure!(
                        number as i128 == value as i128,
                        "integer value at {}/{} cannot be represented without precision loss",
                        sheet.name,
                        address
                    );
                    let number_format = number_format.expect("integer value has a number format");
                    if marked {
                        let format = red_number_formats
                            .entry(number_format.to_owned())
                            .or_insert_with(|| {
                                excel_writer::Format::new()
                                    .set_background_color(excel_writer::Color::RGB(0xFF_00_00))
                                    .set_num_format(number_format)
                            });
                        worksheet.write_number_with_format(
                            row_index,
                            column_index,
                            number,
                            format,
                        )?;
                    } else {
                        let format = number_formats
                            .entry(number_format.to_owned())
                            .or_insert_with(|| {
                                excel_writer::Format::new().set_num_format(number_format)
                            });
                        worksheet.write_number_with_format(
                            row_index,
                            column_index,
                            number,
                            format,
                        )?;
                    }
                }
                ReaderCellValue::Boolean(value) => {
                    if marked {
                        worksheet.write_boolean_with_format(
                            row_index,
                            column_index,
                            value,
                            &red_fill,
                        )?;
                    } else {
                        worksheet.write_boolean(row_index, column_index, value)?;
                    }
                }
                ReaderCellValue::DateTimeSerial(value) => {
                    anyhow::ensure!(
                        value.is_finite(),
                        "non-finite date at {}/{}",
                        sheet.name,
                        address
                    );
                    let number_format = number_format.expect("date serial has a number format");
                    if marked {
                        let format = red_number_formats
                            .entry(number_format.to_owned())
                            .or_insert_with(|| {
                                excel_writer::Format::new()
                                    .set_background_color(excel_writer::Color::RGB(0xFF_00_00))
                                    .set_num_format(number_format)
                            });
                        worksheet.write_number_with_format(
                            row_index,
                            column_index,
                            value,
                            format,
                        )?;
                    } else {
                        let format = number_formats
                            .entry(number_format.to_owned())
                            .or_insert_with(|| {
                                excel_writer::Format::new().set_num_format(number_format)
                            });
                        worksheet.write_number_with_format(
                            row_index,
                            column_index,
                            value,
                            format,
                        )?;
                    }
                }
                ReaderCellValue::DateTimeIso(_) | ReaderCellValue::DurationIso(_) => {
                    anyhow::bail!(
                        "ISO-typed date or duration at {}/{} cannot be preserved by the duplicate output engine",
                        sheet.name,
                        address
                    );
                }
                ReaderCellValue::Error(value) => {
                    anyhow::bail!(
                        "typed Excel error at {}/{} cannot be written as a duplicate result value: {value}",
                        sheet.name,
                        address
                    );
                }
            }
        }

        if action == DuplicateAction::Check {
            if let Some(marked_rows) = marked_rows {
                for row in marked_rows {
                    for column in 1..=max_column {
                        let address = cell_address(*row, column);
                        if !sheet.values.contains_key(&address) {
                            let row_index = u32::try_from(row.saturating_sub(1))
                                .context("worksheet row exceeds XLSX limit")?;
                            let column_index = u16::try_from(column.saturating_sub(1))
                                .context("worksheet column exceeds XLSX limit")?;
                            worksheet.write_blank(row_index, column_index, &red_fill)?;
                        }
                    }
                }
            }
        }
        anyhow::ensure!(
            max_row <= 1_048_576 && max_column <= 16_384,
            "source worksheet {:?} exceeds XLSX dimensions",
            sheet.name
        );
    }

    let profile = inspect_xlsx_profile(input)?;
    anyhow::ensure!(
        profile.inventory_complete,
        "duplicate output cannot rebuild a workbook with unknown OOXML parts"
    );
    let unsupported_features = profile
        .features_present
        .iter()
        .filter(|feature| **feature != WorkbookFeature::Relationships)
        .map(|feature| format!("{feature:?}"))
        .collect::<Vec<_>>();
    anyhow::ensure!(
        unsupported_features.is_empty(),
        "duplicate output cannot rebuild workbook with unsupported features: {}",
        unsupported_features.join(", ")
    );
    let defined_names = excel_reader::read_defined_names(input)?;
    anyhow::ensure!(
        defined_names.is_empty(),
        "duplicate output cannot preserve workbook defined names"
    );

    let parent = output
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    fs::create_dir_all(parent)?;
    let candidate = reserve_duplicate_result_temp(output)?;
    let validation = (|| -> Result<()> {
        workbook
            .save(&candidate)
            .with_context(|| format!("write duplicate result candidate {}", candidate.display()))?;
        let reopened = read_workbook(&candidate).with_context(|| {
            format!("reopen duplicate result candidate {}", candidate.display())
        })?;
        anyhow::ensure!(
            reopened.sheets.len() == source.sheets.len(),
            "duplicate result output worksheet count changed"
        );
        for (expected, actual) in source.sheets.iter().zip(&reopened.sheets) {
            anyhow::ensure!(
                expected.name == actual.name,
                "duplicate result output worksheet order/name changed"
            );
            let expected_values = if action == DuplicateAction::Delete {
                let rows = output_rows_by_sheet
                    .get(&expected.name)
                    .context("duplicate deletion output row map is missing")?;
                let mut values = BTreeMap::new();
                for (address, value) in &expected.values {
                    let (source_row, column) = parse_a1(address)
                        .with_context(|| format!("invalid source cell address {address:?}"))?;
                    if let Some(output_row) = rows.get(&source_row) {
                        values.insert(cell_address(*output_row, column), value.clone());
                    }
                }
                values
            } else {
                expected.values.clone()
            };
            anyhow::ensure!(
                expected_values == actual.values,
                "duplicate result output values/types differ in worksheet {:?}",
                expected.name
            );
            anyhow::ensure!(
                actual.formulas.is_empty() && actual.comments.is_empty(),
                "duplicate result output unexpectedly contains formulas or comments"
            );
        }
        let actual_number_formats = read_numeric_cell_formats(&candidate, &reopened)?;
        anyhow::ensure!(
            expected_number_formats == actual_number_formats,
            "duplicate result output numeric number formats differ from the source rows"
        );
        Ok(())
    })();
    if let Err(error) = validation {
        let _ = fs::remove_file(&candidate);
        return Err(error);
    }
    if let Err(error) = crate::file_commit::create_target_from_candidate(&candidate, output) {
        let _ = fs::remove_file(&candidate);
        return Err(error);
    }
    let _ = fs::remove_file(&candidate);
    Ok(output.to_path_buf())
}

type NumberFormatsBySheet = BTreeMap<String, BTreeMap<String, String>>;

fn read_numeric_cell_formats(
    path: &Path,
    workbook: &excel_reader::WorkbookDump,
) -> Result<NumberFormatsBySheet> {
    let requests = workbook
        .sheets
        .iter()
        .flat_map(|sheet| {
            sheet.values.iter().filter_map(move |(address, value)| {
                matches!(
                    value.get("type").and_then(serde_json::Value::as_str),
                    Some("number" | "integer" | "datetime_serial")
                )
                .then(|| CellNumberFormatRequest {
                    sheet_name: sheet.name.clone(),
                    cell_reference: address.clone(),
                })
            })
        })
        .collect::<Vec<_>>();
    let metadata = read_xlsx_effective_cell_number_formats(path, &requests)
        .with_context(|| format!("read numeric cell formats from {}", path.display()))?;
    anyhow::ensure!(
        metadata.len() == requests.len(),
        "number-format reader returned incomplete metadata for {}",
        path.display()
    );
    let mut formats = NumberFormatsBySheet::new();
    for info in metadata {
        let format_code = info.format_code.with_context(|| {
            format!(
                "unsupported built-in number format id {} at {}!{}",
                info.number_format_id, info.sheet_name, info.cell_reference
            )
        })?;
        formats
            .entry(info.sheet_name)
            .or_default()
            .insert(info.cell_reference, format_code);
    }
    Ok(formats)
}

fn reserve_duplicate_result_temp(output: &Path) -> Result<PathBuf> {
    let parent = output
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let stem = output
        .file_stem()
        .and_then(|stem| stem.to_str())
        .unwrap_or("duplicate-result");
    for _ in 0..1000 {
        let sequence = DUPLICATE_RESULT_TEMP_SEQUENCE.fetch_add(1, Ordering::Relaxed);
        let candidate = parent.join(format!(
            ".{stem}.pytools-dedup-{}-{sequence}.tmp.xlsx",
            std::process::id()
        ));
        match fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&candidate)
        {
            Ok(file) => {
                drop(file);
                return Ok(candidate);
            }
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => {
                return Err(error).with_context(|| {
                    format!("reserve duplicate result candidate {}", candidate.display())
                });
            }
        }
    }
    bail!(
        "cannot reserve a unique duplicate result candidate beside {}",
        output.display()
    )
}

fn duplicate_preservation(action: DuplicateAction) -> BTreeSet<PreservationRequirement> {
    // Match the contract to the operation. Marking changes targeted fills;
    // deleting rows must also move surviving row-attached objects and
    // references. Rewriting the XLSX package means byte-identical package
    // parts are not a valid requirement for either operation.
    match action {
        DuplicateAction::Check | DuplicateAction::Mark => BTreeSet::from([
            PreservationRequirement::CellValuesAndTypes,
            PreservationRequirement::FormulaText,
            PreservationRequirement::CommentsAndPositions,
            PreservationRequirement::Styles,
            PreservationRequirement::MergedRanges,
            PreservationRequirement::SheetIdentityAndOrder,
        ]),
        DuplicateAction::Delete => BTreeSet::from([
            PreservationRequirement::CellValuesAndTypes,
            PreservationRequirement::FormulaText,
            PreservationRequirement::CommentsAndPositions,
            PreservationRequirement::Styles,
            PreservationRequirement::MergedRanges,
            PreservationRequirement::DefinedNames,
            PreservationRequirement::RowAndColumnDimensions,
            PreservationRequirement::Relationships,
            PreservationRequirement::SheetIdentityAndOrder,
        ]),
    }
}

pub fn precheck_dedup_tasks(tasks: &[DedupTask]) -> DuplicatePrecheck {
    let mut result = DuplicatePrecheck::default();
    let mut source_sheets_by_path =
        BTreeMap::<PathBuf, std::result::Result<BTreeSet<String>, String>>::new();
    for task in tasks.iter().filter(|task| task.enabled) {
        let mut errors = Vec::new();
        let Some(input) = task.source_wb.as_ref() else {
            errors.push("source workbook is empty".to_owned());
            result.fail += 1;
            result.issues.extend(
                errors
                    .into_iter()
                    .map(|e| format!("row {}: {e}", task.row_no)),
            );
            continue;
        };
        if !input.is_file() {
            errors.push(format!(
                "source workbook does not exist: {}",
                input.display()
            ));
        }
        if !is_xlsx(input) {
            errors.push(format!(
                "only .xlsx is currently routed to the Rust edit candidate: {}",
                input.display()
            ));
        }
        if task.source_ws.trim().is_empty() {
            errors.push("source worksheet is empty".to_owned());
        }
        if !errors.is_empty() {
            result.fail += 1;
            result.issues.extend(
                errors
                    .into_iter()
                    .map(|e| format!("row {}: {e}", task.row_no)),
            );
            continue;
        }
        let sheet_names = source_sheets_by_path
            .entry(input.clone())
            .or_insert_with(|| {
                read_workbook(input)
                    .with_context(|| format!("read {}", input.display()))
                    .map(|workbook| {
                        workbook
                            .sheets
                            .into_iter()
                            .map(|sheet| sheet.name)
                            .collect::<BTreeSet<_>>()
                    })
                    .map_err(|error| format!("{error:#}"))
            });
        match sheet_names {
            Ok(sheets) if sheets.contains(&task.source_ws) => {
                result.ok += 1;
                if task.key_cols.is_empty() {
                    result.warn += 1;
                    result.issues.push(format!("row {}: key columns are empty; the complete used row will form the duplicate key",task.row_no));
                }
            }
            Ok(_) => {
                result.fail += 1;
                result.issues.push(format!(
                    "row {}: worksheet {:?} not found",
                    task.row_no, task.source_ws
                ));
            }
            Err(error) => {
                result.fail += 1;
                result.issues.push(format!("row {}: {error}", task.row_no));
            }
        }
    }
    result
}

/// Prepare menu 2-5 as one candidate per target workbook. Config duplicates
/// are merged in source order. Existing targets become evidence-gated edit
/// plans; absent targets become new-workbook payloads for rust_xlsxwriter.
pub fn build_append_by_config_plans(
    tasks: &[DedupTask],
    output_dir: &Path,
) -> Result<Vec<AppendTargetPlan>> {
    let mut seen = HashSet::new();
    let mut merged = 0usize;
    let mut grouped = BTreeMap::<PathBuf, Vec<&DedupTask>>::new();
    for task in tasks.iter().filter(|task| task.enabled) {
        let (Some(source), Some(target)) = (&task.source_wb, &task.target_wb) else {
            continue;
        };
        let signature = (
            source.clone(),
            task.source_ws.clone(),
            target.clone(),
            task.target_ws.clone(),
            task.key_cols.clone(),
        );
        if !seen.insert(signature) {
            merged += 1;
            continue;
        }
        grouped.entry(target.clone()).or_default().push(task);
    }

    let mut reports = Vec::new();
    for (target, target_tasks) in grouped {
        anyhow::ensure!(
            target
                .extension()
                .and_then(|ext| ext.to_str())
                .is_some_and(|ext| ext.eq_ignore_ascii_case("xlsx")),
            "2-5 Rust append targets must be .xlsx: {}",
            target.display()
        );
        let target_existed = target.is_file();
        if target.exists() && !target_existed {
            anyhow::bail!(
                "target path exists but is not a regular file: {}",
                target.display()
            );
        }
        let target_sha256 = if target_existed {
            Some(crate::file_commit::sha256_file(&target)?)
        } else {
            None
        };
        let target_dump = if target_existed {
            Some(
                read_workbook(&target)
                    .with_context(|| format!("read append target {}", target.display()))?,
            )
        } else {
            None
        };
        let target_date_values = match target_dump.as_ref() {
            Some(dump) => date_iso_map(&target, &dump.sheets)?,
            None => DateIsoMap::default(),
        };
        let target_1904 = if target_existed {
            uses_1904_date_system(&target)?
        } else {
            false
        };
        let candidate = append_candidate_path(&target, output_dir);
        let mut state = BTreeMap::<String, AppendSheetState>::new();
        let mut operations = Vec::new();
        let mut new_sheets = Vec::<NewAppendSheet>::new();
        let mut report = AppendTargetPlan {
            target: target.clone(),
            candidate: candidate.clone(),
            target_existed,
            target_sha256: target_sha256.clone(),
            task_ok: 0,
            task_skipped: 0,
            appended_rows: 0,
            duplicate_rows_filtered: 0,
            merged_tasks: merged,
            issues: Vec::new(),
            edit_plan: None,
            new_workbook: None,
            replacement_workbook: None,
        };

        for task in target_tasks {
            let Some(source_path) = task.source_wb.as_ref() else {
                continue;
            };
            let target_sheet_name = &task.target_ws;
            if target_sheet_name.trim().is_empty() {
                report.task_skipped += 1;
                report
                    .issues
                    .push(format!("row {}: target worksheet is empty", task.row_no));
                continue;
            }
            if !source_path.is_file() || !is_append_source(source_path) {
                report.task_skipped += 1;
                report.issues.push(format!(
                    "row {}: source workbook is missing or not xlsx/xlsm: {}",
                    task.row_no,
                    source_path.display()
                ));
                continue;
            }
            let source = match read_workbook(source_path) {
                Ok(workbook) => workbook,
                Err(error) => {
                    report.task_skipped += 1;
                    report.issues.push(format!(
                        "row {}: source read failed: {error:#}",
                        task.row_no
                    ));
                    continue;
                }
            };
            let Some(source_sheet) = source
                .sheets
                .iter()
                .find(|sheet| sheet.name == task.source_ws)
            else {
                report.task_skipped += 1;
                report.issues.push(format!(
                    "row {}: source worksheet {:?} not found",
                    task.row_no, task.source_ws
                ));
                continue;
            };
            let source_date_values = date_iso_map(source_path, &source.sheets)?;
            let source_1904 = uses_1904_date_system(source_path)?;
            let source_cols = used_column_count(source_sheet);
            let key_cols = if task.key_cols.is_empty() {
                (1..=source_cols).collect::<Vec<_>>()
            } else {
                task.key_cols.clone()
            };
            if key_cols.contains(&0) {
                report.task_skipped += 1;
                report.issues.push(format!(
                    "row {}: key column numbers must be positive and one-based",
                    task.row_no
                ));
                continue;
            }

            if !state.contains_key(target_sheet_name) {
                let target_sheet = target_dump.as_ref().and_then(|dump| {
                    dump.sheets
                        .iter()
                        .find(|sheet| sheet.name == *target_sheet_name)
                });
                state.insert(
                    target_sheet_name.clone(),
                    AppendSheetState::new(target_sheet, &target_date_values),
                );
            }
            let sheet_state = state
                .get_mut(target_sheet_name)
                .expect("append target state was inserted");
            // V2 requires exact header text and order; trimming here could
            // silently accept a source/target schema mismatch.
            let source_headers = (1..=source_cols)
                .map(|column| cell_text(source_sheet, 1, column, &source_date_values))
                .collect::<Vec<_>>();
            if !sheet_state.is_new_sheet
                && !sheet_state.is_empty_sheet
                && source_headers != sheet_state.headers
            {
                report.task_skipped += 1;
                report.issues.push(format!(
                    "row {}: header_mismatch on target worksheet {:?}",
                    task.row_no, target_sheet_name
                ));
                continue;
            }

            if sheet_state.is_empty_sheet && !sheet_state.header_written {
                let header_row = (1..=source_cols)
                    .map(|column| append_cell(source_sheet, 1, column, source_1904, target_1904))
                    .collect::<Result<Vec<_>>>()?;
                if target_existed {
                    if sheet_state.is_new_sheet {
                        operations.push(EditOperation::CreateSheet {
                            name: target_sheet_name.clone(),
                        });
                    }
                    for (column, value) in header_row.iter().enumerate() {
                        push_append_operation(
                            &mut operations,
                            target_sheet_name,
                            1,
                            column + 1,
                            value.clone(),
                        );
                    }
                } else {
                    let target_sheet = get_or_add_new_sheet(&mut new_sheets, target_sheet_name);
                    target_sheet.rows.push(header_row);
                }
                sheet_state.header_written = true;
                sheet_state.headers = source_headers.clone();
                sheet_state.is_new_sheet = false;
                sheet_state.is_empty_sheet = false;
                sheet_state.next_row = 2;
            }

            let existing_sheet = target_dump.as_ref().and_then(|dump| {
                dump.sheets
                    .iter()
                    .find(|sheet| sheet.name == *target_sheet_name)
            });
            let mut seen_keys = sheet_state
                .keys_by_columns
                .get(&key_cols)
                .cloned()
                .unwrap_or_else(|| {
                    let mut keys = BTreeSet::new();
                    if let Some(existing_sheet) = existing_sheet {
                        for existing_row in 2..=used_row_count(existing_sheet) {
                            let key = key_cols
                                .iter()
                                .map(|column| {
                                    normalize(&cell_text(
                                        existing_sheet,
                                        existing_row,
                                        *column,
                                        &target_date_values,
                                    ))
                                })
                                .collect::<Vec<_>>();
                            if key.iter().any(|value| !value.is_empty()) {
                                keys.insert(key);
                            }
                        }
                    }
                    // Python mutates its in-memory target after every append task.
                    // Seed a new selector from rows accepted by earlier tasks too.
                    for row in &sheet_state.appended_rows {
                        let key = key_cols
                            .iter()
                            .map(|column| {
                                row.get(column.saturating_sub(1))
                                    .map(|value| normalize(value))
                                    .unwrap_or_default()
                            })
                            .collect::<Vec<_>>();
                        if key.iter().any(|value| !value.is_empty()) {
                            keys.insert(key);
                        }
                    }
                    keys
                });

            let mut newly_appended_rows = Vec::<Vec<String>>::new();
            for row in 2..=used_row_count(source_sheet) {
                let key = key_cols
                    .iter()
                    .map(|column| {
                        normalize(&cell_text(source_sheet, row, *column, &source_date_values))
                    })
                    .collect::<Vec<_>>();
                if key.iter().all(String::is_empty) {
                    // Python's append path keeps rows with blank key columns.
                } else if !seen_keys.insert(key) {
                    report.duplicate_rows_filtered += 1;
                    continue;
                }
                let values = (1..=source_cols)
                    .map(|column| append_cell(source_sheet, row, column, source_1904, target_1904))
                    .collect::<Result<Vec<_>>>()?;
                newly_appended_rows.push(
                    (1..=source_cols)
                        .map(|column| {
                            normalize(&cell_text(source_sheet, row, column, &source_date_values))
                        })
                        .collect(),
                );
                let destination_row = sheet_state.next_row.max(2);
                if target_existed {
                    for (column, value) in values.into_iter().enumerate() {
                        push_append_operation(
                            &mut operations,
                            target_sheet_name,
                            destination_row,
                            column + 1,
                            value,
                        );
                    }
                } else {
                    get_or_add_new_sheet(&mut new_sheets, target_sheet_name)
                        .rows
                        .push(values);
                }
                sheet_state.next_row = destination_row + 1;
                report.appended_rows += 1;
            }
            sheet_state
                .keys_by_columns
                .insert(key_cols.clone(), seen_keys);
            // Keep every selector cache in sync with rows appended by this task.
            // A later task may use a different key-column combination.
            for (columns, keys) in &mut sheet_state.keys_by_columns {
                for row in &newly_appended_rows {
                    let key = columns
                        .iter()
                        .map(|column| {
                            row.get(column.saturating_sub(1))
                                .map(|value| normalize(value))
                                .unwrap_or_default()
                        })
                        .collect::<Vec<_>>();
                    if key.iter().any(|value| !value.is_empty()) {
                        keys.insert(key);
                    }
                }
            }
            sheet_state.appended_rows.extend(newly_appended_rows);
            report.task_ok += 1;
        }

        if target_existed && !operations.is_empty() {
            let profile = inspect_xlsx_profile(&target)?;
            // Object-model writers may reserialize known OOXML parts. Require
            // semantic preservation of the listed workbook features rather
            // than byte-identical package-part preservation; unknown parts are
            // rejected by the selected backend's preflight.
            let preserve = BTreeSet::from([
                PreservationRequirement::CellValuesAndTypes,
                PreservationRequirement::FormulaText,
                PreservationRequirement::Styles,
                PreservationRequirement::CommentsAndPositions,
                PreservationRequirement::MergedRanges,
                PreservationRequirement::DefinedNames,
                PreservationRequirement::RowAndColumnDimensions,
                PreservationRequirement::PageSettings,
                PreservationRequirement::HeaderFooterText,
                PreservationRequirement::Relationships,
                PreservationRequirement::SheetIdentityAndOrder,
            ]);
            let plan = ExcelEditPlan {
                request_id: format!(
                    "pytools-dedup-append-{}",
                    target.file_stem().unwrap_or_default().to_string_lossy()
                ),
                input: target.clone(),
                output: candidate.clone(),
                format: WorkbookFormat::Xlsx,
                workbook: profile,
                operations,
                preserve,
                drop_allowed: Default::default(),
            };
            plan.validate().context("validate dedup append edit plan")?;
            report.edit_plan = Some(plan);
        } else if !target_existed && !new_sheets.is_empty() {
            report.new_workbook = Some(NewAppendWorkbook { sheets: new_sheets });
        }
        crate::file_commit::verify_planning_snapshot(&target, target_sha256.as_deref())
            .with_context(|| format!("verify append target snapshot {}", target.display()))?;
        reports.push(report);
    }
    Ok(reports)
}

/// Execute a serialized append payload only when it describes a new workbook.
/// Existing workbook edits are delegated to `apply-edit-plan` and capability
/// evidence through ExcelModifier.
pub fn write_new_append_candidate(plan: &AppendTargetPlan) -> Result<PathBuf> {
    anyhow::ensure!(
        !plan.target_existed,
        "existing workbook append plans must use ExcelModifier"
    );
    let workbook_plan = plan
        .new_workbook
        .as_ref()
        .context("append plan has no new workbook payload")?;
    anyhow::ensure!(
        !plan.target.exists(),
        "append target appeared after planning; regenerate the plan: {}",
        plan.target.display()
    );
    anyhow::ensure!(
        !plan.candidate.exists(),
        "refusing to overwrite {}",
        plan.candidate.display()
    );
    write_append_workbook_payload(workbook_plan, &plan.candidate)
}

/// Write a workbook payload to a fresh XLSX path and verify the reopened cell
/// count and sheet order. Used for both new append targets and replacement
/// sources staged for an ExcelEditPlan.
pub fn write_append_workbook_payload(
    workbook_plan: &NewAppendWorkbook,
    output_path: &Path,
) -> Result<PathBuf> {
    anyhow::ensure!(
        !workbook_plan
            .sheets
            .iter()
            .flat_map(|sheet| &sheet.rows)
            .flatten()
            .any(|cell| matches!(cell, AppendCell::Value(CellValue::Error(_)))),
        "rust_xlsxwriter candidate does not expose a typed error-cell writer; refusing to convert error values to text"
    );
    anyhow::ensure!(
        !output_path.exists(),
        "refusing to overwrite {}",
        output_path.display()
    );
    let header_format = excel_writer::Format::new()
        .set_bold()
        .set_border(excel_writer::FormatBorder::Thin)
        .set_align(excel_writer::FormatAlign::Center)
        .set_align(excel_writer::FormatAlign::Top);
    let basic_header_format = excel_writer::Format::new().set_bold();
    let data_date_format = excel_writer::Format::new().set_num_format("yyyy/m/d");
    let mut workbook = excel_writer::Workbook::new();
    for planned_sheet in &workbook_plan.sheets {
        if planned_sheet.style_header_row {
            anyhow::ensure!(
                planned_sheet.rows.first().is_some_and(|row| {
                    !row.is_empty()
                        && row
                            .iter()
                            .all(|cell| matches!(cell, AppendCell::Value(CellValue::Text(_))))
                }),
                "styled append header row must contain text cells"
            );
        }
        if let Some(column) = planned_sheet.data_date_column {
            anyhow::ensure!(
                planned_sheet
                    .rows
                    .first()
                    .and_then(|row| row.get(column))
                    .is_some_and(|cell| matches!(cell, AppendCell::Value(CellValue::Text(value)) if value == "数据日期")),
                "data-date format column must refer to the 数据日期 header"
            );
        }
        let worksheet = workbook.add_worksheet();
        worksheet.set_name(&planned_sheet.name)?;
        for (row_index, row) in planned_sheet.rows.iter().enumerate() {
            for (column_index, cell) in row.iter().enumerate() {
                let row_index =
                    u32::try_from(row_index).context("worksheet row exceeds XLSX limit")?;
                let column_index =
                    u16::try_from(column_index).context("worksheet column exceeds XLSX limit")?;
                let is_data_date =
                    row_index > 0 && planned_sheet.data_date_column == Some(column_index as usize);
                match cell {
                    AppendCell::Formula(formula) => {
                        if is_data_date {
                            worksheet.write_formula_with_format(
                                row_index,
                                column_index,
                                formula.as_str(),
                                &data_date_format,
                            )?;
                        } else {
                            worksheet.write_formula(row_index, column_index, formula.as_str())?;
                        }
                    }
                    AppendCell::Value(CellValue::Text(value)) => {
                        if planned_sheet.style_header_row && row_index == 0 {
                            worksheet.write_string_with_format(
                                row_index,
                                column_index,
                                value,
                                &header_format,
                            )?;
                        } else if planned_sheet.header_bold && row_index == 0 {
                            worksheet.write_string_with_format(
                                row_index,
                                column_index,
                                value,
                                &basic_header_format,
                            )?;
                        } else if is_data_date {
                            worksheet.write_string_with_format(
                                row_index,
                                column_index,
                                value,
                                &data_date_format,
                            )?;
                        } else {
                            worksheet.write_string(row_index, column_index, value)?;
                        }
                    }
                    AppendCell::Value(CellValue::Decimal(value)) => {
                        let number = value.parse::<f64>().with_context(|| {
                            format!("numeric cell {value:?} is not representable as XLSX number")
                        })?;
                        anyhow::ensure!(number.is_finite(), "numeric cell must be finite");
                        if is_data_date {
                            worksheet.write_number_with_format(
                                row_index,
                                column_index,
                                number,
                                &data_date_format,
                            )?;
                        } else {
                            worksheet.write_number(row_index, column_index, number)?;
                        }
                    }
                    AppendCell::Value(CellValue::Boolean(value)) => {
                        if is_data_date {
                            worksheet.write_boolean_with_format(
                                row_index,
                                column_index,
                                *value,
                                &data_date_format,
                            )?;
                        } else {
                            worksheet.write_boolean(row_index, column_index, *value)?;
                        }
                    }
                    AppendCell::Number {
                        serial,
                        number_format,
                    } => {
                        let number = serial.parse::<f64>().context("parse numeric serial")?;
                        anyhow::ensure!(number.is_finite(), "numeric serial must be finite");
                        let format = excel_writer::Format::new().set_num_format(if is_data_date {
                            "yyyy/m/d"
                        } else {
                            number_format
                        });
                        worksheet.write_number_with_format(
                            row_index,
                            column_index,
                            number,
                            &format,
                        )?;
                    }
                    AppendCell::DateTime {
                        serial,
                        number_format,
                    }
                    | AppendCell::Value(CellValue::DateTime {
                        serial,
                        number_format,
                    }) => {
                        let serial = serial.parse::<f64>().context("parse date/time serial")?;
                        anyhow::ensure!(serial.is_finite(), "date/time serial must be finite");
                        let format = excel_writer::Format::new().set_num_format(if is_data_date {
                            "yyyy/m/d"
                        } else {
                            number_format
                        });
                        if serial < 0.0 {
                            // Keep pre-1900 values as their stored Excel serial. Python/openpyxl
                            // preserves these numbers with a date number format, while
                            // ExcelDateTime rejects them before rust_xlsxwriter can write them.
                            worksheet.write_number_with_format(
                                row_index,
                                column_index,
                                serial,
                                &format,
                            )?;
                        } else {
                            let datetime =
                                excel_writer::ExcelDateTime::from_serial_datetime(serial)?;
                            worksheet.write_datetime_with_format(
                                row_index,
                                column_index,
                                &datetime,
                                &format,
                            )?;
                        }
                    }
                    AppendCell::Value(CellValue::Error(_)) => {
                        unreachable!("typed error values were rejected above")
                    }
                    AppendCell::Value(CellValue::Blank) => {
                        if is_data_date {
                            worksheet.write_blank(row_index, column_index, &data_date_format)?;
                        }
                    }
                }
            }
        }
    }
    if let Some(parent) = output_path.parent() {
        fs::create_dir_all(parent)?;
    }
    fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(output_path)
        .with_context(|| format!("reserve append workbook output {}", output_path.display()))?;
    let validation = (|| -> Result<()> {
        workbook
            .save(output_path)
            .with_context(|| format!("write append workbook {}", output_path.display()))?;
        let reopened = read_workbook(output_path)
            .with_context(|| format!("reopen append workbook {}", output_path.display()))?;
        anyhow::ensure!(
            reopened.sheets.len() == workbook_plan.sheets.len(),
            "append workbook worksheet count does not match its plan"
        );
        for (planned, actual) in workbook_plan.sheets.iter().zip(&reopened.sheets) {
            anyhow::ensure!(
                planned.name == actual.name,
                "append workbook worksheet order/name differs: expected {:?}, found {:?}",
                planned.name,
                actual.name
            );
            let expected = planned
                .rows
                .iter()
                .flatten()
                .filter(|cell| match cell {
                    AppendCell::Formula(_) => true,
                    AppendCell::Number { .. } => true,
                    AppendCell::DateTime { .. } => true,
                    AppendCell::Value(CellValue::Blank) => false,
                    AppendCell::Value(_) => true,
                })
                .count();
            let actual_count = actual
                .values
                .keys()
                .chain(actual.formulas.keys())
                .collect::<BTreeSet<_>>()
                .len();
            anyhow::ensure!(expected==actual_count,"append workbook non-empty cell count differs in {:?}: expected {expected}, found {actual_count}",planned.name);
        }
        Ok(())
    })();
    if let Err(error) = validation {
        let _ = fs::remove_file(output_path);
        return Err(error);
    }
    Ok(output_path.to_path_buf())
}

#[derive(Default)]
struct AppendSheetState {
    is_new_sheet: bool,
    is_empty_sheet: bool,
    header_written: bool,
    headers: Vec<String>,
    next_row: usize,
    keys_by_columns: std::collections::HashMap<Vec<usize>, BTreeSet<Vec<String>>>,
    appended_rows: Vec<Vec<String>>,
}

impl AppendSheetState {
    fn new(sheet: Option<&SheetDump>, date_values: &DateIsoMap) -> Self {
        let Some(sheet) = sheet else {
            return Self {
                is_new_sheet: true,
                is_empty_sheet: true,
                next_row: 1,
                ..Self::default()
            };
        };
        let rows = used_row_count(sheet);
        let headers: Vec<String> = (1..=used_column_count(sheet))
            .map(|column| cell_text(sheet, 1, column, date_values))
            .collect();
        let is_empty_sheet = rows <= 1 && headers.iter().all(String::is_empty);
        let next_row = rows.saturating_add(1).max(2);
        Self {
            is_new_sheet: false,
            is_empty_sheet,
            header_written: false,
            headers,
            next_row,
            keys_by_columns: std::collections::HashMap::new(),
            appended_rows: Vec::new(),
        }
    }
}

fn append_candidate_path(target: &Path, output_dir: &Path) -> PathBuf {
    let stamp = chrono::Local::now().format("%Y%m%d_%H%M%S");
    let stem = target.file_stem().unwrap_or_default().to_string_lossy();
    use std::hash::{Hash, Hasher};
    let mut hasher = std::collections::hash_map::DefaultHasher::new();
    target.hash(&mut hasher);
    let path_key = hasher.finish();
    let mut candidate = output_dir.join(format!(
        "{stem}_dedup_candidate_{stamp}_{path_key:08x}.xlsx"
    ));
    let mut suffix = 1usize;
    while candidate.exists() {
        candidate = output_dir.join(format!(
            "{stem}_dedup_candidate_{stamp}_{path_key:08x}_{suffix}.xlsx"
        ));
        suffix += 1;
    }
    candidate
}

fn is_append_source(path: &Path) -> bool {
    path.extension()
        .and_then(|ext| ext.to_str())
        .is_some_and(|ext| matches!(ext.to_ascii_lowercase().as_str(), "xlsx" | "xlsm"))
}

fn get_or_add_new_sheet<'a>(
    sheets: &'a mut Vec<NewAppendSheet>,
    name: &str,
) -> &'a mut NewAppendSheet {
    if let Some(index) = sheets.iter().position(|sheet| sheet.name == name) {
        return &mut sheets[index];
    }
    sheets.push(NewAppendSheet {
        name: name.to_owned(),
        rows: Vec::new(),
        style_header_row: false,
        header_bold: false,
        data_date_column: None,
    });
    sheets.last_mut().expect("new sheet was just pushed")
}

fn append_cell(
    sheet: &SheetDump,
    row: usize,
    column: usize,
    source_1904: bool,
    target_1904: bool,
) -> Result<AppendCell> {
    let address = cell_address(row, column);
    if let Some(formula) = sheet.formulas.get(&address) {
        return Ok(AppendCell::Formula(format!(
            "={}",
            formula.trim_start_matches('=')
        )));
    }
    let value = sheet
        .values
        .get(&address)
        .unwrap_or(&serde_json::Value::Null);
    let cell = match value.get("type").and_then(serde_json::Value::as_str) {
        Some("string" | "duration_iso") => CellValue::Text(
            value
                .get("value")
                .and_then(serde_json::Value::as_str)
                .unwrap_or_default()
                .to_owned(),
        ),
        Some("datetime_iso") => {
            let iso = value
                .get("value")
                .and_then(serde_json::Value::as_str)
                .unwrap_or_default();
            let serial = iso_datetime_to_excel_serial(iso, target_1904)
                .with_context(|| format!("unsupported ISO datetime value {iso:?}"))?;
            return Ok(AppendCell::DateTime {
                serial: serial.to_string(),
                number_format: "yyyy-mm-dd h:mm:ss".to_owned(),
            });
        }
        Some("number" | "integer") => CellValue::Decimal(
            value
                .get("value")
                .map(ToString::to_string)
                .unwrap_or_else(|| "0".into()),
        ),
        Some("datetime_serial") => {
            let serial = value
                .get("value")
                .and_then(serde_json::Value::as_f64)
                .unwrap_or(0.0);
            let adjusted = match (source_1904, target_1904) {
                (true, false) => serial + 1462.0,
                (false, true) => serial - 1462.0,
                _ => serial,
            };
            return Ok(AppendCell::DateTime {
                serial: adjusted.to_string(),
                number_format: "yyyy-mm-dd h:mm:ss".to_owned(),
            });
        }
        Some("boolean") => CellValue::Boolean(
            value
                .get("value")
                .and_then(serde_json::Value::as_bool)
                .unwrap_or(false),
        ),
        Some("error") => CellValue::Error(
            value
                .get("value")
                .and_then(serde_json::Value::as_str)
                .unwrap_or_default()
                .to_owned(),
        ),
        _ => CellValue::Blank,
    };
    Ok(AppendCell::Value(cell))
}

fn push_append_operation(
    operations: &mut Vec<EditOperation>,
    sheet: &str,
    row: usize,
    column: usize,
    value: AppendCell,
) {
    let cell = cell_address(row, column);
    match value {
        AppendCell::Formula(formula) => operations.push(EditOperation::SetFormulaText {
            sheet: sheet.to_owned(),
            cell,
            formula,
        }),
        AppendCell::Value(value) => operations.push(EditOperation::SetCellValue {
            sheet: sheet.to_owned(),
            cell,
            value,
        }),
        AppendCell::Number {
            serial,
            number_format,
        }
        | AppendCell::DateTime {
            serial,
            number_format,
        } => operations.push(EditOperation::SetCellValue {
            sheet: sheet.to_owned(),
            cell,
            value: CellValue::DateTime {
                serial,
                number_format,
            },
        }),
    }
}

fn append_duplicate_operations(
    operations: &mut Vec<EditOperation>,
    sheet: &SheetDump,
    rows: Vec<usize>,
    action: DuplicateAction,
) {
    match action {
        // The source menu's "check" action marks duplicate rows red and saves
        // the workbook. Keep that behavior while the CLI exposes the action
        // name used by the product menu.
        DuplicateAction::Check | DuplicateAction::Mark => {
            let last_col = used_column_count(sheet);
            for row in rows {
                for col in 1..=last_col {
                    operations.push(EditOperation::SetCellFillColor {
                        sheet: sheet.name.clone(),
                        cell: cell_address(row, col),
                        argb: "FFFF0000".to_owned(),
                    });
                }
            }
        }
        DuplicateAction::Delete => {
            let mut runs = Vec::new();
            let mut iter = rows.into_iter().rev().peekable();
            while let Some(end) = iter.next() {
                let mut start = end;
                while iter.peek().is_some_and(|next| *next + 1 == start) {
                    start = iter.next().expect("peeked duplicate row");
                }
                runs.push((start, end));
            }
            for (start, end) in runs {
                operations.push(EditOperation::DeleteRows {
                    sheet: sheet.name.clone(),
                    first_row: start as u32,
                    count: (end - start + 1) as u32,
                });
            }
        }
    }
}

fn duplicate_rows_for_surviving_rows(
    sheet: &SheetDump,
    key_cols: &[usize],
    surviving_original_rows: &[usize],
    date_values: &DateIsoMap,
) -> Vec<usize> {
    if key_cols.is_empty() || surviving_original_rows.len() <= 1 {
        return Vec::new();
    }
    let mut seen = BTreeSet::new();
    let mut duplicate_current_rows = Vec::new();
    for (current_index, original_row) in surviving_original_rows.iter().copied().enumerate().skip(1)
    {
        let key = key_cols
            .iter()
            .map(|column| normalize(&cell_text(sheet, original_row, *column, date_values)))
            .collect::<Vec<_>>();
        if key.iter().all(String::is_empty) {
            continue;
        }
        if !seen.insert(key) {
            duplicate_current_rows.push(current_index + 1);
        }
    }
    duplicate_current_rows
}

fn duplicate_rows_for_sheet(
    sheet: &SheetDump,
    key_cols: &[usize],
    date_values: &DateIsoMap,
) -> Vec<usize> {
    if key_cols.is_empty() {
        return Vec::new();
    }
    let last_row = used_row_count(sheet);
    let mut seen = BTreeSet::new();
    let mut duplicate = Vec::new();
    for row in 2..=last_row {
        let key = key_cols
            .iter()
            .map(|col| normalize(&cell_text(sheet, row, *col, date_values)))
            .collect::<Vec<_>>();
        if key.iter().all(String::is_empty) {
            continue;
        }
        if !seen.insert(key) {
            duplicate.push(row);
        }
    }
    duplicate
}

fn comment_key_columns(sheet: &SheetDump) -> Vec<usize> {
    let marked = sheet
        .comments
        .iter()
        .filter_map(|(address, text)| {
            (parse_a1(address).is_some_and(|(row, _)| row == 1)
                && normalize(text).contains("标识列"))
            .then(|| parse_a1(address).map(|(_, col)| col))
            .flatten()
        })
        .collect::<Vec<_>>();
    if marked.is_empty() {
        (1..=used_column_count(sheet)).collect()
    } else {
        marked
    }
}

fn cell_text(sheet: &SheetDump, row: usize, col: usize, date_values: &DateIsoMap) -> String {
    let address = cell_address(row, col);
    if let Some(formula) = sheet.formulas.get(&address) {
        return format!("={formula}");
    }
    if let Some(value) = date_values.get(&(sheet.name.clone(), address.clone())) {
        return value.get(..10).unwrap_or(value).to_owned();
    }
    let Some(value) = sheet.values.get(&address) else {
        return String::new();
    };
    match value.get("type").and_then(serde_json::Value::as_str) {
        Some("string" | "datetime_iso" | "duration_iso" | "error") => value
            .get("value")
            .and_then(serde_json::Value::as_str)
            .unwrap_or_default()
            .to_owned(),
        Some("number" | "integer" | "datetime_serial") => value
            .get("value")
            .map(ToString::to_string)
            .unwrap_or_default(),
        Some("boolean") => value
            .get("value")
            .and_then(serde_json::Value::as_bool)
            .map(|value| if value { "True" } else { "False" }.to_owned())
            .unwrap_or_default(),
        _ => String::new(),
    }
}

fn normalize(value: &str) -> String {
    let text = value.trim();
    normalize_python_date_text(text).unwrap_or_else(|| text.to_owned())
}

fn normalize_python_date_text(text: &str) -> Option<String> {
    let chars = text.chars().collect::<Vec<_>>();
    if chars.len() < 8 || chars[0] != '2' || chars[1] != '0' {
        return None;
    }
    let year = chars
        .get(..4)?
        .iter()
        .collect::<String>()
        .parse::<i32>()
        .ok()?;
    let mut cursor = 4;
    if !matches!(*chars.get(cursor)?, '年' | '-' | '.' | '/') {
        return None;
    }
    cursor += 1;
    let month_start = cursor;
    while chars.get(cursor).is_some_and(|ch| ch.is_ascii_digit()) && cursor - month_start < 2 {
        cursor += 1;
    }
    if cursor == month_start || !matches!(*chars.get(cursor)?, '月' | '-' | '.' | '/') {
        return None;
    }
    let month = chars[month_start..cursor]
        .iter()
        .collect::<String>()
        .parse::<u32>()
        .ok()?;
    cursor += 1;
    let day_start = cursor;
    while chars.get(cursor).is_some_and(|ch| ch.is_ascii_digit()) && cursor - day_start < 2 {
        cursor += 1;
    }
    if cursor == day_start {
        return None;
    }
    let day = chars[day_start..cursor]
        .iter()
        .collect::<String>()
        .parse::<u32>()
        .ok()?;
    if chars.get(cursor) == Some(&'日') {
        cursor += 1;
    }
    if cursor < chars.len() {
        if !matches!(chars[cursor], ' ' | 'T') {
            return None;
        }
        let time = chars[cursor + 1..].iter().collect::<String>();
        if !valid_python_time_suffix(&time) {
            return None;
        }
    }
    let date = chrono::NaiveDate::from_ymd_opt(year, month, day)?;
    Some(date.format("%Y-%m-%d").to_string())
}

fn valid_python_time_suffix(time: &str) -> bool {
    let Some((hour, rest)) = time.split_once(':') else {
        return false;
    };
    if !(1..=2).contains(&hour.len()) || !hour.bytes().all(|byte| byte.is_ascii_digit()) {
        return false;
    }
    let (minute, seconds) = match rest.split_once(':') {
        Some((minute, seconds)) => (minute, Some(seconds)),
        None => (rest, None),
    };
    if minute.len() != 2 || !minute.bytes().all(|byte| byte.is_ascii_digit()) {
        return false;
    }
    let Some(seconds) = seconds else {
        return true;
    };
    let whole_seconds = match seconds.split_once('.') {
        Some((whole, fraction))
            if !fraction.is_empty() && fraction.bytes().all(|byte| byte.is_ascii_digit()) =>
        {
            whole
        }
        Some(_) => return false,
        None => seconds,
    };
    whole_seconds.len() == 2 && whole_seconds.bytes().all(|byte| byte.is_ascii_digit())
}

fn date_iso_map(path: &Path, sheets: &[SheetDump]) -> Result<DateIsoMap> {
    let has_date_serial = sheets.iter().any(|sheet| {
        sheet.values.values().any(|value| {
            value.get("type").and_then(serde_json::Value::as_str) == Some("datetime_serial")
        })
    });
    if !has_date_serial {
        return Ok(DateIsoMap::default());
    }
    let max_columns = sheets
        .iter()
        .map(used_column_count)
        .max()
        .unwrap_or(1)
        .max(1);
    let mut dates = DateIsoMap::default();
    for sheet in read_value_columns_with_iso_dates(path, max_columns)? {
        for (address, value) in sheet.values {
            if let ReaderCellValue::DateTimeIso(value) = value {
                dates.insert((sheet.name.clone(), address), value);
            }
        }
    }
    Ok(dates)
}

fn iso_datetime_to_excel_serial(value: &str, date_1904: bool) -> Option<f64> {
    use chrono::Timelike;
    let datetime = chrono::NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S%.f")
        .ok()
        .or_else(|| chrono::NaiveDateTime::parse_from_str(value, "%Y-%m-%d %H:%M:%S%.f").ok())
        .or_else(|| {
            chrono::NaiveDate::parse_from_str(value, "%Y-%m-%d")
                .ok()
                .and_then(|date| date.and_hms_opt(0, 0, 0))
        })?;
    let epoch = if date_1904 {
        chrono::NaiveDate::from_ymd_opt(1904, 1, 1)?
    } else {
        chrono::NaiveDate::from_ymd_opt(1899, 12, 30)?
    };
    let days = datetime.date().signed_duration_since(epoch).num_days();
    let seconds = datetime.time().num_seconds_from_midnight() as f64;
    let fraction = datetime.time().nanosecond() as f64 / 86_400_000_000_000.0;
    Some(days as f64 + seconds / 86_400.0 + fraction)
}

fn make_plan(
    input: &Path,
    output: &Path,
    profile: excel_editor::WorkbookProfile,
    operations: Vec<EditOperation>,
    request: &str,
) -> Result<Option<ExcelEditPlan>> {
    if operations.is_empty() {
        return Ok(None);
    }
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
    let plan = ExcelEditPlan {
        request_id: format!("pytools-{request}"),
        input: input.to_path_buf(),
        output: output.to_path_buf(),
        format: WorkbookFormat::Xlsx,
        workbook: profile,
        operations,
        preserve,
        drop_allowed: Default::default(),
    };
    plan.validate().context("validate duplicate edit plan")?;
    Ok(Some(plan))
}

fn ensure_xlsx(input: &Path, output: &Path) -> Result<()> {
    if !input.is_file() {
        bail!("input workbook does not exist: {}", input.display());
    }
    if !is_xlsx(input) || !is_xlsx(output) {
        bail!("duplicate edit plans currently accept .xlsx only");
    }
    if input.canonicalize().ok() == output.canonicalize().ok() && output.exists() {
        bail!("edit plans cannot overwrite the original workbook");
    }
    Ok(())
}
fn is_xlsx(path: &Path) -> bool {
    path.extension()
        .and_then(|value| value.to_str())
        .is_some_and(|value| value.eq_ignore_ascii_case("xlsx"))
}
fn used_row_count(sheet: &SheetDump) -> usize {
    sheet
        .values
        .keys()
        .chain(sheet.formulas.keys())
        .filter_map(|address| parse_a1(address).map(|x| x.0))
        .max()
        .unwrap_or(0)
}
fn used_column_count(sheet: &SheetDump) -> usize {
    sheet
        .values
        .keys()
        .chain(sheet.formulas.keys())
        .filter_map(|address| parse_a1(address).map(|x| x.1))
        .max()
        .unwrap_or(1)
}
fn parse_a1(value: &str) -> Option<(usize, usize)> {
    let split = value.find(|c: char| c.is_ascii_digit())?;
    let col = value[..split].bytes().try_fold(0usize, |acc, byte| {
        acc.checked_mul(26)?
            .checked_add((byte.to_ascii_uppercase() - b'A' + 1) as usize)
    })?;
    Some((value[split..].parse().ok()?, col))
}
fn cell_address(row: usize, mut col: usize) -> String {
    let mut letters = String::new();
    while col > 0 {
        letters.insert(0, (b'A' + ((col - 1) % 26) as u8) as char);
        col = (col - 1) / 26;
    }
    format!("{letters}{row}")
}

#[cfg(test)]
mod tests {
    use super::{iso_datetime_to_excel_serial, normalize, precheck_dedup_tasks};
    use crate::config::DedupTask;
    use std::{
        fs,
        path::PathBuf,
        time::{SystemTime, UNIX_EPOCH},
    };

    #[test]
    fn date_keys_match_python_date_and_datetime_normalization() {
        for input in [
            "2026-1-31",
            "2026/01/31",
            "2026年1月31日",
            "2026-01-31 9:02",
            "2026-01-31T12:30:45.25",
        ] {
            assert_eq!(normalize(input), "2026-01-31", "{input}");
        }
        assert_eq!(normalize("2026-02-31"), "2026-02-31");
        assert_eq!(normalize("2026-01-31Tnot-a-time"), "2026-01-31Tnot-a-time");
    }

    #[test]
    fn excel_date_epochs_differ_by_1462_days() {
        let date = "2026-01-31T12:30:45.000";
        let serial_1900 = iso_datetime_to_excel_serial(date, false).unwrap();
        let serial_1904 = iso_datetime_to_excel_serial(date, true).unwrap();
        assert!((serial_1900 - serial_1904 - 1462.0).abs() < 1e-12);
    }

    #[test]
    fn dedup_precheck_validates_enabled_rows_sharing_one_source_workbook() -> anyhow::Result<()> {
        let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let root = std::env::temp_dir().join(format!(
            "pytools-dedup-precheck-{}-{nonce}",
            std::process::id()
        ));
        fs::create_dir_all(&root)?;
        let _cleanup = TestDirectory(root.clone());
        let source = root.join("source.xlsx");
        let mut workbook = rust_xlsxwriter::Workbook::new();
        workbook.add_worksheet().set_name("Data")?;
        workbook.save(&source)?;

        let task = |row_no, enabled, source_ws: &str, key_cols| DedupTask {
            row_no,
            enabled,
            task_name: format!("task-{row_no}"),
            source_wb: Some(source.clone()),
            source_ws: source_ws.to_owned(),
            key_cols,
            target_wb: None,
            target_ws: String::new(),
            write_mode: String::new(),
        };
        let tasks = [
            task(2, true, "Data", vec![1]),
            task(3, true, "Data", Vec::new()),
            task(4, false, "Missing", vec![1]),
        ];

        let report = precheck_dedup_tasks(&tasks);
        assert_eq!(report.ok, 2);
        assert_eq!(report.warn, 1);
        assert_eq!(report.fail, 0);
        assert!(report.issues.iter().any(|issue| issue.contains("row 3")));
        assert!(!report.issues.iter().any(|issue| issue.contains("row 4")));
        Ok(())
    }

    struct TestDirectory(PathBuf);

    impl Drop for TestDirectory {
        fn drop(&mut self) {
            let _ = fs::remove_dir_all(&self.0);
        }
    }
}

/// 2-2 按配置删除重复：在源工作簿上就地删除重复行（保留首次出现）。
/// 语义对齐旧 PyTools `dedup_311.py::_delete_duplicates`：按标识列复合键判断，
/// 标识列全空的行保留；删除后保存回源文件。
#[derive(Debug, Default)]
pub struct DeleteDuplicatesRun {
    pub task_ok: usize,
    pub task_skipped: usize,
    pub task_failed: usize,
    pub deleted_rows: usize,
    pub saved_books: Vec<PathBuf>,
    pub issues: Vec<String>,
    pub events: Vec<crate::engine_diag::DiagEvent>,
}

pub fn delete_duplicates_in_place(tasks: &[DedupTask]) -> Result<DeleteDuplicatesRun> {
    use umya_spreadsheet as umya;
    let mut run = DeleteDuplicatesRun::default();
    let enabled: Vec<&DedupTask> = tasks.iter().filter(|task| task.enabled).collect();
    anyhow::ensure!(
        !enabled.is_empty(),
        "未找到启用的去重任务，请检查“去重追加数据配置”工作表和“是否启用”列"
    );
    let mut grouped = std::collections::BTreeMap::<PathBuf, Vec<&DedupTask>>::new();
    for task in enabled {
        let Some(source) = task.source_wb.as_ref() else {
            run.task_skipped += 1;
            run.issues.push(format!("第 {} 行：源工作簿为空", task.row_no));
            continue;
        };
        grouped.entry(source.clone()).or_default().push(task);
    }
    for (path, group) in grouped {
        let read_started = std::time::Instant::now();
        let extension = path
            .extension()
            .and_then(|ext| ext.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase();
        if extension != "xlsx" {
            run.task_skipped += group.len();
            let message = if extension == "xlsm" {
                "xlsm 就地删除暂不支持，已跳过".to_owned()
            } else {
                format!("不支持的源扩展名 .{extension}")
            };
            run.events.push(
                crate::engine_diag::DiagEvent::new(
                    "umya-spreadsheet",
                    crate::engine_diag::DiagStage::ModifyWorkbook,
                    crate::engine_diag::DiagOutcome::Skipped,
                )
                .with_file(path.to_string_lossy().into_owned())
                .with_message(message.clone()),
            );
            run.issues.push(format!("{}：{message}", path.display()));
            continue;
        }
        let mut book = match umya::reader::xlsx::read(&path) {
            Ok(book) => book,
            Err(error) => {
                run.task_failed += group.len();
                run.events.push(
                    crate::engine_diag::DiagEvent::new(
                        "umya-spreadsheet",
                        crate::engine_diag::DiagStage::ReadSource,
                        crate::engine_diag::DiagOutcome::Failed,
                    )
                    .with_file(path.to_string_lossy().into_owned())
                    .with_message(format!("{error:#}")),
                );
                run.issues
                    .push(format!("{}：umya 读取源工作簿失败：{error:#}", path.display()));
                continue;
            }
        };
        run.events.push(crate::excel_core_diag::read_event(
            &path,
            group.first().map(|task| task.row_no),
            read_started.elapsed().as_millis(),
            None,
        ));
        let mut deleted_in_book = 0usize;
        let mut changed = false;
        for task in group {
            let task_started = std::time::Instant::now();
            let sheet = match book.get_sheet_by_name_mut(&task.source_ws) {
                Ok(sheet) => sheet,
                Err(_) => {
                    run.task_failed += 1;
                    run.events.push(
                        crate::engine_diag::DiagEvent::new(
                            "umya-spreadsheet",
                            crate::engine_diag::DiagStage::CopySheet,
                            crate::engine_diag::DiagOutcome::Failed,
                        )
                        .with_row(task.row_no)
                        .with_file(path.to_string_lossy().into_owned())
                        .with_sheet(task.source_ws.clone())
                        .with_message("源工作簿中不存在该工作表"),
                    );
                    run.issues.push(format!(
                        "第 {} 行：源工作簿 {} 中不存在工作表 {:?}",
                        task.row_no,
                        path.display(),
                        task.source_ws
                    ));
                    continue;
                }
            };
            let (max_col, max_row) = sheet.get_highest_column_and_row();
            let used_cols = max_col.max(1);
            let key_cols: Vec<u32> = if task.key_cols.is_empty() {
                (1..=used_cols).collect()
            } else {
                task.key_cols.iter().map(|column| *column as u32).collect()
            };
            let mut seen: std::collections::HashSet<Vec<String>> = std::collections::HashSet::new();
            let mut to_delete: Vec<u32> = Vec::new();
            for row in 2..=(max_row as u32) {
                let key: Vec<String> = key_cols
                    .iter()
                    .map(|column| umya_cell_text(sheet, *column, row))
                    .collect();
                if key.iter().all(|part| part.is_empty()) {
                    continue;
                }
                if seen.contains(&key) {
                    to_delete.push(row);
                } else {
                    seen.insert(key);
                }
            }
            // 自下而上删除，避免行号位移。
            for row in to_delete.iter().rev() {
                sheet.remove_row(*row, 1);
            }
            let deleted = to_delete.len();
            deleted_in_book += deleted;
            run.deleted_rows += deleted;
            run.task_ok += 1;
            if deleted > 0 {
                changed = true;
            }
            run.events.push(
                crate::engine_diag::DiagEvent::new(
                    "umya-spreadsheet",
                    crate::engine_diag::DiagStage::ModifyWorkbook,
                    crate::engine_diag::DiagOutcome::Ok,
                )
                .with_row(task.row_no)
                .with_file(path.to_string_lossy().into_owned())
                .with_sheet(task.source_ws.clone())
                .with_elapsed_ms(task_started.elapsed().as_millis())
                .with_message(format!("删除重复行 {deleted} 行")),
            );
        }
        if changed {
            umya::writer::xlsx::write(&book, &path)
                .with_context(|| format!("就地保存删除结果 {} 失败", path.display()))?;
            run.saved_books.push(path.clone());
            run.events.push(
                crate::engine_diag::DiagEvent::new(
                    "umya-spreadsheet",
                    crate::engine_diag::DiagStage::SaveOutput,
                    crate::engine_diag::DiagOutcome::Ok,
                )
                .with_file(path.to_string_lossy().into_owned())
                .with_message(format!("已就地删除 {deleted_in_book} 行")),
            );
        }
    }
    Ok(run)
}

/// 2-1 按配置检查重复：在源工作簿上就地标记重复行（整行红色填充）。
/// 语义对齐旧 PyTools `dedup_311.py::_mark_duplicates`：首次出现不标记，
/// 后续重复行从第 1 列标到已用列；标识列全部为空的行跳过；标记后保存回源文件。
/// 仅支持 .xlsx 源（.xlsm 就地保存涉及宏保留，暂不支持并如实跳过）。
#[derive(Debug, Default)]
pub struct MarkDuplicatesRun {
    pub task_ok: usize,
    pub task_skipped: usize,
    pub task_failed: usize,
    pub marked_rows: usize,
    pub saved_books: Vec<PathBuf>,
    pub issues: Vec<String>,
    pub events: Vec<crate::engine_diag::DiagEvent>,
}

fn umya_cell_text(sheet: &mut umya_spreadsheet::Worksheet, column: u32, row: u32) -> String {
    let mut letters = String::new();
    let mut rest = column;
    while rest > 0 {
        letters.insert(0, (b'A' + ((rest - 1) % 26) as u8) as char);
        rest = (rest - 1) / 26;
    }
    let address = format!("{letters}{row}");
    sheet
        .get_cell(address.as_str())
        .map(|cell| cell.get_raw_value().to_string())
        .unwrap_or_default()
        .trim()
        .to_owned()
}

fn umya_mark_row_red(sheet: &mut umya_spreadsheet::Worksheet, used_cols: u32, row: u32) {
    for column in 1..=used_cols {
        let cell = sheet.cell_mut((column, row));
        let mut style = cell.get_style().clone();
        {
            let pattern = style.get_fill_mut().get_pattern_fill_mut();
            pattern.set_pattern_type(umya_spreadsheet::PatternValues::Solid);
            pattern.get_foreground_color_mut().set_argb_str("FFFF0000");
        }
        cell.set_style(style);
    }
}

pub fn mark_duplicates_in_place(tasks: &[DedupTask]) -> Result<MarkDuplicatesRun> {
    use umya_spreadsheet as umya;
    let mut run = MarkDuplicatesRun::default();
    let enabled: Vec<&DedupTask> = tasks.iter().filter(|task| task.enabled).collect();
    anyhow::ensure!(
        !enabled.is_empty(),
        "未找到启用的去重任务，请检查“去重追加数据配置”工作表和“是否启用”列"
    );    let mut grouped = std::collections::BTreeMap::<PathBuf, Vec<&DedupTask>>::new();
    for task in enabled {
        let Some(source) = task.source_wb.as_ref() else {
            run.task_skipped += 1;
            run.events.push(
                crate::engine_diag::DiagEvent::new(
                    "umya-spreadsheet",
                    crate::engine_diag::DiagStage::ReadSource,
                    crate::engine_diag::DiagOutcome::Skipped,
                )
                .with_row(task.row_no)
                .with_message("配置的源工作簿路径为空"),
            );
            run.issues.push(format!("第 {} 行：源工作簿为空", task.row_no));
            continue;
        };
        grouped.entry(source.clone()).or_default().push(task);
    }
    for (path, group) in grouped {
        let book_read_started = std::time::Instant::now();
        let extension = path
            .extension()
            .and_then(|ext| ext.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase();
        if extension != "xlsx" {
            run.task_skipped += group.len();
            let message = if extension == "xlsm" {
                "xlsm 就地标记暂不支持，已跳过".to_owned()
            } else {
                format!("不支持的源扩展名 .{extension}")
            };
            run.events.push(
                crate::engine_diag::DiagEvent::new(
                    "umya-spreadsheet",
                    crate::engine_diag::DiagStage::ModifyWorkbook,
                    crate::engine_diag::DiagOutcome::Skipped,
                )
                .with_file(path.to_string_lossy().into_owned())
                .with_message(message.clone()),
            );
            run.issues.push(format!("{}：{message}", path.display()));
            continue;
        }
        let mut book = match umya::reader::xlsx::read(&path) {
            Ok(book) => book,
            Err(error) => {
                run.task_failed += group.len();
                run.events.push(
                    crate::engine_diag::DiagEvent::new(
                        "umya-spreadsheet",
                        crate::engine_diag::DiagStage::ReadSource,
                        crate::engine_diag::DiagOutcome::Failed,
                    )
                    .with_file(path.to_string_lossy().into_owned())
                    .with_message(format!("{error:#}")),
                );
                run.issues
                    .push(format!("{}：umya 读取源工作簿失败：{error:#}", path.display()));
                continue;
            }
        };
        // 读取成功事件（经统一适配层构造，引擎名与阶段口径统一）。
        run.events.push(crate::excel_core_diag::read_event(
            &path,
            group.first().map(|task| task.row_no),
            book_read_started.elapsed().as_millis(),
            None,
        ));
        let mut marked_in_book = 0usize;
        let mut changed = false;
        for task in group {
            let task_started = std::time::Instant::now();
            let sheet = match book.get_sheet_by_name_mut(&task.source_ws) {
                Ok(sheet) => sheet,
                Err(_) => {
                    run.task_failed += 1;
                    run.events.push(
                        crate::engine_diag::DiagEvent::new(
                            "umya-spreadsheet",
                            crate::engine_diag::DiagStage::CopySheet,
                            crate::engine_diag::DiagOutcome::Failed,
                        )
                        .with_row(task.row_no)
                        .with_file(path.to_string_lossy().into_owned())
                        .with_sheet(task.source_ws.clone())
                        .with_message("源工作表中不存在该工作表"),
                    );
                    run.issues.push(format!(
                        "第 {} 行：源工作簿 {} 中不存在工作表 {:?}",
                        task.row_no,
                        path.display(),
                        task.source_ws
                    ));
                    continue;
                }
            };
            let (max_col, max_row) = sheet.get_highest_column_and_row();
            let used_cols = max_col.max(1);
            let key_cols: Vec<u32> = if task.key_cols.is_empty() {
                (1..=used_cols).collect()
            } else {
                task.key_cols.iter().map(|column| *column as u32).collect()
            };
            let mut seen: std::collections::HashSet<Vec<String>> = std::collections::HashSet::new();
            let mut marked = 0usize;
            for row in 2..=(max_row as u32) {
                let key: Vec<String> = key_cols
                    .iter()
                    .map(|column| umya_cell_text(sheet, *column, row))
                    .collect();
                if key.iter().all(|part| part.is_empty()) {
                    continue;
                }
                if seen.contains(&key) {
                    umya_mark_row_red(sheet, used_cols, row);
                    marked += 1;
                } else {
                    seen.insert(key);
                }
            }
            marked_in_book += marked;
            run.marked_rows += marked;
            run.task_ok += 1;
            if marked > 0 {
                changed = true;
            }
            run.events.push(
                crate::engine_diag::DiagEvent::new(
                    "umya-spreadsheet",
                    crate::engine_diag::DiagStage::ModifyWorkbook,
                    crate::engine_diag::DiagOutcome::Ok,
                )
                .with_row(task.row_no)
                .with_file(path.to_string_lossy().into_owned())
                .with_sheet(task.source_ws.clone())
                .with_elapsed_ms(task_started.elapsed().as_millis())
                .with_message(format!("标记重复行 {marked} 行")),
            );
        }
        if changed {
            umya::writer::xlsx::write(&book, &path)
                .with_context(|| format!("就地保存标记结果 {} 失败", path.display()))?;
            run.saved_books.push(path.clone());
            run.events.push(
                crate::engine_diag::DiagEvent::new(
                    "umya-spreadsheet",
                    crate::engine_diag::DiagStage::SaveOutput,
                    crate::engine_diag::DiagOutcome::Ok,
                )
                .with_file(path.to_string_lossy().into_owned())
                .with_message(format!("已就地标记 {marked_in_book} 行")),
            );
        }
    }
    Ok(run)
}
