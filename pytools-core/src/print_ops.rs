//! Print preparation plans for the non-PDF print tools.

use crate::config::PrintTask;
use anyhow::{Context, Result};
use excel_editor::{
    inspect_xlsx_profile, EditOperation, ExcelEditPlan, PageOrientation, PageSettings,
    PreservationRequirement, RangeStyleMode, WorkbookFormat,
};
use excel_reader::{read_workbook, uses_1904_date_system, SheetDump, WorkbookDump};
use rust_xlsxwriter::{Format, Workbook, Worksheet};
use serde::Serialize;
use std::{
    collections::{BTreeMap, BTreeSet, HashMap},
    fs::{self, OpenOptions},
    io,
    path::{Path, PathBuf},
    sync::atomic::{AtomicU64, Ordering},
};

#[derive(Clone, Debug, Default, Serialize)]
pub struct PrintPrecheck {
    pub ok: usize,
    pub warnings: usize,
    pub failed: usize,
    pub issues: Vec<String>,
}

#[derive(Clone, Debug, Serialize)]
pub struct PrintPlanReport {
    pub input: PathBuf,
    pub output: PathBuf,
    pub matching_sheets: usize,
    pub marked_ranges: usize,
    pub plan: Option<ExcelEditPlan>,
}

#[derive(Clone, Debug, Serialize)]
pub struct StaticPrintExecution {
    pub engine: &'static str,
    pub output: PathBuf,
    pub worksheet_count: usize,
    pub imported_ranges: usize,
    pub formulas_materialized: usize,
    pub comments_dropped: bool,
}

#[derive(Clone, Debug)]
struct StaticPrintCell {
    row: usize,
    column: usize,
    value: serde_json::Value,
}

#[derive(Clone, Debug)]
struct StaticPrintMerge {
    first_row: usize,
    first_column: usize,
    last_row: usize,
    last_column: usize,
    text: String,
}

static NEXT_STATIC_PRINT_TEMP_ID: AtomicU64 = AtomicU64::new(0);

/// Check mode 1-3 config tasks. This menu's precheck (3-4) never renders pages
/// and never exports a PDF.
pub fn precheck_print_tasks(tasks: &[PrintTask]) -> PrintPrecheck {
    let mut report = PrintPrecheck::default();
    for task in tasks.iter().filter(|task| task.enabled) {
        let mut errors = Vec::new();
        if !matches!(task.mode, 1 | 2 | 3) {
            errors.push(format!("invalid print mode {}", task.mode));
        }
        let Some(source) = task.source_wb.as_ref() else {
            errors.push("source workbook is empty".into());
            report.failed += 1;
            report.issues.extend(
                errors
                    .into_iter()
                    .map(|e| format!("row {}: {e}", task.row_no)),
            );
            continue;
        };
        if !source.is_file() {
            errors.push(format!(
                "source workbook does not exist: {}",
                source.display()
            ));
        }
        if !is_xlsx(source) {
            errors.push("Rust print edit candidate currently accepts .xlsx only".into());
        }
        if task.source_ws.trim().is_empty() {
            errors.push("source worksheet is empty".into());
        }
        if task
            .target_wb
            .as_ref()
            .is_none_or(|path| path.as_os_str().is_empty())
        {
            errors.push("target workbook is empty".into());
        }
        if task.target_wb.as_ref().is_some_and(|path| !is_xlsx(path)) {
            errors.push("target workbook must be .xlsx for the current candidate".into());
        }
        if let Some(path) = task.target_wb.as_ref() {
            match path.symlink_metadata() {
                Ok(_) => errors.push(format!(
                    "target workbook already exists; PT-3-3 only writes to a new file: {}",
                    path.display()
                )),
                Err(error) if error.kind() != std::io::ErrorKind::NotFound => errors.push(format!(
                    "cannot check target workbook path {}: {error}",
                    path.display()
                )),
                Err(_) => {}
            }
        }
        if !task.source_ranges.trim().is_empty() && parse_ranges(&task.source_ranges).is_err() {
            errors.push("source range syntax is invalid".into());
        }
        if task.fit_wide <= 0 || task.fit_tall < 0 {
            report.warnings += 1;
            report.issues.push(format!(
                "row {}: non-positive fit setting will use a safe default",
                task.row_no
            ));
        }
        if !task.orientation.trim().is_empty()
            && orientation_from_comment(&normalize(&task.orientation)).is_none()
        {
            report.warnings += 1;
            report.issues.push(format!(
                "row {}: invalid orientation will use automatic orientation",
                task.row_no
            ));
        }
        if !errors.is_empty() {
            report.failed += 1;
            report.issues.extend(
                errors
                    .into_iter()
                    .map(|e| format!("row {}: {e}", task.row_no)),
            );
            continue;
        }
        match read_workbook(source) {
            Ok(workbook)
                if workbook
                    .sheets
                    .iter()
                    .any(|sheet| sheet.name == task.source_ws) =>
            {
                report.ok += 1
            }
            Ok(_) => {
                report.failed += 1;
                report.issues.push(format!(
                    "row {}: source worksheet {:?} not found",
                    task.row_no, task.source_ws
                ));
            }
            Err(error) => {
                report.failed += 1;
                report
                    .issues
                    .push(format!("row {}: {error:#}", task.row_no));
            }
        }
    }
    report
}

/// Build the comment-marked 3-1/3-2 page-setup plan. The candidate keeps the
/// source cell addresses and constrains the print area to the marked ranges;
/// this avoids copying cells into a new style table before that operation is
/// separately parity-qualified.
pub fn build_comment_print_plan(
    input: &Path,
    output: &Path,
    drop_comments: bool,
    hide_zero: bool,
) -> Result<PrintPlanReport> {
    anyhow::ensure!(
        input.is_file(),
        "input workbook does not exist: {}",
        input.display()
    );
    anyhow::ensure!(
        is_xlsx(input) && is_xlsx(output),
        "comment print plans currently accept .xlsx only"
    );
    let dump = read_workbook(input).with_context(|| format!("read {}", input.display()))?;
    let mut marked = Vec::<(String, String, Vec<(usize, usize, usize, usize)>)>::new();
    for sheet in &dump.sheets {
        let comment = sheet
            .comments
            .get("A1")
            .map(|text| normalize(text))
            .unwrap_or_default();
        if !comment.contains("打印") && !comment.contains("输出") {
            continue;
        }
        let mut ranges = comment_ranges(sheet);
        if ranges.is_empty() {
            if let Some(bounds) = used_bounds(sheet) {
                ranges.push(bounds);
            }
        }
        if ranges.is_empty() {
            continue;
        }
        marked.push((sheet.name.clone(), comment, ranges));
    }
    let profile = inspect_xlsx_profile(input)?;
    let mut operations = Vec::new();
    let mut output_sheets = Vec::<(
        String,
        String,
        (usize, usize, usize, usize),
        PageOrientation,
    )>::new();
    let mut used_names = dump
        .sheets
        .iter()
        .map(|sheet| sheet.name.clone())
        .collect::<BTreeSet<_>>();
    let mut final_names = BTreeSet::new();
    for (source_sheet, comment, ranges) in &marked {
        for (index, range) in ranges.iter().copied().enumerate() {
            let proposed = if ranges.len() == 1 {
                source_sheet.clone()
            } else {
                format!("{}_{}", source_sheet, index + 1)
            };
            let final_name = unique_sheet_name(&proposed, &mut final_names);
            let (r1, c1, r2, c2) = range;
            let rows = r2 - r1 + 1;
            let columns = c2 - c1 + 1;
            let orientation = orientation_from_comment(comment).unwrap_or_else(|| {
                if columns > rows {
                    PageOrientation::Landscape
                } else {
                    PageOrientation::Portrait
                }
            });
            output_sheets.push((source_sheet.clone(), final_name, range, orientation));
        }
    }
    if !output_sheets.is_empty() {
        used_names.extend(output_sheets.iter().map(|(_, name, _, _)| name.clone()));
        for (index, old_name) in dump
            .sheets
            .iter()
            .map(|sheet| sheet.name.clone())
            .enumerate()
        {
            let mut temp = format!("__pt_print_seed_{}", index + 1);
            let mut suffix = 1;
            while used_names.contains(&temp) {
                temp = format!("__pt_print_seed_{}_{}", index + 1, suffix);
                suffix += 1;
            }
            used_names.insert(temp.clone());
            operations.push(EditOperation::RenameSheet {
                old_name,
                new_name: temp,
            });
        }
        let style_mode = if drop_comments {
            RangeStyleMode::Basic
        } else {
            RangeStyleMode::Full
        };
        for (source_sheet, target_sheet, (r1, c1, r2, c2), orientation) in &output_sheets {
            let source_range = format!(
                "{}{}:{}{}",
                column_letters(*c1),
                r1,
                column_letters(*c2),
                r2
            );
            operations.push(EditOperation::CreateSheet {
                name: target_sheet.clone(),
            });
            if !drop_comments {
                // The legacy openpyxl printer reads `column_dimensions[key].width`
                // for every column in a copied block. Accessing a missing
                // dimension materializes openpyxl's default width (13.0), so
                // seed that width before ImportRange overwrites explicit widths.
                for offset in 0..(c2 - c1 + 1) {
                    operations.push(EditOperation::SetColumnWidth {
                        sheet: target_sheet.clone(),
                        column_index: (offset + 1) as u32,
                        width: 13.0,
                    });
                }
            }
            operations.push(EditOperation::ImportRange {
                source_workbook: input.to_path_buf(),
                source_sheet: source_sheet.clone(),
                source_range,
                target_sheet: target_sheet.clone(),
                target_start: "A1".into(),
                copy_values: true,
                copy_formulas: false,
                style_mode,
                copy_comments: !drop_comments,
                copy_merges: true,
                copy_row_heights: !drop_comments,
                copy_column_widths: !drop_comments,
                copy_row_hidden: false,
                copy_column_hidden: false,
                copy_outline_levels: false,
                copy_all_row_dimensions: false,
                copy_all_column_dimensions: false,
                materialize_blank_cells: false,
                hide_zero,
            });
            let area = format!("A1:{}{}", column_letters(c2 - c1 + 1), r2 - r1 + 1);
            operations.push(EditOperation::SetPageSettings {
                sheet: target_sheet.clone(),
                settings: PageSettings {
                    orientation: Some(*orientation),
                    fit_to_page: Some(true),
                    fit_to_width: Some(1),
                    fit_to_height: Some(0),
                    clear_scale: true,
                    paper_size: Some(9),
                    print_area: Some(area),
                    horizontal_centered: Some(false),
                    vertical_centered: Some(false),
                    footer: None,
                },
            });
        }
        for (index, (_, target_sheet, _, _)) in output_sheets.iter().enumerate() {
            operations.push(EditOperation::SetPageSettings {
                sheet: target_sheet.clone(),
                settings: PageSettings {
                    footer: Some(format!(
                        "第 {} 页 / 共 {} 页",
                        index + 1,
                        output_sheets.len()
                    )),
                    ..PageSettings::default()
                },
            });
        }
        for sheet in &dump.sheets {
            let temp = dump
                .sheets
                .iter()
                .position(|candidate| candidate.name == sheet.name)
                .map(|index| format!("__pt_print_seed_{}", index + 1))
                .unwrap_or_default();
            // Renames may have collision suffixes; resolve through the already queued rename operation.
            let temp_name = operations
                .iter()
                .find_map(|operation| match operation {
                    EditOperation::RenameSheet { old_name, new_name }
                        if old_name == &sheet.name =>
                    {
                        Some(new_name.clone())
                    }
                    _ => None,
                })
                .unwrap_or(temp);
            operations.push(EditOperation::DeleteSheet { name: temp_name });
        }
    }
    let plan = if operations.is_empty() {
        None
    } else {
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
            request_id: if drop_comments {
                "pytools-print-fast-by-comment"
            } else {
                "pytools-print-keep-by-comment"
            }
            .into(),
            input: input.to_path_buf(),
            output: output.to_path_buf(),
            format: WorkbookFormat::Xlsx,
            workbook: profile,
            operations,
            preserve,
            drop_allowed: Default::default(),
        };
        plan.validate()?;
        Some(plan)
    };
    Ok(PrintPlanReport {
        input: input.to_path_buf(),
        output: output.to_path_buf(),
        matching_sheets: marked.len(),
        marked_ranges: output_sheets.len(),
        plan,
    })
}

pub fn build_config_print_plan(task: &PrintTask) -> Result<PrintPlanReport> {
    build_config_print_plans(std::slice::from_ref(task))?
        .into_iter()
        .next()
        .context("print task did not produce a plan")
}

/// Combine all enabled configuration rows for each target into a single edit
/// plan. The configured target path is the output path and must not exist;
/// each workbook is produced as a new file while keeping cross-workbook ranges
/// explicit in ExcelEditPlan.
pub fn build_config_print_plans(tasks: &[PrintTask]) -> Result<Vec<PrintPlanReport>> {
    let mut grouped = std::collections::BTreeMap::<PathBuf, Vec<&PrintTask>>::new();
    for task in tasks.iter().filter(|task| task.enabled) {
        let target = task
            .target_wb
            .as_ref()
            .context("print target workbook is empty")?;
        grouped.entry(target.clone()).or_default().push(task);
    }
    let mut reports = Vec::new();
    for (target, target_tasks) in grouped {
        anyhow::ensure!(
            is_xlsx(&target),
            "print target must be .xlsx: {}",
            target.display()
        );
        match target.symlink_metadata() {
            Ok(_) => anyhow::bail!(
                "print target already exists; PT-3-3 only writes to a new file: {}",
                target.display()
            ),
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => {}
            Err(error) => {
                return Err(error)
                    .with_context(|| format!("check print target path {}", target.display()))
            }
        }
        let first_task = target_tasks
            .first()
            .context("empty print target task group")?;
        let seed_path = first_task
            .source_wb
            .as_ref()
            .context("print source workbook is empty")?;
        anyhow::ensure!(
            is_xlsx(seed_path),
            "new print targets require an .xlsx source seed: {}",
            seed_path.display()
        );
        let dump = read_workbook(seed_path)
            .with_context(|| format!("read seed source {}", seed_path.display()))?;
        let input = seed_path;
        let profile = inspect_xlsx_profile(seed_path)?;
        let initial_sheets = dump
            .sheets
            .iter()
            .map(|sheet| sheet.name.clone())
            .collect::<Vec<_>>();
        let first_row = first_task.row_no;
        let output = target.clone();
        let mut operations = Vec::new();
        let mut target_order = Vec::new();
        let mut seed_renames = Vec::new();
        let mut used_names = initial_sheets.iter().cloned().collect::<BTreeSet<_>>();
        used_names.extend(target_tasks.iter().filter_map(|task| {
            task.target_ws
                .trim()
                .is_empty()
                .then_some(task.source_ws.clone())
                .or_else(|| Some(task.target_ws.trim().chars().take(31).collect()))
        }));
        for (index, old_name) in initial_sheets.iter().enumerate() {
            let mut temp = format!("__pt_seed_{}", index + 1);
            let mut suffix = 1;
            while used_names.contains(&temp) {
                temp = format!("__pt_seed_{}_{}", index + 1, suffix);
                suffix += 1;
            }
            used_names.insert(temp.clone());
            operations.push(EditOperation::RenameSheet {
                old_name: old_name.clone(),
                new_name: temp.clone(),
            });
            seed_renames.push(temp);
        }
        let mut created = BTreeSet::<String>::new();
        let mut next_row = BTreeMap::<String, usize>::new();
        let mut max_columns = BTreeMap::<String, usize>::new();
        let mut range_count = 0usize;
        let mut matching_sheets = BTreeSet::<String>::new();
        let mut final_setup = BTreeMap::<String, PageSettings>::new();
        for task in target_tasks {
            let source = task
                .source_wb
                .as_ref()
                .context("print source workbook is empty")?;
            anyhow::ensure!(
                source.is_file() && is_xlsx(source),
                "print source must be an existing .xlsx workbook: {}",
                source.display()
            );
            let source_dump = read_workbook(source)
                .with_context(|| format!("read print source {}", source.display()))?;
            let sheet = source_dump
                .sheets
                .iter()
                .find(|sheet| sheet.name == task.source_ws)
                .with_context(|| {
                    format!(
                        "source worksheet {:?} not found in {}",
                        task.source_ws,
                        source.display()
                    )
                })?;
            let ranges = if task.source_ranges.trim().is_empty() {
                used_bounds(sheet).into_iter().collect()
            } else {
                parse_ranges(&task.source_ranges)?
            };
            anyhow::ensure!(
                !ranges.is_empty(),
                "print source range is empty on configuration row {}",
                task.row_no
            );
            validate_static_print_values(sheet, &ranges).with_context(|| {
                format!(
                    "validate static print values for {}!{}",
                    source.display(),
                    task.source_ws
                )
            })?;
            let target_sheet = if task.target_ws.trim().is_empty() {
                task.source_ws.clone()
            } else {
                task.target_ws.trim().chars().take(31).collect::<String>()
            };
            if !created.contains(&target_sheet) {
                operations.push(EditOperation::CreateSheet {
                    name: target_sheet.clone(),
                });
                if !target_order.contains(&target_sheet) {
                    target_order.push(target_sheet.clone());
                }
                created.insert(target_sheet.clone());
                next_row.insert(target_sheet.clone(), 1);
                max_columns.insert(target_sheet.clone(), 1);
            }
            let mut block_end_row = 1usize;
            let mut block_width = 1usize;
            for (r1, c1, r2, c2) in ranges.iter().copied() {
                let start_row = *next_row.get(&target_sheet).unwrap_or(&1);
                let source_range =
                    format!("{}{}:{}{}", column_letters(c1), r1, column_letters(c2), r2);
                let style_mode = if task.mode == 3 {
                    RangeStyleMode::Basic
                } else {
                    RangeStyleMode::Full
                };
                operations.push(EditOperation::ImportRange {
                    source_workbook: source.clone(),
                    source_sheet: task.source_ws.clone(),
                    source_range,
                    target_sheet: target_sheet.clone(),
                    target_start: format!("A{start_row}"),
                    copy_values: true,
                    copy_formulas: false,
                    style_mode,
                    copy_comments: false,
                    copy_merges: true,
                    copy_row_heights: task.mode == 1,
                    copy_column_widths: task.mode == 1 || task.mode == 2,
                    copy_row_hidden: false,
                    copy_column_hidden: false,
                    copy_outline_levels: false,
                    copy_all_row_dimensions: false,
                    copy_all_column_dimensions: false,
                    materialize_blank_cells: false,
                    hide_zero: true,
                });
                let rows = r2 - r1 + 1;
                let columns = c2 - c1 + 1;
                if task.mode == 2 {
                    operations.push(EditOperation::SetRowHeight {
                        sheet: target_sheet.clone(),
                        row_index: start_row as u32,
                        height: 15.0,
                    });
                    for row in start_row + 1..start_row + rows {
                        operations.push(EditOperation::SetRowHeight {
                            sheet: target_sheet.clone(),
                            row_index: row as u32,
                            height: 15.0,
                        });
                    }
                    for offset in 0..columns {
                        let source_column = c1 + offset;
                        let max_len = (r1..=r2)
                            .map(|row| {
                                let address = format!("{}{}", column_letters(source_column), row);
                                sheet
                                    .values
                                    .get(&address)
                                    .map(|value| static_print_display_length(value, task.hide_zero))
                                    .unwrap_or_default()
                            })
                            .max()
                            .unwrap_or(0);
                        if max_len > 0 {
                            let width = (max_len as f64 * 0.95 + 1.2).clamp(6.5, 18.0);
                            operations.push(EditOperation::SetColumnWidth {
                                sheet: target_sheet.clone(),
                                column_index: (offset + 1) as u32,
                                width,
                            });
                        }
                    }
                }
                block_end_row = start_row + rows - 1;
                block_width = block_width.max(columns);
                *next_row.entry(target_sheet.clone()).or_insert(1) = start_row + rows + 1;
                let width_entry = max_columns.entry(target_sheet.clone()).or_insert(1);
                *width_entry = (*width_entry).max(columns);
                range_count += 1;
            }
            let orientation = parse_orientation(&task.orientation).unwrap_or_else(|| {
                if block_width > block_end_row {
                    PageOrientation::Landscape
                } else {
                    PageOrientation::Portrait
                }
            });
            let fit_wide = u16::try_from(task.fit_wide.max(1))
                .context("FitToPagesWide exceeds XLSX setting range")?;
            let fit_tall = u16::try_from(task.fit_tall.max(0))
                .context("FitToPagesTall exceeds XLSX setting range")?;
            let final_row = next_row
                .get(&target_sheet)
                .copied()
                .unwrap_or(2)
                .saturating_sub(2)
                .max(1);
            let final_cols = max_columns.get(&target_sheet).copied().unwrap_or(1).max(1);
            final_setup.insert(
                target_sheet.clone(),
                PageSettings {
                    orientation: Some(orientation),
                    fit_to_page: Some(true),
                    fit_to_width: Some(fit_wide),
                    fit_to_height: Some(fit_tall),
                    clear_scale: true,
                    paper_size: Some(9),
                    print_area: Some(format!("A1:{}{}", column_letters(final_cols), final_row)),
                    horizontal_centered: Some(task.center_h),
                    vertical_centered: Some(task.center_v),
                    footer: None,
                },
            );
            matching_sheets.insert(target_sheet);
        }
        anyhow::ensure!(
            !matching_sheets.is_empty(),
            "print target has no enabled, valid tasks"
        );
        for seed in seed_renames {
            operations.push(EditOperation::DeleteSheet { name: seed });
        }
        let total = target_order.len();
        for (index, sheet_name) in target_order.iter().enumerate() {
            let mut settings = final_setup.remove(sheet_name).unwrap_or_default();
            settings.footer = Some(format!("第 {} 页 / 共 {} 页", index + 1, total));
            operations.push(EditOperation::SetPageSettings {
                sheet: sheet_name.clone(),
                settings,
            });
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
            request_id: format!("pytools-print-config-{}", first_row),
            input: input.clone(),
            output: output.clone(),
            format: WorkbookFormat::Xlsx,
            workbook: profile,
            operations,
            preserve,
            drop_allowed: Default::default(),
        };
        plan.validate().context("validate grouped print plan")?;
        reports.push(PrintPlanReport {
            input: input.clone(),
            output,
            matching_sheets: matching_sheets.len(),
            marked_ranges: range_count,
            plan: Some(plan),
        });
    }
    Ok(reports)
}

/// Write configured static print workbooks as new XLSX files using Rust-owned
/// values and the workspace's approved new-workbook writer. Source workbooks
/// are read only; cached formula results are materialized as values, numeric
/// zero may be hidden per the edit plan, and comments are never copied.
pub fn write_static_print_plans(plans: &[ExcelEditPlan]) -> Result<Vec<StaticPrintExecution>> {
    plans.iter().map(write_static_print_plan).collect()
}

fn write_static_print_plan(plan: &ExcelEditPlan) -> Result<StaticPrintExecution> {
    plan.validate()
        .context("validate native static print plan")?;
    anyhow::ensure!(
        plan.format == WorkbookFormat::Xlsx && is_xlsx(&plan.output),
        "native static print writer accepts .xlsx targets only"
    );
    anyhow::ensure!(
        fs::symlink_metadata(&plan.output)
            .map(|_| false)
            .unwrap_or_else(|error| error.kind() == io::ErrorKind::NotFound),
        "static print target already exists or cannot be checked: {}",
        plan.output.display()
    );

    let mut source_books = HashMap::<PathBuf, WorkbookDump>::new();
    let mut source_uses_1904_epoch = HashMap::<PathBuf, bool>::new();
    for operation in &plan.operations {
        if let EditOperation::ImportRange {
            source_workbook, ..
        } = operation
        {
            if !source_books.contains_key(source_workbook) {
                let dump = read_workbook(source_workbook).with_context(|| {
                    format!("read static print source {}", source_workbook.display())
                })?;
                source_books.insert(source_workbook.clone(), dump);
                source_uses_1904_epoch.insert(
                    source_workbook.clone(),
                    uses_1904_date_system(source_workbook).with_context(|| {
                        format!(
                            "read date system for static print source {}",
                            source_workbook.display()
                        )
                    })?,
                );
            }
        }
    }

    let mut worksheet_names = Vec::new();
    let mut cells_by_sheet = BTreeMap::<String, Vec<StaticPrintCell>>::new();
    let mut merges_by_sheet = BTreeMap::<String, Vec<StaticPrintMerge>>::new();
    let mut imported_ranges = 0usize;
    let mut formulas_materialized = 0usize;

    for operation in &plan.operations {
        match operation {
            EditOperation::CreateSheet { name } => {
                anyhow::ensure!(
                    !worksheet_names.contains(name),
                    "static print plan creates worksheet twice: {name:?}"
                );
                worksheet_names.push(name.clone());
            }
            EditOperation::ImportRange {
                source_workbook,
                source_sheet,
                source_range,
                target_sheet,
                target_start,
                copy_values,
                copy_formulas,
                copy_comments,
                copy_merges,
                hide_zero,
                ..
            } => {
                anyhow::ensure!(
                    *copy_values && !*copy_formulas && !*copy_comments,
                    "native static print requires values-only imports without formulas or comments"
                );
                anyhow::ensure!(
                    worksheet_names.contains(target_sheet),
                    "static print import references an uncreated worksheet {target_sheet:?}"
                );
                let source_dump = source_books
                    .get(source_workbook)
                    .context("static print source workbook was not loaded")?;
                let source = source_dump
                    .sheets
                    .iter()
                    .find(|sheet| sheet.name == *source_sheet)
                    .with_context(|| {
                        format!(
                            "worksheet {source_sheet:?} not found in {}",
                            source_workbook.display()
                        )
                    })?;
                let (source_first, source_last) = parse_a1_range(source_range)?;
                let (target_row, target_column) = parse_a1(target_start).with_context(|| {
                    format!("invalid static print target cell {target_start:?}")
                })?;
                let mut covered_by_string_merge = BTreeSet::new();
                let mut target_merges = Vec::new();
                for merge in source.merges.iter().filter(|_| *copy_merges) {
                    let Ok((
                        (merge_first_row, merge_first_column),
                        (merge_last_row, merge_last_column),
                    )) = parse_a1_range(merge)
                    else {
                        continue;
                    };
                    if merge_first_row < source_first.0
                        || merge_first_column < source_first.1
                        || merge_last_row > source_last.0
                        || merge_last_column > source_last.1
                    {
                        continue;
                    }
                    let source_top_left = a1_address(merge_first_row, merge_first_column);
                    let Some(text) = source
                        .values
                        .get(&source_top_left)
                        .filter(|value| {
                            value.get("type").and_then(serde_json::Value::as_str) == Some("string")
                        })
                        .and_then(|value| value.get("value"))
                        .and_then(serde_json::Value::as_str)
                    else {
                        continue;
                    };
                    if source.formulas.contains_key(&source_top_left) {
                        formulas_materialized += 1;
                    }
                    let target_first_row = target_row + (merge_first_row - source_first.0);
                    let target_first_column = target_column + (merge_first_column - source_first.1);
                    let target_last_row = target_row + (merge_last_row - source_first.0);
                    let target_last_column = target_column + (merge_last_column - source_first.1);
                    for row in merge_first_row..=merge_last_row {
                        for column in merge_first_column..=merge_last_column {
                            covered_by_string_merge.insert((row, column));
                        }
                    }
                    target_merges.push(StaticPrintMerge {
                        first_row: target_first_row - 1,
                        first_column: target_first_column - 1,
                        last_row: target_last_row - 1,
                        last_column: target_last_column - 1,
                        text: text.to_owned(),
                    });
                }
                merges_by_sheet
                    .entry(target_sheet.clone())
                    .or_default()
                    .extend(target_merges);
                for row in source_first.0..=source_last.0 {
                    for column in source_first.1..=source_last.1 {
                        if covered_by_string_merge.contains(&(row, column)) {
                            continue;
                        }
                        let address = a1_address(row, column);
                        let Some(value) = source.values.get(&address) else {
                            continue;
                        };
                        anyhow::ensure!(
                            value.get("type").and_then(serde_json::Value::as_str) != Some("error"),
                            "cannot print Excel error value at {}!{}",
                            source_sheet,
                            address
                        );
                        if source.formulas.contains_key(&address) {
                            formulas_materialized += 1;
                        }
                        if *hide_zero && static_print_is_numeric_zero(value) {
                            continue;
                        }
                        let source_uses_1904 = source_uses_1904_epoch
                            .get(source_workbook)
                            .copied()
                            .unwrap_or(false);
                        let mut output_value = value.clone();
                        if value.get("type").and_then(serde_json::Value::as_str)
                            == Some("datetime_serial")
                        {
                            let serial = value
                                .get("value")
                                .and_then(serde_json::Value::as_f64)
                                .context("static print date serial is invalid")?;
                            let serial = date_serial_for_1900_output(serial, source_uses_1904)?;
                            let serial = serde_json::Number::from_f64(serial)
                                .context("static print date serial is not finite")?;
                            output_value["value"] = serde_json::Value::Number(serial);
                        }
                        cells_by_sheet
                            .entry(target_sheet.clone())
                            .or_default()
                            .push(StaticPrintCell {
                                row: target_row - 1 + (row - source_first.0),
                                column: target_column - 1 + (column - source_first.1),
                                value: output_value,
                            });
                    }
                }
                imported_ranges += 1;
            }
            EditOperation::SetPageSettings { sheet, .. } => {
                anyhow::ensure!(
                    worksheet_names.contains(sheet),
                    "page settings reference an uncreated worksheet {sheet:?}"
                );
            }
            EditOperation::SetRowHeight { sheet, .. } => {
                anyhow::ensure!(
                    worksheet_names.contains(sheet),
                    "row height references an uncreated worksheet {sheet:?}"
                );
            }
            EditOperation::SetColumnWidth {
                sheet,
                column_index,
                width,
            } => {
                anyhow::ensure!(
                    worksheet_names.contains(sheet),
                    "column width references an uncreated worksheet {sheet:?}"
                );
                let _ = (column_index, width);
            }
            EditOperation::RenameSheet { .. } | EditOperation::DeleteSheet { .. } => {
                // A new workbook has no seed sheet to rename or delete.
            }
            other => anyhow::bail!(
                "native static print writer does not support plan operation {:?}",
                other.kind()
            ),
        }
    }
    anyhow::ensure!(
        !worksheet_names.is_empty(),
        "static print plan has no target worksheets"
    );

    let mut workbook = Workbook::new();
    for name in &worksheet_names {
        workbook.add_worksheet().set_name(name)?;
    }
    for name in &worksheet_names {
        let worksheet = workbook.worksheet_from_name(name)?;
        if let Some(merges) = merges_by_sheet.get(name) {
            for merge in merges {
                worksheet.merge_range(
                    to_row(merge.first_row)?,
                    to_column(merge.first_column)?,
                    to_row(merge.last_row)?,
                    to_column(merge.last_column)?,
                    &merge.text,
                    &Format::new(),
                )?;
            }
        }
        if let Some(cells) = cells_by_sheet.get(name) {
            for cell in cells {
                write_static_print_cell(worksheet, cell.row, cell.column, &cell.value)?;
            }
        }
    }
    for operation in &plan.operations {
        match operation {
            EditOperation::SetRowHeight {
                sheet,
                row_index,
                height,
            } => {
                workbook
                    .worksheet_from_name(sheet)?
                    .set_row_height(row_index.saturating_sub(1), *height)?;
            }
            EditOperation::SetColumnWidth {
                sheet,
                column_index,
                width,
            } => {
                workbook.worksheet_from_name(sheet)?.set_column_width(
                    u16::try_from(column_index.saturating_sub(1))
                        .context("static print column width index is out of range")?,
                    *width,
                )?;
            }
            EditOperation::SetPageSettings { sheet, settings } => {
                apply_static_print_page_settings(workbook.worksheet_from_name(sheet)?, settings)?;
            }
            _ => {}
        }
    }

    let output_parent = plan
        .output
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    fs::create_dir_all(output_parent).with_context(|| {
        format!(
            "create static print output directory {}",
            output_parent.display()
        )
    })?;
    save_new_workbook_without_overwrite(&mut workbook, &plan.output)?;
    Ok(StaticPrintExecution {
        engine: "rust_xlsxwriter",
        output: plan.output.clone(),
        worksheet_count: worksheet_names.len(),
        imported_ranges,
        formulas_materialized,
        comments_dropped: true,
    })
}

fn apply_static_print_page_settings(
    worksheet: &mut Worksheet,
    settings: &PageSettings,
) -> Result<()> {
    if let Some(orientation) = settings.orientation {
        match orientation {
            PageOrientation::Landscape => {
                worksheet.set_landscape();
            }
            PageOrientation::Portrait => {
                worksheet.set_portrait();
            }
        }
    }
    if settings.fit_to_page == Some(true) {
        worksheet.set_print_fit_to_pages(
            settings.fit_to_width.unwrap_or(1),
            settings.fit_to_height.unwrap_or(0),
        );
    }
    if let Some(paper_size) = settings.paper_size {
        worksheet.set_paper_size(u8::try_from(paper_size).context("paper size is out of range")?);
    }
    if let Some(area) = &settings.print_area {
        let ((first_row, first_column), (last_row, last_column)) = parse_a1_range(area)?;
        worksheet.set_print_area(
            u32::try_from(first_row - 1).context("print area row is out of range")?,
            u16::try_from(first_column - 1).context("print area column is out of range")?,
            u32::try_from(last_row - 1).context("print area row is out of range")?,
            u16::try_from(last_column - 1).context("print area column is out of range")?,
        )?;
    }
    if let Some(centered) = settings.horizontal_centered {
        worksheet.set_print_center_horizontally(centered);
    }
    if let Some(centered) = settings.vertical_centered {
        worksheet.set_print_center_vertically(centered);
    }
    if let Some(footer) = &settings.footer {
        worksheet.set_footer(footer);
    }
    Ok(())
}

fn write_static_print_cell(
    worksheet: &mut Worksheet,
    row: usize,
    column: usize,
    value: &serde_json::Value,
) -> Result<()> {
    let row = to_row(row)?;
    let column = to_column(column)?;
    let value_type = value
        .get("type")
        .and_then(serde_json::Value::as_str)
        .context("static print cell has no value type")?;
    let data = value
        .get("value")
        .context("static print cell has no value")?;
    match value_type {
        "string" | "datetime_iso" | "duration_iso" => {
            worksheet.write_string(row, column, data.as_str().unwrap_or_default())?;
        }
        "integer" => {
            let integer = data.as_i64().context("static print integer is invalid")?;
            let number = integer as f64;
            anyhow::ensure!(
                format!("{number:.0}").parse::<i64>().ok() == Some(integer),
                "integer {integer} cannot be represented exactly as an XLSX numeric cell"
            );
            worksheet.write_number(row, column, number)?;
        }
        "number" => {
            let number = data.as_f64().context("static print number is invalid")?;
            anyhow::ensure!(number.is_finite(), "static print number is not finite");
            worksheet.write_number(row, column, number)?;
        }
        "datetime_serial" => {
            let number = data
                .as_f64()
                .context("static print date serial is invalid")?;
            anyhow::ensure!(number.is_finite(), "static print date serial is not finite");
            let format = Format::new().set_num_format("yyyy-mm-dd h:mm:ss");
            worksheet.write_number_with_format(row, column, number, &format)?;
        }
        "boolean" => {
            worksheet.write_boolean(
                row,
                column,
                data.as_bool().context("static print boolean is invalid")?,
            )?;
        }
        "error" => anyhow::bail!("Excel error values cannot be written to static print output"),
        other => anyhow::bail!("unsupported static print cell type {other:?}"),
    }
    Ok(())
}

fn static_print_is_numeric_zero(value: &serde_json::Value) -> bool {
    matches!(
        value.get("type").and_then(serde_json::Value::as_str),
        Some("number" | "integer")
    ) && value
        .get("value")
        .and_then(serde_json::Value::as_f64)
        .is_some_and(|number| number == 0.0)
}

fn date_serial_for_1900_output(serial: f64, source_uses_1904_epoch: bool) -> Result<f64> {
    let output_serial = if source_uses_1904_epoch {
        serial + 1462.0
    } else {
        serial
    };
    anyhow::ensure!(
        output_serial.is_finite(),
        "static print date serial is not finite"
    );
    Ok(output_serial)
}

fn parse_a1_range(text: &str) -> Result<((usize, usize), (usize, usize))> {
    let (start, end) = text.split_once(':').unwrap_or((text, text));
    let first = parse_a1(start).with_context(|| format!("invalid A1 range start {start:?}"))?;
    let last = parse_a1(end).with_context(|| format!("invalid A1 range end {end:?}"))?;
    anyhow::ensure!(
        last.0 >= first.0 && last.1 >= first.1,
        "invalid A1 range {text:?}"
    );
    Ok((first, last))
}

fn a1_address(row: usize, column: usize) -> String {
    format!("{}{}", column_letters(column), row)
}

fn to_row(row: usize) -> Result<u32> {
    u32::try_from(row).context("static print row exceeds XLSX limits")
}

fn to_column(column: usize) -> Result<u16> {
    u16::try_from(column).context("static print column exceeds XLSX limits")
}

fn save_new_workbook_without_overwrite(workbook: &mut Workbook, output: &Path) -> Result<()> {
    anyhow::ensure!(
        fs::symlink_metadata(output)
            .map(|_| false)
            .unwrap_or_else(|error| error.kind() == io::ErrorKind::NotFound),
        "static print target already exists or cannot be checked: {}",
        output.display()
    );
    let parent = output
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let file_name = output
        .file_name()
        .and_then(|name| name.to_str())
        .context("static print output filename is not valid UTF-8")?;
    let temp_path = (0..1000)
        .find_map(|_| {
            let serial = NEXT_STATIC_PRINT_TEMP_ID.fetch_add(1, Ordering::Relaxed);
            let candidate = parent.join(format!(
                ".{file_name}.pytools-print-{}-{serial}.tmp",
                std::process::id()
            ));
            match OpenOptions::new()
                .write(true)
                .create_new(true)
                .open(&candidate)
            {
                Ok(file) => {
                    drop(file);
                    Some(Ok(candidate))
                }
                Err(error) if error.kind() == io::ErrorKind::AlreadyExists => None,
                Err(error) => Some(Err(error)),
            }
        })
        .context("could not allocate a temporary static print workbook path")??;

    let write_result = (|| -> Result<()> {
        workbook.save(&temp_path).with_context(|| {
            format!(
                "write temporary static print workbook {}",
                temp_path.display()
            )
        })?;
        crate::file_commit::create_target_from_candidate(&temp_path, output).with_context(
            || {
                format!(
                    "publish validated static print workbook without overwriting {}",
                    output.display()
                )
            },
        )?;
        Ok(())
    })();
    let _ = fs::remove_file(&temp_path);
    write_result
}

fn comment_ranges(sheet: &SheetDump) -> Vec<(usize, usize, usize, usize)> {
    let mut starts = std::collections::BTreeMap::new();
    let mut ends = std::collections::BTreeMap::new();
    for (address, text) in &sheet.comments {
        let normalized = normalize(text);
        let Some(id) = number_after(&normalized, "打印区域") else {
            continue;
        };
        let Some((row, col)) = parse_a1(address) else {
            continue;
        };
        if normalized
            .replace(' ', "")
            .contains(&format!("打印区域#{id}"))
        {
            ends.insert(id, (row, col));
        } else {
            starts.insert(id, (row, col));
        }
    }
    starts
        .iter()
        .filter_map(|(id, &(r1, c1))| {
            ends.get(id)
                .map(|&(r2, c2)| (r1.min(r2), c1.min(c2), r1.max(r2), c1.max(c2)))
        })
        .collect()
}
fn number_after(text: &str, key: &str) -> Option<usize> {
    let i = text.find(key)? + key.len();
    let digits = text[i..]
        .trim_start_matches(['#', ' '])
        .chars()
        .take_while(char::is_ascii_digit)
        .collect::<String>();
    digits.parse().ok()
}
fn used_bounds(sheet: &SheetDump) -> Option<(usize, usize, usize, usize)> {
    let mut min_r = usize::MAX;
    let mut min_c = usize::MAX;
    let mut max_r = 0;
    let mut max_c = 0;
    for address in sheet.values.keys().chain(sheet.formulas.keys()) {
        if let Some((r, c)) = parse_a1(address) {
            min_r = min_r.min(r);
            min_c = min_c.min(c);
            max_r = max_r.max(r);
            max_c = max_c.max(c);
        }
    }
    (max_r > 0).then_some((min_r, min_c, max_r, max_c))
}
fn value_to_text(value: &serde_json::Value) -> String {
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
fn static_print_display_length(value: &serde_json::Value, hide_zero: bool) -> usize {
    let value_type = value.get("type").and_then(serde_json::Value::as_str);
    if value_type == Some("datetime_serial") {
        return 19;
    }
    let numeric_zero = matches!(value_type, Some("number" | "integer"))
        && value
            .get("value")
            .and_then(serde_json::Value::as_f64)
            .is_some_and(|number| number == 0.0);
    if hide_zero && numeric_zero {
        0
    } else {
        value_to_text(value).trim().chars().count()
    }
}
fn parse_ranges(text: &str) -> Result<Vec<(usize, usize, usize, usize)>> {
    text.replace('；', ";")
        .split(';')
        .filter(|part| !part.trim().is_empty())
        .map(|part| {
            let (a, b) = part.split_once(':').unwrap_or((part, part));
            let (r1, c1) = parse_a1(a.trim()).context("invalid A1 range start")?;
            let (r2, c2) = parse_a1(b.trim()).context("invalid A1 range end")?;
            anyhow::ensure!(r2 >= r1 && c2 >= c1, "invalid range ordering");
            Ok((r1, c1, r2, c2))
        })
        .collect()
}
fn validate_static_print_values(
    sheet: &SheetDump,
    ranges: &[(usize, usize, usize, usize)],
) -> Result<()> {
    anyhow::ensure!(
        sheet.formula_status == "read_by_calamine",
        "formula metadata is unavailable for static print worksheet {:?}",
        sheet.name
    );
    for address in sheet.formulas.keys() {
        if address_in_ranges(address, ranges) {
            let Some(value) = sheet.values.get(address) else {
                anyhow::bail!(
                    "formula cached value is missing at {}!{}",
                    sheet.name,
                    address
                );
            };
            if value.get("type").and_then(serde_json::Value::as_str) == Some("string")
                && value
                    .get("value")
                    .and_then(serde_json::Value::as_str)
                    .is_some_and(str::is_empty)
            {
                anyhow::bail!(
                    "formula cached value is blank or unavailable at {}!{}",
                    sheet.name,
                    address
                );
            }
        }
    }
    for (address, value) in &sheet.values {
        if address_in_ranges(address, ranges)
            && value.get("type").and_then(serde_json::Value::as_str) == Some("error")
        {
            anyhow::bail!(
                "Excel error value at {}!{}: {}",
                sheet.name,
                address,
                value
                    .get("value")
                    .and_then(serde_json::Value::as_str)
                    .unwrap_or("unknown")
            );
        }
    }
    Ok(())
}

fn address_in_ranges(address: &str, ranges: &[(usize, usize, usize, usize)]) -> bool {
    let Some((row, column)) = parse_a1(address) else {
        return false;
    };
    ranges
        .iter()
        .any(|(start_row, start_column, end_row, end_column)| {
            row >= *start_row && row <= *end_row && column >= *start_column && column <= *end_column
        })
}

fn parse_a1(text: &str) -> Option<(usize, usize)> {
    let text = text.trim().trim_start_matches('$');
    let split = text.find(|c: char| c.is_ascii_digit())?;
    let col = text[..split].bytes().try_fold(0usize, |acc, b| {
        acc.checked_mul(26)?
            .checked_add((b.to_ascii_uppercase() - b'A' + 1) as usize)
    })?;
    Some((text[split..].parse().ok()?, col))
}
fn column_letters(mut col: usize) -> String {
    let mut out = String::new();
    while col > 0 {
        out.insert(0, (b'A' + ((col - 1) % 26) as u8) as char);
        col = (col - 1) / 26;
    }
    out
}
fn unique_sheet_name(proposed: &str, used: &mut BTreeSet<String>) -> String {
    let base = proposed
        .chars()
        .filter(|character| !matches!(character, '[' | ']' | ':' | '*' | '?' | '/' | '\\'))
        .collect::<String>();
    let base = if base.trim().is_empty() {
        "Sheet".to_owned()
    } else {
        base
    };
    let mut candidate = base.chars().take(31).collect::<String>();
    let mut suffix = 1usize;
    while used.contains(&candidate) {
        let tag = format!("_{suffix}");
        let keep = 31usize.saturating_sub(tag.chars().count());
        candidate = format!("{}{}", base.chars().take(keep).collect::<String>(), tag);
        suffix += 1;
    }
    used.insert(candidate.clone());
    candidate
}
fn orientation_from_comment(text: &str) -> Option<PageOrientation> {
    let landscape = text.contains("横向");
    let portrait = text.contains("纵向");
    match (landscape, portrait) {
        (true, false) => Some(PageOrientation::Landscape),
        (false, true) => Some(PageOrientation::Portrait),
        _ => None,
    }
}
fn parse_orientation(text: &str) -> Option<PageOrientation> {
    let text = normalize(text);
    if text.contains("横向") || text.eq_ignore_ascii_case("landscape") {
        Some(PageOrientation::Landscape)
    } else if text.contains("纵向") || text.eq_ignore_ascii_case("portrait") {
        Some(PageOrientation::Portrait)
    } else {
        None
    }
}
fn normalize(text: &str) -> String {
    text.trim().replace('　', " ").replace('\u{00a0}', " ")
}
fn is_xlsx(path: &Path) -> bool {
    path.extension()
        .and_then(|e| e.to_str())
        .is_some_and(|e| e.eq_ignore_ascii_case("xlsx"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    #[test]
    fn static_print_autofit_distinguishes_numeric_zero_from_text_zero() {
        assert_eq!(
            static_print_display_length(&json!({"type": "number", "value": 0}), true),
            0
        );
        assert_eq!(
            static_print_display_length(&json!({"type": "integer", "value": 0}), true),
            0
        );
        assert_eq!(
            static_print_display_length(&json!({"type": "string", "value": "0"}), true),
            1
        );
        assert_eq!(
            static_print_display_length(&json!({"type": "string", "value": "0.00"}), true),
            4
        );
        assert_eq!(
            static_print_display_length(&json!({"type": "number", "value": 0}), false),
            1
        );
        let date_zero = json!({"type": "datetime_serial", "value": 0});
        assert!(!static_print_is_numeric_zero(&date_zero));
        assert_eq!(static_print_display_length(&date_zero, true), 19);
        assert_eq!(date_serial_for_1900_output(0.0, true).unwrap(), 1462.0);
        assert_eq!(date_serial_for_1900_output(42.5, false).unwrap(), 42.5);
    }

    #[test]
    fn static_print_rejects_excel_error_cells_in_selected_ranges() {
        let mut values = BTreeMap::new();
        values.insert("B2".to_owned(), json!({"type": "error", "value": "Div0"}));
        let sheet = SheetDump {
            name: "Source".to_owned(),
            values,
            formulas: BTreeMap::new(),
            formula_status: "read_by_calamine".to_owned(),
            merges: Vec::new(),
            merge_status: "read_by_calamine".to_owned(),
            comments: BTreeMap::new(),
            comments_status: "read_by_namespace_aware_ooxml_supplement".to_owned(),
            conditional_formatting: Vec::new(),
            conditional_formatting_status: "read_by_namespace_aware_ooxml_supplement".to_owned(),
            data_validations: Vec::new(),
            data_validations_status: "read_by_namespace_aware_ooxml_supplement".to_owned(),
        };

        let error = validate_static_print_values(&sheet, &[(1, 2, 3, 4)])
            .expect_err("Excel error cells must prevent static print output");
        assert!(error.to_string().contains("Excel error value at Source!B2"));
    }
}
