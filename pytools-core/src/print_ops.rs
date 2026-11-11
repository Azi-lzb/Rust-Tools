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
    /// 源单元格样式经 OOXML 读取映射后的写入格式；None = 默认格式。
    format: Option<rust_xlsxwriter::Format>,
}

#[derive(Clone, Debug)]
struct StaticPrintMerge {
    first_row: usize,
    first_column: usize,
    last_row: usize,
    last_column: usize,
    text: String,
    /// 锚点单元格样式（若源单元格带样式）。
    format: Option<rust_xlsxwriter::Format>,
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
        let mut seen_target_sheets = BTreeSet::<String>::new();
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
            // 重复目标 Sheet 对齐 Openpyxl 实际行为：清空重写（后写任务覆盖先写
            // 任务的行）。reset 按 task 粒度，同一 task 内的多区域正常堆叠。
            if seen_target_sheets.contains(&task.target_ws) {
                next_row.insert(task.target_ws.clone(), 1);
            }
            seen_target_sheets.insert(task.target_ws.clone());
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
                    hide_zero: task.hide_zero,
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
            // 打印区域末行对齐 openpyxl 引擎语义（apply_setup(ws, out_row - 1, …)）：
            // out_row/next_row 在最后一个区域后已 +1（块间空行），打印区域包含
            // 该尾随空行，即末行 = next_row - 1。此前 Rust 侧多减 1 导致
            // 打印区域比 Openpyxl 输出少一行（实测 8 场景对拍暴露）。
            let final_row = next_row
                .get(&target_sheet)
                .copied()
                .unwrap_or(2)
                .saturating_sub(1)
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
            settings.footer = Some(format!("&C第 {} 页 / 共 {} 页", index + 1, total));
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
    let mut source_styles =
        HashMap::<PathBuf, std::sync::Arc<crate::ooxml_cell_styles::SourceStyles>>::new();
    // 阶段计时（PYTOOLS_PRINT_STAGE_TIMING=1 时输出到 stderr，供性能分析用）。
    let stage_timing_enabled = std::env::var_os("PYTOOLS_PRINT_STAGE_TIMING").is_some();
    let stage_started = std::time::Instant::now();
    let styles_load_ms = std::cell::Cell::new(0u128);
    let ooxml_parse_ms = std::cell::Cell::new(0u128);
    let style_map_ms = std::cell::Cell::new(0u128);
    let region_extract_ms = std::cell::Cell::new(0u128);
    let region_reorg_ms = std::cell::Cell::new(0u128);
    let source_load_started = std::time::Instant::now();
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
    let source_read_ms = source_load_started.elapsed();
    let extract_reorg_started = std::time::Instant::now();

    let mut worksheet_names = Vec::new();
    let mut cells_by_sheet = BTreeMap::<String, Vec<StaticPrintCell>>::new();
    let mut merges_by_sheet = BTreeMap::<String, Vec<StaticPrintMerge>>::new();
    let mut comments_by_sheet = BTreeMap::<String, Vec<(String, String)>>::new();
    // mode 1 源行高/列宽（OOXML 窄读补充）：目标行/列 -> 值，后写覆盖（对齐
    // openpyxl 引擎逐块赋值语义）。在计划自带的 Set 行高列宽操作之前应用。
    let mut layout_row_heights = BTreeMap::<String, BTreeMap<u32, f64>>::new();
    let mut layout_column_widths = BTreeMap::<String, BTreeMap<u32, f64>>::new();
    let mut source_layouts =
        HashMap::<(PathBuf, String), std::sync::Arc<crate::ooxml_cell_styles::SheetLayout>>::new();
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
                style_mode,
                copy_comments,
                copy_merges,
                copy_row_heights,
                copy_column_widths,
                hide_zero,
                ..
            } => {
                anyhow::ensure!(
                    *copy_values && !*copy_formulas,
                    "native static print requires values-only imports without formulas"
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
                // OOXML 窄范围样式读取：styles.xml + worksheet cell s 索引，
                // 只覆盖本次导入涉及的源单元格。
                let source_style_map = source_styles
                    .entry(source_workbook.clone())
                    .or_insert_with(|| {
                        let styles_started = std::time::Instant::now();
                        let loaded = crate::ooxml_cell_styles::SourceStyles::load(source_workbook)
                            .unwrap_or_else(|error| {
                                eprintln!(
                                    "[3-1][Rust] 源样式读取失败（输出将使用默认格式）：{error}"
                                );
                                crate::ooxml_cell_styles::SourceStyles::default()
                            });
                        let elapsed = styles_started.elapsed().as_millis();
                        styles_load_ms.set(styles_load_ms.get() + elapsed);
                        // styles.xml 解析属于 OOXML 窄读的一部分。
                        ooxml_parse_ms.set(ooxml_parse_ms.get() + elapsed);
                        std::sync::Arc::new(loaded)
                    })
                    .clone();
                let indices_started = std::time::Instant::now();
                let cell_style_indices = source_style_map
                    .cell_style_indices(source_workbook, source_sheet)
                    .unwrap_or_default();
                ooxml_parse_ms.set(ooxml_parse_ms.get() + indices_started.elapsed().as_millis());
                // mode 1 源行高列宽：OOXML 窄读（calamine 不暴露）。对齐
                // openpyxl 引擎 copy_block：行高按源行→目标行平移（mode 1），
                // 列宽按源列→目标列平移（mode 1/2），仅显式设置项，后写覆盖。
                if *copy_row_heights || *copy_column_widths {
                    let layout = source_layouts
                        .entry((source_workbook.clone(), source_sheet.clone()))
                        .or_insert_with(|| {
                            std::sync::Arc::new(
                                source_style_map
                                    .sheet_layout(source_workbook, source_sheet)
                                    .unwrap_or_else(|error| {
                                        eprintln!(
                                            "[3-1][Rust] 源行高列宽读取失败（输出使用默认行高列宽）：{error}"
                                        );
                                        crate::ooxml_cell_styles::SheetLayout::default()
                                    }),
                            )
                        })
                        .clone();
                    if *copy_row_heights {
                        for (row, height) in &layout.row_heights {
                            if *row < source_first.0 as u32 || *row > source_last.0 as u32 {
                                continue;
                            }
                            let target = (target_row as u32) + row - source_first.0 as u32;
                            layout_row_heights
                                .entry(target_sheet.clone())
                                .or_default()
                                .insert(target, *height);
                        }
                    }
                    if *copy_column_widths {
                        // openpyxl 语义：`src_ws.column_dimensions[列字母].width`
                        // 只在显式设置时命中，未设置的列取 openpyxl 的
                        // DEFAULT_COLUMN_WIDTH（13.0），并且写入目标时同样落盘
                        // （不是「留空用默认」）。
                        for offset in 0..(source_last.1 - source_first.1 + 1) {
                            let source_column = (source_first.1 + offset) as u32;
                            let width = layout
                                .column_widths
                                .get(&source_column)
                                .copied()
                                .unwrap_or(crate::ooxml_cell_styles::DEFAULT_COLUMN_WIDTH);
                            let target =
                                (target_column as u32) + source_column - source_first.1 as u32;
                            layout_column_widths
                                .entry(target_sheet.clone())
                                .or_default()
                                .insert(target, width);
                        }
                    }
                }
                // Format 去重缓存：descriptor → 已构造 Format。
                let mut format_cache = HashMap::<
                    crate::ooxml_cell_styles::CellStyleInfo,
                    rust_xlsxwriter::Format,
                >::new();
                let mut covered_by_string_merge = BTreeSet::new();
                let mut target_merges = Vec::new();
                let format_of = |address: &str,
                                 cache: &mut HashMap<
                    crate::ooxml_cell_styles::CellStyleInfo,
                    rust_xlsxwriter::Format,
                >| {
                    let map_started = std::time::Instant::now();
                    let result = cell_style_indices
                        .get(address)
                        .map(|index| source_style_map.descriptor(*index))
                        .filter(|descriptor| {
                            *descriptor != crate::ooxml_cell_styles::CellStyleInfo::default()
                        })
                        .map(|descriptor| {
                            cache
                                .entry(descriptor.clone())
                                .or_insert_with(|| {
                                    build_writer_format(
                                        &descriptor,
                                        matches!(style_mode, RangeStyleMode::Full),
                                    )
                                })
                                .clone()
                        });
                    style_map_ms.set(style_map_ms.get() + map_started.elapsed().as_millis());
                    result
                };
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
                    // 合并区必须整块保留（原 Python 复制所有完全落在区域内的合并）。
                    // 锚点可能不是字符串（数字/日期），按显示文本取锚点值。
                    let text = source
                        .values
                        .get(&source_top_left)
                        .and_then(|value| value.get("value"))
                        .map(display_text_of_value)
                        .unwrap_or_default();
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
                        format: format_of(&source_top_left, &mut format_cache),
                    });
                }
                merges_by_sheet
                    .entry(target_sheet.clone())
                    .or_default()
                    .extend(target_merges);
                if *copy_comments {
                    for (address, text) in &source.comments {
                        let Some((row, column)) = parse_a1(address) else {
                            continue;
                        };
                        if row < source_first.0
                            || row > source_last.0
                            || column < source_first.1
                            || column > source_last.1
                        {
                            continue;
                        }
                        let target_address = a1_address(
                            target_row + (row - source_first.0),
                            target_column + (column - source_first.1),
                        );
                        comments_by_sheet
                            .entry(target_sheet.clone())
                            .or_default()
                            .push((target_address, text.clone()));
                    }
                }
                let extract_started = std::time::Instant::now();
                for row in source_first.0..=source_last.0 {
                    for column in source_first.1..=source_last.1 {
                        if covered_by_string_merge.contains(&(row, column)) {
                            continue;
                        }
                        let address = a1_address(row, column);
                        let Some(value) = source.values.get(&address) else {
                            continue;
                        };
                        // 原 PyTools print_310.py 不拒绝错误值，按显示文本输出。
                        if source.formulas.contains_key(&address) {
                            formulas_materialized += 1;
                        }
                        // 零值隐藏：对齐 Openpyxl 行为——值清空但保留单元格
                        // 样式（数字格式等），避免输出出现格式断档。
                        let hidden_zero = *hide_zero && static_print_is_numeric_zero(value);
                        let cell_format = if hidden_zero {
                            None
                        } else {
                            format_of(&address, &mut format_cache)
                        };
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
                            // 纯时间格式（hh:mm:ss 等）且串行值 < 1（纯时刻，无日期
                            // 部分）：串行值在 1900/1904 系统下显示一致，不做跨纪元
                            // 平移，与 openpyxl 引擎对 1904 簿 time 单元格的取值一致。
                            // 串行值 ≥ 1 含整天数（如 =B3+1），仍是真实日期时间，照常平移。
                            let time_only = source_uses_1904
                                && serial < 1.0
                                && cell_style_indices
                                    .get(&address)
                                    .map(|index| source_style_map.descriptor(*index))
                                    .and_then(|descriptor| descriptor.num_fmt)
                                    .as_deref()
                                    .map(numfmt_is_time_only)
                                    .unwrap_or(false);
                            let serial = if time_only {
                                serial
                            } else {
                                date_serial_for_1900_output(serial, source_uses_1904)?
                            };
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
                                value: if hidden_zero {
                                    serde_json::json!({"type": "blank", "value": null})
                                } else {
                                    output_value
                                },
                                format: if hidden_zero {
                                    format_of(&address, &mut format_cache)
                                } else {
                                    cell_format
                                },
                            });
                    }
                }
                region_extract_ms
                    .set(region_extract_ms.get() + extract_started.elapsed().as_millis());
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

    let extract_reorg_ms = extract_reorg_started.elapsed();
    // 区域提取 = 逐单元格读取/取值；区域重组 = 其余坐标平移、合并/批注归集
    // 与目标分组工作。两者之和即本阶段的 extract_reorg_ms。
    let region_reorg_ms = std::cell::Cell::new(
        extract_reorg_ms
            .as_millis()
            .saturating_sub(region_extract_ms.get()),
    );
    let target_create_started = std::time::Instant::now();
    let mut workbook = Workbook::new();
    for name in &worksheet_names {
        workbook.add_worksheet().set_name(name)?;
    }
    let target_create_ms = target_create_started.elapsed();
    let target_write_started = std::time::Instant::now();
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
                    merge.format.as_ref().unwrap_or(&Format::new()),
                )?;
            }
        }
        if let Some(cells) = cells_by_sheet.get(name) {
            for cell in cells {
                write_static_print_cell(
                    worksheet,
                    cell.row,
                    cell.column,
                    &cell.value,
                    cell.format.as_ref(),
                )?;
            }
        }
        if let Some(comments) = comments_by_sheet.get(name) {
            for (address, text) in comments {
                let Some((row, column)) = parse_a1(address) else {
                    continue;
                };
                let note = rust_xlsxwriter::Note::new(text.as_str());
                worksheet.insert_note(to_row(row)?, to_column(column)?, &note)?;
            }
        }
    }
    let target_write_ms = target_write_started.elapsed();
    let page_settings_started = std::time::Instant::now();
    // 先应用 mode 1 源行高列宽（逐块收集，后写覆盖），再应用计划显式
    // Set 行高列宽（mode 2 自适应等），与 openpyxl 引擎的执行顺序一致。
    for (sheet, heights) in &layout_row_heights {
        let worksheet = workbook.worksheet_from_name(sheet)?;
        for (row, height) in heights {
            worksheet.set_row_height(row.saturating_sub(1), *height)?;
        }
    }
    for (sheet, widths) in &layout_column_widths {
        let worksheet = workbook.worksheet_from_name(sheet)?;
        for (column, width) in widths {
            set_column_width_exact(worksheet, *column, *width)?;
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
                set_column_width_exact(
                    workbook.worksheet_from_name(sheet)?,
                    *column_index,
                    *width,
                )?;
            }
            EditOperation::SetPageSettings { sheet, settings } => {
                apply_static_print_page_settings(workbook.worksheet_from_name(sheet)?, settings)?;
            }
            _ => {}
        }
    }

    let page_settings_ms = page_settings_started.elapsed();
    let save_started = std::time::Instant::now();
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
    if stage_timing_enabled {
        eprintln!(
            "[3-1][stage-timing] {}",
            serde_json::json!({
                "output": plan.output.file_name().map(|name| name.to_string_lossy().into_owned()),
                // 目标 §14.4 的 12 项阶段细分（按 Rust 实现的可测边界映射）。
                "source_read_ms": u64::try_from(source_read_ms.as_millis()).unwrap_or(u64::MAX),
                "ooxml_parse_ms": u64::try_from(ooxml_parse_ms.get()).unwrap_or(u64::MAX),
                "styles_load_ms": u64::try_from(styles_load_ms.get()).unwrap_or(u64::MAX),
                "style_map_ms": u64::try_from(style_map_ms.get()).unwrap_or(u64::MAX),
                "region_extract_ms": u64::try_from(region_extract_ms.get()).unwrap_or(u64::MAX),
                "region_reorg_ms": u64::try_from(region_reorg_ms.get()).unwrap_or(u64::MAX),                "extract_reorg_ms": u64::try_from(extract_reorg_ms.as_millis()).unwrap_or(u64::MAX),
                "target_create_ms": u64::try_from(target_create_ms.as_millis()).unwrap_or(u64::MAX),
                "target_write_ms": u64::try_from(target_write_ms.as_millis()).unwrap_or(u64::MAX),
                "page_settings_ms": u64::try_from(page_settings_ms.as_millis()).unwrap_or(u64::MAX),
                "save_ms": u64::try_from(save_started.elapsed().as_millis()).unwrap_or(u64::MAX),
                "total_ms": u64::try_from(stage_started.elapsed().as_millis()).unwrap_or(u64::MAX),
            })
        );
    }
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

/// 将 OOXML 样式描述映射为 rust_xlsxwriter 写入格式。
///
/// `full` = mode 1/2（含边框）；Basic = mode 3（font/fill/align/numfmt，无边框）。
/// 把 OOXML 颜色（8 位 ARGB 或 6 位 RGB）转成 rust_xlsxwriter 可用的 6 位 RGB。
///
/// rust_xlsxwriter 的 `Color::from(&str)` 用 `u32::from_str_radix` 解析，
/// 8 位 ARGB（如 `FF000000`）会超过 `0xFFFFFF` 上限，被 `is_valid()` 判为越界
/// 并**静默丢弃颜色**，同时在 stderr 打印告警。因此写入前统一截掉 alpha。
fn writer_rgb6(color: &str) -> Option<String> {
    let value = color.trim().trim_start_matches('#');
    if value.len() == 8 {
        let rgb = &value[2..];
        if rgb.chars().all(|ch| ch.is_ascii_hexdigit()) {
            return Some(format!("#{rgb}"));
        }
        return None;
    }
    if value.len() == 6 && value.chars().all(|ch| ch.is_ascii_hexdigit()) {
        return Some(format!("#{value}"));
    }
    None
}

/// 把字符宽度写成与源/openpyxl **完全相同**的 width 值。
///
/// rust_xlsxwriter 的 `set_column_width(w)` 走「字符宽 → 像素 → 再转回字符宽」
/// 往返换算（MDW=7、padding=5），写出的 width 会比输入大约 0.71
/// （12.0 → 12.7109375），与 openpyxl 直接写原始 width 的行为不一致。
/// 这里改走像素入口反算：像素 p 在存盘时按 `floor(p*256/7)/256` 还原成
/// 字符宽，因此先反解出能还原出目标 width 的像素值即可精确对齐。
fn set_column_width_exact(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    column_index: u32,
    width: f64,
) -> Result<()> {
    // 目标 width 的 1/256 量化值（Excel 列宽上限 255 字符 → 65280）。
    let target = (width * 256.0).round();
    anyhow::ensure!(
        (0.0..=65280.0).contains(&target),
        "static print column width {width} is out of range"
    );
    let target = target as u64;
    // 反解像素：floor(px*256/7) == target，px 落在 [target*7/256, (target+1)*7/256)。
    let low = (target * 7).div_ceil(256);
    let high = ((target + 1) * 7).div_ceil(256) - 1;
    let pixels = (low..=high).find(|candidate| (*candidate * 256 / 7) == target);
    let Some(pixels) = pixels else {
        // 目标宽度无法用像素精确表达时退回字符宽接口（存在 0.71 的取整差）。
        worksheet.set_column_width(
            u16::try_from(column_index.saturating_sub(1))
                .context("static print column width index is out of range")?,
            width,
        )?;
        return Ok(());
    };
    worksheet.set_column_width_pixels(
        u16::try_from(column_index.saturating_sub(1))
            .context("static print column width index is out of range")?,
        u32::try_from(pixels).unwrap_or(u32::MAX),
    )?;
    Ok(())
}

fn build_writer_format(
    descriptor: &crate::ooxml_cell_styles::CellStyleInfo,
    full: bool,
) -> rust_xlsxwriter::Format {
    use rust_xlsxwriter::FormatAlign;
    use rust_xlsxwriter::FormatBorder;
    use rust_xlsxwriter::FormatUnderline;

    let mut format = Format::new();
    if let Some(num_fmt) = &descriptor.num_fmt {
        format = format.set_num_format(num_fmt.as_str());
    }
    if let Some(font) = &descriptor.font {
        if font.bold {
            format = format.set_bold();
        }
        if font.italic {
            format = format.set_italic();
        }
        if font.strike {
            format = format.set_font_strikethrough();
        }
        if let Some(size) = &font.size {
            if let Ok(size) = size.parse::<f64>() {
                format = format.set_font_size(size);
            }
        }
        if let Some(name) = &font.name {
            format = format.set_font_name(name.as_str());
        }
        if let Some(color) = &font.color {
            // OOXML 颜色是 8 位 ARGB；rust_xlsxwriter 只接受 6 位 RGB
            // （0x000000..0xFFFFFF），传 8 位会被判越界并**静默丢弃颜色**。
            // 这里截掉 alpha（Excel 写出时同样补 FF 前缀）。
            if let Some(rgb) = writer_rgb6(color) {
                format = format.set_font_color(rgb.as_str());
            }
        }
        if let Some(style) = &font.underline {
            let underline = match style.as_str() {
                "double" => FormatUnderline::Double,
                "singleAccounting" => FormatUnderline::SingleAccounting,
                "doubleAccounting" => FormatUnderline::DoubleAccounting,
                _ => FormatUnderline::Single,
            };
            format = format.set_underline(underline);
        }
    }
    if let Some(fill) = &descriptor.fill {
        if fill.pattern.as_deref() == Some("solid") {
            if let Some(color) = &fill.fg_color {
                if let Some(rgb) = writer_rgb6(color) {
                    format = format.set_background_color(rgb.as_str());
                }
            }
        }
    }
    if let Some(align) = &descriptor.align {
        if let Some(horizontal) = &align.horizontal {
            let align_value = match horizontal.as_str() {
                "left" => FormatAlign::Left,
                "right" => FormatAlign::Right,
                "center" | "centerContinuous" => FormatAlign::CenterAcross,
                "fill" => FormatAlign::Fill,
                "justify" => FormatAlign::Justify,
                "distributed" => FormatAlign::Distributed,
                _ => FormatAlign::General,
            };
            if align_value != FormatAlign::General {
                format = format.set_align(align_value);
            }
        }
        if let Some(vertical) = &align.vertical {
            let align_value = match vertical.as_str() {
                "top" => FormatAlign::Top,
                "bottom" => FormatAlign::Bottom,
                "center" => FormatAlign::VerticalCenter,
                "justify" => FormatAlign::Justify,
                "distributed" => FormatAlign::VerticalDistributed,
                _ => FormatAlign::General,
            };
            if align_value != FormatAlign::General {
                format = format.set_align(align_value);
            }
        }
        if align.wrap_text {
            format = format.set_text_wrap();
        }
    }
    if full {
        if let Some(border) = &descriptor.border {
            fn apply_edge(
                mut format: rust_xlsxwriter::Format,
                edge: &Option<(Option<String>, Option<String>)>,
                setter: fn(rust_xlsxwriter::Format, FormatBorder) -> rust_xlsxwriter::Format,
            ) -> rust_xlsxwriter::Format {
                let Some((style, color)) = edge else {
                    return format;
                };
                let border_kind = match style.as_deref() {
                    Some("thin") => FormatBorder::Thin,
                    Some("medium") => FormatBorder::Medium,
                    Some("thick") => FormatBorder::Thick,
                    Some("dashed") => FormatBorder::Dashed,
                    Some("dotted") => FormatBorder::Dotted,
                    Some("double") => FormatBorder::Double,
                    Some("hair") => FormatBorder::Hair,
                    Some("mediumDashed") => FormatBorder::MediumDashed,
                    Some("dashDot") => FormatBorder::DashDot,
                    Some("mediumDashDot") => FormatBorder::MediumDashDot,
                    Some("dashDotDot") => FormatBorder::DashDotDot,
                    Some("mediumDashDotDot") => FormatBorder::MediumDashDotDot,
                    Some("slantDashDot") => FormatBorder::SlantDashDot,
                    _ => return format,
                };
                setter(format, border_kind)
            }
            format = apply_edge(
                format,
                &border.left,
                rust_xlsxwriter::Format::set_border_left,
            );
            format = apply_edge(
                format,
                &border.right,
                rust_xlsxwriter::Format::set_border_right,
            );
            format = apply_edge(format, &border.top, rust_xlsxwriter::Format::set_border_top);
            format = apply_edge(
                format,
                &border.bottom,
                rust_xlsxwriter::Format::set_border_bottom,
            );
        }
    }
    format
}

fn write_static_print_cell(
    worksheet: &mut Worksheet,
    row: usize,
    column: usize,
    value: &serde_json::Value,
    cell_format: Option<&rust_xlsxwriter::Format>,
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
            match cell_format {
                Some(format) => {
                    worksheet.write_string_with_format(
                        row,
                        column,
                        data.as_str().unwrap_or_default(),
                        format,
                    );
                }
                None => {
                    worksheet.write_string(row, column, data.as_str().unwrap_or_default())?;
                }
            };
        }
        "integer" => {
            let integer = data.as_i64().context("static print integer is invalid")?;
            let number = integer as f64;
            anyhow::ensure!(
                format!("{number:.0}").parse::<i64>().ok() == Some(integer),
                "integer {integer} cannot be represented exactly as an XLSX numeric cell"
            );
            match cell_format {
                Some(format) => {
                    worksheet.write_number_with_format(row, column, number, format);
                }
                None => {
                    worksheet.write_number(row, column, number)?;
                }
            };
        }
        "number" => {
            let number = data.as_f64().context("static print number is invalid")?;
            anyhow::ensure!(number.is_finite(), "static print number is not finite");
            match cell_format {
                Some(format) => {
                    worksheet.write_number_with_format(row, column, number, format);
                }
                None => {
                    worksheet.write_number(row, column, number)?;
                }
            };
        }
        "datetime_serial" => {
            let number = data
                .as_f64()
                .context("static print date serial is invalid")?;
            anyhow::ensure!(number.is_finite(), "static print date serial is not finite");
            let format = match cell_format {
                // 源单元格自带数字格式（OOXML 读取）：保持源显示口径。
                Some(format) => format.clone(),
                None => Format::new().set_num_format("yyyy-mm-dd h:mm:ss"),
            };
            worksheet.write_number_with_format(row, column, number, &format)?;
        }
        "boolean" => {
            match cell_format {
                Some(format) => {
                    worksheet.write_boolean_with_format(
                        row,
                        column,
                        data.as_bool().context("static print boolean is invalid")?,
                        format,
                    );
                }
                None => {
                    worksheet.write_boolean(
                        row,
                        column,
                        data.as_bool().context("static print boolean is invalid")?,
                    )?;
                }
            };
        }
        "error" => {
            // 原 Python 打印工具原样输出错误值（如 #DIV/0!）。xlsxwriter 无
            // 类型化错误写入 API，按旧版可见文本写入，保持与源表显示一致。
            match cell_format {
                Some(format) => {
                    worksheet.write_string_with_format(
                        row,
                        column,
                        data.as_str().unwrap_or("#N/A"),
                        format,
                    );
                }
                None => {
                    worksheet.write_string(row, column, data.as_str().unwrap_or("#N/A"))?;
                }
            };
        }
        "blank" => {
            match cell_format {
                Some(format) => {
                    worksheet.write_blank(row, column, format)?;
                }
                None => {}
            };
        }
        other => anyhow::bail!("unsupported static print cell type {other:?}"),
    }
    Ok(())
}

/// 单元格显示文本（用于合并区锚点；数字/日期/错误值都可取到字符串）。
fn display_text_of_value(value: &serde_json::Value) -> String {
    if let Some(text) = value.as_str() {
        return text.to_owned();
    }
    if let Some(number) = value.as_f64() {
        if number.fract() == 0.0 && number.abs() < 1e15 {
            return format!("{number:.0}");
        }
        return number.to_string();
    }
    if let Some(boolean) = value.as_bool() {
        return if boolean { "TRUE" } else { "FALSE" }.to_owned();
    }
    String::new()
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

/// 判断数字格式是否为纯时间格式（hh:mm:ss、mm:ss、[h]:mm 等）。
///
/// 纯时间串行值是「天的小数」，在 1900/1904 两种日期系统下显示一致；
/// 不做跨纪元平移，使值与 openpyxl 引擎（openpyxl 按源簿纪元读 time 单元格）一致。
/// 规则：清洗掉引号字面量、反斜杠转义与普通 [括号] 段后，
/// 含时间记号（h/s 或保留的 [h]/[m]/[mm]/[ss]）且不含日期记号（y/d）。
fn numfmt_is_time_only(format_code: &str) -> bool {
    let mut cleaned = String::with_capacity(format_code.len());
    let mut has_elapsed_token = false;
    let mut chars = format_code.chars().peekable();
    while let Some(ch) = chars.next() {
        match ch {
            // 引号字面量整段丢弃
            '"' => {
                for skipped in chars.by_ref() {
                    if skipped == '"' {
                        break;
                    }
                }
            }
            // 反斜杠转义的单个字面量（如 yyyy\-mm\-dd）丢弃
            '\\' => {
                chars.next();
            }
            // [括号] 段：颜色/条件（[Red]、[DBNum1]、[$-409]）丢弃；
            // 经过时间记号（[h]、[m]、[mm]、[ss]）保留记号字母
            '[' => {
                let mut section = String::new();
                for skipped in chars.by_ref() {
                    if skipped == ']' {
                        break;
                    }
                    section.push(skipped);
                }
                let lowered = section.to_ascii_lowercase();
                if lowered.starts_with('h') || lowered.starts_with('m') || lowered.starts_with('s')
                {
                    cleaned.push_str(&lowered);
                    has_elapsed_token = true;
                }
            }
            _ => cleaned.push(ch),
        }
    }
    let lowered = cleaned.to_ascii_lowercase();
    // 单独的 m 是月份记号；时间记号要求 h/s（或保留的经过时间括号段）
    let has_time = lowered.contains('h') || lowered.contains('s') || has_elapsed_token;
    let has_date = lowered.contains('y') || lowered.contains('d');
    has_time && !has_date
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
            // 原 PyTools print_310.py 不拒绝错误值，直接输出其显示文本；
            // 这里仅记录性放行，保持与旧工具一致的可用性。
            let _ = (address, value);
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
    fn static_print_allows_excel_error_cells_in_selected_ranges() {
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

        // 原 PyTools print_310.py 不拒绝错误值，直接输出其显示文本；
        // 因此这里要求放行而不是报错。
        validate_static_print_values(&sheet, &[(1, 2, 3, 4)])
            .expect("Excel error cells must not block static print output");
    }

    #[test]
    fn time_only_numfmt_detection() {
        assert!(numfmt_is_time_only("hh:mm:ss"));
        assert!(numfmt_is_time_only("h:mm"));
        assert!(numfmt_is_time_only("mm:ss"));
        assert!(numfmt_is_time_only("[h]:mm"));
        assert!(numfmt_is_time_only("[mm]"));
        // 引号/转义字面量不参与记号判断
        assert!(numfmt_is_time_only("h\"时\"mm\"分\""));
        // 日期/日期时间格式不做跨纪元豁免
        assert!(!numfmt_is_time_only("yyyy\\-mm\\-dd"));
        assert!(!numfmt_is_time_only("yyyy-mm-dd h:mm:ss"));
        assert!(!numfmt_is_time_only("dd/mm/yyyy"));
        assert!(!numfmt_is_time_only("m/d/yy h:mm"));
        // 颜色/条件括号段不是时间记号；月份 mm 不是时间
        assert!(!numfmt_is_time_only("[Red]mm"));
        assert!(!numfmt_is_time_only("mm"));
        assert!(!numfmt_is_time_only("0.00%"));
    }
}
