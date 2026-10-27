//! Rust implementation of the compact timeline extraction workflow (PT-1-2).

use crate::{
    config::{GlobalConfig, PathMapRule, TimelineRule},
    dedup_ops::{AppendCell, AppendTargetPlan, NewAppendSheet, NewAppendWorkbook},
    rule_engine::{extract_cells_from_matrix, pick_rules_for_workbook, ExtractedCell},
    workbook_extract::{matrix_from_dump, WorkbookCell},
};
use anyhow::{anyhow, Context, Result};
use chrono::{Local, NaiveDate, NaiveDateTime};
use excel_editor::{
    inspect_xlsx_profile, CellValue, EditOperation, ExcelEditPlan, PreservationRequirement,
    RangeStyleMode, WorkbookFormat,
};
use excel_reader::{
    read_value_columns_with_iso_dates, read_workbook, CellValue as ReaderCellValue,
};
use excel_writer::{Format, Workbook, Worksheet};
use rayon::prelude::*;
use serde::Serialize;
use std::{
    collections::{BTreeMap, BTreeSet, HashSet},
    fs::{self, OpenOptions},
    path::{Path, PathBuf},
    sync::mpsc,
    time::{Duration, Instant},
};

const OUTPUT_SHEET: &str = "时序提取结果";
const OUTPUT_HEADERS: [&str; 7] = [
    "源文件",
    "工作表名",
    "规则名称",
    "数据日期",
    "行头路径",
    "列头路径",
    "数值",
];

#[derive(Clone, Debug, Default, Serialize)]
pub struct TimelineExtractProgress {
    pub total_inputs: usize,
    pub processed_inputs: usize,
    pub completed_inputs: usize,
    pub failed_inputs: usize,
    pub current_input: Option<usize>,
    pub matched_rule_sheets: usize,
    pub extracted_rows: usize,
    pub output_rows: usize,
}

#[derive(Clone, Debug, Serialize)]
pub struct TimelineExtractRunEvent {
    pub sequence: usize,
    pub timestamp: String,
    pub feature_id: &'static str,
    pub phase: String,
    pub status: String,
    pub user_message: String,
    pub progress: TimelineExtractProgress,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub input_index: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stage_elapsed_ms: Option<u128>,
    pub total_elapsed_ms: u128,
    pub components: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub actual_engine: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_type: Option<String>,
}

#[derive(Clone, Debug, Serialize)]
pub struct TimelineExtractResult {
    pub total_files: usize,
    pub processed_files: usize,
    pub files: usize,
    pub failed_files: usize,
    pub matched_rule_sheets: usize,
    pub rows: usize,
    pub configured_target_rules: usize,
    pub target_writes_deferred: bool,
    pub path: Option<PathBuf>,
    #[serde(skip)]
    pub target_append_plans: Vec<TimelineTargetAppendPlan>,
    pub read_elapsed_ms: u128,
    pub write_elapsed_ms: u128,
    pub total_elapsed_ms: u128,
}

#[derive(Clone, Debug)]
pub struct TimelineTargetAppendPlan {
    pub plan: AppendTargetPlan,
    pub requested_rows: usize,
    /// True when configured rows could not be safely applied in the current
    /// run and need a later, explicitly authorized execution.
    pub deferred: bool,
}

/// Compatibility entry point for callers of the original timeline discovery API.
pub fn discover_timeline_extract_sources(global: &GlobalConfig) -> Result<Vec<PathBuf>> {
    crate::source_discovery::discover_configured_source_files(global)
}

struct FileExtraction {
    cells: Vec<ExtractedCell<WorkbookCell>>,
    matched_rule_sheets: usize,
    elapsed: Duration,
}

struct EventEmitter<'a> {
    started: Instant,
    sequence: usize,
    progress: TimelineExtractProgress,
    callback: Option<&'a mut dyn FnMut(TimelineExtractRunEvent)>,
}

impl EventEmitter<'_> {
    fn emit(
        &mut self,
        phase: &str,
        status: &str,
        message: impl Into<String>,
        input_index: Option<usize>,
        elapsed: Option<Duration>,
        components: &[&str],
        engine: Option<&str>,
        error_type: Option<&str>,
    ) {
        let Some(callback) = self.callback.as_mut() else {
            return;
        };
        self.sequence += 1;
        let mut recorded_components = vec!["Rust".to_owned()];
        recorded_components.extend(components.iter().map(|component| (*component).to_owned()));
        callback(TimelineExtractRunEvent {
            sequence: self.sequence,
            timestamp: Local::now().to_rfc3339(),
            feature_id: "PT-1-2",
            phase: phase.to_owned(),
            status: status.to_owned(),
            user_message: message.into(),
            progress: self.progress.clone(),
            input_index,
            stage_elapsed_ms: elapsed.map(|value| value.as_millis()),
            total_elapsed_ms: self.started.elapsed().as_millis(),
            components: recorded_components,
            actual_engine: engine.map(str::to_owned),
            error_type: error_type.map(str::to_owned),
        });
    }
}

fn record_file_extraction(
    emitter: &mut EventEmitter<'_>,
    index: usize,
    total: usize,
    result: &Result<FileExtraction>,
) {
    emitter.progress.processed_inputs += 1;
    emitter.progress.current_input = Some(index + 1);
    match result {
        Ok(file) => {
            emitter.progress.completed_inputs += 1;
            emitter.progress.matched_rule_sheets += file.matched_rule_sheets;
            emitter.progress.extracted_rows += file.cells.len();
            emitter.emit(
                "workbook_extract",
                "completed",
                format!(
                    "已处理 {}/{} 个工作簿，匹配 {} 个工作表",
                    emitter.progress.processed_inputs, total, file.matched_rule_sheets
                ),
                Some(index + 1),
                Some(file.elapsed),
                &["pytools-core", "excel-reader", "calamine", "rule_engine"],
                Some("calamine"),
                None,
            );
        }
        Err(_) => {
            emitter.progress.failed_inputs += 1;
            emitter.emit(
                "workbook_extract",
                "failed",
                format!("第 {}/{} 个工作簿读取或抽取失败", index + 1, total),
                Some(index + 1),
                None,
                &["pytools-core", "excel-reader", "calamine", "rule_engine"],
                Some("calamine"),
                Some("workbook_extract_error"),
            );
        }
    }
}

/// Extract configured row/column paths from supported source workbooks and
/// write the compact seven-column result used by the original `run_timeline_slim`.
pub fn run_timeline_extract(
    source_paths: &[PathBuf],
    rules: &[TimelineRule],
    path_maps: &[PathMapRule],
    global: &GlobalConfig,
) -> Result<TimelineExtractResult> {
    run_timeline_extract_inner(source_paths, rules, path_maps, global, None)
}

pub fn run_timeline_extract_with_progress<F>(
    source_paths: &[PathBuf],
    rules: &[TimelineRule],
    path_maps: &[PathMapRule],
    global: &GlobalConfig,
    mut on_event: F,
) -> Result<TimelineExtractResult>
where
    F: FnMut(TimelineExtractRunEvent),
{
    run_timeline_extract_inner(source_paths, rules, path_maps, global, Some(&mut on_event))
}

fn run_timeline_extract_inner(
    source_paths: &[PathBuf],
    rules: &[TimelineRule],
    path_maps: &[PathMapRule],
    global: &GlobalConfig,
    callback: Option<&mut dyn FnMut(TimelineExtractRunEvent)>,
) -> Result<TimelineExtractResult> {
    let started = Instant::now();
    let supported_extensions = global
        .source_exts
        .iter()
        .map(|value| normalize_extension(value))
        .chain(std::iter::once(".xls".to_owned()))
        .collect::<HashSet<_>>();
    let sources = source_paths
        .iter()
        .filter(|path| {
            path.is_file()
                && path
                    .extension()
                    .map(|extension| {
                        supported_extensions
                            .contains(&format!(".{}", extension.to_string_lossy().to_lowercase()))
                    })
                    .unwrap_or(false)
        })
        .cloned()
        .collect::<Vec<_>>();
    let mut emitter = EventEmitter {
        started,
        sequence: 0,
        progress: TimelineExtractProgress {
            total_inputs: sources.len(),
            ..TimelineExtractProgress::default()
        },
        callback,
    };
    emitter.emit(
        "run",
        "started",
        format!(
            "开始时序提取：{} 个输入，{} 条启用规则",
            sources.len(),
            rules.iter().filter(|rule| rule.enabled).count()
        ),
        None,
        None,
        &["pytools-core"],
        None,
        None,
    );

    let read_started = Instant::now();
    let source_count = sources.len();
    let mut results = (0..sources.len())
        .map(|_| None)
        .collect::<Vec<Option<Result<FileExtraction>>>>();
    if global.error_policy == "fail_fast" {
        for (index, path) in sources.iter().enumerate() {
            let result = extract_one_file(path, rules, path_maps, &global.timeline_rule_match_mode);
            record_file_extraction(&mut emitter, index, source_count, &result);
            let failed = result.is_err();
            results[index] = Some(result);
            if failed {
                break;
            }
        }
    } else {
        let (sender, receiver) = mpsc::channel::<(usize, Result<FileExtraction>)>();
        std::thread::scope(|scope| {
            let worker_sender = sender.clone();
            let sources_ref = &sources;
            let rules_ref = rules;
            let path_maps_ref = path_maps;
            let rule_match_mode = global.timeline_rule_match_mode.as_str();
            scope.spawn(move || {
                sources_ref
                    .par_iter()
                    .enumerate()
                    .for_each(|(index, path)| {
                        let result =
                            extract_one_file(path, rules_ref, path_maps_ref, rule_match_mode);
                        let _ = worker_sender.send((index, result));
                    });
            });
            drop(sender);
            for (index, result) in receiver {
                record_file_extraction(&mut emitter, index, source_count, &result);
                results[index] = Some(result);
            }
        });
    }
    let read_elapsed = read_started.elapsed();

    let mut cells = Vec::new();
    let mut matched_rule_sheets = 0usize;
    let mut first_error = None;
    let mut seen = HashSet::<(String, String, String, String, String, String)>::new();
    for (index, result) in results.into_iter().enumerate() {
        let Some(result) = result else {
            first_error.get_or_insert_with(|| {
                anyhow!("workbook extraction worker did not return a result")
            });
            continue;
        };
        match result {
            Ok(file) => {
                matched_rule_sheets += file.matched_rule_sheets;
                for cell in file.cells {
                    let key = (
                        cell.source_file.clone(),
                        cell.sheet_name.clone(),
                        cell.data_date.clone(),
                        cell.row_path.clone(),
                        cell.col_path.clone(),
                        python_value_text(&cell.value),
                    );
                    if seen.insert(key) {
                        cells.push(cell);
                    }
                }
            }
            Err(error) => {
                first_error.get_or_insert_with(|| {
                    error.context(format!(
                        "process source workbook {}",
                        sources[index].display()
                    ))
                });
            }
        }
    }
    if global.error_policy == "fail_fast" {
        if let Some(error) = first_error {
            emitter.emit(
                "run",
                "failed",
                "时序提取失败",
                None,
                Some(read_elapsed),
                &["pytools-core"],
                None,
                Some("workbook_extract_error"),
            );
            return Err(error);
        }
    }

    let write_started = Instant::now();
    let path = if cells.is_empty() {
        None
    } else {
        emitter.progress.output_rows = cells.len();
        emitter.emit(
            "workbook_write",
            "started",
            "正在写入时序提取结果",
            None,
            None,
            &["pytools-core", "excel-writer", "rust_xlsxwriter"],
            None,
            None,
        );
        Some(write_output_workbook(&cells, &global.output_dir)?)
    };
    let write_elapsed = write_started.elapsed();
    emitter.emit(
        "workbook_write",
        if path.is_some() {
            "completed"
        } else {
            "skipped"
        },
        if path.is_some() {
            format!("时序提取完成，输出 {} 行", cells.len())
        } else {
            "无匹配结果，不生成输出工作簿".to_owned()
        },
        None,
        Some(write_elapsed),
        &["pytools-core", "excel-writer", "rust_xlsxwriter"],
        path.as_ref().map(|_| "rust_xlsxwriter"),
        None,
    );
    let output_rule_names = cells
        .iter()
        .map(|cell| cell.rule_name.as_str())
        .collect::<HashSet<_>>();
    let configured_target_rules = rules
        .iter()
        .filter(|rule| {
            rule.target_write_enabled
                && rule.target_wb_path.is_some()
                && rule
                    .target_sheet
                    .as_ref()
                    .is_some_and(|sheet| !sheet.trim().is_empty())
                && output_rule_names.contains(rule.name.as_str())
        })
        .count();
    let target_append_plans = if cells.is_empty() {
        Vec::new()
    } else {
        build_timeline_target_append_plans(
            &cells,
            rules,
            &global.timeline_rule_match_mode,
            &global.output_dir,
            path.as_deref(),
        )?
    };
    let target_writes_deferred = target_append_plans.iter().any(|plan| plan.deferred);
    if !target_append_plans.is_empty() {
        let requested_rows = target_append_plans
            .iter()
            .map(|plan| plan.requested_rows)
            .sum::<usize>();
        emitter.emit(
            "target_append_plan",
            "completed",
            format!(
                "已准备 {} 个目标簿追加计划，涉及 {} 行",
                target_append_plans.len(),
                requested_rows
            ),
            None,
            None,
            &["pytools-core", "excel-reader", "calamine", "excel-editor"],
            Some("calamine"),
            None,
        );
    }
    let partial = emitter.progress.failed_inputs > 0 || target_writes_deferred;
    let end_message = if target_writes_deferred {
        format!(
            "时序提取生成 {} 行；配置的目标簿写入尚待安全提交",
            cells.len(),
        )
    } else if emitter.progress.failed_inputs > 0 {
        format!(
            "时序提取部分完成：{} 个工作簿失败，仍输出 {} 行",
            emitter.progress.failed_inputs,
            cells.len()
        )
    } else {
        format!("时序提取结束：输出 {} 行", cells.len())
    };
    emitter.emit(
        "run",
        if partial { "partial" } else { "completed" },
        end_message,
        None,
        Some(started.elapsed()),
        &["pytools-core"],
        None,
        None,
    );
    Ok(TimelineExtractResult {
        total_files: emitter.progress.total_inputs,
        processed_files: emitter.progress.processed_inputs,
        files: emitter.progress.completed_inputs,
        failed_files: emitter.progress.failed_inputs,
        matched_rule_sheets,
        rows: cells.len(),
        configured_target_rules,
        target_writes_deferred,
        path,
        target_append_plans,
        read_elapsed_ms: read_elapsed.as_millis(),
        write_elapsed_ms: write_elapsed.as_millis(),
        total_elapsed_ms: started.elapsed().as_millis(),
    })
}

fn extract_one_file(
    path: &Path,
    rules: &[TimelineRule],
    path_maps: &[PathMapRule],
    rule_match_mode: &str,
) -> Result<FileExtraction> {
    let started = Instant::now();
    let filename = path
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or_else(|| anyhow!("source workbook has no valid file name"))?;
    let matched_rules = pick_rules_for_workbook(rules, filename, rule_match_mode);
    if matched_rules.is_empty() {
        return Ok(FileExtraction {
            cells: Vec::new(),
            matched_rule_sheets: 0,
            elapsed: started.elapsed(),
        });
    }
    let workbook =
        read_workbook(path).with_context(|| format!("read source workbook {}", path.display()))?;
    let mut cells = Vec::new();
    let mut matched_rule_sheets = 0usize;
    for rule in matched_rules {
        for sheet in &workbook.sheets {
            if !crate::path_map::match_all_keywords(&sheet.name, &rule.sheet_keyword) {
                continue;
            }
            let (mut matrix, merges) = matrix_from_dump(sheet)?;
            clear_non_anchor_merge_values(&mut matrix, &merges);
            for row in &mut matrix {
                for value in row {
                    if matches!(value, Some(WorkbookCell::Error(_))) {
                        *value = None;
                    }
                }
            }
            let extracted = extract_cells_from_matrix(
                &matrix,
                &merges,
                filename,
                rule,
                &sheet.name,
                path_maps,
                true,
                &Default::default(),
            )
            .map_err(|error| anyhow!(error))
            .with_context(|| {
                format!(
                    "extract rule {:?} from worksheet {:?}",
                    rule.name, sheet.name
                )
            })?;
            if !extracted.is_empty() {
                matched_rule_sheets += 1;
                cells.extend(extracted);
            }
        }
    }
    Ok(FileExtraction {
        cells,
        matched_rule_sheets,
        elapsed: started.elapsed(),
    })
}

struct TargetSheetRows<'a> {
    name: String,
    rows: Vec<&'a ExtractedCell<WorkbookCell>>,
}

struct TargetWorkbookRows<'a> {
    path: PathBuf,
    sheets: Vec<TargetSheetRows<'a>>,
    requested_rows: usize,
}

fn build_timeline_target_append_plans(
    rows: &[ExtractedCell<WorkbookCell>],
    rules: &[TimelineRule],
    rule_match_mode: &str,
    output_dir: &Path,
    summary_path: Option<&Path>,
) -> Result<Vec<TimelineTargetAppendPlan>> {
    let mut grouped = Vec::<TargetWorkbookRows<'_>>::new();
    for row in rows {
        for rule in pick_rules_for_workbook(rules, &row.source_file, rule_match_mode)
            .into_iter()
            .filter(|rule| rule.name == row.rule_name && rule.target_write_enabled)
        {
            let (Some(target), Some(target_sheet)) = (&rule.target_wb_path, &rule.target_sheet)
            else {
                continue;
            };
            if target_sheet.trim().is_empty() {
                continue;
            }
            anyhow::ensure!(
                target
                    .extension()
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
                "PT-1-2 target append currently supports .xlsx targets only: {}",
                target.display()
            );
            if summary_path.is_some_and(|summary| same_path_lexically(target, summary)) {
                anyhow::bail!(
                    "PT-1-2 target workbook cannot be the generated summary workbook: {}",
                    target.display()
                );
            }

            let workbook_index = grouped
                .iter()
                .position(|workbook| workbook.path == *target)
                .unwrap_or_else(|| {
                    grouped.push(TargetWorkbookRows {
                        path: target.clone(),
                        sheets: Vec::new(),
                        requested_rows: 0,
                    });
                    grouped.len() - 1
                });
            let workbook = &mut grouped[workbook_index];
            workbook.requested_rows += 1;
            let sheet_index = workbook
                .sheets
                .iter()
                .position(|sheet| sheet.name == *target_sheet)
                .unwrap_or_else(|| {
                    workbook.sheets.push(TargetSheetRows {
                        name: target_sheet.clone(),
                        rows: Vec::new(),
                    });
                    workbook.sheets.len() - 1
                });
            workbook.sheets[sheet_index].rows.push(row);
        }
    }

    let mut plans = Vec::with_capacity(grouped.len());
    for target_rows in grouped {
        plans.push(build_timeline_target_append_plan(target_rows, output_dir)?);
    }
    Ok(plans)
}

fn build_timeline_target_append_plan(
    target_rows: TargetWorkbookRows<'_>,
    output_dir: &Path,
) -> Result<TimelineTargetAppendPlan> {
    let target = target_rows.path;
    let candidate = timeline_target_candidate_path(&target, output_dir);
    let metadata = fs::symlink_metadata(&target);
    let target_existed = match metadata {
        Ok(metadata) => {
            anyhow::ensure!(
                !metadata.file_type().is_symlink() && metadata.is_file(),
                "PT-1-2 target must be a regular non-symlink file: {}",
                target.display()
            );
            true
        }
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => false,
        Err(error) => {
            return Err(error)
                .with_context(|| format!("inspect PT-1-2 target {}", target.display()));
        }
    };
    let target_sha256 = target_existed
        .then(|| crate::file_commit::sha256_file(&target))
        .transpose()?;
    let mut plan = AppendTargetPlan {
        target: target.clone(),
        candidate: candidate.clone(),
        target_existed,
        target_sha256: target_sha256.clone(),
        task_ok: 0,
        task_skipped: 0,
        appended_rows: 0,
        duplicate_rows_filtered: 0,
        merged_tasks: target_rows.sheets.len().saturating_sub(1),
        issues: Vec::new(),
        edit_plan: None,
        new_workbook: None,
        replacement_workbook: None,
    };

    let (target_dump, target_values) = if target_existed {
        anyhow::ensure!(
            !excel_reader::uses_1904_date_system(&target)?,
            "PT-1-2 append to 1904-date-system targets is not validated; target left unchanged: {}",
            target.display()
        );
        let dump = read_workbook(&target)
            .with_context(|| format!("read PT-1-2 target workbook {}", target.display()))?;
        let values = read_value_columns_with_iso_dates(&target, 16_384)
            .with_context(|| format!("read PT-1-2 target values {}", target.display()))?;
        (Some(dump), values)
    } else {
        (None, Vec::new())
    };

    let mut replacement_sheets = Vec::<NewAppendSheet>::new();
    for target_sheet in &target_rows.sheets {
        let existing = target_values
            .iter()
            .find(|sheet| sheet.name == target_sheet.name);
        let existing_headers = existing.map(timeline_target_headers).unwrap_or_default();
        if existing.is_some()
            && OUTPUT_HEADERS[..6]
                .iter()
                .any(|required| !existing_headers.iter().any(|header| header == required))
        {
            plan.task_skipped += 1;
            plan.issues.push(format!(
                "target worksheet {:?}: header_mismatch; existing data was preserved",
                target_sheet.name
            ));
            continue;
        }

        let mut output_rows = vec![OUTPUT_HEADERS
            .iter()
            .map(|header| AppendCell::Value(CellValue::Text((*header).to_owned())))
            .collect::<Vec<_>>()];
        let mut seen_keys = HashSet::<Vec<String>>::new();
        if let Some(existing) = existing {
            for row_index in 2..=existing.max_row {
                let row = timeline_existing_target_row(existing, &existing_headers, row_index)?;
                seen_keys.insert(timeline_target_key_from_values(&row));
                output_rows.push(
                    row.iter()
                        .enumerate()
                        .map(|(column, value)| timeline_append_cell(value.as_ref(), column))
                        .collect::<Result<Vec<_>>>()?,
                );
            }
        }

        let mut batch_keys = HashSet::<Vec<String>>::new();
        let mut appended_rows = 0usize;
        for extracted in &target_sheet.rows {
            let key = timeline_new_target_key(extracted);
            if !batch_keys.insert(key.clone()) || seen_keys.contains(&key) {
                plan.duplicate_rows_filtered += 1;
                continue;
            }
            seen_keys.insert(key);
            output_rows.push(timeline_new_target_row(extracted)?);
            appended_rows += 1;
        }
        plan.task_ok += 1;
        if appended_rows == 0 {
            continue;
        }
        plan.appended_rows += appended_rows;
        replacement_sheets.push(NewAppendSheet {
            name: target_sheet.name.clone(),
            rows: output_rows,
            style_header_row: false,
            header_bold: true,
            data_date_column: Some(3),
        });
    }

    if target_existed && !replacement_sheets.is_empty() {
        let profile = inspect_xlsx_profile(&target)
            .with_context(|| format!("inspect PT-1-2 target workbook {}", target.display()))?;
        anyhow::ensure!(
            profile.inventory_complete,
            "PT-1-2 target OOXML inventory is incomplete; target left unchanged: {}",
            target.display()
        );
        let dump = target_dump.as_ref().expect("existing target was read");
        let staging_workbook = timeline_replacement_source_path(&candidate);
        let mut operations = Vec::<EditOperation>::new();
        let mut used_sheet_names = dump
            .sheets
            .iter()
            .map(|sheet| sheet.name.to_lowercase())
            .collect::<HashSet<_>>();
        let mut append_index = dump.sheets.len();
        for replacement in &replacement_sheets {
            let range = timeline_target_sheet_range(replacement)?;
            if let Some(index) = dump
                .sheets
                .iter()
                .position(|sheet| sheet.name == replacement.name)
            {
                let temporary_name = timeline_temporary_sheet_name(&mut used_sheet_names)?;
                operations.push(EditOperation::CreateSheetAt {
                    name: temporary_name.clone(),
                    index,
                });
                operations.push(timeline_import_target_sheet(
                    &staging_workbook,
                    &replacement.name,
                    &range,
                    &temporary_name,
                ));
                operations.push(EditOperation::DeleteSheet {
                    name: replacement.name.clone(),
                });
                operations.push(EditOperation::RenameSheet {
                    old_name: temporary_name,
                    new_name: replacement.name.clone(),
                });
            } else {
                operations.push(EditOperation::CreateSheetAt {
                    name: replacement.name.clone(),
                    index: append_index,
                });
                operations.push(timeline_import_target_sheet(
                    &staging_workbook,
                    &replacement.name,
                    &range,
                    &replacement.name,
                ));
                append_index += 1;
            }
        }
        let preserve = BTreeSet::from([
            PreservationRequirement::CellValuesAndTypes,
            PreservationRequirement::FormulaText,
            PreservationRequirement::CommentsAndPositions,
            PreservationRequirement::Styles,
            PreservationRequirement::MergedRanges,
            PreservationRequirement::ConditionalFormatting,
            PreservationRequirement::DataValidations,
            PreservationRequirement::DefinedNames,
            PreservationRequirement::RowAndColumnDimensions,
            PreservationRequirement::PageSettings,
            PreservationRequirement::HeaderFooterText,
            PreservationRequirement::Relationships,
            PreservationRequirement::SheetIdentityAndOrder,
        ]);
        let mut drop_allowed = BTreeMap::new();
        for replacement in &replacement_sheets {
            if let Some(old_sheet) = dump
                .sheets
                .iter()
                .find(|sheet| sheet.name == replacement.name)
            {
                let mut allowed = BTreeSet::from([
                    PreservationRequirement::FormulaText,
                    PreservationRequirement::CommentsAndPositions,
                    PreservationRequirement::Styles,
                    PreservationRequirement::MergedRanges,
                    PreservationRequirement::RowAndColumnDimensions,
                    PreservationRequirement::PageSettings,
                    PreservationRequirement::HeaderFooterText,
                    PreservationRequirement::Relationships,
                ]);
                if !old_sheet.conditional_formatting.is_empty() {
                    allowed.insert(PreservationRequirement::ConditionalFormatting);
                }
                if !old_sheet.data_validations.is_empty() {
                    allowed.insert(PreservationRequirement::DataValidations);
                }
                drop_allowed.insert(replacement.name.clone(), allowed);
            }
        }
        let edit_plan = ExcelEditPlan {
            request_id: format!(
                "pytools-pt-1-2-append-{}",
                target.file_stem().unwrap_or_default().to_string_lossy()
            ),
            input: target.clone(),
            output: candidate.clone(),
            format: WorkbookFormat::Xlsx,
            workbook: profile,
            operations,
            preserve,
            drop_allowed,
        };
        edit_plan
            .validate()
            .context("validate PT-1-2 target append edit plan")?;
        plan.edit_plan = Some(edit_plan);
        plan.replacement_workbook = Some(NewAppendWorkbook {
            sheets: replacement_sheets,
        });
    } else if !target_existed && !replacement_sheets.is_empty() {
        plan.new_workbook = Some(NewAppendWorkbook {
            sheets: replacement_sheets,
        });
    }

    crate::file_commit::verify_planning_snapshot(&target, target_sha256.as_deref())
        .with_context(|| format!("verify PT-1-2 target snapshot {}", target.display()))?;
    let deferred = plan.appended_rows > 0;
    Ok(TimelineTargetAppendPlan {
        plan,
        requested_rows: target_rows.requested_rows,
        deferred,
    })
}

fn timeline_target_headers(sheet: &excel_reader::SheetValues) -> Vec<String> {
    (1..=sheet.max_column)
        .map(|column| {
            sheet
                .values
                .get(&crate::rule_engine::cell_address(1, column))
                .map(timeline_reader_cell_text)
                .unwrap_or_else(|| format!("Unnamed: {}", column - 1))
        })
        .collect()
}

fn timeline_existing_target_row(
    sheet: &excel_reader::SheetValues,
    headers: &[String],
    row: usize,
) -> Result<Vec<Option<WorkbookCell>>> {
    OUTPUT_HEADERS
        .iter()
        .map(|header| {
            let Some(column) = headers.iter().position(|existing| existing == header) else {
                return Ok(None);
            };
            sheet
                .values
                .get(&crate::rule_engine::cell_address(row, column + 1))
                .map(timeline_reader_cell)
                .transpose()
        })
        .collect()
}

fn timeline_reader_cell(value: &ReaderCellValue) -> Result<WorkbookCell> {
    Ok(match value {
        ReaderCellValue::Empty => return Ok(WorkbookCell::Text(String::new())),
        ReaderCellValue::Text(value) => WorkbookCell::Text(value.clone()),
        ReaderCellValue::Number(value) => WorkbookCell::Number(*value),
        ReaderCellValue::Integer(value) => WorkbookCell::Integer(*value),
        ReaderCellValue::Boolean(value) => WorkbookCell::Boolean(*value),
        ReaderCellValue::DateTimeSerial(value) => WorkbookCell::DateTimeSerial(*value),
        ReaderCellValue::DateTimeIso(value) => WorkbookCell::DateTimeIso(value.clone()),
        ReaderCellValue::DurationIso(value) => WorkbookCell::DurationIso(value.clone()),
        ReaderCellValue::Error(value) => WorkbookCell::Error(value.clone()),
    })
}

fn timeline_reader_cell_text(value: &ReaderCellValue) -> String {
    match value {
        ReaderCellValue::Empty => String::new(),
        ReaderCellValue::Text(value)
        | ReaderCellValue::DurationIso(value)
        | ReaderCellValue::Error(value) => value.clone(),
        ReaderCellValue::DateTimeIso(value) => crate::summary_ops::python_datetime_iso_text(value),
        ReaderCellValue::DateTimeSerial(value) => {
            crate::summary_ops::python_datetime_serial_text(*value)
                .unwrap_or_else(|| crate::summary_ops::python_float_text(*value))
        }
        ReaderCellValue::Number(value) => crate::summary_ops::python_float_text(*value),
        ReaderCellValue::Integer(value) => value.to_string(),
        ReaderCellValue::Boolean(value) => if *value { "True" } else { "False" }.to_owned(),
    }
}

fn timeline_target_key_from_values(row: &[Option<WorkbookCell>]) -> Vec<String> {
    [0, 1, 3, 4, 5]
        .into_iter()
        .map(|index| {
            row.get(index)
                .and_then(Option::as_ref)
                .map(timeline_existing_key_cell)
                .unwrap_or_default()
        })
        .collect()
}

fn timeline_existing_key_cell(value: &WorkbookCell) -> String {
    match value {
        WorkbookCell::Text(value) | WorkbookCell::Error(value) => value.trim().to_owned(),
        WorkbookCell::Integer(value) => value.to_string(),
        WorkbookCell::Number(value) => timeline_python_real_text(*value),
        WorkbookCell::Boolean(value) => if *value { "True" } else { "False" }.to_owned(),
        WorkbookCell::DateTimeSerial(value) => {
            crate::summary_ops::python_datetime_serial_text(*value)
                .map(|value| timeline_datetime_dedup_key(&value))
                .unwrap_or_else(|| crate::summary_ops::python_float_text(*value))
        }
        WorkbookCell::DurationSerial(value) => crate::summary_ops::python_float_text(*value),
        WorkbookCell::DateTimeIso(value) => {
            timeline_datetime_dedup_key(&crate::summary_ops::python_datetime_iso_text(value))
        }
        WorkbookCell::DurationIso(value) => value.clone(),
    }
}

fn timeline_datetime_dedup_key(value: &str) -> String {
    NaiveDateTime::parse_from_str(value, "%Y-%m-%d %H:%M:%S%.f")
        .or_else(|_| NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S%.f"))
        .map(|datetime| {
            if datetime.time() == chrono::NaiveTime::MIN {
                datetime.format("%Y-%m-%d").to_string()
            } else {
                datetime.format("%Y-%m-%d %H:%M:%S").to_string()
            }
        })
        .unwrap_or_else(|_| value.replace('T', " "))
}

fn timeline_python_real_text(value: f64) -> String {
    if value.is_finite()
        && value.fract() == 0.0
        && value >= i64::MIN as f64
        && value <= i64::MAX as f64
    {
        (value as i64).to_string()
    } else {
        format!("{value:.15}")
            .trim_end_matches('0')
            .trim_end_matches('.')
            .to_owned()
    }
}

fn timeline_new_target_key(cell: &ExtractedCell<WorkbookCell>) -> Vec<String> {
    let normalized_date = NaiveDate::parse_from_str(&cell.data_date, "%Y-%m-%d")
        .map(|date| date.format("%Y-%m-%d").to_string())
        .unwrap_or_else(|_| cell.data_date.trim().to_owned());
    vec![
        cell.source_file.trim().to_owned(),
        cell.sheet_name.trim().to_owned(),
        normalized_date,
        cell.row_path.trim().to_owned(),
        cell.col_path.trim().to_owned(),
    ]
}

fn timeline_new_target_row(cell: &ExtractedCell<WorkbookCell>) -> Result<Vec<AppendCell>> {
    let mut output = vec![
        AppendCell::Value(CellValue::Text(cell.source_file.clone())),
        AppendCell::Value(CellValue::Text(cell.sheet_name.clone())),
        AppendCell::Value(CellValue::Text(cell.rule_name.clone())),
        timeline_date_append_cell(&cell.data_date),
        AppendCell::Value(CellValue::Text(cell.row_path.clone())),
        AppendCell::Value(CellValue::Text(cell.col_path.clone())),
    ];
    output.push(timeline_append_cell(Some(&cell.value), 6)?);
    Ok(output)
}

fn timeline_date_append_cell(value: &str) -> AppendCell {
    if value.is_empty() {
        return AppendCell::Value(CellValue::Blank);
    }
    date_to_excel_serial(value)
        .map(|serial| AppendCell::DateTime {
            serial: serial.to_string(),
            number_format: "yyyy/m/d".to_owned(),
        })
        .unwrap_or_else(|| AppendCell::Value(CellValue::Text(value.to_owned())))
}

fn timeline_append_cell(value: Option<&WorkbookCell>, column: usize) -> Result<AppendCell> {
    let Some(value) = value else {
        return Ok(AppendCell::Value(CellValue::Blank));
    };
    let number_format = if column == 3 {
        "yyyy/m/d"
    } else {
        "YYYY-MM-DD HH:MM:SS"
    };
    Ok(match value {
        WorkbookCell::Text(value) if value.is_empty() => AppendCell::Value(CellValue::Blank),
        WorkbookCell::Text(value) | WorkbookCell::Error(value) => {
            AppendCell::Value(CellValue::Text(value.clone()))
        }
        WorkbookCell::Integer(value) => AppendCell::Value(CellValue::Decimal(value.to_string())),
        WorkbookCell::Number(value) => AppendCell::Value(CellValue::Decimal(value.to_string())),
        WorkbookCell::Boolean(value) => AppendCell::Value(CellValue::Boolean(*value)),
        WorkbookCell::DateTimeSerial(serial) => AppendCell::DateTime {
            serial: serial.to_string(),
            number_format: number_format.to_owned(),
        },
        WorkbookCell::DurationSerial(serial) => AppendCell::Number {
            serial: serial.to_string(),
            number_format: "0".to_owned(),
        },
        WorkbookCell::DateTimeIso(value) => {
            let serial = datetime_to_excel_serial(value)
                .with_context(|| format!("unsupported PT-1-2 target date {value:?}"))?;
            AppendCell::DateTime {
                serial: serial.to_string(),
                number_format: number_format.to_owned(),
            }
        }
        WorkbookCell::DurationIso(value) => {
            let serial = iso_duration_serial(value)
                .with_context(|| format!("unsupported PT-1-2 target duration {value:?}"))?;
            AppendCell::DateTime {
                serial: serial.to_string(),
                number_format: "[h]:mm:ss.000".to_owned(),
            }
        }
    })
}

fn timeline_target_sheet_range(sheet: &NewAppendSheet) -> Result<String> {
    anyhow::ensure!(
        !sheet.rows.is_empty(),
        "PT-1-2 replacement worksheet has no rows"
    );
    let columns = sheet.rows.iter().map(Vec::len).max().unwrap_or(0);
    anyhow::ensure!(columns > 0, "PT-1-2 replacement worksheet has no columns");
    anyhow::ensure!(
        sheet.rows.iter().all(|row| row.len() == columns),
        "PT-1-2 replacement worksheet rows must have a consistent column count"
    );
    Ok(format!(
        "A1:{}",
        crate::rule_engine::cell_address(sheet.rows.len(), columns)
    ))
}

fn timeline_import_target_sheet(
    source_workbook: &Path,
    source_sheet: &str,
    source_range: &str,
    target_sheet: &str,
) -> EditOperation {
    EditOperation::ImportRange {
        source_workbook: source_workbook.to_path_buf(),
        source_sheet: source_sheet.to_owned(),
        source_range: source_range.to_owned(),
        target_sheet: target_sheet.to_owned(),
        target_start: "A1".to_owned(),
        copy_values: true,
        copy_formulas: false,
        style_mode: RangeStyleMode::Full,
        copy_comments: false,
        copy_merges: false,
        copy_row_heights: false,
        copy_column_widths: false,
        copy_row_hidden: false,
        copy_column_hidden: false,
        copy_outline_levels: false,
        copy_all_row_dimensions: false,
        copy_all_column_dimensions: false,
        materialize_blank_cells: true,
        hide_zero: false,
    }
}

fn timeline_temporary_sheet_name(used: &mut HashSet<String>) -> Result<String> {
    for suffix in 1..=10_000usize {
        let candidate = format!("__pt_replace_{suffix}");
        if used.insert(candidate.to_lowercase()) {
            return Ok(candidate);
        }
    }
    anyhow::bail!("could not allocate a PT-1-2 temporary worksheet name")
}

fn timeline_target_candidate_path(target: &Path, output_dir: &Path) -> PathBuf {
    let timestamp = Local::now().format("%Y%m%d_%H%M%S");
    let stem = target.file_stem().unwrap_or_default().to_string_lossy();
    for suffix in 0..10_000usize {
        let tail = if suffix == 0 {
            String::new()
        } else {
            format!("_{suffix}")
        };
        let candidate = output_dir.join(format!("{stem}_pt12_target_{timestamp}{tail}.xlsx"));
        if !candidate.exists() {
            return candidate;
        }
    }
    output_dir.join(format!(
        "{}_pt12_target_{}.xlsx",
        stem,
        chrono::Utc::now().timestamp_millis()
    ))
}

fn timeline_replacement_source_path(candidate: &Path) -> PathBuf {
    let stem = candidate.file_stem().unwrap_or_default().to_string_lossy();
    candidate.with_file_name(format!(".{stem}.replacement-source.xlsx"))
}

fn same_path_lexically(left: &Path, right: &Path) -> bool {
    let current = std::env::current_dir().ok();
    let absolute = |path: &Path| {
        if path.is_absolute() {
            path.to_path_buf()
        } else {
            current
                .as_ref()
                .map_or_else(|| path.to_path_buf(), |base| base.join(path))
        }
    };
    absolute(left) == absolute(right)
}

fn clear_non_anchor_merge_values(
    matrix: &mut [Vec<Option<WorkbookCell>>],
    merges: &[crate::rule_engine::MergeRange],
) {
    for merge in merges {
        for row in merge.min_row..=merge.max_row {
            for column in merge.min_col..=merge.max_col {
                if row == merge.min_row && column == merge.min_col {
                    continue;
                }
                if let Some(cell) = matrix.get_mut(row).and_then(|cells| cells.get_mut(column)) {
                    *cell = None;
                }
            }
        }
    }
}

fn normalize_extension(extension: &str) -> String {
    let extension = extension.trim().to_lowercase();
    if extension.starts_with('.') {
        extension
    } else {
        format!(".{extension}")
    }
}

fn python_value_text(value: &WorkbookCell) -> String {
    match value {
        WorkbookCell::Text(value)
        | WorkbookCell::Error(value)
        | WorkbookCell::DurationIso(value) => value.clone(),
        WorkbookCell::Integer(value) => value.to_string(),
        WorkbookCell::Number(value) => crate::summary_ops::python_float_text(*value),
        WorkbookCell::DurationSerial(value) => crate::summary_ops::python_float_text(*value),
        WorkbookCell::Boolean(value) => if *value { "True" } else { "False" }.to_owned(),
        WorkbookCell::DateTimeSerial(serial) => {
            crate::summary_ops::python_datetime_serial_text(*serial)
                .unwrap_or_else(|| crate::summary_ops::python_float_text(*serial))
        }
        WorkbookCell::DateTimeIso(value) => crate::summary_ops::python_datetime_iso_text(value),
    }
}

fn write_output_workbook(
    rows: &[ExtractedCell<WorkbookCell>],
    output_dir: &Path,
) -> Result<PathBuf> {
    fs::create_dir_all(output_dir)
        .with_context(|| format!("create timeline output directory {}", output_dir.display()))?;
    let output_path = reserve_output_path(output_dir)?;
    let write_result = (|| -> Result<()> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();
        worksheet.set_name(OUTPUT_SHEET)?;
        let header_format = Format::new().set_bold();
        for (column, header) in OUTPUT_HEADERS.iter().enumerate() {
            worksheet.write_string_with_format(0, column as u16, *header, &header_format)?;
        }
        for (index, cell) in rows.iter().enumerate() {
            let row = u32::try_from(index + 1).context("timeline output exceeds XLSX row limit")?;
            worksheet.write_string(row, 0, &cell.source_file)?;
            worksheet.write_string(row, 1, &cell.sheet_name)?;
            worksheet.write_string(row, 2, &cell.rule_name)?;
            write_data_date(worksheet, row, 3, &cell.data_date)?;
            worksheet.write_string(row, 4, &cell.row_path)?;
            worksheet.write_string(row, 5, &cell.col_path)?;
            write_value(worksheet, row, 6, &cell.value)?;
        }
        workbook.save(&output_path).with_context(|| {
            format!("save timeline extraction output {}", output_path.display())
        })?;
        Ok(())
    })();
    if let Err(error) = write_result {
        let _ = fs::remove_file(&output_path);
        return Err(error);
    }
    Ok(output_path)
}

fn reserve_output_path(output_dir: &Path) -> Result<PathBuf> {
    let timestamp = Local::now().format("%Y%m%d_%H%M%S");
    for suffix in 1..=10_000usize {
        let label = if suffix == 1 {
            String::new()
        } else {
            format!("_{}", suffix)
        };
        let candidate = output_dir.join(format!("{timestamp}_时序提取（快速精简）{label}.xlsx"));
        match OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&candidate)
        {
            Ok(_) => return Ok(candidate),
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => {
                return Err(error).with_context(|| {
                    format!("reserve timeline output path {}", candidate.display())
                });
            }
        }
    }
    Err(anyhow!("could not reserve a unique timeline output path"))
}

fn write_data_date(worksheet: &mut Worksheet, row: u32, column: u16, value: &str) -> Result<()> {
    let format = Format::new().set_num_format("yyyy/m/d");
    if let Some(serial) = date_to_excel_serial(value) {
        worksheet.write_number_with_format(row, column, serial, &format)?;
    } else if value.is_empty() {
        worksheet.write_blank(row, column, &format)?;
    } else {
        worksheet.write_string_with_format(row, column, value, &format)?;
    }
    Ok(())
}

fn write_value(
    worksheet: &mut Worksheet,
    row: u32,
    column: u16,
    value: &WorkbookCell,
) -> Result<()> {
    match value {
        WorkbookCell::Text(value) | WorkbookCell::Error(value) => {
            worksheet.write_string(row, column, value)?;
        }
        WorkbookCell::Integer(value) => {
            worksheet.write_number(row, column, *value as f64)?;
        }
        WorkbookCell::Number(value) => {
            worksheet.write_number(row, column, *value)?;
        }
        WorkbookCell::DurationSerial(value) => {
            let format = Format::new().set_num_format("0");
            worksheet.write_number_with_format(row, column, *value, &format)?;
        }
        WorkbookCell::Boolean(value) => {
            worksheet.write_boolean(row, column, *value)?;
        }
        WorkbookCell::DateTimeSerial(value) => {
            let format = Format::new().set_num_format("YYYY-MM-DD HH:MM:SS");
            worksheet.write_number_with_format(row, column, *value, &format)?;
        }
        WorkbookCell::DateTimeIso(value) => {
            if let Some(serial) = datetime_to_excel_serial(value) {
                let format = Format::new().set_num_format("YYYY-MM-DD HH:MM:SS");
                worksheet.write_number_with_format(row, column, serial, &format)?;
            } else {
                worksheet.write_string(row, column, value)?;
            }
        }
        WorkbookCell::DurationIso(value) => {
            if let Some(serial) = iso_duration_serial(value) {
                let format = Format::new().set_num_format("[h]:mm:ss.000");
                worksheet.write_number_with_format(row, column, serial, &format)?;
            } else {
                worksheet.write_string(row, column, value)?;
            }
        }
    }
    Ok(())
}

fn date_to_excel_serial(value: &str) -> Option<f64> {
    let date = NaiveDate::parse_from_str(value, "%Y-%m-%d").ok()?;
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 30)?;
    Some(date.signed_duration_since(epoch).num_days() as f64)
}

fn datetime_to_excel_serial(value: &str) -> Option<f64> {
    let datetime = chrono::DateTime::parse_from_rfc3339(value)
        .map(|value| value.naive_utc())
        .or_else(|_| NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S%.f"))
        .or_else(|_| NaiveDateTime::parse_from_str(value, "%Y-%m-%d %H:%M:%S%.f"))
        .or_else(|_| {
            NaiveDate::parse_from_str(value, "%Y-%m-%d")
                .map(|date| date.and_hms_opt(0, 0, 0).expect("valid midnight"))
        })
        .ok()?;
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 30)?.and_hms_opt(0, 0, 0)?;
    Some(datetime.signed_duration_since(epoch).num_milliseconds() as f64 / 86_400_000.0)
}

fn iso_duration_serial(value: &str) -> Option<f64> {
    let text = value.strip_prefix('P')?;
    let mut in_time = false;
    let mut number = String::new();
    let mut seconds = 0.0;
    for character in text.chars() {
        if character == 'T' {
            in_time = true;
            continue;
        }
        if character.is_ascii_digit() || character == '.' || character == '-' {
            number.push(character);
            continue;
        }
        let amount = number.parse::<f64>().ok()?;
        number.clear();
        let multiplier = match (in_time, character) {
            (false, 'D') => 86_400.0,
            (true, 'H') => 3_600.0,
            (true, 'M') => 60.0,
            (true, 'S') => 1.0,
            _ => return None,
        };
        seconds += amount * multiplier;
    }
    if number.is_empty() {
        Some(seconds / 86_400.0)
    } else {
        None
    }
}
