//! Candidate Rust implementation of pytools institution-name normalization.
//!
//! The workbook adapter currently accepts `.xlsx` only. Apply mode publishes a
//! validated copy beside each source; it never replaces the selected source.

use anyhow::{anyhow, bail, ensure, Context, Result};
use excel_editor::{
    CellValue as PlannedCellValue, EditOperation, ExcelEditPlan, PreservationRequirement,
    WorkbookFormat, WorkbookProfile,
};
use excel_reader::{
    read_defined_names, read_ooxml_comments, read_workbook,
    read_xlsx_effective_cell_number_formats, CellNumberFormatRequest, CellValue as ReaderCellValue,
    DefinedNameParseStatus, DefinedNameRecord, DefinedNameScope, SheetDump, WorkbookComment,
};
use excel_writer::{Format, Note, Workbook as NewWorkbook};
use rust_xlsxwriter::Formula;
use serde::Serialize;
use serde_json::Value;
use std::{
    collections::{BTreeMap, BTreeSet},
    fs::{self, OpenOptions},
    io::Write,
    path::{Path, PathBuf},
    time::{Instant, SystemTime, UNIX_EPOCH},
};

#[cfg(test)]
use excel_editor::{open_xlsx, save_xlsx};
#[cfg(test)]
use rayon::prelude::*;
#[cfg(test)]
use std::collections::HashSet;

const BRANCH_SHEET_KEY: &str = "分机构";
pub const INSTITUTION_NORMALIZE_OUTPUT_DIRECTORY: &str = "PyTools_机构名称规范化结果";

#[cfg(test)]
#[derive(Debug, Default, Clone, PartialEq, Eq, Serialize)]
pub struct NormalizeReport {
    pub files: usize,
    pub branch_sheets: usize,
    pub mapped_sheets: usize,
    pub mapped_rows: usize,
    pub failed_files: usize,
    pub mapping_empty: bool,
    pub errors: Vec<String>,
}

#[derive(Debug, Default, Clone, PartialEq)]
pub struct NormalizePlanResult {
    pub branch_sheets: usize,
    pub mapped_sheets: usize,
    pub mapped_rows: usize,
    pub mapping_empty: bool,
    pub plan: Option<ExcelEditPlan>,
}

#[derive(Debug, Clone)]
pub struct InstitutionNormalizeRunOptions {
    pub apply: bool,
    pub capabilities: Option<PathBuf>,
    pub platform: Option<String>,
    pub plan_dir: Option<PathBuf>,
    /// 就地修改：直接把规范化结果写回源工作簿（不再生成结果目录副本）。
    pub in_place: bool,
    /// 就地修改使用的引擎名称（"umya" 或 "openpyxl"）。
    pub engine: Option<String>,
}

#[derive(Debug, Clone, Default, Serialize)]
pub struct InstitutionNormalizeProgress {
    pub total_files: usize,
    pub current_file: Option<usize>,
    pub processed_files: usize,
    pub completed_files: usize,
    pub failed_files: usize,
    pub modified_files: usize,
}

#[derive(Debug, Clone, Serialize)]
pub struct InstitutionNormalizeEvent {
    pub sequence: usize,
    pub timestamp: String,
    pub phase: String,
    pub status: String,
    pub user_message: String,
    pub progress: InstitutionNormalizeProgress,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub file_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub elapsed_ms: Option<u128>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub components: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub actual_engine: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub diagnostic: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub route_diagnostic: Option<Value>,
}

#[derive(Debug, Clone, Default, Serialize)]
pub struct InstitutionNormalizeStageTimings {
    pub source_fingerprint_ms: u128,
    pub profile_inspection_ms: u128,
    pub plan_build_ms: u128,
    pub plan_persistence_ms: u128,
    pub source_stability_check_ms: u128,
    pub engine_execution_ms: u128,
    pub safe_commit_ms: u128,
    pub total_ms: u128,
}

#[derive(Debug, Clone, Serialize)]
pub struct InstitutionNormalizeFileResult {
    pub file_name: String,
    pub status: String,
    pub branch_sheets: usize,
    pub mapped_sheets: usize,
    pub mapped_rows: usize,
    pub operation_count: usize,
    pub plan_path: Option<PathBuf>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub output_path: Option<PathBuf>,
    pub actual_engine: Option<String>,
    pub stage_components: BTreeMap<String, Vec<String>>,
    pub execution: Option<Value>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub commit: Option<crate::file_commit::TargetCommitReport>,
    pub timings: InstitutionNormalizeStageTimings,
    pub error_stage: Option<String>,
    pub error_type: Option<String>,
    pub user_message: String,
    pub diagnostic: Option<String>,
    pub route_diagnostic: Option<Value>,
    pub cleanup_warning: Option<String>,
}

#[derive(Debug, Clone, Serialize)]
pub struct InstitutionNormalizeRunReport {
    pub schema_version: u32,
    pub product: String,
    pub feature_id: String,
    pub run_id: String,
    pub started_at: String,
    pub mode: String,
    pub status: String,
    pub total_inputs: usize,
    pub completed_files: usize,
    pub failed_files: usize,
    pub modified_files: usize,
    pub branch_sheets: usize,
    pub mapped_sheets: usize,
    pub mapped_rows: usize,
    pub mapping_empty: bool,
    pub configuration_elapsed_ms: u128,
    pub elapsed_ms: u128,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub failure: Option<InstitutionNormalizeRunFailure>,
    pub events: Vec<InstitutionNormalizeEvent>,
    pub files: Vec<InstitutionNormalizeFileResult>,
}

#[derive(Debug, Clone, Serialize)]
pub struct InstitutionNormalizeRunFailure {
    pub stage: String,
    pub error_type: String,
    pub user_message: String,
    pub diagnostic: String,
}

#[derive(Debug)]
struct InstitutionNormalizeFileFailure {
    stage: &'static str,
    error_type: &'static str,
    diagnostic: String,
    route_diagnostic: Option<Value>,
    actual_engine: Option<String>,
    execution: Option<Value>,
    timings: InstitutionNormalizeStageTimings,
    stage_components: BTreeMap<String, Vec<String>>,
}

impl InstitutionNormalizeFileFailure {
    fn attach_state(
        mut self,
        timings: &InstitutionNormalizeStageTimings,
        stage_components: &BTreeMap<String, Vec<String>>,
        actual_engine: Option<&str>,
    ) -> Self {
        self.timings = timings.clone();
        self.stage_components = stage_components.clone();
        self.actual_engine = actual_engine.map(str::to_owned);
        self
    }

    fn attach_execution(mut self, execution: Option<Value>) -> Self {
        self.execution = execution;
        self
    }
}

/// Run PT-7-2 against an explicit workbook list. Planning is the default. Apply
/// creates a new workbook copy with Rust; a supplied ExcelModifier capability
/// remains an explicit alternative route.
pub fn run_institution_normalize_files<F>(
    config_path: impl AsRef<Path>,
    inputs: &[PathBuf],
    options: InstitutionNormalizeRunOptions,
    mut on_event: F,
) -> Result<InstitutionNormalizeRunReport>
where
    F: FnMut(&InstitutionNormalizeEvent),
{
    ensure!(!inputs.is_empty(), "provide at least one workbook path");
    ensure!(
        options.capabilities.is_some() == options.platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    ensure!(
        options.apply || options.capabilities.is_none(),
        "--capabilities and --platform require --apply"
    );
    ensure!(
        options.apply || options.plan_dir.is_some(),
        "plan-only mode requires --plan-dir so the generated plans can be reviewed"
    );

    let started_at = chrono::Utc::now().to_rfc3339();
    let run_nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .context("system clock is before Unix epoch")?
        .as_nanos();
    let run_id = format!("pytools-pt-7-2-{}-{run_nonce}", std::process::id());
    let run_started = Instant::now();
    let mut report = InstitutionNormalizeRunReport {
        schema_version: 1,
        product: "pytools".to_owned(),
        feature_id: "PT-7-2".to_owned(),
        run_id: run_id.clone(),
        started_at,
        mode: if options.in_place {
            "in_place".to_owned()
        } else if options.apply {
            "apply".to_owned()
        } else {
            "plan_only".to_owned()
        },
        status: "running".to_owned(),
        total_inputs: inputs.len(),
        completed_files: 0,
        failed_files: 0,
        modified_files: 0,
        branch_sheets: 0,
        mapped_sheets: 0,
        mapped_rows: 0,
        mapping_empty: false,
        configuration_elapsed_ms: 0,
        elapsed_ms: 0,
        failure: None,
        events: Vec::new(),
        files: Vec::new(),
    };

    emit_institution_normalize_event(
        &mut report,
        &mut on_event,
        "configuration_read",
        "started",
        "正在读取机构映射配置。".to_owned(),
        InstitutionNormalizeProgress {
            total_files: inputs.len(),
            ..Default::default()
        },
        None,
        None,
        vec!["Rust".to_owned(), "calamine".to_owned()],
        None,
        None,
        None,
        None,
    );
    let config_started = Instant::now();
    let mapping = match crate::config::load_institution_mapping_ordered(config_path.as_ref()) {
        Ok(mapping) => mapping,
        Err(error) => {
            report.configuration_elapsed_ms = config_started.elapsed().as_millis();
            report.status = "failed".to_owned();
            let error_type = "configuration_read_failed".to_owned();
            let user_message = "机构映射配置读取失败，未处理所选工作簿。".to_owned();
            let diagnostic = format!("PT-7-2 stage=configuration_read: {error:#}");
            report.failure = Some(InstitutionNormalizeRunFailure {
                stage: "configuration_read".to_owned(),
                error_type: error_type.clone(),
                user_message: user_message.clone(),
                diagnostic: diagnostic.clone(),
            });
            let config_elapsed_ms = report.configuration_elapsed_ms;
            emit_institution_normalize_event(
                &mut report,
                &mut on_event,
                "configuration_read",
                "failed",
                user_message,
                InstitutionNormalizeProgress {
                    total_files: inputs.len(),
                    ..Default::default()
                },
                None,
                Some(config_elapsed_ms),
                vec!["Rust".to_owned(), "calamine".to_owned()],
                None,
                Some(error_type.clone()),
                Some(diagnostic.clone()),
                None,
            );
            report.elapsed_ms = run_started.elapsed().as_millis();
            let elapsed_ms = report.elapsed_ms;
            let final_progress = current_institution_progress(&report, inputs.len());
            emit_institution_normalize_event(
                &mut report,
                &mut on_event,
                "task",
                "failed",
                "任务失败：映射配置未能读取，未处理工作簿。".to_owned(),
                final_progress,
                None,
                Some(elapsed_ms),
                vec!["Rust".to_owned()],
                None,
                Some(error_type),
                None,
                None,
            );
            return Ok(report);
        }
    };
    report.configuration_elapsed_ms = config_started.elapsed().as_millis();
    let configuration_elapsed_ms = report.configuration_elapsed_ms;
    emit_institution_normalize_event(
        &mut report,
        &mut on_event,
        "configuration_read",
        "completed",
        format!("机构映射配置已读取，共 {} 条有效映射。", mapping.len()),
        InstitutionNormalizeProgress {
            total_files: inputs.len(),
            ..Default::default()
        },
        None,
        Some(configuration_elapsed_ms),
        vec!["Rust".to_owned(), "calamine".to_owned()],
        None,
        None,
        None,
        None,
    );

    if mapping.is_empty() {
        report.mapping_empty = true;
        emit_institution_normalize_event(
            &mut report,
            &mut on_event,
            "normalization",
            "skipped",
            "机构映射表为空，未处理所选工作簿。".to_owned(),
            InstitutionNormalizeProgress {
                total_files: inputs.len(),
                ..Default::default()
            },
            None,
            None,
            vec!["Rust".to_owned()],
            None,
            None,
            None,
            None,
        );
        report.status = "completed".to_owned();
        report.elapsed_ms = run_started.elapsed().as_millis();
        let elapsed_ms = report.elapsed_ms;
        let final_progress = current_institution_progress(&report, inputs.len());
        emit_institution_normalize_event(
            &mut report,
            &mut on_event,
            "task",
            "completed",
            "任务结束：机构映射表为空，未处理工作簿。".to_owned(),
            final_progress,
            None,
            Some(elapsed_ms),
            vec!["Rust".to_owned()],
            None,
            None,
            None,
            None,
        );
        return Ok(report);
    }

    for (index, input) in inputs.iter().enumerate() {
        let file_name = input
            .file_name()
            .unwrap_or_default()
            .to_string_lossy()
            .into_owned();
        let mut progress = current_institution_progress(&report, inputs.len());
        progress.current_file = Some(index + 1);
        emit_institution_normalize_event(
            &mut report,
            &mut on_event,
            "file_processing",
            "started",
            format!("正在处理 {}/{}：{}", index + 1, inputs.len(), file_name),
            progress.clone(),
            Some(file_name.clone()),
            None,
            vec!["Rust".to_owned()],
            None,
            None,
            None,
            None,
        );

        let candidate = institution_normalize_candidate(input, &run_id, index + 1)?;
        let file_started = Instant::now();
        let outcome = build_and_optionally_apply_institution_plan(
            input,
            &candidate,
            index + 1,
            inputs.len(),
            &run_id,
            &mapping,
            &options,
            &mut report,
            &mut on_event,
            progress,
        );
        let mut file_result = match outcome {
            Ok(mut result) => {
                match cleanup_institution_normalize_candidate(&candidate) {
                    Ok(()) => {}
                    Err(error) => {
                        result.cleanup_warning =
                            Some(format!("could not remove temporary candidate: {error:#}"));
                    }
                }
                result
            }
            Err(mut failure) => {
                if let Err(error) = cleanup_institution_normalize_candidate(&candidate) {
                    failure
                        .diagnostic
                        .push_str(&format!("; temporary candidate cleanup failed: {error:#}"));
                }
                report.failed_files += 1;
                InstitutionNormalizeFileResult {
                    file_name: file_name.clone(),
                    status: "failed".to_owned(),
                    branch_sheets: 0,
                    mapped_sheets: 0,
                    mapped_rows: 0,
                    operation_count: 0,
                    plan_path: None,
                    output_path: None,
                    actual_engine: failure.actual_engine.clone(),
                    stage_components: failure.stage_components.clone(),
                    execution: failure.execution.clone(),
                    commit: None,
                    timings: failure.timings.clone(),
                    error_stage: Some(failure.stage.to_owned()),
                    error_type: Some(failure.error_type.to_owned()),
                    user_message: user_message_for_normalize_failure(
                        failure.error_type,
                        &file_name,
                        index + 1,
                        inputs.len(),
                    ),
                    diagnostic: Some(failure.diagnostic),
                    route_diagnostic: failure.route_diagnostic,
                    cleanup_warning: None,
                }
            }
        };

        if file_result.status != "failed" {
            report.completed_files += 1;
            report.branch_sheets += file_result.branch_sheets;
            report.mapped_sheets += file_result.mapped_sheets;
            report.mapped_rows += file_result.mapped_rows;
            if file_result.status == "applied" {
                report.modified_files += 1;
            }
        }
        file_result.timings.total_ms = file_started.elapsed().as_millis();
        let event_status = if file_result.status == "failed" {
            "failed"
        } else {
            "completed"
        };
        let mut components = vec!["Rust".to_owned()];
        if let Some(engine) = file_result.actual_engine.as_deref() {
            components.push(engine.to_owned());
            components.push(engine_language(engine).to_owned());
        }
        let mut progress = current_institution_progress(&report, inputs.len());
        progress.current_file = Some(index + 1);
        emit_institution_normalize_event(
            &mut report,
            &mut on_event,
            "file_processing",
            event_status,
            file_result.user_message.clone(),
            progress,
            Some(file_name),
            Some(file_result.timings.total_ms),
            components,
            file_result.actual_engine.clone(),
            file_result.error_type.clone(),
            file_result.diagnostic.clone(),
            file_result.route_diagnostic.clone(),
        );
        report.files.push(file_result);
    }

    report.status = if report.failed_files == 0 {
        "completed".to_owned()
    } else {
        "partial_failure".to_owned()
    };
    report.elapsed_ms = run_started.elapsed().as_millis();
    let final_status = report.status.clone();
    let final_message = format!(
        "处理结束：成功 {} 个，失败 {} 个，修改 {} 个文件。",
        report.completed_files, report.failed_files, report.modified_files
    );
    let final_progress = current_institution_progress(&report, inputs.len());
    let elapsed_ms = report.elapsed_ms;
    emit_institution_normalize_event(
        &mut report,
        &mut on_event,
        "task",
        &final_status,
        final_message,
        final_progress,
        None,
        Some(elapsed_ms),
        vec!["Rust".to_owned()],
        None,
        None,
        None,
        None,
    );
    Ok(report)
}

fn build_and_optionally_apply_institution_plan(
    input: &Path,
    candidate: &Path,
    index: usize,
    total_files: usize,
    run_id: &str,
    mapping: &[(String, String)],
    options: &InstitutionNormalizeRunOptions,
    report: &mut InstitutionNormalizeRunReport,
    on_event: &mut impl FnMut(&InstitutionNormalizeEvent),
    progress: InstitutionNormalizeProgress,
) -> std::result::Result<InstitutionNormalizeFileResult, InstitutionNormalizeFileFailure> {
    let file_name = input
        .file_name()
        .unwrap_or_default()
        .to_string_lossy()
        .into_owned();
    let mut timings = InstitutionNormalizeStageTimings::default();
    let mut stage_components = BTreeMap::new();
    let mut actual_engine = None;
    if !has_extension(input, "xlsx") {
        return Err(normalize_file_failure(
            "format_check",
            "unsupported_format",
            "this PT-7-2 candidate currently supports .xlsx only; convert other formats before retrying",
            None,
        )
        .attach_state(&timings, &stage_components, actual_engine.as_deref()));
    }
    let metadata = fs::symlink_metadata(input).map_err(|error| {
        normalize_file_failure(
            "input_check",
            "input_unavailable",
            format!("inspect input workbook: {error}"),
            None,
        )
        .attach_state(&timings, &stage_components, actual_engine.as_deref())
    })?;
    if !metadata.is_file() || metadata.file_type().is_symlink() {
        return Err(normalize_file_failure(
            "input_check",
            "unsafe_input",
            "input must be a regular non-symlink file",
            None,
        )
        .attach_state(&timings, &stage_components, actual_engine.as_deref()));
    }

    let source_hash = if options.apply {
        let started = Instant::now();
        stage_components.insert(
            "source_fingerprint".to_owned(),
            vec!["Rust".to_owned(), "sha2".to_owned()],
        );
        emit_institution_stage(
            report,
            on_event,
            "source_fingerprint",
            "started",
            format!("正在校验源文件：{file_name}。"),
            &progress,
            &file_name,
            None,
            vec!["Rust".to_owned(), "sha2".to_owned()],
            None,
            None,
            None,
            None,
        );
        let hash_result = crate::file_commit::sha256_file(input);
        timings.source_fingerprint_ms = started.elapsed().as_millis();
        let hash = match hash_result {
            Ok(hash) => hash,
            Err(error) => {
                let failure = normalize_file_failure(
                    "source_fingerprint",
                    "input_unavailable",
                    format!("hash input workbook: {error:#}"),
                    None,
                )
                .attach_state(
                    &timings,
                    &stage_components,
                    actual_engine.as_deref(),
                );
                emit_institution_stage(
                    report,
                    on_event,
                    "source_fingerprint",
                    "failed",
                    format!("源文件安全检查失败：{file_name}。"),
                    &progress,
                    &file_name,
                    Some(timings.source_fingerprint_ms),
                    vec!["Rust".to_owned(), "sha2".to_owned()],
                    None,
                    Some(failure.error_type.to_owned()),
                    Some(failure.diagnostic.clone()),
                    None,
                );
                return Err(failure);
            }
        };
        emit_institution_stage(
            report,
            on_event,
            "source_fingerprint",
            "completed",
            format!("源文件安全检查完成：{file_name}。"),
            &progress,
            &file_name,
            Some(timings.source_fingerprint_ms),
            vec!["Rust".to_owned(), "sha2".to_owned()],
            None,
            None,
            None,
            None,
        );
        Some(hash)
    } else {
        None
    };

    let profile_started = Instant::now();
    stage_components.insert(
        "profile_inspection".to_owned(),
        vec!["Rust".to_owned(), "zip".to_owned(), "roxmltree".to_owned()],
    );
    emit_institution_stage(
        report,
        on_event,
        "profile_inspection",
        "started",
        format!("正在检查工作簿结构：{file_name}。"),
        &progress,
        &file_name,
        None,
        vec!["Rust".to_owned(), "zip".to_owned(), "roxmltree".to_owned()],
        None,
        None,
        None,
        None,
    );
    let profile_result = excel_editor::inspect_xlsx_profile(input);
    timings.profile_inspection_ms = profile_started.elapsed().as_millis();
    let profile = match profile_result {
        Ok(profile) => profile,
        Err(error) => {
            let failure = normalize_file_failure(
                "profile_inspection",
                "workbook_inspection_failed",
                format!("inspect workbook package: {error:#}"),
                None,
            )
            .attach_state(&timings, &stage_components, actual_engine.as_deref());
            emit_institution_stage(
                report,
                on_event,
                "profile_inspection",
                "failed",
                format!("工作簿结构检查失败：{file_name}。"),
                &progress,
                &file_name,
                Some(timings.profile_inspection_ms),
                vec!["Rust".to_owned(), "zip".to_owned(), "roxmltree".to_owned()],
                None,
                Some(failure.error_type.to_owned()),
                Some(failure.diagnostic.clone()),
                None,
            );
            return Err(failure);
        }
    };
    emit_institution_stage(
        report,
        on_event,
        "profile_inspection",
        "completed",
        format!("工作簿结构检查完成：{file_name}。"),
        &progress,
        &file_name,
        Some(timings.profile_inspection_ms),
        vec!["Rust".to_owned(), "zip".to_owned(), "roxmltree".to_owned()],
        None,
        None,
        None,
        None,
    );

    let plan_started = Instant::now();
    stage_components.insert(
        "plan_build".to_owned(),
        vec!["Rust".to_owned(), "calamine".to_owned()],
    );
    emit_institution_stage(
        report,
        on_event,
        "plan_build",
        "started",
        format!("正在按映射规则检查工作表：{file_name}。"),
        &progress,
        &file_name,
        None,
        vec!["Rust".to_owned(), "calamine".to_owned()],
        None,
        None,
        None,
        None,
    );
    let plan_result_result = build_institution_normalize_plan(
        input,
        candidate,
        profile,
        format!("{run_id}-{index}"),
        mapping,
    );
    timings.plan_build_ms = plan_started.elapsed().as_millis();
    let plan_result = match plan_result_result {
        Ok(plan_result) => plan_result,
        Err(error) => {
            let failure = normalize_file_failure(
                "plan_build",
                "planning_failed",
                format!("build workbook edit plan: {error:#}"),
                None,
            )
            .attach_state(&timings, &stage_components, actual_engine.as_deref());
            emit_institution_stage(
                report,
                on_event,
                "plan_build",
                "failed",
                format!("映射规则处理失败：{file_name}。"),
                &progress,
                &file_name,
                Some(timings.plan_build_ms),
                vec!["Rust".to_owned(), "calamine".to_owned()],
                None,
                Some(failure.error_type.to_owned()),
                Some(failure.diagnostic.clone()),
                None,
            );
            return Err(failure);
        }
    };
    emit_institution_stage(
        report,
        on_event,
        "plan_build",
        "completed",
        format!(
            "映射规则检查完成：{file_name}，命中 {} 行。",
            plan_result.mapped_rows
        ),
        &progress,
        &file_name,
        Some(timings.plan_build_ms),
        vec!["Rust".to_owned(), "calamine".to_owned()],
        None,
        None,
        None,
        None,
    );

    let Some(plan) = plan_result.plan else {
        return Ok(InstitutionNormalizeFileResult {
            file_name: file_name.clone(),
            status: "no_changes".to_owned(),
            branch_sheets: plan_result.branch_sheets,
            mapped_sheets: 0,
            mapped_rows: 0,
            operation_count: 0,
            plan_path: None,
            output_path: None,
            actual_engine: None,
            stage_components,
            execution: None,
            commit: None,
            timings,
            error_stage: None,
            error_type: None,
            user_message: format!("完成 {index}/{total_files}：{file_name}，无需修改。"),
            diagnostic: None,
            route_diagnostic: None,
            cleanup_warning: None,
        });
    };

    let operation_count = plan.operations.len();
    let plan_path = if let Some(plan_dir) = options.plan_dir.as_deref() {
        let path = plan_dir.join(format!("{index:04}_institution_normalize.json"));
        stage_components.insert(
            "plan_persistence".to_owned(),
            vec!["Rust".to_owned(), "serde_json".to_owned()],
        );
        emit_institution_stage(
            report,
            on_event,
            "plan_persistence",
            "started",
            format!("正在保存修改计划：{file_name}。"),
            &progress,
            &file_name,
            None,
            vec!["Rust".to_owned(), "serde_json".to_owned()],
            None,
            None,
            None,
            None,
        );
        let write_started = Instant::now();
        let write_result = write_json_create_new(&path, &plan);
        timings.plan_persistence_ms = write_started.elapsed().as_millis();
        if let Err(error) = write_result {
            let failure = normalize_file_failure(
                "plan_persistence",
                "plan_output_failed",
                format!("write reviewable plan: {error:#}"),
                None,
            )
            .attach_state(&timings, &stage_components, actual_engine.as_deref());
            emit_institution_stage(
                report,
                on_event,
                "plan_persistence",
                "failed",
                format!("修改计划保存失败：{file_name}。"),
                &progress,
                &file_name,
                Some(timings.plan_persistence_ms),
                vec!["Rust".to_owned(), "serde_json".to_owned()],
                None,
                Some(failure.error_type.to_owned()),
                Some(failure.diagnostic.clone()),
                None,
            );
            return Err(failure);
        }
        emit_institution_stage(
            report,
            on_event,
            "plan_persistence",
            "completed",
            format!("修改计划已保存：{file_name}。"),
            &progress,
            &file_name,
            Some(timings.plan_persistence_ms),
            vec!["Rust".to_owned(), "serde_json".to_owned()],
            None,
            None,
            None,
            None,
        );
        Some(path)
    } else {
        None
    };

    if !options.apply {
        return Ok(InstitutionNormalizeFileResult {
            file_name: file_name.clone(),
            status: "planned".to_owned(),
            branch_sheets: plan_result.branch_sheets,
            mapped_sheets: plan_result.mapped_sheets,
            mapped_rows: plan_result.mapped_rows,
            operation_count,
            plan_path,
            output_path: None,
            actual_engine: None,
            stage_components,
            execution: None,
            commit: None,
            timings,
            error_stage: None,
            error_type: None,
            user_message: format!(
                "已生成 {index}/{total_files}：{file_name} 的修改计划，涉及 {} 行；尚未修改原文件。",
                plan_result.mapped_rows
            ),
            diagnostic: None,
            route_diagnostic: None,
            cleanup_warning: None,
        });
    }

    let output_path = if options.in_place {
        // 就地模式：输出即源文件，不再预留副本路径。
        input.to_path_buf()
    } else {
        institution_normalize_output_path(input).map_err(|error| {
            normalize_file_failure(
                "copy_target_reservation",
                "copy_target_unavailable",
                format!("reserve normalized copy path: {error:#}"),
                None,
            )
            .attach_state(&timings, &stage_components, actual_engine.as_deref())
        })?
    };

    let source_check_started = Instant::now();
    let expected_source_hash = source_hash.expect("apply mode computes source hash");
    stage_components.insert(
        "source_stability_check".to_owned(),
        vec!["Rust".to_owned(), "sha2".to_owned()],
    );
    emit_institution_stage(
        report,
        on_event,
        "source_stability_check",
        "started",
        format!("正在确认源文件未变化：{file_name}。"),
        &progress,
        &file_name,
        None,
        vec!["Rust".to_owned(), "sha2".to_owned()],
        None,
        None,
        None,
        None,
    );
    let current_hash_result = crate::file_commit::sha256_file(input);
    timings.source_stability_check_ms = source_check_started.elapsed().as_millis();
    let current_source_hash = match current_hash_result {
        Ok(hash) => hash,
        Err(error) => {
            let failure = normalize_file_failure(
                "source_stability_check",
                "input_unavailable",
                format!("recheck input workbook: {error:#}"),
                None,
            )
            .attach_state(&timings, &stage_components, actual_engine.as_deref());
            emit_institution_stage(
                report,
                on_event,
                "source_stability_check",
                "failed",
                format!("源文件一致性检查失败：{file_name}。"),
                &progress,
                &file_name,
                Some(timings.source_stability_check_ms),
                vec!["Rust".to_owned(), "sha2".to_owned()],
                None,
                Some(failure.error_type.to_owned()),
                Some(failure.diagnostic.clone()),
                None,
            );
            return Err(failure);
        }
    };
    if current_source_hash != expected_source_hash {
        let failure = normalize_file_failure(
            "source_stability_check",
            "input_changed",
            "input workbook changed while the modification plan was being built",
            None,
        )
        .attach_state(&timings, &stage_components, actual_engine.as_deref());
        emit_institution_stage(
            report,
            on_event,
            "source_stability_check",
            "failed",
            format!("源文件在计划期间发生变化，已停止修改：{file_name}。"),
            &progress,
            &file_name,
            Some(timings.source_stability_check_ms),
            vec!["Rust".to_owned(), "sha2".to_owned()],
            None,
            Some(failure.error_type.to_owned()),
            Some(failure.diagnostic.clone()),
            None,
        );
        return Err(failure);
    }
    emit_institution_stage(
        report,
        on_event,
        "source_stability_check",
        "completed",
        format!("源文件一致性检查通过：{file_name}。"),
        &progress,
        &file_name,
        Some(timings.source_stability_check_ms),
        vec!["Rust".to_owned(), "sha2".to_owned()],
        None,
        None,
        None,
        None,
    );

    let modifier_route = options.capabilities.is_some();
    let selected_components = if modifier_route {
        vec!["Rust".to_owned(), "ExcelModifier".to_owned()]
    } else {
        vec![
            "Rust".to_owned(),
            "excel-writer".to_owned(),
            "rust_xlsxwriter".to_owned(),
        ]
    };
    let apply_started = Instant::now();
    stage_components.insert("engine_execution".to_owned(), selected_components.clone());
    emit_institution_stage(
        report,
        on_event,
        "engine_execution",
        "started",
        if modifier_route {
            format!("正在按已登记能力生成规范化副本：{file_name}。")
        } else {
            format!("正在用 Rust 新建并验证规范化工作簿：{file_name}。")
        },
        &progress,
        &file_name,
        None,
        selected_components.clone(),
        None,
        None,
        None,
        None,
    );
    let execution_result = if options.in_place {
        let engine = match options.engine.as_deref() {
            Some("openpyxl") => crate::inplace_edit::InPlaceEngine::Openpyxl,
            _ => crate::inplace_edit::InPlaceEngine::Umya,
        };
        crate::inplace_edit::execute_in_place(&plan, engine).map(|outcome| {
            let engine_name = engine.label().to_owned();
            let mut components = vec!["Rust".to_owned(), "in-place".to_owned()];
            components.push(engine_name.clone());
            (
                Some(serde_json::json!({
                    "engine": engine_name,
                    "mode": "in_place",
                    "execution": outcome.execution,
                    "commit": outcome.commit,
                })),
                engine_name,
                components,
            )
        })
    } else if let (Some(capability_file), Some(platform)) =
        (options.capabilities.as_deref(), options.platform.as_deref())
    {
        crate::edit_plan::apply_plan_with_capability_file(&plan, capability_file, platform).map(
            |execution| {
                let engine = execution.engine.to_string();
                let mut components = vec!["Rust".to_owned(), "ExcelModifier".to_owned()];
                components.push(engine.clone());
                components.push(engine_language(&engine).to_owned());
                (serde_json::to_value(execution).ok(), engine, components)
            },
        )
    } else {
        create_rust_new_workbook_from_cell_plan(input, candidate, &plan).map(|()| {
            (
                Some(serde_json::json!({
                    "engine": "rust_xlsxwriter",
                    "execution": "created_and_validated_workbook",
                    "operation_count": plan.operations.len(),
                })),
                "rust_xlsxwriter".to_owned(),
                selected_components.clone(),
            )
        })
    };
    timings.engine_execution_ms = apply_started.elapsed().as_millis();
    let execution = execution_result.map_err(|error| {
        let modifier_error = error.downcast_ref::<excel_editor::ExcelModifierError>();
        let route_diagnostic = modifier_error.and_then(|error| serde_json::to_value(error).ok());
        let error_type = match modifier_error.map(|error| error.kind) {
            Some(
                excel_editor::ExcelModifierErrorKind::NoEligibleBackend
                | excel_editor::ExcelModifierErrorKind::AllBackendsRejectedBeforeWrite,
            ) => "no_eligible_backend",
            _ if modifier_route => "backend_execution_failed",
            _ => "workbook_rebuild_failed",
        };
        let failure = normalize_file_failure(
            "engine_execution",
            error_type,
            if modifier_route {
                format!("ExcelModifier execution failed: {error:#}")
            } else {
                format!("Rust could not create a safe normalized workbook copy: {error:#}")
            },
            route_diagnostic,
        )
        .attach_state(&timings, &stage_components, actual_engine.as_deref());
        emit_institution_stage(
            report,
            on_event,
            "engine_execution",
            "failed",
            format!("规范化副本生成失败：{file_name}；原文件未被修改。"),
            &progress,
            &file_name,
            Some(timings.engine_execution_ms),
            selected_components.clone(),
            None,
            Some(failure.error_type.to_owned()),
            Some(failure.diagnostic.clone()),
            failure.route_diagnostic.clone(),
        );
        failure
    })?;
    let (execution_value, execution_engine, execution_components) = execution;
    actual_engine = Some(execution_engine.clone());
    stage_components.insert("engine_execution".to_owned(), execution_components);
    emit_institution_stage(
        report,
        on_event,
        "engine_execution",
        "completed",
        if modifier_route {
            format!("工作簿修改引擎已生成规范化副本：{file_name}。")
        } else {
            format!("Rust 新工作簿已生成并通过内容校验：{file_name}。")
        },
        &progress,
        &file_name,
        Some(timings.engine_execution_ms),
        selected_components,
        actual_engine.clone(),
        None,
        None,
        None,
    );

    let commit_started = Instant::now();
    stage_components.insert(
        "safe_commit".to_owned(),
        vec!["Rust".to_owned(), "file_commit".to_owned()],
    );
    emit_institution_stage(
        report,
        on_event,
        "safe_commit",
        "started",
        format!("正在验证源文件并另存修改副本：{file_name}。"),
        &progress,
        &file_name,
        None,
        vec!["Rust".to_owned(), "file_commit".to_owned()],
        actual_engine.clone(),
        None,
        None,
        None,
    );
    let commit_result = (|| -> Result<crate::file_commit::TargetCommitReport> {
        crate::file_commit::verify_planning_snapshot(input, Some(&expected_source_hash))
            .context("source changed during normalization; refusing to publish a copy")?;
        let output_parent = output_path
            .parent()
            .context("normalized copy path has no parent directory")?;
        ensure_institution_normalize_output_directory(output_parent, true)?;
        crate::file_commit::create_target_from_candidate(candidate, &output_path)
            .with_context(|| format!("publish normalized copy to {}", output_path.display()))
    })();
    timings.safe_commit_ms = commit_started.elapsed().as_millis();
    let commit = commit_result.map_err(|error| {
        let mut failure = normalize_file_failure(
            "safe_commit",
            "commit_failed",
            format!("safe copy publication failed: {error:#}"),
            None,
        )
        .attach_state(&timings, &stage_components, actual_engine.as_deref())
        .attach_execution(execution_value.clone());
        if let Err(cleanup_error) = cleanup_institution_normalize_candidate(candidate) {
            failure.diagnostic.push_str(&format!(
                "; validated candidate cleanup failed: {cleanup_error:#}"
            ));
        }
        emit_institution_stage(
            report,
            on_event,
            "safe_commit",
            "failed",
            format!("另存修改副本失败：{file_name}；原文件未被替换。"),
            &progress,
            &file_name,
            Some(timings.safe_commit_ms),
            vec!["Rust".to_owned(), "file_commit".to_owned()],
            actual_engine.clone(),
            Some(failure.error_type.to_owned()),
            Some(failure.diagnostic.clone()),
            None,
        );
        failure
    })?;
    emit_institution_stage(
        report,
        on_event,
        "safe_commit",
        "completed",
        format!("规范化副本已验证并保存：{}。", output_path.display()),
        &progress,
        &file_name,
        Some(timings.safe_commit_ms),
        vec!["Rust".to_owned(), "file_commit".to_owned()],
        actual_engine.clone(),
        None,
        None,
        None,
    );
    let mut cleanup_warning = commit.cleanup_warning.clone();

    Ok(InstitutionNormalizeFileResult {
        file_name: file_name.clone(),
        status: "applied".to_owned(),
        branch_sheets: plan_result.branch_sheets,
        mapped_sheets: plan_result.mapped_sheets,
        mapped_rows: plan_result.mapped_rows,
        operation_count,
        plan_path,
        output_path: Some(output_path),
        actual_engine,
        stage_components,
        execution: execution_value,
        commit: Some(commit.clone()),
        timings,
        error_stage: None,
        error_type: None,
        user_message: format!(
            "完成 {index}/{total_files}：{file_name}，修改 {} 行；副本已保存到 {}。",
            plan_result.mapped_rows,
            commit.target.display()
        ),
        diagnostic: None,
        route_diagnostic: None,
        cleanup_warning: cleanup_warning.take(),
    })
}

fn emit_institution_normalize_event<F>(
    report: &mut InstitutionNormalizeRunReport,
    on_event: &mut F,
    phase: &str,
    status: &str,
    user_message: String,
    progress: InstitutionNormalizeProgress,
    file_name: Option<String>,
    elapsed_ms: Option<u128>,
    components: Vec<String>,
    actual_engine: Option<String>,
    error_type: Option<String>,
    diagnostic: Option<String>,
    route_diagnostic: Option<Value>,
) where
    F: FnMut(&InstitutionNormalizeEvent),
{
    let event = InstitutionNormalizeEvent {
        sequence: report.events.len() + 1,
        timestamp: chrono::Utc::now().to_rfc3339(),
        phase: phase.to_owned(),
        status: status.to_owned(),
        user_message,
        progress,
        file_name,
        elapsed_ms,
        components,
        actual_engine,
        error_type,
        diagnostic,
        route_diagnostic,
    };
    on_event(&event);
    report.events.push(event);
}

#[allow(clippy::too_many_arguments)]
fn emit_institution_stage<F>(
    report: &mut InstitutionNormalizeRunReport,
    on_event: &mut F,
    phase: &str,
    status: &str,
    message: String,
    progress: &InstitutionNormalizeProgress,
    file_name: &str,
    elapsed_ms: Option<u128>,
    components: Vec<String>,
    actual_engine: Option<String>,
    error_type: Option<String>,
    diagnostic: Option<String>,
    route_diagnostic: Option<Value>,
) where
    F: FnMut(&InstitutionNormalizeEvent),
{
    emit_institution_normalize_event(
        report,
        on_event,
        phase,
        status,
        message,
        progress.clone(),
        Some(file_name.to_owned()),
        elapsed_ms,
        components,
        actual_engine,
        error_type,
        diagnostic,
        route_diagnostic,
    );
}

fn current_institution_progress(
    report: &InstitutionNormalizeRunReport,
    total_files: usize,
) -> InstitutionNormalizeProgress {
    InstitutionNormalizeProgress {
        total_files,
        current_file: None,
        processed_files: report.completed_files + report.failed_files,
        completed_files: report.completed_files,
        failed_files: report.failed_files,
        modified_files: report.modified_files,
    }
}

fn institution_normalize_output_path(input: &Path) -> Result<PathBuf> {
    let parent = input
        .parent()
        .filter(|path| !path.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    let result_dir = parent.join(INSTITUTION_NORMALIZE_OUTPUT_DIRECTORY);
    ensure_institution_normalize_output_directory(&result_dir, false)?;
    let file_name = input
        .file_name()
        .context("input workbook has no file name")?;
    let first_choice = result_dir.join(file_name);
    if !institution_normalize_path_exists(&first_choice)? {
        return Ok(first_choice);
    }
    let stem = input
        .file_stem()
        .context("input workbook has no file stem")?
        .to_string_lossy();
    let extension = input
        .extension()
        .context("input workbook has no file extension")?
        .to_string_lossy();
    for index in 1..=9999 {
        let candidate = result_dir.join(format!("{stem}_{index}.{extension}"));
        if !institution_normalize_path_exists(&candidate)? {
            return Ok(candidate);
        }
    }
    bail!(
        "could not allocate a non-overwriting normalization copy for {}",
        input.display()
    )
}

fn ensure_institution_normalize_output_directory(path: &Path, create: bool) -> Result<()> {
    if create {
        match fs::create_dir(path) {
            Ok(()) => {}
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => {}
            Err(error) => {
                return Err(error).with_context(|| {
                    format!("create normalization output directory {}", path.display())
                })
            }
        }
    }
    match fs::symlink_metadata(path) {
        Ok(metadata) => {
            ensure!(
                metadata.is_dir() && !metadata.file_type().is_symlink(),
                "normalization output path is not a regular directory: {}",
                path.display()
            );
            Ok(())
        }
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(()),
        Err(error) => Err(error)
            .with_context(|| format!("inspect normalization output directory {}", path.display())),
    }
}

fn institution_normalize_path_exists(path: &Path) -> Result<bool> {
    match fs::symlink_metadata(path) {
        Ok(_) => Ok(true),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(false),
        Err(error) => Err(error)
            .with_context(|| format!("inspect normalization output path {}", path.display())),
    }
}

pub fn create_rust_new_workbook_from_cell_plan(
    input: &Path,
    candidate: &Path,
    plan: &ExcelEditPlan,
) -> Result<()> {
    ensure!(
        !candidate.exists(),
        "normalization candidate path already exists: {}",
        candidate.display()
    );
    ensure!(
        plan.operations.iter().all(|operation| matches!(
            operation,
            EditOperation::SetCellValue { .. }
                | EditOperation::DeleteRows { .. }
                | EditOperation::InsertRows { .. }
        )),
        "native workbook creation supports cell-value and row insert/delete operations only"
    );

    let profile = excel_editor::inspect_xlsx_profile(input)?;
    ensure!(
        profile.inventory_complete,
        "source workbook has unknown package parts; refusing to rebuild it"
    );
    let unsupported_features = [
        excel_editor::WorkbookFeature::ConditionalFormatting,
        excel_editor::WorkbookFeature::DataValidations,
        excel_editor::WorkbookFeature::Macros,
        excel_editor::WorkbookFeature::ExternalLinks,
        excel_editor::WorkbookFeature::Charts,
        excel_editor::WorkbookFeature::PivotTables,
        excel_editor::WorkbookFeature::Slicers,
        excel_editor::WorkbookFeature::Drawings,
        excel_editor::WorkbookFeature::Images,
        excel_editor::WorkbookFeature::SheetProtection,
        excel_editor::WorkbookFeature::ActiveX,
        excel_editor::WorkbookFeature::DigitalSignatures,
        excel_editor::WorkbookFeature::FreezePanes,
        excel_editor::WorkbookFeature::AutoFilter,
    ];
    let present_unsupported = unsupported_features
        .iter()
        .filter(|feature| profile.features_present.contains(feature))
        .map(|feature| format!("{feature:?}"))
        .collect::<Vec<_>>();
    ensure!(
        present_unsupported.is_empty(),
        "source workbook contains objects the new-workbook route cannot preserve: {}",
        present_unsupported.join(", ")
    );

    let mut source = read_workbook(input).with_context(|| format!("read {}", input.display()))?;
    let mut comments = read_ooxml_comments(input)?;
    let mut defined_names = read_defined_names(input)?;
    let mut expected_names = defined_names.clone();
    let mut formula_book =
        if plan.operations.iter().any(|operation| {
            matches!(
                operation,
                EditOperation::DeleteRows { .. } | EditOperation::InsertRows { .. }
            )
        }) {
            Some(excel_editor::open_xlsx(input).with_context(|| {
                format!("open {} for row-reference simulation", input.display())
            })?)
        } else {
            None
        };

    let date_requests = source
        .sheets
        .iter()
        .flat_map(|sheet| {
            sheet.values.iter().filter_map(move |(address, value)| {
                (!sheet.formulas.contains_key(address)
                    && value.get("type").and_then(Value::as_str) == Some("datetime_serial"))
                .then(|| CellNumberFormatRequest {
                    sheet_name: sheet.name.clone(),
                    cell_reference: address.clone(),
                })
            })
        })
        .collect::<Vec<_>>();
    let mut date_formats = if date_requests.is_empty() {
        BTreeMap::new()
    } else {
        read_xlsx_effective_cell_number_formats(input, &date_requests)?
            .into_iter()
            .map(|item| ((item.sheet_name, item.cell_reference), item.format_code))
            .collect::<BTreeMap<_, _>>()
    };

    let mut replacements = BTreeMap::<(String, String), PlannedCellValue>::new();
    for operation in &plan.operations {
        match operation {
            EditOperation::SetCellValue { sheet, cell, value } => {
                let cell = normalize_workbook_cell_address(cell)?;
                ensure!(
                    replacements
                        .insert((sheet.clone(), cell.clone()), value.clone())
                        .is_none(),
                    "native workbook plan writes the same cell more than once: {sheet}!{cell}"
                );
            }
            EditOperation::DeleteRows {
                sheet,
                first_row,
                count,
            } => {
                ensure!(*count > 0, "row deletion count must be positive");
                let source_sheet = source
                    .sheets
                    .iter_mut()
                    .find(|candidate| candidate.name == *sheet)
                    .with_context(|| format!("row deletion references unknown sheet {sheet:?}"))?;
                transform_sheet_rows(source_sheet, *first_row, *count, false)?;
                transform_comment_rows(&mut comments, sheet, *first_row, *count, false)?;
                transform_number_format_rows(&mut date_formats, sheet, *first_row, *count, false)?;
                transform_replacement_rows(&mut replacements, sheet, *first_row, *count, false)?;
                transform_defined_name_rows(&mut expected_names, sheet, *first_row, *count, false)?;
                formula_book
                    .as_mut()
                    .expect("row operations initialize formula reference simulation")
                    .remove_row(sheet, *first_row, *count);
            }
            EditOperation::InsertRows {
                sheet,
                first_row,
                count,
            } => {
                ensure!(*count > 0, "row insertion count must be positive");
                let source_sheet = source
                    .sheets
                    .iter_mut()
                    .find(|candidate| candidate.name == *sheet)
                    .with_context(|| format!("row insertion references unknown sheet {sheet:?}"))?;
                transform_sheet_rows(source_sheet, *first_row, *count, true)?;
                transform_comment_rows(&mut comments, sheet, *first_row, *count, true)?;
                transform_number_format_rows(&mut date_formats, sheet, *first_row, *count, true)?;
                transform_replacement_rows(&mut replacements, sheet, *first_row, *count, true)?;
                transform_defined_name_rows(&mut expected_names, sheet, *first_row, *count, true)?;
                formula_book
                    .as_mut()
                    .expect("row operations initialize formula reference simulation")
                    .insert_new_row(sheet, *first_row, *count);
            }
            _ => bail!("unsupported operation in Rust new-workbook route"),
        }
    }

    if let Some(formula_book) = formula_book.as_ref() {
        for sheet in &mut source.sheets {
            let output_sheet = formula_book
                .sheet_by_name(&sheet.name)
                .with_context(|| format!("simulated worksheet {:?} disappeared", sheet.name))?;
            sheet.formulas = output_sheet
                .cells()
                .into_iter()
                .filter_map(|cell| {
                    let formula = cell.formula();
                    (!formula.is_empty())
                        .then(|| (cell.coordinate().to_string(), formula.to_owned()))
                })
                .collect();
        }
    }

    defined_names = expected_names.clone();
    let expected_replacements = replacements.clone();

    let mut workbook = NewWorkbook::new();
    for sheet in &source.sheets {
        let worksheet = workbook.add_worksheet();
        worksheet
            .set_name(&sheet.name)
            .with_context(|| format!("create worksheet {:?}", sheet.name))?;
        worksheet.set_formula_result_default("");
        let ranges = sheet
            .merges
            .iter()
            .map(|range| parse_workbook_cell_range(range))
            .collect::<Result<Vec<_>>>()?;
        for (first_row, first_column, last_row, last_column) in &ranges {
            ensure!(
                !sheet
                    .values
                    .keys()
                    .chain(sheet.formulas.keys())
                    .any(|address| {
                        parse_workbook_cell_address(address).is_ok_and(|(row, column)| {
                            row >= *first_row
                                && row <= *last_row
                                && column >= *first_column
                                && column <= *last_column
                                && (row, column) != (*first_row, *first_column)
                        })
                    }),
                "merged range in {:?} contains non-anchor cell data that cannot be copied safely",
                sheet.name
            );
            worksheet.merge_range(
                *first_row,
                *first_column,
                *last_row,
                *last_column,
                "",
                &Format::new(),
            )?;
        }

        for (address, formula) in &sheet.formulas {
            let key = (sheet.name.clone(), address.clone());
            if let Some(replacement) = replacements.remove(&key) {
                let (row, column) = parse_workbook_cell_address(address)?;
                write_new_plan_value(worksheet, row, column, &replacement)?;
                continue;
            }
            let (row, column) = parse_workbook_cell_address(&address)?;
            let formula = if formula.starts_with('=') {
                formula.clone()
            } else {
                format!("={formula}")
            };
            worksheet.write_formula(row, column, Formula::new(formula))?;
            if let Some(cached_value) = sheet.values.get(address) {
                if let Some(result) = formula_cached_result(cached_value)? {
                    worksheet.set_formula_result(row, column, result);
                }
            }
        }

        for (address, value) in &sheet.values {
            if sheet.formulas.contains_key(address) {
                continue;
            }
            let (row, column) = parse_workbook_cell_address(address)?;
            if cell_is_merged_non_anchor(row, column, &ranges) {
                continue;
            }
            let key = (sheet.name.clone(), address.clone());
            if let Some(replacement) = replacements.remove(&key) {
                write_new_plan_value(worksheet, row, column, &replacement)?;
            } else {
                write_institution_source_value(
                    worksheet,
                    row,
                    column,
                    value,
                    date_formats.get(&key).and_then(Option::as_deref),
                )?;
            }
        }
        let missing_cells = replacements
            .keys()
            .filter(|(target_sheet, address)| {
                target_sheet == &sheet.name
                    && !sheet.values.contains_key(address)
                    && !sheet.formulas.contains_key(address)
            })
            .cloned()
            .collect::<Vec<_>>();
        for (target_sheet, address) in missing_cells {
            let (row, column) = parse_workbook_cell_address(&address)?;
            ensure!(
                !cell_is_merged_non_anchor(row, column, &ranges),
                "normalization target {target_sheet}!{address} is inside a merged range"
            );
            let key = (target_sheet, address);
            let replacement = replacements
                .remove(&key)
                .expect("collected replacement must still be present");
            write_new_plan_value(worksheet, row, column, &replacement)?;
        }

        for (address, comment) in comments.get(&sheet.name).into_iter().flatten() {
            let (row, column) = parse_workbook_cell_address(address)?;
            let note = match comment.author.as_deref() {
                Some(author) => Note::new(&comment.text)
                    .add_author_prefix(false)
                    .set_author(author),
                None => Note::new(&comment.text),
            };
            worksheet.insert_note(row, column, &note)?;
        }
    }
    ensure!(
        replacements.is_empty(),
        "normalization plan references an unknown worksheet or unsupported cell"
    );

    for defined_name in &defined_names {
        let name = match &defined_name.scope {
            DefinedNameScope::Workbook => defined_name.name.clone(),
            DefinedNameScope::Worksheet {
                sheet_name: Some(sheet_name),
                ..
            } => format!("'{}'!{}", sheet_name.replace('\'', "''"), defined_name.name),
            scope => bail!(
                "cannot preserve defined name {:?} with scope {scope:?}",
                defined_name.name
            ),
        };
        workbook.define_name(name, &defined_name.raw_reference)?;
    }

    workbook.save(candidate).with_context(|| {
        format!(
            "write normalized workbook candidate {}",
            candidate.display()
        )
    })?;
    validate_institution_normalize_candidate(
        candidate,
        &source,
        &expected_replacements,
        &comments,
        &defined_names,
    )?;
    Ok(())
}

fn validate_institution_normalize_candidate(
    candidate: &Path,
    source: &excel_reader::WorkbookDump,
    replacements: &BTreeMap<(String, String), PlannedCellValue>,
    comments: &BTreeMap<String, BTreeMap<String, excel_reader::WorkbookComment>>,
    defined_names: &[excel_reader::DefinedNameRecord],
) -> Result<()> {
    let output = read_workbook(candidate).with_context(|| {
        format!(
            "reopen normalized workbook candidate {}",
            candidate.display()
        )
    })?;
    ensure!(
        source
            .sheets
            .iter()
            .map(|sheet| &sheet.name)
            .collect::<Vec<_>>()
            == output
                .sheets
                .iter()
                .map(|sheet| &sheet.name)
                .collect::<Vec<_>>(),
        "new workbook changed worksheet names, count or order"
    );
    for (source_sheet, output_sheet) in source.sheets.iter().zip(&output.sheets) {
        let mut expected_formulas = source_sheet.formulas.clone();
        let replacement_addresses = replacements
            .keys()
            .filter(|(sheet_name, _)| sheet_name == &source_sheet.name)
            .map(|(_, address)| address.clone())
            .collect::<BTreeSet<_>>();
        for address in &replacement_addresses {
            expected_formulas.remove(address);
        }
        let expected_formula_values = expected_formulas.keys().cloned().collect::<BTreeSet<_>>();
        let expected_values = source_sheet
            .values
            .iter()
            .filter(|(address, _)| {
                !expected_formula_values.contains(*address)
                    && !replacement_addresses.contains(*address)
            })
            .map(|(address, value)| (address.clone(), value.clone()))
            .collect::<BTreeMap<_, _>>();
        let actual_values = output_sheet
            .values
            .iter()
            .filter(|(address, _)| !output_sheet.formulas.contains_key(*address))
            .map(|(address, value)| (address.clone(), value.clone()))
            .collect::<BTreeMap<_, _>>();
        let actual_unmodified_values = actual_values
            .iter()
            .filter(|(address, _)| !replacement_addresses.contains(*address))
            .map(|(address, value)| (address.clone(), value.clone()))
            .collect::<BTreeMap<_, _>>();
        ensure!(
            expected_values == actual_unmodified_values,
            "new workbook changed stored cell values or types in {:?}",
            source_sheet.name
        );
        for address in replacement_addresses {
            let replacement = &replacements[&(source_sheet.name.clone(), address.clone())];
            match replacement {
                PlannedCellValue::Blank => ensure!(
                    actual_values.get(&address).is_none_or(|value| {
                        value.get("type").and_then(Value::as_str) == Some("empty")
                    }),
                    "new workbook did not clear cell {:?}!{address}",
                    source_sheet.name
                ),
                value => ensure!(
                    actual_values
                        .get(&address)
                        .is_some_and(|actual| planned_cell_value_matches_probe(value, actual)),
                    "new workbook wrote an incorrect value or type to {:?}!{address}",
                    source_sheet.name
                ),
            }
        }
        let normalize_formula = |formula: &str| formula.trim_start_matches('=').trim().to_owned();
        ensure!(
            expected_formulas.len() == output_sheet.formulas.len()
                && expected_formulas.iter().all(|(address, formula)| {
                    output_sheet.formulas.get(address).is_some_and(|actual| {
                        normalize_formula(formula) == normalize_formula(actual)
                    })
                }),
            "new workbook changed formula text in {:?}",
            source_sheet.name
        );
        let mut expected_merges = source_sheet.merges.clone();
        let mut actual_merges = output_sheet.merges.clone();
        expected_merges.sort();
        actual_merges.sort();
        ensure!(
            expected_merges == actual_merges,
            "new workbook changed merged ranges in {:?}",
            source_sheet.name
        );
    }
    ensure!(
        &read_ooxml_comments(candidate)? == comments,
        "new workbook changed comments or authors"
    );
    ensure!(
        same_defined_name_references(&read_defined_names(candidate)?, &defined_names),
        "new workbook changed defined names or their references"
    );
    Ok(())
}

fn transform_sheet_rows(
    sheet: &mut SheetDump,
    first_row: u32,
    count: u32,
    insert: bool,
) -> Result<()> {
    sheet.values = transform_address_map(&sheet.values, first_row, count, insert)?;
    sheet.comments = transform_address_map(&sheet.comments, first_row, count, insert)?;
    let mut merges = Vec::with_capacity(sheet.merges.len());
    for range in &sheet.merges {
        if let Some(range) = transform_merge_range(range, first_row, count, insert)? {
            merges.push(range);
        }
    }
    sheet.merges = merges;
    Ok(())
}

fn transform_address_map<T: Clone>(
    entries: &BTreeMap<String, T>,
    first_row: u32,
    count: u32,
    insert: bool,
) -> Result<BTreeMap<String, T>> {
    let mut transformed = BTreeMap::new();
    for (address, value) in entries {
        if let Some(address) = transform_row_address(address, first_row, count, insert)? {
            ensure!(
                transformed.insert(address.clone(), value.clone()).is_none(),
                "row operation maps multiple cells to {address}"
            );
        }
    }
    Ok(transformed)
}

fn transform_row_address(
    address: &str,
    first_row: u32,
    count: u32,
    insert: bool,
) -> Result<Option<String>> {
    ensure!(
        first_row > 0 && count > 0,
        "row index and count must be positive"
    );
    let (row, column) = parse_workbook_cell_address(address)?;
    let first = first_row - 1;
    let end = first
        .checked_add(count)
        .context("row operation range overflows")?;
    let adjusted = if insert {
        if row < first {
            Some(row)
        } else {
            Some(
                row.checked_add(count)
                    .context("inserted row index overflows")?,
            )
        }
    } else if row < first {
        Some(row)
    } else if row < end {
        None
    } else {
        Some(row - count)
    };
    adjusted
        .map(|row| {
            ensure!(row < 1_048_576, "row operation exceeds XLSX row limits");
            Ok(format_workbook_cell_address(row, column))
        })
        .transpose()
}

fn transform_merge_range(
    range: &str,
    first_row: u32,
    count: u32,
    insert: bool,
) -> Result<Option<String>> {
    let (first, first_column, last, last_column) = parse_workbook_cell_range(range)?;
    let operation_first = first_row - 1;
    let operation_end = operation_first
        .checked_add(count)
        .context("merged-range row operation overflows")?;
    let (new_first, new_last) = if insert {
        if first >= operation_first {
            (
                first.checked_add(count).context("merged row overflows")?,
                last.checked_add(count).context("merged row overflows")?,
            )
        } else if last >= operation_first {
            (
                first,
                last.checked_add(count).context("merged row overflows")?,
            )
        } else {
            (first, last)
        }
    } else {
        let overlap_first = first.max(operation_first);
        let overlap_end = (last + 1).min(operation_end);
        if overlap_first >= overlap_end {
            if first >= operation_end {
                (first - count, last - count)
            } else {
                (first, last)
            }
        } else {
            let remaining_rows = (last - first + 1) - (overlap_end - overlap_first);
            if remaining_rows == 0 {
                return Ok(None);
            }
            let new_first = if first < operation_first {
                first
            } else {
                operation_first
            };
            (new_first, new_first + remaining_rows - 1)
        }
    };
    ensure!(
        new_last < 1_048_576,
        "row operation moves merged range beyond XLSX limits"
    );
    Ok(Some(format!(
        "{}:{}",
        format_workbook_cell_address(new_first, first_column),
        format_workbook_cell_address(new_last, last_column)
    )))
}

fn transform_comment_rows(
    comments: &mut BTreeMap<String, BTreeMap<String, WorkbookComment>>,
    sheet: &str,
    first_row: u32,
    count: u32,
    insert: bool,
) -> Result<()> {
    if let Some(sheet_comments) = comments.get_mut(sheet) {
        *sheet_comments = transform_address_map(sheet_comments, first_row, count, insert)?;
    }
    Ok(())
}

fn transform_number_format_rows(
    formats: &mut BTreeMap<(String, String), Option<String>>,
    sheet: &str,
    first_row: u32,
    count: u32,
    insert: bool,
) -> Result<()> {
    let mut transformed = BTreeMap::new();
    for ((format_sheet, address), format) in formats.iter() {
        let address = if format_sheet == sheet {
            transform_row_address(address, first_row, count, insert)?
        } else {
            Some(address.clone())
        };
        if let Some(address) = address {
            transformed.insert((format_sheet.clone(), address), format.clone());
        }
    }
    *formats = transformed;
    Ok(())
}

fn transform_replacement_rows(
    replacements: &mut BTreeMap<(String, String), PlannedCellValue>,
    sheet: &str,
    first_row: u32,
    count: u32,
    insert: bool,
) -> Result<()> {
    let mut transformed = BTreeMap::new();
    for ((replacement_sheet, address), value) in replacements.iter() {
        let address = if replacement_sheet == sheet {
            transform_row_address(address, first_row, count, insert)?
        } else {
            Some(address.clone())
        };
        if let Some(address) = address {
            transformed.insert((replacement_sheet.clone(), address), value.clone());
        } else {
            bail!("row deletion removes a cell already written by the edit plan")
        }
    }
    *replacements = transformed;
    Ok(())
}

fn format_workbook_cell_address(row: u32, column: u16) -> String {
    let mut number = u32::from(column) + 1;
    let mut letters = String::new();
    while number > 0 {
        let remainder = ((number - 1) % 26) as u8;
        letters.insert(0, (b'A' + remainder) as char);
        number = (number - 1) / 26;
    }
    format!("{letters}{}", row + 1)
}

fn transform_defined_name_rows(
    names: &mut Vec<DefinedNameRecord>,
    sheet: &str,
    first_row: u32,
    count: u32,
    insert: bool,
) -> Result<()> {
    let mut transformed = Vec::with_capacity(names.len());
    for mut name in names.drain(..) {
        if name.parse_status == DefinedNameParseStatus::Areas {
            let mut areas = Vec::with_capacity(name.areas.len());
            for mut area in name.areas {
                if area.sheet_name == sheet {
                    let Some(reference) =
                        transform_defined_name_range(&area.reference, first_row, count, insert)?
                    else {
                        continue;
                    };
                    area.reference = reference;
                }
                areas.push(area);
            }
            if areas.is_empty() {
                continue;
            }
            name.raw_reference = areas
                .iter()
                .map(|area| {
                    format!(
                        "'{}'!{}",
                        area.sheet_name.replace('\'', "''"),
                        area.reference
                    )
                })
                .collect::<Vec<_>>()
                .join(",");
            name.areas = areas;
        } else {
            let local_to_target = matches!(
                &name.scope,
                DefinedNameScope::Worksheet {
                    sheet_name: Some(scope_sheet),
                    ..
                } if scope_sheet == sheet
            );
            let references_target = name
                .raw_reference
                .to_ascii_lowercase()
                .contains(&format!("{}!", sheet.to_ascii_lowercase()));
            if local_to_target {
                match try_transform_defined_name_range(
                    &name.raw_reference,
                    first_row,
                    count,
                    insert,
                )? {
                    Some(Some(reference)) => name.raw_reference = reference,
                    Some(None) => continue,
                    None if name.raw_reference.chars().all(|character| {
                        character.is_ascii_digit() || matches!(character, '.' | '-' | '+')
                    }) => {}
                    None => bail!(
                        "local defined name {:?} uses an expression that cannot be safely adjusted for row operations",
                        name.name
                    ),
                }
            } else if references_target {
                bail!(
                    "defined name {:?} uses an expression that cannot be safely adjusted for row operations",
                    name.name
                );
            }
        }
        transformed.push(name);
    }
    *names = transformed;
    Ok(())
}

fn transform_defined_name_range(
    reference: &str,
    first_row: u32,
    count: u32,
    insert: bool,
) -> Result<Option<String>> {
    try_transform_defined_name_range(reference, first_row, count, insert)?
        .context("defined-name area is not a supported A1 range")
}

fn try_transform_defined_name_range(
    reference: &str,
    first_row: u32,
    count: u32,
    insert: bool,
) -> Result<Option<Option<String>>> {
    let (start_text, end_text) = reference
        .split_once(':')
        .map_or((reference, reference), |(start, end)| (start, end));
    let single_cell = !reference.contains(':');
    let (start_row, start_column) = match parse_workbook_cell_address(start_text) {
        Ok(value) => value,
        Err(_) => return Ok(None),
    };
    let (end_row, end_column) = match parse_workbook_cell_address(end_text) {
        Ok(value) => value,
        Err(_) => return Ok(None),
    };
    ensure!(
        start_row <= end_row && start_column <= end_column,
        "defined-name range has reversed bounds: {reference:?}"
    );
    let operation_first = first_row - 1;
    let operation_end = operation_first
        .checked_add(count)
        .context("defined-name row operation overflows")?;
    let (new_start, new_end) = if insert {
        if start_row >= operation_first {
            (
                start_row
                    .checked_add(count)
                    .context("defined-name row overflows")?,
                end_row
                    .checked_add(count)
                    .context("defined-name row overflows")?,
            )
        } else if end_row >= operation_first {
            (
                start_row,
                end_row
                    .checked_add(count)
                    .context("defined-name row overflows")?,
            )
        } else {
            (start_row, end_row)
        }
    } else {
        let overlap_first = start_row.max(operation_first);
        let overlap_end = (end_row + 1).min(operation_end);
        if overlap_first >= overlap_end {
            if start_row >= operation_end {
                (start_row - count, end_row - count)
            } else {
                (start_row, end_row)
            }
        } else {
            let remaining = (end_row - start_row + 1) - (overlap_end - overlap_first);
            if remaining == 0 {
                return Ok(Some(None));
            }
            let new_start = if start_row < operation_first {
                start_row
            } else {
                operation_first
            };
            (new_start, new_start + remaining - 1)
        }
    };
    ensure!(
        new_end < 1_048_576,
        "defined-name range exceeds XLSX row limits"
    );
    let new_start = format_workbook_cell_address_with_locks(start_text, new_start, start_column)?;
    if single_cell {
        Ok(Some(Some(new_start)))
    } else {
        let new_end = format_workbook_cell_address_with_locks(end_text, new_end, end_column)?;
        Ok(Some(Some(format!("{new_start}:{new_end}"))))
    }
}

fn format_workbook_cell_address_with_locks(
    original: &str,
    row: u32,
    column: u16,
) -> Result<String> {
    let cleaned = original.replace('$', "");
    let first_digit = cleaned
        .find(|character: char| character.is_ascii_digit())
        .context("defined-name cell reference has no row number")?;
    let column_text = &cleaned[..first_digit];
    let row_locked = original
        .as_bytes()
        .get(first_digit + usize::from(original.starts_with('$')))
        == Some(&b'$');
    let column_locked = original.starts_with('$');
    let mut formatted = format_workbook_cell_address(row, column);
    let formatted_digit = formatted
        .find(|character: char| character.is_ascii_digit())
        .context("formatted cell reference has no row number")?;
    let formatted_column = &formatted[..formatted_digit];
    ensure!(
        column_text.eq_ignore_ascii_case(formatted_column),
        "defined-name column changed unexpectedly for {original:?}"
    );
    let row_text = &formatted[formatted_digit..];
    let column_lock = if column_locked { "$" } else { "" };
    let row_lock = if row_locked { "$" } else { "" };
    formatted = format!("{column_lock}{formatted_column}{row_lock}{row_text}");
    Ok(formatted)
}

fn same_defined_name_references(left: &[DefinedNameRecord], right: &[DefinedNameRecord]) -> bool {
    let as_keys = |records: &[DefinedNameRecord]| {
        records
            .iter()
            .map(|record| {
                (
                    record.name.clone(),
                    defined_name_scope_key(&record.scope),
                    record.raw_reference.clone(),
                )
            })
            .collect::<BTreeSet<_>>()
    };
    as_keys(left) == as_keys(right)
}

fn defined_name_scope_key(scope: &DefinedNameScope) -> String {
    match scope {
        DefinedNameScope::Workbook => "workbook".to_owned(),
        DefinedNameScope::Worksheet {
            local_sheet_id,
            sheet_name,
            sheet_kind,
        } => format!(
            "worksheet:{local_sheet_id}:{}:{}",
            sheet_name.as_deref().unwrap_or_default(),
            sheet_kind.as_deref().unwrap_or_default()
        ),
        DefinedNameScope::UnresolvedWorksheet { local_sheet_id } => {
            format!("unresolved:{local_sheet_id}")
        }
        DefinedNameScope::InvalidWorksheetScope { raw_local_sheet_id } => {
            format!("invalid:{raw_local_sheet_id}")
        }
    }
}

fn write_institution_source_value(
    worksheet: &mut excel_writer::Worksheet,
    row: u32,
    column: u16,
    value: &Value,
    number_format: Option<&str>,
) -> Result<()> {
    match ReaderCellValue::from_probe_json(value)? {
        ReaderCellValue::Empty => {}
        ReaderCellValue::Text(text) => {
            worksheet.write_string(row, column, &text)?;
        }
        ReaderCellValue::Number(number) => {
            worksheet.write_number(row, column, number)?;
        }
        ReaderCellValue::Integer(number) => {
            ensure!(
                number.unsigned_abs() <= (1_u64 << 53),
                "integer outside exact XLSX numeric range at row {}, column {}",
                row + 1,
                column + 1
            );
            worksheet.write_number(row, column, number as f64)?;
        }
        ReaderCellValue::Boolean(value) => {
            worksheet.write_boolean(row, column, value)?;
        }
        ReaderCellValue::DateTimeSerial(number) => {
            let number_format = number_format.context("date cell has no readable number format")?;
            worksheet.write_number_with_format(
                row,
                column,
                number,
                &Format::new().set_num_format(number_format),
            )?;
        }
        ReaderCellValue::DateTimeIso(value) | ReaderCellValue::DurationIso(value) => {
            bail!("new workbook route cannot yet preserve ISO-typed date/time value {value:?}")
        }
        ReaderCellValue::Error(value) => {
            bail!("new workbook route cannot preserve typed Excel error value {value:?}")
        }
    }
    Ok(())
}

fn planned_cell_value_matches_probe(expected: &PlannedCellValue, actual: &Value) -> bool {
    let Ok(actual) = ReaderCellValue::from_probe_json(actual) else {
        return false;
    };
    match (expected, actual) {
        (PlannedCellValue::Text(expected), ReaderCellValue::Text(actual)) => expected == &actual,
        (PlannedCellValue::Decimal(expected), actual) => {
            let Ok(expected) = expected.parse::<f64>() else {
                return false;
            };
            let actual = match actual {
                ReaderCellValue::Number(value) | ReaderCellValue::DateTimeSerial(value) => value,
                ReaderCellValue::Integer(value) => value as f64,
                _ => return false,
            };
            actual.is_finite() && actual == expected
        }
        (PlannedCellValue::Boolean(expected), ReaderCellValue::Boolean(actual)) => {
            *expected == actual
        }
        (PlannedCellValue::Blank, ReaderCellValue::Empty) => true,
        _ => false,
    }
}

fn write_new_plan_value(
    worksheet: &mut excel_writer::Worksheet,
    row: u32,
    column: u16,
    value: &PlannedCellValue,
) -> Result<()> {
    match value {
        PlannedCellValue::Text(text) => {
            worksheet.write_string(row, column, text)?;
        }
        PlannedCellValue::Decimal(text) => {
            let number = text
                .parse::<f64>()
                .with_context(|| format!("planned decimal {text:?} is not a number"))?;
            ensure!(
                number.is_finite(),
                "planned decimal must be finite: {text:?}"
            );
            worksheet.write_number(row, column, number)?;
        }
        PlannedCellValue::Boolean(value) => {
            worksheet.write_boolean(row, column, *value)?;
        }
        PlannedCellValue::Blank => {}
        _ => bail!("new workbook route does not support this planned cell value type"),
    }
    Ok(())
}

fn formula_cached_result(value: &Value) -> Result<Option<String>> {
    Ok(match ReaderCellValue::from_probe_json(value)? {
        ReaderCellValue::Empty => Some(String::new()),
        ReaderCellValue::Text(value)
        | ReaderCellValue::DateTimeIso(value)
        | ReaderCellValue::DurationIso(value)
        | ReaderCellValue::Error(value) => Some(value),
        ReaderCellValue::Number(value) | ReaderCellValue::DateTimeSerial(value) => {
            Some(value.to_string())
        }
        ReaderCellValue::Integer(value) => Some(value.to_string()),
        ReaderCellValue::Boolean(value) => Some(if value { "TRUE" } else { "FALSE" }.to_owned()),
    })
}

fn parse_workbook_cell_range(value: &str) -> Result<(u32, u16, u32, u16)> {
    let (first, last) = value
        .split_once(':')
        .with_context(|| format!("invalid merged cell range {value:?}"))?;
    let (first_row, first_column) = parse_workbook_cell_address(first)?;
    let (last_row, last_column) = parse_workbook_cell_address(last)?;
    ensure!(
        first_row <= last_row && first_column <= last_column,
        "merged range has reversed bounds: {value:?}"
    );
    Ok((first_row, first_column, last_row, last_column))
}

fn parse_workbook_cell_address(value: &str) -> Result<(u32, u16)> {
    let value = value.replace('$', "");
    let first_digit = value
        .find(|character: char| character.is_ascii_digit())
        .context("cell address has no row number")?;
    let (column, row) = value.split_at(first_digit);
    ensure!(
        !column.is_empty()
            && column
                .chars()
                .all(|character| character.is_ascii_alphabetic()),
        "cell address has an invalid column: {value:?}"
    );
    let column = column.chars().try_fold(0_u32, |current, character| {
        current
            .checked_mul(26)
            .and_then(|current| {
                current.checked_add(character.to_ascii_uppercase() as u32 - 'A' as u32 + 1)
            })
            .context("cell column exceeds XLSX limits")
    })?;
    let row = row
        .parse::<u32>()
        .context("cell row is not a positive integer")?;
    ensure!(
        (1..=1_048_576).contains(&row) && (1..=16_384).contains(&column),
        "cell address is outside XLSX limits: {value:?}"
    );
    Ok((row - 1, (column - 1) as u16))
}

fn normalize_workbook_cell_address(value: &str) -> Result<String> {
    let (row, column) = parse_workbook_cell_address(value)?;
    let mut number = u32::from(column) + 1;
    let mut letters = String::new();
    while number > 0 {
        let remainder = ((number - 1) % 26) as u8;
        letters.insert(0, (b'A' + remainder) as char);
        number = (number - 1) / 26;
    }
    Ok(format!("{letters}{}", row + 1))
}

fn cell_is_merged_non_anchor(row: u32, column: u16, ranges: &[(u32, u16, u32, u16)]) -> bool {
    ranges
        .iter()
        .any(|(first_row, first_column, last_row, last_column)| {
            row >= *first_row
                && row <= *last_row
                && column >= *first_column
                && column <= *last_column
                && (row, column) != (*first_row, *first_column)
        })
}

fn cleanup_institution_normalize_candidate(candidate: &Path) -> Result<()> {
    let file_name = candidate
        .file_name()
        .and_then(|value| value.to_str())
        .context("temporary candidate has no UTF-8 file name")?;
    ensure!(
        file_name.starts_with(".pytools-institution-normalize-")
            && candidate
                .extension()
                .is_some_and(|extension| extension == "xlsx"),
        "refusing to clean up a path that is not a generated PT-7-2 candidate"
    );
    let metadata = match fs::symlink_metadata(candidate) {
        Ok(metadata) => metadata,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => return Ok(()),
        Err(error) => {
            return Err(error)
                .with_context(|| format!("inspect temporary candidate {}", candidate.display()))
        }
    };
    ensure!(
        metadata.file_type().is_file() && !metadata.file_type().is_symlink(),
        "refusing to remove non-regular temporary candidate {}",
        candidate.display()
    );
    fs::remove_file(candidate)
        .with_context(|| format!("remove temporary candidate {}", candidate.display()))
}

fn institution_normalize_candidate(input: &Path, run_id: &str, index: usize) -> Result<PathBuf> {
    let parent = input
        .parent()
        .filter(|path| !path.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    for attempt in 0..1000usize {
        let candidate = parent.join(format!(
            ".pytools-institution-normalize-{run_id}-{index}-{attempt}.xlsx"
        ));
        if !candidate.exists() {
            return Ok(candidate);
        }
    }
    bail!(
        "could not reserve a unique PT-7-2 candidate beside {}",
        input.display()
    )
}

fn write_json_create_new<T: Serialize>(path: &Path, value: &T) -> Result<()> {
    if let Some(parent) = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
    {
        fs::create_dir_all(parent)
            .with_context(|| format!("create plan directory {}", parent.display()))?;
    }
    let mut file = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(path)
        .with_context(|| format!("create plan file {} without overwrite", path.display()))?;
    let result = (|| -> Result<()> {
        file.write_all(&serde_json::to_vec_pretty(value)?)
            .with_context(|| format!("write plan file {}", path.display()))?;
        file.sync_all()
            .with_context(|| format!("sync plan file {}", path.display()))?;
        Ok(())
    })();
    drop(file);
    if let Err(error) = result {
        let _ = fs::remove_file(path);
        return Err(error);
    }
    Ok(())
}

fn normalize_file_failure(
    stage: &'static str,
    error_type: &'static str,
    diagnostic: impl Into<String>,
    route_diagnostic: Option<Value>,
) -> InstitutionNormalizeFileFailure {
    InstitutionNormalizeFileFailure {
        stage,
        error_type,
        diagnostic: diagnostic.into(),
        route_diagnostic,
        actual_engine: None,
        execution: None,
        timings: InstitutionNormalizeStageTimings::default(),
        stage_components: BTreeMap::new(),
    }
}

fn user_message_for_normalize_failure(
    error_type: &str,
    file_name: &str,
    index: usize,
    total: usize,
) -> String {
    match error_type {
        "unsupported_format" => {
            format!("跳过 {index}/{total}：{file_name}，当前候选仅支持 .xlsx；原文件未修改。")
        }
        "no_eligible_backend" => format!(
            "失败 {index}/{total}：{file_name}，没有已验收的修改引擎可处理此文件；原文件未修改。"
        ),
        "backend_execution_failed" => {
            format!("失败 {index}/{total}：{file_name}，工作簿修改未通过验证；原文件未安全替换。")
        }
        _ => format!("失败 {index}/{total}：{file_name}，该文件未完成安全处理；请查看开发诊断。"),
    }
}

fn engine_language(engine: &str) -> &'static str {
    match engine {
        "openpyxl_worker" => "Python",
        _ => "Rust",
    }
}

/// Normalize one in-memory worksheet matrix, preserving mapping insertion order.
///
/// Rows and columns use one-based Excel semantics at the API boundary: row 1 is
/// the header, and column A is preferred whenever it has data; otherwise B is
/// used. The supplied mapping is ordered because substring matches use the
/// first matching key after exact-match lookup.
pub fn normalize_branch_sheet_values(
    sheet_name: &str,
    rows: &mut [Vec<Option<String>>],
    mapping: &[(String, String)],
) -> usize {
    if !sheet_name.contains(BRANCH_SHEET_KEY) || mapping.is_empty() {
        return 0;
    }

    let last_a = last_nonblank_row(rows, 0);
    let last_b = last_nonblank_row(rows, 1);
    let last_row = last_a.max(last_b);
    if last_row < 2 {
        return 0;
    }

    let Some(target_column) = target_column(rows, last_row) else {
        return 0;
    };

    let mut changed = 0;
    for row in rows.iter_mut().take(last_row).skip(1) {
        let Some(value) = row.get_mut(target_column) else {
            continue;
        };
        let Some(raw) = value.as_deref() else {
            continue;
        };
        let original = raw.trim();
        if original.is_empty() {
            continue;
        }
        let normalized = normalize_name(original, mapping);
        if normalized != original {
            *value = Some(normalized);
            changed += 1;
        }
    }
    changed
}

/// Build an engine-independent cell-update plan for one `.xlsx` workbook.
///
/// `workbook_profile` must come from a complete, independently validated
/// package inventory before the plan is executed. This function preserves the
/// caller's inventory status; it does not claim that reading cell values is a
/// complete OOXML feature inventory. The existing umya candidate remains
/// The old direct umya candidate is compiled only into this module's tests;
/// production callers must route the resulting plan through ExcelModifier.
pub fn build_institution_normalize_plan(
    input: impl AsRef<Path>,
    output: impl AsRef<Path>,
    workbook_profile: WorkbookProfile,
    request_id: impl Into<String>,
    mapping: &[(String, String)],
) -> Result<NormalizePlanResult> {
    if mapping.is_empty() {
        return Ok(NormalizePlanResult {
            mapping_empty: true,
            ..NormalizePlanResult::default()
        });
    }

    let input = input.as_ref();
    let output = output.as_ref();
    if !has_extension(input, "xlsx") || !has_extension(output, "xlsx") {
        return Err(anyhow!(
            "institution normalization edit plans currently accept only .xlsx"
        ));
    }
    if workbook_profile.format != WorkbookFormat::Xlsx {
        return Err(anyhow!(
            "institution normalization edit plans require an XLSX workbook profile"
        ));
    }

    let dump = read_workbook(input).with_context(|| format!("read {}", input.display()))?;
    let mut result = NormalizePlanResult::default();
    let mut operations = Vec::new();

    for sheet in dump
        .sheets
        .iter()
        .filter(|sheet| sheet.name.contains(BRANCH_SHEET_KEY))
    {
        result.branch_sheets += 1;
        if sheet.formula_status.starts_with("read_error:") {
            return Err(anyhow!(
                "cannot plan institution normalization because formula text could not be read for worksheet {}: {}",
                sheet.name,
                sheet.formula_status
            ));
        }

        let max_row = sheet
            .values
            .keys()
            .chain(sheet.formulas.keys())
            .filter_map(|address| target_column_row(address))
            .map(|(_, row)| row)
            .max()
            .unwrap_or_default();
        if max_row < 2 {
            continue;
        }

        let mut rows = vec![vec![None, None]; max_row];
        for row in 1..=max_row {
            for column in 0..2 {
                let address = target_cell_address(column, row);
                rows[row - 1][column] = cell_python_string_from_dump(sheet, &address)?;
            }
        }
        let original_rows = rows.clone();
        let changed_rows = normalize_branch_sheet_values(&sheet.name, &mut rows, mapping);
        if changed_rows == 0 {
            continue;
        }

        result.mapped_sheets += 1;
        result.mapped_rows += changed_rows;
        for row_index in 1..rows.len() {
            for column_index in 0..2 {
                if rows[row_index][column_index] == original_rows[row_index][column_index] {
                    continue;
                }
                let value = rows[row_index][column_index]
                    .as_ref()
                    .expect("normalization only writes nonblank mapped values");
                operations.push(EditOperation::SetCellValue {
                    sheet: sheet.name.clone(),
                    cell: target_cell_address(column_index, row_index + 1),
                    value: PlannedCellValue::Text(value.clone()),
                });
            }
        }
    }

    if !operations.is_empty() {
        // openpyxl rewrites known package XML (including core properties), so
        // PT-7-2 asks for explicit semantic object preservation instead of the
        // raw UnmodifiedPackageParts contract. Unknown package parts remain a
        // pre-write rejection at the backend boundary.
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
            PreservationRequirement::SheetIdentityAndOrder,
        ]);
        let mut preserve = preserve;
        if workbook_profile
            .features_present
            .contains(&excel_editor::WorkbookFeature::ConditionalFormatting)
        {
            preserve.insert(PreservationRequirement::ConditionalFormatting);
        }
        if workbook_profile
            .features_present
            .contains(&excel_editor::WorkbookFeature::DataValidations)
        {
            preserve.insert(PreservationRequirement::DataValidations);
        }
        let plan = ExcelEditPlan {
            request_id: request_id.into(),
            input: input.to_path_buf(),
            output: output.to_path_buf(),
            format: WorkbookFormat::Xlsx,
            workbook: workbook_profile,
            operations,
            preserve,
            drop_allowed: Default::default(),
        };
        plan.validate()
            .context("validate institution normalization edit plan")?;
        result.plan = Some(plan);
    }

    Ok(result)
}

fn target_column_row(address: &str) -> Option<(usize, usize)> {
    let first_digit = address.find(|character: char| character.is_ascii_digit())?;
    let column = match &address[..first_digit] {
        "A" => 0,
        "B" => 1,
        _ => return None,
    };
    let row = address[first_digit..].parse::<usize>().ok()?;
    (row > 0).then_some((column, row))
}

fn target_cell_address(column: usize, row: usize) -> String {
    format!("{}{}", if column == 0 { "A" } else { "B" }, row)
}

fn has_extension(path: &Path, expected: &str) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| extension.eq_ignore_ascii_case(expected))
}

fn cell_python_string_from_dump(sheet: &SheetDump, address: &str) -> Result<Option<String>> {
    if let Some(formula) = sheet.formulas.get(address) {
        return Ok(Some(format!("={formula}")));
    }
    let Some(value) = sheet.values.get(address) else {
        return Ok(None);
    };
    let value = ReaderCellValue::from_probe_json(value)
        .with_context(|| format!("decode {address} on worksheet {}", sheet.name))?;
    Ok(match value {
        ReaderCellValue::Empty => None,
        ReaderCellValue::Text(value)
        | ReaderCellValue::DateTimeIso(value)
        | ReaderCellValue::DurationIso(value) => Some(value),
        ReaderCellValue::Number(value) => Some(crate::summary_ops::python_float_text(value)),
        ReaderCellValue::DateTimeSerial(value) => Some(value.to_string()),
        ReaderCellValue::Integer(value) => Some(value.to_string()),
        ReaderCellValue::Boolean(value) => Some(if value { "True" } else { "False" }.to_owned()),
        ReaderCellValue::Error(value) => {
            return Err(anyhow!(
                "cannot plan institution normalization for error-valued cell {address} on worksheet {}: {value}",
                sheet.name
            ));
        }
    })
}

/// Process `.xlsx` workbooks in place. Inputs are processed concurrently when
/// their paths are unique; repeated paths are processed sequentially to avoid
/// racing writes to the same file. This legacy implementation is test-only so
/// production code cannot bypass ExcelModifier.
#[cfg(test)]
pub fn run_normalize_institution(
    target_workbooks: &[PathBuf],
    mapping: &[(String, String)],
) -> NormalizeReport {
    if mapping.is_empty() {
        return NormalizeReport {
            mapping_empty: true,
            ..NormalizeReport::default()
        };
    }

    let mut seen = HashSet::new();
    let unique_paths = target_workbooks.iter().all(|path| {
        let key = std::fs::canonicalize(path).unwrap_or_else(|_| path.clone());
        seen.insert(key)
    });
    let results = if unique_paths {
        target_workbooks
            .par_iter()
            .map(|path| (path, process_workbook(path, mapping)))
            .collect::<Vec<_>>()
    } else {
        target_workbooks
            .iter()
            .map(|path| (path, process_workbook(path, mapping)))
            .collect::<Vec<_>>()
    };

    let mut report = NormalizeReport::default();
    for (path, result) in results {
        match result {
            Ok((branch_sheets, mapped_sheets, mapped_rows)) => {
                report.files += 1;
                report.branch_sheets += branch_sheets;
                report.mapped_sheets += mapped_sheets;
                report.mapped_rows += mapped_rows;
            }
            Err(error) => {
                report.failed_files += 1;
                report.errors.push(format!("{}: {error:#}", path.display()));
            }
        }
    }
    report
}

#[cfg(test)]
fn process_workbook(path: &Path, mapping: &[(String, String)]) -> Result<(usize, usize, usize)> {
    if !path
        .extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
    {
        return Err(anyhow!(
            "当前机构名称规范化候选后端只接受 .xlsx，未验证其他格式"
        ));
    }

    let mut workbook = open_xlsx(path).with_context(|| format!("open {}", path.display()))?;
    let sheet_names = workbook
        .sheet_collection()
        .iter()
        .filter(|sheet| sheet.name().contains(BRANCH_SHEET_KEY))
        .map(|sheet| sheet.name().to_owned())
        .collect::<Vec<_>>();

    let mut branch_sheets = 0;
    let mut mapped_sheets = 0;
    let mut mapped_rows = 0;
    for sheet_name in sheet_names {
        branch_sheets += 1;
        let sheet = workbook
            .sheet_by_name_mut(&sheet_name)
            .with_context(|| format!("find worksheet {sheet_name}"))?;
        let original_column_numbers = sheet
            .column_dimensions()
            .iter()
            .map(|column| column.col_num())
            .collect::<HashSet<_>>();
        let max_row = sheet.highest_row() as usize;
        let mut values = (0..max_row)
            .map(|row_index| {
                (1_u32..=2)
                    .map(|column_index| {
                        sheet
                            .cell((column_index, row_index as u32 + 1))
                            .map(cell_python_string)
                    })
                    .collect::<Vec<_>>()
            })
            .collect::<Vec<_>>();
        let changed = normalize_branch_sheet_values(&sheet_name, &mut values, mapping);
        if changed == 0 {
            continue;
        }

        let target_column = target_column(&values, values.len()).unwrap();
        for (row_index, row) in values.iter().enumerate().skip(1) {
            let Some(value) = row.get(target_column).and_then(Option::as_ref) else {
                continue;
            };
            let original = sheet
                .cell((target_column as u32 + 1, row_index as u32 + 1))
                .map(cell_python_string)
                .unwrap_or_default();
            if *value != original {
                sheet
                    .cell_mut((target_column as u32 + 1, row_index as u32 + 1))
                    .set_value_string(value.clone());
            }
        }
        sheet
            .column_dimensions_mut()
            .retain(|column| original_column_numbers.contains(&column.col_num()));
        mapped_sheets += 1;
        mapped_rows += changed;
    }

    if mapped_rows > 0 {
        save_xlsx(&workbook, path).with_context(|| format!("save {}", path.display()))?;
    }
    Ok((branch_sheets, mapped_sheets, mapped_rows))
}

#[cfg(test)]
fn cell_python_string(cell: &excel_editor::Cell) -> String {
    if cell.is_formula() {
        format!("={}", cell.formula())
    } else if cell.data_type() == "b" {
        match cell.value().as_ref() {
            "TRUE" => "True".to_owned(),
            "FALSE" => "False".to_owned(),
            value => value.to_owned(),
        }
    } else {
        cell.value().into_owned()
    }
}

fn last_nonblank_row(rows: &[Vec<Option<String>>], column: usize) -> usize {
    rows.iter()
        .enumerate()
        .rev()
        .find(|(_, row)| {
            row.get(column)
                .and_then(Option::as_deref)
                .is_some_and(|value| !value.trim().is_empty())
        })
        .map_or(0, |(index, _)| index + 1)
}

fn target_column(rows: &[Vec<Option<String>>], last_row: usize) -> Option<usize> {
    let has_data = |column: usize| {
        rows.iter().take(last_row).skip(1).any(|row| {
            row.get(column)
                .and_then(Option::as_deref)
                .is_some_and(|value| !value.trim().is_empty())
        })
    };
    if has_data(0) {
        Some(0)
    } else if has_data(1) {
        Some(1)
    } else {
        None
    }
}

fn normalize_name(original: &str, mapping: &[(String, String)]) -> String {
    if let Some((_, standard)) = mapping.iter().find(|(key, _)| key == original) {
        return standard.clone();
    }
    let original_lower = original.to_lowercase();
    mapping
        .iter()
        .find(|(key, _)| !key.is_empty() && original_lower.contains(&key.to_lowercase()))
        .map_or_else(|| original.to_owned(), |(_, standard)| standard.clone())
}

#[cfg(test)]
mod tests {
    use super::*;
    use excel_reader::read_workbook;
    use serde_json::{json, Value};
    use std::time::{SystemTime, UNIX_EPOCH};

    #[test]
    fn floating_mapping_keys_use_python_float_string_format() {
        let sheet = SheetDump {
            name: "示例分机构".to_owned(),
            values: BTreeMap::from([(
                "A2".to_owned(),
                json!({"type": "number", "value": 0.00001}),
            )]),
            formulas: BTreeMap::new(),
            formula_status: "read_by_test".to_owned(),
            merges: Vec::new(),
            merge_status: "read_by_test".to_owned(),
            comments: BTreeMap::new(),
            comments_status: "read_by_test".to_owned(),
            conditional_formatting: Vec::new(),
            conditional_formatting_status: "read_by_test".to_owned(),
            data_validations: Vec::new(),
            data_validations_status: "read_by_test".to_owned(),
        };

        assert_eq!(
            cell_python_string_from_dump(&sheet, "A2")
                .unwrap()
                .as_deref(),
            Some("1e-05")
        );

        let mut rows = vec![
            vec![Some("机构".to_owned())],
            vec![cell_python_string_from_dump(&sheet, "A2").unwrap()],
        ];
        let mapping = vec![("1e-05".to_owned(), "Float Standard".to_owned())];
        assert_eq!(
            normalize_branch_sheet_values("示例分机构", &mut rows, &mapping),
            1
        );
        assert_eq!(rows[1][0].as_deref(), Some("Float Standard"));
    }

    #[test]
    fn workbook_normalization_matches_python_baseline_and_preserves_fixture_metadata() {
        let fixture_dir = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures");
        let baseline: Value = serde_json::from_str(include_str!(
            "../probe/reports/institution_normalize_baseline.json"
        ))
        .expect("Python baseline report must be valid JSON");
        let mapping = baseline["mapping"]
            .as_array()
            .unwrap()
            .iter()
            .map(|entry| {
                (
                    entry[0].as_str().unwrap().to_owned(),
                    entry[1].as_str().unwrap().to_owned(),
                )
            })
            .collect::<Vec<_>>();
        let source = fixture_dir.join("institution_normalize.xlsx");
        let unique = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let edited = std::env::temp_dir().join(format!(
            "pytools-institution-normalize-{}-{unique}.xlsx",
            std::process::id()
        ));
        std::fs::copy(&source, &edited).unwrap();

        let report = run_normalize_institution(std::slice::from_ref(&edited), &mapping);
        assert_eq!(
            report.files,
            baseline["stats"]["files"].as_u64().unwrap() as usize
        );
        assert_eq!(
            report.branch_sheets,
            baseline["stats"]["branch_sheets"].as_u64().unwrap() as usize
        );
        assert_eq!(
            report.mapped_sheets,
            baseline["stats"]["mapped_sheets"].as_u64().unwrap() as usize
        );
        assert_eq!(
            report.mapped_rows,
            baseline["stats"]["mapped_rows"].as_u64().unwrap() as usize
        );
        assert_eq!(report.failed_files, 0, "{:?}", report.errors);

        let dump = read_workbook(&edited).unwrap();
        let actual = dump
            .sheets
            .iter()
            .map(|sheet| {
                json!({
                    "name": sheet.name,
                    "values": sheet.values,
                    "formulas": sheet.formulas,
                    "merges": sheet.merges,
                    "comments": sheet.comments,
                })
            })
            .collect::<Vec<_>>();
        let expected = baseline["sheets"]
            .as_array()
            .unwrap()
            .iter()
            .map(|sheet| {
                json!({
                    "name": sheet["name"],
                    "values": sheet["values"],
                    "formulas": sheet["formulas"],
                    "merges": sheet["merges"],
                    "comments": sheet["comments"],
                })
            })
            .collect::<Vec<_>>();
        assert_eq!(json!(actual), json!(expected));
        std::fs::remove_file(edited).unwrap();
    }

    #[test]
    fn helper_uses_exact_match_then_first_case_insensitive_substring() {
        let mapping = vec![
            ("AB".to_owned(), "Short Standard".to_owned()),
            ("ABC Bank".to_owned(), "Exact Standard".to_owned()),
            ("branch".to_owned(), "Branch Standard".to_owned()),
        ];
        let mut rows = vec![
            vec![Some("机构".to_owned()), Some("名称".to_owned())],
            vec![Some("ABC Bank".to_owned()), Some("ignored".to_owned())],
            vec![Some(" abc bank office ".to_owned()), None],
            vec![Some("   ".to_owned()), None],
        ];
        assert_eq!(
            normalize_branch_sheet_values("报表分机构", &mut rows, &mapping),
            2
        );
        assert_eq!(rows[1][0].as_deref(), Some("Exact Standard"));
        assert_eq!(rows[2][0].as_deref(), Some("Short Standard"));

        let mut fallback_rows = vec![
            vec![Some("机构".to_owned()), Some("名称".to_owned())],
            vec![None, Some("Branch Office".to_owned())],
        ];
        assert_eq!(
            normalize_branch_sheet_values("分机构-备用", &mut fallback_rows, &mapping),
            1
        );
        assert_eq!(fallback_rows[1][1].as_deref(), Some("Branch Standard"));
        assert_eq!(
            normalize_branch_sheet_values("普通表", &mut rows, &mapping),
            0
        );
    }

    #[test]
    fn empty_mapping_skips_all_files() {
        let missing = PathBuf::from("missing.xlsx");
        let report = run_normalize_institution(&[missing], &[]);

        assert!(report.mapping_empty);
        assert_eq!(report.files, 0);
        assert_eq!(report.failed_files, 0);
    }

    #[test]
    fn repeated_paths_are_processed_sequentially_without_double_mapping() {
        let fixture =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/institution_normalize.xlsx");
        let unique = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let edited = std::env::temp_dir().join(format!(
            "pytools-institution-normalize-duplicate-{}-{unique}.xlsx",
            std::process::id()
        ));
        std::fs::copy(fixture, &edited).unwrap();
        let mapping = vec![("ABC Bank".to_owned(), "Exact Standard".to_owned())];

        let report = run_normalize_institution(&[edited.clone(), edited.clone()], &mapping);
        assert_eq!(report.files, 2);
        assert_eq!(report.branch_sheets, 6);
        assert_eq!(report.mapped_sheets, 1);
        assert_eq!(report.mapped_rows, 2);
        assert_eq!(report.failed_files, 0, "{:?}", report.errors);
        std::fs::remove_file(edited).unwrap();
    }
}
