use anyhow::{bail, Context, Result};
use chrono::Local;
use pytools_core::archive_ops::{
    archive_by_date_copy, archive_by_date_preview, archive_by_type_copy, archive_by_type_preview,
    execute_archive_plan, ArchiveResult,
};
use pytools_core::batch_ops::{
    build_quarterly_rural_loan_plan, build_remove_foreign_plan, build_rural_loan_plan,
    build_village_split_plan,
};
use pytools_core::batch_rename::{
    run_batch_rename_with_progress as run_batch_rename_business, BatchRenameProgressEvent,
};
use pytools_core::config::{
    load_dedup_tasks, load_extract_tasks, load_foreign_banks, load_global,
    load_institution_mapping, load_timeline_rules, GlobalConfig,
};
use pytools_core::config_admin::{build_config_repair_plan, initialize_new_config, view_config};
use pytools_core::conversion::{
    convert_excel_files_with_progress, convert_one, convert_word_files_with_progress,
    ConversionProgressEvent,
};
use pytools_core::dedup_ops::{
    build_append_by_config_plans, build_comment_duplicate_plan, build_config_duplicate_plan,
    precheck_dedup_tasks, write_config_duplicate_check_candidate,
    write_config_duplicate_delete_candidate, write_existing_append_candidate,
    write_new_append_candidate, AppendTargetPlan, DuplicateAction,
};
use pytools_core::edit_plan::{apply_plan_file, apply_plan_with_capability_file};
use pytools_core::file_commit::{
    copy_workbook_to_new_target, create_target_from_candidate, replace_target_from_candidate,
    sha256_file, verify_planning_snapshot,
};
use pytools_core::file_sort_prefix::{
    run_content_prefix_rename_xlsx_candidate,
    run_content_prefix_rename_xlsx_candidate_with_progress,
    run_content_prefix_unrename_xlsx_candidate,
    run_content_prefix_unrename_xlsx_candidate_with_progress,
    run_file_sort_prefix_with_progress as run_file_sort_prefix_with_progress_business,
    run_remove_file_sort_prefix_with_progress as run_remove_file_sort_prefix_with_progress_business,
    FileSortProgressEvent,
};
use pytools_core::local_filename_map::{
    run_local_filename_map_candidate, run_local_filename_map_with_progress,
    LocalFilenameMapProgressEvent,
};
use pytools_core::print_ops::{
    build_comment_print_plan, build_config_print_plans, precheck_print_tasks,
    write_static_print_plans,
};
use pytools_core::quick::{open_config, open_latest_output, open_output_directory};
use pytools_core::region_sum::{run_region_sum_check_with_progress, RegionSumRunEvent};
use pytools_core::sheet_extract::run_extract_sheets_candidate_with_progress;
use pytools_core::sheet_rename::{build_sheet_rename_plan, run_batch_sheet_rename};
use pytools_core::source_discovery::discover_configured_source_files;
use pytools_core::submission_check::{
    run_submission_check, run_submission_check_with_progress, SubmissionCheckProgressEvent,
    SubmissionCheckReport,
};
use pytools_core::summary_ops::{
    summarize_by_comments_with_progress, summarize_by_named_ranges_with_progress,
    summarize_used_ranges_with_progress, SummaryProgress, SummaryRunEvent,
};
use pytools_core::timeline_extract::{
    run_timeline_extract, run_timeline_extract_with_progress, TimelineExtractProgress,
    TimelineExtractResult, TimelineExtractRunEvent,
};
use pytools_core::timeline_ops::{
    prepare_wide_target_replacement_source, run_header_path_compare_with_progress,
    run_timeline_precheck, run_wide_summary_with_progress,
    validate_wide_target_replacement_candidate, HeaderPathCompareProgress,
    HeaderPathCompareRunEvent, WideSummaryProgress, WideSummaryRunEvent,
};
use std::fs::{self, File, OpenOptions};
use std::io::{self, BufRead, IsTerminal, Write};
use std::path::{Path, PathBuf};
use std::time::{Instant, SystemTime, UNIX_EPOCH};

thread_local! {
    static TERMINAL_BATCH_LINE: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

fn finish_terminal_batch_line() {
    if TERMINAL_BATCH_LINE.with(|width| width.replace(0) > 0) {
        eprintln!();
    }
}

fn emit_batch_event_view(
    log_view: &str,
    total: usize,
    current: Option<usize>,
    status: &str,
    phase_change: bool,
    message: &str,
    event_value: Option<&serde_json::Value>,
    elapsed_ms: Option<u128>,
) {
    use pytools_core::run_logging::{BatchEventKind, BatchLogPolicy, Level};

    if log_view == "developer" {
        finish_terminal_batch_line();
        if let Some(value) = event_value {
            if let Ok(line) = serde_json::to_string(value) {
                eprintln!("{line}");
            }
        }
        return;
    }
    if log_view == "off" {
        finish_terminal_batch_line();
        return;
    }

    let kind = if matches!(status, "failed" | "error" | "partial_failure") {
        BatchEventKind::Failure
    } else if matches!(
        status,
        "warning" | "business_issue" | "mismatch" | "partial"
    ) {
        BatchEventKind::BusinessIssue
    } else if phase_change {
        BatchEventKind::PhaseChange
    } else {
        BatchEventKind::Routine
    };
    let mut policy = BatchLogPolicy::new(total);
    if let Some(elapsed) = elapsed_ms {
        policy = policy.with_elapsed_ms(elapsed);
    }
    let interactive_terminal = io::stderr().is_terminal();
    let should_emit = if interactive_terminal {
        policy.should_emit_interactive(current, kind)
    } else {
        policy.should_emit(current, kind)
    };
    if !should_emit {
        return;
    }
    let level = match kind {
        BatchEventKind::Failure => Level::Error,
        BatchEventKind::BusinessIssue => Level::Warn,
        BatchEventKind::Routine | BatchEventKind::PhaseChange => Level::Info,
    };
    let inline =
        total > 8 && interactive_terminal && kind == BatchEventKind::Routine && current.is_some();
    if inline {
        let line = pytools_core::run_logging::terminal_line(
            level,
            &policy.format_progress(current.unwrap_or_default(), message),
        );
        let previous = TERMINAL_BATCH_LINE.with(|width| width.get());
        let width = line.chars().count();
        eprint!("\r{line}{}", " ".repeat(previous.saturating_sub(width)));
        let _ = io::stderr().flush();
        TERMINAL_BATCH_LINE.with(|last| last.set(width));
        if current.is_some_and(|index| index >= total) {
            finish_terminal_batch_line();
        }
        return;
    }
    finish_terminal_batch_line();
    let text = if total > 8 {
        current.map_or_else(
            || message.to_owned(),
            |index| policy.format_progress(index, message),
        )
    } else {
        message.to_owned()
    };
    eprintln!("{}", pytools_core::run_logging::terminal_line(level, &text));
}

fn emit_unified_terminal_summary_with_outputs(
    summary: &pytools_core::run_logging::RunSummary,
    output_paths: &[PathBuf],
) {
    use pytools_core::run_logging::{
        format_terminal_summary, format_terminal_summary_with_outputs, is_summary_frame_line,
        summary_line_level, terminal_line,
    };
    finish_terminal_batch_line();
    let paths = output_paths
        .iter()
        .map(|path| display_path_for_user(path))
        .collect::<Vec<_>>();
    let lines = if paths.is_empty() {
        format_terminal_summary(summary, 8)
    } else {
        format_terminal_summary_with_outputs(summary, 8, &paths)
    };
    for line in lines {
        if is_summary_frame_line(&line) {
            // Decorative rules and blank separators read better untagged.
            eprintln!("{line}");
            continue;
        }
        let level = summary_line_level(&line, summary.overall_status);
        eprintln!("{}", terminal_line(level, &line));
    }
}

fn write_unified_terminal_summary<W: Write>(
    output: &mut W,
    summary: &pytools_core::run_logging::RunSummary,
    output_paths: &[PathBuf],
) -> io::Result<()> {
    use pytools_core::run_logging::{
        format_terminal_summary, format_terminal_summary_with_outputs, is_summary_frame_line,
        summary_line_level, terminal_line,
    };
    finish_terminal_batch_line();
    let paths = output_paths
        .iter()
        .map(|path| display_path_for_user(path))
        .collect::<Vec<_>>();
    let lines = if paths.is_empty() {
        format_terminal_summary(summary, 8)
    } else {
        format_terminal_summary_with_outputs(summary, 8, &paths)
    };
    for line in lines {
        if is_summary_frame_line(&line) {
            writeln!(output, "{line}")?;
            continue;
        }
        let level = summary_line_level(&line, summary.overall_status);
        writeln!(output, "{}", terminal_line(level, &line))?;
    }
    Ok(())
}

/// Serialized form of a summary event as shown on the developer terminal:
/// the raw event plus the display `level` and `event_name` fields, so the
/// history file and the stderr stream never drift apart.
fn summary_event_json(event: &SummaryRunEvent) -> serde_json::Value {
    let level = match event.status.as_str() {
        "failed" => pytools_core::run_logging::Level::Error,
        "warning" | "partial" | "business_issue" => pytools_core::run_logging::Level::Warn,
        _ => pytools_core::run_logging::Level::Info,
    };
    let mut value = serde_json::to_value(event).unwrap_or(serde_json::Value::Null);
    if value.is_object() {
        value["level"] = serde_json::json!(level);
        value["event_name"] = serde_json::json!(event.phase);
    }
    value
}

fn emit_summary_batch_event_view(log_view: &str, event: &SummaryRunEvent) {
    let current = event.progress.current_input_index;
    let phase_change = current.is_none()
        || matches!(
            event.phase.as_str(),
            "run" | "row_extract" | "output_save" | "output_write"
        );
    let level = match event.status.as_str() {
        "failed" => pytools_core::run_logging::Level::Error,
        "warning" | "partial" | "business_issue" => pytools_core::run_logging::Level::Warn,
        _ => pytools_core::run_logging::Level::Info,
    };
    let mut value = serde_json::to_value(event).ok();
    if let Some(value) = value.as_mut() {
        value["level"] = serde_json::json!(level);
        value["event_name"] = serde_json::json!(event.phase);
    }
    emit_batch_event_view(
        log_view,
        event.progress.total_inputs,
        current,
        &event.status,
        phase_change,
        &event.user_message,
        value.as_ref(),
        Some(event.total_elapsed_ms),
    );
}

fn main() {
    if let Err(error) = run() {
        eprintln!(
            "{}",
            pytools_core::run_logging::terminal_line(
                pytools_core::run_logging::Level::Error,
                // 输出完整错误链：阶段上下文之外还必须暴露根因（如"目标在计划后已出现"）。
                &format!("命令执行失败：{error:#}")
            )
        );
        std::process::exit(2);
    }
}

fn run() -> Result<()> {
    let mut arguments = std::env::args().skip(1);
    let Some(command) = arguments.next() else {
        return run_interactive_menu(Vec::new());
    };
    if command == "--help" || command == "-h" {
        print_usage();
        return Ok(());
    }
    let arguments = arguments.collect::<Vec<_>>();
    if command.starts_with('-') {
        let mut menu_arguments = Vec::with_capacity(arguments.len() + 1);
        menu_arguments.push(command);
        menu_arguments.extend(arguments);
        if menu_arguments
            .iter()
            .any(|argument| matches!(argument.as_str(), "--help" | "-h"))
        {
            print_usage();
            return Ok(());
        }
        return run_interactive_menu(menu_arguments);
    }
    if arguments
        .iter()
        .any(|argument| matches!(argument.as_str(), "--help" | "-h"))
    {
        print_usage();
        return Ok(());
    }
    match command.as_str() {
        "menu" | "interactive" => run_interactive_menu(arguments),
        "timeline-precheck" => run_timeline_precheck_cli(arguments),
        "timeline-compare" => run_timeline_compare_cli(arguments, false),
        "timeline-extract" => run_timeline_extract_cli(arguments),
        "wide-summary" => run_wide_summary_cli(arguments, false),
        "wide-summary-row-suffix" => run_wide_summary_cli(arguments, true),
        "summary-used-range" => run_summary_used_range_cli(arguments, false),
        "summary-comments" => run_summary_comments_cli(arguments, false),
        "summary-named-ranges" => run_summary_named_ranges_cli(arguments, false),
        "summary-named-validation" => run_named_validation_extract_cli(arguments, false),
        "archive-date-preview" => run_archive_cli(arguments, "date-preview"),
        "archive-type-preview" => run_archive_cli(arguments, "type-preview"),
        "archive-date-copy" => run_archive_cli(arguments, "date-copy"),
        "archive-type-copy" => run_archive_cli(arguments, "type-copy"),
        "archive-execute-plan" => run_archive_cli(arguments, "plan"),
        "convert-excel" => run_convert_batch_cli(arguments, true, false),
        "convert-word" => run_convert_batch_cli(arguments, false, false),
        "convert-file" => run_convert_file_cli(arguments),
        "config-init" => run_config_init_cli(arguments),
        "config-repair-plan" => run_config_repair_plan_cli(arguments),
        "config-view" => run_config_view_cli(arguments),
        "dedup-comment-plan" => run_dedup_comment_plan(arguments),
        "dedup-config-plan" => run_dedup_config_plan(arguments),
        "dedup-precheck" => run_dedup_precheck(arguments),
        "dedup-append-plan" => run_dedup_append_plan(arguments),
        "dedup-append-execute" => run_dedup_append_execute(arguments),
        "apply-edit-plan" => run_apply_edit_plan(arguments),
        "print-comment-plan" => run_print_comment_plan(arguments),
        "print-config-precheck" => run_print_config_precheck(arguments),
        "print-config-plan" => run_print_config_plan(arguments),
        "village-split-plan" => run_village_split_plan(arguments),
        "village-split" => run_village_split_batch(arguments, false),
        "remove-foreign-bank-plan" => run_remove_foreign_bank_plan(arguments),
        "remove-foreign-bank" => run_remove_foreign_bank_batch(arguments, false),
        "rural-loan-plan" => run_rural_loan_plan(arguments),
        "quarterly-rural-loan-plan" => run_quarterly_rural_loan_plan(arguments),
        "institution-normalize-plan" => run_institution_normalize_plan(arguments),
        "institution-normalize" => run_institution_normalize_batch(arguments, false),
        "fx-header" => run_fx_header_batch(arguments, false),
        "fx-header-plan" => run_fx_header_plan(arguments),
        "rename-sheets-plan" => run_rename_sheets_plan(arguments),
        "rename-sheets" => run_rename_sheets(arguments, false),
        "submission-check" => run_submission_check_cli(arguments, false),
        "open-config" => run_quick_cli(arguments, "config"),
        "open-latest-output" => run_quick_cli(arguments, "latest"),
        "open-output-dir" => run_quick_cli(arguments, "output"),
        "extract-sheets" | "extract-sheets-area" => run_extract_sheets(arguments, false),
        "excel-open-repair" => run_excel_open_repair_cli(arguments, false),
        "region-sum-check" => run_region_sum(arguments, false),
        "file-sort-prefix" => run_file_sort_prefix(arguments, false),
        "remove-file-sort-prefix" => run_remove_file_sort_prefix(arguments, false),
        "rename-by-content" => run_content_prefix_rename_cli(arguments, false),
        "unrename-by-content" => run_content_prefix_unrename_cli(arguments, false),
        "rename-local-map" => run_local_filename_map(arguments),
        "rename-batch" => run_batch_rename(arguments, false),
        _ => {
            print_usage();
            bail!("unsupported command {command:?}")
        }
    }
}

fn run_interactive_menu(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    anyhow::ensure!(
        options.positional.is_empty(),
        "menu accepts options only; pass workbook paths after selecting a menu item"
    );
    let config_path = options
        .values
        .get("config")
        .map(PathBuf::from)
        .unwrap_or_else(default_app_config_path);
    // Interactive menus use the bundled Openpyxl worker for existing XLSX
    // edits. Legacy capability files remain accepted by the parser for CLI
    // compatibility, but are not required to enter the menus.
    let (mut global, mut config_error) = match load_global(&config_path) {
        Ok(global) => (global, None),
        Err(error) => (
            fallback_menu_global(&config_path)?,
            Some(format!("{error:#}")),
        ),
    };
    if let Some(log_dir) = options.values.get("log-dir") {
        let log_dir = PathBuf::from(log_dir);
        global.log_dir = log_dir.clone();
        global.log_dir_override = Some(log_dir);
    }
    let log_view = options
        .values
        .get("log-view")
        .cloned()
        .unwrap_or_else(|| global.log_view.clone());
    anyhow::ensure!(
        matches!(log_view.as_str(), "user" | "developer" | "off"),
        "--log-view must be user, developer, or off"
    );
    if let Err(error) = pytools_core::run_logging::retain_completed_events(
        &global.log_dir.join("run-events"),
        global.log_retention_days,
        global.log_max_bytes,
    ) {
        eprintln!(
            "{}",
            pytools_core::run_logging::terminal_line(
                pytools_core::run_logging::Level::Warn,
                &format!("日志保留策略执行失败：{error}")
            )
        );
    }

    let stdin = io::stdin();
    let allow_native_dialog = stdin.is_terminal();
    let mut input = stdin.lock();
    let stdout = io::stdout();
    let mut output = stdout.lock();
    run_interactive_main_menu(
        &mut input,
        &mut output,
        &config_path,
        &mut global,
        &mut config_error,
        &log_view,
        allow_native_dialog,
        None,
        None,
    )
}

fn fallback_menu_global(config_path: &Path) -> Result<GlobalConfig> {
    let base = config_path
        .parent()
        .and_then(Path::parent)
        .unwrap_or_else(|| Path::new("."));
    let base = if base.is_absolute() {
        base.to_path_buf()
    } else {
        std::env::current_dir()?.join(base)
    };
    Ok(GlobalConfig::new(
        base.join("input"),
        base.join("output"),
        pytools_core::run_logging::default_log_dir(),
    ))
}

/// 版本与构建时刻（让用户一眼确认当前运行的是哪一次构建）。
fn build_banner() -> String {
    let stamp = option_env!("PYTOOLS_BUILD_UTC")
        .and_then(|seconds| seconds.parse::<i64>().ok())
        .and_then(|seconds| chrono::DateTime::from_timestamp(seconds, 0))
        .map(|utc| {
            utc.with_timezone(&Local)
                .format("%Y-%m-%d %H:%M")
                .to_string()
        })
        .unwrap_or_else(|| "本机增量构建".to_owned());
    format!("  v{}  构建 {}", env!("CARGO_PKG_VERSION"), stamp)
}

/// Locate the application configuration without any launcher script, so a
/// double-click on dist/RustTools.exe finds the project-root config.xlsx.
/// Search order (first existing wins):
///   1. PYTOOLS_CONFIG_PATH environment override
///   2. <exe dir>/config/config.xlsx        (packaged layout)
///   3. <exe dir>/config.xlsx
///   4. <exe dir>/../config/config.xlsx     (repo worktree layout)
///   5. <exe dir>/../config.xlsx
///   6. <cwd>/config/config.xlsx
///   7. <cwd>/config.xlsx
/// When nothing exists, return the exe-parent candidate so 7-1 初始化 places
/// a new configuration next to the project data instead of inside dist/.
fn default_app_config_path() -> PathBuf {
    if let Some(config_path) = std::env::var_os("PYTOOLS_CONFIG_PATH") {
        return PathBuf::from(config_path);
    }

    let exe_dir = std::env::current_exe()
        .ok()
        .and_then(|executable| executable.parent().map(Path::to_path_buf))
        .unwrap_or_else(|| PathBuf::from("."));
    let exe_parent = exe_dir
        .parent()
        .map(Path::to_path_buf)
        .unwrap_or_else(|| exe_dir.clone());
    let cwd = std::env::current_dir().unwrap_or_else(|_| PathBuf::from("."));

    let mut candidates = Vec::new();
    for base in [exe_dir.clone(), exe_parent.clone(), cwd.clone()] {
        candidates.push(base.join("config").join("config.xlsx"));
        candidates.push(base.join("config.xlsx"));
    }
    if let Some(found) = candidates.iter().find(|candidate| candidate.is_file()) {
        return found.clone();
    }

    // Default for 初始化/修复配置: the executable's own directory (portable
    // install layout) — never a parent that the install does not control.
    candidates[0].clone()
}

fn run_interactive_main_menu<R, W>(
    input: &mut R,
    output: &mut W,
    config_path: &Path,
    global: &mut GlobalConfig,
    config_error: &mut Option<String>,
    log_view: &str,
    allow_native_dialog: bool,
    capabilities: Option<&Path>,
    platform: Option<&str>,
) -> Result<()>
where
    R: BufRead,
    W: Write,
{
    loop {
        writeln!(output, "======== rustTools 菜单 ========")?;
        writeln!(output, "{}", build_banner())?;
        writeln!(output, "  1  数据提取与汇总")?;
        writeln!(output, "  2  数据去重与追加")?;
        writeln!(output, "  3  报表打印")?;
        writeln!(output, "  4  文件处理")?;
        writeln!(output, "  5  专用报表处理")?;
        writeln!(output, "  6  归档整理")?;
        writeln!(output, "  7  配置与快捷操作")?;
        writeln!(output, "  0  退出")?;
        writeln!(output, "==============================")?;
        if let Some(error) = config_error.as_deref() {
            if log_view == "developer" {
                writeln!(
                    output,
                    "[配置不可用] {}；请先使用 7-1 初始化/修复配置表（{error}）",
                    display_path_for_user(&config_path)
                )?;
            } else {
                writeln!(
                    output,
                    "[配置不可用] {}；请先使用 7-1 初始化/修复配置表。",
                    display_path_for_user(&config_path)
                )?;
            }
        }
        let Some(choice) = read_menu_choice(input, output, "主菜单选择> ")? else {
            writeln!(output, "再见。")?;
            return Ok(());
        };
        let category_id = match choice.as_str() {
            "1" => "data.extract_and_summarize",
            "2" => "data.deduplicate_and_append",
            "3" => "reports.static_print",
            "4" => "files.process_and_convert",
            "5" => "reports.specialized",
            "6" => "archive.organize",
            "7" => "configuration.and_quick_actions",
            "0" => {
                writeln!(output, "再见。")?;
                return Ok(());
            }
            _ => {
                writeln!(output, "仅接受 1、2、3、4、5、6、7 或 0")?;
                continue;
            }
        };
        if matches!(
            category_id,
            "data.deduplicate_and_append"
                | "reports.static_print"
                | "files.process_and_convert"
                | "reports.specialized"
                | "archive.organize"
        ) && config_error.is_some()
        {
            writeln!(
                output,
                "当前配置不可用，该工具暂不能运行；请先进入 7 配置与快捷操作。"
            )?;
            continue;
        }
        match category_id {
            "data.extract_and_summarize" => run_interactive_data_menu(
                input,
                output,
                config_path,
                global,
                config_error,
                log_view,
                allow_native_dialog,
                capabilities,
                platform,
            )?,
            "data.deduplicate_and_append" => run_interactive_dedup_menu(
                input,
                output,
                config_path,
                global,
                log_view,
                allow_native_dialog,
                capabilities,
                platform,
            )?,
            "reports.static_print" => run_interactive_print_menu(
                input,
                output,
                config_path,
                global,
                log_view,
                allow_native_dialog,
                capabilities,
                platform,
            )?,
            "files.process_and_convert" => run_interactive_convert_menu(
                input,
                output,
                config_path,
                global,
                log_view,
                allow_native_dialog,
                capabilities,
                platform,
            )?,
            "reports.specialized" => run_interactive_batch_menu(
                input,
                output,
                config_path,
                global,
                log_view,
                allow_native_dialog,
            )?,
            "archive.organize" => run_interactive_archive_menu(
                input,
                output,
                config_path,
                global,
                log_view,
                allow_native_dialog,
            )?,
            "configuration.and_quick_actions" => run_interactive_config_menu(
                input,
                output,
                config_path,
                global,
                config_error,
                log_view,
                capabilities,
                platform,
            )?,
            _ => unreachable!("resolved PyTools V2 category ID"),
        }
    }
}

fn run_interactive_data_menu<R, W>(
    input: &mut R,
    output: &mut W,
    config_path: &Path,
    global: &GlobalConfig,
    config_error: &Option<String>,
    log_view: &str,
    allow_native_dialog: bool,
    capabilities: Option<&Path>,
    platform: Option<&str>,
) -> Result<()>
where
    R: BufRead,
    W: Write,
{
    loop {
        writeln!(output, "------ 数据提取与汇总 ------")?;
        writeln!(
            output,
            "  1  时序宽表汇总  【时序提取规则；路径标准化映射】"
        )?;
        writeln!(output, "  2  按使用区域快速汇总")?;
        writeln!(
            output,
            "  3  按模板命名区域汇总  【表头区域生成列名；表结构区域校验；汇总区域提取数据】"
        )?;
        writeln!(output, "  4  按模板校验区域提取校验  【临时复制源文件并调用 Excel/WPS COM 重算；原文件不修改】")?;
        writeln!(
            output,
            "  5  表头路径比对与映射  【时序提取规则；路径标准化映射】"
        )?;
        writeln!(
            output,
            "  6  时序配置预校验  【时序提取规则；路径标准化映射；全局配置】"
        )?;
        writeln!(output, "  0  返回")?;
        writeln!(output, "--------------------------")?;
        let Some(choice) = read_menu_choice(input, output, "子菜单选择> ")? else {
            return Ok(());
        };

        let feature_id = match choice.as_str() {
            "1" => "data.timeline_wide_summary",
            "2" => "data.used_range_summary",
            "3" => "data.named_range_summary",
            "4" => "data.named_validation_extract",
            "5" => "data.header_path_compare",
            "6" => "data.timeline_precheck",
            "0" => return Ok(()),
            _ => {
                writeln!(output, "仅接受 1、2、3、4、5、6 或 0")?;
                continue;
            }
        };

        match feature_id {
            "data.timeline_wide_summary" => {
                if config_error.is_some() {
                    writeln!(output, "当前配置不可用；请先进入 7-1 初始化/修复配置。")?;
                    return Ok(());
                }
                let Some(sources) = choose_source_files_or_folders_impl(
                    "选择时序宽表汇总源文件或文件夹",
                    NativeFileFilter::ExcelWorkbooks,
                    input,
                    output,
                    allow_native_dialog,
                    true,
                    false,
                )?
                else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let mut arguments = config_cli_arguments(config_path);
                arguments.extend(summary_cli_arguments(global, log_view));
                if let (Some(capabilities), Some(platform)) = (capabilities, platform) {
                    arguments.push("--capabilities".to_owned());
                    arguments.push(capabilities.to_string_lossy().into_owned());
                    arguments.push("--platform".to_owned());
                    arguments.push(platform.to_owned());
                }
                arguments.extend(
                    sources
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_interactive_wide_summary_menu(output, arguments, false)?;
                return Ok(());
            }
            "data.used_range_summary" => {
                let Some(sources) = choose_workbooks(
                    "选择按使用区域快速汇总的源文件（可多选）",
                    true,
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let mut arguments = summary_cli_arguments(global, log_view);
                arguments.extend(
                    sources
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_summary_used_range_cli(arguments, true)?;
                return Ok(());
            }
            "data.named_range_summary" => {
                let Some(template) = choose_workbooks(
                    "选择按模板命名区域汇总的模板工作簿",
                    false,
                    input,
                    output,
                    allow_native_dialog,
                )?
                .and_then(|paths| paths.into_iter().next()) else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let Some(mut sources) = choose_source_files_or_folders_impl(
                    "选择按模板命名区域汇总的源文件或文件夹",
                    NativeFileFilter::ExcelWorkbooks,
                    input,
                    output,
                    allow_native_dialog,
                    true,
                    true,
                )?
                else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let selected_count = sources.len();
                sources.retain(|path| {
                    !path
                        .file_name()
                        .is_some_and(|name| name.to_string_lossy().starts_with("~$"))
                        && !path
                            .file_stem()
                            .is_some_and(|name| name.to_string_lossy().contains("_审核版"))
                        && !path.components().any(|part| part.as_os_str() == "执行结果")
                });
                if sources.len() != selected_count {
                    writeln!(
                        output,
                        "→ 按汇总规则排除 {} 个临时文件、审核副本或执行结果",
                        selected_count - sources.len()
                    )?;
                }
                if sources.is_empty() {
                    writeln!(output, "没有可汇总的源工作簿。")?;
                    return Ok(());
                }
                let mut arguments = vec![
                    "--template".to_owned(),
                    template.to_string_lossy().into_owned(),
                ];
                arguments.extend(summary_cli_arguments(global, log_view));
                arguments.extend(
                    sources
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                match run_summary_named_ranges_cli(arguments, true) {
                    Ok(()) => return Ok(()),
                    Err(error) => {
                        writeln!(output, "按模板命名区域汇总失败：{error:#}")?;
                        continue;
                    }
                }
            }
            "data.named_validation_extract" => {
                let Some(template) = choose_workbooks(
                    "选择按模板校验区域提取校验的模板工作簿",
                    false,
                    input,
                    output,
                    allow_native_dialog,
                )?
                .and_then(|paths| paths.into_iter().next()) else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let Some(mut sources) = choose_source_files_or_folders_impl(
                    "选择需要提取校验结果的源工作簿或文件夹",
                    NativeFileFilter::ExcelWorkbooks,
                    input,
                    output,
                    allow_native_dialog,
                    true,
                    true,
                )?
                else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let selected_count = sources.len();
                sources.retain(|path| {
                    !path
                        .file_name()
                        .is_some_and(|name| name.to_string_lossy().starts_with("~$"))
                        && !path
                            .file_stem()
                            .is_some_and(|name| name.to_string_lossy().contains("_审核版"))
                        && !path.components().any(|part| part.as_os_str() == "执行结果")
                });
                if sources.len() != selected_count {
                    writeln!(
                        output,
                        "→ 按校验规则排除 {} 个临时文件、审核副本或执行结果",
                        selected_count - sources.len()
                    )?;
                }
                if sources.is_empty() {
                    writeln!(output, "没有可校验的源工作簿。")?;
                    return Ok(());
                }
                let mut arguments = vec![
                    "--template".to_owned(),
                    template.to_string_lossy().into_owned(),
                ];
                arguments.extend(summary_cli_arguments(global, log_view));
                arguments.extend(
                    sources
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                match run_named_validation_extract_cli(arguments, true) {
                    Ok(()) => return Ok(()),
                    Err(error) => {
                        writeln!(output, "按模板校验区域提取校验失败：{error:#}")?;
                        continue;
                    }
                }
            }
            "data.header_path_compare" => {
                if config_error.is_some() {
                    writeln!(output, "当前配置不可用；请先进入 7-1 初始化/修复配置。")?;
                    return Ok(());
                }
                let Some(template) = choose_workbooks(
                    "选择表头路径比对模板",
                    false,
                    input,
                    output,
                    allow_native_dialog,
                )?
                .and_then(|paths| paths.into_iter().next()) else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let Some(sources) = choose_workbooks(
                    "选择表头路径比对源文件（可多选）",
                    true,
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let mut arguments = config_cli_arguments(config_path);
                arguments.extend([
                    "--log-dir".to_owned(),
                    global.log_dir.to_string_lossy().into_owned(),
                    "--log-view".to_owned(),
                    log_view.to_owned(),
                    "--template".to_owned(),
                    template.to_string_lossy().into_owned(),
                ]);
                arguments.extend(
                    sources
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_timeline_compare_cli(arguments, true)?;
                return Ok(());
            }
            "data.timeline_precheck" => {
                if config_error.is_some() {
                    writeln!(output, "当前配置不可用；请先进入 7-1 初始化/修复配置。")?;
                    return Ok(());
                }
                run_interactive_timeline_precheck(output, config_path, global, log_view)?;
                return Ok(());
            }
            _ => unreachable!("resolved stable PyTools feature ID"),
        }
    }
}

fn run_interactive_quick_menu<R, W>(
    input: &mut R,
    output: &mut W,
    config_path: &Path,
    global: &GlobalConfig,
) -> Result<()>
where
    R: BufRead,
    W: Write,
{
    loop {
        writeln!(output, "------ 快捷工具 ------")?;
        writeln!(output, "  1  打开配置文件（config.xlsx）")?;
        writeln!(output, "  2  打开 output 最新文件")?;
        writeln!(output, "  3  打开 output 文件夹")?;
        writeln!(output, "  0  返回")?;
        writeln!(output, "----------------------")?;
        let Some(choice) = read_menu_choice(input, output, "子菜单选择> ")? else {
            return Ok(());
        };
        match choice.as_str() {
            "1" => match open_config(config_path) {
                Ok(()) => writeln!(
                    output,
                    "[完成] 已打开配置文件: {}",
                    display_path_for_user(&config_path)
                )?,
                Err(error) => writeln!(output, "[失败] 打开配置文件失败: {error:#}")?,
            },
            "2" => match open_latest_output(&global.output_dir) {
                Ok(Some(path)) => writeln!(
                    output,
                    "[完成] 已打开最新输出文件: {}",
                    display_path_for_user(&path)
                )?,
                Ok(None) => writeln!(
                    output,
                    "[提示] 未找到输出文件: {}",
                    display_path_for_user(&global.output_dir)
                )?,
                Err(error) => writeln!(output, "[失败] 打开输出文件失败: {error:#}")?,
            },
            "3" => match open_output_directory(&global.output_dir) {
                Ok(()) => writeln!(
                    output,
                    "[完成] 已打开输出文件夹: {}",
                    display_path_for_user(&global.output_dir)
                )?,
                Err(error) => writeln!(output, "[失败] 打开输出文件夹失败: {error:#}")?,
            },
            "0" => return Ok(()),
            _ => writeln!(output, "仅接受 1、2、3 或 0")?,
        }
        return Ok(());
    }
}

#[derive(Debug)]
struct ConfigMenuRunFacts {
    status: &'static str,
    metrics: serde_json::Value,
    stages: Vec<serde_json::Value>,
    actual_engines: Vec<String>,
    actual_components: Vec<String>,
    excel_modifier_invoked: bool,
    excel_writer_invoked: bool,
    error_type: Option<&'static str>,
}

impl ConfigMenuRunFacts {
    fn new() -> Self {
        Self {
            status: "succeeded",
            metrics: serde_json::json!({}),
            stages: Vec::new(),
            actual_engines: Vec::new(),
            actual_components: Vec::new(),
            excel_modifier_invoked: false,
            excel_writer_invoked: false,
            error_type: None,
        }
    }

    fn add_engine(&mut self, engine: impl Into<String>) {
        let engine = engine.into();
        if !self.actual_engines.contains(&engine) {
            self.actual_engines.push(engine);
        }
    }

    fn add_component(&mut self, component: impl Into<String>) {
        let component = component.into();
        if !self.actual_components.contains(&component) {
            self.actual_components.push(component);
        }
    }

    fn add_stage(&mut self, name: &str, elapsed: std::time::Duration) {
        self.stages.push(serde_json::json!({
            "name": name,
            "elapsed_ms": elapsed.as_millis(),
        }));
    }
}

fn config_unified_status(facts: &ConfigMenuRunFacts) -> pytools_core::run_logging::TaskStatus {
    use pytools_core::run_logging::TaskStatus;
    match facts.status {
        "failed" => TaskStatus::Failed,
        "partial" => TaskStatus::PartialSuccess,
        _ if facts.metrics["operation"] == "already_complete" => TaskStatus::NoOp,
        _ if facts.metrics["operation"] == "candidate_retained" => TaskStatus::Cancelled,
        _ => TaskStatus::Success,
    }
}

fn finish_config_menu_action<W: Write>(
    output: &mut W,
    global: &GlobalConfig,
    log_view: &str,
    feature_id: &str,
    action: &str,
    label: &str,
    started_at: &chrono::DateTime<Local>,
    elapsed_ms: u128,
    facts: &ConfigMenuRunFacts,
) -> Result<()> {
    if log_view == "user" {
        if facts.status == "failed" {
            writeln!(
                output,
                "{label}失败，耗时 {elapsed_ms_display_seconds:.3} 秒。",
                elapsed_ms_display_seconds = elapsed_ms as f64 / 1000.0
            )?;
        } else if facts.status == "partial" {
            writeln!(
                output,
                "{label}完成，但部分配置未能解析：{}，耗时 {elapsed_ms_display_seconds:.3} 秒。",
                config_menu_summary(feature_id, &facts.metrics),
                elapsed_ms_display_seconds = elapsed_ms as f64 / 1000.0
            )?;
        } else {
            writeln!(
                output,
                "{label}完成：{}，耗时 {elapsed_ms_display_seconds:.3} 秒。",
                config_menu_summary(feature_id, &facts.metrics),
                elapsed_ms_display_seconds = elapsed_ms as f64 / 1000.0
            )?;
        }
    } else if log_view == "developer" {
        let history_result = write_config_menu_run_history(
            global, feature_id, action, started_at, elapsed_ms, facts,
        );
        match history_result {
            Ok(path) => writeln!(
                output, "PT-5 完成：{feature_id}，状态 {}，{}，实际引擎：{}，实际组件：{}，ExcelModifier {}，Excel writer {}，耗时 {elapsed_ms_display_seconds:.3} 秒；运行记录：{}",
                facts.status,
                config_menu_summary(feature_id, &facts.metrics),
                if facts.actual_engines.is_empty() { "未调用".to_owned() } else { facts.actual_engines.join(", ") },
                if facts.actual_components.is_empty() { "未确认".to_owned() } else { facts.actual_components.join(", ") },
                if facts.excel_modifier_invoked { "已调用" } else { "未调用" },
                if facts.excel_writer_invoked { "已调用" } else { "未调用" },
                display_path_for_user(&path)
            , elapsed_ms_display_seconds = elapsed_ms as f64 / 1000.0)?,
            Err(error) => writeln!(
                output,
                "PT-5 完成：{feature_id}，状态 {}，{}，运行记录写入失败：{error:#}",
                facts.status,
                config_menu_summary(feature_id, &facts.metrics)
            )?,
        }
    }
    Ok(())
}

fn config_menu_summary(feature_id: &str, metrics: &serde_json::Value) -> String {
    if feature_id == "PT-5-2" {
        return format!(
            "读取 {} 张工作表、{} 行",
            metrics["sheet_count"].as_u64().unwrap_or_default(),
            metrics["non_empty_rows"].as_u64().unwrap_or_default()
        );
    }
    match metrics["operation"]
        .as_str()
        .unwrap_or("configuration_action")
    {
        "initialized" => format!(
            "新建配置表 {} 张工作表",
            metrics["sheets"].as_u64().unwrap_or_default()
        ),
        "already_complete" => "配置完整，无需修复".to_owned(),
        "repair_plan" => format!(
            "生成修复计划 {} 项操作",
            metrics["planned_operations"].as_u64().unwrap_or_default()
        ),
        "candidate_ready" => "候选配置已验证".to_owned(),
        "committed" => "配置已安全更新".to_owned(),
        "candidate_retained" => "候选已保留，当前配置未替换".to_owned(),
        _ => "配置操作未完成".to_owned(),
    }
}

fn write_config_menu_run_history(
    global: &GlobalConfig,
    feature_id: &str,
    action: &str,
    started_at: &chrono::DateTime<Local>,
    elapsed_ms: u128,
    facts: &ConfigMenuRunFacts,
) -> Result<PathBuf> {
    fs::create_dir_all(&global.log_dir).with_context(|| {
        format!(
            "create config log directory {}",
            display_path_for_user(&global.log_dir)
        )
    })?;
    let mut stages = facts.stages.clone();
    if stages.is_empty() {
        stages.push(serde_json::json!({
            "name": "configuration_menu_action",
            "elapsed_ms": elapsed_ms,
        }));
    }
    let history = serde_json::json!({
        "schema_version": 1,
        "feature_id": feature_id,
        "action": action,
        "status": facts.status,
        "started_at": started_at.to_rfc3339(),
        "elapsed_ms": elapsed_ms,
        "stages": stages,
        "metrics": facts.metrics,
        "actual_engines": facts.actual_engines,
        "actual_components": facts.actual_components,
        "excel_modifier_invoked": facts.excel_modifier_invoked,
        "excel_writer_invoked": facts.excel_writer_invoked,
        "plan_only": facts.metrics["execution_mode"] == "plan_only",
        "error_type": facts.error_type,
    });
    let timestamp = Local::now().format("%Y%m%d_%H%M%S_%f");
    for suffix in 1..=1000 {
        let suffix = if suffix == 1 {
            String::new()
        } else {
            format!("-{suffix}")
        };
        let path = global
            .log_dir
            .join(format!("pytools-config-{action}-{timestamp}{suffix}.json"));
        match write_json_report_create_new(&path, &history) {
            Ok(()) => return Ok(path),
            Err(_error) if path.exists() => continue,
            Err(error) => return Err(error),
        }
    }
    bail!("could not allocate a unique configuration run history filename")
}

fn run_interactive_config_menu<R, W>(
    input: &mut R,
    output: &mut W,
    config_path: &Path,
    global: &mut GlobalConfig,
    config_error: &mut Option<String>,
    log_view: &str,
    capabilities: Option<&Path>,
    platform: Option<&str>,
) -> Result<()>
where
    R: BufRead,
    W: Write,
{
    loop {
        writeln!(output, "------ 配置与快捷操作 ------")?;
        writeln!(output, "  1  初始化 / 修复配置（需二次确认）")?;
        writeln!(output, "  2  重新加载并查看配置状态")?;
        writeln!(output, "  3  打开配置文件")?;
        writeln!(output, "  4  打开最近输出")?;
        writeln!(output, "  5  打开输出目录")?;
        writeln!(output, "  0  返回")?;
        writeln!(output, "--------------------------")?;
        let Some(choice) = read_menu_choice(input, output, "子菜单选择> ")? else {
            return Ok(());
        };
        let action_id = match choice.as_str() {
            "1" => "config.initialize_or_repair",
            "2" => "config.reload_and_view",
            "3" => "quick.open_config",
            "4" => "quick.open_latest_output",
            "5" => "quick.open_output_folder",
            "0" => return Ok(()),
            _ => {
                writeln!(output, "仅接受 1、2、3、4、5 或 0")?;
                continue;
            }
        };
        match action_id {
            "config.initialize_or_repair" => {
                let mut diagnostic =
                    ConfiguredRunLog::new(global, "7-1", "初始化或修复配置", log_view);
                writeln!(
                    output,
                    "[二次确认] 7-1 将新建配置表或生成配置修复候选；如后续确认替换，当前配置文件会被更新。"
                )?;
                writeln!(output, "配置路径：{}", display_path_for_user(&config_path))?;
                writeln!(output, "  1  确认继续")?;
                writeln!(output, "  0  取消，不修改配置")?;
                let Some(confirmation) = read_menu_choice(input, output, "确认执行 7-1> ")?
                else {
                    diagnostic.finish_as(
                        false,
                        None,
                        &[],
                        &[],
                        Some(pytools_core::run_logging::TaskStatus::Cancelled),
                    );
                    return Ok(());
                };
                if confirmation != "1" {
                    diagnostic.finish_as(
                        false,
                        None,
                        &[],
                        &[],
                        Some(pytools_core::run_logging::TaskStatus::Cancelled),
                    );
                    writeln!(output, "[已取消] 配置文件未修改。")?;
                    continue;
                }
                let started_at = Local::now();
                let started = Instant::now();
                let mut facts = ConfigMenuRunFacts::new();
                if log_view == "user" {
                    writeln!(output, "开始初始化/修复配置表")?;
                } else if log_view == "developer" {
                    writeln!(output, "PT-5 开始：PT-5-1 initialize_or_repair")?;
                }
                if let Err(error) = run_interactive_config_initialize_or_repair(
                    input,
                    output,
                    config_path,
                    global,
                    config_error,
                    log_view,
                    capabilities,
                    platform,
                    &mut facts,
                ) {
                    facts.status = "failed";
                    facts.error_type = Some("config_initialize_or_repair_failed");
                    if log_view == "developer" {
                        writeln!(output, "[配置操作失败] {error:#}")?;
                    } else {
                        writeln!(
                            output,
                            "[配置操作失败] 请检查配置文件是否为有效的 .xlsx，再重试或联系开发者。"
                        )?;
                    }
                }
                diagnostic.record_stages(&facts.stages);
                let status = config_unified_status(&facts);
                diagnostic.finish_as(facts.status == "failed", None, &[], &[], Some(status));
                finish_config_menu_action(
                    output,
                    global,
                    log_view,
                    "PT-5-1",
                    "initialize_or_repair",
                    "初始化/修复配置表",
                    &started_at,
                    started.elapsed().as_millis(),
                    &facts,
                )?;
            }
            "config.reload_and_view" => {
                let mut diagnostic = ConfiguredRunLog::new(global, "7-2", "重新加载配置", log_view);
                let started_at = Local::now();
                let view_started = Instant::now();
                let mut facts = ConfigMenuRunFacts::new();
                let had_last_good_config = config_error.is_none();
                if log_view == "user" {
                    writeln!(output, "开始读取并查看当前配置")?;
                } else if log_view == "developer" {
                    writeln!(output, "PT-5 开始：PT-5-2 view")?;
                }
                facts.add_engine("calamine");
                facts.add_component("pytools_core::config::view_config");
                match view_config(config_path) {
                    Ok(view) => {
                        let refreshed_global = load_global(config_path);
                        facts.add_engine("calamine");
                        facts.add_component("pytools_core::config::view_config");
                        let non_empty_rows = view
                            .sheets
                            .iter()
                            .map(|sheet| {
                                sheet
                                    .rows
                                    .iter()
                                    .skip(1)
                                    .filter(|row| {
                                        row.iter().any(|value| {
                                            !value.is_null()
                                                && value
                                                    .as_str()
                                                    .is_none_or(|text| !text.is_empty())
                                        })
                                    })
                                    .count() as u64
                            })
                            .sum::<u64>();
                        facts.metrics = serde_json::json!({
                            "sheet_count": view.sheets.len(),
                            "non_empty_rows": non_empty_rows,
                            "global_config_loaded": false,
                        });
                        writeln!(output, "配置文件: {}", display_path_for_user(&config_path))?;
                        if let Ok(modified) =
                            fs::metadata(config_path).and_then(|metadata| metadata.modified())
                        {
                            let modified = chrono::DateTime::<Local>::from(modified);
                            writeln!(output, "最后修改: {}", modified.format("%Y-%m-%d %H:%M:%S"))?;
                        }
                        writeln!(output, "Sheet 列表（{}）:", view.sheets.len())?;
                        for sheet in &view.sheets {
                            let row_count = sheet
                                .rows
                                .iter()
                                .skip(1)
                                .filter(|row| {
                                    row.iter().any(|value| {
                                        !value.is_null()
                                            && value.as_str().is_none_or(|text| !text.is_empty())
                                    })
                                })
                                .count();
                            let column_count = sheet.rows.first().map(Vec::len).unwrap_or(0);
                            writeln!(
                                output,
                                "  - {}: {} 行 × {} 列",
                                sheet.name, row_count, column_count
                            )?;
                        }
                        if log_view == "developer" {
                            writeln!(output, "配置值不写入运行日志。")?;
                            writeln!(output, "实际读取组件: excel-reader / Calamine")?;
                        } else {
                            facts.add_component("pytools_core::config::load_global");
                            match &refreshed_global {
                                Ok(config) => {
                                    facts.metrics["global_config_loaded"] = serde_json::json!(true);
                                    writeln!(output, "全局配置解析:")?;
                                    writeln!(
                                        output,
                                        "  输入目录: {}",
                                        display_path_for_user(&config.input_dir)
                                    )?;
                                    writeln!(
                                        output,
                                        "  输出目录: {}",
                                        display_path_for_user(&config.output_dir)
                                    )?;
                                    writeln!(
                                        output,
                                        "  日志目录: {}",
                                        display_path_for_user(&config.log_dir)
                                    )?;
                                    writeln!(output, "  源文件扩展名: {:?}", config.source_exts)?;
                                }
                                Err(error) => {
                                    facts.status = if had_last_good_config {
                                        "partial"
                                    } else {
                                        "failed"
                                    };
                                    if log_view == "developer" {
                                        if had_last_good_config {
                                            writeln!(output, "  [全局配置解析失败] {error:#}；继续使用上次成功加载的配置。")?;
                                        } else {
                                            writeln!(output, "  [全局配置解析失败] {error:#}")?;
                                        }
                                    } else {
                                        facts.error_type = Some("global_config_parse_failed");
                                        if had_last_good_config {
                                            writeln!(output, "  [全局配置解析失败] 磁盘配置不可用，继续使用上次成功加载的配置。")?;
                                        } else {
                                            writeln!(output, "  [全局配置解析失败] 请使用 7-1 初始化/修复配置表。")?;
                                        }
                                    }
                                }
                            }
                            facts.add_component("pytools_core::config::load_timeline_rules");
                            if let Ok(rules) = load_timeline_rules(config_path) {
                                let enabled =
                                    rules.iter().filter(|rule| rule.enabled).collect::<Vec<_>>();
                                writeln!(output, "时序提取规则（启用）: {} 条", enabled.len())?;
                                for rule in enabled {
                                    writeln!(
                                        output,
                                        "  · {}  wb={}  sheet={}",
                                        rule.name, rule.wb_keyword, rule.sheet_keyword
                                    )?;
                                }
                            }
                            facts.add_component("pytools_core::config::load_institution_mapping");
                            facts.add_component("pytools_core::config::load_foreign_banks");
                            if let (Ok(mapping), Ok(foreign)) = (
                                load_institution_mapping(config_path),
                                load_foreign_banks(config_path),
                            ) {
                                writeln!(
                                    output,
                                    "机构映射: {} 条；外资行: {} 个",
                                    mapping.len(),
                                    foreign.len()
                                )?;
                            }
                            facts.add_component("pytools_core::config::load_extract_tasks");
                            if let Ok(tasks) = load_extract_tasks(config_path) {
                                writeln!(output, "工作表提取任务（启用）: {} 条", tasks.len())?;
                            }
                        }
                        if log_view != "off" {
                            writeln!(
                                output,
                                "用时={:.3} 秒",
                                (view_started.elapsed().as_millis()) as f64 / 1000.0
                            )?;
                        }
                        facts.add_component("pytools_core::config::load_global");
                        match refreshed_global {
                            Ok(mut refreshed) => {
                                // 与 refresh_interactive_global 相同：重载配置
                                // 不得丢弃 --log-dir 显式覆盖。
                                refreshed.log_dir_override = global.log_dir_override.clone();
                                if let Some(log_dir) = &global.log_dir_override {
                                    refreshed.log_dir = log_dir.clone();
                                }
                                *global = refreshed;
                                *config_error = None;
                                facts.metrics["global_config_loaded"] = serde_json::json!(true);
                            }
                            Err(error) => {
                                facts.status = if had_last_good_config {
                                    "partial"
                                } else {
                                    "failed"
                                };
                                facts.error_type = Some("global_config_refresh_failed");
                                facts.metrics["global_config_loaded"] =
                                    serde_json::json!(had_last_good_config);
                                if !had_last_good_config {
                                    *config_error = Some(format!("{error:#}"));
                                }
                                if log_view == "developer" {
                                    if had_last_good_config {
                                        writeln!(output, "[全局配置刷新失败] {error:#}；继续使用上次成功加载的配置。")?;
                                    } else {
                                        writeln!(output, "[全局配置不可用] {error:#}")?;
                                    }
                                } else if had_last_good_config {
                                    writeln!(output, "[全局配置刷新失败] 磁盘配置无效，继续使用上次成功加载的配置。")?;
                                } else {
                                    writeln!(
                                        output,
                                        "[全局配置不可用] 请使用 7-1 初始化/修复配置表。"
                                    )?;
                                }
                            }
                        }
                    }
                    Err(_error) if !config_path.is_file() => {
                        facts.status = if had_last_good_config {
                            "partial"
                        } else {
                            "failed"
                        };
                        facts.error_type = Some(if had_last_good_config {
                            "config_missing_using_last_good"
                        } else {
                            "config_missing"
                        });
                        facts.metrics = serde_json::json!({
                            "sheet_count": 0,
                            "non_empty_rows": 0,
                            "global_config_loaded": had_last_good_config,
                        });
                        if had_last_good_config {
                            writeln!(
                                output,
                                "[缺失] 配置文件 {} 不存在；继续使用上次成功加载的配置。",
                                display_path_for_user(&config_path)
                            )?;
                        } else {
                            writeln!(
                                output,
                                "[缺失] 配置文件 {} 不存在，请先选择 7-1 初始化/修复配置表。",
                                display_path_for_user(&config_path)
                            )?;
                        }
                    }
                    Err(error) => {
                        facts.status = if had_last_good_config {
                            "partial"
                        } else {
                            "failed"
                        };
                        facts.error_type = Some(if had_last_good_config {
                            "config_view_failed_using_last_good"
                        } else {
                            "config_view_failed"
                        });
                        facts.metrics = serde_json::json!({
                            "global_config_loaded": had_last_good_config,
                        });
                        if !had_last_good_config {
                            *config_error = Some(format!("{error:#}"));
                        }
                        if log_view == "developer" {
                            if had_last_good_config {
                                writeln!(
                                    output,
                                    "[读取失败] {error:#}；继续使用上次成功加载的配置。"
                                )?;
                            } else {
                                writeln!(output, "[读取失败] {error:#}")?;
                            }
                        } else if had_last_good_config {
                            writeln!(
                                output,
                                "[读取失败] 当前配置读取失败，继续使用上次成功加载的配置。"
                            )?;
                        } else {
                            writeln!(
                                output,
                                "[读取失败] 请检查配置文件，或使用 7-1 初始化/修复配置表。"
                            )?;
                        }
                    }
                }
                facts.add_stage("configuration_read_and_summary", view_started.elapsed());
                diagnostic.record_stages(&facts.stages);
                let status = config_unified_status(&facts);
                diagnostic.finish_as(facts.status == "failed", None, &[], &[], Some(status));
                finish_config_menu_action(
                    output,
                    global,
                    log_view,
                    "PT-5-2",
                    "view",
                    "查看当前配置",
                    &started_at,
                    view_started.elapsed().as_millis(),
                    &facts,
                )?;
            }
            "quick.open_config" => log_quick_menu_action(
                output,
                global,
                log_view,
                "QUICK-OPEN-CONFIG",
                "打开配置文件",
                open_config(config_path).map(|()| Some(config_path.to_path_buf())),
                "没有可打开的配置文件".to_owned(),
            )?,
            "quick.open_latest_output" => log_quick_menu_action(
                output,
                global,
                log_view,
                "QUICK-OPEN-LATEST",
                "打开最近输出文件",
                open_latest_output(&global.output_dir),
                format!(
                    "输出目录中没有可打开的文件：{}",
                    display_path_for_user(&global.output_dir)
                ),
            )?,
            "quick.open_output_folder" => log_quick_menu_action(
                output,
                global,
                log_view,
                "QUICK-OPEN-OUTPUT",
                "打开输出目录",
                open_output_directory(&global.output_dir).map(|()| Some(global.output_dir.clone())),
                "输出目录没有打开".to_owned(),
            )?,
            _ => unreachable!("resolved PyTools V2 configuration action ID"),
        }
    }
}

fn run_interactive_config_initialize_or_repair<R, W>(
    input: &mut R,
    output: &mut W,
    config_path: &Path,
    global: &mut GlobalConfig,
    config_error: &mut Option<String>,
    log_view: &str,
    capabilities: Option<&Path>,
    platform: Option<&str>,
    facts: &mut ConfigMenuRunFacts,
) -> Result<()>
where
    R: BufRead,
    W: Write,
{
    anyhow::ensure!(
        config_path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
        "rustTools configuration must use the .xlsx format"
    );

    if !config_path.exists() {
        let initialize_started = Instant::now();
        if log_view != "off" {
            writeln!(output, "正在新建 rustTools 配置表……")?;
        }
        facts.add_engine("rust_xlsxwriter");
        facts.add_component("pytools_core::config::initialize_new_config");
        facts.excel_writer_invoked = true;
        let result = initialize_new_config(config_path)?;
        facts.metrics = serde_json::json!({
            "operation": "initialized",
            "execution_mode": "initialized",
            "sheets": result.sheets,
        });
        facts.add_stage("configuration_initialization", initialize_started.elapsed());
        if log_view == "off" {
            writeln!(
                output,
                "[完成] 已新建配置表：{}（{} 张工作表）",
                display_path_for_user(&result.path),
                result.sheets
            )?;
        } else {
            writeln!(
                output,
                "[完成] 已新建配置表：{}（{} 张工作表，用时={:.3} 秒）",
                display_path_for_user(&result.path),
                result.sheets,
                (initialize_started.elapsed().as_millis()) as f64 / 1000.0
            )?;
        }
        if log_view == "developer" {
            writeln!(output, "实际写入引擎: rust_xlsxwriter")?;
        }
        refresh_interactive_global(config_path, global, config_error);
        if let Some(error) = config_error.as_deref() {
            if log_view == "developer" {
                writeln!(output, "[全局配置加载失败] {error}")?;
            } else {
                writeln!(
                    output,
                    "[全局配置加载失败] 请检查新建配置表，必要时选择 7-1 修复。"
                )?;
            }
        }
        return Ok(());
    }

    let output_dir = unique_menu_action_dir(&global.output_dir, "config-repair");
    let candidate_path = output_dir.join("config-candidate.xlsx");
    let plan_path = output_dir.join("repair-plan.json");
    let original_sha256 = sha256_file(config_path)?;
    let plan_started = Instant::now();
    facts.add_engine("calamine");
    facts.add_component("pytools_core::config::build_config_repair_plan");
    facts.add_component("excel_reader::read_xlsx_sheet_names_ooxml");
    let Some(plan) = build_config_repair_plan(config_path, &candidate_path)? else {
        facts.metrics = serde_json::json!({
            "operation": "already_complete",
            "execution_mode": "not_needed",
            "planned_operations": 0,
            "plan_file_written": false,
            "candidate_generated": false,
            "committed": false,
        });
        facts.add_stage("configuration_read_and_repair_plan", plan_started.elapsed());
        writeln!(
            output,
            "[无需修复] 当前配置已经完整：{}",
            display_path_for_user(&config_path)
        )?;
        refresh_interactive_global(config_path, global, config_error);
        return Ok(());
    };
    facts.metrics = serde_json::json!({
        "operation": "repair_plan",
        "execution_mode": "plan_only",
        "planned_operations": plan.operations.len(),
        "plan_file_written": global.emit_plan_files,
        "candidate_generated": false,
        "committed": false,
    });
    facts.add_stage("configuration_read_and_repair_plan", plan_started.elapsed());
    fs::create_dir_all(&output_dir)?;
    if global.emit_plan_files {
        fs::write(&plan_path, serde_json::to_vec_pretty(&plan)?)?;
        if log_view == "off" {
            writeln!(
                output,
                "[计划已生成] 共 {} 项操作：{}",
                plan.operations.len(),
                display_path_for_user(&plan_path)
            )?;
        } else {
            writeln!(
                output,
                "[计划已生成] 共 {} 项操作，用时={:.3} 秒：{}",
                plan.operations.len(),
                (plan_started.elapsed().as_millis()) as f64 / 1000.0,
                display_path_for_user(&plan_path)
            )?;
        }
    } else {
        if log_view == "off" {
            writeln!(
                output,
                "[计划已生成] 共 {} 项操作（全局配置 输出附带计划文件=否，计划未落盘）",
                plan.operations.len()
            )?;
        } else {
            writeln!(
                output, "[计划已生成] 共 {} 项操作，用时={:.3} 秒（全局配置 输出附带计划文件=否，计划未落盘）",
                plan.operations.len(), (plan_started.elapsed().as_millis()) as f64 / 1000.0)?;
        }
    }

    let _ = (capabilities, platform);

    verify_planning_snapshot(config_path, Some(&original_sha256))?;
    let execution_started = Instant::now();
    facts.excel_modifier_invoked = true;
    facts.add_component("excel_editor::OpenpyxlXlsxBackend");
    let execution = pytools_core::inplace_edit::execute_candidate(
        &plan,
        pytools_core::inplace_edit::InPlaceEngine::Openpyxl,
    )?;
    facts.add_engine("openpyxl");
    let _candidate_view = view_config(&candidate_path)?;
    facts.add_component("pytools_core::config::view_config");
    let _candidate_global = load_global(&candidate_path)?;
    facts.add_component("pytools_core::config::load_global");
    facts.add_stage("candidate_modify_and_validate", execution_started.elapsed());
    facts.metrics["operation"] = serde_json::json!("candidate_ready");
    facts.metrics["execution_mode"] = serde_json::json!("candidate_ready");
    facts.metrics["candidate_generated"] = serde_json::json!(true);
    if log_view == "off" {
        writeln!(
            output,
            "[候选已验证] {}；源配置仍未替换。",
            display_path_for_user(&candidate_path)
        )?;
    } else {
        writeln!(
            output,
            "[候选已验证] {}；源配置仍未替换；执行用时={:.3} 秒。",
            display_path_for_user(&candidate_path),
            (execution_started.elapsed().as_millis()) as f64 / 1000.0
        )?;
    }
    if log_view == "developer" {
        writeln!(
            output,
            "实际执行信息: {}",
            serde_json::to_string(&execution)?
        )?;
    }
    let confirm = read_menu_choice(
        input,
        output,
        "确认用该候选替换当前配置？输入 y 确认；默认保留原文件> ",
    )?;
    if confirm
        .as_deref()
        .is_some_and(|answer| matches!(answer.to_ascii_lowercase().as_str(), "y" | "yes"))
    {
        let commit_started = Instant::now();
        facts.add_component("pytools_core::config::replace_target_from_candidate");
        let committed =
            replace_target_from_candidate(&candidate_path, config_path, &original_sha256)?;
        facts.metrics["operation"] = serde_json::json!("committed");
        facts.metrics["execution_mode"] = serde_json::json!("committed");
        facts.metrics["committed"] = serde_json::json!(true);
        facts.add_stage("configuration_safe_commit", commit_started.elapsed());
        if log_view == "off" {
            writeln!(
                output,
                "[完成] 配置已安全更新：{}；原文件备份：{}",
                display_path_for_user(&committed.target),
                committed
                    .backup_path
                    .as_deref()
                    .map(|path| display_path_for_user(&path).to_string())
                    .unwrap_or_else(|| "无".to_owned())
            )?;
        } else {
            writeln!(
                output,
                "[完成] 配置已安全更新：{}；原文件备份：{}；提交用时={:.3} 秒",
                display_path_for_user(&committed.target),
                committed
                    .backup_path
                    .as_deref()
                    .map(|path| display_path_for_user(&path).to_string())
                    .unwrap_or_else(|| "无".to_owned()),
                (commit_started.elapsed().as_millis()) as f64 / 1000.0
            )?;
        }
        refresh_interactive_global(config_path, global, config_error);
        if let Some(error) = config_error.as_deref() {
            anyhow::bail!("updated configuration cannot be loaded: {error}");
        }
    } else {
        facts.metrics["operation"] = serde_json::json!("candidate_retained");
        facts.metrics["execution_mode"] = serde_json::json!("candidate_retained");
        writeln!(output, "未替换当前配置；候选文件保留在输出目录。")?;
    }
    Ok(())
}

fn refresh_interactive_global(
    config_path: &Path,
    global: &mut GlobalConfig,
    config_error: &mut Option<String>,
) {
    match load_global(config_path) {
        Ok(mut refreshed) => {
            // 重新加载配置不得丢弃 --log-dir 显式覆盖：配置初始化/重载后
            // 日志目录仍必须留在用户指定位置。
            refreshed.log_dir_override = global.log_dir_override.clone();
            if let Some(log_dir) = &global.log_dir_override {
                refreshed.log_dir = log_dir.clone();
            }
            *global = refreshed;
            *config_error = None;
        }
        Err(error) => *config_error = Some(format!("{error:#}")),
    }
}

fn run_interactive_archive_menu<R, W>(
    input: &mut R,
    output: &mut W,
    config_path: &Path,
    global: &GlobalConfig,
    log_view: &str,
    allow_native_dialog: bool,
) -> Result<()>
where
    R: BufRead,
    W: Write,
{
    loop {
        writeln!(output, "------ 归档整理 ------")?;
        writeln!(output, "  1  按日期归档（预览）")?;
        writeln!(
            output,
            "  2  按类型归档（预览，使用归档类型配置）  【归档类型配置】"
        )?;
        writeln!(output, "  3  执行归档计划")?;
        writeln!(output, "  4  按日期归档（直接复制）")?;
        writeln!(
            output,
            "  5  按类型归档（直接复制，使用归档类型配置）  【归档类型配置】"
        )?;
        writeln!(output, "  0  返回")?;
        writeln!(output, "----------------------")?;
        let Some(choice) = read_menu_choice(input, output, "子菜单选择> ")? else {
            return Ok(());
        };
        let (feature_id, operation, description) = match choice.as_str() {
            "1" => ("PT-8-1", "date-preview", "按日期归档预览"),
            "2" => ("PT-8-2", "type-preview", "按类型归档预览"),
            "3" => ("PT-8-3", "plan", "执行归档计划"),
            "4" => ("PT-8-4", "date-copy", "按日期直接复制归档"),
            "5" => ("PT-8-5", "type-copy", "按类型直接复制归档"),
            "0" => return Ok(()),
            _ => {
                writeln!(output, "仅接受 1、2、3、4、5 或 0")?;
                continue;
            }
        };

        let mut arguments = vec![
            "--output-dir".to_owned(),
            global.output_dir.to_string_lossy().into_owned(),
        ];
        let input_count = if operation == "plan" {
            let Some(plan) = choose_archive_plan_workbook(
                "选择归档计划 Excel 文件",
                input,
                output,
                allow_native_dialog,
            )?
            else {
                writeln!(output, "[已取消]")?;
                return Ok(());
            };
            arguments.extend(["--plan".to_owned(), plan.to_string_lossy().into_owned()]);
            1
        } else {
            let Some(paths) =
                choose_archive_paths(description, input, output, allow_native_dialog)?
            else {
                writeln!(output, "[已取消]")?;
                return Ok(());
            };
            let Some(target_root) = choose_archive_target_root(input, output, allow_native_dialog)?
            else {
                writeln!(output, "[已取消]")?;
                return Ok(());
            };
            arguments.extend([
                "--target-root".to_owned(),
                target_root.to_string_lossy().into_owned(),
                "--config".to_owned(),
                config_path.to_string_lossy().into_owned(),
            ]);
            arguments.extend(paths.iter().map(|path| path.to_string_lossy().into_owned()));
            paths.len()
        };

        run_archive_menu_operation(
            arguments,
            operation,
            feature_id,
            description,
            input_count,
            global,
            log_view,
        )?;
        return Ok(());
    }
}

fn choose_archive_plan_workbook<R, W>(
    title: &str,
    input: &mut R,
    output: &mut W,
    allow_native_dialog: bool,
) -> Result<Option<PathBuf>>
where
    R: BufRead,
    W: Write,
{
    let Some(mut selected) = choose_input_paths(
        title,
        false,
        NativeFileFilter::XlsxOnly,
        input,
        output,
        allow_native_dialog,
    )?
    else {
        return Ok(None);
    };
    if selected.len() != 1 {
        writeln!(output, "[提示] 归档计划执行需要且只需要一个 .xlsx 工作簿。")?;
        return Ok(None);
    }
    let path = selected.remove(0);
    if !path
        .extension()
        .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
    {
        writeln!(output, "[跳过] 归档计划当前只接受 .xlsx 工作簿。")?;
        return Ok(None);
    }
    Ok(Some(path))
}

fn choose_archive_paths<R, W>(
    title: &str,
    input: &mut R,
    output: &mut W,
    allow_native_dialog: bool,
) -> Result<Option<Vec<PathBuf>>>
where
    R: BufRead,
    W: Write,
{
    writeln!(output, "→ {title}")?;
    writeln!(output, "  1  选择文件（可多选）")?;
    writeln!(output, "  2  选择文件夹（仅当前层）")?;
    writeln!(output, "  3  选择文件夹（递归子目录）")?;
    writeln!(output, "  0  取消")?;
    let Some(mode) = read_menu_choice(input, output, "选择方式> ")? else {
        return Ok(None);
    };
    match mode.as_str() {
        "0" => Ok(None),
        "1" => choose_input_paths(
            title,
            true,
            NativeFileFilter::AllFiles,
            input,
            output,
            allow_native_dialog,
        ),
        "2" | "3" => {
            let recursive = mode == "3";
            let native_folder = allow_native_dialog
                .then(|| native_folder_dialog(title))
                .flatten();
            let mut folders = native_folder.into_iter().collect::<Vec<_>>();
            if folders.is_empty() {
                writeln!(
                    output,
                    "→ 可手动输入文件夹路径（多个路径用分号分隔），或输入 0 取消："
                )?;
                let Some(raw) = read_menu_choice(input, output, "文件夹路径> ")? else {
                    return Ok(None);
                };
                let Some(paths) = parse_manual_path_list(&raw) else {
                    return Ok(None);
                };
                folders.extend(paths);
            } else {
                print_selected_paths(output, &folders)?;
            }
            let files = collect_menu_source_files(&folders, recursive, NativeFileFilter::AllFiles)?;
            writeln!(
                output,
                "→ 已选择文件夹 {} 个，{}共识别文件 {} 个",
                folders.len(),
                if recursive { "递归" } else { "当前层" },
                files.len()
            )?;
            print_selected_paths(output, &files)?;
            Ok(Some(files))
        }
        _ => {
            writeln!(output, "仅接受 1、2、3 或 0")?;
            Ok(None)
        }
    }
}

fn choose_archive_target_root<R, W>(
    input: &mut R,
    output: &mut W,
    allow_native_dialog: bool,
) -> Result<Option<PathBuf>>
where
    R: BufRead,
    W: Write,
{
    writeln!(output, "→ 请选择归档目标根目录")?;
    if allow_native_dialog {
        if let Some(path) = native_folder_dialog("选择归档目标根目录") {
            print_selected_paths(output, std::slice::from_ref(&path))?;
            return Ok(Some(path));
        }
        writeln!(output, "[提示] 文件夹弹窗不可用或已取消。")?;
    }
    writeln!(output, "→ 可手动输入目标文件夹路径，或输入 0 取消：")?;
    let Some(raw) = read_menu_choice(input, output, "目标根目录> ")? else {
        return Ok(None);
    };
    let Some(mut paths) = parse_manual_path_list(&raw) else {
        return Ok(None);
    };
    anyhow::ensure!(
        paths.len() == 1,
        "archive target root must be a single directory"
    );
    let path = paths.remove(0);
    anyhow::ensure!(
        path.is_dir(),
        "archive target root does not exist or is not a directory: {}",
        display_path_for_user(&path)
    );
    print_selected_paths(output, std::slice::from_ref(&path))?;
    Ok(Some(path))
}

fn config_cli_arguments(config_path: &Path) -> Vec<String> {
    vec![
        "--config".to_owned(),
        config_path.to_string_lossy().into_owned(),
    ]
}

fn run_interactive_timeline_menu<R, W>(
    input: &mut R,
    output: &mut W,
    config_path: &Path,
    global: &GlobalConfig,
    log_view: &str,
    allow_native_dialog: bool,
) -> Result<()>
where
    R: BufRead,
    W: Write,
{
    loop {
        writeln!(output, "------ 时序工具 ------")?;
        writeln!(output, "  1  表头路径比对")?;
        writeln!(output, "  2  时序提取（PT-1-2，当前迁移范围排除）")?;
        writeln!(output, "  3  宽表汇总（行头加后缀）")?;
        writeln!(output, "  4  宽表汇总（行头不加后缀）")?;
        writeln!(output, "  5  预校验")?;
        writeln!(output, "  0  返回")?;
        writeln!(output, "----------------------")?;
        let Some(choice) = read_menu_choice(input, output, "子菜单选择> ")? else {
            return Ok(());
        };
        match choice.as_str() {
            "1" => {
                let Some(template) = choose_workbooks(
                    "选择表头路径比对模板",
                    false,
                    input,
                    output,
                    allow_native_dialog,
                )?
                .and_then(|paths| paths.into_iter().next()) else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let Some(sources) = choose_workbooks(
                    "选择表头路径比对源文件（可多选）",
                    true,
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let mut arguments = config_cli_arguments(config_path);
                arguments.extend([
                    "--log-dir".to_owned(),
                    global.log_dir.to_string_lossy().into_owned(),
                    "--log-view".to_owned(),
                    log_view.to_owned(),
                ]);
                arguments.push("--template".to_owned());
                arguments.push(template.to_string_lossy().into_owned());
                arguments.extend(
                    sources
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_timeline_compare_cli(arguments, true)?;
                return Ok(());
            }
            "2" => {
                writeln!(
                    output,
                    "[范围排除] PT-1-2 时序提取不在已确认迁移范围内；不会启动该候选功能。"
                )?;
                return Ok(());
            }
            "3" | "4" => {
                let Some(sources) = choose_workbooks(
                    if choice == "3" {
                        "选择行头加后缀宽表汇总源文件（可多选）"
                    } else {
                        "选择行头不加后缀宽表汇总源文件（可多选）"
                    },
                    true,
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let mut arguments = config_cli_arguments(config_path);
                arguments.extend(summary_cli_arguments(global, log_view));
                arguments.extend(
                    sources
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_interactive_wide_summary_menu(output, arguments, choice == "3")?;
                return Ok(());
            }
            "5" => {
                run_interactive_timeline_precheck(output, config_path, global, log_view)?;
                return Ok(());
            }
            "0" => return Ok(()),
            _ => writeln!(output, "仅接受 1、2、3、4、5 或 0")?,
        }
    }
}

fn run_interactive_wide_summary_menu<W: Write>(
    output: &mut W,
    arguments: Vec<String>,
    row_suffix: bool,
) -> Result<()>
where
    W: Write,
{
    // 菜单输出为人类可读行（带时间戳）；developer 视图保留机器可读 JSON，与 CLI 行为一致。
    let developer_view = arguments
        .iter()
        .position(|argument| argument == "--log-view")
        .and_then(|index| arguments.get(index + 1))
        .map(|value| value == "developer")
        .unwrap_or(false);
    let report = run_wide_summary_cli_report(arguments, row_suffix)?;
    if developer_view {
        writeln!(output, "{}", serde_json::to_string(&report)?)?;
        return Ok(());
    }
    let (status, level) = match report["overall_status"]
        .as_str()
        .unwrap_or_default()
        .to_ascii_uppercase()
        .as_str()
    {
        "SUCCESS" => ("完成", pytools_core::run_logging::Level::Info),
        "PARTIAL_SUCCESS" => ("部分完成", pytools_core::run_logging::Level::Warn),
        "NO_OP" => ("未生成汇总", pytools_core::run_logging::Level::Info),
        "FAILED" => ("失败", pytools_core::run_logging::Level::Error),
        "CANCELLED" => ("已取消", pytools_core::run_logging::Level::Warn),
        _ => ("状态记录缺失", pytools_core::run_logging::Level::Warn),
    };
    let count = |name: &str| {
        report["input_stats"]
            .get(name)
            .and_then(serde_json::Value::as_u64)
            .map(|value| value.to_string())
            .unwrap_or_else(|| "未采集".to_owned())
    };
    writeln!(
        output,
        "{}",
        pytools_core::run_logging::terminal_line(level, &format!(
            "宽表汇总：{status}，选择 {} 个，尝试 {} 个，成功 {} 个，失败 {} 个，跳过 {} 个，未完成 {} 个",
            count("selected"), count("attempted"), count("succeeded"), count("failed"), count("skipped"), count("pending")
        ))
    )?;
    if let Some(output_path) = report["output"].as_str() {
        writeln!(
            output,
            "{}",
            pytools_core::run_logging::terminal_line(
                pytools_core::run_logging::Level::Info,
                &format!("汇总文件：{output_path}")
            )
        )?;
    }
    match report["target_write_status"].as_str() {
        Some("not_configured") => {
            writeln!(
                output,
                "{}",
                pytools_core::run_logging::terminal_line(
                    pytools_core::run_logging::Level::Warn,
                    "目标未写入：请检查规则的目标写入开关、目标路径和工作表"
                )
            )?;
        }
        Some(other) if other != "not_run" && other != "not_applicable" => {
            if let Some(items) = report["target_writes"].as_array() {
                let mut seen = std::collections::BTreeSet::new();
                for item in items {
                    if let Some(target) = item["target"].as_str() {
                        if seen.insert(target.to_owned()) {
                            let state = match item["status"].as_str() {
                                Some("succeeded") => "成功",
                                Some("failed") => "失败",
                                Some("skipped") => "跳过",
                                _ => "状态不可用",
                            };
                            let result_level = match item["status"].as_str() {
                                Some("failed") => pytools_core::run_logging::Level::Error,
                                Some("skipped") => pytools_core::run_logging::Level::Warn,
                                _ => pytools_core::run_logging::Level::Info,
                            };
                            writeln!(
                                output,
                                "{}",
                                pytools_core::run_logging::terminal_line(
                                    result_level,
                                    &format!("目标写入：{target}（{state}）")
                                )
                            )?;
                            if item["status"] == "failed" {
                                writeln!(output, "{}", pytools_core::run_logging::terminal_line(pytools_core::run_logging::Level::Error, "写入失败；请检查目标工作簿、工作表、文件占用状态和写入权限"))?;
                            }
                        }
                    }
                }
            }
        }
        _ => {}
    }
    if let Some(failure) = report["failure"].as_object() {
        let stage = failure["stage"].as_str().unwrap_or("wide_summary");
        let code = summary_event_error_code(failure["error_type"].as_str(), stage);
        writeln!(
            output,
            "{}",
            pytools_core::run_logging::terminal_line(
                pytools_core::run_logging::Level::Error,
                &format!("失败阶段：{stage}，错误代码：{code}")
            )
        )?;
    }
    if let Some(elapsed) = report["total_elapsed_ms"].as_u64() {
        writeln!(
            output,
            "{}",
            pytools_core::run_logging::terminal_line(
                pytools_core::run_logging::Level::Info,
                &format!(
                    "总耗时 {elapsed_display_seconds:.3} 秒",
                    elapsed_display_seconds = elapsed as f64 / 1000.0
                )
            )
        )?;
    }
    Ok(())
}
fn run_interactive_dedup_menu<R, W>(
    input: &mut R,
    output: &mut W,
    config_path: &Path,
    global: &GlobalConfig,
    log_view: &str,
    allow_native_dialog: bool,
    capabilities: Option<&Path>,
    platform: Option<&str>,
) -> Result<()>
where
    R: BufRead,
    W: Write,
{
    loop {
        writeln!(output, "------ 数据去重与追加 ------")?;
        writeln!(
            output,
            "  1  按配置检查重复  【去重追加数据配置：源数据工作簿/源数据工作表/标识列序号】"
        )?;
        writeln!(
            output,
            "  2  按配置删除重复  【去重追加数据配置：源数据工作簿/源数据工作表/标识列序号】"
        )?;
        writeln!(output, "  3  按配置去重追加到目标工作簿  【去重追加数据配置：源数据/标识列/目标工作簿/目标工作表】")?;
        writeln!(output, "  4  去重/追加配置预校验  【去重追加数据配置】")?;
        writeln!(output, "  0  返回")?;
        writeln!(output, "----------------------")?;
        let Some(menu_choice) = read_menu_choice(input, output, "子菜单选择> ")? else {
            return Ok(());
        };
        let feature_id = match menu_choice.as_str() {
            "1" => "dedup.config_check",
            "2" => "dedup.config_delete",
            "3" => "dedup.append_to_history",
            "4" => "dedup.config_precheck",
            "0" => return Ok(()),
            _ => {
                writeln!(output, "仅接受 1、2、3、4 或 0")?;
                continue;
            }
        };
        let choice = match feature_id {
            "dedup.config_check" => "3",
            "dedup.config_delete" => "4",
            "dedup.append_to_history" => "5",
            "dedup.config_precheck" => "6",
            _ => unreachable!("resolved stable PyTools feature ID"),
        };
        match choice {
            "1" | "2" => {
                let Some(input_path) = choose_workbooks(
                    if choice == "1" {
                        "选择按批注检查并标记重复的工作簿"
                    } else {
                        "选择按批注删除重复的工作簿"
                    },
                    false,
                    input,
                    output,
                    allow_native_dialog,
                )?
                .and_then(|paths| paths.into_iter().next()) else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let candidate = unique_workbook_candidate(&input_path, &global.output_dir, "dedup");
                let mut arguments = config_cli_arguments(config_path);
                arguments.extend([
                    "--input".to_owned(),
                    input_path.to_string_lossy().into_owned(),
                    "--output".to_owned(),
                    candidate.to_string_lossy().into_owned(),
                    "--action".to_owned(),
                    if choice == "1" {
                        "check".to_owned()
                    } else {
                        "delete".to_owned()
                    },
                ]);
                let action_name = if choice == "1" {
                    "comment-check"
                } else {
                    "comment-delete"
                };
                let action_output_dir = unique_menu_output_dir(&global.output_dir, action_name);
                let plan_path = action_output_dir.join("plan.json");
                arguments.push("--plan-output".to_owned());
                arguments.push(plan_path.to_string_lossy().into_owned());
                let (feature_id, action, label) = if choice == "1" {
                    ("PT-2-1", "comment_check", "按批注检查重复并生成计划")
                } else {
                    ("PT-2-2", "comment_delete", "按批注删除重复并生成计划")
                };
                run_dedup_menu_action(
                    output,
                    global,
                    log_view,
                    feature_id,
                    action,
                    label,
                    &[
                        "Rust duplicate planning",
                        "Calamine via excel-reader",
                        "OOXML comments and workbook profile metadata",
                    ],
                    arguments,
                    dedup_comment_plan_result,
                )?;
                return Ok(());
            }
            "3" | "4" => {
                if choice == "3" {
                    // 2-1 改为在源工作簿上就地标红标记重复行（对齐旧版 Python
                    // `dedup_311.py`），不再生成标记副本工作簿。
                    let label = "按配置检查重复（在源工作簿上标红标记重复行）";
                    writeln!(
                        output,
                        "{} [INFO] 开始处理：{label}",
                        pytools_core::engine_diag::now_prefix()
                    )?;
                    let mut log = ConfiguredRunLog::new(global, "2-1", label, log_view);
                    let mut report_tasks = Vec::new();
                    let mut progress_write_error = None;
                    let result = match load_dedup_tasks(config_path) {
                        Ok(tasks) => {
                            let enabled_count = tasks.iter().filter(|task| task.enabled).count();
                            writeln!(
                                output,
                                "{} [INFO] 已启用配置：{enabled_count} 项",
                                pytools_core::engine_diag::now_prefix()
                            )?;
                            let result =
                                pytools_core::dedup_ops::mark_duplicates_in_place_with_progress(
                                    &tasks,
                                    |task| {
                                        if progress_write_error.is_none() {
                                            progress_write_error =
                                                write_dedup_task_start(output, task, false).err();
                                        }
                                    },
                                );
                            report_tasks = tasks;
                            result
                        }
                        Err(error) => Err(error),
                    };
                    let summary = match &result {
                        Ok(run) => {
                            log.set_business_metric(
                                "标记重复行数",
                                serde_json::json!(run.marked_rows),
                            );
                            log.set_business_metric(
                                "保存工作簿数",
                                serde_json::json!(run.saved_books.len()),
                            );
                            if run.marked_rows > 0 {
                                log.record_business_issue(
                                    pytools_core::run_logging::BusinessIssue::new(
                                        "重复数据",
                                        "配置检查发现重复行并已标记",
                                    )
                                    .count(run.marked_rows),
                                );
                            }
                            log.finish(
                                false,
                                Some((run.task_ok, run.task_skipped, run.task_failed)),
                                &run.events,
                                &run.saved_books,
                            )
                        }
                        Err(error) => {
                            log.set_error_summary(pytools_core::run_logging::ErrorSummary {
                                stage: "duplicate_mark".to_owned(),
                                code: "DEDUP_OPERATION_FAILED".to_owned(),
                                message: format!("{label}失败：{error:#}"),
                                file_name: None,
                                sheet_name: None,
                            });
                            log.finish(true, None, &[], &[])
                        }
                    };
                    if let Some(error) = progress_write_error {
                        return Err(error.into());
                    }
                    match result {
                        Ok(run) => {
                            write_dedup_task_results(
                                output,
                                &report_tasks,
                                &run.events,
                                &run.saved_books,
                                "标记重复行",
                            )?;
                            for event in run.events.iter().filter(|event| {
                                matches!(
                                    event.outcome,
                                    pytools_core::engine_diag::DiagOutcome::Skipped
                                        | pytools_core::engine_diag::DiagOutcome::Failed
                                )
                            }) {
                                writeln!(output, "{}", event.user_line())?;
                            }
                            write_unified_terminal_summary(output, &summary, &run.saved_books)?;
                        }
                        Err(_error) => {
                            write_unified_terminal_summary(output, &summary, &[])?;
                        }
                    }
                    return Ok(());
                }
                if choice == "4" {
                    // 2-2 改为在源工作簿上就地删除重复行（保留首次出现），
                    // 不再生成删除结果副本工作簿。
                    let label = "按配置删除重复（在源工作簿上就地删除重复行）";
                    writeln!(
                        output,
                        "{} [INFO] 开始处理：{label}",
                        pytools_core::engine_diag::now_prefix()
                    )?;
                    let mut log = ConfiguredRunLog::new(global, "2-2", label, log_view);
                    let mut report_tasks = Vec::new();
                    let mut progress_write_error = None;
                    let result = match load_dedup_tasks(config_path) {
                        Ok(tasks) => {
                            let enabled_count = tasks.iter().filter(|task| task.enabled).count();
                            writeln!(
                                output,
                                "{} [INFO] 已启用配置：{enabled_count} 项",
                                pytools_core::engine_diag::now_prefix()
                            )?;
                            let result =
                                pytools_core::dedup_ops::delete_duplicates_in_place_with_progress(
                                    &tasks,
                                    |task| {
                                        if progress_write_error.is_none() {
                                            progress_write_error =
                                                write_dedup_task_start(output, task, false).err();
                                        }
                                    },
                                );
                            report_tasks = tasks;
                            result
                        }
                        Err(error) => Err(error),
                    };
                    let summary = match &result {
                        Ok(run) => {
                            log.set_business_metric(
                                "删除重复行数",
                                serde_json::json!(run.deleted_rows),
                            );
                            log.set_business_metric(
                                "保存工作簿数",
                                serde_json::json!(run.saved_books.len()),
                            );
                            log.finish(
                                false,
                                Some((run.task_ok, run.task_skipped, run.task_failed)),
                                &run.events,
                                &run.saved_books,
                            )
                        }
                        Err(error) => {
                            log.set_error_summary(pytools_core::run_logging::ErrorSummary {
                                stage: "duplicate_delete".to_owned(),
                                code: "DEDUP_OPERATION_FAILED".to_owned(),
                                message: format!("{label}失败：{error:#}"),
                                file_name: None,
                                sheet_name: None,
                            });
                            log.finish(true, None, &[], &[])
                        }
                    };
                    if let Some(error) = progress_write_error {
                        return Err(error.into());
                    }
                    match result {
                        Ok(run) => {
                            write_dedup_task_results(
                                output,
                                &report_tasks,
                                &run.events,
                                &run.saved_books,
                                "删除重复行",
                            )?;
                            for event in run.events.iter().filter(|event| {
                                matches!(
                                    event.outcome,
                                    pytools_core::engine_diag::DiagOutcome::Skipped
                                        | pytools_core::engine_diag::DiagOutcome::Failed
                                )
                            }) {
                                writeln!(output, "{}", event.user_line())?;
                            }
                            write_unified_terminal_summary(output, &summary, &run.saved_books)?;
                        }
                        Err(_error) => {
                            write_unified_terminal_summary(output, &summary, &[])?;
                        }
                    }
                    return Ok(());
                }
            }
            "5" => {
                run_interactive_dedup_append_menu(
                    input,
                    output,
                    config_path,
                    global,
                    log_view,
                    capabilities,
                    platform,
                )?;
                return Ok(());
            }
            "6" => {
                run_dedup_menu_action(
                    output,
                    global,
                    log_view,
                    "PT-2-6",
                    "config_precheck",
                    "按配置预校验去重任务",
                    &["Rust configuration validation", "Calamine via excel-reader"],
                    config_cli_arguments(config_path),
                    dedup_precheck_result,
                )?;
                return Ok(());
            }
            "0" => return Ok(()),
            _ => unreachable!("stable dedup feature maps to a supported operation"),
        }
    }
}

fn run_interactive_dedup_append_menu<R, W>(
    _input: &mut R,
    output: &mut W,
    config_path: &Path,
    global: &GlobalConfig,
    log_view: &str,
    capabilities: Option<&Path>,
    platform: Option<&str>,
) -> Result<()>
where
    R: BufRead,
    W: Write,
{
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let mut diagnostic = ConfiguredRunLog::new(global, "2-3", "按配置去重追加", log_view);
    let started_at = Local::now();
    let started = Instant::now();
    match log_view {
        "user" => writeln!(
            output,
            "{}",
            pytools_core::run_logging::terminal_line(
                pytools_core::run_logging::Level::Info,
                "开始处理：读取启用配置并去重追加到配置目标工作簿"
            )
        )?,
        "developer" => writeln!(output, "PT-2 开始：PT-2-5 config_append")?,
        _ => {}
    }

    let planning_started = Instant::now();
    let action_output_dir = unique_menu_output_dir(&global.output_dir, "config-append");
    let mut append_tasks = Vec::new();
    let planning_result = match load_dedup_tasks(config_path) {
        Ok(tasks) => {
            let enabled_count = tasks.iter().filter(|task| task.enabled).count();
            if log_view == "user" {
                writeln!(
                    output,
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        pytools_core::run_logging::Level::Info,
                        &format!("已启用配置：{enabled_count} 项")
                    )
                )?;
                for task in tasks.iter().filter(|task| task.enabled) {
                    write_dedup_task_start(output, task, true)?;
                }
                output.flush()?;
            }
            append_tasks = tasks;
            dedup_append_plan_bundle_with_tasks(
                &append_tasks,
                &action_output_dir,
                global.emit_plan_files,
            )
        }
        Err(error) => Err(error),
    };
    let (mut report, plans) = match planning_result {
        Ok(result) => result,
        Err(error) => {
            diagnostic.set_error_summary(pytools_core::run_logging::ErrorSummary {
                stage: "workbook_read_and_plan".to_owned(),
                code: "DEDUP_APPEND_PLAN_FAILED".to_owned(),
                message: format!("读取启用配置或生成追加计划失败：{error:#}"),
                file_name: None,
                sheet_name: None,
            });
            let summary = diagnostic.finish(true, None, &[], &[]);
            let elapsed_ms = started.elapsed().as_millis();
            let stages = vec![serde_json::json!({
                "name": "workbook_read_and_plan",
                "status": "failed",
                "elapsed_ms": planning_started.elapsed().as_millis()
            })];
            if log_view == "developer" {
                match write_dedup_append_menu_run_history(
                    global,
                    &started_at,
                    elapsed_ms,
                    "failed",
                    &serde_json::json!({}),
                    &stages,
                    &["Rust append planning".to_owned()],
                    false,
                    false,
                    false,
                    Some("dedup_operation_failed"),
                ) {
                    Ok(path) => writeln!(output, "PT-2 失败：PT-2-5，耗时 {elapsed_ms_display_seconds:.3} 秒；运行记录：{}；错误：{error:#}", display_path_for_user(&path), elapsed_ms_display_seconds = elapsed_ms as f64 / 1000.0)?,
                    Err(log_error) => writeln!(output, "PT-2 失败：PT-2-5，耗时 {elapsed_ms_display_seconds:.3} 秒；运行记录写入失败：{log_error:#}；错误：{error:#}", elapsed_ms_display_seconds = elapsed_ms as f64 / 1000.0)?,
                }
            } else {
                write_unified_terminal_summary(output, &summary, &[])?;
            }
            return Ok(());
        }
    };
    let planning_elapsed_ms = planning_started.elapsed().as_millis();

    let mut executable_indices = Vec::new();
    let mut diagnostic_events = Vec::new();
    for (index, plan) in plans.iter().enumerate() {
        if plan.appended_rows == 0 {
            report[index]["execution"] = serde_json::Value::String("no_changes".to_owned());
            report[index]["status"] = serde_json::Value::String("no_changes".to_owned());
        } else {
            report[index]["status"] = serde_json::Value::String("pending".to_owned());
            executable_indices.push(index);
        }
        if log_view == "user" {
            for issue in &plan.issues {
                writeln!(
                    output,
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        pytools_core::run_logging::Level::Warn,
                        &format!("配置跳过：{}；{issue}", display_path_for_user(&plan.target))
                    )
                )?;
            }
        }
    }

    let mut modifier_invoked = false;
    let mut writer_invoked = false;
    for index in executable_indices.iter().copied() {
        let plan = &plans[index];
        if log_view == "user" {
            // 目标路径集中展示在「目标汇总结果」；写入过程只报状态与数量。
            writeln!(
                output,
                "{}",
                pytools_core::run_logging::terminal_line(
                    pytools_core::run_logging::Level::Info,
                    &format!(
                        "正在写入目标工作簿（追加 {} 行，过滤重复 {} 行）",
                        plan.appended_rows, plan.duplicate_rows_filtered
                    )
                )
            )?;
        }
        let mut executor_invoked = false;
        let direct_existing = plan.target_existed && capabilities.is_none();
        let execution = (|| -> Result<serde_json::Value> {
            if direct_existing {
                executor_invoked = true;
                return run_dedup_append_direct_existing(plan);
            }
            let plan_file = write_temporary_dedup_append_plan(plan)?;
            let mut arguments = vec![
                "--plan".to_owned(),
                plan_file.path().to_string_lossy().into_owned(),
                "--commit-target".to_owned(),
            ];
            if plan.target_existed {
                arguments.push("--capabilities".to_owned());
                arguments.push(
                    capabilities
                        .context("capabilities missing for existing target")?
                        .to_string_lossy()
                        .into_owned(),
                );
                arguments.push("--platform".to_owned());
                arguments.push(
                    platform
                        .context("platform missing for existing target")?
                        .to_owned(),
                );
            }
            executor_invoked = true;
            run_dedup_append_execute_report(arguments)
        })();
        if executor_invoked {
            if plan.target_existed && !direct_existing {
                modifier_invoked = true;
            } else if !plan.target_existed {
                writer_invoked = true;
            }
        }

        match execution {
            Ok(execution) => {
                let engine = execution["engine"]
                    .as_str()
                    .or_else(|| execution["candidate_engine"].as_str())
                    .unwrap_or("unknown")
                    .to_owned();
                diagnostic_events.push(
                    pytools_core::engine_diag::DiagEvent::new(
                        &engine,
                        pytools_core::engine_diag::DiagStage::SaveOutput,
                        pytools_core::engine_diag::DiagOutcome::Ok,
                    )
                    .with_file(plan.target.to_string_lossy().into_owned())
                    .with_message("目标工作簿追加完成"),
                );
                report[index]["execution"] = serde_json::Value::String("committed".to_owned());
                report[index]["status"] = serde_json::Value::String("succeeded".to_owned());
                report[index]["execution_result"] = execution;
                if log_view == "user" {
                    // 路径由「目标汇总结果」统一展示一次。
                    writeln!(
                        output,
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            pytools_core::run_logging::Level::Info,
                            "已写入目标工作簿",
                        )
                    )?;
                }
            }
            Err(error) => {
                let error_message = format!("{error:#}");
                let stage = if error_message.contains("backend capability")
                    || error_message.contains("capabilities missing")
                    || error_message.contains("platform missing")
                {
                    pytools_core::engine_diag::DiagStage::ReadConfig
                } else if error_message.contains("stage=excel_modifier")
                    || error_message.contains("stage=direct_append_candidate")
                {
                    pytools_core::engine_diag::DiagStage::ModifyWorkbook
                } else {
                    pytools_core::engine_diag::DiagStage::SaveOutput
                };
                diagnostic_events.push(
                    pytools_core::engine_diag::DiagEvent::new(
                        if plan.target_existed {
                            "openpyxl"
                        } else {
                            "rust_xlsxwriter"
                        },
                        stage,
                        pytools_core::engine_diag::DiagOutcome::Failed,
                    )
                    .with_file(plan.target.to_string_lossy().into_owned())
                    .with_message(error_message.clone()),
                );
                report[index]["execution"] = serde_json::Value::String("failed".to_owned());
                report[index]["status"] = serde_json::Value::String("failed".to_owned());
                report[index]["error"] = serde_json::Value::String(error_message);
                break;
            }
        }
    }
    for index in executable_indices {
        if report[index]["status"] == "pending" {
            report[index]["execution"] = serde_json::Value::String("not_processed".to_owned());
            report[index]["status"] = serde_json::Value::String("not_processed".to_owned());
        }
    }

    if log_view == "user" {
        for (index, plan) in plans.iter().enumerate() {
            let config_rows = append_tasks
                .iter()
                .filter(|task| task.enabled && task.target_wb.as_ref() == Some(&plan.target))
                .map(|task| format!("第 {} 行「{}」", task.row_no, dedup_task_label(task)))
                .collect::<Vec<_>>()
                .join("、");
            let status = report[index]["status"].as_str().unwrap_or("unknown");
            let committed = status == "succeeded" && plan.appended_rows > 0;
            let config_rows_label = if config_rows.is_empty() {
                "未匹配到配置行"
            } else {
                config_rows.as_str()
            };
            writeln!(
                output,
                "{}",
                pytools_core::run_logging::terminal_line(
                    if status == "failed" {
                        pytools_core::run_logging::Level::Error
                    } else {
                        pytools_core::run_logging::Level::Info
                    },
                    // 追加/过滤数量由统一摘要集中展示；这里只定位目标与保存状态。
                    &format!(
                        "目标工作簿：{}；关联配置 {}{}",
                        display_path_for_user(&plan.target),
                        config_rows_label,
                        if committed {
                            "；目标工作簿已保存"
                        } else {
                            "；目标工作簿未修改"
                        }
                    )
                )
            )?;
        }
        for task in append_tasks.iter().filter(|task| {
            task.enabled
                && task
                    .target_wb
                    .as_ref()
                    .is_none_or(|target| !plans.iter().any(|plan| &plan.target == target))
        }) {
            writeln!(
                output,
                "{}",
                pytools_core::run_logging::terminal_line(
                    pytools_core::run_logging::Level::Warn,
                    &format!(
                        "配置结果：第 {} 行「{}」；未生成目标计划，未执行去重",
                        task.row_no,
                        dedup_task_label(task)
                    )
                )
            )?;
        }
    }

    let elapsed_ms = started.elapsed().as_millis();
    let mut metrics = dedup_menu_metrics(&report);
    let failed_targets = report
        .as_array()
        .into_iter()
        .flatten()
        .filter(|item| item["status"] == "failed")
        .count();
    let blocked_targets = report
        .as_array()
        .into_iter()
        .flatten()
        .filter(|item| item["status"] == "planned_only")
        .count();
    let unprocessed_targets = report
        .as_array()
        .into_iter()
        .flatten()
        .filter(|item| item["status"] == "not_processed")
        .count();
    let failed = failed_targets > 0;
    let committed_targets = report
        .as_array()
        .into_iter()
        .flatten()
        .filter(|item| item["status"] == "succeeded")
        .count();
    let committed_rows: usize = plans
        .iter()
        .enumerate()
        .filter(|(index, _)| report[*index]["status"] == "succeeded")
        .map(|(_, plan)| plan.appended_rows)
        .sum();
    let filtered_rows: usize = plans
        .iter()
        .enumerate()
        .filter(|(index, _)| {
            report[*index]["status"] == "succeeded" || report[*index]["status"] == "no_changes"
        })
        .map(|(_, plan)| plan.duplicate_rows_filtered)
        .sum();
    let skipped_tasks: usize = plans.iter().map(|plan| plan.task_skipped).sum();
    metrics["committed_appended_rows"] = serde_json::json!(committed_rows);
    metrics["processed_duplicate_rows_filtered"] = serde_json::json!(filtered_rows);
    metrics["unprocessed_targets"] = serde_json::json!(unprocessed_targets);
    let mut actual_engines = vec![
        "Rust append planning".to_owned(),
        "Calamine via excel-reader".to_owned(),
        "OOXML workbook profile metadata".to_owned(),
    ];
    let mut stages = vec![serde_json::json!({
        "name": "workbook_read_and_plan",
        "status": "succeeded",
        "elapsed_ms": planning_elapsed_ms
    })];
    for (index, item) in report.as_array().into_iter().flatten().enumerate() {
        if let Some(execution) = item.get("execution_result") {
            if let Some(engine) = execution["engine"]
                .as_str()
                .or_else(|| execution["candidate_engine"].as_str())
            {
                if !actual_engines.iter().any(|known| known == engine) {
                    actual_engines.push(engine.to_owned());
                }
            }
            if let Some(execution_stages) = execution["execution_summary"]["stages"].as_array() {
                for stage in execution_stages {
                    stages.push(serde_json::json!({
                        "target_index": index,
                        "name": stage["name"],
                        "status": stage["status"],
                        "engine": stage.get("engine"),
                        "elapsed_ms": stage.get("elapsed_ms")
                    }));
                }
            }
        }
    }
    let run_status = if committed_targets > 0
        && (failed || blocked_targets > 0 || unprocessed_targets > 0 || skipped_tasks > 0)
    {
        "partial"
    } else if failed {
        "failed"
    } else if blocked_targets > 0 {
        "blocked"
    } else {
        "succeeded"
    };
    let unified_status = match run_status {
        "failed" | "blocked" => pytools_core::run_logging::TaskStatus::Failed,
        "partial" => pytools_core::run_logging::TaskStatus::PartialSuccess,
        _ if committed_targets == 0 => pytools_core::run_logging::TaskStatus::NoOp,
        _ => pytools_core::run_logging::TaskStatus::Success,
    };
    let committed = report
        .as_array()
        .into_iter()
        .flatten()
        .enumerate()
        .filter(|(_, item)| item["status"] == "succeeded")
        .filter_map(|(index, _)| plans.get(index).map(|plan| plan.target.clone()))
        .collect::<Vec<_>>();
    for (name, value) in [
        ("写入目标数", serde_json::json!(committed_targets)),
        ("追加行数", serde_json::json!(committed_rows)),
        ("过滤重复行数", serde_json::json!(filtered_rows)),
        ("跳过配置数", serde_json::json!(skipped_tasks)),
        ("受阻目标数", serde_json::json!(blocked_targets)),
        ("失败目标数", serde_json::json!(failed_targets)),
        ("未处理目标数", serde_json::json!(unprocessed_targets)),
    ] {
        diagnostic.set_business_metric(name, value);
    }
    if blocked_targets > 0 {
        diagnostic.set_error_summary(pytools_core::run_logging::ErrorSummary {
            stage: "target_commit".to_owned(),
            code: "EXCEL_MODIFIER_CAPABILITY_REQUIRED".to_owned(),
            message: "目标工作簿需要可用的 Excel 编辑能力；本次未修改受阻目标".to_owned(),
            file_name: report
                .as_array()
                .into_iter()
                .flatten()
                .find(|item| item["status"] == "planned_only")
                .and_then(|item| item["target"].as_str())
                .map(|path| pytools_core::run_logging::portable_file_name(Path::new(path))),
            sheet_name: None,
        });
    } else if failed_targets > 0 {
        diagnostic.set_error_summary(pytools_core::run_logging::ErrorSummary {
            stage: "target_commit".to_owned(),
            code: "DEDUP_APPEND_TARGET_FAILED".to_owned(),
            message: report
                .as_array()
                .into_iter()
                .flatten()
                .find(|item| item["status"] == "failed")
                .and_then(|item| item["error"].as_str())
                .unwrap_or("目标写入失败")
                .to_owned(),
            file_name: report
                .as_array()
                .into_iter()
                .flatten()
                .find(|item| item["status"] == "failed")
                .and_then(|item| item["target"].as_str())
                .map(|path| pytools_core::run_logging::portable_file_name(Path::new(path))),
            sheet_name: None,
        });
    }
    let summary = diagnostic.finish_as(
        failed || blocked_targets > 0,
        None,
        &diagnostic_events,
        &committed,
        Some(unified_status),
    );

    match log_view {
        "user" => {
            if plans.is_empty() {
                writeln!(
                    output,
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        pytools_core::run_logging::Level::Warn,
                        if append_tasks.iter().any(|task| task.enabled) {
                            "启用配置未生成可处理的目标；请检查源工作簿和目标工作簿路径，目标未修改。"
                        } else {
                            "未找到启用配置，目标未修改。"
                        }
                    )
                )?;
            }
            write_unified_terminal_summary(output, &summary, &committed)?;
            for item in report.as_array().into_iter().flatten() {
                if let Some(error) = item["error"].as_str() {
                    writeln!(
                        output,
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            pytools_core::run_logging::Level::Error,
                            &format!(
                                "写入失败：{}；错误：{error}",
                                item["target"].as_str().unwrap_or("未知目标")
                            )
                        )
                    )?;
                }
            }
        }
        "developer" => {
            let outcome = match run_status {
                "failed" => "失败",
                "partial" => "部分完成",
                "blocked" => "受阻",
                _ => "完成",
            };
            match write_dedup_append_menu_run_history(
                global,
                &started_at,
                elapsed_ms,
                run_status,
                &metrics,
                &stages,
                &actual_engines,
                modifier_invoked,
                writer_invoked,
                blocked_targets > 0 && committed_targets == 0,
                failed
                    .then_some("dedup_append_execution_failed")
                    .or_else(|| {
                        (blocked_targets > 0)
                            .then_some("excel_modifier_capability_required")
                    }),
            ) {
                Ok(path) => writeln!(output, "PT-2 {outcome}：PT-2-5，提交 {committed_targets} 个目标，实际组件：{}，ExcelModifier {}，Excel writer {}，耗时 {elapsed_ms_display_seconds:.3} 秒；运行记录：{}", actual_engines.join(", "), if modifier_invoked { "已调用" } else { "未调用" }, if writer_invoked { "已调用" } else { "未调用" }, display_path_for_user(&path), elapsed_ms_display_seconds = elapsed_ms as f64 / 1000.0)?,
                Err(error) => writeln!(output, "PT-2 {outcome}：PT-2-5，提交 {committed_targets} 个目标，实际组件：{}，ExcelModifier {}，Excel writer {}，耗时 {elapsed_ms_display_seconds:.3} 秒；运行记录写入失败：{error:#}", actual_engines.join(", "), if modifier_invoked { "已调用" } else { "未调用" }, if writer_invoked { "已调用" } else { "未调用" }, elapsed_ms_display_seconds = elapsed_ms as f64 / 1000.0)?,
            }
        }
        "off" if failed => {
            writeln!(output, "处理失败：按启用配置追加到目标。")?;
            for item in report.as_array().into_iter().flatten() {
                if let Some(error) = item["error"].as_str() {
                    writeln!(output, "  错误：{error}")?;
                }
            }
        }
        _ => {}
    }
    // 菜单默认不打印 JSON（user 视图保持纯可读输出）；
    // developer 视图保留机器可读 JSON 以兼容排障与自动化。
    if log_view == "developer" {
        writeln!(output, "{}", serde_json::to_string_pretty(&report)?)?;
    }
    Ok(())
}

fn write_dedup_append_menu_run_history(
    global: &GlobalConfig,
    started_at: &chrono::DateTime<Local>,
    elapsed_ms: u128,
    status: &str,
    metrics: &serde_json::Value,
    stages: &[serde_json::Value],
    actual_engines: &[String],
    excel_modifier_invoked: bool,
    excel_writer_invoked: bool,
    plan_only: bool,
    error_type: Option<&str>,
) -> Result<PathBuf> {
    fs::create_dir_all(&global.log_dir).with_context(|| {
        format!(
            "create dedup log directory {}",
            display_path_for_user(&global.log_dir)
        )
    })?;
    let timestamp = Local::now().format("%Y%m%d_%H%M%S_%f");
    let history = serde_json::json!({
        "schema_version": 1,
        "feature_id": "PT-2-5",
        "action": "config_append",
        "status": status,
        "started_at": started_at.to_rfc3339(),
        "elapsed_ms": elapsed_ms,
        "stages": stages,
        "metrics": metrics,
        "actual_engines": actual_engines,
        "actual_components": actual_engines,
        "excel_modifier_invoked": excel_modifier_invoked,
        "excel_writer_invoked": excel_writer_invoked,
        "plan_only": plan_only,
        "error_type": error_type,
    });
    for suffix in 1..=1000 {
        let suffix = if suffix == 1 {
            String::new()
        } else {
            format!("-{suffix}")
        };
        let path = global.log_dir.join(format!(
            "pytools-dedup-config_append-{timestamp}{suffix}.json"
        ));
        match OpenOptions::new().write(true).create_new(true).open(&path) {
            Ok(mut file) => {
                serde_json::to_writer_pretty(&mut file, &history)?;
                file.write_all(b"\n")?;
                file.sync_all()?;
                return Ok(path);
            }
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => {
                return Err(error)
                    .with_context(|| format!("create {}", display_path_for_user(&path)))
            }
        }
    }
    bail!("unable to allocate a unique dedup append run history path")
}

fn write_temporary_dedup_append_plan(plan: &AppendTargetPlan) -> Result<TemporaryFile> {
    let stamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .context("system clock before Unix epoch")?
        .as_nanos();
    for suffix in 0..1000 {
        let path = std::env::temp_dir().join(format!(
            "pytools-dedup-append-{}-{stamp}-{suffix}.json",
            std::process::id()
        ));
        match OpenOptions::new().write(true).create_new(true).open(&path) {
            Ok(mut file) => {
                let temporary = TemporaryFile::new(path);
                serde_json::to_writer_pretty(&mut file, plan)?;
                file.write_all(b"\n")?;
                file.sync_all()?;
                return Ok(temporary);
            }
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => {
                return Err(error).with_context(|| {
                    format!("create temporary plan {}", display_path_for_user(&path))
                })
            }
        }
    }
    bail!("unable to allocate a temporary dedup append plan")
}

/// 去重任务的可读标识：有名称用名称；无名任务用
/// 「源工作表 → 目标文件：目标工作表」表达业务身份，替代无信息量的“未命名”。
fn dedup_task_label(task: &pytools_core::config::DedupTask) -> String {
    let name = task.task_name.trim();
    if !name.is_empty() {
        return name.to_owned();
    }
    let source_sheet = if task.source_ws.trim().is_empty() {
        "（未指定源工作表）"
    } else {
        task.source_ws.trim()
    };
    let target_file = task
        .target_wb
        .as_ref()
        .and_then(|target| {
            target
                .file_stem()
                .map(|stem| stem.to_string_lossy().into_owned())
        })
        .unwrap_or_else(|| "（未配置目标文件）".to_owned());
    let target_sheet = if task.target_ws.trim().is_empty() {
        "（未指定目标工作表）"
    } else {
        task.target_ws.trim()
    };
    format!("{source_sheet} → {target_file}：{target_sheet}")
}

fn write_dedup_task_start<W: Write>(
    output: &mut W,
    task: &pytools_core::config::DedupTask,
    include_target: bool,
) -> std::io::Result<()> {
    let task_label = dedup_task_label(task);
    let source = task
        .source_wb
        .as_deref()
        .map(|path| path.to_string_lossy().into_owned())
        .unwrap_or_else(|| "（未配置）".to_owned());
    let source_sheet = if task.source_ws.trim().is_empty() {
        "（未指定）"
    } else {
        task.source_ws.trim()
    };
    let key_columns = if task.key_cols.is_empty() {
        "全部已用列（默认）".to_owned()
    } else {
        task.key_cols
            .iter()
            .map(usize::to_string)
            .collect::<Vec<_>>()
            .join(", ")
    };
    let mut line = format!(
        "{} [INFO] 正在检查：配置工作表「{}」第 {} 行「{task_label}」；源文件：{}；源工作表：{}；标识列：{}",
        pytools_core::engine_diag::now_prefix(), task.config_sheet_name, task.row_no,
        source, source_sheet, key_columns
    );
    if include_target {
        let target = task
            .target_wb
            .as_deref()
            .map(|path| path.to_string_lossy().into_owned())
            .unwrap_or_else(|| "（未配置）".to_owned());
        let target_sheet = if task.target_ws.trim().is_empty() {
            "（未指定）"
        } else {
            task.target_ws.trim()
        };
        line.push_str(&format!("；目标文件：{target}；目标工作表：{target_sheet}"));
    }
    writeln!(output, "{line}")?;
    output.flush()
}

fn write_dedup_task_results<W: Write>(
    output: &mut W,
    tasks: &[pytools_core::config::DedupTask],
    events: &[pytools_core::engine_diag::DiagEvent],
    saved_books: &[PathBuf],
    duplicate_label: &str,
) -> std::io::Result<()> {
    use pytools_core::engine_diag::DiagOutcome;

    for task in tasks.iter().filter(|task| task.enabled) {
        let task_event = events
            .iter()
            .find(|event| {
                event.row_no == Some(task.row_no)
                    && event.sheet.as_deref() == Some(task.source_ws.as_str())
            })
            .or_else(|| {
                events.iter().find(|event| {
                    event.row_no == Some(task.row_no) && event.outcome != DiagOutcome::Ok
                })
            });
        let workbook_event = task.source_wb.as_ref().and_then(|source| {
            let source = source.to_string_lossy();
            events.iter().find(|event| {
                event.row_no.is_none()
                    && event.file.as_deref() == Some(source.as_ref())
                    && matches!(event.outcome, DiagOutcome::Skipped | DiagOutcome::Failed)
            })
        });
        let event = task_event.or(workbook_event);
        let prefix = pytools_core::engine_diag::now_prefix();
        if let Some(event) = event
            .filter(|event| matches!(event.outcome, DiagOutcome::Skipped | DiagOutcome::Failed))
        {
            let level = if event.outcome == DiagOutcome::Failed {
                "ERROR"
            } else {
                "WARN"
            };
            let event_message = event.message.clone();
            writeln!(
                output,
                "{prefix} [{level}] 配置结果：第 {} 行「{}」；{}：{event_message}；源工作簿未修改",
                task.row_no,
                dedup_task_label(task),
                if event.outcome == DiagOutcome::Failed {
                    "失败"
                } else {
                    "已跳过"
                },
            )?;
            continue;
        }

        // The per-task diagnostic describes the scan result, independently
        // of whether this operation marks or deletes duplicates.
        let duplicate_count = task_event.and_then(|event| {
            event
                .message
                .strip_prefix("发现重复行 ")
                .and_then(|message| message.split_once(" 行"))
                .and_then(|(count, _)| count.parse::<usize>().ok())
        });
        let Some(duplicate_count) = duplicate_count else {
            writeln!(
                output,
                "{prefix} [WARN] 配置结果：第 {} 行「{}」；未找到对应的重复检查结果",
                task.row_no,
                dedup_task_label(task)
            )?;
            continue;
        };
        let saved = task
            .source_wb
            .as_ref()
            .is_some_and(|source| saved_books.iter().any(|saved_path| saved_path == source));
        // 有结果的行按功能语义给结论；无变化的行明确说明源工作簿未修改。
        let (level, outcome_text) = if duplicate_count > 0 && saved {
            (
                "INFO",
                format!("{duplicate_label} {duplicate_count} 行，已更新并保存源工作簿"),
            )
        } else if duplicate_count > 0 {
            (
                "WARN",
                format!("{duplicate_label} {duplicate_count} 行；源工作簿未保存"),
            )
        } else {
            ("INFO", "未发现重复数据，源工作簿未修改".to_owned())
        };
        writeln!(
            output,
            "{prefix} [{level}] 配置结果：第 {} 行「{}」；{outcome_text}",
            task.row_no,
            dedup_task_label(task)
        )?;
    }
    Ok(())
}

/// Structured diagnostics for configuration-driven operations.  Their task
/// counts must not be presented as per-file counts when inputs are grouped.
struct ConfiguredRunLog {
    tracker: pytools_core::run_logging::RunTracker,
    writer: Option<pytools_core::run_logging::EventWriter>,
    log_dir: PathBuf,
    log_view: String,
    enabled: bool,
    error_summary: Option<pytools_core::run_logging::ErrorSummary>,
}

fn path_for_log_view(path: &Path, log_view: &str) -> String {
    pytools_core::run_logging::path_for_log_view(path, log_view == "developer")
}

fn diagnostic_path_for_log_view(
    writer: Option<&pytools_core::run_logging::EventWriter>,
    log_view: &str,
) -> Option<String> {
    writer.map(|writer| {
        pytools_core::run_logging::diagnostic_path_for_log_view(
            writer.path(),
            log_view == "developer",
        )
    })
}

impl ConfiguredRunLog {
    fn new(global: &GlobalConfig, feature: &str, label: &str, log_view: &str) -> Self {
        use pytools_core::run_logging::{EventWriter, Level, RunTracker};
        let tracker = RunTracker::new(feature, label, &[]);
        let writer = if log_view == "off" {
            None
        } else {
            match EventWriter::create(&global.log_dir.join("run-events"), feature, &tracker.run_id)
            {
                Ok(mut writer) => {
                    if let Err(error) = writer.write(&tracker.event(
                        Level::Info,
                        "run_started",
                        "configuration",
                        "started",
                        format!("开始：{}", tracker.feature_name),
                    )) {
                        eprintln!(
                            "{}",
                            pytools_core::run_logging::terminal_line(
                                Level::Warn,
                                &format!("诊断日志写入失败：{error}")
                            )
                        );
                    }
                    Some(writer)
                }
                Err(error) => {
                    eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            Level::Warn,
                            &format!("无法创建诊断日志：{error}；业务处理继续")
                        )
                    );
                    None
                }
            }
        };
        Self {
            tracker,
            writer,
            log_dir: global.log_dir.clone(),
            log_view: log_view.to_owned(),
            enabled: log_view != "off",
            error_summary: None,
        }
    }

    fn set_business_metric(&mut self, name: &str, value: serde_json::Value) {
        self.tracker.set_business_metric(name, value);
    }

    fn record_business_issue(&mut self, issue: pytools_core::run_logging::BusinessIssue) {
        self.tracker.record_business_issue(issue);
    }

    fn set_error_summary(&mut self, error: pytools_core::run_logging::ErrorSummary) {
        if self.error_summary.is_none() {
            self.error_summary = Some(error);
        }
    }

    fn record_stages(&mut self, stages: &[serde_json::Value]) {
        if let Some(writer) = self.writer.as_mut() {
            for stage in stages {
                let name = stage["name"].as_str().unwrap_or("configuration");
                let mut event = self.tracker.event(
                    pytools_core::run_logging::Level::Debug,
                    "configuration_stage",
                    name,
                    "completed",
                    format!("{}：配置阶段「{name}」完成", self.tracker.feature_name),
                );
                event.elapsed_ms = stage["elapsed_ms"].as_u64().map(u128::from);
                if let Err(error) = writer.write(&event) {
                    eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            pytools_core::run_logging::Level::Warn,
                            &format!("诊断日志写入失败：{error}")
                        )
                    );
                    break;
                }
            }
        }
    }

    fn finish(
        self,
        failed: bool,
        task_counts: Option<(usize, usize, usize)>,
        events: &[pytools_core::engine_diag::DiagEvent],
        outputs: &[PathBuf],
    ) -> pytools_core::run_logging::RunSummary {
        self.finish_as(failed, task_counts, events, outputs, None)
    }

    fn finish_as(
        mut self,
        failed: bool,
        task_counts: Option<(usize, usize, usize)>,
        events: &[pytools_core::engine_diag::DiagEvent],
        outputs: &[PathBuf],
        override_status: Option<pytools_core::run_logging::TaskStatus>,
    ) -> pytools_core::run_logging::RunSummary {
        use pytools_core::engine_diag::DiagOutcome;
        use pytools_core::run_logging::{Level, TaskStatus};
        let developer_view = self.log_view == "developer";
        let path = self.writer.as_ref().map(|writer| {
            pytools_core::run_logging::diagnostic_path_for_log_view(writer.path(), developer_view)
        });
        let mut summary = self.tracker.finish(
            failed,
            false,
            false,
            outputs
                .iter()
                .map(|path| pytools_core::run_logging::path_for_log_view(path, developer_view))
                .collect(),
            path,
        );
        summary.error_summary = self.error_summary.clone();
        summary.input_stats_available = false;
        if let Some((ok, skipped, failures)) = task_counts {
            summary.overall_status = if failed || (failures > 0 && ok == 0) {
                TaskStatus::Failed
            } else if failures > 0 {
                TaskStatus::PartialSuccess
            } else if ok == 0 && skipped > 0 {
                TaskStatus::NoOp
            } else {
                TaskStatus::Success
            };
            summary
                .business_metrics
                .insert("成功配置任务".to_owned(), serde_json::json!(ok));
            summary
                .business_metrics
                .insert("跳过配置任务".to_owned(), serde_json::json!(skipped));
            summary
                .business_metrics
                .insert("失败配置任务".to_owned(), serde_json::json!(failures));
        }
        if let Some(status) = override_status {
            summary.overall_status = status;
        }
        if let Some(writer) = self.writer.as_mut() {
            for source in events {
                let level = match source.outcome {
                    DiagOutcome::Failed => Level::Error,
                    DiagOutcome::Skipped => Level::Warn,
                    DiagOutcome::Started => Level::Info,
                    DiagOutcome::Ok => Level::Debug,
                };
                let mut event = self.tracker.event(
                    level,
                    "engine_operation",
                    &source.stage.to_string(),
                    &source.outcome.to_string(),
                    format!("{}：{}", source.stage, source.outcome),
                );
                event.timestamp = source.timestamp_rfc3339.clone();
                event.engine = Some(source.engine.clone());
                event.config_row = source.row_no;
                event.file_name = source
                    .file
                    .as_deref()
                    .map(|value| pytools_core::run_logging::portable_file_name(Path::new(value)));
                event.sheet_name = source.sheet.clone();
                event.elapsed_ms = source.elapsed_ms;
                if source.outcome == DiagOutcome::Failed {
                    event.error_code = Some(
                        match source.stage {
                            pytools_core::engine_diag::DiagStage::ReadConfig => {
                                "CONFIG_READ_FAILED"
                            }
                            pytools_core::engine_diag::DiagStage::ReadSource => "EXCEL_READ_FAILED",
                            pytools_core::engine_diag::DiagStage::CopySheet => "SHEET_COPY_FAILED",
                            pytools_core::engine_diag::DiagStage::ModifyWorkbook => {
                                "EXCEL_MODIFY_FAILED"
                            }
                            pytools_core::engine_diag::DiagStage::SaveOutput => {
                                "OUTPUT_WRITE_FAILED"
                            }
                        }
                        .to_owned(),
                    );
                    if self.log_view == "developer" {
                        event.details = Some(serde_json::json!({"diagnostic": source.message}));
                    }
                }
                if let Err(error) = writer.write(&event) {
                    eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            Level::Warn,
                            &format!("诊断日志写入失败：{error}")
                        )
                    );
                    break;
                }
            }
            if let Err(error) = writer.write(&self.tracker.event(
                if failed { Level::Error } else { Level::Info },
                "run_finished",
                "run",
                &format!("{:?}", summary.overall_status),
                format!("{}结束", self.tracker.feature_name),
            )) {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("诊断日志写入失败：{error}")
                    )
                );
            }
        }
        if self.enabled {
            match serde_json::to_value(&summary) {
                Ok(mut value) => {
                    if let Some((ok, skipped, failures)) = task_counts {
                        value["task_stats"] = serde_json::json!({"succeeded": ok, "skipped": skipped, "failed": failures});
                    }
                    let path = self.log_dir.join(format!(
                        "{}_{}.run.json",
                        self.tracker.feature_id, self.tracker.run_id
                    ));
                    match write_json_report_create_new(&path, &value) {
                        Ok(()) => eprintln!(
                            "{}",
                            pytools_core::run_logging::terminal_line(
                                Level::Info,
                                &format!(
                                    "运行编号 {}；运行记录：{}",
                                    self.tracker.run_id,
                                    display_path_for_user(&path)
                                )
                            )
                        ),
                        Err(error) => eprintln!(
                            "{}",
                            pytools_core::run_logging::terminal_line(
                                Level::Warn,
                                &format!("运行记录写入失败：{error}；业务结果仍保留")
                            )
                        ),
                    }
                }
                Err(error) => eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("运行记录序列化失败：{error}；业务结果仍保留")
                    )
                ),
            }
        }
        summary
    }
}

/// Shared run log adapter for file operations that already expose structured
/// reports or progress callbacks. It keeps logging failures out of the business
/// result and only persists raw error chains in explicit developer mode.
struct UnifiedRunSession {
    tracker: pytools_core::run_logging::RunTracker,
    writer: Option<pytools_core::run_logging::EventWriter>,
    log_dir: PathBuf,
    log_view: String,
    input_stats_available: bool,
    first_error: Option<pytools_core::run_logging::ErrorSummary>,
    warned_log_failure: bool,
}

impl UnifiedRunSession {
    fn new(
        log_dir: &Path,
        feature_id: &str,
        feature_name: &str,
        log_view: &str,
        run_id: Option<String>,
    ) -> Self {
        use pytools_core::run_logging::{EventWriter, Level, RunTracker};
        let tracker = run_id.map_or_else(
            || RunTracker::new(feature_id, feature_name, &[]),
            |run_id| RunTracker::with_id(feature_id, feature_name, &[], run_id),
        );
        let mut session = Self {
            tracker,
            writer: None,
            log_dir: log_dir.to_path_buf(),
            log_view: log_view.to_owned(),
            input_stats_available: true,
            first_error: None,
            warned_log_failure: false,
        };
        if log_view != "off" {
            match EventWriter::create(
                &log_dir.join("run-events"),
                feature_id,
                &session.tracker.run_id,
            ) {
                Ok(writer) => session.writer = Some(writer),
                Err(error) => session.warn_log_failure("无法创建诊断日志", &error),
            }
        }
        session.progress(
            Level::Info,
            "run_started",
            "run",
            "started",
            &format!("开始：{}", session.tracker.feature_name),
            None,
            None,
            None,
        );
        session
    }

    fn set_inputs(&mut self, files: &[PathBuf], discovered: Option<usize>) {
        if self.tracker.set_inputs(files).is_err() {
            self.input_stats_available = false;
        }
        if let Some(discovered) = discovered {
            self.tracker.set_discovered(discovered);
        }
    }

    fn set_discovered(&mut self, discovered: usize) {
        self.tracker.set_discovered(discovered);
    }

    fn set_business_metric(&mut self, name: &str, value: serde_json::Value) {
        self.tracker.set_business_metric(name, value);
    }

    fn record_business_issue(&mut self, issue: pytools_core::run_logging::BusinessIssue) {
        self.tracker.record_business_issue(issue);
    }

    fn mark_inputs_unavailable(&mut self) {
        self.input_stats_available = false;
    }

    fn set_existing_timing(&mut self, started_at: String, elapsed_ms: u128) {
        self.tracker.set_existing_timing(started_at, elapsed_ms);
    }

    fn progress(
        &mut self,
        level: pytools_core::run_logging::Level,
        event_name: &str,
        stage: &str,
        status: &str,
        message: &str,
        elapsed_ms: Option<u128>,
        engine: Option<&str>,
        details: Option<serde_json::Value>,
    ) {
        self.progress_at(
            level, event_name, stage, status, message, elapsed_ms, engine, details, None,
        );
    }

    fn progress_at(
        &mut self,
        level: pytools_core::run_logging::Level,
        event_name: &str,
        stage: &str,
        status: &str,
        message: &str,
        elapsed_ms: Option<u128>,
        engine: Option<&str>,
        details: Option<serde_json::Value>,
        timestamp: Option<&str>,
    ) {
        let mut event = self
            .tracker
            .event(level, event_name, stage, status, message);
        if let Some(timestamp) = timestamp {
            event.timestamp = timestamp.to_owned();
        }
        event.elapsed_ms = elapsed_ms;
        event.engine = engine.map(str::to_owned);
        if self.log_view == "developer" {
            event.details = details;
        }
        self.write_event(&event);
    }

    fn file_result(
        &mut self,
        index: usize,
        status: pytools_core::run_logging::FileStatus,
        attempted: bool,
        stage: &str,
        message: &str,
        error_code: Option<&str>,
        details: Option<String>,
    ) {
        use pytools_core::run_logging::{FileStatus, Level};
        let prior = self.tracker.state(index);
        if prior == Some(FileStatus::Pending) {
            if attempted {
                let _ = self.tracker.transition(index, FileStatus::InProgress);
            }
            let _ = self.tracker.transition(index, status);
        }
        let input = self.tracker.input_result(index);
        let mut event = self.tracker.event(
            match status {
                FileStatus::Failed => Level::Error,
                FileStatus::Skipped => Level::Warn,
                _ => Level::Info,
            },
            "file_result",
            stage,
            match status {
                FileStatus::Succeeded => "succeeded",
                FileStatus::Failed => "failed",
                FileStatus::Skipped => "skipped",
                FileStatus::Cancelled => "cancelled",
                FileStatus::Pending => "pending",
                FileStatus::InProgress => "in_progress",
            },
            message,
        );
        event.file_index = Some(index);
        event.file_name = input.map(|item| item.file_name.clone());
        event.error_code = error_code.map(str::to_owned);
        if self.log_view == "developer" {
            event.details =
                details.map(|error_chain| serde_json::json!({"error_chain": error_chain}));
        }
        if status == FileStatus::Failed && self.first_error.is_none() {
            self.first_error = Some(pytools_core::run_logging::ErrorSummary {
                stage: stage.to_owned(),
                code: error_code.unwrap_or("INTERNAL_OPERATION_FAILED").to_owned(),
                message: message.to_owned(),
                file_name: event.file_name.clone(),
                sheet_name: None,
            });
        }
        self.write_event(&event);
    }

    fn discovered_file_event(
        &mut self,
        index: usize,
        path: &Path,
        status: pytools_core::run_logging::FileStatus,
        stage: &str,
        message: &str,
        error_code: Option<&str>,
        details: Option<String>,
    ) {
        use pytools_core::run_logging::{FileStatus, Level};
        let mut event = self.tracker.event(
            match status {
                FileStatus::Failed => Level::Error,
                FileStatus::Skipped => Level::Warn,
                _ => Level::Info,
            },
            "discovered_file_result",
            stage,
            match status {
                FileStatus::Succeeded => "succeeded",
                FileStatus::Failed => "failed",
                FileStatus::Skipped => "skipped",
                FileStatus::Cancelled => "cancelled",
                FileStatus::Pending => "pending",
                FileStatus::InProgress => "in_progress",
            },
            message,
        );
        event.file_index = Some(index);
        event.file_name = Some(pytools_core::run_logging::portable_file_name(path));
        event.error_code = error_code.map(str::to_owned);
        if self.log_view == "developer" {
            event.details =
                details.map(|error_chain| serde_json::json!({"diagnostic": error_chain}));
        }
        if status == FileStatus::Failed && self.first_error.is_none() {
            self.first_error = Some(pytools_core::run_logging::ErrorSummary {
                stage: stage.to_owned(),
                code: error_code.unwrap_or("INTERNAL_OPERATION_FAILED").to_owned(),
                message: message.to_owned(),
                file_name: event.file_name.clone(),
                sheet_name: None,
            });
        }
        self.write_event(&event);
    }

    fn operation_failure(
        &mut self,
        stage: &str,
        code: &str,
        message: &str,
        error_chain: &anyhow::Error,
    ) {
        use pytools_core::run_logging::Level;
        self.first_error
            .get_or_insert_with(|| pytools_core::run_logging::ErrorSummary {
                stage: stage.to_owned(),
                code: code.to_owned(),
                message: message.to_owned(),
                file_name: None,
                sheet_name: None,
            });
        self.progress(
            Level::Error,
            "operation_failed",
            stage,
            "failed",
            message,
            None,
            None,
            Some(serde_json::json!({"error_chain": format!("{error_chain:#}")})),
        );
    }

    fn finish(
        self,
        critical_failure: bool,
        cancelled: bool,
        no_op: bool,
        outputs: &[PathBuf],
        output_stats: Option<pytools_core::run_logging::OutputStats>,
        record_stats: Option<pytools_core::run_logging::RecordStats>,
        sheet_stats: Option<pytools_core::run_logging::SheetStats>,
    ) -> (pytools_core::run_logging::RunSummary, Option<PathBuf>) {
        self.finish_with_status(
            critical_failure,
            cancelled,
            no_op,
            outputs,
            output_stats,
            record_stats,
            sheet_stats,
            None,
        )
    }

    fn finish_with_status(
        mut self,
        critical_failure: bool,
        cancelled: bool,
        no_op: bool,
        outputs: &[PathBuf],
        output_stats: Option<pytools_core::run_logging::OutputStats>,
        record_stats: Option<pytools_core::run_logging::RecordStats>,
        sheet_stats: Option<pytools_core::run_logging::SheetStats>,
        status_override: Option<pytools_core::run_logging::TaskStatus>,
    ) -> (pytools_core::run_logging::RunSummary, Option<PathBuf>) {
        use pytools_core::run_logging::{Level, TaskStatus};
        let diagnostic_path = self.writer.as_ref().map(|writer| {
            pytools_core::run_logging::diagnostic_path_for_log_view(
                writer.path(),
                self.log_view == "developer",
            )
        });
        let mut summary = self.tracker.finish(
            critical_failure,
            cancelled,
            no_op,
            outputs
                .iter()
                .map(|path| {
                    pytools_core::run_logging::path_for_log_view(path, self.log_view == "developer")
                })
                .collect(),
            diagnostic_path,
        );
        summary.input_stats_available = self.input_stats_available;
        summary.output_stats = output_stats;
        summary.record_stats = record_stats;
        summary.sheet_stats = sheet_stats;
        summary.error_summary = self.first_error.clone();
        if let Some(status) = status_override {
            summary.overall_status = status;
        }
        if !critical_failure
            && summary.input_stats.failed == 0
            && summary.input_stats.pending == 0
            && no_op
        {
            summary.overall_status = TaskStatus::NoOp;
        }
        self.progress(
            if matches!(summary.overall_status, TaskStatus::Failed) {
                Level::Error
            } else {
                Level::Info
            },
            "run_finished",
            "run",
            &format!("{:?}", summary.overall_status),
            &format!("{}结束", self.tracker.feature_name),
            Some(summary.total_elapsed_ms),
            None,
            None,
        );
        if self.log_view == "off" {
            return (summary, None);
        }
        let path = self.log_dir.join(format!(
            "{}_{}.run.json",
            self.tracker.feature_id, self.tracker.run_id
        ));
        match write_json_report_create_new(&path, &summary) {
            Ok(()) => (summary, Some(path)),
            Err(error) => {
                self.warn_log_failure("运行汇总写入失败；业务结果仍保留", &error);
                (summary, None)
            }
        }
    }

    fn write_event(&mut self, event: &pytools_core::run_logging::LogEvent) {
        let error = self
            .writer
            .as_mut()
            .and_then(|writer| writer.write(event).err());
        if let Some(error) = error {
            self.warn_log_failure("诊断日志写入失败；业务处理继续", &error);
        }
    }

    fn warn_log_failure(&mut self, action: &str, error: &dyn std::fmt::Display) {
        if self.warned_log_failure {
            return;
        }
        self.warned_log_failure = true;
        eprintln!(
            "{}：{}",
            pytools_core::run_logging::terminal_line(
                pytools_core::run_logging::Level::Warn,
                action
            ),
            error
        );
    }
}

fn log_quick_menu_action<W: Write>(
    output: &mut W,
    global: &GlobalConfig,
    log_view: &str,
    feature_id: &str,
    label: &str,
    result: Result<Option<PathBuf>>,
    no_result_message: String,
) -> Result<()> {
    use pytools_core::run_logging::{terminal_line, Level, TaskStatus};
    let mut diagnostic = ConfiguredRunLog::new(global, feature_id, label, log_view);
    match result {
        Ok(Some(path)) => {
            diagnostic.finish_as(
                false,
                Some((1, 0, 0)),
                &[],
                std::slice::from_ref(&path),
                Some(TaskStatus::Success),
            );
            writeln!(
                output,
                "{}",
                terminal_line(
                    Level::Info,
                    &format!("已执行：{label}；位置：{}", display_path_for_user(&path))
                )
            )?;
        }
        Ok(None) => {
            diagnostic.finish_as(false, Some((0, 1, 0)), &[], &[], Some(TaskStatus::NoOp));
            writeln!(output, "{}", terminal_line(Level::Warn, &no_result_message))?;
        }
        Err(error) => {
            diagnostic.finish_as(true, Some((0, 0, 1)), &[], &[], Some(TaskStatus::Failed));
            let message = if log_view == "developer" {
                format!("{label}失败：{error:#}")
            } else {
                format!("{label}失败；请检查文件是否存在以及系统是否允许打开")
            };
            writeln!(output, "{}", terminal_line(Level::Error, &message))?;
        }
    }
    Ok(())
}

fn run_dedup_menu_action<W: Write>(
    output: &mut W,
    global: &GlobalConfig,
    log_view: &str,
    feature_id: &str,
    action: &str,
    label: &str,
    actual_components: &[&str],
    arguments: Vec<String>,
    run: fn(Vec<String>) -> Result<serde_json::Value>,
) -> Result<()> {
    let mut diagnostic = ConfiguredRunLog::new(global, feature_id, label, log_view);
    let started_at = Local::now();
    let started = Instant::now();
    match log_view {
        "user" => writeln!(output, "开始：{label}")?,
        "developer" => writeln!(output, "PT-2 开始：{feature_id} {action}")?,
        _ => {}
    }

    let report = match run(arguments) {
        Ok(report) => report,
        Err(error) => {
            diagnostic.set_error_summary(pytools_core::run_logging::ErrorSummary {
                stage: action.to_owned(),
                code: "DEDUP_OPERATION_FAILED".to_owned(),
                message: format!("{label}失败：{error:#}"),
                file_name: None,
                sheet_name: None,
            });
            let summary = diagnostic.finish(true, None, &[], &[]);
            if log_view == "user" {
                write_unified_terminal_summary(output, &summary, &[])?;
            }
            let elapsed_ms = started.elapsed().as_millis();
            if log_view == "developer" {
                match write_dedup_menu_run_history(
                    global,
                    feature_id,
                    action,
                    "failed",
                    &started_at,
                    elapsed_ms,
                    &serde_json::json!({}),
                    &[],
                    Some("dedup_operation_failed"),
                ) {
                    Ok(path) => writeln!(output, "PT-2 失败：{feature_id}，耗时 {elapsed_ms_display_seconds:.3} 秒；运行记录：{}；错误：{error:#}", display_path_for_user(&path), elapsed_ms_display_seconds = elapsed_ms as f64 / 1000.0)?,
                    Err(log_error) => writeln!(output, "PT-2 失败：{feature_id}，耗时 {elapsed_ms_display_seconds:.3} 秒；运行记录写入失败：{log_error:#}；错误：{error:#}", elapsed_ms_display_seconds = elapsed_ms as f64 / 1000.0)?,
                }
            }
            return Ok(());
        }
    };

    let elapsed_ms = started.elapsed().as_millis();
    let metrics = dedup_menu_metrics(&report);
    let precheck_issues =
        action == "config_precheck" && metrics["failed_tasks"].as_u64().unwrap_or_default() > 0;
    let writes_duplicate_output = matches!(action, "config_check" | "config_delete");
    let failed_outputs =
        writes_duplicate_output && metrics["failed_targets"].as_u64().unwrap_or_default() > 0;
    let status = if failed_outputs && metrics["written_outputs"].as_u64().unwrap_or_default() > 0 {
        "partial"
    } else if failed_outputs {
        "failed"
    } else {
        "succeeded"
    };
    let mut components = actual_components.to_vec();
    if writes_duplicate_output
        && metrics["written_outputs"].as_u64().unwrap_or_default() > 0
        && !components.contains(&"rust_xlsxwriter")
    {
        components.push("rust_xlsxwriter");
    }
    let error_type = if failed_outputs {
        Some("dedup_result_output_failures")
    } else {
        None
    };
    let mut summary_metrics = Vec::<(&str, &str)>::new();
    match action {
        "config_precheck" => summary_metrics.extend([
            ("检查任务数", "processed_groups"),
            ("通过配置数", "passed_tasks"),
            ("警告配置数", "warning_tasks"),
            ("失败配置数", "failed_tasks"),
        ]),
        "config_check" => summary_metrics.extend([
            ("重复行数", "duplicate_rows"),
            ("命中工作表数", "hit_sheets"),
            ("保存工作簿数", "written_outputs"),
            ("失败目标数", "failed_targets"),
            ("跳过配置数", "skipped_tasks"),
        ]),
        "config_delete" => summary_metrics.extend([
            ("删除重复行数", "duplicate_rows"),
            ("命中工作表数", "hit_sheets"),
            ("保存工作簿数", "written_outputs"),
            ("失败目标数", "failed_targets"),
            ("跳过配置数", "skipped_tasks"),
        ]),
        "config_append" => summary_metrics.extend([
            ("追加行数", "appended_rows"),
            ("过滤重复行数", "duplicate_rows_filtered"),
            ("写入目标数", "committed_targets"),
            ("失败目标数", "failed_targets"),
            ("跳过配置数", "skipped_tasks"),
        ]),
        _ => summary_metrics.push(("处理组数", "processed_groups")),
    }
    for (label, key) in summary_metrics {
        if let Some(value) = metrics.get(key) {
            diagnostic.set_business_metric(label, value.clone());
        }
    }
    if precheck_issues {
        diagnostic.record_business_issue(
            pytools_core::run_logging::BusinessIssue::new(
                "配置校验",
                "预校验发现未通过的配置项；处理流程已正常完成",
            )
            .count(metrics["failed_tasks"].as_u64().unwrap_or_default() as usize),
        );
    }
    let precheck_warnings =
        action == "config_precheck" && metrics["warning_tasks"].as_u64().unwrap_or_default() > 0;
    if precheck_warnings {
        diagnostic.record_business_issue(
            pytools_core::run_logging::BusinessIssue::new(
                "配置警告",
                "预校验发现需要关注的配置项；处理流程已正常完成",
            )
            .count(metrics["warning_tasks"].as_u64().unwrap_or_default() as usize),
        );
    }
    if action == "config_check" && metrics["duplicate_rows"].as_u64().unwrap_or_default() > 0 {
        diagnostic.record_business_issue(
            pytools_core::run_logging::BusinessIssue::new("重复数据", "检查发现重复行")
                .count(metrics["duplicate_rows"].as_u64().unwrap_or_default() as usize),
        );
    }
    let summary_status = if failed_outputs {
        if metrics["written_outputs"].as_u64().unwrap_or_default() > 0 {
            pytools_core::run_logging::TaskStatus::PartialSuccess
        } else {
            pytools_core::run_logging::TaskStatus::Failed
        }
    } else {
        pytools_core::run_logging::TaskStatus::Success
    };
    if failed_outputs {
        let first_failure = report
            .as_array()
            .into_iter()
            .flatten()
            .find(|item| item["status"] == "failed");
        diagnostic.set_error_summary(pytools_core::run_logging::ErrorSummary {
            stage: action.to_owned(),
            code: "DEDUP_RESULT_OUTPUT_FAILED".to_owned(),
            message: first_failure
                .and_then(|item| item["error"].as_str())
                .unwrap_or("重复检查结果写入失败")
                .to_owned(),
            file_name: first_failure
                .and_then(|item| item["report"]["workbook"].as_str())
                .map(|path| pytools_core::run_logging::portable_file_name(Path::new(path))),
            sheet_name: None,
        });
    }
    let summary = diagnostic.finish_as(failed_outputs, None, &[], &[], Some(summary_status));

    match log_view {
        "user" => {
            write_unified_terminal_summary(output, &summary, &[])?;
            if let Some(issues) = report["issues"].as_array() {
                for issue in issues.iter().filter_map(serde_json::Value::as_str) {
                    writeln!(output, "  {issue}")?;
                }
            }
            if let Some(items) = report.as_array() {
                for item in items.iter().filter(|item| item["status"] == "failed") {
                    let workbook = item["report"]["workbook"].as_str().unwrap_or("工作簿");
                    let error = item["error"].as_str().unwrap_or("未知输出错误");
                    writeln!(output, "  {workbook}：{error}")?;
                }
            }
        }
        "developer" => {
            let outcome = match status {
                "failed" => "失败",
                "partial" => "部分完成",
                _ => "完成",
            };
            match write_dedup_menu_run_history(
                global,
                feature_id,
                action,
                status,
                &started_at,
                elapsed_ms,
                &metrics,
                    &components,
                    error_type,
            ) {
                Ok(path) => writeln!(output, "PT-2 {outcome}：{feature_id}，{}，实际组件：{}，ExcelModifier 未调用，耗时 {elapsed_ms_display_seconds:.3} 秒；运行记录：{}", dedup_menu_summary(action, &metrics), components.join(", "), display_path_for_user(&path), elapsed_ms_display_seconds = elapsed_ms as f64 / 1000.0)?,
                Err(error) => writeln!(output, "PT-2 {outcome}：{feature_id}，{}，实际组件：{}，ExcelModifier 未调用，耗时 {elapsed_ms_display_seconds:.3} 秒；运行记录写入失败：{error:#}", dedup_menu_summary(action, &metrics), components.join(", "), elapsed_ms_display_seconds = elapsed_ms as f64 / 1000.0)?,
            }
        }
        _ => {}
    }

    // 菜单默认不打印 JSON（user 视图保持纯可读输出）；
    // developer 视图保留机器可读 JSON 以兼容排障与自动化。
    if log_view == "developer" {
        writeln!(output, "{}", serde_json::to_string_pretty(&report)?)?;
    }
    Ok(())
}

fn dedup_menu_metrics(report: &serde_json::Value) -> serde_json::Value {
    let items = report
        .as_array()
        .map(|items| items.iter().collect::<Vec<_>>())
        .unwrap_or_else(|| vec![report]);
    let mut duplicate_rows = 0u64;
    let mut hit_sheets = 0u64;
    let mut skipped_tasks = 0u64;
    let mut appended_rows = 0u64;
    let mut duplicate_rows_filtered = 0u64;
    let mut planned_operations = 0u64;
    let mut committed_targets = 0u64;
    let mut planned_only_targets = 0u64;
    let mut failed_targets = 0u64;
    let mut written_outputs = 0u64;
    for item in &items {
        let details = item.get("report").unwrap_or(item);
        duplicate_rows += details["duplicate_rows"].as_u64().unwrap_or_default();
        hit_sheets += details["hit_sheets"].as_u64().unwrap_or_default();
        skipped_tasks += details["skipped_tasks"]
            .as_u64()
            .or_else(|| item["task_skipped"].as_u64())
            .unwrap_or_default();
        appended_rows += item["appended_rows"].as_u64().unwrap_or_default();
        duplicate_rows_filtered += item["duplicate_rows_filtered"].as_u64().unwrap_or_default();
        planned_operations += details["plan"]["operations"]
            .as_array()
            .map_or(0, |operations| operations.len() as u64);
        committed_targets += u64::from(item["execution"] == "committed");
        planned_only_targets +=
            u64::from(item["status"] == "planned_only" || item["status"] == "cancelled");
        failed_targets += u64::from(item["status"] == "failed");
        written_outputs += u64::from(item["output_file"].is_string());
    }
    serde_json::json!({
        "processed_groups": items.len(),
        "duplicate_rows": duplicate_rows,
        "hit_sheets": hit_sheets,
        "skipped_tasks": skipped_tasks,
        "appended_rows": appended_rows,
        "duplicate_rows_filtered": duplicate_rows_filtered,
        "planned_operations": planned_operations,
        "committed_targets": committed_targets,
        "planned_only_targets": planned_only_targets,
        "failed_targets": failed_targets,
        "written_outputs": written_outputs,
        "passed_tasks": report["ok"].as_u64().unwrap_or_default(),
        "warning_tasks": report["warn"].as_u64().unwrap_or_default(),
        "failed_tasks": report["fail"].as_u64().unwrap_or_default(),
    })
}

fn dedup_menu_summary(action: &str, metrics: &serde_json::Value) -> String {
    if action == "config_precheck" {
        return format!(
            "通过 {} 项、警告 {} 项、失败 {} 项",
            metrics["passed_tasks"].as_u64().unwrap_or_default(),
            metrics["warning_tasks"].as_u64().unwrap_or_default(),
            metrics["failed_tasks"].as_u64().unwrap_or_default()
        );
    }
    if action == "config_append" {
        return format!(
            "提交 {} 个目标，保留计划 {} 个目标，追加 {} 行，过滤重复 {} 行，跳过 {} 项",
            metrics["committed_targets"].as_u64().unwrap_or_default(),
            metrics["planned_only_targets"].as_u64().unwrap_or_default(),
            metrics["appended_rows"].as_u64().unwrap_or_default(),
            metrics["duplicate_rows_filtered"]
                .as_u64()
                .unwrap_or_default(),
            metrics["skipped_tasks"].as_u64().unwrap_or_default()
        );
    }
    if matches!(action, "config_check" | "config_delete") {
        let verb = if action == "config_delete" {
            "删除重复行"
        } else {
            "标记重复行"
        };
        return format!(
            "生成 {} 个结果工作簿，{} {} 条，失败 {} 个工作簿，跳过 {} 项",
            metrics["written_outputs"].as_u64().unwrap_or_default(),
            verb,
            metrics["duplicate_rows"].as_u64().unwrap_or_default(),
            metrics["failed_targets"].as_u64().unwrap_or_default(),
            metrics["skipped_tasks"].as_u64().unwrap_or_default()
        );
    }
    format!(
        "处理 {} 个工作簿组，命中 {} 张表，发现 {} 条重复行，计划操作 {} 项",
        metrics["processed_groups"].as_u64().unwrap_or_default(),
        metrics["hit_sheets"].as_u64().unwrap_or_default(),
        metrics["duplicate_rows"].as_u64().unwrap_or_default(),
        metrics["planned_operations"].as_u64().unwrap_or_default()
    )
}

fn write_dedup_menu_run_history(
    global: &GlobalConfig,
    feature_id: &str,
    action: &str,
    status: &str,
    started_at: &chrono::DateTime<Local>,
    elapsed_ms: u128,
    metrics: &serde_json::Value,
    actual_components: &[&str],
    error_type: Option<&str>,
) -> Result<PathBuf> {
    fs::create_dir_all(&global.log_dir).with_context(|| {
        format!(
            "create dedup log directory {}",
            display_path_for_user(&global.log_dir)
        )
    })?;
    let timestamp = Local::now().format("%Y%m%d_%H%M%S_%f");
    let writer_invoked = actual_components.contains(&"rust_xlsxwriter");
    let stage = if action == "config_precheck" {
        "configuration_precheck"
    } else if writer_invoked {
        "read_plan_write_and_verify_new_workbook"
    } else {
        "workbook_read_and_plan"
    };
    let history = serde_json::json!({
        "schema_version": 1,
        "feature_id": feature_id,
        "action": action,
        "status": status,
        "started_at": started_at.to_rfc3339(),
        "elapsed_ms": elapsed_ms,
        "stages": [{"name": stage, "elapsed_ms": elapsed_ms}],
        "metrics": metrics,
        "actual_engines": actual_components,
        "excel_modifier_invoked": false,
        "excel_writer_invoked": writer_invoked,
        "plan_only": action != "config_precheck" && !writer_invoked,
        "error_type": error_type,
    });
    for suffix in 1..=1000 {
        let suffix = if suffix == 1 {
            String::new()
        } else {
            format!("-{suffix}")
        };
        let path = global
            .log_dir
            .join(format!("pytools-dedup-{action}-{timestamp}{suffix}.json"));
        match OpenOptions::new().write(true).create_new(true).open(&path) {
            Ok(mut file) => {
                serde_json::to_writer_pretty(&mut file, &history)?;
                file.write_all(b"\n")?;
                file.sync_all()?;
                return Ok(path);
            }
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => {
                return Err(error)
                    .with_context(|| format!("create {}", display_path_for_user(&path)))
            }
        }
    }
    bail!("unable to allocate a unique dedup run history path")
}

fn unique_menu_output_dir(output_root: &Path, action: &str) -> PathBuf {
    let timestamp = Local::now().format("%Y%m%d_%H%M%S_%f");
    let stem = format!("dedup-{action}-{timestamp}");
    let mut candidate = output_root.join(&stem);
    let mut suffix = 2;
    while candidate.exists() {
        candidate = output_root.join(format!("{stem}-{suffix}"));
        suffix += 1;
    }
    candidate
}

fn unique_menu_action_dir(output_root: &Path, action: &str) -> PathBuf {
    let timestamp = Local::now().format("%Y%m%d_%H%M%S_%f");
    let stem = format!("pytools-{action}-{timestamp}");
    let mut candidate = output_root.join(&stem);
    let mut suffix = 2;
    while candidate.exists() {
        candidate = output_root.join(format!("{stem}-{suffix}"));
        suffix += 1;
    }
    candidate
}

/// 3-1/3-2：按配置直接执行静态打印，不做 y 确认；两个历史入口统一走 Openpyxl。
/// 运行日志统一带本机时间戳；失败显示阶段、引擎与原始错误；输出文件去重显示。
fn run_configured_static_print_menu<W>(
    output: &mut W,
    config_path: &Path,
    global: &GlobalConfig,
    feature_id: &str,
    engine: pytools_core::static_print_engines::StaticPrintEngine,
    log_view: &str,
) -> Result<()>
where
    W: Write,
{
    use pytools_core::engine_diag::{DiagOutcome, DiagStage};
    use pytools_core::run_logging::{EventWriter, Level, RunTracker, TaskStatus};
    let tracker = RunTracker::new(feature_id, "按配置生成静态打印文件", &[]);
    let mut diagnostic = if log_view == "off" {
        None
    } else {
        match EventWriter::create(
            &global.log_dir.join("run-events"),
            feature_id,
            &tracker.run_id,
        ) {
            Ok(writer) => Some(writer),
            Err(error) => {
                writeln!(
                    output,
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("无法创建诊断日志：{error}；业务处理继续")
                    )
                )?;
                None
            }
        }
    };
    let mut log_failed = false;
    if let Some(writer) = diagnostic.as_mut() {
        if let Err(error) = writer.write(&tracker.event(
            Level::Info,
            "run_started",
            "configuration",
            "started",
            "开始按配置生成打印文件",
        )) {
            writeln!(
                output,
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Warn,
                    &format!("诊断日志写入失败：{error}")
                )
            )?;
            log_failed = true;
        }
    }
    writeln!(
        output,
        "{} [INFO] 开始按打印配置生成静态打印文件（{}）",
        pytools_core::engine_diag::now_prefix(),
        engine.label()
    )?;
    let print_tasks_for_display = pytools_core::config::load_print_tasks(config_path).ok();
    if let Some(tasks) = print_tasks_for_display.as_ref() {
        let enabled_count = tasks.iter().filter(|task| task.enabled).count();
        writeln!(
            output,
            "{} [INFO] 已启用配置：{enabled_count} 项",
            pytools_core::engine_diag::now_prefix()
        )?;
        for task in tasks.iter().filter(|task| task.enabled) {
            let source = task
                .source_wb
                .as_deref()
                .map(|path| path.to_string_lossy().into_owned())
                .unwrap_or_else(|| "（未配置）".to_owned());
            let target = task
                .target_wb
                .as_deref()
                .map(|path| path.to_string_lossy().into_owned())
                .unwrap_or_else(|| "（未配置）".to_owned());
            writeln!(
                output,
                "{} [INFO] 正在执行打印配置：工作表「打印配置」第 {} 行；源文件：{}；源工作表：{}；目标文件：{}；目标工作表：{}；打印区域：{}",
                pytools_core::engine_diag::now_prefix(),
                task.row_no,
                source,
                if task.source_ws.trim().is_empty() { "（未指定）" } else { task.source_ws.trim() },
                target,
                if task.target_ws.trim().is_empty() { "（未指定）" } else { task.target_ws.trim() },
                if task.source_ranges.trim().is_empty() { "（按配置默认范围）" } else { task.source_ranges.trim() }
            )?;
        }
    }
    output.flush()?;
    let run_print =
        || pytools_core::static_print_engines::run_configured_static_print(config_path, engine);
    // 打印仍在当前线程执行；动画使用 stderr，避免争用菜单持有的 stdout 锁。
    let result = if io::stdout().is_terminal() && io::stderr().is_terminal() {
        std::thread::scope(|scope| {
            let (stop, stopped) = std::sync::mpsc::channel::<()>();
            let indicator = scope.spawn(move || {
                let frames = ["  处理中.  ", "  处理中.. ", "  处理中..."];
                let mut frame = 0;
                let mut stderr = io::stderr();
                loop {
                    if write!(stderr, "\r{}", frames[frame])
                        .and_then(|_| stderr.flush())
                        .is_err()
                    {
                        break;
                    }
                    match stopped.recv_timeout(std::time::Duration::from_millis(400)) {
                        Err(std::sync::mpsc::RecvTimeoutError::Timeout) => {
                            frame = (frame + 1) % frames.len();
                        }
                        _ => break,
                    }
                }
                let _ = stderr.write_all(b"\r                    \r");
                let _ = stderr.flush();
            });
            let result = run_print();
            drop(stop);
            let _ = indicator.join();
            result
        })
    } else {
        // 非交互输出（管道/重定向）不打印占位行：最终摘要已包含结果与耗时。
        run_print()
    };
    if let (Some(tasks), Ok(run)) = (print_tasks_for_display.as_ref(), result.as_ref()) {
        for task in tasks.iter().filter(|task| task.enabled) {
            let task_event = run
                .events
                .iter()
                .find(|event| {
                    event.row_no == Some(task.row_no)
                        && matches!(event.outcome, DiagOutcome::Skipped | DiagOutcome::Failed)
                })
                .or_else(|| {
                    run.events.iter().find(|event| {
                        event.row_no == Some(task.row_no) && event.outcome == DiagOutcome::Ok
                    })
                });
            let target_saved = task
                .target_wb
                .as_ref()
                .is_some_and(|target| run.outputs.iter().any(|output_path| output_path == target));
            if let Some(event) = task_event
                .filter(|event| matches!(event.outcome, DiagOutcome::Skipped | DiagOutcome::Failed))
            {
                writeln!(
                    output,
                    "{} [{}] 打印配置执行结果：第 {} 行；状态：{}；原因：{}；保存：否",
                    pytools_core::engine_diag::now_prefix(),
                    if event.outcome == DiagOutcome::Failed {
                        "ERROR"
                    } else {
                        "WARN"
                    },
                    task.row_no,
                    event.outcome,
                    event.message
                )?;
            } else {
                writeln!(
                    output,
                    "{} [INFO] 打印配置执行结果：第 {} 行；重复行：不适用；修改：{}；保存：{}",
                    pytools_core::engine_diag::now_prefix(),
                    task.row_no,
                    if target_saved { "是" } else { "否" },
                    if target_saved { "是" } else { "否" }
                )?;
            }
        }
    }
    let outputs = result
        .as_ref()
        .ok()
        .map(|run| {
            run.outputs
                .iter()
                .map(|path| path_for_log_view(path, log_view))
                .collect()
        })
        .unwrap_or_default();
    let diagnostic_path = diagnostic_path_for_log_view(diagnostic.as_ref(), log_view);
    let mut summary = tracker.finish(
        result.is_err(),
        false,
        false,
        outputs,
        diagnostic_path.clone(),
    );
    summary.input_stats_available = false; // 配置任务数不等于源文件数。
    if let Ok(run) = &result {
        summary.overall_status = if run.task_failed > 0 && run.task_ok > 0 {
            TaskStatus::PartialSuccess
        } else if run.task_failed > 0 {
            TaskStatus::Failed
        } else if run.task_ok == 0 {
            TaskStatus::NoOp
        } else {
            TaskStatus::Success
        };
        summary
            .business_metrics
            .insert("成功配置项".to_owned(), serde_json::json!(run.task_ok));
        summary
            .business_metrics
            .insert("跳过配置项".to_owned(), serde_json::json!(run.task_skipped));
        summary
            .business_metrics
            .insert("失败配置项".to_owned(), serde_json::json!(run.task_failed));
        summary.output_stats = Some(pytools_core::run_logging::OutputStats {
            workbooks_saved: Some(run.outputs.len()),
            sheets_saved: None,
            saved: Some(run.task_failed == 0 && !run.outputs.is_empty()),
        });
        if let Some(source) = run
            .events
            .iter()
            .find(|event| event.outcome == DiagOutcome::Failed)
        {
            summary.error_summary = Some(pytools_core::run_logging::ErrorSummary {
                stage: source.stage.to_string(),
                code: match source.stage {
                    DiagStage::ReadConfig => "CONFIG_READ_FAILED",
                    DiagStage::ReadSource => "EXCEL_READ_FAILED",
                    DiagStage::CopySheet => "SHEET_COPY_FAILED",
                    DiagStage::ModifyWorkbook => "EXCEL_MODIFY_FAILED",
                    DiagStage::SaveOutput => "OUTPUT_WRITE_FAILED",
                }
                .to_owned(),
                message: format!("{}失败；请检查对应配置、工作簿和写入权限", source.stage),
                file_name: source
                    .file
                    .as_deref()
                    .map(|file| pytools_core::run_logging::portable_file_name(Path::new(file))),
                sheet_name: source.sheet.clone(),
            });
        }
        if let Some(writer) = diagnostic.as_mut() {
            for source in &run.events {
                let level = match source.outcome {
                    DiagOutcome::Failed => Level::Error,
                    DiagOutcome::Skipped => Level::Warn,
                    DiagOutcome::Started => Level::Info,
                    DiagOutcome::Ok => Level::Debug,
                };
                let mut event = tracker.event(
                    level,
                    "engine_operation",
                    &source.stage.to_string(),
                    &source.outcome.to_string(),
                    format!("{}：{}", source.stage, source.outcome),
                );
                event.timestamp = source.timestamp_rfc3339.clone();
                event.engine = Some(source.engine.clone());
                event.config_row = source.row_no;
                event.file_name = source
                    .file
                    .as_deref()
                    .map(|file| pytools_core::run_logging::portable_file_name(Path::new(file)));
                event.sheet_name = source.sheet.clone();
                event.elapsed_ms = source.elapsed_ms;
                if source.outcome == DiagOutcome::Failed {
                    event.error_code = Some(
                        match source.stage {
                            DiagStage::ReadConfig => "CONFIG_READ_FAILED",
                            DiagStage::ReadSource => "EXCEL_READ_FAILED",
                            DiagStage::CopySheet => "SHEET_COPY_FAILED",
                            DiagStage::ModifyWorkbook => "EXCEL_MODIFY_FAILED",
                            DiagStage::SaveOutput => "OUTPUT_WRITE_FAILED",
                        }
                        .to_owned(),
                    );
                    if log_view == "developer" {
                        event.details = Some(serde_json::json!({"diagnostic": source.message}));
                    }
                }
                if let Err(error) = writer.write(&event) {
                    if !log_failed {
                        writeln!(
                            output,
                            "{}",
                            pytools_core::run_logging::terminal_line(
                                Level::Warn,
                                &format!("诊断日志写入失败：{error}")
                            )
                        )?;
                        log_failed = true;
                    }
                    break;
                }
            }
        }
    } else {
        summary.error_summary = Some(pytools_core::run_logging::ErrorSummary {
            stage: "static_print".to_owned(),
            code: "EXCEL_OPERATION_FAILED".to_owned(),
            message: result
                .as_ref()
                .err()
                .map(|error| format!("打印任务中断：{error:#}"))
                .unwrap_or_else(|| "打印任务中断".to_owned()),
            file_name: None,
            sheet_name: None,
        });
    }
    if let (Err(error), Some(writer)) = (&result, diagnostic.as_mut()) {
        let mut failure = tracker.event(
            Level::Error,
            "operation_failed",
            "static_print",
            "failed",
            "打印任务中断；请检查配置、文件格式和读写权限",
        );
        failure.error_code = Some("EXCEL_OPERATION_FAILED".to_owned());
        if log_view == "developer" {
            failure.details = Some(serde_json::json!({"error_chain": format!("{error:#}")}));
        }
        if let Err(write_error) = writer.write(&failure) {
            if !log_failed {
                writeln!(
                    output,
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("诊断日志写入失败：{write_error}")
                    )
                )?;
                log_failed = true;
            }
        }
    }
    if let Some(writer) = diagnostic.as_mut() {
        let event = tracker.event(
            if result.is_err()
                || matches!(
                    summary.overall_status,
                    TaskStatus::Failed | TaskStatus::PartialSuccess
                )
            {
                Level::Error
            } else {
                Level::Info
            },
            "run_finished",
            "run",
            &format!("{:?}", summary.overall_status),
            "静态打印文件生成结束",
        );
        if let Err(error) = writer.write(&event) {
            if !log_failed {
                writeln!(
                    output,
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("诊断日志写入失败：{error}")
                    )
                )?;
            }
        }
    }
    if log_view != "off" {
        let mut report = serde_json::to_value(&summary)?;
        if let Ok(run) = &result {
            report["task_stats"] = serde_json::json!({"succeeded": run.task_ok, "failed": run.task_failed, "skipped": run.task_skipped});
        }
        let path = global
            .log_dir
            .join(format!("{feature_id}_{}.run.json", tracker.run_id));
        match write_json_report_create_new(&path, &report) {
            Ok(()) => writeln!(
                output,
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Info,
                    &format!(
                        "运行编号 {}；运行记录：{}",
                        tracker.run_id,
                        display_path_for_user(&path)
                    )
                )
            )?,
            Err(error) => writeln!(
                output,
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Warn,
                    &format!("运行记录写入失败：{error}；业务结果仍保留")
                )
            )?,
        }
    }
    if log_view == "user" {
        let level = match summary.overall_status {
            TaskStatus::Failed => Level::Error,
            TaskStatus::PartialSuccess | TaskStatus::Cancelled => Level::Warn,
            TaskStatus::Success if !summary.business_issues.is_empty() => Level::Warn,
            TaskStatus::Success | TaskStatus::NoOp => Level::Info,
        };
        for line in pytools_core::run_logging::format_terminal_summary(&summary, 8) {
            writeln!(
                output,
                "{}",
                pytools_core::run_logging::terminal_line(level, &line)
            )?;
        }
    }
    match result {
        Ok(run) => {
            // 跳过原因与错误详情：带各自发生时刻、引擎、阶段和原始错误。
            let detail_events = run
                .events
                .iter()
                .filter(|event| matches!(event.outcome, DiagOutcome::Skipped | DiagOutcome::Failed))
                .collect::<Vec<_>>();
            if !detail_events.is_empty() {
                for event in detail_events {
                    writeln!(output, "{}", event.user_line())?;
                }
            } else {
                if !run.issues.is_empty() {
                    writeln!(
                        output,
                        "{} [WARN] 有 {} 项打印问题，请检查配置及源文件。",
                        pytools_core::engine_diag::now_prefix(),
                        run.issues.len()
                    )?;
                }
            }
        }
        Err(_) => {}
    }
    output.flush()?;
    Ok(())
}

fn run_interactive_print_menu<R, W>(
    input: &mut R,
    output: &mut W,
    config_path: &Path,
    global: &GlobalConfig,
    log_view: &str,
    allow_native_dialog: bool,
    capabilities: Option<&Path>,
    platform: Option<&str>,
) -> Result<()>
where
    R: BufRead,
    W: Write,
{
    // 打印改为“按配置直接生成”：不再需要 y 确认，也不再提供仅出计划的入口。
    // 保留旧菜单编号；3-1 与 3-2 当前统一调用 openpyxl。
    let _ = (allow_native_dialog, capabilities, platform);
    loop {
        writeln!(output, "------ 报表打印 ------")?;
        writeln!(output, "  1  按打印配置生成：openpyxl（旧 3-1 入口）  【打印配置：源工作簿/源工作表/打印区域/目标工作簿】")?;
        writeln!(
            output,
            "  2  按打印配置生成：openpyxl（旧 3-2 入口）  【打印配置：源工作簿/源工作表/打印区域/目标工作簿】"
        )?;
        writeln!(output, "  3  打印配置预校验  【打印配置】")?;
        writeln!(output, "  0  返回")?;
        writeln!(output, "----------------------")?;
        let Some(menu_choice) = read_menu_choice(input, output, "子菜单选择> ")? else {
            return Ok(());
        };
        match menu_choice.as_str() {
            "1" => {
                run_configured_static_print_menu(
                    output,
                    config_path,
                    global,
                    "3-1",
                    pytools_core::static_print_engines::StaticPrintEngine::Openpyxl,
                    log_view,
                )?;
            }
            "2" => {
                run_configured_static_print_menu(
                    output,
                    config_path,
                    global,
                    "3-2",
                    pytools_core::static_print_engines::StaticPrintEngine::Openpyxl,
                    log_view,
                )?;
            }
            "3" => {
                run_print_menu_action(
                    output,
                    global,
                    log_view,
                    "PT-3-4",
                    "config_precheck",
                    "按配置预校验打印任务",
                    || {
                        Ok(serde_json::to_value(build_print_config_precheck_report(
                            config_cli_arguments(config_path),
                        )?)?)
                    },
                )?;
            }
            "0" => return Ok(()),
            _ => {
                writeln!(output, "仅接受 1、2、3 或 0")?;
            }
        }
    }
}

fn run_print_menu_action<W, F>(
    output: &mut W,
    global: &GlobalConfig,
    log_view: &str,
    feature_id: &str,
    action: &str,
    label: &str,
    run: F,
) -> Result<()>
where
    W: Write,
    F: FnOnce() -> Result<serde_json::Value>,
{
    let mut diagnostic = ConfiguredRunLog::new(global, feature_id, label, log_view);
    let started_at = Local::now();
    let started = Instant::now();
    match log_view {
        "user" => writeln!(output, "开始：{label}")?,
        "developer" => writeln!(output, "PT-3 开始：{feature_id} {action}")?,
        _ => {}
    }

    let report = match run() {
        Ok(report) => report,
        Err(error) => {
            diagnostic.set_error_summary(pytools_core::run_logging::ErrorSummary {
                stage: action.to_owned(),
                code: "PRINT_OPERATION_FAILED".to_owned(),
                message: format!("{label}失败：{error:#}"),
                file_name: None,
                sheet_name: None,
            });
            let summary = diagnostic.finish(true, None, &[], &[]);
            if log_view == "user" {
                write_unified_terminal_summary(output, &summary, &[])?;
            }
            let elapsed_ms = started.elapsed().as_millis();
            if log_view == "developer" {
                match write_print_menu_run_history(
                    global,
                    feature_id,
                    action,
                    "failed",
                    &started_at,
                    elapsed_ms,
                    &serde_json::json!({}),
                    &[],
                    &[],
                    false,
                    false,
                    Some("print_menu_action_failed"),
                ) {
                    Ok(path) => writeln!(output, "PT-3 失败：{feature_id}，耗时 {elapsed_ms_display_seconds:.3} 秒；运行记录：{}；错误：{error:#}", display_path_for_user(&path), elapsed_ms_display_seconds = elapsed_ms as f64 / 1000.0)?,
                    Err(log_error) => writeln!(output, "PT-3 失败：{feature_id}，耗时 {elapsed_ms_display_seconds:.3} 秒；运行记录写入失败：{log_error:#}；错误：{error:#}", elapsed_ms_display_seconds = elapsed_ms as f64 / 1000.0)?,
                }
            } else if log_view != "user" {
                writeln!(output, "{label}失败；错误：{error:#}")?;
            }
            return Ok(());
        }
    };

    let elapsed_ms = started.elapsed().as_millis();
    let metrics = print_menu_metrics(action, &report);
    let precheck_issues =
        action == "config_precheck" && metrics["failed_tasks"].as_u64().unwrap_or_default() > 0;
    let precheck_warnings =
        action == "config_precheck" && metrics["warning_tasks"].as_u64().unwrap_or_default() > 0;
    for (label, key) in match action {
        "config_precheck" => vec![
            ("通过配置数", "passed_tasks"),
            ("警告配置数", "warning_tasks"),
            ("失败配置数", "failed_tasks"),
        ],
        "config_plan" => vec![
            ("计划组数", "planned_groups"),
            ("执行组数", "executed_groups"),
            ("匹配工作表数", "matching_sheets"),
            ("标记区域数", "marked_ranges"),
            ("计划操作数", "planned_operations"),
            ("计划文件数", "plan_files"),
        ],
        _ => vec![
            ("匹配工作表数", "matching_sheets"),
            ("标记区域数", "marked_ranges"),
        ],
    } {
        if let Some(value) = metrics.get(key) {
            diagnostic.set_business_metric(label, value.clone());
        }
    }
    if precheck_issues {
        diagnostic.record_business_issue(
            pytools_core::run_logging::BusinessIssue::new("配置校验", "打印配置预校验发现未通过项")
                .count(metrics["failed_tasks"].as_u64().unwrap_or_default() as usize),
        );
    }
    if precheck_warnings {
        diagnostic.record_business_issue(
            pytools_core::run_logging::BusinessIssue::new("配置警告", "打印配置预校验发现警告项")
                .count(metrics["warning_tasks"].as_u64().unwrap_or_default() as usize),
        );
    }
    let status = "succeeded";
    let summary_outputs = report
        .as_array()
        .into_iter()
        .flatten()
        .filter_map(|item| {
            item["plan_file"]
                .as_str()
                .or_else(|| item["output_file"].as_str())
        })
        .map(PathBuf::from)
        .chain(report["plan_file"].as_str().map(PathBuf::from))
        .collect::<Vec<_>>();
    let summary = diagnostic.finish_as(
        false,
        None,
        &[],
        &summary_outputs,
        Some(pytools_core::run_logging::TaskStatus::Success),
    );
    let actual_engines = print_menu_actual_engines(&report);
    let writer_invoked = actual_engines
        .iter()
        .any(|engine| engine == "rust_xlsxwriter");
    let modifier_invoked = actual_engines
        .iter()
        .any(|engine| engine != "calamine" && engine != "rust_xlsxwriter");
    let mut actual_components = print_menu_components(action);
    if modifier_invoked {
        actual_components.push("excel-editor::ExcelModifier");
    }
    if writer_invoked {
        actual_components.push("rust_xlsxwriter::Workbook");
    }

    match log_view {
        "user" => {
            write_unified_terminal_summary(output, &summary, &summary_outputs)?;
            if let Some(issues) = report["issues"].as_array() {
                for issue in issues.iter().filter_map(serde_json::Value::as_str) {
                    writeln!(output, "  {issue}")?;
                }
            }
        }
        "developer" => {
            let outcome = "完成";
            match write_print_menu_run_history(
                global,
                feature_id,
                action,
                status,
                &started_at,
                elapsed_ms,
                &metrics,
                &actual_engines,
                &actual_components,
                modifier_invoked,
                writer_invoked,
                None,
            ) {
                Ok(path) => writeln!(output, "PT-3 {outcome}：{feature_id}，{}，实际引擎：{}，实际组件：{}，ExcelModifier {}，Excel writer {}，耗时 {elapsed_ms_display_seconds:.3} 秒；运行记录：{}", print_menu_summary(action, &metrics), actual_engines.join(", "), actual_components.join(", "), if modifier_invoked { "已调用" } else { "未调用" }, if writer_invoked { "已调用" } else { "未调用" }, display_path_for_user(&path), elapsed_ms_display_seconds = elapsed_ms as f64 / 1000.0)?,
                Err(error) => writeln!(output, "PT-3 {outcome}：{feature_id}，{}，实际引擎：{}，实际组件：{}，ExcelModifier {}，Excel writer {}，耗时 {elapsed_ms_display_seconds:.3} 秒；运行记录写入失败：{error:#}", print_menu_summary(action, &metrics), actual_engines.join(", "), actual_components.join(", "), if modifier_invoked { "已调用" } else { "未调用" }, if writer_invoked { "已调用" } else { "未调用" }, elapsed_ms_display_seconds = elapsed_ms as f64 / 1000.0)?,
            }
        }
        _ if log_view == "off" && precheck_issues => {
            writeln!(
                output,
                "预校验发现问题：{label}；{}。",
                print_menu_summary(action, &metrics)
            )?;
        }
        _ => {}
    }

    // 菜单默认不打印 JSON（user 视图保持纯可读输出）；
    // developer 视图保留机器可读 JSON 以兼容排障与自动化。
    if log_view == "developer" {
        writeln!(output, "{}", serde_json::to_string_pretty(&report)?)?;
    }
    Ok(())
}

fn print_menu_components(action: &str) -> Vec<&'static str> {
    match action {
        "comment-keep-format" | "comment-fast-copy" => vec![
            "pytools_core::print_ops::build_comment_print_plan",
            "excel_reader::Calamine",
        ],
        "config_plan" => vec![
            "pytools_core::config::load_print_tasks",
            "pytools_core::print_ops::build_config_print_plans",
            "excel_reader::Calamine",
        ],
        "config_precheck" => vec![
            "pytools_core::config::load_print_tasks",
            "pytools_core::print_ops::precheck_print_tasks",
            "excel_reader::Calamine",
        ],
        _ => Vec::new(),
    }
}

fn print_menu_actual_engines(report: &serde_json::Value) -> Vec<String> {
    fn collect(value: &serde_json::Value, engines: &mut Vec<String>) {
        match value {
            serde_json::Value::Array(items) => {
                for item in items {
                    collect(item, engines);
                }
            }
            serde_json::Value::Object(fields) => {
                if let Some(engine) = fields.get("engine").and_then(serde_json::Value::as_str) {
                    if !engines.iter().any(|existing| existing == engine) {
                        engines.push(engine.to_owned());
                    }
                }
                for (name, value) in fields {
                    if name != "engine" {
                        collect(value, engines);
                    }
                }
            }
            _ => {}
        }
    }

    let mut engines = vec!["calamine".to_owned()];
    collect(report, &mut engines);
    engines
}

fn print_menu_metrics(action: &str, report: &serde_json::Value) -> serde_json::Value {
    let count = |value: &serde_json::Value| value.as_u64().unwrap_or_default();
    match action {
        "comment-keep-format" | "comment-fast-copy" => serde_json::json!({
            "matching_sheets": count(&report["matching_sheets"]),
            "marked_ranges": count(&report["marked_ranges"]),
            "execution": if report["execution"].is_object() { "executed" } else { "plan_only" },
        }),
        "config_plan" => {
            let items = report.as_array().cloned().unwrap_or_default();
            let mut matching_sheets = 0u64;
            let mut marked_ranges = 0u64;
            let mut planned_operations = 0u64;
            let mut plan_files = 0u64;
            let mut executed_groups = 0u64;
            for item in &items {
                let details = &item["report"];
                matching_sheets += count(&details["matching_sheets"]);
                marked_ranges += count(&details["marked_ranges"]);
                planned_operations += details["plan"]["operations"]
                    .as_array()
                    .map_or(0, |operations| operations.len() as u64);
                plan_files += u64::from(item["plan_file"].is_string());
                executed_groups += u64::from(item["execution"].is_object());
            }
            serde_json::json!({
                "planned_groups": items.len(),
                "executed_groups": executed_groups,
                "execution": if executed_groups == 0 { "plan_only" } else if executed_groups == items.len() as u64 { "executed" } else { "partial" },
                "matching_sheets": matching_sheets,
                "marked_ranges": marked_ranges,
                "planned_operations": planned_operations,
                "plan_files": plan_files,
            })
        }
        "config_precheck" => serde_json::json!({
            "passed_tasks": count(&report["ok"]),
            "warning_tasks": count(&report["warnings"]),
            "failed_tasks": count(&report["failed"]),
        }),
        _ => serde_json::json!({}),
    }
}

fn print_menu_summary(action: &str, metrics: &serde_json::Value) -> String {
    let count = |name: &str| metrics[name].as_u64().unwrap_or_default();
    if action == "config_precheck" {
        return format!(
            "通过 {} 项、警告 {} 项、失败 {} 项",
            count("passed_tasks"),
            count("warning_tasks"),
            count("failed_tasks")
        );
    }
    if action == "config_plan" {
        let execution_summary = match metrics["execution"].as_str() {
            Some("executed") => format!("实际生成 {} 组静态打印工作簿", count("executed_groups")),
            Some("partial") => format!(
                "实际生成 {} 组静态打印工作簿，其余未执行",
                count("executed_groups")
            ),
            _ => "仅生成计划，未创建静态打印工作簿".to_owned(),
        };
        return format!(
            "{}；配置 {} 组，命中 {} 张表、{} 个打印区域，计划操作 {} 项",
            execution_summary,
            count("planned_groups"),
            count("matching_sheets"),
            count("marked_ranges"),
            count("planned_operations")
        );
    }
    format!(
        "命中 {} 张工作表、{} 个打印区域，{}",
        count("matching_sheets"),
        count("marked_ranges"),
        if metrics["execution"] == "executed" {
            "已生成工作簿"
        } else {
            "已生成计划"
        }
    )
}

fn write_print_menu_run_history(
    global: &GlobalConfig,
    feature_id: &str,
    action: &str,
    status: &str,
    started_at: &chrono::DateTime<Local>,
    elapsed_ms: u128,
    metrics: &serde_json::Value,
    actual_engines: &[String],
    actual_components: &[&str],
    excel_modifier_invoked: bool,
    excel_writer_invoked: bool,
    error_type: Option<&str>,
) -> Result<PathBuf> {
    fs::create_dir_all(&global.log_dir).with_context(|| {
        format!(
            "create print log directory {}",
            display_path_for_user(&global.log_dir)
        )
    })?;
    let timestamp = Local::now().format("%Y%m%d_%H%M%S_%f");
    let history = serde_json::json!({
        "schema_version": 1,
        "feature_id": feature_id,
        "action": action,
        "status": status,
        "started_at": started_at.to_rfc3339(),
        "elapsed_ms": elapsed_ms,
        "stages": [{
            "name": if action == "config_precheck" { "configuration_precheck" } else { "workbook_read_and_plan" },
            "elapsed_ms": elapsed_ms
        }],
        "metrics": metrics,
        "actual_engines": actual_engines,
        "actual_components": actual_components,
        "excel_modifier_invoked": excel_modifier_invoked,
        "excel_writer_invoked": excel_writer_invoked,
        "plan_only": action != "config_precheck" && !excel_modifier_invoked && !excel_writer_invoked,
        "error_type": error_type,
    });
    for suffix in 1..=1000 {
        let suffix = if suffix == 1 {
            String::new()
        } else {
            format!("-{suffix}")
        };
        let path = global
            .log_dir
            .join(format!("pytools-print-{action}-{timestamp}{suffix}.json"));
        match OpenOptions::new().write(true).create_new(true).open(&path) {
            Ok(mut file) => {
                serde_json::to_writer_pretty(&mut file, &history)?;
                file.write_all(b"\n")?;
                return Ok(path);
            }
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => {
                return Err(error).with_context(|| {
                    format!("create print run history {}", display_path_for_user(&path))
                })
            }
        }
    }
    bail!("could not allocate a unique print run history filename")
}

fn run_interactive_convert_menu<R, W>(
    input: &mut R,
    output: &mut W,
    config_path: &Path,
    global: &GlobalConfig,
    log_view: &str,
    allow_native_dialog: bool,
    _capabilities: Option<&Path>,
    _platform: Option<&str>,
) -> Result<()>
where
    R: BufRead,
    W: Write,
{
    loop {
        writeln!(output, "------ 文件处理 ------")?;
        writeln!(
            output,
            "  1  文件名机构标准化  【重命名配置：简称/全称/代码；键/值】"
        )?;
        writeln!(
            output,
            "  2  文件排序前缀管理（添加 / 清理）  【文件排序配置：是否启用/排序前缀/匹配关键词】"
        )?;
        writeln!(output, "  3  工作簿内容命名前缀（添加 / 清理）  【重命名配置：内容重命名启用/内容重命名取值规则】")?;
        writeln!(
            output,
            "  4  Office 格式转换（Excel 工作簿 → XLSX；支持 .xlsx 规范另存，DOC/WPS/WPT → DOCX）  【全局配置：excel目的格式/word目的格式】"
        )?;
        writeln!(
            output,
            "  5  按配置修改 Sheet 名  【重命名配置：原表名/新表名（默认 J:K）】"
        )?;
        writeln!(output, "  6  机构提交完整性检查  【提交检查配置】")?;
        writeln!(
            output,
            "  7  修复 Excel COM 无法打开的工作簿（生成修复副本；原件不变）  【自动补全角括号表名引号/修正冻结窗格矛盾】"
        )?;
        writeln!(output, "  0  返回")?;
        writeln!(output, "----------------------")?;
        let Some(menu_choice) = read_menu_choice(input, output, "子菜单选择> ")? else {
            return Ok(());
        };
        let operation_id = match menu_choice.as_str() {
            "1" => "files.institution_filename_normalize",
            "2" => {
                writeln!(output, "文件排序前缀操作：")?;
                writeln!(output, "  1  添加前缀")?;
                writeln!(output, "  2  清理前缀")?;
                let Some(action) = read_menu_choice(input, output, "选择操作> ")? else {
                    return Ok(());
                };
                match action.as_str() {
                    "1" => "files.sort_prefix.add",
                    "2" => "files.sort_prefix.remove",
                    "0" => continue,
                    _ => {
                        writeln!(output, "仅接受 1、2 或 0")?;
                        continue;
                    }
                }
            }
            "3" => {
                writeln!(output, "工作簿内容命名前缀操作：")?;
                writeln!(output, "  1  添加内容前缀")?;
                writeln!(output, "  2  清理内容前缀")?;
                writeln!(output, "  0  返回")?;
                let Some(action) = read_menu_choice(input, output, "选择操作> ")? else {
                    return Ok(());
                };
                match action.as_str() {
                    "1" => "files.name_from_workbook_cells",
                    "2" => "files.name_from_workbook_cells.remove",
                    "0" => continue,
                    _ => {
                        writeln!(output, "仅接受 1、2 或 0")?;
                        continue;
                    }
                }
            }
            "4" => {
                writeln!(output, "选择格式转换：")?;
                writeln!(output, "  1  Excel 格式转换或 .xlsx 规范另存为 XLSX")?;
                writeln!(output, "  2  Word 格式转换（DOC / WPS / WPT → DOCX）")?;
                let Some(format_choice) = read_menu_choice(input, output, "选择格式> ")? else {
                    return Ok(());
                };
                match format_choice.as_str() {
                    "1" => "files.office_conversion.excel",
                    "2" => "files.office_conversion.word",
                    "0" => continue,
                    _ => {
                        writeln!(output, "仅接受 1、2 或 0")?;
                        continue;
                    }
                }
            }
            "5" => "files.rename_sheets",
            "6" => "files.submission_check",
            "7" => "files.excel_open_repair",
            "0" => return Ok(()),
            _ => {
                writeln!(output, "仅接受 1、2、3、4、5、6、7 或 0")?;
                continue;
            }
        };
        let (title, filter) = match operation_id {
            "files.institution_filename_normalize" => {
                ("批量重命名文件", NativeFileFilter::AllFiles)
            }
            "files.sort_prefix.add" => ("添加文件排序前缀", NativeFileFilter::AllFiles),
            "files.sort_prefix.remove" => ("取消文件排序前缀", NativeFileFilter::AllFiles),
            "files.name_from_workbook_cells" => {
                ("按工作簿内容重命名文件", NativeFileFilter::ExcelWorkbooks)
            }
            "files.name_from_workbook_cells.remove" => {
                ("清理工作簿内容命名前缀", NativeFileFilter::ExcelWorkbooks)
            }
            "files.office_conversion.excel" => {
                ("批量转换 Excel 文件", NativeFileFilter::ExcelConversion)
            }
            "files.office_conversion.word" => {
                ("批量转换 Word 文件", NativeFileFilter::WordDocuments)
            }
            "files.rename_sheets" => ("批量修改 Sheet 名", NativeFileFilter::ExcelWorkbooks),
            "files.submission_check" => ("检查提交文件", NativeFileFilter::AllFiles),
            "files.excel_open_repair" => (
                "修复 Excel COM 无法打开的工作簿",
                NativeFileFilter::ExcelWorkbooks,
            ),
            _ => unreachable!("resolved stable file operation ID"),
        };
        let Some(paths) =
            choose_source_files_or_folders(title, filter, input, output, allow_native_dialog)?
        else {
            writeln!(output, "[已取消]")?;
            return Ok(());
        };
        if paths.is_empty() {
            writeln!(output, "[提示] 所选范围内没有符合格式的文件。")?;
            return Ok(());
        }

        let mut arguments = config_cli_arguments(config_path);
        match operation_id {
            "files.institution_filename_normalize" => {
                arguments.push("--log-dir".to_owned());
                arguments.push(global.log_dir.to_string_lossy().into_owned());
                arguments.push("--log-view".to_owned());
                arguments.push(log_view.to_owned());
                arguments.extend(
                    paths
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_batch_rename(arguments, true)?;
            }
            "files.office_conversion.excel" => {
                arguments.push("--log-dir".to_owned());
                arguments.push(global.log_dir.to_string_lossy().into_owned());
                arguments.push("--log-view".to_owned());
                arguments.push(log_view.to_owned());
                arguments.extend(
                    paths
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_convert_batch_menu(arguments, true, output)?;
            }
            "files.office_conversion.word" => {
                arguments.push("--log-dir".to_owned());
                arguments.push(global.log_dir.to_string_lossy().into_owned());
                arguments.push("--log-view".to_owned());
                arguments.push(log_view.to_owned());
                arguments.extend(
                    paths
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_convert_batch_menu(arguments, false, output)?;
            }
            "files.rename_sheets" => {
                let execute = {
                    writeln!(
                        output,
                        "[确认] 将直接修改所选源工作簿的工作表名称；先生成临时候选并校验，再安全替换原文件。"
                    )?;
                    writeln!(
                        output,
                        "[引擎] .xlsx 使用 openpyxl 改名并更新公式/命名区域引用；.xls 请先在 Windows 通过 4-4 转为 .xlsx。"
                    )?;
                    read_menu_choice(
                        input,
                        output,
                        "输入 y/yes/是 执行；其他输入取消，不生成计划文件> ",
                    )?
                    .is_some_and(|answer| {
                        matches!(
                            answer.trim().to_ascii_lowercase().as_str(),
                            "y" | "yes" | "是"
                        )
                    })
                };
                if execute {
                    let mut arguments = config_cli_arguments(config_path);
                    arguments.extend([
                        "--log-dir".to_owned(),
                        global.log_dir.to_string_lossy().into_owned(),
                        "--log-view".to_owned(),
                        log_view.to_owned(),
                    ]);
                    arguments.extend(
                        paths
                            .into_iter()
                            .map(|path| path.to_string_lossy().into_owned()),
                    );
                    match run_rename_sheets(arguments, true) {
                        Ok(()) => writeln!(
                            output,
                            "[完成] 已处理所选源工作簿，具体改名及跳过情况见上方汇总。"
                        )?,
                        Err(error) => writeln!(
                            output,
                            "[失败] 工作表改名未全部完成；已处理文件以汇总为准。错误：{error:#}"
                        )?,
                    }
                } else {
                    writeln!(output, "[已取消] 未执行改名，也未生成计划或输出文件。")?;
                }
            }
            "files.name_from_workbook_cells" | "files.name_from_workbook_cells.remove" => {
                arguments.push("--log-dir".to_owned());
                arguments.push(global.log_dir.to_string_lossy().into_owned());
                arguments.push("--log-view".to_owned());
                arguments.push(log_view.to_owned());
                arguments.extend(
                    paths
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                if operation_id == "files.name_from_workbook_cells.remove" {
                    run_content_prefix_unrename_cli(arguments, true)?;
                } else {
                    run_content_prefix_rename_cli(arguments, true)?;
                }
            }
            "files.submission_check" => {
                arguments.push("--output-dir".to_owned());
                arguments.push(global.output_dir.to_string_lossy().into_owned());
                arguments.push("--log-dir".to_owned());
                arguments.push(global.log_dir.to_string_lossy().into_owned());
                arguments.push("--log-view".to_owned());
                arguments.push(log_view.to_owned());
                arguments.extend(
                    paths
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_submission_check_cli(arguments, true)?;
            }
            "files.excel_open_repair" => {
                writeln!(
                    output,
                    "[确认] 将检测所选工作簿并修复 Excel 无法打开的两类问题：含全角括号表名的裸引用、冻结窗格坐标矛盾。"
                )?;
                writeln!(
                    output,
                    "[说明] 修复只写入源文件旁的“<名称>-Excel修复.<扩展名>”副本；原始文件保持不变，无需修复的文件不会产出副本。"
                )?;
                if !read_menu_choice(input, output, "输入 y/yes/是 执行；其他输入取消> ")?
                    .is_some_and(|answer| {
                        matches!(
                            answer.trim().to_ascii_lowercase().as_str(),
                            "y" | "yes" | "是"
                        )
                    })
                {
                    writeln!(output, "[已取消] 未执行修复，也未生成任何副本。")?;
                    return Ok(());
                }
                arguments.push("--log-dir".to_owned());
                arguments.push(global.log_dir.to_string_lossy().into_owned());
                arguments.push("--log-view".to_owned());
                arguments.push(log_view.to_owned());
                arguments.extend(
                    paths
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_excel_open_repair_cli(arguments, true)?;
            }
            "files.sort_prefix.add" | "files.sort_prefix.remove" => {
                arguments.push("--output-dir".to_owned());
                arguments.push(global.output_dir.to_string_lossy().into_owned());
                arguments.push("--log-dir".to_owned());
                arguments.push(global.log_dir.to_string_lossy().into_owned());
                arguments.push("--log-view".to_owned());
                arguments.push(log_view.to_owned());
                arguments.extend(
                    paths
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                if operation_id == "files.sort_prefix.add" {
                    run_file_sort_prefix(arguments, true)?;
                } else {
                    run_remove_file_sort_prefix(arguments, true)?;
                }
            }
            _ => unreachable!("stable file operation ID was validated"),
        }
        return Ok(());
    }
}

fn run_interactive_summary_menu<R, W>(
    input: &mut R,
    output: &mut W,
    global: &GlobalConfig,
    log_view: &str,
    allow_native_dialog: bool,
) -> Result<()>
where
    R: BufRead,
    W: Write,
{
    loop {
        writeln!(output, "------ 汇总工具 ------")?;
        writeln!(output, "  1  按使用区域汇总（全量）")?;
        writeln!(output, "  2  按批注汇总（模板+源文件）")?;
        writeln!(output, "  0  返回")?;
        writeln!(output, "----------------------")?;
        let Some(choice) = read_menu_choice(input, output, "子菜单选择> ")? else {
            return Ok(());
        };
        match choice.as_str() {
            "1" => {
                let Some(sources) = choose_workbooks(
                    "选择按使用区域汇总的源文件（可多选）",
                    true,
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let mut arguments = summary_cli_arguments(global, log_view);
                arguments.extend(
                    sources
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_summary_used_range_cli(arguments, true)?;
                return Ok(());
            }
            "2" => {
                let Some(template) = choose_workbooks(
                    "选择批注汇总模板文件",
                    false,
                    input,
                    output,
                    allow_native_dialog,
                )?
                .and_then(|paths| paths.into_iter().next()) else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let Some(sources) = choose_workbooks(
                    "选择按批注汇总的源文件（可多选）",
                    true,
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let mut arguments = summary_cli_arguments(global, log_view);
                arguments.push("--template".to_owned());
                arguments.push(template.to_string_lossy().into_owned());
                arguments.extend(
                    sources
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_summary_comments_cli(arguments, true)?;
                return Ok(());
            }
            "0" => return Ok(()),
            _ => writeln!(output, "输入无效，请重试。")?,
        }
    }
}

/// 菜单 5 的就地修改统一使用受保护的 openpyxl 执行路径。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum SpecialReportEngine {
    Openpyxl,
}

impl SpecialReportEngine {
    fn label(self) -> &'static str {
        match self {
            Self::Openpyxl => "openpyxl（本机 Python，保留宏）",
        }
    }

    /// 该引擎可就地修改的扩展名。
    fn accepts(self, path: &Path) -> bool {
        let extension = path
            .extension()
            .and_then(|ext| ext.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase();
        match self {
            Self::Openpyxl => matches!(extension.as_str(), "xlsx" | "xlsm" | "xltx" | "xltm"),
        }
    }
}

fn run_interactive_batch_menu<R, W>(
    input: &mut R,
    output: &mut W,
    config_path: &Path,
    global: &GlobalConfig,
    log_view: &str,
    allow_native_dialog: bool,
) -> Result<()>
where
    R: BufRead,
    W: Write,
{
    loop {
        writeln!(output, "------ 专用报表处理 ------")?;
        writeln!(output, "  1  拆分村镇银行数据（直接修改所选源工作簿）")?;
        writeln!(output, "  2  机构名称规范化（直接修改所选源工作簿）  【机构映射表：原始机构名称/映射后机构名称】")?;
        writeln!(
            output,
            "  3  删除分机构表中的外资行（生成“金融局”副本，原文件不变）  【机构映射表：是否为外资行】"
        )?;
        writeln!(output, "  4  修改外汇工作表页眉单位（直接修改原工作簿）")?;
        writeln!(output, "  5  地区总分校验（读取并生成核对结果）")?;
        writeln!(
            output,
            "  6  工作表数据提取（普通无批注版）  【工作表提取：关键字/提取区域/输出设置】"
        )?;
        writeln!(output, "  7  涉农贷款比上月修正（直接修改本期源工作簿）")?;
        writeln!(
            output,
            "  8  季度涉农贷款修正及分地区核对（直接修改本期源工作簿）"
        )?;
        writeln!(output, "  0  返回")?;
        writeln!(output, "----------------------------")?;
        let Some(choice) = read_menu_choice(input, output, "子菜单选择> ")? else {
            return Ok(());
        };
        let feature_id = match choice.as_str() {
            "1" => "PT-7-1",
            "2" => "PT-7-2",
            "3" => "PT-7-3",
            "4" => "PT-7-4",
            "5" => "PT-7-5",
            "6" => "PT-7-6",
            "7" => "PT-7-7",
            "8" => "PT-7-8",
            "0" => return Ok(()),
            _ => {
                writeln!(output, "仅接受 1、2、3、4、5、6、7、8 或 0")?;
                continue;
            }
        };
        // 专用报表菜单的既有工作簿修改统一使用项目内置 openpyxl worker，不路由到 umya。
        let engine: Option<SpecialReportEngine> = if matches!(
            feature_id,
            "PT-7-1" | "PT-7-2" | "PT-7-3" | "PT-7-4" | "PT-7-7" | "PT-7-8"
        ) {
            Some(SpecialReportEngine::Openpyxl)
        } else {
            None
        };
        match feature_id {
            "PT-7-1" => {
                writeln!(output, "→ 第一次文件选择：选择 本期商业银行共享一批")?;
                let Some(current) = choose_single_excel_workbook(
                    "选择 本期商业银行共享一批",
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消] 未选择本期工作簿。")?;
                    return Ok(());
                };
                writeln!(output, "→ 第二次文件选择：选择 去年同期商业银行共享一批")?;
                let Some(previous) = choose_single_excel_workbook(
                    "选择 去年同期商业银行共享一批",
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消] 未选择去年同期商业银行共享一批。")?;
                    return Ok(());
                };
                writeln!(
                    output,
                    "→ 第三次文件选择：选择 本外币、人民币、外币、外发用报表（修改分机构表）"
                )?;
                let Some(selected) = choose_source_files_or_folders(
                    "选择 本外币、人民币、外币、外发用报表（修改分机构表）（可多选）",
                    NativeFileFilter::ExcelWorkbooks,
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消] 未选择目标工作簿。")?;
                    return Ok(());
                };
                let mut targets = Vec::new();
                for path in selected {
                    if path.components().any(|component| {
                        component.as_os_str().to_string_lossy() == "PyTools_专用报表处理结果"
                    }) {
                        writeln!(
                            output,
                            "[跳过] {}：这是专用报表结果副本目录，不重复处理。",
                            display_path_for_user(&path)
                        )?;
                        continue;
                    }
                    let accepted = path.extension().is_some_and(|extension| {
                        matches!(
                            extension.to_string_lossy().to_ascii_lowercase().as_str(),
                            "xls" | "xlsx"
                        )
                    });
                    if accepted {
                        targets.push(path);
                    } else {
                        writeln!(
                            output,
                            "[跳过] {}：PT-7-1 当前只接受 .xls/.xlsx；其他格式不处理。",
                            display_path_for_user(&path)
                        )?;
                    }
                }
                if targets.is_empty() {
                    writeln!(output, "[提示] 没有可处理的 .xls/.xlsx 目标工作簿。")?;
                    return Ok(());
                }
                let current_ext = current
                    .extension()
                    .and_then(|extension| extension.to_str())
                    .unwrap_or_default()
                    .to_ascii_lowercase();
                let previous_ext = previous
                    .extension()
                    .and_then(|extension| extension.to_str())
                    .unwrap_or_default()
                    .to_ascii_lowercase();
                if current_ext != previous_ext {
                    writeln!(
                        output,
                        "[中止] 本期和去年同期工作簿格式必须相同，均选 .xls 或均选 .xlsx。"
                    )?;
                    return Ok(());
                }
                let mut xls_targets = Vec::new();
                let mut xlsx_targets = Vec::new();
                for target in targets {
                    match target.extension().and_then(|extension| extension.to_str()) {
                        Some(extension) if extension.eq_ignore_ascii_case("xls") => {
                            xls_targets.push(target)
                        }
                        Some(extension) if extension.eq_ignore_ascii_case("xlsx") => {
                            xlsx_targets.push(target)
                        }
                        _ => {}
                    }
                }
                writeln!(
                    output,
                    "[就地模式] .xls 由 Excel/WPS COM 修改安全候选后替换源文件；.xlsx 保持 Openpyxl 路径。两者都不生成额外结果副本。"
                )?;
                for (extension, selected_targets, engine_name) in [
                    ("xls", xls_targets, "com"),
                    ("xlsx", xlsx_targets, "openpyxl"),
                ] {
                    if selected_targets.is_empty() {
                        continue;
                    }
                    if extension != current_ext {
                        writeln!(
                            output,
                            "[跳过] 选择的目标 .{extension} 与本期/去年同期 .{current_ext} 不一致；为避免跨格式计算，本次不处理该组。"
                        )?;
                        continue;
                    }
                    let mut arguments = vec![
                        "--current".to_owned(),
                        current.to_string_lossy().into_owned(),
                        "--previous".to_owned(),
                        previous.to_string_lossy().into_owned(),
                        "--output-dir".to_owned(),
                        global.output_dir.to_string_lossy().into_owned(),
                        "--log-dir".to_owned(),
                        global.log_dir.to_string_lossy().into_owned(),
                        "--log-view".to_owned(),
                        log_view.to_owned(),
                        "--in-place".to_owned(),
                        engine_name.to_owned(),
                    ];
                    arguments.extend(
                        selected_targets
                            .into_iter()
                            .map(|path| path.to_string_lossy().into_owned()),
                    );
                    run_village_split_batch(arguments, true)?;
                }
                return Ok(());
            }
            "PT-7-2" => {
                let Some(selected) = choose_source_files_or_folders(
                    "选择需要规范化机构名称的目标工作簿",
                    NativeFileFilter::ExcelWorkbooks,
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let mut targets = Vec::new();
                for path in selected {
                    if path.components().any(|component| {
                        component.as_os_str().to_string_lossy()
                            == pytools_core::institution_normalize::INSTITUTION_NORMALIZE_OUTPUT_DIRECTORY
                    }) {
                        writeln!(
                            output,
                            "[跳过] {}：这是机构名称规范化结果目录中的副本，不重复处理。",
                            display_path_for_user(&path)
                        )?;
                        continue;
                    }
                    let accepted = if path
                        .extension()
                        .is_some_and(|extension| extension.eq_ignore_ascii_case("xls"))
                    {
                        cfg!(windows)
                    } else {
                        match engine {
                            Some(SpecialReportEngine::Openpyxl) => {
                                path.extension().is_some_and(|e| {
                                    let e = e.to_string_lossy().to_ascii_lowercase();
                                    matches!(e.as_str(), "xlsx" | "xlsm" | "xltx" | "xltm")
                                })
                            }
                            _ => path.extension().is_some_and(|e| {
                                let e = e.to_string_lossy().to_ascii_lowercase();
                                matches!(e.as_str(), "xlsx" | "xltx")
                            }),
                        }
                    };
                    if accepted {
                        targets.push(path);
                    } else if path
                        .extension()
                        .is_some_and(|e| e.to_string_lossy().eq_ignore_ascii_case("xls"))
                    {
                        writeln!(
                            output,
                            "[跳过] {}：旧版 .xls 机构规范化需要 Windows Excel/WPS COM。",
                            display_path_for_user(&path)
                        )?;
                    } else {
                        writeln!(
                            output,
                            "[跳过] {}：当前 openpyxl 引擎不支持该格式。",
                            display_path_for_user(&path)
                        )?;
                    }
                }
                if targets.is_empty() {
                    writeln!(output, "[提示] 没有可处理的 Excel 工作簿。")?;
                    return Ok(());
                }
                writeln!(output, "[就地模式] .xls 使用 Excel/WPS COM；.xlsx 使用 Openpyxl，直接改写源工作簿中的机构名称。")?;

                let mut arguments = vec![
                    "--config".to_owned(),
                    config_path.to_string_lossy().into_owned(),
                    "--view".to_owned(),
                    log_view.to_owned(),
                    "--in-place".to_owned(),
                    "auto".to_owned(),
                ];
                if log_view != "off" {
                    let nonce = SystemTime::now()
                        .duration_since(UNIX_EPOCH)
                        .map(|duration| duration.as_nanos())
                        .unwrap_or_default();
                    let report_path = global.log_dir.join(format!(
                        "PT-7-2_{}_{}_{}.json",
                        Local::now().format("%Y%m%dT%H%M%S"),
                        std::process::id(),
                        nonce
                    ));
                    arguments.extend([
                        "--report-output".to_owned(),
                        report_path.to_string_lossy().into_owned(),
                    ]);
                }
                arguments.push("--apply".to_owned());
                arguments.extend(
                    targets
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_institution_normalize_batch(arguments, true)?;
                return Ok(());
            }
            "PT-7-3" => {
                let foreign = pytools_core::config::load_foreign_banks(config_path)?;
                if foreign.is_empty() {
                    writeln!(
                        output,
                        "[已中止] 外资行配置为空，请在机构映射表 C 列将外资行标记为 1。"
                    )?;
                    return Ok(());
                }
                let Some(selected) = choose_source_files_or_folders(
                    "选择要生成外发用副本的工作簿（删除分机构表外资行）",
                    NativeFileFilter::XlsxOnly,
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let mut targets = Vec::new();
                for path in selected {
                    if path.components().any(|component| {
                        component.as_os_str().to_string_lossy() == "PyTools_专用报表处理结果"
                    }) {
                        writeln!(
                            output,
                            "[跳过] {}：这是专用报表结果副本目录，不重复处理。",
                            display_path_for_user(&path)
                        )?;
                        continue;
                    }
                    if path
                        .extension()
                        .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
                    {
                        targets.push(path);
                    } else {
                        writeln!(
                            output,
                            "[跳过] {}：PT-7-3 当前只接受 .xlsx；未验收格式请先显式转换。",
                            display_path_for_user(&path)
                        )?;
                    }
                }
                if targets.is_empty() {
                    writeln!(output, "[提示] 没有可处理的 .xlsx 工作簿。")?;
                    return Ok(());
                }
                writeln!(output, "[副本模式] 使用 openpyxl 删除分机构表中的外资行，生成“(金融局)”副本；原工作簿保持不变。")?;

                let mut arguments = vec![
                    "--config".to_owned(),
                    config_path.to_string_lossy().into_owned(),
                    "--output-dir".to_owned(),
                    global.output_dir.to_string_lossy().into_owned(),
                    "--view".to_owned(),
                    log_view.to_owned(),
                    "--log-dir".to_owned(),
                    global.log_dir.to_string_lossy().into_owned(),
                ];
                if log_view != "off" {
                    let nonce = SystemTime::now()
                        .duration_since(UNIX_EPOCH)
                        .map(|duration| duration.as_nanos())
                        .unwrap_or_default();
                    let report_path = global.log_dir.join(format!(
                        "PT-7-3_{}_{}_{}.json",
                        Local::now().format("%Y%m%dT%H%M%S"),
                        std::process::id(),
                        nonce
                    ));
                    arguments.extend([
                        "--report-output".to_owned(),
                        report_path.to_string_lossy().into_owned(),
                    ]);
                }
                arguments.push("--create-copy".to_owned());
                arguments.extend(
                    targets
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_remove_foreign_bank_batch(arguments, true)?;
                return Ok(());
            }
            "PT-7-4" => {
                let Some(selected) = choose_workbooks(
                    "选择要修改外汇页眉的 Excel 文件（可多选）",
                    true,
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消] 未选择外汇工作簿。")?;
                    return Ok(());
                };
                let mut targets = Vec::new();
                for path in selected {
                    if path
                        .extension()
                        .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
                    {
                        targets.push(path);
                    } else if path
                        .extension()
                        .is_some_and(|e| e.to_string_lossy().eq_ignore_ascii_case("xls"))
                    {
                        writeln!(
                            output,
                            "[跳过] {}：PT-7-4 当前仅验收 .xlsx；.xls 请先显式转换，含 VBA 的 .xlsm 暂不修改。",
                            display_path_for_user(&path)
                        )?;
                    } else {
                        writeln!(
                            output,
                            "[跳过] {}：PT-7-4 当前仅验收 .xlsx；.xls 请先显式转换，含 VBA 的 .xlsm 暂不修改。",
                            display_path_for_user(&path)
                        )?;
                    }
                }
                if targets.is_empty() {
                    writeln!(output, "[提示] 没有可处理的 .xlsx 工作簿。")?;
                    return Ok(());
                }
                writeln!(
                    output,
                    "[就地模式] 优先使用 OOXML 定点补丁修改外汇页眉；仅遇到 OOXML 不支持的结构时回退到受保护的 Openpyxl，候选校验后安全替换原文件。"
                )?;
                let mut arguments = vec![
                    "--output-dir".to_owned(),
                    global.output_dir.to_string_lossy().into_owned(),
                    "--log-dir".to_owned(),
                    global.log_dir.to_string_lossy().into_owned(),
                    "--view".to_owned(),
                    log_view.to_owned(),
                ];
                arguments.extend(["--in-place".to_owned(), "openpyxl".to_owned()]);
                arguments.extend(
                    targets
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_fx_header_batch(arguments, true)?;
                return Ok(());
            }
            "PT-7-5" => {
                let Some(selected) = choose_source_files_or_folders(
                    "选择要做地区总分校验的 Excel 文件",
                    NativeFileFilter::XlsxOrXlsm,
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let mut sources = Vec::new();
                for path in selected {
                    if path.extension().is_some_and(|extension| {
                        extension.eq_ignore_ascii_case("xlsx")
                            || extension.eq_ignore_ascii_case("xlsm")
                    }) {
                        sources.push(path);
                    } else {
                        writeln!(
                            output,
                            "[跳过] {}：当前 Rust 核对入口只支持 .xlsx/.xlsm；请先显式转换为 .xlsx",
                            display_path_for_user(&path)
                        )?;
                    }
                }
                if sources.is_empty() {
                    writeln!(output, "[提示] 没有可处理的 .xlsx/.xlsm 文件。")?;
                    return Ok(());
                }
                let mut arguments = config_cli_arguments(config_path);
                arguments.extend([
                    "--output-dir".to_owned(),
                    global.output_dir.to_string_lossy().into_owned(),
                    "--log-dir".to_owned(),
                    global.log_dir.to_string_lossy().into_owned(),
                    "--log-view".to_owned(),
                    log_view.to_owned(),
                ]);
                arguments.extend(
                    sources
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_region_sum(arguments, true)?;
                return Ok(());
            }
            "PT-7-6" => {
                let tasks = load_extract_tasks(config_path).with_context(|| {
                    format!(
                        "load worksheet-extraction tasks from {}",
                        display_path_for_user(&config_path)
                    )
                })?;
                if tasks.is_empty() {
                    writeln!(output, "[已中止] '工作表提取'配置无启用任务。")?;
                    return Ok(());
                }
                writeln!(
                    output,
                    "[就绪] 已加载 {} 条启用的工作表提取任务。",
                    tasks.len()
                )?;
                let Some(append_choice) = read_menu_choice(
                    input,
                    output,
                    "若任务配置的输出工作簿已存在，是否追加新工作表？输入 y/yes/是 确认，其他输入保持新建输出模式> ",
                )? else {
                    writeln!(output, "[已取消] 未选择工作表提取输出模式。")?;
                    return Ok(());
                };
                let append_existing = matches!(
                    append_choice.trim().to_ascii_lowercase().as_str(),
                    "y" | "yes" | "是"
                );
                let Some(selected) = choose_workbooks(
                    "选择要提取工作表数据的源文件（可多选）",
                    true,
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消] 未选择工作表提取源文件。")?;
                    return Ok(());
                };
                let mut sources = Vec::new();
                for path in selected {
                    if path.extension().is_some_and(|extension| {
                        extension.eq_ignore_ascii_case("xlsx")
                            || extension.eq_ignore_ascii_case("xlsm")
                    }) {
                        sources.push(path);
                    } else {
                        writeln!(
                            output,
                            "[跳过] {}：PT-7-6 区域提取仅接受 .xlsx/.xlsm；整表样式导入仍仅接受 .xlsx。",
                            display_path_for_user(&path)
                        )?;
                    }
                }
                if sources.is_empty() {
                    writeln!(output, "[提示] 没有可处理的 .xlsx/.xlsm 工作簿。")?;
                    return Ok(());
                }
                let mut arguments = config_cli_arguments(config_path);
                arguments.extend([
                    "--output-dir".to_owned(),
                    global.output_dir.to_string_lossy().into_owned(),
                    "--log-dir".to_owned(),
                    global.log_dir.to_string_lossy().into_owned(),
                    "--log-view".to_owned(),
                    log_view.to_owned(),
                ]);
                if append_existing {
                    arguments.push("--append-existing".to_owned());
                    writeln!(
                        output,
                        "[追加模式] 已明确选择向配置指定的已有输出工作簿追加工作表。"
                    )?;
                }
                arguments.extend(
                    sources
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_extract_sheets(arguments, true)?;
                return Ok(());
            }
            "PT-7-7" | "PT-7-8" => {
                let quarterly = feature_id == "PT-7-8";
                let operation = if quarterly {
                    "季度涉农贷款比上月修正及分地区核对"
                } else {
                    "涉农贷款比上月修正"
                };
                let current_title = if quarterly {
                    "选择本期季度涉农贷款工作簿"
                } else {
                    "选择本期涉农贷款工作簿"
                };
                let previous_title = if quarterly {
                    "选择上期季度涉农贷款工作簿"
                } else {
                    "选择上期涉农贷款工作簿"
                };
                let Some(current) = choose_single_rural_loan_workbook(
                    current_title,
                    feature_id,
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消] 未选择本期工作簿。")?;
                    return Ok(());
                };
                let Some(previous) = choose_single_rural_loan_workbook(
                    previous_title,
                    feature_id,
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消] 未选择上期工作簿。")?;
                    return Ok(());
                };
                writeln!(
                    output,
                    "[就地模式] {operation}：.xls 使用 Excel/WPS COM 修改隔离候选，校验快照后备份并替换源文件；.xlsx 使用 Openpyxl 候选路径。"
                )?;
                run_rural_loan_menu(
                    &current, &previous, global, log_view, None, None, quarterly, engine,
                )?;
                return Ok(());
            }
            _ => unreachable!("resolved specialized report feature ID"),
        }
    }
}

fn summary_cli_arguments(global: &GlobalConfig, log_view: &str) -> Vec<String> {
    vec![
        "--output-dir".to_owned(),
        global.output_dir.to_string_lossy().into_owned(),
        "--log-dir".to_owned(),
        global.log_dir.to_string_lossy().into_owned(),
        "--log-view".to_owned(),
        log_view.to_owned(),
    ]
}

fn read_menu_choice<R, W>(input: &mut R, output: &mut W, prompt: &str) -> Result<Option<String>>
where
    R: BufRead,
    W: Write,
{
    write!(output, "{prompt}")?;
    output.flush()?;
    let mut line = String::new();
    if input.read_line(&mut line)? == 0 {
        return Ok(None);
    }
    Ok(Some(line.trim().to_owned()))
}

fn choose_workbooks<R, W>(
    title: &str,
    allow_multiple: bool,
    input: &mut R,
    output: &mut W,
    allow_native_dialog: bool,
) -> Result<Option<Vec<PathBuf>>>
where
    R: BufRead,
    W: Write,
{
    choose_input_paths(
        title,
        allow_multiple,
        NativeFileFilter::ExcelWorkbooks,
        input,
        output,
        allow_native_dialog,
    )
}

fn choose_single_excel_workbook<R, W>(
    title: &str,
    input: &mut R,
    output: &mut W,
    allow_native_dialog: bool,
) -> Result<Option<PathBuf>>
where
    R: BufRead,
    W: Write,
{
    let Some(mut selected) = choose_input_paths(
        title,
        false,
        NativeFileFilter::ExcelWorkbooks,
        input,
        output,
        allow_native_dialog,
    )?
    else {
        return Ok(None);
    };
    if selected.len() != 1 {
        writeln!(
            output,
            "[提示] 此步骤需要且只需要选择一个 .xls 或 .xlsx 工作簿。"
        )?;
        return Ok(None);
    }
    let path = selected.remove(0);
    let supported = path.extension().is_some_and(|extension| {
        extension.eq_ignore_ascii_case("xls") || extension.eq_ignore_ascii_case("xlsx")
    });
    if !supported {
        writeln!(
            output,
            "[跳过] {}：PT-7-1 当前只接受 .xls/.xlsx；未验收格式不执行修改。",
            display_path_for_user(&path)
        )?;
        return Ok(None);
    }
    Ok(Some(path))
}

fn choose_single_rural_loan_workbook<R, W>(
    title: &str,
    feature_id: &str,
    input: &mut R,
    output: &mut W,
    allow_native_dialog: bool,
) -> Result<Option<PathBuf>>
where
    R: BufRead,
    W: Write,
{
    let file_filter = if feature_id == "PT-7-7" {
        NativeFileFilter::ExcelWorkbooks
    } else {
        NativeFileFilter::XlsxOnly
    };
    let Some(mut selected) = choose_input_paths(
        title,
        false,
        file_filter,
        input,
        output,
        allow_native_dialog,
    )?
    else {
        return Ok(None);
    };
    if selected.len() != 1 {
        writeln!(
            output,
            "[提示] {feature_id} 此步骤需要且只需要一个工作簿。5-7 接受 .xls/.xlsx；5-8 接受 .xlsx。"
        )?;
        return Ok(None);
    }
    let path = selected.remove(0);
    let extension = path
        .extension()
        .and_then(|extension| extension.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    let supported = if feature_id == "PT-7-7" {
        matches!(extension.as_str(), "xls" | "xlsx")
    } else {
        extension == "xlsx"
    };
    if !supported {
        writeln!(
            output,
            "[跳过] {}：{feature_id} 当前格式不受支持；5-7 接受 .xls/.xlsx，5-8 暂只接受 .xlsx。",
            display_path_for_user(&path)
        )?;
        return Ok(None);
    }
    Ok(Some(path))
}

#[derive(Clone, Copy)]
enum NativeFileFilter {
    ExcelWorkbooks,
    ExcelConversion,
    XlsxOrXlsm,
    XlsxOnly,
    WordDocuments,
    AllFiles,
}

impl NativeFileFilter {
    fn dialog_filter(self) -> &'static str {
        match self {
            Self::ExcelWorkbooks => {
                "Excel files (*.xlsx;*.xlsm;*.xls)\0*.xlsx;*.xlsm;*.xls\0All files (*.*)\0*.*\0\0"
            }
            Self::ExcelConversion => {
                "Excel files (*.csv;*.et;*.xls;*.xlsx;*.xlsm;*.xlsb)\0*.csv;*.et;*.xls;*.xlsx;*.xlsm;*.xlsb\0All files (*.*)\0*.*\0\0"
            }
            Self::XlsxOrXlsm => {
                "Excel files (*.xlsx;*.xlsm)\0*.xlsx;*.xlsm\0All files (*.*)\0*.*\0\0"
            }
            Self::XlsxOnly => "Excel files (*.xlsx)\0*.xlsx\0All files (*.*)\0*.*\0\0",
            Self::WordDocuments => {
                "Word files (*.doc;*.docx)\0*.doc;*.docx\0All files (*.*)\0*.*\0\0"
            }
            Self::AllFiles => "All files (*.*)\0*.*\0\0",
        }
    }

    fn accepts_folder_file(self, path: &Path) -> bool {
        let extension = path
            .extension()
            .and_then(|value| value.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase();
        match self {
            Self::ExcelWorkbooks => matches!(extension.as_str(), "xlsx" | "xlsm" | "xls"),
            Self::ExcelConversion => {
                matches!(
                    extension.as_str(),
                    "csv" | "et" | "xls" | "xlsx" | "xlsm" | "xlsb"
                )
            }
            Self::XlsxOrXlsm => matches!(extension.as_str(), "xlsx" | "xlsm"),
            Self::XlsxOnly => extension == "xlsx",
            Self::WordDocuments => matches!(extension.as_str(), "doc" | "docx"),
            Self::AllFiles => true,
        }
    }
}

fn choose_input_paths<R, W>(
    title: &str,
    allow_multiple: bool,
    filter: NativeFileFilter,
    input: &mut R,
    output: &mut W,
    allow_native_dialog: bool,
) -> Result<Option<Vec<PathBuf>>>
where
    R: BufRead,
    W: Write,
{
    choose_input_paths_impl(
        title,
        allow_multiple,
        filter,
        input,
        output,
        allow_native_dialog,
        false,
    )
}

fn choose_input_paths_impl<R, W>(
    title: &str,
    allow_multiple: bool,
    filter: NativeFileFilter,
    input: &mut R,
    output: &mut W,
    allow_native_dialog: bool,
    count_only: bool,
) -> Result<Option<Vec<PathBuf>>>
where
    R: BufRead,
    W: Write,
{
    writeln!(output, "→ 正在打开文件选择窗口：{title}")?;
    if allow_native_dialog {
        if let Some(paths) = native_file_dialog(title, allow_multiple, filter) {
            if !paths.is_empty() {
                if count_only {
                    writeln!(output, "→ 已选择文件 {} 个", paths.len())?;
                } else {
                    print_selected_paths(output, &paths)?;
                }
                return Ok(Some(paths));
            }
        }
        writeln!(output, "[提示] 文件弹窗不可用或已取消。")?;
    }
    writeln!(output, "→ 可手动输入路径（分号分隔），或输入 0 取消：")?;
    let Some(raw) = read_menu_choice(input, output, "文件路径> ")? else {
        return Ok(None);
    };
    let Some(paths) = parse_manual_path_list(&raw) else {
        return Ok(None);
    };
    if count_only {
        writeln!(output, "→ 已选择文件 {} 个", paths.len())?;
    } else {
        print_selected_paths(output, &paths)?;
    }
    Ok(Some(paths))
}

fn choose_source_files_or_folders<R, W>(
    title: &str,
    filter: NativeFileFilter,
    input: &mut R,
    output: &mut W,
    allow_native_dialog: bool,
) -> Result<Option<Vec<PathBuf>>>
where
    R: BufRead,
    W: Write,
{
    choose_source_files_or_folders_impl(
        title,
        filter,
        input,
        output,
        allow_native_dialog,
        false,
        false,
    )
}

fn choose_source_files_or_folders_impl<R, W>(
    title: &str,
    filter: NativeFileFilter,
    input: &mut R,
    output: &mut W,
    allow_native_dialog: bool,
    count_only: bool,
    summary_folder_rules: bool,
) -> Result<Option<Vec<PathBuf>>>
where
    R: BufRead,
    W: Write,
{
    writeln!(output, "→ {title}")?;
    writeln!(output, "  1  选择文件（可多选）")?;
    writeln!(output, "  2  选择文件夹（仅当前层）")?;
    writeln!(output, "  3  选择文件夹（递归子目录）")?;
    writeln!(output, "  0  取消")?;
    let Some(mode) = read_menu_choice(input, output, "选择方式> ")? else {
        return Ok(None);
    };
    if mode == "0" {
        return Ok(None);
    }
    if mode == "1" {
        return choose_input_paths_impl(
            title,
            true,
            filter,
            input,
            output,
            allow_native_dialog,
            count_only,
        );
    }
    anyhow::ensure!(matches!(mode.as_str(), "2" | "3"), "仅接受 1、2、3 或 0");

    let native_folder = allow_native_dialog
        .then(|| native_folder_dialog(title))
        .flatten();
    let mut folders = native_folder.into_iter().collect::<Vec<_>>();
    if folders.is_empty() {
        writeln!(
            output,
            "→ 可手动输入文件夹路径（多个路径用分号分隔），或输入 0 取消："
        )?;
        let Some(raw) = read_menu_choice(input, output, "文件夹路径> ")? else {
            return Ok(None);
        };
        let Some(paths) = parse_manual_path_list(&raw) else {
            return Ok(None);
        };
        folders.extend(paths);
    } else {
        print_selected_paths(output, &folders)?;
        writeln!(output, "→ 可补充其它文件夹路径（分号分隔，留空继续）：")?;
        if let Some(raw) = read_menu_choice(input, output, "补充路径> ")? {
            if let Some(paths) = parse_manual_path_list(&raw) {
                folders.extend(paths);
            }
        }
    }
    let paths = if summary_folder_rules {
        let mut unique = std::collections::BTreeMap::new();
        for folder in &folders {
            let root = fs::canonicalize(folder).with_context(|| {
                format!("读取汇总源文件夹失败：{}", display_path_for_user(&folder))
            })?;
            let mut candidates = collect_menu_source_files_impl(
                std::slice::from_ref(folder),
                mode == "3",
                filter,
                true,
            )?;
            candidates.retain(|path| {
                !path
                    .file_stem()
                    .is_some_and(|name| name.to_string_lossy().contains("_审核版"))
                    && !path
                        .strip_prefix(&root)
                        .unwrap_or(path)
                        .components()
                        .any(|part| part.as_os_str() == "执行结果")
            });
            if mode == "3"
                && candidates
                    .iter()
                    .any(|path| path.parent() != Some(root.as_path()))
            {
                candidates.retain(|path| path.parent() != Some(root.as_path()));
            }
            for path in candidates {
                unique
                    .entry(path.to_string_lossy().to_lowercase())
                    .or_insert(path);
            }
        }
        unique.into_values().collect::<Vec<_>>()
    } else {
        collect_menu_source_files(&folders, mode == "3", filter)?
    };
    writeln!(
        output,
        "→ 已选 {} 个文件夹，{}共识别文件 {} 个",
        folders.len(),
        if mode == "3" { "递归" } else { "当前层" },
        paths.len()
    )?;
    if !count_only {
        print_selected_paths(output, &paths)?;
    }
    Ok(Some(paths))
}

fn collect_menu_source_files(
    folders: &[PathBuf],
    recursive: bool,
    filter: NativeFileFilter,
) -> Result<Vec<PathBuf>> {
    collect_menu_source_files_impl(folders, recursive, filter, false)
}

fn collect_menu_source_files_impl(
    folders: &[PathBuf],
    recursive: bool,
    filter: NativeFileFilter,
    skip_summary_results: bool,
) -> Result<Vec<PathBuf>> {
    let mut files = Vec::new();
    for folder in folders {
        anyhow::ensure!(
            folder.is_dir(),
            "selected folder does not exist or is not a directory: {}",
            display_path_for_user(&folder)
        );
        collect_menu_source_files_from_folder(
            folder,
            folder,
            recursive,
            filter,
            skip_summary_results,
            &mut files,
        )?;
    }
    let mut unique = std::collections::BTreeMap::new();
    for path in files {
        let resolved = fs::canonicalize(&path).unwrap_or(path);
        let identity = resolved.to_string_lossy().to_lowercase();
        unique.entry(identity).or_insert(resolved);
    }
    Ok(unique.into_values().collect())
}

fn collect_menu_source_files_from_folder(
    root: &Path,
    folder: &Path,
    recursive: bool,
    filter: NativeFileFilter,
    skip_summary_results: bool,
    files: &mut Vec<PathBuf>,
) -> Result<()> {
    let mut entries = fs::read_dir(folder)
        .with_context(|| format!("list selected folder {}", display_path_for_user(&folder)))?
        .map(|entry| entry.map(|entry| entry.path()))
        .collect::<std::io::Result<Vec<_>>>()?;
    entries.sort_by_key(|path| {
        path.strip_prefix(root)
            .unwrap_or(path)
            .to_string_lossy()
            .to_lowercase()
    });
    for path in entries {
        let metadata = fs::symlink_metadata(&path)
            .with_context(|| format!("inspect selected path {}", display_path_for_user(&path)))?;
        if metadata.file_type().is_symlink() {
            continue;
        }
        if metadata.is_file() {
            if !path
                .file_name()
                .is_some_and(|name| name.to_string_lossy().starts_with("~$"))
                && filter.accepts_folder_file(&path)
            {
                files.push(path);
            }
        } else if recursive
            && metadata.is_dir()
            && !(skip_summary_results && path.file_name().is_some_and(|name| name == "执行结果"))
        {
            collect_menu_source_files_from_folder(
                root,
                &path,
                true,
                filter,
                skip_summary_results,
                files,
            )?;
        }
    }
    Ok(())
}

fn parse_manual_path_list(raw: &str) -> Option<Vec<PathBuf>> {
    let raw = raw.trim();
    if raw.is_empty() || raw == "0" {
        return None;
    }
    let paths = raw
        .split(';')
        .map(|part| {
            part.trim()
                .trim_matches(|character| matches!(character, '\"' | '\''))
        })
        .filter(|part| !part.is_empty())
        .map(PathBuf::from)
        .collect::<Vec<_>>();
    (!paths.is_empty()).then_some(paths)
}

fn print_selected_paths<W: Write>(output: &mut W, paths: &[PathBuf]) -> Result<()> {
    writeln!(output, "→ 已选择文件 {} 个：", paths.len())?;
    let labels = pytools_core::run_logging::disambiguated_file_names(paths);
    let shown = if paths.len() > 12 { 5 } else { paths.len() };
    for (index, label) in labels.iter().take(shown).enumerate() {
        writeln!(output, "  {}. {}", index + 1, label)?;
    }
    if shown < paths.len() {
        writeln!(
            output,
            "  …… 其余 {} 个文件将按批次处理",
            paths.len() - shown
        )?;
    }
    Ok(())
}

#[cfg(windows)]
fn native_file_dialog(
    title: &str,
    allow_multiple: bool,
    file_filter: NativeFileFilter,
) -> Option<Vec<PathBuf>> {
    use std::os::windows::ffi::OsStringExt;
    use windows_sys::Win32::UI::Controls::Dialogs::{
        GetOpenFileNameW, OFN_ALLOWMULTISELECT, OFN_EXPLORER, OFN_FILEMUSTEXIST, OFN_HIDEREADONLY,
        OFN_NOCHANGEDIR, OFN_PATHMUSTEXIST, OPENFILENAMEW,
    };

    let mut filter = file_filter
        .dialog_filter()
        .encode_utf16()
        .collect::<Vec<_>>();
    let mut title = title
        .encode_utf16()
        .chain(std::iter::once(0))
        .collect::<Vec<_>>();
    let mut buffer = vec![0u16; 32_768];
    let mut dialog = OPENFILENAMEW::default();
    dialog.lStructSize = std::mem::size_of::<OPENFILENAMEW>() as u32;
    dialog.lpstrFilter = filter.as_mut_ptr();
    dialog.nFilterIndex = 1;
    dialog.lpstrFile = buffer.as_mut_ptr();
    dialog.nMaxFile = buffer.len() as u32;
    dialog.lpstrTitle = title.as_mut_ptr();
    dialog.Flags = OFN_EXPLORER
        | OFN_FILEMUSTEXIST
        | OFN_PATHMUSTEXIST
        | OFN_NOCHANGEDIR
        | OFN_HIDEREADONLY
        | if allow_multiple {
            OFN_ALLOWMULTISELECT
        } else {
            0
        };

    let selected = unsafe { GetOpenFileNameW(&mut dialog) };
    if selected == 0 {
        return None;
    }
    let parts = buffer
        .split(|unit| *unit == 0)
        .take_while(|part| !part.is_empty())
        .collect::<Vec<_>>();
    if parts.is_empty() {
        return None;
    }
    if parts.len() == 1 {
        return Some(vec![PathBuf::from(std::ffi::OsString::from_wide(parts[0]))]);
    }
    let directory = PathBuf::from(std::ffi::OsString::from_wide(parts[0]));
    Some(
        parts[1..]
            .iter()
            .map(|name| directory.join(std::ffi::OsString::from_wide(name)))
            .collect(),
    )
}

#[cfg(not(windows))]
fn native_file_dialog(
    _title: &str,
    _allow_multiple: bool,
    _file_filter: NativeFileFilter,
) -> Option<Vec<PathBuf>> {
    None
}

#[cfg(windows)]
fn native_folder_dialog(title: &str) -> Option<PathBuf> {
    use windows::core::HSTRING;
    use windows::Win32::System::Com::{CoCreateInstance, CoTaskMemFree, CLSCTX_INPROC_SERVER};
    use windows::Win32::UI::Shell::{
        FileOpenDialog, IFileOpenDialog, FOS_FORCEFILESYSTEM, FOS_PATHMUSTEXIST, FOS_PICKFOLDERS,
        SIGDN_FILESYSPATH,
    };
    use windows_sys::Win32::System::Ole::{OleInitialize, OleUninitialize};

    if unsafe { OleInitialize(std::ptr::null()) } < 0 {
        return None;
    }
    let selected = (|| -> windows::core::Result<PathBuf> {
        unsafe {
            let dialog: IFileOpenDialog =
                CoCreateInstance(&FileOpenDialog, None, CLSCTX_INPROC_SERVER)?;
            dialog.SetOptions(
                dialog.GetOptions()? | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST,
            )?;
            dialog.SetTitle(&HSTRING::from(title))?;
            dialog.Show(None)?;
            let item = dialog.GetResult()?;
            let path = item.GetDisplayName(SIGDN_FILESYSPATH)?;
            let text = path.to_string();
            CoTaskMemFree(Some(path.0.cast()));
            Ok(PathBuf::from(text?))
        }
    })()
    .ok();
    unsafe { OleUninitialize() };
    selected
}

#[cfg(not(windows))]
fn native_folder_dialog(_title: &str) -> Option<PathBuf> {
    None
}

fn run_quick_cli(arguments: Vec<String>, action: &str) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let global = load_global(config)?;
    match action {
        "config" => open_config(
            options
                .values
                .get("config")
                .map(Path::new)
                .context("--config is required")?,
        )?,
        "latest" => {
            let path = open_latest_output(&global.output_dir)?;
            // 机器 JSON 也不泄漏 verbatim 前缀；防御性兜底，正常情况下
            // config 解析已返回可读形式。
            println!(
                "{}",
                serde_json::json!({"latest": path.as_ref().map(|p| display_path_for_user(p))})
            );
        }
        "output" => open_output_directory(&global.output_dir)?,
        _ => unreachable!(),
    };
    Ok(())
}

fn run_config_init_cli(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let path = required_path(&options, "path")?;
    let result = initialize_new_config(&path)?;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}
fn run_config_view_cli(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let path = required_path(&options, "config")?;
    let result = view_config(path)?;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn run_config_repair_plan_cli(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let input = required_path(&options, "input")?;
    let output = required_path(&options, "output")?;
    let capabilities = options.values.get("capabilities");
    let platform = options.values.get("platform");
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let Some(plan) = build_config_repair_plan(&input, &output)? else {
        println!(
            "{{\"status\":\"already_complete\",\"input\":{}}}",
            serde_json::to_string(&input.to_string_lossy())?
        );
        return Ok(());
    };
    if let Some(plan_path) = options.values.get("plan-output") {
        fs::write(plan_path, serde_json::to_vec_pretty(&plan)?)?;
    }
    if let (Some(capabilities), Some(platform)) = (capabilities, platform) {
        let execution = apply_plan_with_capability_file(&plan, capabilities, platform)?;
        println!(
            "{}",
            serde_json::to_string_pretty(&serde_json::json!({
                "plan": plan,
                "plan_file": options.values.get("plan-output"),
                "execution": execution,
            }))?
        );
    } else {
        println!("{}", serde_json::to_string_pretty(&plan)?);
    }
    Ok(())
}

fn run_dedup_comment_plan(arguments: Vec<String>) -> Result<()> {
    let result = dedup_comment_plan_result(arguments)?;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn dedup_comment_plan_result(arguments: Vec<String>) -> Result<serde_json::Value> {
    let options = parse_cli_arguments(arguments)?;
    let input = required_path(&options, "input")?;
    let output = required_path(&options, "output")?;
    let action = parse_duplicate_action(
        options
            .values
            .get("action")
            .map(String::as_str)
            .unwrap_or("check"),
    )?;
    let report = build_comment_duplicate_plan(&input, &output, action)?;
    let capabilities = options.values.get("capabilities");
    let platform = options.values.get("platform");
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let mut result = serde_json::json!({
        "workbook": report.workbook,
        "hit_sheets": report.hit_sheets,
        "duplicate_rows": report.duplicate_rows,
        "skipped_tasks": report.skipped_tasks,
        "plan": report.plan,
        "execution": if capabilities.is_some() { "not_needed" } else { "plan_only" },
    });
    let planned = result.get("plan").filter(|plan| !plan.is_null()).cloned();
    if let Some(plan_value) = planned {
        if let Some(plan_path) = options.values.get("plan-output") {
            if let Some(parent) = Path::new(plan_path)
                .parent()
                .filter(|path| !path.as_os_str().is_empty())
            {
                fs::create_dir_all(parent)?;
            }
            fs::write(plan_path, serde_json::to_vec_pretty(&plan_value)?)?;
            result["plan_file"] = serde_json::json!(plan_path);
        }
        if let (Some(capabilities), Some(platform)) = (capabilities, platform) {
            let plan: excel_editor::ExcelEditPlan = serde_json::from_value(plan_value)?;
            let execution = apply_plan_with_capability_file(&plan, capabilities, platform)?;
            result["execution"] = serde_json::json!(execution);
        }
    }
    Ok(result)
}

/// 全局配置 `输出附带计划文件=否` 时跳过计划 JSON 等开发期附带文件的落盘，
/// 实际使用只保留候选副本等核心输出（PyCode 无此开关，默认行为不变）。
fn plan_files_enabled(config: &Path) -> bool {
    pytools_core::config::load_global(config)
        .map(|global| global.emit_plan_files)
        .unwrap_or(true)
}

fn run_dedup_config_plan(arguments: Vec<String>) -> Result<()> {
    let result = dedup_config_plan_result(arguments)?;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn dedup_config_check_result(arguments: Vec<String>) -> Result<serde_json::Value> {
    dedup_config_result(arguments, DuplicateAction::Check)
}

fn dedup_config_delete_result(arguments: Vec<String>) -> Result<serde_json::Value> {
    dedup_config_result(arguments, DuplicateAction::Delete)
}

fn dedup_config_result(
    arguments: Vec<String>,
    action: DuplicateAction,
) -> Result<serde_json::Value> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let output_dir = options
        .values
        .get("output-dir")
        .map(PathBuf::from)
        .unwrap_or(load_global(&config)?.output_dir);
    let emit_plan_files = plan_files_enabled(&config);
    let tasks = load_dedup_tasks(&config)?;
    let mut outputs = std::collections::BTreeMap::new();
    let output_suffix = if action == DuplicateAction::Delete {
        "dedup-delete"
    } else {
        "dedup-check"
    };
    for task in tasks.iter().filter(|task| task.enabled) {
        if let Some(input) = &task.source_wb {
            outputs
                .entry(input.clone())
                .or_insert_with(|| unique_workbook_candidate(input, &output_dir, output_suffix));
        }
    }
    fs::create_dir_all(&output_dir)?;
    let reports = build_config_duplicate_plan(&tasks, &outputs, action)?;
    let mut results = Vec::with_capacity(reports.len());
    for (index, report) in reports.into_iter().enumerate() {
        let input = report
            .workbook
            .as_ref()
            .context("configured duplicate report has no source workbook")?;
        let output = outputs
            .get(input)
            .context("configured duplicate output mapping is missing")?;
        let mut plan_path = None;
        if emit_plan_files {
            if let Some(plan) = &report.plan {
                let path = output_dir.join(format!(
                    "dedup-{}-plan-{}-{index}.json",
                    if action == DuplicateAction::Delete {
                        "delete"
                    } else {
                        "check"
                    },
                    Local::now().format("%Y%m%d_%H%M%S")
                ));
                fs::write(&path, serde_json::to_vec_pretty(plan)?)?;
                plan_path = Some(path);
            }
        }

        let written = if action == DuplicateAction::Delete {
            write_config_duplicate_delete_candidate(input, output, report.plan.as_ref())
        } else {
            write_config_duplicate_check_candidate(input, output, report.plan.as_ref())
        };
        match written {
            Ok(path) => results.push(serde_json::json!({
                "report": report,
                "output_file": path,
                "plan_file": plan_path,
                "execution": "written",
                "status": "succeeded",
            })),
            Err(error) => results.push(serde_json::json!({
                "report": report,
                "output_file": null,
                "plan_file": plan_path,
                "execution": "failed",
                "status": "failed",
                "error": format!("{error:#}"),
            })),
        }
    }
    Ok(serde_json::Value::Array(results))
}

fn dedup_config_plan_result(arguments: Vec<String>) -> Result<serde_json::Value> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let output_dir = options
        .values
        .get("output-dir")
        .map(PathBuf::from)
        .unwrap_or(load_global(&config)?.output_dir);
    let emit_plan_files = plan_files_enabled(&config);
    let action = parse_duplicate_action(
        options
            .values
            .get("action")
            .map(String::as_str)
            .unwrap_or("check"),
    )?;
    let tasks = pytools_core::config::load_dedup_tasks(&config)?;
    let mut outputs = std::collections::BTreeMap::new();
    for task in tasks.iter().filter(|task| task.enabled) {
        if let Some(input) = &task.source_wb {
            outputs
                .entry(input.clone())
                .or_insert_with(|| unique_workbook_candidate(input, &output_dir, "dedup"));
        }
    }
    fs::create_dir_all(&output_dir)?;
    let reports = build_config_duplicate_plan(&tasks, &outputs, action)?;
    let capabilities = options.values.get("capabilities");
    let platform = options.values.get("platform");
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let mut results = Vec::with_capacity(reports.len());
    for (index, report) in reports.iter().enumerate() {
        let mut plan_path = None;
        let mut execution = if capabilities.is_some() {
            serde_json::json!("not_needed")
        } else {
            serde_json::json!("plan_only")
        };
        if let Some(plan) = &report.plan {
            if emit_plan_files {
                let path = output_dir.join(format!(
                    "dedup-plan-{}-{index}.json",
                    chrono::Local::now().format("%Y%m%d_%H%M%S")
                ));
                fs::write(&path, serde_json::to_vec_pretty(plan)?)?;
                plan_path = Some(path.clone());
            }
            if let (Some(capabilities), Some(platform)) = (capabilities, platform) {
                execution = serde_json::json!(apply_plan_with_capability_file(
                    plan,
                    capabilities,
                    platform
                )?);
            }
        }
        results.push(serde_json::json!({
            "report": report,
            "plan_file": plan_path,
            "execution": execution,
        }));
    }
    Ok(serde_json::Value::Array(results))
}

fn run_dedup_precheck(arguments: Vec<String>) -> Result<()> {
    let report = dedup_precheck_result(arguments)?;
    println!("{}", serde_json::to_string_pretty(&report)?);
    if report["fail"].as_u64().unwrap_or_default() > 0 {
        bail!(
            "dedup precheck found {} failure(s)",
            report["fail"].as_u64().unwrap_or_default()
        );
    }
    Ok(())
}

fn dedup_precheck_result(arguments: Vec<String>) -> Result<serde_json::Value> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let tasks = pytools_core::config::load_dedup_tasks(config)?;
    let report = precheck_dedup_tasks(&tasks);
    Ok(serde_json::to_value(report)?)
}

fn run_dedup_append_plan(arguments: Vec<String>) -> Result<()> {
    let report = dedup_append_plan_result(arguments)?;
    println!("{}", serde_json::to_string_pretty(&report)?);
    Ok(())
}

fn dedup_append_plan_result(arguments: Vec<String>) -> Result<serde_json::Value> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let global = load_global(&config)?;
    let output_dir = options
        .values
        .get("output-dir")
        .map(PathBuf::from)
        .unwrap_or(global.output_dir);
    Ok(dedup_append_plan_bundle(&config, &output_dir, global.emit_plan_files)?.0)
}

fn dedup_append_plan_bundle(
    config: &Path,
    output_dir: &Path,
    emit_plan_files: bool,
) -> Result<(serde_json::Value, Vec<AppendTargetPlan>)> {
    let tasks = load_dedup_tasks(config)?;
    dedup_append_plan_bundle_with_tasks(&tasks, output_dir, emit_plan_files)
}

fn dedup_append_plan_bundle_with_tasks(
    tasks: &[pytools_core::config::DedupTask],
    output_dir: &Path,
    emit_plan_files: bool,
) -> Result<(serde_json::Value, Vec<AppendTargetPlan>)> {
    let reports = build_append_by_config_plans(tasks, output_dir)?;
    fs::create_dir_all(&output_dir)?;
    let timestamp = Local::now().format("%Y%m%d_%H%M%S_%f");
    let mut output = Vec::new();
    for (index, report) in reports.iter().enumerate() {
        let plan_path = if emit_plan_files {
            let path = output_dir.join(format!("dedup-append-{timestamp}-{index}.json"));
            write_json_report_create_new(&path, report)?;
            Some(path)
        } else {
            None
        };
        output.push(serde_json::json!({
            "target": report.target,
            "candidate": report.candidate,
            "target_existed": report.target_existed,
            "task_ok": report.task_ok,
            "task_skipped": report.task_skipped,
            "appended_rows": report.appended_rows,
            "duplicate_rows_filtered": report.duplicate_rows_filtered,
            "merged_tasks": report.merged_tasks,
            "issues": report.issues,
            "plan_file": plan_path,
            "execution": if report.target_existed { "requires apply-edit-plan capability evidence" } else { "run dedup-append-execute" },
        }));
    }
    Ok((serde_json::Value::Array(output), reports))
}

fn run_dedup_append_execute(arguments: Vec<String>) -> Result<()> {
    let report = run_dedup_append_execute_report(arguments)?;
    println!("{}", serde_json::to_string_pretty(&report)?);
    Ok(())
}

fn run_dedup_append_direct_existing(plan: &AppendTargetPlan) -> Result<serde_json::Value> {
    let started = Instant::now();
    let expected_hash = plan
        .target_sha256
        .as_deref()
        .context("目标快照缺失，请重新读取配置后再追加")?;
    // .xls 目标：COM 追加路由（随包 worker + Excel/WPS）。
    if plan
        .target
        .extension()
        .and_then(|ext| ext.to_str())
        .is_some_and(|ext| ext.eq_ignore_ascii_case("xls"))
        && plan.xls_operations.is_some()
    {
        #[cfg(windows)]
        {
            let xls_candidate = unique_workbook_candidate_preserving_extension(
                &plan.target,
                plan.target.parent().unwrap_or(Path::new(".")),
                "dedup_append_xls",
            );
            let result = (|| -> Result<serde_json::Value> {
                fs::copy(&plan.target, &xls_candidate).with_context(|| {
                    format!(
                        "复制 .xls 目标候选 {}",
                        display_path_for_user(&xls_candidate)
                    )
                })?;
                pytools_core::file_commit::verify_planning_snapshot(
                    &plan.target,
                    Some(expected_hash),
                )?;
                let mut com_runtime = pytools_core::conversion_engine::ConversionRuntime::default();
                // 从 SetCellValue 操作序列化值中提取 2D 行数据，
                // 按目标工作表分组后追加到末尾。
                let ops = plan
                    .xls_operations
                    .as_ref()
                    .expect("xls_operations checked");
                // 按工作表分组 → 每表取最大行/列 → 构建 2D 值矩阵
                let mut sheet_data: std::collections::BTreeMap<
                    String,
                    std::collections::BTreeMap<(u32, u32), serde_json::Value>,
                > = Default::default();
                for op in ops {
                    let sheet = op.get("sheet").and_then(|v| v.as_str()).unwrap_or("");
                    let cell = op.get("cell").and_then(|v| v.as_str()).unwrap_or("");
                    if sheet.is_empty() || cell.is_empty() {
                        continue;
                    }
                    let digits = cell.rfind(|c: char| c.is_ascii_digit()).unwrap_or(0);
                    let row: u32 = cell[digits..].parse().unwrap_or(0);
                    let col_str = &cell[..digits];
                    let mut col = 0u32;
                    for c in col_str.chars() {
                        col = col * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
                    }
                    let value = op.get("value").cloned().unwrap_or(serde_json::Value::Null);
                    sheet_data
                        .entry(sheet.to_owned())
                        .or_default()
                        .insert((row, col), value);
                }
                // 对每个工作表：计算 start_row（现有最大行 + 1），构建 2D 数组，调用 COM
                let append_started = Instant::now();
                for (sheet_name, cells) in &sheet_data {
                    let max_row = cells.keys().map(|(r, _)| *r).max().unwrap_or(0);
                    let max_col = cells.keys().map(|(_, c)| *c).max().unwrap_or(0);
                    let mut rows_2d = Vec::with_capacity(max_row as usize);
                    for r in 1..=max_row {
                        let mut row_vals = Vec::with_capacity(max_col as usize);
                        for c in 1..=max_col {
                            row_vals.push(
                                cells
                                    .get(&(r, c))
                                    .cloned()
                                    .unwrap_or(serde_json::Value::Null),
                            );
                        }
                        rows_2d.push(row_vals);
                    }
                    let mut com_runtime =
                        pytools_core::conversion_engine::ConversionRuntime::default();
                    // 目标 .xls 已有数据，追加到末尾。start_row = 1（第 1 行开始写入）。
                    let com =
                        com_runtime.append_rows_com(&xls_candidate, sheet_name, 1, &rows_2d)?;
                    anyhow::ensure!(
                        com["verified"].as_bool() == Some(true),
                        "Excel/WPS COM 追加未通过校验"
                    );
                    let _ = append_started;
                }
                pytools_core::file_commit::replace_target_from_candidate(
                    &xls_candidate,
                    &plan.target,
                    expected_hash,
                )?;
                Ok(serde_json::json!({
                    "engine": "excel_com",
                    "route": "direct_append_xls_com",
                    "candidate": xls_candidate,
                    "target_commit": expected_hash,
                }))
            })();
            let commit_result = match (result, remove_regular_candidate(&xls_candidate)) {
                (Ok(report), Ok(())) => Ok(report),
                (Ok(_), Err(error)) => {
                    Err(error.context("clean temporary .xls dedup append candidate"))
                }
                (Err(error), Ok(())) => Err(error),
                (Err(error), Err(cleanup_error)) => Err(error.context(format!(
                    "temporary candidate cleanup also failed: {cleanup_error:#}"
                ))),
            };
            let elapsed = started.elapsed().as_millis();
            return match commit_result {
                Ok(mut report_value) => {
                    report_value["elapsed_ms"] = serde_json::json!(elapsed);
                    Ok(report_value)
                }
                Err(error) => Err(error),
            };
        }
        #[cfg(not(windows))]
        {
            anyhow::bail!("2-3 的 .xls 目标追加需要 Windows Excel/WPS COM");
        }
    }
    let candidate_started = Instant::now();
    let candidate =
        write_existing_append_candidate(plan).context("stage=direct_append_candidate")?;
    let candidate_elapsed_ms = candidate_started.elapsed().as_millis();
    let commit = replace_target_from_candidate(&candidate, &plan.target, expected_hash)
        .context("stage=target_commit")?;
    Ok(serde_json::json!({
        "engine": "openpyxl_worker",
        "route": "direct_append",
        "candidate": candidate,
        "target_commit": &commit,
        "execution_summary": {
            "function": "append_target_execute",
            "status": "succeeded",
            "task_success_count": plan.task_ok,
            "task_skipped_count": plan.task_skipped,
            "appended_rows": plan.appended_rows,
            "duplicate_rows_filtered": plan.duplicate_rows_filtered,
            "elapsed_ms": started.elapsed().as_millis(),
            "stages": [
                {
                    "name": "existing_workbook_candidate_and_validation",
                    "status": "succeeded",
                    "engine": "openpyxl",
                    "elapsed_ms": candidate_elapsed_ms
                },
                {
                    "name": "target_commit",
                    "status": "succeeded",
                    "elapsed_ms": commit.elapsed_ms
                }
            ]
        }
    }))
}

fn run_dedup_append_execute_report(arguments: Vec<String>) -> Result<serde_json::Value> {
    let run_started = Instant::now();
    let (commit_target, arguments) = extract_boolean_flag(arguments, "commit-target")?;
    let options = parse_cli_arguments(arguments)?;
    let plan_path = required_path(&options, "plan")?;
    let plan: AppendTargetPlan = serde_json::from_slice(&fs::read(&plan_path)?)
        .with_context(|| format!("decode append plan {}", display_path_for_user(&plan_path)))?;
    let report = if plan.target_existed {
        let edit_plan = plan
            .edit_plan
            .as_ref()
            .context("append plan does not contain an edit plan")?;
        edit_plan.validate().context("stage=edit_plan_validation")?;
        let expected_target_hash = if commit_target {
            let expected = plan.target_sha256.as_deref().context(
                "append plan has no target SHA-256; regenerate the plan before committing",
            )?;
            let actual = sha256_file(&plan.target)?;
            anyhow::ensure!(
                actual == expected,
                "configured target changed after planning; refusing to execute the commit: {}",
                display_path_for_user(&plan.target)
            );
            Some(expected.to_owned())
        } else {
            None
        };
        let staging = if plan.replacement_workbook.is_some() {
            Some(TemporaryFile::new(prepare_wide_target_replacement_source(
                &plan,
            )?))
        } else {
            None
        };
        let candidate_preexisted = plan.candidate.exists();
        let edit_started = Instant::now();
        let execution = match pytools_core::inplace_edit::execute_candidate(
            edit_plan,
            pytools_core::inplace_edit::InPlaceEngine::Openpyxl,
        ) {
            Ok(execution) => execution,
            Err(error) => {
                if !candidate_preexisted && plan.candidate.exists() {
                    let _ = fs::remove_file(&plan.candidate);
                }
                return Err(error).context("stage=excel_modifier");
            }
        };
        let edit_elapsed_ms = edit_started.elapsed().as_millis();
        let mut stages = vec![serde_json::json!({
            "name": "excel_modify",
            "status": "succeeded",
            "engine": execution.engine,
            "elapsed_ms": edit_elapsed_ms
        })];
        if let Some(staging) = &staging {
            let validation_started = Instant::now();
            if let Err(error) = validate_wide_target_replacement_candidate(&plan, staging.path()) {
                if !candidate_preexisted && plan.candidate.exists() {
                    fs::remove_file(&plan.candidate).with_context(|| {
                        format!(
                            "remove invalid wide target candidate {}",
                            display_path_for_user(&plan.candidate)
                        )
                    })?;
                }
                return Err(error).context("stage=replacement_contract_validation");
            }
            stages.push(serde_json::json!({
                "name": "replacement_contract_validation",
                "status": "succeeded",
                "elapsed_ms": validation_started.elapsed().as_millis()
            }));
        }
        let mut output = serde_json::to_value(&execution)?;
        if staging.is_some() {
            output["replacement_validation"] = serde_json::Value::String("passed".to_owned());
        }
        if let Some(expected_target_hash) = expected_target_hash {
            let commit =
                replace_target_from_candidate(&plan.candidate, &plan.target, &expected_target_hash)
                    .context("stage=target_commit")?;
            stages.push(serde_json::json!({
                "name": "target_commit",
                "status": "succeeded",
                "elapsed_ms": commit.elapsed_ms
            }));
            output["target_commit"] = serde_json::to_value(commit)?;
        } else {
            stages.push(serde_json::json!({
                "name": "target_commit",
                "status": "not_requested"
            }));
        }
        output["execution_summary"] = serde_json::json!({
            "function": "append_target_execute",
            "status": "succeeded",
            "task_success_count": plan.task_ok,
            "task_skipped_count": plan.task_skipped,
            "appended_rows": plan.appended_rows,
            "duplicate_rows_filtered": plan.duplicate_rows_filtered,
            "elapsed_ms": run_started.elapsed().as_millis(),
            "stages": stages
        });
        output
    } else {
        let write_started = Instant::now();
        let candidate =
            write_new_append_candidate(&plan).context("stage=new_workbook_write_and_reopen")?;
        let mut stages = vec![serde_json::json!({
            "name": "new_workbook_write_and_reopen",
            "status": "succeeded",
            "engine": "rust_xlsxwriter",
            "elapsed_ms": write_started.elapsed().as_millis()
        })];
        let mut output = serde_json::json!({
            "candidate": candidate,
            "candidate_engine": "rust_xlsxwriter",
            "status": "created_and_reopened"
        });
        if commit_target {
            let commit = create_target_from_candidate(&candidate, &plan.target)
                .context("stage=target_commit")?;
            stages.push(serde_json::json!({
                "name": "target_commit",
                "status": "succeeded",
                "elapsed_ms": commit.elapsed_ms
            }));
            output["target_commit"] = serde_json::to_value(commit)?;
            output["status"] = serde_json::Value::String("created_candidate_and_target".into());
        } else {
            stages.push(serde_json::json!({
                "name": "target_commit",
                "status": "not_requested"
            }));
        }
        output["execution_summary"] = serde_json::json!({
            "function": "append_target_execute",
            "status": "succeeded",
            "task_success_count": plan.task_ok,
            "task_skipped_count": plan.task_skipped,
            "appended_rows": plan.appended_rows,
            "duplicate_rows_filtered": plan.duplicate_rows_filtered,
            "elapsed_ms": run_started.elapsed().as_millis(),
            "stages": stages
        });
        output
    };
    Ok(report)
}

struct TemporaryFile(PathBuf);

impl TemporaryFile {
    fn new(path: PathBuf) -> Self {
        Self(path)
    }

    fn path(&self) -> &Path {
        &self.0
    }
}

impl Drop for TemporaryFile {
    fn drop(&mut self) {
        let _ = fs::remove_file(&self.0);
    }
}

fn run_apply_edit_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let plan = required_path(&options, "plan")?;
    let capabilities = required_path(&options, "capabilities")?;
    let platform = options
        .values
        .get("platform")
        .context("--platform is required")?;
    let report = apply_plan_file(plan, capabilities, platform)?;
    println!("{}", serde_json::to_string_pretty(&report)?);
    Ok(())
}

fn parse_duplicate_action(value: &str) -> Result<DuplicateAction> {
    match value {
        "check" => Ok(DuplicateAction::Check),
        "mark" => Ok(DuplicateAction::Mark),
        "delete" => Ok(DuplicateAction::Delete),
        _ => bail!("--action must be check, mark, or delete"),
    }
}

fn run_print_comment_plan(arguments: Vec<String>) -> Result<()> {
    let result = build_print_comment_plan_report(arguments)?;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn build_print_comment_plan_report(arguments: Vec<String>) -> Result<serde_json::Value> {
    let options = parse_cli_arguments(arguments)?;
    let input = required_path(&options, "input")?;
    let output = required_path(&options, "output")?;
    let drop_comments = options
        .values
        .get("drop-comments")
        .is_some_and(|value| parse_cli_bool(value));
    let hide_zero = options
        .values
        .get("hide-zero")
        .is_some_and(|value| parse_cli_bool(value));
    let report = build_comment_print_plan(&input, &output, drop_comments, hide_zero)?;
    let capabilities = options.values.get("capabilities");
    let platform = options.values.get("platform");
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let mut result = serde_json::json!({
        "input": report.input,
        "output": report.output,
        "matching_sheets": report.matching_sheets,
        "marked_ranges": report.marked_ranges,
        "plan": report.plan,
        "execution": if capabilities.is_some() { "not_needed" } else { "plan_only" },
    });
    let planned = result.get("plan").filter(|plan| !plan.is_null()).cloned();
    if let Some(plan_value) = planned {
        if let Some(plan_path) = options.values.get("plan-output") {
            if let Some(parent) = Path::new(plan_path)
                .parent()
                .filter(|path| !path.as_os_str().is_empty())
            {
                fs::create_dir_all(parent)?;
            }
            fs::write(plan_path, serde_json::to_vec_pretty(&plan_value)?)?;
            result["plan_file"] = serde_json::json!(plan_path);
        }
        if let (Some(capabilities), Some(platform)) = (capabilities, platform) {
            let plan: excel_editor::ExcelEditPlan = serde_json::from_value(plan_value)?;
            let execution = apply_plan_with_capability_file(&plan, capabilities, platform)?;
            result["execution"] = serde_json::json!(execution);
        }
    }
    Ok(result)
}

fn parse_cli_bool(value: &str) -> bool {
    matches!(
        value.trim().to_ascii_lowercase().as_str(),
        "1" | "true" | "yes" | "y" | "是" | "on"
    )
}

fn run_print_config_precheck(arguments: Vec<String>) -> Result<()> {
    let report = build_print_config_precheck_report(arguments)?;
    println!("{}", serde_json::to_string_pretty(&report)?);
    if report.failed > 0 {
        bail!("print precheck found {} failure(s)", report.failed);
    }
    Ok(())
}

fn build_print_config_precheck_report(
    arguments: Vec<String>,
) -> Result<pytools_core::print_ops::PrintPrecheck> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let tasks = pytools_core::config::load_print_tasks(config)?;
    Ok(precheck_print_tasks(&tasks))
}

fn run_print_config_plan(arguments: Vec<String>) -> Result<()> {
    let report = build_print_config_plan_report(arguments)?;
    println!("{}", serde_json::to_string_pretty(&report)?);
    Ok(())
}

fn build_print_config_plan_report(arguments: Vec<String>) -> Result<serde_json::Value> {
    build_print_config_report(arguments, false)
}

fn build_print_config_native_report(arguments: Vec<String>) -> Result<serde_json::Value> {
    build_print_config_report(arguments, true)
}

fn build_print_config_report(
    arguments: Vec<String>,
    execute_with_rust_writer: bool,
) -> Result<serde_json::Value> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let output_dir = match options.values.get("output-dir") {
        Some(path) => PathBuf::from(path),
        None => load_global(&config)?.output_dir,
    };
    let emit_plan_files = plan_files_enabled(&config);
    let tasks = pytools_core::config::load_print_tasks(config)?;
    let reports = build_config_print_plans(&tasks)?;
    if !reports.is_empty() {
        fs::create_dir_all(&output_dir)?;
    }
    let capabilities = options.values.get("capabilities");
    let platform = options.values.get("platform");
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let mut output = Vec::new();
    for (index, report) in reports.iter().enumerate() {
        let plan = report
            .plan
            .as_ref()
            .context("grouped print report has no edit plan")?;
        let path = emit_plan_files
            .then(|| -> Result<PathBuf> {
                let path = output_dir.join(format!("print-target-{index}.plan.json"));
                fs::write(&path, serde_json::to_vec_pretty(plan)?)?;
                Ok(path)
            })
            .transpose()?;
        let execution = if let (Some(capabilities), Some(platform)) = (capabilities, platform) {
            let output_parent = plan
                .output
                .parent()
                .filter(|parent| !parent.as_os_str().is_empty())
                .unwrap_or_else(|| std::path::Path::new("."));
            fs::create_dir_all(output_parent).with_context(|| {
                format!(
                    "create print target directory {}",
                    display_path_for_user(&output_parent)
                )
            })?;
            serde_json::json!(apply_plan_with_capability_file(
                plan,
                capabilities,
                platform
            )?)
        } else if execute_with_rust_writer {
            let result = write_static_print_plans(std::slice::from_ref(plan))?
                .into_iter()
                .next()
                .context("Rust static print writer did not return an output")?;
            serde_json::to_value(result)?
        } else {
            serde_json::json!("plan_only")
        };
        output.push(serde_json::json!({
            "report": report,
            "plan_file": path,
            "execution": execution,
        }));
    }
    Ok(serde_json::Value::Array(output))
}

fn run_village_split_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let current = required_path(&options, "current")?;
    let previous = required_path(&options, "previous")?;
    let input = required_path(&options, "input")?;
    let output = required_path(&options, "output")?;
    let report = build_village_split_plan(&current, &previous, &input, &output)?;
    write_optional_plan(&options, &report.plan)?;
    let execution = execute_optional_plan(&options, report.plan.as_ref())?;
    let mut result = serde_json::to_value(report)?;
    result["execution"] = execution;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

/// Extend an existing per-file legacy report without replacing its run ID or
/// treating a plan-only item as a completed workbook modification.
fn attach_legacy_batch_diagnostics(
    tracker: &mut pytools_core::run_logging::RunTracker,
    report: &mut serde_json::Value,
    log_dir: &Path,
    log_view: &str,
) -> Option<pytools_core::run_logging::RunSummary> {
    use pytools_core::run_logging::{EventWriter, FileStatus, Level, TaskStatus};
    if log_view == "off" {
        return None;
    }
    let Some(items) = report["items"].as_array() else {
        return None;
    };
    let files = items
        .iter()
        .map(|item| {
            item["input"]
                .as_str()
                .map(PathBuf::from)
                .unwrap_or_default()
        })
        .collect::<Vec<_>>();
    if tracker.set_inputs(&files).is_err() {
        return None;
    }
    let mut diagnostic = match EventWriter::create(
        &log_dir.join("run-events"),
        &tracker.feature_id,
        &tracker.run_id,
    ) {
        Ok(writer) => Some(writer),
        Err(error) => {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Warn,
                    &format!("无法创建诊断日志：{error}")
                )
            );
            None
        }
    };
    if let Some(writer) = diagnostic.as_mut() {
        let mut event = tracker.event(
            Level::Info,
            "run_started",
            "run",
            "started",
            format!("开始：{}", tracker.feature_name),
        );
        event.timestamp = tracker.started_at.clone();
        if let Err(error) = writer.write(&event) {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Warn,
                    &format!("诊断日志写入失败：{error}")
                )
            );
        }
    }
    let mut outputs = Vec::new();
    let mut incomplete = false;
    for (index, item) in items.iter().enumerate() {
        let position = index + 1;
        let status = item["status"].as_str().unwrap_or("");
        let state = match status {
            "failed" => FileStatus::Failed,
            "committed" | "copy_created" | "candidate_created" | "no_changes" | "succeeded"
            | "created" | "written" => FileStatus::Succeeded,
            "planned" | "plan_only" | "plan_only_no_changes" | "skipped" => FileStatus::Skipped,
            _ => {
                incomplete = true;
                continue;
            }
        };
        if status != "skipped" {
            let _ = tracker.transition(position, FileStatus::InProgress);
        }
        let _ = tracker.transition(position, state);
        if matches!(status, "copy_created" | "created" | "written") {
            if let Some(path) = item["output_copy"]
                .as_str()
                .or_else(|| item["output"].as_str())
                .or_else(|| item["output_path"].as_str())
            {
                outputs.push(path.to_owned());
            }
        } else if status == "committed" {
            if let Some(path) = item["input"].as_str() {
                outputs.push(path.to_owned());
            }
        }
        if let Some(writer) = diagnostic.as_mut() {
            let level = if state == FileStatus::Failed {
                Level::Error
            } else if state == FileStatus::Skipped {
                Level::Warn
            } else {
                Level::Info
            };
            let stage = item["error_stage"].as_str().unwrap_or("file_summary");
            let status_message = match status {
                "committed" => format!("{}已写入源工作簿", tracker.feature_name),
                "copy_created" | "created" | "written" => {
                    format!("{}已生成输出工作簿", tracker.feature_name)
                }
                "candidate_created" => format!("{}候选工作簿已生成", tracker.feature_name),
                "no_changes" => format!("{}未发现需要修改的内容", tracker.feature_name),
                "planned" => format!("{}已生成计划，尚未写入工作簿", tracker.feature_name),
                "plan_only" => format!("{}仅生成计划，未写入工作簿", tracker.feature_name),
                "plan_only_no_changes" => {
                    format!("{}未发现需要修改的内容，未写入工作簿", tracker.feature_name)
                }
                "skipped" => format!("{}跳过当前工作簿", tracker.feature_name),
                "failed" => format!("{}处理失败", tracker.feature_name),
                _ => format!("{}处理结果：{status}", tracker.feature_name),
            };
            let mut event = tracker.event(level, "file_result", stage, status, status_message);
            event.file_index = Some(position);
            event.file_name = Some(pytools_core::run_logging::portable_file_name(&files[index]));
            event.elapsed_ms = item["elapsed_ms"].as_u64().map(u128::from);
            if state == FileStatus::Failed {
                event.error_code = Some("EXCEL_OPERATION_FAILED".to_owned());
            }
            if let Err(error) = writer.write(&event) {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("诊断日志写入失败：{error}")
                    )
                );
                break;
            }
        }
    }
    let failed = report["status"] == "failed"
        || (report["failed_files"].as_u64().unwrap_or(0) > 0
            && tracker.input_stats().succeeded == 0);
    let output_paths = outputs
        .iter()
        .map(|path| path_for_log_view(Path::new(path), log_view))
        .collect();
    let mut summary = tracker.finish(
        failed,
        false,
        items.is_empty(),
        output_paths,
        diagnostic_path_for_log_view(diagnostic.as_ref(), log_view),
    );
    summary.input_stats_available = !incomplete;
    if report["status"] == "partial" {
        summary.overall_status = TaskStatus::PartialSuccess;
    }
    for (field, label) in [
        ("planned_files", "计划文件数"),
        ("committed_files", "已写入文件数"),
        ("created_copies", "生成副本数"),
        ("matched_sheets", "匹配工作表数"),
        ("changed_cells", "修改单元格数"),
        ("fx_sheets", "命中目标工作表数"),
        ("modified_sheets", "修改工作表数"),
        ("saved_files", "保存文件数"),
    ] {
        if !report[field].is_null() {
            summary
                .business_metrics
                .insert(label.to_owned(), report[field].clone());
        }
    }
    if let Some(item) = items.iter().find(|item| item["status"] == "failed") {
        summary.error_summary = Some(pytools_core::run_logging::ErrorSummary {
            stage: item["error_stage"]
                .as_str()
                .unwrap_or("workbook_operation")
                .to_owned(),
            code: "EXCEL_OPERATION_FAILED".to_owned(),
            message: item["error"]
                .as_str()
                .unwrap_or("工作簿处理失败")
                .to_owned(),
            file_name: item["input"]
                .as_str()
                .map(|path| pytools_core::run_logging::portable_file_name(Path::new(path))),
            sheet_name: None,
        });
    }
    if let Some(writer) = diagnostic.as_mut() {
        if let Err(error) = writer.write(&tracker.event(
            if failed { Level::Error } else { Level::Info },
            "run_finished",
            "run",
            &format!("{:?}", summary.overall_status),
            format!("{}结束", tracker.feature_name),
        )) {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Warn,
                    &format!("诊断日志写入失败：{error}")
                )
            );
        }
    }
    let path = log_dir
        .join("run-unified")
        .join(format!("{}_{}.json", tracker.feature_id, tracker.run_id));
    match write_json_report_create_new(&path, &summary) {
        Ok(()) => {
            report["run_summary_path"] = serde_json::json!(path.to_string_lossy());
            if log_view == "user" {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Info,
                        &format!(
                            "运行编号 {}；运行记录：{}",
                            tracker.run_id,
                            pytools_core::run_logging::path_for_log_view(&path, false)
                        )
                    )
                );
            }
        }
        Err(error) => eprintln!(
            "{}",
            pytools_core::run_logging::terminal_line(
                Level::Warn,
                &format!("统一运行记录写入失败：{error}")
            )
        ),
    }
    Some(summary)
}

fn legacy_report_output_paths(report: &serde_json::Value) -> Vec<PathBuf> {
    report["items"]
        .as_array()
        .into_iter()
        .flatten()
        .filter_map(|item| {
            item["output_copy"]
                .as_str()
                .or_else(|| item["output"].as_str())
                .or_else(|| item["output_path"].as_str())
                .or_else(|| {
                    (item["status"] == "committed")
                        .then(|| item["input"].as_str())
                        .flatten()
                })
                .map(PathBuf::from)
        })
        .collect()
}

fn run_village_split_batch(arguments: Vec<String>, interactive: bool) -> Result<()> {
    let (commit_target, arguments) = extract_boolean_flag(arguments, "commit-target")?;
    let (create_copy, arguments) = extract_boolean_flag(arguments, "create-copy")?;
    let options = parse_cli_arguments(arguments)?;
    let in_place = options.values.get("in-place").cloned();
    let capabilities = options.values.get("capabilities");
    let platform = options.values.get("platform");
    let current = required_path(&options, "current")?;
    let previous = required_path(&options, "previous")?;
    let current_extension = current
        .extension()
        .and_then(|extension| extension.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    let previous_extension = previous
        .extension()
        .and_then(|extension| extension.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    anyhow::ensure!(
        current_extension == previous_extension
            && matches!(current_extension.as_str(), "xls" | "xlsx"),
        "PT-7-1 本期与去年同期工作簿必须都是 .xls 或都是 .xlsx"
    );
    let legacy_xls_com = current_extension == "xls";
    if legacy_xls_com {
        anyhow::ensure!(
            cfg!(windows),
            "PT-7-1 的 .xls 路径需要 Windows Excel/WPS COM"
        );
        anyhow::ensure!(
            in_place.as_deref() == Some("com"),
            "PT-7-1 .xls 就地修改需要显式指定 --in-place com"
        );
        anyhow::ensure!(
            !create_copy && !commit_target && capabilities.is_none(),
            "PT-7-1 .xls COM 路径仅支持经快照保护的源文件就地修改"
        );
    }
    anyhow::ensure!(
        !options.positional.is_empty(),
        "provide at least one target workbook"
    );
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );

    anyhow::ensure!(
        !create_copy || (capabilities.is_none() && !commit_target),
        "--create-copy cannot be combined with an ExcelModifier capability"
    );
    let log_view = options
        .values
        .get("log-view")
        .map(String::as_str)
        .unwrap_or("user");
    anyhow::ensure!(
        matches!(log_view, "user" | "developer" | "off"),
        "--log-view must be user, developer, or off"
    );
    let config = options
        .values
        .get("config")
        .map(|path| load_global(path))
        .transpose()?;
    let output_dir = match options.values.get("output-dir") {
        Some(path) => PathBuf::from(path),
        None => config
            .as_ref()
            .map(|config| config.output_dir.clone())
            .unwrap_or_else(|| PathBuf::from("output")),
    };
    let log_dir = match options.values.get("log-dir") {
        Some(path) => PathBuf::from(path),
        None => config
            .as_ref()
            .map(|config| config.log_dir.clone())
            .unwrap_or_else(pytools_core::run_logging::default_log_dir),
    };
    let emit_plan_files = config
        .as_ref()
        .map(|global| global.emit_plan_files)
        .unwrap_or(true);
    let plan_dir = unique_menu_action_dir(&output_dir, "village-split");
    fs::create_dir_all(&output_dir).with_context(|| {
        format!(
            "create output directory {}",
            display_path_for_user(&output_dir)
        )
    })?;
    if emit_plan_files && !create_copy && !legacy_xls_com {
        fs::create_dir_all(&plan_dir).with_context(|| {
            format!("create plan directory {}", display_path_for_user(&plan_dir))
        })?;
    }

    let started = Instant::now();
    let started_at = Local::now().to_rfc3339();
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_nanos())
        .unwrap_or_default();
    let run_id = format!("pytools-pt-7-1-{}-{nonce}", std::process::id());
    let mut unified_tracker = pytools_core::run_logging::RunTracker::with_id(
        "PT-7-1",
        "拆分村镇银行数据",
        &[],
        run_id.clone(),
    );
    let mut planned_files = 0usize;
    let mut committed_files = 0usize;
    let mut completed_files = 0usize;
    let mut failed_files = 0usize;
    let mut matched_sheets = 0usize;
    let mut changed_cells = 0usize;
    let total = options.positional.len();
    let mut events = Vec::<SummaryRunEvent>::new();
    let mut progress = SummaryProgress {
        total_inputs: total,
        total_files: total,
        ..SummaryProgress::default()
    };
    let mut items = Vec::new();
    let mut excel_com_runtime = pytools_core::conversion_engine::ConversionRuntime::default();
    for (index, input) in options.positional.iter().enumerate() {
        let display_name = input
            .file_name()
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_else(|| display_path_for_user(&input).to_string());
        progress.current_file = Some(index + 1);
        progress.processed_files = index;
        progress.completed_files = completed_files;
        progress.failed_files = failed_files;
        if log_view == "user" {
            eprintln!(
                "[处理中] {}/{} 拆分村镇银行数据：{}",
                index + 1,
                total,
                display_name
            );
        }
        emit_village_split_event(
            &mut events,
            log_view,
            started,
            "planning",
            "started",
            format!(
                "开始拆分第 {}/{} 个工作簿中的村镇银行数据",
                index + 1,
                total
            ),
            &progress,
            None,
            &["pytools-core"],
            None,
            None,
        );
        let file_started = Instant::now();
        let mut current_stage = "planning";
        let candidate = if legacy_xls_com {
            unique_workbook_candidate_preserving_extension(input, &output_dir, "village_split")
        } else {
            unique_workbook_candidate(input, &output_dir, "village_split")
        };
        let plan_path = plan_dir.join(format!("target-{:04}.plan.json", index + 1));
        let item_result = (|| -> Result<serde_json::Value> {
            let target_extension = input
                .extension()
                .and_then(|extension| extension.to_str())
                .unwrap_or_default()
                .to_ascii_lowercase();
            anyhow::ensure!(
                target_extension == current_extension,
                "PT-7-1 目标格式 .{target_extension} 与参考工作簿格式 .{current_extension} 不一致"
            );
            if legacy_xls_com {
                current_stage = "com_candidate_copy";
                let expected_target_hash = sha256_file(input)?;
                let expected_current_hash = sha256_file(&current)?;
                let expected_previous_hash = sha256_file(&previous)?;
                fs::copy(input, &candidate).with_context(|| {
                    format!(
                        "复制 .xls 候选 {} → {}",
                        display_path_for_user(input),
                        display_path_for_user(&candidate)
                    )
                })?;
                let isolated_current = unique_workbook_candidate_preserving_extension(
                    &current,
                    &output_dir,
                    "village_current_reference",
                );
                let isolated_previous = unique_workbook_candidate_preserving_extension(
                    &previous,
                    &output_dir,
                    "village_previous_reference",
                );
                if let Err(error) = fs::copy(&current, &isolated_current) {
                    let _ = remove_regular_candidate(&isolated_current);
                    return Err(error).with_context(|| {
                        format!(
                            "为 Excel/WPS COM 隔离本期参考工作簿失败：{}",
                            display_path_for_user(&current)
                        )
                    });
                }
                if let Err(error) = fs::copy(&previous, &isolated_previous) {
                    let cleanup_error = remove_regular_candidate(&isolated_current).err();
                    let error = anyhow::Error::from(error).context(format!(
                        "为 Excel/WPS COM 隔离上期参考工作簿失败：{}",
                        display_path_for_user(&previous)
                    ));
                    return Err(match cleanup_error {
                        Some(cleanup_error) => {
                            error.context(format!("本期参考副本清理也失败：{cleanup_error:#}"))
                        }
                        None => error,
                    });
                }
                current_stage = "com_execution";
                let com_result = excel_com_runtime.village_split_com(
                    &isolated_current,
                    &isolated_previous,
                    &candidate,
                );
                let cleanup_current = remove_regular_candidate(&isolated_current);
                let cleanup_previous = remove_regular_candidate(&isolated_previous);
                let com_result = match com_result {
                    Ok(result) => {
                        cleanup_current.context("清理本期 COM 参考副本失败")?;
                        cleanup_previous.context("清理上期 COM 参考副本失败")?;
                        result
                    }
                    Err(error) => {
                        let cleanup_error =
                            cleanup_current.err().or_else(|| cleanup_previous.err());
                        return Err(match cleanup_error {
                            Some(cleanup_error) => error.context(format!(
                                "stage=pywin32_com_village_split；参考副本清理也失败：{cleanup_error:#}"
                            )),
                            None => error.context("stage=pywin32_com_village_split"),
                        });
                    }
                };
                verify_planning_snapshot(&current, Some(&expected_current_hash))
                    .context("本期参考工作簿在 COM 处理期间发生变化")?;
                verify_planning_snapshot(&previous, Some(&expected_previous_hash))
                    .context("去年同期参考工作簿在 COM 处理期间发生变化")?;
                if com_result["matrix_empty"] == true {
                    emit_village_split_event(
                        &mut events,
                        log_view,
                        started,
                        "com_execution",
                        "no_changes",
                        "COM 读取到的村镇银行参考数据为空；目标未提交",
                        &progress,
                        Some(file_started),
                        &["pywin32", "Excel/WPS COM"],
                        com_result["application"].as_str(),
                        None,
                    );
                    return Ok(serde_json::json!({
                        "input": input,
                        "candidate": candidate,
                        "output_copy": null,
                        "plan_file": null,
                        "status": "no_changes",
                        "matched_sheets": 0,
                        "changed_cells": 0,
                        "warnings": ["COM 读取到的村镇银行参考数据为空，候选未提交"],
                        "execution": com_result,
                        "target_commit": null,
                        "elapsed_ms": file_started.elapsed().as_millis(),
                    }));
                }
                let filled_sheets = com_result["filled_sheets"].as_u64().unwrap_or_default();
                let branch_sheets = com_result["branch_sheets"].as_u64().unwrap_or_default();
                let skipped_sheets = com_result["skipped_sheets"].as_u64().unwrap_or_default();
                emit_village_split_event(
                    &mut events,
                    log_view,
                    started,
                    "com_execution",
                    "completed",
                    format!(
                        "COM 已处理分机构工作表：匹配 {filled_sheets} 个，跳过 {skipped_sheets} 个"
                    ),
                    &progress,
                    Some(file_started),
                    &["pywin32", "Excel/WPS COM"],
                    com_result["application"].as_str(),
                    None,
                );
                let commit = if filled_sheets > 0 {
                    current_stage = "safe_commit";
                    emit_village_split_event(
                        &mut events,
                        log_view,
                        started,
                        "safe_commit",
                        "started",
                        "正在校验源文件快照并安全提交 .xls 候选",
                        &progress,
                        None,
                        &["pytools-core", "file-commit"],
                        None,
                        None,
                    );
                    verify_planning_snapshot(input, Some(&expected_target_hash))?;
                    let commit =
                        replace_target_from_candidate(&candidate, input, &expected_target_hash)
                            .context("stage=target_commit")?;
                    emit_village_split_event(
                        &mut events,
                        log_view,
                        started,
                        "safe_commit",
                        "completed",
                        "候选校验通过，已备份并安全替换 .xls 源工作簿",
                        &progress,
                        Some(file_started),
                        &["pytools-core", "file-commit"],
                        None,
                        None,
                    );
                    Some(serde_json::to_value(commit)?)
                } else {
                    None
                };
                return Ok(serde_json::json!({
                    "input": input,
                    "candidate": candidate,
                    "output_copy": null,
                    "plan_file": null,
                    "status": if commit.is_some() { "committed" } else { "no_changes" },
                    "matched_sheets": filled_sheets,
                    "changed_cells": filled_sheets * 18,
                    "warnings": if branch_sheets == 0 { vec!["未发现名称包含“分机构”的工作表".to_owned()] } else if skipped_sheets > 0 { vec![format!("跳过 {skipped_sheets} 个无法识别或未找到村镇银行行的工作表")] } else { Vec::<String>::new() },
                    "execution": com_result,
                    "target_commit": commit,
                    "elapsed_ms": file_started.elapsed().as_millis(),
                }));
            }
            if in_place.is_none() {
                anyhow::ensure!(
                    target_extension == "xlsx",
                    "PT-7-1 计划模式当前只接受 .xlsx"
                );
            }
            // 就地模式同样需要规划期快照，用于提交前校验源未被改动。
            let expected_target_hash = if commit_target || create_copy || in_place.is_some() {
                Some(sha256_file(input)?)
            } else {
                None
            };
            let planning_started = Instant::now();
            let report = build_village_split_plan(&current, &previous, input, &candidate)?;
            if let Some(expected_hash) = expected_target_hash.as_deref() {
                verify_planning_snapshot(input, Some(expected_hash))?;
            }
            emit_village_split_event(
                &mut events,
                log_view,
                started,
                "planning",
                if report.plan.is_some() {
                    "completed"
                } else {
                    "no_changes"
                },
                if report.plan.is_some() {
                    "已完成业务计算并生成工作簿修改计划"
                } else {
                    "没有需要修改的工作表"
                },
                &progress,
                Some(planning_started),
                &["pytools-core", "excel-reader", "calamine"],
                Some("calamine"),
                None,
            );
            let execution = if let Some(plan) = report.plan.as_ref() {
                current_stage = "plan_persistence";
                let persistence_started = Instant::now();
                if emit_plan_files && !create_copy {
                    fs::write(&plan_path, serde_json::to_vec_pretty(plan)?).with_context(|| {
                        format!("write edit plan {}", display_path_for_user(&plan_path))
                    })?;
                    emit_village_split_event(
                        &mut events,
                        log_view,
                        started,
                        "plan_persistence",
                        "completed",
                        "已保存工作簿修改计划",
                        &progress,
                        Some(persistence_started),
                        &["pytools-core"],
                        None,
                        None,
                    );
                }
                current_stage = "engine_execution";
                let execution_started = Instant::now();
                let execution = if create_copy {
                    current_stage = "engine_execution";
                    create_rust_report_copy(
                        input,
                        &candidate,
                        plan,
                        expected_target_hash.as_deref(),
                    )?
                } else if in_place.is_some() {
                    // 就地修改：execute_optional_plan 读取 --in-place 后直接改写目标。
                    current_stage = "in_place_execution";
                    execute_optional_plan(&options, Some(plan))?
                } else {
                    execute_optional_plan(&options, Some(plan))?
                };
                let engine = execution.get("engine").and_then(serde_json::Value::as_str);
                let plan_only = execution.as_str() == Some("plan_only");
                emit_village_split_event(
                    &mut events,
                    log_view,
                    started,
                    "engine_execution",
                    if plan_only { "skipped" } else { "completed" },
                    if plan_only {
                        "已生成计划；当前没有已验收的执行 capability，原工作簿未修改"
                    } else if create_copy {
                        "已使用 Rust 新建并校验专用报表副本；源工作簿保持不变"
                    } else {
                        "已完成工作簿候选修改与引擎验证"
                    },
                    &progress,
                    Some(execution_started),
                    if plan_only || create_copy {
                        &["pytools-core"]
                    } else {
                        &["pytools-core", "excel-editor"]
                    },
                    engine,
                    None,
                );
                execution
            } else {
                serde_json::Value::String("no_changes".to_owned())
            };
            let commit = if in_place.is_some() && report.plan.is_some() {
                current_stage = "safe_commit";
                let commit_started = Instant::now();
                let commit = execution
                    .get("commit")
                    .cloned()
                    .context("in-place execution completed without a target commit report")?;
                emit_village_split_event(
                    &mut events,
                    log_view,
                    started,
                    "safe_commit",
                    "completed",
                    "候选校验通过；已备份并安全替换所选源工作簿",
                    &progress,
                    Some(commit_started),
                    &["pytools-core", "file-commit"],
                    execution.get("engine").and_then(serde_json::Value::as_str),
                    None,
                );
                Some(serde_json::to_value(commit)?)
            } else if commit_target && report.plan.is_some() {
                current_stage = "safe_commit";
                let commit_started = Instant::now();
                let commit = replace_target_from_candidate(
                    &candidate,
                    input,
                    expected_target_hash
                        .as_deref()
                        .context("target snapshot missing before commit")?,
                )?;
                emit_village_split_event(
                    &mut events,
                    log_view,
                    started,
                    "safe_commit",
                    "completed",
                    "已验证候选并安全提交到目标工作簿",
                    &progress,
                    Some(commit_started),
                    &["pytools-core", "file-commit"],
                    execution.get("engine").and_then(serde_json::Value::as_str),
                    None,
                );
                Some(serde_json::to_value(commit)?)
            } else {
                None
            };
            let output_copy = execution
                .get("output")
                .and_then(serde_json::Value::as_str)
                .map(PathBuf::from);
            Ok(serde_json::json!({
                "input": input,
                "candidate": candidate,
                "output_copy": output_copy,
                "plan_file": if report.plan.is_some() && emit_plan_files && !create_copy { Some(plan_path) } else { None },
                "status": if commit.is_some() {
                    "committed"
                } else if report.plan.is_none() {
                    "no_changes"
                } else if create_copy {
                    "copy_created"
                } else if execution.as_str() == Some("plan_only") {
                    "planned"
                } else {
                    "candidate_created"
                },
                "matched_sheets": matched_sheets,
                "changed_cells": report.changed_cells,
                "warnings": report.warnings,
                "execution": execution,
                "target_commit": commit,
                "elapsed_ms": file_started.elapsed().as_millis(),
            }))
        })();
        let item_result = match (item_result, remove_regular_candidate(&candidate)) {
            (Ok(item), Ok(())) => Ok(item),
            (Ok(_), Err(error)) => Err(error.context("clean temporary village-split candidate")),
            (Err(error), Ok(())) => Err(error),
            (Err(error), Err(cleanup_error)) => Err(error.context(format!(
                "temporary candidate cleanup also failed: {cleanup_error:#}"
            ))),
        };
        match item_result {
            Ok(item) => {
                if item["plan_file"].is_string() {
                    planned_files += 1;
                }
                if item["status"] == "committed" {
                    committed_files += 1;
                }
                completed_files += 1;
                matched_sheets += item["matched_sheets"].as_u64().unwrap_or_default() as usize;
                changed_cells += item["changed_cells"].as_u64().unwrap_or_default() as usize;
                let (status, message) = match item["status"].as_str().unwrap_or_default() {
                    "committed" => ("completed", "当前工作簿已安全提交"),
                    "copy_created" => ("completed", "已生成拆分结果副本，源工作簿保持不变"),
                    "planned" => ("planned", "当前工作簿修改计划已生成，原文件未修改"),
                    "candidate_created" => ("completed", "当前工作簿候选已生成"),
                    _ => ("no_changes", "当前工作簿无需修改"),
                };
                items.push(item);
                progress.processed_files = index + 1;
                progress.completed_files = completed_files;
                emit_village_split_event(
                    &mut events,
                    log_view,
                    started,
                    "file_summary",
                    status,
                    message,
                    &progress,
                    Some(file_started),
                    &["pytools-core"],
                    None,
                    None,
                );
            }
            Err(error) => {
                failed_files += 1;
                progress.processed_files = index + 1;
                progress.failed_files = failed_files;
                let failure_message = format!(
                    "工作簿处理失败：{}；阶段：{}；原因：{error:#}",
                    display_path_for_user(&input),
                    current_stage
                );
                emit_village_split_event(
                    &mut events,
                    log_view,
                    started,
                    current_stage,
                    "failed",
                    &failure_message,
                    &progress,
                    Some(file_started),
                    &["pytools-core"],
                    None,
                    Some("village_split_stage_error"),
                );
                items.push(serde_json::json!({
                    "input": input,
                    "status": "failed",
                    "error_stage": current_stage,
                    "error": format!("{error:#}"),
                    "elapsed_ms": file_started.elapsed().as_millis(),
                }));
            }
        }
    }
    let status = if failed_files == 0 {
        "succeeded"
    } else if completed_files == 0 {
        "failed"
    } else {
        "partial"
    };
    let mut result = serde_json::json!({
        "feature_id": "PT-7-1",
        "files": total,
        "planned_files": planned_files,
        "completed_files": completed_files,
        "committed_files": committed_files,
        "failed_files": failed_files,
        "matched_sheets": matched_sheets,
        "changed_cells": changed_cells,
        "created_copies": items.iter().filter(|item| item["status"] == "copy_created").count(),
        "mode": if create_copy { "create_copy" } else if in_place.is_some() { "in_place" } else if commit_target { "commit_target" } else if capabilities.is_some() { "candidate" } else { "plan_only" },
        "status": status,
        "plan_dir": plan_dir,
        "elapsed_ms": started.elapsed().as_millis(),
        "items": items,
    });
    let unified_summary =
        attach_legacy_batch_diagnostics(&mut unified_tracker, &mut result, &log_dir, log_view);
    let run_log_path = if log_view == "off" {
        None
    } else {
        let path = log_dir.join(format!("PT-7-1_{run_id}.run.json"));
        let safe_items = result["items"]
            .as_array()
            .into_iter()
            .flatten()
            .map(|item| {
                serde_json::json!({
                    "status": item["status"],
                    "error_stage": item.get("error_stage"),
                    "error_type": if item["status"] == "failed" {
                        Some("village_split_stage_error")
                    } else {
                        None
                    },
                    "matched_sheets": item.get("matched_sheets"),
                    "changed_cells": item.get("changed_cells"),
                    "elapsed_ms": item.get("elapsed_ms"),
                })
            })
            .collect::<Vec<_>>();
        let history = serde_json::json!({
            "schema_version": 1,
            "product": "pytools",
            "feature_id": "PT-7-1",
            "function": "village-split",
            "run_id": run_id,
            "started_at": started_at,
            "status": status,
            "total_inputs": total,
            "completed_files": completed_files,
            "failed_files": failed_files,
            "planned_files": planned_files,
            "committed_files": committed_files,
            "matched_sheets": matched_sheets,
            "changed_cells": changed_cells,
            "elapsed_ms": started.elapsed().as_millis(),
            "items": safe_items,
            "events": events,
        });
        match write_json_report_create_new(&path, &history) {
            Ok(()) => {
                result["run_id"] = serde_json::Value::String(run_id.clone());
                result["run_log_path"] =
                    serde_json::Value::String(display_path_for_user(&path).to_string());
                Some(path)
            }
            Err(error) => {
                eprintln!("pytools: warning: could not save PT-7-1 run history: {error:#}");
                None
            }
        }
    };
    if log_view == "user" {
        if let Some(path) = run_log_path {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    pytools_core::run_logging::Level::Info,
                    &format!(
                        "运行记录已保存：{}",
                        pytools_core::run_logging::path_for_log_view(&path, false)
                    )
                )
            );
        }
    }
    if !interactive || log_view == "developer" {
        println!("{}", serde_json::to_string_pretty(&result)?);
    } else if log_view == "user" {
        if let Some(summary) = unified_summary.as_ref() {
            let output_paths = legacy_report_output_paths(&result);
            emit_unified_terminal_summary_with_outputs(summary, &output_paths);
        } else {
            println!("{}", pytools_core::run_logging::terminal_line(if failed_files > 0 { pytools_core::run_logging::Level::Warn } else { pytools_core::run_logging::Level::Info }, &format!( "村镇银行数据拆分结束：输入 {total} 个，已拆分 {completed_files} 个，失败 {failed_files} 个；耗时 {:.3} 秒", (started.elapsed().as_millis()) as f64 / 1000.0)));
        }
    }
    // off 视图：静默。
    Ok(())
}

fn emit_village_split_event(
    events: &mut Vec<SummaryRunEvent>,
    log_view: &str,
    run_started: Instant,
    phase: &str,
    status: &str,
    user_message: impl Into<String>,
    progress: &SummaryProgress,
    stage_started: Option<Instant>,
    components: &[&str],
    actual_engine: Option<&str>,
    error_type: Option<&str>,
) {
    if log_view == "off" {
        return;
    }
    let event = SummaryRunEvent {
        sequence: events.len() + 1,
        timestamp: Local::now().to_rfc3339(),
        phase: phase.to_owned(),
        status: status.to_owned(),
        user_message: user_message.into(),
        progress: progress.clone(),
        file_name: None,
        stage_elapsed_ms: stage_started.map(|started| started.elapsed().as_millis()),
        total_elapsed_ms: run_started.elapsed().as_millis(),
        components: components
            .iter()
            .map(|component| (*component).to_owned())
            .collect(),
        actual_engine: actual_engine.map(str::to_owned),
        error_type: error_type.map(str::to_owned),
    };
    if log_view == "developer" {
        if let Ok(encoded) = serde_json::to_string(&event) {
            eprintln!("{encoded}");
        } else {
            eprintln!("{{\"phase\":\"progress_event\",\"status\":\"serialization_failed\"}}");
        }
    } else {
        eprintln!(
            "{}",
            pytools_core::run_logging::terminal_line(
                if event.status == "failed" {
                    pytools_core::run_logging::Level::Error
                } else {
                    pytools_core::run_logging::Level::Info
                },
                &event.user_message
            )
        );
    }
    events.push(event);
}

fn run_remove_foreign_bank_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let input = required_path(&options, "input")?;
    let output = required_path(&options, "output")?;
    let foreign = pytools_core::config::load_foreign_banks(config)?;
    let report = build_remove_foreign_plan(&input, &output, &foreign)?;
    write_optional_plan(&options, &report.plan)?;
    let execution = execute_optional_plan(&options, report.plan.as_ref())?;
    let mut result = serde_json::to_value(report)?;
    result["execution"] = execution;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn run_remove_foreign_bank_batch(arguments: Vec<String>, interactive: bool) -> Result<()> {
    let (create_copy, arguments) = extract_boolean_flag(arguments, "create-copy")?;
    let options = parse_cli_arguments(arguments)?;
    let in_place = options.values.get("in-place").cloned();
    let config = required_path(&options, "config")?;
    anyhow::ensure!(
        !options.positional.is_empty(),
        "provide at least one source workbook"
    );
    // 就地模式由 execute_optional_plan 读取 --in-place 后直接改写源工作簿；
    // 此处仅保证标志被识别（避免未使用告警）并保留原有创建副本路径。
    let _ = &in_place;
    let capabilities = options.values.get("capabilities");
    let platform = options.values.get("platform");
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let execute = capabilities.is_some() || in_place.is_some();
    anyhow::ensure!(
        !create_copy || !execute,
        "--create-copy cannot be combined with an ExcelModifier capability"
    );
    let view = options.values.get("view").map(String::as_str);
    if let Some(view) = view {
        anyhow::ensure!(
            matches!(view, "user" | "developer" | "off"),
            "--view must be user, developer, or off"
        );
    }
    anyhow::ensure!(
        view.is_some() || !options.values.contains_key("report-output"),
        "--report-output requires --view user or developer"
    );
    let plan_dir = options.values.get("plan-dir").map(PathBuf::from);
    anyhow::ensure!(
        (!execute && !create_copy) || plan_dir.is_none(),
        "--plan-dir is only supported in plan-only mode"
    );
    let output_dir = options
        .values
        .get("output-dir")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("output"));
    let report_output = if view == Some("off") {
        anyhow::ensure!(
            !options.values.contains_key("report-output"),
            "--report-output cannot be used with --view off"
        );
        None
    } else if view.is_some() {
        let path = options
            .values
            .get("report-output")
            .map(PathBuf::from)
            .unwrap_or_else(|| {
                let nonce = SystemTime::now()
                    .duration_since(UNIX_EPOCH)
                    .map(|duration| duration.as_nanos())
                    .unwrap_or_default();
                let log_dir = options
                    .values
                    .get("log-dir")
                    .map(PathBuf::from)
                    .unwrap_or_else(pytools_core::run_logging::default_log_dir);
                log_dir.join(format!(
                    "PT-7-3_{}_{}_{}.json",
                    Local::now().format("%Y%m%dT%H%M%S"),
                    std::process::id(),
                    nonce
                ))
            });
        anyhow::ensure!(
            path.extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| extension.eq_ignore_ascii_case("json")),
            "--report-output must use a .json extension"
        );
        anyhow::ensure!(
            !path.exists(),
            "run history file already exists; refusing to overwrite {}",
            display_path_for_user(&path)
        );
        Some(path)
    } else {
        None
    };

    let run_started = Instant::now();
    let started_at = Local::now().to_rfc3339();
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_nanos())
        .unwrap_or_default();
    let run_id = format!("pytools-pt-7-3-{}-{nonce}", std::process::id());
    let mut unified_tracker = pytools_core::run_logging::RunTracker::with_id(
        "PT-7-3",
        "删除分机构表中的外资行",
        &[],
        run_id.clone(),
    );
    let configuration_started = Instant::now();
    let foreign = pytools_core::config::load_foreign_banks(&config)?;
    let configuration_elapsed_ms = configuration_started.elapsed().as_millis();
    let mut result = serde_json::json!({
        "files": 0,
        "ok_files": 0,
        "branch_sheets": 0,
        "deleted_rows": 0,
        "failed_files": 0,
        "planned_files": 0,
        "planned_branch_sheets": 0,
        "planned_deleted_rows": 0,
        "foreign_empty": foreign.is_empty(),
        "mode": if in_place.is_some() { "in_place" } else if create_copy { "create_copy" } else if execute { "execute" } else { "plan_only" },
        "items": [],
    });

    let mut events = Vec::<SummaryRunEvent>::new();
    let mut progress = SummaryProgress {
        total_inputs: options.positional.len(),
        total_files: options.positional.len(),
        ..SummaryProgress::default()
    };
    if view.is_some() {
        emit_foreign_bank_event(
            &mut events,
            view,
            run_started,
            "configuration_load",
            "completed",
            "外资行配置读取完成。",
            &progress,
            Some(configuration_elapsed_ms),
            &["excel-reader", "calamine"],
            Some("calamine"),
            None,
        );
    }

    if foreign.is_empty() {
        if view.is_some() {
            emit_foreign_bank_event(
                &mut events,
                view,
                run_started,
                "configuration_precheck",
                "skipped",
                "外资行配置为空，未处理工作簿。",
                &progress,
                None,
                &["pytools-core::config"],
                None,
                Some("empty_foreign_bank_mapping"),
            );
        }
    } else {
        result["files"] = serde_json::json!(options.positional.len());
        for (file_index, input) in options.positional.iter().cloned().enumerate() {
            progress.current_file = Some(file_index + 1);
            progress.processed_files = file_index + 1;
            if view.is_some() {
                emit_foreign_bank_event(
                    &mut events,
                    view,
                    run_started,
                    "file_start",
                    "started",
                    format!(
                        "开始检查第 {}/{} 个工作簿。",
                        file_index + 1,
                        progress.total_files
                    ),
                    &progress,
                    None,
                    &[],
                    None,
                    None,
                );
            }
            let started = Instant::now();
            let output = match foreign_bank_copy_path(&input) {
                Ok(path) => path,
                Err(error) => {
                    result["failed_files"] =
                        serde_json::json!(result["failed_files"].as_u64().unwrap_or_default() + 1);
                    result["items"]
                        .as_array_mut()
                        .unwrap()
                        .push(serde_json::json!({
                            "input": input,
                            "status": "failed",
                            "error_stage": "input_validation",
                            "error": format!("{error:#}"),
                            "elapsed_ms": started.elapsed().as_millis(),
                        }));
                    progress.failed_files += 1;
                    if view.is_some() {
                        emit_foreign_bank_event(
                            &mut events,
                            view,
                            run_started,
                            "input_validation",
                            "failed",
                            "工作簿路径不符合处理要求。",
                            &progress,
                            Some(started.elapsed().as_millis()),
                            &[],
                            None,
                            Some("input_validation"),
                        );
                    }
                    continue;
                }
            };
            let extension_is_xlsx = input
                .extension()
                .and_then(|value| value.to_str())
                .is_some_and(|value| value.eq_ignore_ascii_case("xlsx"));
            let extension_is_xls = input
                .extension()
                .and_then(|value| value.to_str())
                .is_some_and(|value| value.eq_ignore_ascii_case("xls"));
            if extension_is_xls {
                // .xls：Calamine 规划删行 → 隔离 .xls 候选经 Excel/WPS COM 删行
                // → 发布 (金融局).xls 副本；原始 .xls 保持不变。
                #[cfg(windows)]
                {
                    let item = (|| -> anyhow::Result<serde_json::Value> {
                        if std::env::var_os("PYTOOLS_CONVERT_ENGINES_DISABLED").is_some() {
                            anyhow::bail!(
                                "已禁用外部引擎（PYTOOLS_CONVERT_ENGINES_DISABLED）；.xls 外资行删除不可用"
                            );
                        }
                        if !create_copy {
                            anyhow::bail!("5-3 的 .xls 路径仅支持生成 (金融局) 副本模式");
                        }
                        let plan_started = Instant::now();
                        // .xls 无法走 ExcelEditPlan（其 profile 检查仅支持
                        // .xlsx），改为 Calamine 直读规划删行，语义与
                        // build_remove_foreign_plan 一致：分机构表按 A 列
                        // 归一化匹配外资行，整段倒序删行。
                        let dump = excel_reader::read_workbook(&input).with_context(|| {
                            format!("读取 .xls 工作簿 {}", display_path_for_user(&input))
                        })?;
                        let mut matched_sheets = 0usize;
                        let mut deletions: Vec<(String, u32, u32)> = Vec::new();
                        let column_name = |mut col: usize| -> String {
                            let mut letters = Vec::new();
                            while col > 0 {
                                letters.push((b'A' + ((col - 1) % 26) as u8) as char);
                                col = (col - 1) / 26;
                            }
                            letters.reverse();
                            letters.into_iter().collect()
                        };
                        for sheet in dump
                            .sheets
                            .iter()
                            .filter(|sheet| sheet.name.contains("分机构"))
                        {
                            matched_sheets += 1;
                            let max_row = sheet
                                .values
                                .keys()
                                .filter_map(|address| {
                                    let digits = address.rfind(|c: char| c.is_ascii_digit())?;
                                    let row: usize = address[digits..].parse().ok()?;
                                    Some(row)
                                })
                                .max()
                                .unwrap_or(0);
                            let mut foreign_rows = Vec::new();
                            for row in 1..=max_row {
                                let mut col = 1usize;
                                let mut cell_text = String::new();
                                loop {
                                    let address = format!("{}{}", column_name(col), row);
                                    match sheet.values.get(&address) {
                                        Some(value) => match value.get("value") {
                                            Some(serde_json::Value::String(text)) => {
                                                cell_text = text.clone();
                                            }
                                            _ => {}
                                        },
                                        None => {}
                                    }
                                    if cell_text.is_empty() && col < 3 {
                                        col += 1;
                                        continue;
                                    }
                                    break;
                                }
                                let normalized = cell_text
                                    .chars()
                                    .map(|c| {
                                        if matches!(c, '\r' | '\n' | '\t' | '\u{3000}' | '\u{00a0}')
                                        {
                                            ' '
                                        } else {
                                            c
                                        }
                                    })
                                    .collect::<String>()
                                    .split_whitespace()
                                    .collect::<Vec<_>>()
                                    .join(" ");
                                if foreign.contains(&normalized) {
                                    foreign_rows.push(row);
                                }
                            }
                            foreign_rows.sort_unstable();
                            let mut runs: Vec<(u32, u32)> = Vec::new();
                            let mut iter = foreign_rows.iter().rev().peekable();
                            while let Some(end) = iter.next() {
                                let mut start = *end;
                                while iter.peek().is_some_and(|next| *next + 1 == start) {
                                    start = *iter.next().expect("peeked foreign row");
                                }
                                runs.push((start as u32, *end as u32));
                            }
                            for (first, last) in runs {
                                deletions.push((sheet.name.clone(), first, last - first + 1));
                            }
                        }
                        if deletions.is_empty() {
                            copy_workbook_to_new_target(&input, &output).map(|copied| {
                                serde_json::json!({
                                    "engine": "file_copy",
                                    "execution": "created_unchanged_workbook_copy",
                                    "output": output,
                                    "sha256": copied.sha256_after,
                                })
                            })?;
                            return Ok(serde_json::json!({
                                "input": input,
                                "output": output,
                                "status": "written",
                                "branch_sheets": matched_sheets,
                                "deleted_rows": 0,
                                "execution": {"engine": "file_copy"},
                                "elapsed_ms": started.elapsed().as_millis(),
                            }));
                        }
                        let expected_source_hash = sha256_file(&input)?;
                        let xls_candidate = unique_workbook_candidate_preserving_extension(
                            &input,
                            &output_dir,
                            "foreign_bank_xls",
                        );
                        let exec = (|| -> Result<serde_json::Value> {
                            fs::copy(&input, &xls_candidate).with_context(|| {
                                format!("复制 .xls 候选 {}", display_path_for_user(&xls_candidate))
                            })?;
                            verify_planning_snapshot(&input, Some(expected_source_hash.as_str()))?;
                            let mut com_runtime =
                                pytools_core::conversion_engine::ConversionRuntime::default();
                            let com = com_runtime.delete_rows_com(&xls_candidate, &deletions)?;
                            anyhow::ensure!(
                                com["verified"].as_bool() == Some(true),
                                "Excel/WPS COM 删行未通过校验"
                            );
                            create_target_from_candidate(&xls_candidate, &output)?;
                            Ok(serde_json::json!({
                                "engine": "excel_com",
                                "execution": "created_and_validated_workbook_copy",
                                "output": output,
                                "operation_count": deletions.len(),
                            }))
                        })();
                        let cleanup = remove_regular_candidate(&xls_candidate);
                        let execution = match (exec, cleanup) {
                            (Ok(report), Ok(())) => Ok(report),
                            (Ok(_), Err(error)) => {
                                Err(error.context("clean temporary foreign-bank candidate"))
                            }
                            (Err(error), Ok(())) => Err(error),
                            (Err(error), Err(cleanup_error)) => Err(error.context(format!(
                                "temporary candidate cleanup also failed: {cleanup_error:#}"
                            ))),
                        }?;
                        result["ok_files"] =
                            serde_json::json!(result["ok_files"].as_u64().unwrap_or_default() + 1);
                        result["branch_sheets"] = serde_json::json!(
                            result["branch_sheets"].as_u64().unwrap_or_default()
                                + matched_sheets as u64
                        );
                        result["deleted_rows"] = serde_json::json!(
                            result["deleted_rows"].as_u64().unwrap_or_default()
                                + deletions.len() as u64
                        );
                        emit_foreign_bank_event(
                            &mut events,
                            view,
                            run_started,
                            "engine_execution",
                            "completed",
                            "已通过 Excel/WPS COM 删行并生成 (金融局) .xls 副本",
                            &progress,
                            Some(plan_started.elapsed().as_millis()),
                            &["pytools-core", "excel-com"],
                            Some("excel_com"),
                            None,
                        );
                        return Ok(serde_json::json!({
                            "input": input,
                            "output": output,
                            "status": "written",
                            "branch_sheets": matched_sheets,
                            "deleted_rows": deletions.len(),
                            "execution": execution,
                            "elapsed_ms": started.elapsed().as_millis(),
                        }));
                    })()?;
                    result["items"].as_array_mut().unwrap().push(item);
                    continue;
                }
                #[cfg(not(windows))]
                {
                    anyhow::bail!("5-3 的 .xls 处理需要 Windows Excel/WPS COM");
                }
            }
            if !extension_is_xlsx {
                result["failed_files"] =
                    serde_json::json!(result["failed_files"].as_u64().unwrap_or_default() + 1);
                result["items"]
                    .as_array_mut()
                    .unwrap()
                    .push(serde_json::json!({
                        "input": input,
                        "output": output,
                        "status": "failed",
                        "error_stage": "input_validation",
                        "error": "Rust row-removal candidate currently accepts .xlsx only",
                        "elapsed_ms": started.elapsed().as_millis(),
                    }));
                progress.failed_files += 1;
                if view.is_some() {
                    emit_foreign_bank_event(
                        &mut events,
                        view,
                        run_started,
                        "input_validation",
                        "failed",
                        "输入格式尚未验收；本项只接受 .xlsx。",
                        &progress,
                        Some(started.elapsed().as_millis()),
                        &[],
                        None,
                        Some("unsupported_format"),
                    );
                }
                continue;
            }
            let plan_started = Instant::now();
            let report = match build_remove_foreign_plan(&input, &output, &foreign) {
                Ok(report) => report,
                Err(error) => {
                    result["failed_files"] =
                        serde_json::json!(result["failed_files"].as_u64().unwrap_or_default() + 1);
                    result["items"]
                        .as_array_mut()
                        .unwrap()
                        .push(serde_json::json!({
                            "input": input,
                            "output": output,
                            "status": "failed",
                            "error_stage": "planning",
                            "error": format!("{error:#}"),
                            "elapsed_ms": started.elapsed().as_millis(),
                        }));
                    progress.failed_files += 1;
                    if view.is_some() {
                        emit_foreign_bank_event(
                            &mut events,
                            view,
                            run_started,
                            "plan_build",
                            "failed",
                            "工作簿检查或修改计划生成失败。",
                            &progress,
                            Some(plan_started.elapsed().as_millis()),
                            &["pytools-core::batch_ops", "excel-reader"],
                            None,
                            Some("planning"),
                        );
                    }
                    continue;
                }
            };
            let plan_elapsed_ms = plan_started.elapsed().as_millis();
            let planned_deleted_rows = report
                .plan
                .as_ref()
                .map(|plan| {
                    plan.operations
                        .iter()
                        .map(|operation| match operation {
                            excel_editor::EditOperation::DeleteRows { count, .. } => {
                                *count as usize
                            }
                            _ => 0,
                        })
                        .sum::<usize>()
                })
                .unwrap_or_default();
            result["planned_files"] =
                serde_json::json!(result["planned_files"].as_u64().unwrap_or_default() + 1);
            result["planned_branch_sheets"] = serde_json::json!(
                result["planned_branch_sheets"].as_u64().unwrap_or_default()
                    + report.matched_sheets as u64
            );
            result["planned_deleted_rows"] = serde_json::json!(
                result["planned_deleted_rows"].as_u64().unwrap_or_default()
                    + planned_deleted_rows as u64
            );

            let mut item = serde_json::json!({
                "input": input,
                "output": if in_place.is_some() { &input } else { &output },
                "status": if execute { "pending" } else if report.plan.is_some() { "plan_only" } else { "plan_only_no_changes" },
                "planned_branch_sheets": report.matched_sheets,
                "planned_deleted_rows": planned_deleted_rows,
                "branch_sheets": 0,
                "deleted_rows": 0,
                "plan": report.plan.clone(),
            });

            if !execute && !create_copy && !(plan_dir.is_some() && report.plan.is_some()) {
                progress.completed_files += 1;
            }
            if view.is_some() {
                emit_foreign_bank_event(
                    &mut events,
                    view,
                    run_started,
                    "plan_build",
                    "completed",
                    format!(
                        "已完成检查：{} 个目标工作表，{} 行待删除。",
                        report.matched_sheets, planned_deleted_rows
                    ),
                    &progress,
                    Some(plan_elapsed_ms),
                    &["pytools-core::batch_ops", "excel-reader", "ExcelEditPlan"],
                    None,
                    None,
                );
            }

            if let (Some(plan_dir), Some(plan)) = (plan_dir.as_deref(), report.plan.as_ref()) {
                let plan_path = foreign_bank_plan_path(plan_dir, file_index, &input);
                let persistence_started = Instant::now();
                match write_json_report_create_new(&plan_path, plan) {
                    Ok(()) => {
                        item["plan_path"] = serde_json::json!(plan_path);
                        progress.completed_files += 1;
                        if view.is_some() {
                            emit_foreign_bank_event(
                                &mut events,
                                view,
                                run_started,
                                "plan_persistence",
                                "completed",
                                "修改计划已保存。",
                                &progress,
                                Some(persistence_started.elapsed().as_millis()),
                                &["excel-editor::ExcelEditPlan"],
                                None,
                                None,
                            );
                        }
                    }
                    Err(error) => {
                        result["failed_files"] = serde_json::json!(
                            result["failed_files"].as_u64().unwrap_or_default() + 1
                        );
                        item["status"] = serde_json::json!("failed");
                        item["error_stage"] = serde_json::json!("plan_persistence");
                        item["error"] = serde_json::json!(format!("{error:#}"));
                        item["elapsed_ms"] = serde_json::json!(started.elapsed().as_millis());
                        progress.failed_files += 1;
                        if view.is_some() {
                            emit_foreign_bank_event(
                                &mut events,
                                view,
                                run_started,
                                "plan_persistence",
                                "failed",
                                "修改计划保存失败。",
                                &progress,
                                Some(persistence_started.elapsed().as_millis()),
                                &["excel-editor::ExcelEditPlan"],
                                None,
                                Some("plan_persistence"),
                            );
                        }
                        result["items"].as_array_mut().unwrap().push(item);
                        continue;
                    }
                }
            }

            if create_copy {
                let execution_started = Instant::now();
                let execution: Result<serde_json::Value> = if let Some(plan) = report.plan.as_ref()
                {
                    fs::create_dir_all(&output_dir).with_context(|| {
                        format!(
                            "create temporary report directory {}",
                            display_path_for_user(&output_dir)
                        )
                    })?;
                    let candidate = unique_workbook_candidate(&input, &output_dir, "foreign_bank");
                    let expected_source_hash = sha256_file(&input)?;
                    let result = (|| -> Result<serde_json::Value> {
                        let mut candidate_plan = plan.clone();
                        candidate_plan.output = candidate.clone();
                        pytools_core::inplace_edit::execute_candidate(
                            &candidate_plan,
                            pytools_core::inplace_edit::InPlaceEngine::Openpyxl,
                        )?;
                        verify_planning_snapshot(&input, Some(&expected_source_hash))?;
                        create_target_from_candidate(&candidate, &output)?;
                        Ok(serde_json::json!({
                            "engine": "openpyxl",
                            "execution": "created_and_validated_workbook_copy",
                            "output": output,
                            "operation_count": plan.operations.len()
                        }))
                    })();
                    let cleanup = remove_regular_candidate(&candidate);
                    match (result, cleanup) {
                        (Ok(report), Ok(())) => Ok(report),
                        (Ok(_), Err(error)) => {
                            Err(error.context("clean temporary foreign-bank candidate"))
                        }
                        (Err(error), Ok(())) => Err(error),
                        (Err(error), Err(cleanup_error)) => Err(error.context(format!(
                            "temporary candidate cleanup also failed: {cleanup_error:#}"
                        ))),
                    }
                } else {
                    copy_workbook_to_new_target(&input, &output).map(|report| {
                        serde_json::json!({
                            "engine": "file_copy",
                            "execution": "created_unchanged_workbook_copy",
                            "output": output,
                            "sha256": report.sha256_after
                        })
                    })
                };
                match execution {
                    Ok(execution) => {
                        result["ok_files"] =
                            serde_json::json!(result["ok_files"].as_u64().unwrap_or_default() + 1);
                        result["branch_sheets"] = serde_json::json!(
                            result["branch_sheets"].as_u64().unwrap_or_default()
                                + report.matched_sheets as u64
                        );
                        result["deleted_rows"] = serde_json::json!(
                            result["deleted_rows"].as_u64().unwrap_or_default()
                                + planned_deleted_rows as u64
                        );
                        item["status"] = serde_json::json!("written");
                        item["branch_sheets"] = serde_json::json!(report.matched_sheets);
                        item["deleted_rows"] = serde_json::json!(planned_deleted_rows);
                        let actual_engine = foreign_bank_actual_engine(&execution);
                        item["execution"] = execution;
                        progress.completed_files += 1;
                        if view.is_some() {
                            emit_foreign_bank_event(
                                &mut events,
                                view,
                                run_started,
                                "engine_execution",
                                "completed",
                                "外发结果副本已生成，原工作簿保持不变。",
                                &progress,
                                Some(execution_started.elapsed().as_millis()),
                                &["pytools-core", "openpyxl"],
                                actual_engine.as_deref(),
                                None,
                            );
                        }
                    }
                    Err(error) => {
                        result["failed_files"] = serde_json::json!(
                            result["failed_files"].as_u64().unwrap_or_default() + 1
                        );
                        item["status"] = serde_json::json!("failed");
                        item["error_stage"] = serde_json::json!("execution");
                        item["error"] = serde_json::json!(format!("{error:#}"));
                        progress.failed_files += 1;
                        if view.is_some() {
                            emit_foreign_bank_event(
                                &mut events,
                                view,
                                run_started,
                                "engine_execution",
                                "failed",
                                &format!(
                                    "外发副本生成失败：{}；原因：{error:#}；未发布结果文件。",
                                    display_path_for_user(&input)
                                ),
                                &progress,
                                Some(execution_started.elapsed().as_millis()),
                                &["pytools-core", "openpyxl"],
                                None,
                                Some("execution"),
                            );
                        }
                    }
                }
            } else if !execute {
                item["execution"] = serde_json::json!("plan_only");
            } else {
                let execution_started = Instant::now();
                let execution: Result<serde_json::Value> = if let Some(plan) =
                    item.get("plan").filter(|value| !value.is_null()).cloned()
                {
                    let plan: excel_editor::ExcelEditPlan = serde_json::from_value(plan)?;
                    if in_place.is_some() {
                        execute_optional_plan(&options, Some(&plan))
                    } else {
                        apply_plan_with_capability_file(
                            &plan,
                            capabilities.expect("validated capability path"),
                            platform.expect("validated platform"),
                        )
                        .and_then(|report| {
                            serde_json::to_value(report).context("serialize edit result")
                        })
                    }
                } else if in_place.is_some() {
                    Ok(serde_json::json!({"engine": "openpyxl", "execution": "no_changes"}))
                } else {
                    copy_workbook_to_new_target(&input, &output).and_then(|report| {
                        serde_json::to_value(report).context("serialize copy result")
                    })
                };
                match execution {
                    Ok(execution) => {
                        result["ok_files"] =
                            serde_json::json!(result["ok_files"].as_u64().unwrap_or_default() + 1);
                        result["branch_sheets"] = serde_json::json!(
                            result["branch_sheets"].as_u64().unwrap_or_default()
                                + report.matched_sheets as u64
                        );
                        result["deleted_rows"] = serde_json::json!(
                            result["deleted_rows"].as_u64().unwrap_or_default()
                                + planned_deleted_rows as u64
                        );
                        item["status"] = serde_json::json!("written");
                        item["branch_sheets"] = serde_json::json!(report.matched_sheets);
                        item["deleted_rows"] = serde_json::json!(planned_deleted_rows);
                        let actual_engine = foreign_bank_actual_engine(&execution);
                        item["execution"] = execution;
                        progress.completed_files += 1;
                        if view.is_some() {
                            emit_foreign_bank_event(
                                &mut events,
                                view,
                                run_started,
                                "engine_execution",
                                "completed",
                                if in_place.is_some() {
                                    "已校验并安全更新原工作簿中的外资行。"
                                } else {
                                    "外发副本已生成，原工作簿保持不变。"
                                },
                                &progress,
                                Some(execution_started.elapsed().as_millis()),
                                &["openpyxl"],
                                actual_engine.as_deref(),
                                None,
                            );
                        }
                    }
                    Err(error) => {
                        result["failed_files"] = serde_json::json!(
                            result["failed_files"].as_u64().unwrap_or_default() + 1
                        );
                        item["status"] = serde_json::json!("failed");
                        item["error_stage"] = serde_json::json!("execution");
                        item["error"] = serde_json::json!(format!("{error:#}"));
                        progress.failed_files += 1;
                        if view.is_some() {
                            emit_foreign_bank_event(
                                &mut events,
                                view,
                                run_started,
                                "engine_execution",
                                "failed",
                                &format!(
                                    "外资行删除失败：{}；原因：{error:#}",
                                    display_path_for_user(&input)
                                ),
                                &progress,
                                Some(execution_started.elapsed().as_millis()),
                                &["openpyxl"],
                                None,
                                Some("execution"),
                            );
                        }
                    }
                }
            }
            item["elapsed_ms"] = serde_json::json!(started.elapsed().as_millis());
            result["items"].as_array_mut().unwrap().push(item);
        }
    }

    let total_elapsed_ms = run_started.elapsed().as_millis();
    let unified_summary = if let Some(current_view) = view {
        let diagnostic_dir = options
            .values
            .get("log-dir")
            .map(PathBuf::from)
            .unwrap_or_else(pytools_core::run_logging::default_log_dir);
        attach_legacy_batch_diagnostics(
            &mut unified_tracker,
            &mut result,
            &diagnostic_dir,
            current_view,
        )
    } else {
        None
    };
    let terminal_output_paths = legacy_report_output_paths(&result);
    if let Some(view) = view {
        let human_summary = format!( "外资行删除结束：检查工作簿 {} 个，失败 {} 个；耗时 {total_elapsed_ms_display_seconds:.3} 秒", result["files"], result["failed_files"], total_elapsed_ms_display_seconds = total_elapsed_ms as f64 / 1000.0);
        let user_view = foreign_bank_user_view(&result, total_elapsed_ms, Some(&events));
        if let Some(path) = report_output.as_deref() {
            let history = serde_json::json!({
                "schema_version": 1,
                "product": "pytools",
                "feature_id": "PT-7-3",
                "run_id": &run_id,
                "started_at": &started_at,
                "status": user_view["status"],
                "mode": result["mode"],
                "summary": user_view["summary"],
                "events": &events,
            });
            if let Err(error) = write_json_report_create_new(path, &history) {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        pytools_core::run_logging::Level::Warn,
                        &format!("运行记录写入失败：{error}；业务结果仍保留")
                    )
                );
            }
        }
        let output = match view {
            "developer" => {
                result["elapsed_ms"] = serde_json::json!(total_elapsed_ms);
                result["run_id"] = serde_json::json!(run_id);
                result["started_at"] = serde_json::json!(started_at);
                result["events"] = serde_json::to_value(&events)?;
                if let Some(path) = report_output {
                    result["history_report"] = serde_json::json!(path);
                }
                result
            }
            "user" => {
                let mut output = user_view;
                output["run_id"] = serde_json::json!(run_id);
                output["started_at"] = serde_json::json!(started_at);
                if let Some(path) = report_output {
                    output["history_report"] = serde_json::json!(path);
                }
                output
            }
            "off" => foreign_bank_user_view(&result, total_elapsed_ms, None),
            _ => unreachable!("view was validated"),
        };
        if interactive && view == "user" {
            if let Some(summary) = unified_summary.as_ref() {
                emit_unified_terminal_summary_with_outputs(summary, &terminal_output_paths);
            } else {
                println!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        if output["status"] == "completed_with_errors" {
                            pytools_core::run_logging::Level::Warn
                        } else {
                            pytools_core::run_logging::Level::Info
                        },
                        &human_summary
                    )
                );
            }
        } else if interactive && view == "off" {
            println!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    if output["status"] == "completed_with_errors" {
                        pytools_core::run_logging::Level::Warn
                    } else {
                        pytools_core::run_logging::Level::Info
                    },
                    &human_summary
                )
            );
        } else {
            println!("{}", serde_json::to_string_pretty(&output)?);
        }
    } else {
        println!("{}", serde_json::to_string_pretty(&result)?);
    }
    Ok(())
}

fn emit_foreign_bank_event(
    events: &mut Vec<SummaryRunEvent>,
    view: Option<&str>,
    run_started: Instant,
    phase: &str,
    status: &str,
    user_message: impl Into<String>,
    progress: &SummaryProgress,
    stage_elapsed_ms: Option<u128>,
    components: &[&str],
    actual_engine: Option<&str>,
    error_type: Option<&str>,
) {
    let Some(view) = view.filter(|view| *view != "off") else {
        return;
    };
    let event = SummaryRunEvent {
        sequence: events.len() + 1,
        timestamp: Local::now().to_rfc3339(),
        phase: phase.to_owned(),
        status: status.to_owned(),
        user_message: user_message.into(),
        progress: progress.clone(),
        file_name: None,
        stage_elapsed_ms,
        total_elapsed_ms: run_started.elapsed().as_millis(),
        components: components.iter().map(|value| (*value).to_owned()).collect(),
        actual_engine: actual_engine.map(str::to_owned),
        error_type: error_type.map(str::to_owned),
    };
    if view == "developer" {
        if let Ok(encoded) = serde_json::to_string(&event) {
            eprintln!("{encoded}");
        }
    } else {
        let message = format!(
            "{}（进度 {}/{}, 已完成 {}，失败 {}）",
            event.user_message,
            event.progress.processed_files,
            event.progress.total_files,
            event.progress.completed_files,
            event.progress.failed_files
        );
        eprintln!(
            "{}",
            pytools_core::run_logging::terminal_line(
                if event.status == "failed" {
                    pytools_core::run_logging::Level::Error
                } else {
                    pytools_core::run_logging::Level::Info
                },
                &message
            )
        );
    }
    events.push(event);
}

fn foreign_bank_plan_path(plan_dir: &Path, file_index: usize, input: &Path) -> PathBuf {
    let stem = input
        .file_stem()
        .map(|stem| stem.to_string_lossy())
        .unwrap_or_default();
    let safe_stem = stem
        .chars()
        .map(|character| {
            if character.is_ascii_alphanumeric() || matches!(character, '-' | '_') {
                character
            } else {
                '_'
            }
        })
        .collect::<String>();
    plan_dir.join(format!("{:03}_{safe_stem}.json", file_index + 1))
}

fn foreign_bank_actual_engine(execution: &serde_json::Value) -> Option<String> {
    execution
        .get("engine")
        .and_then(serde_json::Value::as_str)
        .map(str::to_owned)
        .or_else(|| {
            (execution
                .get("operation")
                .and_then(serde_json::Value::as_str)
                == Some("copied"))
            .then(|| "filesystem_copy".to_owned())
        })
}

fn foreign_bank_user_view(
    result: &serde_json::Value,
    elapsed_ms: u128,
    events: Option<&[SummaryRunEvent]>,
) -> serde_json::Value {
    let is_plan_only = result["mode"] == "plan_only";
    let foreign_empty = result["foreign_empty"].as_bool().unwrap_or(false);
    let failed_files = result["failed_files"].as_u64().unwrap_or_default();
    let status = if foreign_empty {
        "not_run"
    } else if failed_files > 0 {
        "completed_with_errors"
    } else {
        "completed"
    };
    let completed_files = if is_plan_only {
        serde_json::json!(result["items"]
            .as_array()
            .map(|items| {
                items
                    .iter()
                    .filter(|item| {
                        matches!(
                            item["status"].as_str(),
                            Some("plan_only" | "plan_only_no_changes")
                        )
                    })
                    .count()
            })
            .unwrap_or_default())
    } else {
        result["ok_files"].clone()
    };
    let mut output = serde_json::json!({
        "schema_version": 1,
        "product": "pytools",
        "feature_id": "PT-7-3",
        "status": status,
        "mode": if is_plan_only { "plan_only" } else { "execute_copy" },
        "summary": {
            "total_files": result["files"],
            "completed_files": completed_files,
            "failed_files": result["failed_files"],
            "planned_files": result["planned_files"],
            "planned_branch_sheets": result["planned_branch_sheets"],
            "planned_deleted_rows": result["planned_deleted_rows"],
            "branch_sheets": result["branch_sheets"],
            "deleted_rows": result["deleted_rows"],
            "foreign_empty": foreign_empty,
            "elapsed_ms": elapsed_ms
        }
    });
    let issues = result["items"]
        .as_array()
        .into_iter()
        .flatten()
        .filter(|item| item["status"] == "failed")
        .map(|item| {
            let input = item["input"].as_str().unwrap_or_default();
            serde_json::json!({
                "file_name": Path::new(input)
                    .file_name()
                    .map(|name| name.to_string_lossy().into_owned()),
                "stage": item.get("error_stage"),
                "error": item.get("error"),
            })
        })
        .collect::<Vec<_>>();
    if !issues.is_empty() {
        output["issues"] = serde_json::Value::Array(issues);
    }
    if let Some(events) = events {
        output["events"] = serde_json::Value::Array(
            events
                .iter()
                .map(|event| {
                    serde_json::json!({
                        "sequence": event.sequence,
                        "timestamp": event.timestamp,
                        "phase": foreign_bank_user_phase(&event.phase),
                        "status": foreign_bank_user_status(&event.status),
                        "message": event.user_message,
                        "progress": event.progress,
                        "elapsed_ms": event.stage_elapsed_ms,
                        "total_elapsed_ms": event.total_elapsed_ms
                    })
                })
                .collect(),
        );
    }
    output
}

fn foreign_bank_user_phase(phase: &str) -> &str {
    match phase {
        "configuration_load" => "配置读取",
        "configuration_precheck" => "配置检查",
        "file_start" | "input_validation" => "输入检查",
        "plan_build" => "修改计划",
        "plan_persistence" => "计划保存",
        "engine_execution" => "副本生成",
        other => other,
    }
}

fn foreign_bank_user_status(status: &str) -> &str {
    match status {
        "started" => "进行中",
        "completed" => "完成",
        "failed" => "失败",
        "skipped" => "跳过",
        other => other,
    }
}

fn foreign_bank_copy_path(input: &Path) -> Result<PathBuf> {
    let stem = input.file_stem().with_context(|| {
        format!(
            "source workbook has no file stem: {}",
            display_path_for_user(&input)
        )
    })?;
    let extension = input.extension().with_context(|| {
        format!(
            "source workbook has no extension: {}",
            display_path_for_user(&input)
        )
    })?;
    let mut file_name = stem.to_os_string();
    file_name.push("(金融局).");
    file_name.push(extension);
    Ok(input.with_file_name(file_name))
}

#[derive(Clone, Debug, serde::Serialize)]
struct RuralLoanMenuProgress {
    total_inputs: usize,
    processed_inputs: usize,
    completed_inputs: usize,
    matched_entities: usize,
    failed_regions: usize,
}

#[derive(Clone, Debug, serde::Serialize)]
struct RuralLoanMenuEvent {
    sequence: usize,
    timestamp: String,
    feature_id: String,
    phase: String,
    status: String,
    user_message: String,
    progress: RuralLoanMenuProgress,
    #[serde(skip_serializing_if = "Option::is_none")]
    stage_elapsed_ms: Option<u128>,
    total_elapsed_ms: u128,
    components: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    actual_engine: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    error_type: Option<String>,
}

fn new_rural_loan_menu_event(
    run_started: &Instant,
    feature_id: &str,
    phase: &str,
    status: &str,
    user_message: impl Into<String>,
    progress: &RuralLoanMenuProgress,
    stage_elapsed_ms: Option<u128>,
    components: &[&str],
    actual_engine: Option<&str>,
    error_type: Option<&str>,
) -> RuralLoanMenuEvent {
    RuralLoanMenuEvent {
        sequence: 0,
        timestamp: Local::now().to_rfc3339(),
        feature_id: feature_id.to_owned(),
        phase: phase.to_owned(),
        status: status.to_owned(),
        user_message: user_message.into(),
        progress: progress.clone(),
        stage_elapsed_ms,
        total_elapsed_ms: run_started.elapsed().as_millis(),
        components: components
            .iter()
            .map(|component| (*component).to_owned())
            .collect(),
        actual_engine: actual_engine.map(str::to_owned),
        error_type: error_type.map(str::to_owned),
    }
}

fn emit_rural_loan_menu_event(
    log_view: &str,
    events: &mut Vec<RuralLoanMenuEvent>,
    mut event: RuralLoanMenuEvent,
) {
    if log_view == "off" {
        return;
    }
    event.sequence = events.len() + 1;
    match log_view {
        "user" => eprintln!(
            "{}",
            pytools_core::run_logging::terminal_line(
                pytools_core::run_logging::Level::Info,
                &event.user_message
            )
        ),
        "developer" => match serde_json::to_string(&event) {
            Ok(encoded) => eprintln!("{encoded}"),
            Err(_) => {
                eprintln!("{{\"phase\":\"progress_event\",\"status\":\"serialization_failed\"}}")
            }
        },
        _ => unreachable!("log view validated before event emission"),
    }
    events.push(event);
}

fn rural_loan_output_metrics(summary: &serde_json::Value, quarterly: bool) -> (usize, usize, bool) {
    if quarterly {
        let regions = summary
            .get("regions")
            .and_then(serde_json::Value::as_array)
            .map(Vec::as_slice)
            .unwrap_or_default();
        let failed_regions = regions
            .iter()
            .filter(|region| !region.get("error").is_none_or(serde_json::Value::is_null))
            .count();
        (
            summary
                .get("matched")
                .and_then(serde_json::Value::as_u64)
                .unwrap_or_default() as usize,
            failed_regions,
            summary
                .get("consistent")
                .and_then(serde_json::Value::as_bool)
                .unwrap_or(false),
        )
    } else {
        (
            summary
                .get("matched")
                .and_then(serde_json::Value::as_u64)
                .unwrap_or_default() as usize,
            0,
            summary
                .get("consistent")
                .and_then(serde_json::Value::as_bool)
                .unwrap_or(false),
        )
    }
}

fn rural_loan_paths_equal(left: &Path, right: &Path) -> Result<bool> {
    let left = fs::canonicalize(left).with_context(|| {
        format!(
            "resolve current workbook path {}",
            display_path_for_user(&left)
        )
    })?;
    let right_parent = right.parent().unwrap_or_else(|| Path::new("."));
    let right_parent = fs::canonicalize(right_parent).with_context(|| {
        format!(
            "resolve output directory {}",
            display_path_for_user(&right_parent)
        )
    })?;
    let right_name = right
        .file_name()
        .context("candidate workbook path has no file name")?;
    let right = right_parent.join(right_name);
    #[cfg(windows)]
    {
        Ok(left
            .to_string_lossy()
            .eq_ignore_ascii_case(&right.to_string_lossy()))
    }
    #[cfg(not(windows))]
    {
        Ok(left == right)
    }
}

fn run_rural_loan_menu(
    current: &Path,
    previous: &Path,
    global: &GlobalConfig,
    log_view: &str,
    capabilities: Option<&Path>,
    platform: Option<&str>,
    quarterly: bool,
    engine: Option<SpecialReportEngine>,
) -> Result<()> {
    let (feature_id, function_name, suffix, operation_name) = if quarterly {
        (
            "PT-7-8",
            "quarterly-rural-loan-plan",
            "quarterly_rural_loan",
            "季度涉农贷款修正及分地区核对",
        )
    } else {
        (
            "PT-7-7",
            "rural-loan-plan",
            "rural_loan",
            "涉农贷款比上月修正",
        )
    };
    let current_extension = current
        .extension()
        .and_then(|extension| extension.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    let previous_extension = previous
        .extension()
        .and_then(|extension| extension.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    let legacy_xls_com = !quarterly && current_extension == "xls" && previous_extension == "xls";
    anyhow::ensure!(
        matches!(log_view, "user" | "developer" | "off"),
        "--log-view must be user, developer, or off"
    );
    anyhow::ensure!(
        engine.is_some() || capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    if quarterly {
        anyhow::ensure!(
            current_extension == "xlsx" && previous_extension == "xlsx",
            "{feature_id} current and prior inputs currently support .xlsx only"
        );
    } else {
        anyhow::ensure!(
            current_extension == previous_extension
                && matches!(current_extension.as_str(), "xls" | "xlsx"),
            "{feature_id} 本期与上期工作簿必须均为 .xls 或均为 .xlsx"
        );
        if legacy_xls_com {
            anyhow::ensure!(cfg!(windows), "5-7 的 .xls 路径需要 Windows Excel/WPS COM");
            anyhow::ensure!(engine.is_some(), "5-7 的 .xls 路径需要就地安全提交模式");
            anyhow::ensure!(
                capabilities.is_none() && platform.is_none(),
                "5-7 的 .xls COM 路径不接受 ExcelModifier capability"
            );
        }
    }

    let run_started = Instant::now();
    let started_at = Local::now().to_rfc3339();
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_nanos())
        .unwrap_or_default();
    let run_id = format!(
        "pytools-{}-{}-{nonce}",
        feature_id.to_ascii_lowercase(),
        std::process::id()
    );
    let mut unified = UnifiedRunSession::new(
        &global.log_dir,
        feature_id,
        operation_name,
        log_view,
        Some(run_id.clone()),
    );
    let selected_inputs = [current.to_path_buf(), previous.to_path_buf()];
    unified.set_inputs(&selected_inputs, Some(selected_inputs.len()));
    let mut events = Vec::<RuralLoanMenuEvent>::new();
    let mut progress = RuralLoanMenuProgress {
        total_inputs: 2,
        processed_inputs: 0,
        completed_inputs: 0,
        matched_entities: 0,
        failed_regions: 0,
    };
    emit_rural_loan_menu_event(
        log_view,
        &mut events,
        new_rural_loan_menu_event(
            &run_started,
            feature_id,
            "planning",
            "started",
            format!("正在读取本期和上期工作簿并执行{operation_name}"),
            &progress,
            None,
            if legacy_xls_com {
                &["Python", "pywin32", "Excel/WPS COM"]
            } else {
                &["Rust", "pytools-core", "excel-reader", "calamine"]
            },
            legacy_xls_com.then_some("Excel/WPS COM"),
            None,
        ),
    );

    let mut phase = "planning";
    let mut phase_started = Instant::now();
    let mut safe_summary = None::<serde_json::Value>;
    let mut output_copy = None::<PathBuf>;
    let mut candidate_cleanup_path = None::<PathBuf>;
    let mut target_commit = None::<serde_json::Value>;
    let mut com_runtime = pytools_core::conversion_engine::ConversionRuntime::default();
    let result = (|| -> Result<serde_json::Value> {
        fs::create_dir_all(&global.output_dir).with_context(|| {
            format!(
                "create output directory {}",
                display_path_for_user(&global.output_dir)
            )
        })?;
        let candidate = if legacy_xls_com {
            let candidate_dir = current.parent().unwrap_or(&global.output_dir);
            unique_workbook_candidate_preserving_extension(current, candidate_dir, suffix)
        } else {
            unique_workbook_candidate(current, &global.output_dir, suffix)
        };
        candidate_cleanup_path = Some(candidate.clone());
        anyhow::ensure!(
            !rural_loan_paths_equal(current, &candidate)?,
            "candidate output must not be the current-period workbook"
        );
        let plan_dir =
            unique_menu_action_dir(&global.output_dir.join("plans").join(feature_id), "run");
        if capabilities.is_some() {
            fs::create_dir_all(&plan_dir).with_context(|| {
                format!("create plan directory {}", display_path_for_user(&plan_dir))
            })?;
        }
        let plan_path = plan_dir.join("rural-loan-adjustment.plan.json");
        let expected_target_hash = Some(sha256_file(current)?);
        let expected_previous_hash = legacy_xls_com.then(|| sha256_file(previous)).transpose()?;

        phase_started = Instant::now();
        let (report, summary, com_result) = if legacy_xls_com {
            phase = "com_candidate_copy";
            fs::copy(current, &candidate).with_context(|| {
                format!(
                    "复制 5-7 .xls 候选 {} → {}",
                    display_path_for_user(current),
                    display_path_for_user(&candidate)
                )
            })?;
            let isolated_previous = unique_workbook_candidate_preserving_extension(
                previous,
                candidate.parent().unwrap_or(&global.output_dir),
                "rural_loan_previous_reference",
            );
            fs::copy(previous, &isolated_previous).with_context(|| {
                format!(
                    "为 Excel/WPS COM 隔离上期参考工作簿失败：{}",
                    display_path_for_user(previous)
                )
            })?;
            phase = "com_execution";
            let response = com_runtime.rural_loan_com(&candidate, &isolated_previous);
            let cleanup_previous = remove_regular_candidate(&isolated_previous);
            let response = match response {
                Ok(response) => {
                    cleanup_previous.context("清理上期 COM 参考副本失败")?;
                    response
                }
                Err(error) => match cleanup_previous {
                    Ok(()) => return Err(error.context("stage=pywin32_com_rural_loan")),
                    Err(cleanup_error) => {
                        return Err(error.context(format!(
                        "stage=pywin32_com_rural_loan；上期参考副本清理也失败：{cleanup_error:#}"
                    )))
                    }
                },
            };
            verify_planning_snapshot(current, expected_target_hash.as_deref())
                .context("本期工作簿在 COM 处理期间发生变化")?;
            verify_planning_snapshot(previous, expected_previous_hash.as_deref())
                .context("上期工作簿在 COM 处理期间发生变化")?;
            let summary = response
                .get("summary")
                .cloned()
                .context("5-7 COM worker 未返回业务统计")?;
            let report = pytools_core::batch_ops::PlanBuildReport {
                input: current.to_path_buf(),
                output: candidate.clone(),
                matched_sheets: 2,
                changed_cells: response["changed_cells"].as_u64().unwrap_or_default() as usize,
                plan: None,
                warnings: Vec::new(),
            };
            (report, summary, Some(response))
        } else if quarterly {
            let (report, summary) = build_quarterly_rural_loan_plan(current, previous, &candidate)?;
            (report, serde_json::to_value(summary)?, None)
        } else {
            let (report, summary) = build_rural_loan_plan(current, previous, &candidate)?;
            (report, serde_json::to_value(summary)?, None)
        };
        if let Some(expected_hash) = expected_target_hash.as_deref() {
            verify_planning_snapshot(current, Some(expected_hash))?;
        }
        let (matched_entities, failed_regions, consistent) =
            rural_loan_output_metrics(&summary, quarterly);
        let summary_metrics = serde_json::json!({
            "matched_entities": matched_entities,
            "failed_regions": failed_regions,
            "consistent": consistent
        });
        safe_summary = Some(summary_metrics.clone());
        progress.processed_inputs = 2;
        progress.completed_inputs = 2;
        progress.matched_entities = matched_entities;
        progress.failed_regions = failed_regions;
        unified.file_result(
            1,
            pytools_core::run_logging::FileStatus::Succeeded,
            true,
            "planning",
            "本期工作簿已完成读取和计算",
            None,
            None,
        );
        unified.file_result(
            2,
            pytools_core::run_logging::FileStatus::Succeeded,
            true,
            "planning",
            "上期工作簿已完成读取和计算",
            None,
            None,
        );
        let plan_elapsed_ms = phase_started.elapsed().as_millis();
        let plan_file = if capabilities.is_some() {
            if let Some(plan) = report.plan.as_ref() {
                write_json_report_create_new(&plan_path, plan)?;
                Some(plan_path.clone())
            } else {
                None
            }
        } else {
            None
        };
        emit_rural_loan_menu_event(
            log_view,
            &mut events,
            new_rural_loan_menu_event(
                &run_started,
                feature_id,
                if legacy_xls_com {
                    "com_execution"
                } else {
                    "planning"
                },
                "completed",
                if legacy_xls_com {
                    format!(
                        "Excel/WPS COM 已完成 .xls 业务处理：匹配机构 {} 个；{}",
                        matched_entities,
                        if consistent {
                            "分地区总数一致"
                        } else {
                            "分地区总数不一致"
                        }
                    )
                } else if quarterly {
                    format!(
                        "季度计算完成：地区 {} 个、匹配机构 {} 个；{}",
                        summary["regions"].as_array().map_or(0, Vec::len),
                        matched_entities,
                        if consistent {
                            "汇总一致"
                        } else {
                            "存在差异或地区错误"
                        }
                    )
                } else {
                    format!(
                        "计算完成：匹配机构 {} 个；{}",
                        matched_entities,
                        if consistent {
                            "分地区总数一致"
                        } else {
                            "分地区总数不一致"
                        }
                    )
                },
                &progress,
                Some(plan_elapsed_ms),
                if legacy_xls_com {
                    &["Python", "pywin32", "Excel/WPS COM"]
                } else {
                    &["Rust", "pytools-core", "excel-reader", "calamine"]
                },
                if legacy_xls_com {
                    com_result
                        .as_ref()
                        .and_then(|response| response["application"].as_str())
                        .or(Some("Excel/WPS COM"))
                } else {
                    Some("calamine")
                },
                None,
            ),
        );

        let mut execution = serde_json::json!("no_changes");
        if legacy_xls_com {
            let response = com_result.as_ref().context("5-7 COM 结果缺失")?;
            if report.changed_cells > 0 {
                phase = "safe_commit";
                phase_started = Instant::now();
                emit_rural_loan_menu_event(
                    log_view,
                    &mut events,
                    new_rural_loan_menu_event(
                        &run_started,
                        feature_id,
                        "safe_commit",
                        "started",
                        "正在校验源文件快照并备份后安全替换 .xls 本期工作簿",
                        &progress,
                        None,
                        &["Rust", "pytools-core", "file-commit"],
                        None,
                        None,
                    ),
                );
                let expected_hash = expected_target_hash
                    .as_deref()
                    .context("missing current-period planning snapshot")?;
                verify_planning_snapshot(current, Some(expected_hash))?;
                let commit = replace_target_from_candidate(&candidate, current, expected_hash)
                    .context("stage=target_commit")?;
                target_commit = Some(serde_json::to_value(commit)?);
                execution = serde_json::json!({
                    "mode": "in_place",
                    "engine": "pywin32_com",
                    "application": response["application"],
                    "changed_cells": report.changed_cells,
                });
                emit_rural_loan_menu_event(
                    log_view,
                    &mut events,
                    new_rural_loan_menu_event(
                        &run_started,
                        feature_id,
                        "safe_commit",
                        "completed",
                        "COM 候选已通过快照校验，已备份并安全替换本期 .xls 工作簿",
                        &progress,
                        Some(phase_started.elapsed().as_millis()),
                        &["Rust", "pytools-core", "file-commit"],
                        Some("pywin32 Excel/WPS COM"),
                        None,
                    ),
                );
            }
        } else if let Some(plan) = report.plan.as_ref() {
            if engine.is_some() {
                // 就地修改：直接改写本期工作簿（候选校验 + 快照替换）。
                phase = "in_place_execution";
                phase_started = Instant::now();
                let in_place_engine = pytools_core::inplace_edit::InPlaceEngine::Openpyxl;
                emit_rural_loan_menu_event(
                    log_view,
                    &mut events,
                    new_rural_loan_menu_event(
                        &run_started,
                        feature_id,
                        "in_place_execution",
                        "started",
                        format!("正在使用 {} 就地修改本期工作簿", in_place_engine.label()),
                        &progress,
                        None,
                        &["Rust", "pytools-core", "in-place"],
                        Some(in_place_engine.label()),
                        None,
                    ),
                );
                let outcome = pytools_core::inplace_edit::execute_in_place(plan, in_place_engine)
                    .context("stage=in_place_execution")?;
                execution = serde_json::json!({
                    "mode": "in_place",
                    "engine": in_place_engine.label(),
                    "execution": outcome.execution,
                });
                target_commit = Some(serde_json::to_value(&outcome.commit)?);
                emit_rural_loan_menu_event(
                    log_view,
                    &mut events,
                    new_rural_loan_menu_event(
                        &run_started,
                        feature_id,
                        "in_place_execution",
                        "completed",
                        "本期工作簿已就地改写并通过校验",
                        &progress,
                        Some(phase_started.elapsed().as_millis()),
                        &["Rust", "pytools-core", "in-place"],
                        Some(in_place_engine.label()),
                        None,
                    ),
                );
            } else {
                match (capabilities, platform) {
                    (Some(capabilities), Some(platform)) => {
                        phase = "engine_execution";
                        phase_started = Instant::now();
                        emit_rural_loan_menu_event(
                            log_view,
                            &mut events,
                            new_rural_loan_menu_event(
                                &run_started,
                                feature_id,
                                "engine_execution",
                                "started",
                                "正在通过 ExcelModifier 生成并校验候选工作簿",
                                &progress,
                                None,
                                &["Rust", "pytools-core", "excel-editor", "ExcelModifier"],
                                None,
                                None,
                            ),
                        );
                        let edit_result =
                            apply_plan_with_capability_file(plan, capabilities, platform)
                                .context("stage=excel_modifier")?;
                        execution = serde_json::to_value(&edit_result)?;
                        let actual_engine = edit_result.engine.to_string();
                        emit_rural_loan_menu_event(
                            log_view,
                            &mut events,
                            new_rural_loan_menu_event(
                                &run_started,
                                feature_id,
                                "engine_execution",
                                "completed",
                                "候选工作簿已生成并通过 ExcelModifier 校验",
                                &progress,
                                Some(phase_started.elapsed().as_millis()),
                                &["Rust", "pytools-core", "excel-editor", "ExcelModifier"],
                                Some(&actual_engine),
                                None,
                            ),
                        );

                        phase = "safe_commit";
                        phase_started = Instant::now();
                        emit_rural_loan_menu_event(
                            log_view,
                            &mut events,
                            new_rural_loan_menu_event(
                                &run_started,
                                feature_id,
                                "safe_commit",
                                "started",
                                "正在按计划快照安全更新本期工作簿",
                                &progress,
                                None,
                                &["Rust", "pytools-core"],
                                None,
                                None,
                            ),
                        );
                        let expected_hash = expected_target_hash
                            .as_deref()
                            .context("missing current-period planning snapshot")?;
                        let commit =
                            replace_target_from_candidate(&candidate, current, expected_hash)
                                .context("stage=target_commit")?;
                        target_commit = Some(serde_json::to_value(commit)?);
                        emit_rural_loan_menu_event(
                            log_view,
                            &mut events,
                            new_rural_loan_menu_event(
                                &run_started,
                                feature_id,
                                "safe_commit",
                                "completed",
                                "本期工作簿已安全更新，目标快照校验通过",
                                &progress,
                                Some(phase_started.elapsed().as_millis()),
                                &["Rust", "pytools-core"],
                                None,
                                None,
                            ),
                        );
                    }
                    _ => {
                        phase = "engine_execution";
                        phase_started = Instant::now();
                        let copy_path =
                            specialized_report_copy_path(current, "PyTools_专用报表处理结果")?;
                        emit_rural_loan_menu_event(
                            log_view,
                            &mut events,
                            new_rural_loan_menu_event(
                                &run_started,
                                feature_id,
                                "engine_execution",
                                "started",
                                "正在使用 Rust 新建并校验专用报表结果工作簿",
                                &progress,
                                None,
                                &["Rust", "pytools-core", "excel-writer", "rust_xlsxwriter"],
                                None,
                                None,
                            ),
                        );
                        pytools_core::institution_normalize::create_rust_new_workbook_from_cell_plan(
                        current, &candidate, plan,
                    )
                    .context("stage=rust_new_workbook")?;
                        verify_planning_snapshot(current, expected_target_hash.as_deref())
                            .context("stage=source_stability_check")?;
                        create_target_from_candidate(&candidate, &copy_path).with_context(
                            || {
                                format!(
                                    "publish result copy to {}",
                                    display_path_for_user(&copy_path)
                                )
                            },
                        )?;
                        output_copy = Some(copy_path.clone());
                        execution = serde_json::json!({
                            "engine": "rust_xlsxwriter",
                            "execution": "created_and_validated_workbook_copy",
                            "output": copy_path,
                            "operation_count": plan.operations.len()
                        });
                        emit_rural_loan_menu_event(
                            log_view,
                            &mut events,
                            new_rural_loan_menu_event(
                                &run_started,
                                feature_id,
                                "engine_execution",
                                "completed",
                                format!("结果工作簿已保存，本期源文件未修改"),
                                &progress,
                                Some(phase_started.elapsed().as_millis()),
                                &["Rust", "pytools-core", "excel-writer", "rust_xlsxwriter"],
                                Some("rust_xlsxwriter"),
                                None,
                            ),
                        );
                    }
                }
            }
        } else {
            emit_rural_loan_menu_event(
                log_view,
                &mut events,
                new_rural_loan_menu_event(
                    &run_started,
                    feature_id,
                    "engine_execution",
                    "skipped",
                    "没有需要写入的单元格；本期工作簿未修改",
                    &progress,
                    None,
                    &["Rust", "pytools-core"],
                    None,
                    None,
                ),
            );
        }

        emit_rural_loan_menu_event(
            log_view,
            &mut events,
            new_rural_loan_menu_event(
                &run_started,
                feature_id,
                "completed",
                "succeeded",
                format!(
                    "{operation_name}完成，总耗时 {:.3} 秒",
                    (run_started.elapsed().as_millis()) as f64 / 1000.0
                ),
                &progress,
                Some(run_started.elapsed().as_millis()),
                &["Rust", "pytools-core"],
                None,
                None,
            ),
        );

        Ok(serde_json::json!({
            "feature_id": feature_id,
            "function": function_name,
            "matched_sheets": report.matched_sheets,
            "changed_cells": report.changed_cells,
            "warnings": report.warnings,
            "candidate": candidate,
            "output_copy": output_copy,
            "plan_file": plan_file,
            "summary_metrics": summary_metrics,
            "summary": summary,
            "execution": execution,
            "target_commit": target_commit
        }))
    })();

    if let Some(candidate) = candidate_cleanup_path.as_deref() {
        if let Err(error) = remove_regular_candidate(candidate) {
            eprintln!("pytools: warning: could not remove temporary report candidate: {error:#}");
        }
    }

    if result.is_err() && log_view != "off" {
        let error_type = match phase {
            "com_execution" => "excel_com_operation_error",
            "engine_execution" => "excel_modifier_error",
            "safe_commit" => "target_commit_error",
            _ => "rural_loan_planning_error",
        };
        emit_rural_loan_menu_event(
            log_view,
            &mut events,
            new_rural_loan_menu_event(
                &run_started,
                feature_id,
                phase,
                "failed",
                "涉农贷款处理失败；请查看运行记录中的错误阶段",
                &progress,
                Some(phase_started.elapsed().as_millis()),
                &["Rust", "pytools-core"],
                None,
                Some(error_type),
            ),
        );
    }

    for event in &events {
        let level = if event.status == "failed" {
            pytools_core::run_logging::Level::Error
        } else if event.status == "skipped" {
            pytools_core::run_logging::Level::Warn
        } else {
            pytools_core::run_logging::Level::Info
        };
        let details = event
            .error_type
            .as_ref()
            .map(|error_type| serde_json::json!({"error_type": error_type}));
        unified.progress_at(
            level,
            &event.phase,
            &event.phase,
            &event.status,
            &event.user_message,
            event.stage_elapsed_ms,
            event.actual_engine.as_deref(),
            details,
            Some(&event.timestamp),
        );
    }
    if let Err(error) = result.as_ref() {
        unified.operation_failure(
            phase,
            match phase {
                "safe_commit" => "OUTPUT_COMMIT_FAILED",
                "com_execution" => "EXCEL_COM_OPERATION_FAILED",
                "engine_execution" => "EXCEL_EDIT_FAILED",
                _ => "RURAL_LOAN_FAILED",
            },
            "涉农贷款处理失败；请检查运行记录中的失败阶段",
            error,
        );
    }
    if result.is_err() && progress.completed_inputs == 0 {
        unified.mark_inputs_unavailable();
    }
    unified.set_existing_timing(started_at.clone(), run_started.elapsed().as_millis());
    let mut unified_outputs = output_copy.iter().cloned().collect::<Vec<_>>();
    if target_commit.is_some() {
        unified_outputs.push(current.to_path_buf());
    }
    let unified_output_stats =
        (!unified_outputs.is_empty()).then_some(pytools_core::run_logging::OutputStats {
            workbooks_saved: Some(unified_outputs.len()),
            sheets_saved: None,
            saved: Some(true),
        });
    let status_override = if result.is_ok() && progress.failed_regions > 0 {
        Some(pytools_core::run_logging::TaskStatus::PartialSuccess)
    } else {
        None
    };
    let (unified_summary, unified_summary_path) = unified.finish_with_status(
        result.is_err(),
        false,
        result.is_ok() && unified_outputs.is_empty() && progress.matched_entities == 0,
        &unified_outputs,
        unified_output_stats,
        None,
        None,
        status_override,
    );

    let run_log_path = if log_view == "off" {
        None
    } else {
        let mut actual_engines = Vec::<String>::new();
        for engine in events
            .iter()
            .filter_map(|event| event.actual_engine.as_ref())
        {
            if !actual_engines.contains(engine) {
                actual_engines.push(engine.clone());
            }
        }
        let failure = result.as_ref().err().map(|_| {
            serde_json::json!({
                "phase": phase,
                "error_type": events.last().and_then(|event| event.error_type.as_deref()).unwrap_or("rural_loan_error")
            })
        });
        let history = serde_json::json!({
            "schema_version": 1,
            "product": "pytools",
            "feature_id": feature_id,
            "function": function_name,
            "run_id": run_id.clone(),
            "started_at": started_at,
            "status": if result.is_err() { "failed" } else { "succeeded" },
            "total_inputs": 2,
            "completed_inputs": progress.completed_inputs,
            "failed_inputs": if result.is_err() { 1 } else { 0 },
            "summary_metrics": safe_summary,
            "actual_engines": actual_engines,
            "failure": failure,
            "elapsed_ms": run_started.elapsed().as_millis(),
            "events": events
        });
        let path = global
            .log_dir
            .join(format!("{feature_id}_{run_id}.run.json"));
        match write_json_report_create_new(&path, &history) {
            Ok(()) => Some(path),
            Err(error) => {
                eprintln!("pytools: warning: could not save {feature_id} run history: {error:#}");
                None
            }
        }
    };

    match result {
        Ok(mut result) => {
            if log_view != "off" {
                result["run_id"] = serde_json::Value::String(run_id);
                if let Some(path) = run_log_path {
                    result["run_log_path"] =
                        serde_json::Value::String(display_path_for_user(&path).to_string());
                    if log_view == "user" {
                        eprintln!(
                            "{}",
                            pytools_core::run_logging::terminal_line(
                                pytools_core::run_logging::Level::Info,
                                &format!("运行记录已保存：{}", path_for_log_view(&path, log_view))
                            )
                        );
                    }
                }
            }
            if log_view == "user" {
                println!("{}", pytools_core::run_logging::terminal_line(
                    if unified_summary.overall_status == pytools_core::run_logging::TaskStatus::PartialSuccess { pytools_core::run_logging::Level::Warn } else { pytools_core::run_logging::Level::Info },
                    &format!( "{operation_name}结束：匹配机构 {} 个，失败地区 {} 个；运行状态 {:?}，耗时 {:.3} 秒；运行编号 {}{}", progress.matched_entities, progress.failed_regions, unified_summary.overall_status, (unified_summary.total_elapsed_ms) as f64 / 1000.0, unified_summary.run_id, unified_summary_path.as_ref().map(|path| format!("；运行记录 {}", display_path_for_user(&path))).unwrap_or_default()),
                ));
            } else {
                let mut result = result;
                result["run_id"] = serde_json::json!(unified_summary.run_id.clone());
                if let Some(path) = unified_summary_path.as_ref() {
                    result["run_summary_path"] = serde_json::json!(path);
                }
                println!("{}", serde_json::to_string_pretty(&result)?);
            }
            Ok(())
        }
        Err(error) => {
            if log_view == "user" {
                if let Some(path) = run_log_path {
                    eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            pytools_core::run_logging::Level::Info,
                            &format!(
                                "运行记录已保存：{}",
                                pytools_core::run_logging::path_for_log_view(&path, false)
                            )
                        )
                    );
                }
                if let Some(path) = unified_summary_path {
                    eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            pytools_core::run_logging::Level::Info,
                            &format!(
                                "结构化日志已保存：{}",
                                pytools_core::run_logging::path_for_log_view(&path, false)
                            )
                        )
                    );
                }
                Err(anyhow::anyhow!(
                    "涉农贷款处理失败；请查看运行记录中的错误阶段。"
                ))
            } else {
                Err(error)
            }
        }
    }
}

fn run_rural_loan_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let current = required_path(&options, "current")?;
    let previous = required_path(&options, "previous")?;
    let output = required_path(&options, "output")?;
    let (report, summary) = build_rural_loan_plan(&current, &previous, &output)?;
    write_optional_plan(&options, &report.plan)?;
    let execution = execute_optional_plan(&options, report.plan.as_ref())?;
    println!(
        "{}",
        serde_json::to_string_pretty(
            &serde_json::json!({"edit":report,"summary":summary,"execution":execution})
        )?
    );
    Ok(())
}

fn run_quarterly_rural_loan_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let current = required_path(&options, "current")?;
    let previous = required_path(&options, "previous")?;
    let output = required_path(&options, "output")?;
    let (report, summary) = build_quarterly_rural_loan_plan(&current, &previous, &output)?;
    write_optional_plan(&options, &report.plan)?;
    let execution = execute_optional_plan(&options, report.plan.as_ref())?;
    println!(
        "{}",
        serde_json::to_string_pretty(
            &serde_json::json!({"edit":report,"summary":summary,"execution":execution})
        )?
    );
    Ok(())
}

fn run_institution_normalize_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let input = required_path(&options, "input")?;
    let output = required_path(&options, "output")?;
    let profile = excel_editor::inspect_xlsx_profile(&input)?;
    let mapping = pytools_core::config::load_institution_mapping_ordered(config)?;
    let report = pytools_core::institution_normalize::build_institution_normalize_plan(
        &input,
        &output,
        profile,
        "pytools-institution-normalize",
        &mapping,
    )?;
    write_optional_plan(&options, &report.plan)?;
    let execution = execute_optional_plan(&options, report.plan.as_ref())?;
    println!(
        "{}",
        serde_json::to_string_pretty(&serde_json::json!({
            "branch_sheets": report.branch_sheets,
            "mapped_sheets": report.mapped_sheets,
            "mapped_rows": report.mapped_rows,
            "mapping_empty": report.mapping_empty,
            "plan_written": report.plan.is_some(),
            "execution": execution,
        }))?
    );
    Ok(())
}

fn run_institution_normalize_batch(arguments: Vec<String>, interactive: bool) -> Result<()> {
    let (apply, arguments) = extract_boolean_flag(arguments, "apply")?;
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let view = options
        .values
        .get("view")
        .map(String::as_str)
        .unwrap_or("user");
    anyhow::ensure!(
        matches!(view, "user" | "developer" | "off"),
        "--view must be user, developer, or off"
    );
    let plan_dir = options.values.get("plan-dir").map(PathBuf::from);
    let report_output = if view == "off" {
        anyhow::ensure!(
            !options.values.contains_key("report-output"),
            "--report-output cannot be used with --view off"
        );
        None
    } else {
        let path = options
            .values
            .get("report-output")
            .map(PathBuf::from)
            .unwrap_or_else(|| {
                let nonce = std::time::SystemTime::now()
                    .duration_since(std::time::UNIX_EPOCH)
                    .map(|duration| duration.as_nanos())
                    .unwrap_or_default();
                pytools_core::run_logging::default_log_dir().join(format!(
                    "PT-7-2_{}_{}_{}.json",
                    chrono::Local::now().format("%Y%m%dT%H%M%S"),
                    std::process::id(),
                    nonce
                ))
            });
        anyhow::ensure!(
            path.extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| extension.eq_ignore_ascii_case("json")),
            "--report-output must use a .json extension"
        );
        anyhow::ensure!(
            !path.exists(),
            "run history file already exists; refusing to overwrite {}",
            display_path_for_user(&path)
        );
        if let Some(parent) = path
            .parent()
            .filter(|parent| !parent.as_os_str().is_empty())
        {
            fs::create_dir_all(parent).with_context(|| {
                format!(
                    "create history directory {}",
                    display_path_for_user(&parent)
                )
            })?;
        }
        Some(path)
    };
    let capability = options.values.get("capabilities").map(PathBuf::from);
    let platform = options.values.get("platform").cloned();
    anyhow::ensure!(
        capability.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    anyhow::ensure!(
        apply || capability.is_none(),
        "--capabilities and --platform require --apply"
    );
    let in_place = options.values.contains_key("in-place");
    let run_options = pytools_core::institution_normalize::InstitutionNormalizeRunOptions {
        apply,
        capabilities: capability,
        platform,
        plan_dir,
        in_place,
        engine: options.values.get("in-place").cloned(),
    };
    let mut emitted_events = Vec::new();
    let report = pytools_core::institution_normalize::run_institution_normalize_files(
        &config,
        &options.positional,
        run_options,
        |event| {
            if view == "developer" {
                if let Ok(mut value) = serde_json::to_value(event) {
                    redact_diagnostic_fields(&mut value);
                    if let Ok(encoded) = serde_json::to_string(&value) {
                        eprintln!("{encoded}");
                    }
                }
            } else if view == "user" {
                let current = event
                    .progress
                    .current_file
                    .unwrap_or(event.progress.processed_files);
                eprintln!(
                    "{}（进度 {current}/{}, 成功 {}，失败 {}）",
                    event.user_message,
                    event.progress.total_files,
                    event.progress.completed_files,
                    event.progress.failed_files
                );
            }
            emitted_events.push(event.clone());
        },
    )?;
    anyhow::ensure!(
        emitted_events.len() == report.events.len(),
        "internal progress event count mismatch"
    );

    let diagnostic_dir = report_output
        .as_deref()
        .and_then(Path::parent)
        .map(Path::to_path_buf)
        .unwrap_or_else(pytools_core::run_logging::default_log_dir);
    let (unified_path, unified_summary) = if view != "off" {
        write_normalize_unified_log(&report, &options.positional, &diagnostic_dir, view)
    } else {
        (None, None)
    };

    if let Some(report_output) = report_output.as_deref() {
        let mut safe_report = serde_json::to_value(&report)?;
        redact_diagnostic_fields(&mut safe_report);
        if let Err(error) = write_json_report_create_new(report_output, &safe_report) {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    pytools_core::run_logging::Level::Warn,
                    &format!("运行记录写入失败：{error}；业务结果仍保留")
                )
            );
        }
    }
    let output = if view == "developer" {
        let mut value = serde_json::to_value(&report)?;
        redact_diagnostic_fields(&mut value);
        value
    } else {
        institution_normalize_user_view(&report)
    };
    let mut output = output;
    if view == "off" {
        if let Some(object) = output.as_object_mut() {
            object.remove("events");
        }
    }
    if let Some(report_output) = report_output {
        output["history_report"] =
            serde_json::Value::String(display_path_for_user(&report_output).to_string());
    }
    if let Some(path) = unified_path {
        output["run_summary_path"] = serde_json::json!(path.to_string_lossy());
    }
    if interactive && view == "user" {
        if let Some(summary) = unified_summary.as_ref() {
            let outputs = report
                .files
                .iter()
                .filter_map(|file| file.output_path.clone())
                .collect::<Vec<_>>();
            emit_unified_terminal_summary_with_outputs(summary, &outputs);
        } else {
            println!("{}", pytools_core::run_logging::terminal_line(if report.failed_files > 0 { pytools_core::run_logging::Level::Warn } else { pytools_core::run_logging::Level::Info }, &format!( "机构名称规范化结束：选择 {} 个，已处理 {} 个，失败 {} 个；耗时 {:.3} 秒；运行编号 {}", report.total_inputs, report.completed_files, report.failed_files, (report.elapsed_ms) as f64 / 1000.0, report.run_id)));
        }
    } else if interactive && view == "off" {
        println!("{}", pytools_core::run_logging::terminal_line(if report.failed_files > 0 { pytools_core::run_logging::Level::Warn } else { pytools_core::run_logging::Level::Info }, &format!( "机构名称规范化结束：选择 {} 个，已处理 {} 个，失败 {} 个；耗时 {:.3} 秒；运行编号 {}", report.total_inputs, report.completed_files, report.failed_files, (report.elapsed_ms) as f64 / 1000.0, report.run_id)));
    } else {
        println!("{}", serde_json::to_string_pretty(&output)?);
    }
    Ok(())
}

fn redact_diagnostic_fields(value: &mut serde_json::Value) {
    match value {
        serde_json::Value::Object(fields) => {
            for (name, field) in fields {
                let key = name.to_ascii_lowercase();
                if key == "route_diagnostic" {
                    let kind = field
                        .get("kind")
                        .and_then(serde_json::Value::as_str)
                        .map(str::to_owned);
                    *field = kind
                        .map(|kind| serde_json::json!({"kind": kind}))
                        .unwrap_or(serde_json::Value::Null);
                } else if matches!(key.as_str(), "diagnostic" | "password" | "token" | "secret") {
                    *field = serde_json::Value::Null;
                } else {
                    redact_diagnostic_fields(field);
                }
            }
        }
        serde_json::Value::Array(items) => items.iter_mut().for_each(redact_diagnostic_fields),
        _ => {}
    }
}

fn write_normalize_unified_log(
    report: &pytools_core::institution_normalize::InstitutionNormalizeRunReport,
    inputs: &[PathBuf],
    log_dir: &Path,
    log_view: &str,
) -> (
    Option<PathBuf>,
    Option<pytools_core::run_logging::RunSummary>,
) {
    use pytools_core::run_logging::{EventWriter, FileStatus, Level, RunTracker};
    let mut tracker =
        RunTracker::with_id("PT-7-2", "机构名称规范化", inputs, report.run_id.clone());
    tracker.set_existing_timing(report.started_at.clone(), report.elapsed_ms);
    let mut writer =
        match EventWriter::create(&log_dir.join("run-events"), "PT-7-2", &report.run_id) {
            Ok(writer) => Some(writer),
            Err(error) => {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("无法创建诊断日志：{error}")
                    )
                );
                None
            }
        };
    if let Some(writer) = writer.as_mut() {
        let mut start = tracker.event(
            Level::Info,
            "run_started",
            "configuration",
            "started",
            "开始机构名称规范化",
        );
        start.timestamp = report.started_at.clone();
        if let Err(error) = writer.write(&start) {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Warn,
                    &format!("诊断日志写入失败：{error}")
                )
            );
        }
    }
    if report.mapping_empty {
        tracker.skip_unfinished();
    }
    let mut outputs = Vec::new();
    for (index, file) in report.files.iter().enumerate() {
        let position = index + 1;
        if tracker.state(position) != Some(FileStatus::Pending) {
            continue;
        }
        let _ = tracker.transition(position, FileStatus::InProgress);
        let state = match file.status.as_str() {
            "failed" => FileStatus::Failed,
            "planned" => FileStatus::Skipped,
            "applied" | "no_changes" => FileStatus::Succeeded,
            _ => continue,
        };
        let _ = tracker.transition(position, state);
        if let Some(path) = &file.output_path {
            outputs.push(path_for_log_view(path, log_view));
        }
        if report.mode == "in_place" && file.status == "applied" {
            outputs.push(path_for_log_view(&inputs[index], log_view));
        }
        if let Some(writer) = writer.as_mut() {
            let level = if state == FileStatus::Failed {
                Level::Error
            } else if state == FileStatus::Skipped {
                Level::Warn
            } else {
                Level::Info
            };
            let mut event = tracker.event(
                level,
                "file_result",
                file.error_stage.as_deref().unwrap_or("normalization"),
                &file.status,
                &file.user_message,
            );
            event.file_index = Some(position);
            event.file_name = Some(file.file_name.clone());
            event.engine = file.actual_engine.clone();
            event.elapsed_ms = Some(file.timings.total_ms);
            if state == FileStatus::Failed {
                event.error_code = Some("EXCEL_OPERATION_FAILED".to_owned());
            }
            if let Err(error) = writer.write(&event) {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("诊断日志写入失败：{error}")
                    )
                );
                break;
            }
        }
    }
    let mut summary = tracker.finish(
        report.status == "failed",
        false,
        report.mapping_empty,
        outputs,
        diagnostic_path_for_log_view(writer.as_ref(), log_view),
    );
    summary.business_metrics.insert(
        "映射工作表数".to_owned(),
        serde_json::json!(report.mapped_sheets),
    );
    summary
        .business_metrics
        .insert("映射行数".to_owned(), serde_json::json!(report.mapped_rows));
    summary.business_metrics.insert(
        "修改文件数".to_owned(),
        serde_json::json!(report.modified_files),
    );
    if let Some(failure) = &report.failure {
        summary.error_summary = Some(pytools_core::run_logging::ErrorSummary {
            stage: user_friendly_normalize_error_stage(&failure.stage).to_owned(),
            code: "INSTITUTION_NORMALIZE_FAILED".to_owned(),
            message: failure.user_message.clone(),
            file_name: report
                .files
                .iter()
                .find(|file| file.status == "failed")
                .map(|file| file.file_name.clone()),
            sheet_name: None,
        });
    }
    if let Some(writer) = writer.as_mut() {
        if let Err(error) = writer.write(&tracker.event(
            if report.failed_files > 0 {
                Level::Error
            } else {
                Level::Info
            },
            "run_finished",
            "run",
            &format!("{:?}", summary.overall_status),
            "机构名称规范化结束",
        )) {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Warn,
                    &format!("诊断日志写入失败：{error}")
                )
            );
        }
    }
    let path = log_dir
        .join("run-unified")
        .join(format!("PT-7-2_{}.json", report.run_id));
    let summary_path = match write_json_report_create_new(&path, &summary) {
        Ok(()) => Some(path),
        Err(error) => {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Warn,
                    &format!("统一运行记录写入失败：{error}")
                )
            );
            None
        }
    };
    (summary_path, Some(summary))
}

fn institution_normalize_user_view(
    report: &pytools_core::institution_normalize::InstitutionNormalizeRunReport,
) -> serde_json::Value {
    serde_json::json!({
        "schema_version": report.schema_version,
        "product": report.product,
        "feature_id": report.feature_id,
        "run_id": report.run_id,
        "started_at": report.started_at,
        "status": user_friendly_normalize_run_status(&report.status),
        "mode": match report.mode.as_str() {
            "in_place" => "就地修改源工作簿",
            "apply" => "执行修改并另存副本",
            _ => "仅生成计划，未修改文件",
        },
        "failure": report.failure.as_ref().map(|failure| serde_json::json!({
            "stage": user_friendly_normalize_error_stage(&failure.stage),
            "message": failure.user_message
        })),
        "progress": report.events.last().map(|event| &event.progress),
        "summary": {
            "total_inputs": report.total_inputs,
            "completed_files": report.completed_files,
            "failed_files": report.failed_files,
            "modified_files": report.modified_files,
            "branch_sheets": report.branch_sheets,
            "mapped_sheets": report.mapped_sheets,
            "mapped_rows": report.mapped_rows,
            "mapping_empty": report.mapping_empty,
            "configuration_elapsed_ms": report.configuration_elapsed_ms,
            "elapsed_ms": report.elapsed_ms
        },
        "events": report.events.iter().map(|event| serde_json::json!({
            "sequence": event.sequence,
            "timestamp": event.timestamp,
            "phase": user_friendly_normalize_phase(&event.phase),
            "status": user_friendly_normalize_event_status(&event.status),
            "message": event.user_message,
            "progress": event.progress,
            "file_name": event.file_name,
            "elapsed_ms": event.elapsed_ms
        })).collect::<Vec<_>>(),
        "files": report.files.iter().map(|file| serde_json::json!({
            "file_name": file.file_name,
            "status": user_friendly_normalize_status(&file.status),
            "branch_sheets": file.branch_sheets,
            "mapped_sheets": file.mapped_sheets,
            "mapped_rows": file.mapped_rows,
            "operation_count": file.operation_count,
            "plan_saved": file.plan_path.is_some(),
            "output_path": file.output_path,
            "stages": [
                {"name": "文件检查", "elapsed_ms": file.timings.source_fingerprint_ms},
                {"name": "工作簿结构检查", "elapsed_ms": file.timings.profile_inspection_ms},
                {"name": "业务规则与修改计划", "elapsed_ms": file.timings.plan_build_ms},
                {"name": "计划保存", "elapsed_ms": file.timings.plan_persistence_ms},
                {"name": "输入一致性复核", "elapsed_ms": file.timings.source_stability_check_ms},
                {"name": "修改引擎执行", "elapsed_ms": file.timings.engine_execution_ms},
                {"name": "安全提交", "elapsed_ms": file.timings.safe_commit_ms},
                {"name": "文件总耗时", "elapsed_ms": file.timings.total_ms}
            ],
            "error_stage": file.error_stage.as_deref().map(user_friendly_normalize_error_stage),
            "message": file.user_message,
            "cleanup_warning": file.cleanup_warning
        })).collect::<Vec<_>>()
    })
}

fn user_friendly_normalize_phase(phase: &str) -> &str {
    match phase {
        "source_fingerprint" => "文件安全检查",
        "profile_inspection" => "工作簿结构检查",
        "plan_build" => "规则检查与修改计划",
        "plan_persistence" => "保存修改计划",
        "source_stability_check" => "输入一致性复核",
        "engine_execution" => "应用修改",
        "safe_commit" => "验证并提交结果",
        "configuration_read" => "读取映射配置",
        "normalization" => "机构名称规范化",
        "file_processing" => "处理工作簿",
        "task" => "任务汇总",
        _ => "处理步骤",
    }
}

fn user_friendly_normalize_run_status(status: &str) -> &str {
    match status {
        "completed" => "已完成",
        "partial_failure" => "部分文件失败",
        "failed" => "失败",
        _ => "处理中",
    }
}

fn user_friendly_normalize_event_status(status: &str) -> &str {
    match status {
        "started" => "进行中",
        "completed" => "已完成",
        "partial_failure" => "部分文件失败",
        "failed" => "失败",
        "skipped" => "已跳过",
        _ => "处理中",
    }
}

fn user_friendly_normalize_status(status: &str) -> &str {
    match status {
        "planned" => "已生成计划，尚未修改",
        "applied" => "已修改",
        "no_changes" => "无需修改",
        "failed" => "失败",
        other => other,
    }
}

fn user_friendly_normalize_error_stage(stage: &str) -> &str {
    match stage {
        "configuration_read" => "读取映射配置",
        "format_check" => "文件格式检查",
        "input_check" => "输入文件检查",
        "source_fingerprint" | "source_stability_check" => "源文件安全检查",
        "profile_inspection" => "工作簿结构检查",
        "plan_build" => "规则计算与修改计划",
        "plan_persistence" => "修改计划保存",
        "engine_execution" => "工作簿修改",
        "safe_commit" => "安全提交",
        other => other,
    }
}

fn write_json_report_create_new<T: serde::Serialize>(path: &Path, report: &T) -> Result<()> {
    if let Some(parent) = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
    {
        fs::create_dir_all(parent).with_context(|| {
            format!("create report directory {}", display_path_for_user(&parent))
        })?;
    }
    let mut output = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(path)
        .with_context(|| {
            format!(
                "create run report {} without overwriting",
                display_path_for_user(&path)
            )
        })?;
    output
        .write_all(&serde_json::to_vec_pretty(report)?)
        .with_context(|| format!("write run report {}", display_path_for_user(&path)))?;
    output
        .sync_all()
        .with_context(|| format!("sync run report {}", display_path_for_user(&path)))?;
    Ok(())
}

fn run_fx_header_batch(arguments: Vec<String>, interactive: bool) -> Result<()> {
    let (create_copy, arguments) = extract_boolean_flag(arguments, "create-copy")?;
    let (commit_target, arguments) = extract_boolean_flag(arguments, "commit-target")?;
    let options = parse_cli_arguments(arguments)?;
    let in_place = options.values.get("in-place").cloned();
    anyhow::ensure!(
        !options.positional.is_empty(),
        "provide at least one FX header workbook"
    );
    let capabilities = options.values.get("capabilities");
    let platform = options.values.get("platform");
    anyhow::ensure!(
        !create_copy || (!commit_target && capabilities.is_none()),
        "--create-copy cannot be combined with --commit-target or ExcelModifier capability evidence"
    );
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    // 就地模式自带引擎与回读校验，不依赖外部 capability 文件。
    anyhow::ensure!(
        capabilities.is_none() || commit_target || in_place.is_some(),
        "FX header execution requires --commit-target with matching capability evidence"
    );

    let view = options
        .values
        .get("view")
        .map(String::as_str)
        .unwrap_or("user");
    anyhow::ensure!(
        matches!(view, "user" | "developer" | "off"),
        "--view must be user, developer, or off"
    );
    anyhow::ensure!(
        view != "off" || !options.values.contains_key("report-output"),
        "--report-output cannot be used with --view off"
    );

    let output_dir = options
        .values
        .get("output-dir")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("output"));
    let log_dir = options
        .values
        .get("log-dir")
        .map(PathBuf::from)
        .unwrap_or_else(pytools_core::run_logging::default_log_dir);
    let plan_dir = if create_copy || in_place.is_some() || (commit_target && capabilities.is_none())
    {
        None
    } else {
        Some(
            options
                .values
                .get("plan-dir")
                .map(PathBuf::from)
                .unwrap_or_else(|| {
                    unique_menu_action_dir(&output_dir.join("plans").join("PT-7-4"), "run")
                }),
        )
    };
    fs::create_dir_all(&output_dir).with_context(|| {
        format!(
            "create output directory {}",
            display_path_for_user(&output_dir)
        )
    })?;
    if let Some(plan_dir) = plan_dir.as_deref() {
        fs::create_dir_all(plan_dir).with_context(|| {
            format!("create plan directory {}", display_path_for_user(&plan_dir))
        })?;
    }

    let run_started = Instant::now();
    let started_at = Local::now().to_rfc3339();
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_nanos())
        .unwrap_or_default();
    let run_id = format!("pytools-pt-7-4-{}-{nonce}", std::process::id());
    let mut unified_tracker = pytools_core::run_logging::RunTracker::with_id(
        "PT-7-4",
        "修改外汇工作表页眉",
        &[],
        run_id.clone(),
    );
    let report_output = if view == "off" {
        None
    } else {
        let path = options
            .values
            .get("report-output")
            .map(PathBuf::from)
            .unwrap_or_else(|| {
                log_dir.join(format!(
                    "PT-7-4_{}_{}_{}.json",
                    Local::now().format("%Y%m%dT%H%M%S"),
                    std::process::id(),
                    nonce
                ))
            });
        anyhow::ensure!(
            path.extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| extension.eq_ignore_ascii_case("json")),
            "--report-output must use a .json extension"
        );
        anyhow::ensure!(
            !path.exists(),
            "run history file already exists; refusing to overwrite {}",
            display_path_for_user(&path)
        );
        Some(path)
    };

    let mut events = Vec::<SummaryRunEvent>::new();
    let mut progress = SummaryProgress {
        total_inputs: options.positional.len(),
        total_files: options.positional.len(),
        ..SummaryProgress::default()
    };
    let mut items = Vec::with_capacity(options.positional.len());
    let mut completed_files = 0usize;
    let mut failed_files = 0usize;
    let mut planned_files = 0usize;
    let mut committed_files = 0usize;
    let mut created_copies = 0usize;
    let mut fx_sheets = 0usize;
    let mut modified_sheets = 0usize;

    for (index, input) in options.positional.iter().enumerate() {
        progress.current_file = Some(index + 1);
        progress.processed_files = index;
        progress.completed_files = completed_files;
        progress.failed_files = failed_files;
        if view == "user" {
            eprintln!(
                "[处理中] {}/{} 检查外汇页眉",
                index + 1,
                options.positional.len()
            );
        }
        emit_fx_header_event(
            &mut events,
            view,
            run_started,
            "input_validation",
            "started",
            format!(
                "开始检查第 {}/{} 个工作簿的外汇页眉并准备生成副本",
                index + 1,
                options.positional.len()
            ),
            &progress,
            None,
            &["pytools-core"],
            None,
            None,
        );

        let file_started = Instant::now();
        let candidate = unique_workbook_candidate(input, &output_dir, "fx_header");
        let mut current_stage = "input_validation";
        let item_result = (|| -> Result<serde_json::Value> {
            if in_place.is_none() {
                anyhow::ensure!(
                    input
                        .extension()
                        .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
                    "PT-7-4 currently accepts .xlsx only"
                );
            }
            // .xls 就地修改：随包 worker 以 Excel/WPS COM 改页眉（万元→万美元），
            // 语义对齐 PyCode feature_1_4_fx_header_com；候选+快照保护原件。
            if input
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("xls"))
                && in_place.is_some()
            {
                #[cfg(windows)]
                {
                    if std::env::var_os("PYTOOLS_CONVERT_ENGINES_DISABLED").is_some() {
                        anyhow::bail!(
                            "已禁用外部引擎（PYTOOLS_CONVERT_ENGINES_DISABLED）；.xls 页眉修改不可用"
                        );
                    }
                    let expected_target_hash = sha256_file(input)?;
                    let xls_candidate = unique_workbook_candidate_preserving_extension(
                        input,
                        &output_dir,
                        "fx_header_xls",
                    );
                    let result = (|| -> Result<serde_json::Value> {
                        fs::copy(input, &xls_candidate).with_context(|| {
                            format!("复制 .xls 候选 {}", display_path_for_user(&xls_candidate))
                        })?;
                        verify_planning_snapshot(input, Some(expected_target_hash.as_str()))?;
                        let mut com_runtime =
                            pytools_core::conversion_engine::ConversionRuntime::default();
                        let com = com_runtime.fx_header_com(&xls_candidate)?;
                        anyhow::ensure!(
                            com["verified"].as_bool() == Some(true),
                            "Excel/WPS COM 页眉修改未通过校验"
                        );
                        let commit = replace_target_from_candidate(
                            &xls_candidate,
                            input,
                            expected_target_hash.as_str(),
                        )?;
                        Ok(serde_json::json!({
                            "input": input,
                            "output_copy": null,
                            "plan_file": null,
                            "status": "committed",
                            "fx_sheets": com["fx_sheets"].as_u64().unwrap_or_default(),
                            "modified_sheets": com["modified_sheets"].as_u64().unwrap_or_default(),
                            "execution": {
                                "engine": "excel_com",
                                "application": com["application"],
                                "fx_sheets": com["fx_sheets"],
                                "modified_sheets": com["modified_sheets"],
                            },
                            "target_commit": serde_json::to_value(commit)?,
                            "elapsed_ms": file_started.elapsed().as_millis(),
                        }))
                    })();
                    let legacy_result = match (result, remove_regular_candidate(&xls_candidate)) {
                        (Ok(item), Ok(())) => Ok(item),
                        (Ok(item), Err(error)) => {
                            Err(error.context("clean .xls fx-header candidate"))
                        }
                        (Err(error), Ok(())) => Err(error),
                        (Err(error), Err(cleanup_error)) => Err(error).context(format!(
                            "temporary candidate cleanup also failed: {cleanup_error:#}"
                        )),
                    };
                    return legacy_result;
                }
                #[cfg(not(windows))]
                {
                    anyhow::bail!("5-4 的 .xls 页眉修改需要 Windows Excel/WPS COM");
                }
            }
            let metadata = fs::symlink_metadata(input).with_context(|| {
                format!("inspect input workbook {}", display_path_for_user(&input))
            })?;
            anyhow::ensure!(
                metadata.is_file() && !metadata.file_type().is_symlink(),
                "input workbook must be a regular, non-symlink file: {}",
                display_path_for_user(&input)
            );
            // 就地模式同样需要规划期快照，用于提交前校验源未被改动。
            let expected_target_hash = if commit_target || create_copy || in_place.is_some() {
                Some(sha256_file(input)?)
            } else {
                None
            };
            current_stage = "profile_inspection";
            let profile_started = Instant::now();
            let profile = excel_editor::inspect_xlsx_profile(input)?;
            emit_fx_header_event(
                &mut events,
                view,
                run_started,
                "profile_inspection",
                "completed",
                "已检查工作簿格式与对象特征",
                &progress,
                Some(profile_started),
                &["excel-editor", "excel-reader"],
                Some("inspect_xlsx_profile"),
                None,
            );

            current_stage = "plan_build";
            let plan_started = Instant::now();
            let inspection = pytools_core::fx_header::build_fx_header_plan(
                input,
                &candidate,
                profile,
                format!("{run_id}-{:04}", index + 1),
            )?;
            if let Some(expected_hash) = expected_target_hash.as_deref() {
                verify_planning_snapshot(input, Some(expected_hash))?;
            }
            emit_fx_header_event(
                &mut events,
                view,
                run_started,
                "plan_build",
                if inspection.plan.is_some() {
                    "completed"
                } else {
                    "no_changes"
                },
                if inspection.plan.is_some() {
                    "已完成页眉检查并生成修改计划"
                } else {
                    "没有需要修改的外汇页眉"
                },
                &progress,
                Some(plan_started),
                &["pytools-core", "excel-reader"],
                Some("read_header_inventory"),
                None,
            );

            let plan_file = if plan_dir.is_some() {
                if let Some(plan) = inspection.plan.as_ref() {
                    current_stage = "plan_persistence";
                    let path = plan_dir
                        .as_deref()
                        .context("plan directory missing in plan mode")?
                        .join(format!("target-{:04}.plan.json", index + 1));
                    let persistence_started = Instant::now();
                    write_json_report_create_new(&path, plan)?;
                    emit_fx_header_event(
                        &mut events,
                        view,
                        run_started,
                        "plan_persistence",
                        "completed",
                        "已保存页眉修改计划",
                        &progress,
                        Some(persistence_started),
                        &["pytools-core"],
                        None,
                        None,
                    );
                    Some(path)
                } else {
                    None
                }
            } else {
                None
            };

            let execution = if let Some(plan) = inspection.plan.as_ref() {
                current_stage = "engine_execution";
                let execution_started = Instant::now();
                let execution = if create_copy || (commit_target && capabilities.is_none()) {
                    create_fx_header_copy(
                        input,
                        &candidate,
                        plan,
                        expected_target_hash
                            .as_deref()
                            .context("source snapshot missing before header modification")?,
                        create_copy,
                    )?
                } else if in_place.is_some() {
                    current_stage = "in_place_execution";
                    execute_fx_header_in_place(
                        input,
                        &candidate,
                        plan,
                        expected_target_hash
                            .as_deref()
                            .context("source snapshot missing before header modification")?,
                        &options,
                    )?
                } else {
                    execute_optional_plan(&options, Some(plan))?
                };
                let engine = execution.get("engine").and_then(serde_json::Value::as_str);
                emit_fx_header_event(
                    &mut events,
                    view,
                    run_started,
                    "engine_execution",
                    "completed",
                    if create_copy {
                        "已通过 OOXML 页眉补丁校验并生成独立副本"
                    } else if in_place.is_some() {
                        if engine == Some("ooxml_patch") {
                            "已通过 OOXML 页眉定点补丁校验候选"
                        } else {
                            "OOXML 定点补丁不适用，已由 Openpyxl 保护回退并校验候选"
                        }
                    } else {
                        "已通过 ExcelModifier 修改并验证候选工作簿"
                    },
                    &progress,
                    Some(execution_started),
                    if engine == Some("ooxml_patch") {
                        &["pytools-core", "ooxml-patch"][..]
                    } else {
                        &["pytools-core", "excel-editor", "openpyxl_worker"][..]
                    },
                    engine,
                    None,
                );
                execution
            } else {
                serde_json::Value::String("no_changes".to_owned())
            };

            let commit = if in_place.is_some() && inspection.plan.is_some() {
                let commit = execution
                    .get("commit")
                    .cloned()
                    .context("in-place FX header execution returned no safe-commit report")?;
                emit_fx_header_event(
                    &mut events,
                    view,
                    run_started,
                    "safe_commit",
                    "completed",
                    "页眉修改候选已校验并安全替换原工作簿",
                    &progress,
                    Some(file_started),
                    &["pytools-core", "file-commit"],
                    execution.get("engine").and_then(serde_json::Value::as_str),
                    None,
                );
                Some(serde_json::to_value(commit)?)
            } else if commit_target && inspection.plan.is_some() {
                current_stage = "safe_commit";
                let commit_started = Instant::now();
                let commit = replace_target_from_candidate(
                    &candidate,
                    input,
                    expected_target_hash
                        .as_deref()
                        .context("target snapshot missing before commit")?,
                )?;
                emit_fx_header_event(
                    &mut events,
                    view,
                    run_started,
                    "safe_commit",
                    "completed",
                    "已验证候选并安全提交到原工作簿",
                    &progress,
                    Some(commit_started),
                    &["pytools-core", "file-commit", "excel-reader", "calamine"],
                    execution.get("engine").and_then(serde_json::Value::as_str),
                    None,
                );
                Some(serde_json::to_value(commit)?)
            } else {
                None
            };

            let output_copy = if create_copy {
                execution.get("output").cloned()
            } else {
                None
            };
            Ok(serde_json::json!({
                "input": input,
                "candidate": if inspection.plan.is_some() && !create_copy && in_place.is_none() { Some(&candidate) } else { None },
                "output_copy": output_copy,
                "plan_file": plan_file,
                "status": if commit.is_some() {
                    "committed"
                } else if create_copy && inspection.plan.is_some() {
                    "copy_created"
                } else if inspection.plan.is_none() {
                    "no_changes"
                } else {
                    "planned"
                },
                "fx_sheets": inspection.fx_sheets,
                "modified_sheets": inspection.modified_sheets,
                "execution": if inspection.plan.is_some() { execution } else { serde_json::Value::String("no_changes".to_owned()) },
                "target_commit": commit,
                "elapsed_ms": file_started.elapsed().as_millis(),
            }))
        })();

        let item_result = if create_copy || commit_target || in_place.is_some() {
            match (item_result, remove_regular_candidate(&candidate)) {
                (Ok(item), Ok(())) => Ok(item),
                (Err(error), Ok(())) => Err(error),
                (Ok(_), Err(cleanup_error)) => {
                    Err(cleanup_error).context("remove temporary FX header copy candidate")
                }
                (Err(error), Err(cleanup_error)) => Err(error.context(format!(
                    "temporary FX header candidate cleanup failed: {cleanup_error:#}"
                ))),
            }
        } else {
            item_result
        };

        match item_result {
            Ok(item) => {
                if item["plan_file"].is_string() {
                    planned_files += 1;
                }
                if item["status"] == "committed" {
                    committed_files += 1;
                }
                if item["status"] == "copy_created" {
                    created_copies += 1;
                }
                completed_files += 1;
                fx_sheets += item["fx_sheets"].as_u64().unwrap_or_default() as usize;
                modified_sheets += item["modified_sheets"].as_u64().unwrap_or_default() as usize;
                progress.processed_files = index + 1;
                progress.completed_files = completed_files;
                items.push(item);
                let item = items.last().expect("successful item was just appended");
                let (status, message) = match item["status"].as_str().unwrap_or_default() {
                    "committed" => ("completed", "当前工作簿的页眉修改已安全提交"),
                    "copy_created" => ("completed", "页眉修改副本已生成，源工作簿保持不变"),
                    "planned" => ("planned", "当前工作簿的页眉修改计划已保存，原文件未修改"),
                    _ => ("no_changes", "当前工作簿无需修改"),
                };
                emit_fx_header_event(
                    &mut events,
                    view,
                    run_started,
                    "file_summary",
                    status,
                    message,
                    &progress,
                    Some(file_started),
                    &["pytools-core"],
                    None,
                    None,
                );
            }
            Err(error) => {
                failed_files += 1;
                progress.processed_files = index + 1;
                progress.failed_files = failed_files;
                emit_fx_header_event(
                    &mut events,
                    view,
                    run_started,
                    current_stage,
                    "failed",
                    "当前工作簿处理失败；请查看开发者运行记录中的失败阶段",
                    &progress,
                    Some(file_started),
                    &["pytools-core"],
                    None,
                    Some("fx_header_stage_error"),
                );
                items.push(serde_json::json!({
                    "input": input,
                    "status": "failed",
                    "error_stage": current_stage,
                    "error": format!("{error:#}"),
                    "elapsed_ms": file_started.elapsed().as_millis(),
                }));
            }
        }
    }

    let total_files = options.positional.len();
    let status = if failed_files == 0 {
        "succeeded"
    } else if completed_files == 0 {
        "failed"
    } else {
        "partial"
    };
    progress.current_file = None;
    progress.processed_files = total_files;
    progress.completed_files = completed_files;
    progress.failed_files = failed_files;
    emit_fx_header_event(
        &mut events,
        view,
        run_started,
        "run_summary",
        status,
        format!("外汇页眉处理完成：成功 {completed_files}，失败 {failed_files}"),
        &progress,
        None,
        &["pytools-core"],
        None,
        None,
    );

    let mut result = serde_json::json!({
        "schema_version": 1,
        "product": "pytools",
        "feature_id": "PT-7-4",
        "function": "fx-header",
        "run_id": run_id,
        "started_at": started_at,
        "status": status,
        "mode": if in_place.is_some() { "in_place" } else if create_copy { "create_copy" } else if commit_target { "commit_target" } else { "plan_only" },
        "files": total_files,
        "completed_files": completed_files,
        "failed_files": failed_files,
        "planned_files": planned_files,
        "committed_files": committed_files,
        "created_copies": created_copies,
        "saved_files": committed_files + created_copies,
        "fx_sheets": fx_sheets,
        "modified_sheets": modified_sheets,
        "elapsed_ms": run_started.elapsed().as_millis(),
        "plan_dir": plan_dir,
        "items": items,
    });
    let unified_summary =
        attach_legacy_batch_diagnostics(&mut unified_tracker, &mut result, &log_dir, view);
    let terminal_output_paths = legacy_report_output_paths(&result);

    let mut history_warning = None::<String>;
    if let Some(path) = report_output.as_deref() {
        let safe_items = result["items"]
            .as_array()
            .into_iter()
            .flatten()
            .map(|item| {
                serde_json::json!({
                    "status": item["status"],
                    "error_stage": item.get("error_stage"),
                    "error_type": if item["status"] == "failed" { Some("fx_header_stage_error") } else { None },
                    "fx_sheets": item.get("fx_sheets"),
                    "modified_sheets": item.get("modified_sheets"),
                    "elapsed_ms": item.get("elapsed_ms"),
                })
            })
            .collect::<Vec<_>>();
        let history = serde_json::json!({
            "schema_version": 1,
            "product": "pytools",
            "feature_id": "PT-7-4",
            "function": "fx-header",
            "run_id": result["run_id"],
            "started_at": result["started_at"],
            "status": result["status"],
            "mode": result["mode"],
            "summary": {
                "total_files": total_files,
                "completed_files": completed_files,
                "failed_files": failed_files,
                "planned_files": planned_files,
                "committed_files": committed_files,
                "created_copies": created_copies,
                "fx_sheets": fx_sheets,
                "modified_sheets": modified_sheets,
                "elapsed_ms": run_started.elapsed().as_millis(),
            },
            "items": safe_items,
            "events": events,
        });
        match write_json_report_create_new(path, &history) {
            Ok(()) => result["history_report"] = serde_json::json!(path),
            Err(error) => {
                let warning = format!("could not save PT-7-4 run history: {error:#}");
                eprintln!("pytools: warning: {warning}");
                history_warning = Some(warning);
            }
        }
    }

    let output = match view {
        "developer" => {
            result["events"] = serde_json::to_value(&events)?;
            if let Some(warning) = history_warning {
                result["history_warning"] = serde_json::json!(warning);
            }
            result
        }
        "user" | "off" => {
            let mut user = fx_header_user_view(
                &result,
                run_started.elapsed().as_millis(),
                if view == "user" { Some(&events) } else { None },
            );
            user["run_id"] = result["run_id"].clone();
            user["started_at"] = result["started_at"].clone();
            if let Some(path) = result.get("history_report") {
                user["history_report"] = path.clone();
            }
            if let Some(warning) = history_warning {
                user["history_warning"] = serde_json::json!(warning);
            }
            user
        }
        _ => unreachable!("view was validated"),
    };
    if interactive && view == "user" {
        if let Some(summary) = unified_summary.as_ref() {
            emit_unified_terminal_summary_with_outputs(summary, &terminal_output_paths);
        } else {
            println!("{}", pytools_core::run_logging::terminal_line(if failed_files > 0 { pytools_core::run_logging::Level::Warn } else { pytools_core::run_logging::Level::Info }, &format!( "外汇页眉处理结束：输入 {total_files} 个，已完成 {completed_files} 个，失败 {failed_files} 个；耗时 {:.3} 秒", (run_started.elapsed().as_millis()) as f64 / 1000.0)));
        }
    } else if interactive && view == "off" {
        println!("{}", pytools_core::run_logging::terminal_line(if failed_files > 0 { pytools_core::run_logging::Level::Warn } else { pytools_core::run_logging::Level::Info }, &format!( "外汇页眉处理结束：输入 {total_files} 个，已完成 {completed_files} 个，失败 {failed_files} 个；耗时 {:.3} 秒", (run_started.elapsed().as_millis()) as f64 / 1000.0)));
    } else {
        println!("{}", serde_json::to_string_pretty(&output)?);
    }
    Ok(())
}

fn create_fx_header_copy(
    input: &Path,
    candidate: &Path,
    plan: &excel_editor::ExcelEditPlan,
    expected_source_hash: &str,
    publish_copy: bool,
) -> Result<serde_json::Value> {
    let updates = plan
        .operations
        .iter()
        .map(|operation| match operation {
            excel_editor::EditOperation::ReplaceHeaderText {
                sheet,
                search,
                replacement,
            } => Ok(ooxml_patch::HeaderTextUpdate {
                sheet_name: sheet.clone(),
                search: search.clone(),
                replacement: replacement.clone(),
            }),
            other => bail!("PT-7-4 copy route cannot execute {:?}", other.kind()),
        })
        .collect::<Result<Vec<_>>>()?;
    anyhow::ensure!(!updates.is_empty(), "FX header copy has no operations");
    verify_planning_snapshot(input, Some(expected_source_hash))?;
    let preflight = ooxml_patch::preflight_header_text_patch(input, &updates)
        .context("preflight FX header text patch")?;
    anyhow::ensure!(
        preflight.replacements > 0,
        "FX header patch found no matching text"
    );
    let written = ooxml_patch::replace_header_text(input, candidate, &updates)
        .context("write temporary FX header copy")?;
    anyhow::ensure!(
        written.replacements == preflight.replacements
            && written.modified_sheets == preflight.modified_sheets,
        "FX header copy output differs from the preflight result"
    );
    ooxml_patch::validate_header_text_patch(input, candidate, &updates)
        .context("validate FX header copy; unexpected workbook changes were detected")?;
    verify_planning_snapshot(input, Some(expected_source_hash))?;

    if !publish_copy {
        return Ok(serde_json::json!({
            "engine": "ooxml_patch",
            "execution": "validated_header_candidate",
            "modified_sheets": written.modified_sheets,
            "replacements": written.replacements,
        }));
    }
    let output = specialized_report_copy_path(input, "PyTools_专用报表处理结果")?;
    let publication = create_target_from_candidate(candidate, &output).with_context(|| {
        format!(
            "publish FX header copy to {}",
            display_path_for_user(&output)
        )
    })?;
    Ok(serde_json::json!({
        "engine": "ooxml_patch",
        "execution": "created_and_validated_workbook_copy",
        "output": output,
        "modified_sheets": written.modified_sheets,
        "replacements": written.replacements,
        "package_parts": written.package_parts,
        "publication": publication,
    }))
}

fn execute_fx_header_in_place(
    input: &Path,
    candidate: &Path,
    plan: &excel_editor::ExcelEditPlan,
    expected_source_hash: &str,
    options: &ParsedCli,
) -> Result<serde_json::Value> {
    match create_fx_header_copy(input, candidate, plan, expected_source_hash, false) {
        Ok(mut execution) => {
            let commit = replace_target_from_candidate(candidate, input, expected_source_hash)
                .context("safely replace source after validating OOXML header patch")?;
            execution["execution"] = serde_json::json!("validated_ooxml_patch_and_safe_replace");
            execution["commit"] = serde_json::to_value(commit)?;
            Ok(execution)
        }
        Err(ooxml_error) => {
            remove_regular_candidate(candidate)
                .context("remove rejected OOXML header candidate before Openpyxl fallback")?;
            verify_planning_snapshot(input, Some(expected_source_hash))
                .context("source changed after OOXML header route; refusing fallback")?;
            let mut fallback = execute_optional_plan(options, Some(plan))
                .with_context(|| format!("OOXML header patch failed ({ooxml_error:#}); protected Openpyxl fallback also failed"))?;
            fallback["fallback_reason"] = serde_json::json!(format!("{ooxml_error:#}"));
            fallback["route"] = serde_json::json!("openpyxl_protected_fallback");
            Ok(fallback)
        }
    }
}

fn emit_fx_header_event(
    events: &mut Vec<SummaryRunEvent>,
    view: &str,
    run_started: Instant,
    phase: &str,
    status: &str,
    user_message: impl Into<String>,
    progress: &SummaryProgress,
    stage_started: Option<Instant>,
    components: &[&str],
    actual_engine: Option<&str>,
    error_type: Option<&str>,
) {
    if view == "off" {
        return;
    }
    let event = SummaryRunEvent {
        sequence: events.len() + 1,
        timestamp: Local::now().to_rfc3339(),
        phase: phase.to_owned(),
        status: status.to_owned(),
        user_message: user_message.into(),
        progress: progress.clone(),
        file_name: None,
        stage_elapsed_ms: stage_started.map(|started| started.elapsed().as_millis()),
        total_elapsed_ms: run_started.elapsed().as_millis(),
        components: components.iter().map(|value| (*value).to_owned()).collect(),
        actual_engine: actual_engine.map(str::to_owned),
        error_type: error_type.map(str::to_owned),
    };
    if view == "developer" {
        if let Ok(encoded) = serde_json::to_string(&event) {
            eprintln!("{encoded}");
        }
    } else {
        let message = format!(
            "{}（进度 {}/{}, 已完成 {}，失败 {}，已耗时 {:.3} 秒）",
            event.user_message,
            event.progress.processed_files,
            event.progress.total_files,
            event.progress.completed_files,
            event.progress.failed_files,
            (event.total_elapsed_ms) as f64 / 1000.0
        );
        eprintln!(
            "{}",
            pytools_core::run_logging::terminal_line(
                if event.status == "failed" {
                    pytools_core::run_logging::Level::Error
                } else {
                    pytools_core::run_logging::Level::Info
                },
                &message
            )
        );
    }
    events.push(event);
}

fn fx_header_user_view(
    result: &serde_json::Value,
    elapsed_ms: u128,
    events: Option<&[SummaryRunEvent]>,
) -> serde_json::Value {
    let outputs = result["items"]
        .as_array()
        .into_iter()
        .flatten()
        .filter_map(|item| item["output_copy"].as_str())
        .collect::<Vec<_>>();
    let mut user = serde_json::json!({
        "schema_version": 1,
        "product": "pytools",
        "feature_id": "PT-7-4",
        "status": result["status"],
        "mode": result["mode"],
        "summary": {
            "total_files": result["files"],
            "completed_files": result["completed_files"],
            "failed_files": result["failed_files"],
            "planned_files": result["planned_files"],
            "committed_files": result["committed_files"],
            "created_copies": result["created_copies"],
            "fx_sheets": result["fx_sheets"],
            "modified_sheets": result["modified_sheets"],
            "elapsed_ms": elapsed_ms,
        },
        "outputs": outputs,
    });
    if let Some(events) = events {
        user["events"] = serde_json::Value::Array(
            events
                .iter()
                .map(|event| {
                    serde_json::json!({
                        "sequence": event.sequence,
                        "timestamp": event.timestamp,
                        "phase": fx_header_user_phase(&event.phase),
                        "status": event.status,
                        "message": event.user_message,
                        "progress": event.progress,
                        "elapsed_ms": event.stage_elapsed_ms,
                        "total_elapsed_ms": event.total_elapsed_ms,
                    })
                })
                .collect(),
        );
    }
    user
}

fn fx_header_user_phase(phase: &str) -> &str {
    match phase {
        "input_validation" | "profile_inspection" => "输入检查",
        "plan_build" => "页眉检查",
        "plan_persistence" => "计划保存",
        "engine_execution" => "页眉修改",
        "safe_commit" => "文件提交",
        "file_summary" => "文件完成",
        "run_summary" => "任务汇总",
        other => other,
    }
}

fn run_fx_header_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let input = required_path(&options, "input")?;
    let output = required_path(&options, "output")?;
    let profile = excel_editor::inspect_xlsx_profile(&input)?;
    let report = pytools_core::fx_header::build_fx_header_plan(
        &input,
        &output,
        profile,
        "pytools-fx-header",
    )?;
    write_optional_plan(&options, &report.plan)?;
    let execution = execute_optional_plan(&options, report.plan.as_ref())?;
    println!(
        "{}",
        serde_json::to_string_pretty(&serde_json::json!({
            "fx_sheets": report.fx_sheets,
            "modified_sheets": report.modified_sheets,
            "plan_written": report.plan.is_some(),
            "execution": execution,
        }))?
    );
    Ok(())
}

fn write_optional_plan(
    options: &ParsedCli,
    plan: &Option<excel_editor::ExcelEditPlan>,
) -> Result<()> {
    if let (Some(path), Some(plan)) = (options.values.get("plan-output"), plan) {
        if let Some(parent) = Path::new(path).parent() {
            if !parent.as_os_str().is_empty() {
                fs::create_dir_all(parent)?;
            }
        }
        fs::write(path, serde_json::to_vec_pretty(plan)?)?;
    }
    Ok(())
}

fn execute_optional_plan(
    options: &ParsedCli,
    plan: Option<&excel_editor::ExcelEditPlan>,
) -> Result<serde_json::Value> {
    let capabilities = options.values.get("capabilities");
    let platform = options.values.get("platform");
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let Some(plan) = plan else {
        return Ok(serde_json::json!("no_changes"));
    };
    // 既有工作簿就地修改统一使用 Openpyxl；旧的 umya 参数只作为兼容别名。
    if let Some(engine_name) = options.values.get("in-place") {
        let engine = match engine_name.as_str() {
            "openpyxl" | "umya" => pytools_core::inplace_edit::InPlaceEngine::Openpyxl,
            other => anyhow::bail!("--in-place 仅接受 openpyxl，收到 {other:?}"),
        };
        let outcome = pytools_core::inplace_edit::execute_in_place(plan, engine)?;
        return Ok(serde_json::json!({
            "mode": "in_place",
            "engine": engine.label(),
            "target": outcome.target,
            "execution": outcome.execution,
            "commit": outcome.commit,
        }));
    }
    match (capabilities, platform) {
        (Some(capabilities), Some(platform)) => Ok(serde_json::to_value(
            apply_plan_with_capability_file(plan, capabilities, platform)?,
        )?),
        _ => Ok(serde_json::json!("plan_only")),
    }
}

fn run_rename_sheets_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let input = required_path(&options, "input")?;
    let output = required_path(&options, "output")?;
    let report = build_sheet_rename_plan(&config, &input, &output)?;
    write_optional_plan(&options, &report.plan)?;
    let execution = execute_optional_plan(&options, report.plan.as_ref())?;
    let mut result = serde_json::to_value(report)?;
    result["execution"] = execution;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn run_rename_sheets(arguments: Vec<String>, interactive: bool) -> Result<()> {
    let (recursive, arguments) = extract_boolean_flag(arguments, "recursive")?;
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let log_dir = options
        .values
        .get("log-dir")
        .map(PathBuf::from)
        .unwrap_or_else(pytools_core::run_logging::default_log_dir);
    let log_view = options
        .values
        .get("log-view")
        .map(String::as_str)
        .unwrap_or("user");
    anyhow::ensure!(
        matches!(log_view, "user" | "developer" | "off"),
        "--log-view must be user, developer, or off"
    );
    let mut unified =
        UnifiedRunSession::new(&log_dir, "PT-4-5", "按配置修改 Sheet 名", log_view, None);
    unified.set_inputs(&options.positional, None);
    let result = run_batch_sheet_rename(&config, &options.positional, recursive);
    let report = match result {
        Ok(report) => report,
        Err(error) => {
            unified.mark_inputs_unavailable();
            unified.operation_failure(
                "sheet_rename",
                "SHEET_RENAME_FAILED",
                "工作表重命名失败",
                &error,
            );
            let _ = unified.finish(true, false, false, &[], None, None, None);
            return Err(error);
        }
    };
    unified.set_discovered(report.files_scanned);
    let mut output_paths = Vec::new();
    for (index, item) in report.files.iter().enumerate() {
        let status = match item.status.as_str() {
            "renamed" => {
                if let Some(path) = item.output.as_ref() {
                    output_paths.push(path.clone());
                }
                pytools_core::run_logging::FileStatus::Succeeded
            }
            "unchanged" | "skipped" => pytools_core::run_logging::FileStatus::Skipped,
            _ => pytools_core::run_logging::FileStatus::Failed,
        };
        let code = (status == pytools_core::run_logging::FileStatus::Failed)
            .then_some("SHEET_RENAME_FAILED");
        let detail = item
            .reason
            .clone()
            .unwrap_or_else(|| match item.engine.as_deref() {
                Some(engine) => format!("{}（引擎：{}）", item.status, engine),
                None => item.status.clone(),
            });
        unified.discovered_file_event(
            index + 1,
            &item.input,
            status,
            "sheet_rename",
            &detail,
            code,
            None,
        );
    }
    for index in 1..=options.positional.len() {
        unified.file_result(
            index,
            pytools_core::run_logging::FileStatus::Succeeded,
            true,
            "input_scan",
            "已完成所选文件范围的工作表重命名扫描",
            None,
            None,
        );
    }
    let output_stats =
        (!output_paths.is_empty()).then_some(pytools_core::run_logging::OutputStats {
            workbooks_saved: Some(output_paths.len()),
            sheets_saved: None,
            saved: Some(true),
        });
    let (summary, summary_path) = unified.finish(
        false,
        false,
        report.rename_ok == 0,
        &output_paths,
        output_stats,
        None,
        None,
    );
    if interactive && log_view == "user" {
        println!("{}", pytools_core::run_logging::terminal_line(if report.skip > 0 { pytools_core::run_logging::Level::Warn } else { pytools_core::run_logging::Level::Info }, &format!( "Sheet 改名结束：扫描 {} 个文件，处理 {} 个工作簿，成功改名 {} 个工作表，跳过 {} 项；耗时 {:.3} 秒；运行编号 {}{}", report.files_scanned, report.workbooks, report.rename_ok, report.skip, (report.elapsed_ms) as f64 / 1000.0, summary.run_id, summary_path.as_ref().map(|path| format!("；运行记录 {}", display_path_for_user(&path))).unwrap_or_default())));
    } else {
        let mut value = serde_json::to_value(&report)?;
        value["run_id"] = serde_json::json!(summary.run_id.clone());
        if let Some(path) = summary_path {
            value["run_summary_path"] = serde_json::json!(path);
        }
        println!("{}", serde_json::to_string_pretty(&value)?);
    }
    Ok(())
}

fn run_excel_open_repair_cli(arguments: Vec<String>, interactive: bool) -> Result<()> {
    let (recursive, arguments) = extract_boolean_flag(arguments, "recursive")?;
    let options = parse_cli_arguments(arguments)?;
    if options.positional.is_empty() {
        bail!("provide at least one file or directory path");
    }
    let log_dir = options
        .values
        .get("log-dir")
        .map(PathBuf::from)
        .unwrap_or_else(pytools_core::run_logging::default_log_dir);
    let log_view = options
        .values
        .get("log-view")
        .map(String::as_str)
        .unwrap_or("user");
    anyhow::ensure!(
        matches!(log_view, "user" | "developer" | "off"),
        "--log-view must be user, developer, or off"
    );
    let mut unified = UnifiedRunSession::new(
        &log_dir,
        "PT-4-11",
        "修复 Excel COM 无法打开的工作簿（菜单 4-7）",
        log_view,
        None,
    );
    unified.set_inputs(&options.positional, None);
    let result = pytools_core::excel_open_repair_batch::run_excel_open_repair_batch(
        &options.positional,
        recursive,
    );
    let report = match result {
        Ok(report) => report,
        Err(error) => {
            unified.mark_inputs_unavailable();
            unified.operation_failure(
                "excel_open_repair",
                "EXCEL_OPEN_REPAIR_FAILED",
                "Excel 打开修复失败",
                &error,
            );
            let _ = unified.finish(true, false, false, &[], None, None, None);
            return Err(error);
        }
    };
    unified.set_discovered(report.files_scanned);
    let mut output_paths = Vec::new();
    for (index, item) in report.files.iter().enumerate() {
        let status = match item.status.as_str() {
            "repaired" => {
                if let Some(path) = item.output.as_ref() {
                    output_paths.push(path.clone());
                }
                pytools_core::run_logging::FileStatus::Succeeded
            }
            "unchanged" | "skipped" => pytools_core::run_logging::FileStatus::Skipped,
            _ => pytools_core::run_logging::FileStatus::Failed,
        };
        let code = (status == pytools_core::run_logging::FileStatus::Failed)
            .then_some("EXCEL_OPEN_REPAIR_FAILED");
        let detail = item.reason.clone().unwrap_or_else(|| {
            format!(
                "已生成修复副本：引用补引号 {} 处、冻结窗格修正 {} 处",
                item.quoted, item.pane_fixed
            )
        });
        unified.discovered_file_event(
            index + 1,
            &item.input,
            status,
            "excel_open_repair",
            &detail,
            code,
            None,
        );
    }
    for index in 1..=options.positional.len() {
        unified.file_result(
            index,
            pytools_core::run_logging::FileStatus::Succeeded,
            true,
            "input_scan",
            "已完成所选范围的 Excel 打开修复扫描",
            None,
            None,
        );
    }
    let output_stats =
        (!output_paths.is_empty()).then_some(pytools_core::run_logging::OutputStats {
            workbooks_saved: Some(output_paths.len()),
            sheets_saved: None,
            saved: Some(true),
        });
    let has_failure = report.files.iter().any(|item| item.status == "failed");
    let (summary, summary_path) = unified.finish(
        false,
        false,
        has_failure
            || (report.repaired == 0
                && report.workbooks > 0
                && report.unchanged == report.workbooks
                && report.skipped == 0),
        &output_paths,
        output_stats,
        None,
        None,
    );
    if interactive && log_view == "user" {
        println!(
            "{}",
            pytools_core::run_logging::terminal_line(
                if report.skipped > 0 {
                    pytools_core::run_logging::Level::Warn
                } else {
                    pytools_core::run_logging::Level::Info
                },
                &format!(
                    "Excel 打开修复结束：扫描 {} 个文件，修复 {} 个（引用补引号 {} 处、冻结窗格修正 {} 处），无需修复 {} 个，跳过 {} 个；耗时 {:.3} 秒；运行记录 {}",
                    report.files_scanned,
                    report.repaired,
                    report.quoted_total,
                    report.pane_fixed_total,
                    report.unchanged,
                    report.skipped,
                    report.elapsed_ms as f64 / 1000.0,
                    summary_path
                        .as_ref()
                        .map(|path| display_path_for_user(path))
                        .unwrap_or_default()
                )
            )
        );
    } else {
        let mut value = serde_json::to_value(&report)?;
        value["run_id"] = serde_json::json!(summary.run_id.clone());
        if let Some(path) = summary_path {
            value["run_summary_path"] = serde_json::json!(path);
        }
        println!("{}", serde_json::to_string_pretty(&value)?);
    }
    Ok(())
}

fn run_submission_check_cli(arguments: Vec<String>, interactive: bool) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let output_dir = options
        .values
        .get("output-dir")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("output"));
    if options.positional.is_empty() {
        bail!("provide at least one file or directory path");
    }
    let log_view = options
        .values
        .get("log-view")
        .map(String::as_str)
        .unwrap_or("user");
    anyhow::ensure!(
        matches!(log_view, "user" | "developer" | "off"),
        "--log-view must be user, developer, or off"
    );
    let log_dir = options
        .values
        .get("log-dir")
        .map(PathBuf::from)
        .or_else(|| load_global(&config).ok().map(|global| global.log_dir))
        .unwrap_or_else(pytools_core::run_logging::default_log_dir);
    let mut unified =
        UnifiedRunSession::new(&log_dir, "PT-4-8", "机构提交完整性检查", log_view, None);
    unified.set_inputs(&options.positional, None);
    let mut events = Vec::<serde_json::Value>::new();
    let report_result = if log_view == "off" {
        run_submission_check(&config, &options.positional, &output_dir)
    } else {
        run_submission_check_with_progress(&config, &options.positional, &output_dir, |event| {
            let value = serde_json::to_value(&event).unwrap_or(serde_json::Value::Null);
            let level = if event.status == "failed" {
                pytools_core::run_logging::Level::Error
            } else if event.status == "skipped" {
                pytools_core::run_logging::Level::Warn
            } else {
                pytools_core::run_logging::Level::Info
            };
            unified.progress(
                level,
                &event.phase,
                &event.phase,
                &event.status,
                &event.user_message,
                event.stage_elapsed_ms,
                event.actual_engine.as_deref(),
                event
                    .error_type
                    .as_ref()
                    .map(|error| serde_json::json!({"error_type": error})),
            );
            emit_submission_check_event_view(log_view, &event, &value);
            if log_view == "developer" {
                events.push(value);
            }
        })
    };
    let history_result = if log_view == "developer" {
        Some(write_submission_check_run_history(
            &log_dir,
            report_result.as_ref().ok(),
            &events,
        ))
    } else {
        None
    };
    if let Some(Err(error)) = &history_result {
        eprintln!("pytools: warning: could not write structured run history: {error:#}");
    }
    let report = match report_result {
        Ok(report) => report,
        Err(error) => {
            unified.mark_inputs_unavailable();
            let failed_event = events
                .iter()
                .rev()
                .find(|event| event["status"] == "failed");
            let stage = failed_event
                .and_then(|event| event["phase"].as_str())
                .unwrap_or("submission_check");
            let error_code = match stage {
                "configuration_load" => "CONFIG_READ_FAILED",
                "source_discovery" => "FILE_DISCOVERY_FAILED",
                "report_write" | "report_publish" => "OUTPUT_WRITE_FAILED",
                _ => "SUBMISSION_CHECK_FAILED",
            };
            unified.operation_failure(stage, error_code, "提交完整性检查失败", &error);
            let _ = unified.finish(true, false, false, &[], None, None, None);
            if let Some(Ok(path)) = &history_result {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        pytools_core::run_logging::Level::Info,
                        &format!(
                            "pytools: developer run history: {}",
                            display_path_for_user(&path)
                        )
                    )
                );
            }
            return Err(error);
        }
    };
    unified.set_discovered(report.files);
    for index in 1..=options.positional.len() {
        unified.file_result(
            index,
            pytools_core::run_logging::FileStatus::Succeeded,
            true,
            "submission_check",
            "所选输入位置已纳入提交完整性报告",
            None,
            None,
        );
    }
    unified.set_business_metric("已提交配置项", serde_json::json!(report.submitted));
    unified.set_business_metric("缺失配置项", serde_json::json!(report.missing));
    unified.set_business_metric("未匹配文件", serde_json::json!(report.unmatched));
    unified.set_business_metric("多重匹配文件", serde_json::json!(report.multi));
    for (category, count) in [
        ("缺失提交", report.missing),
        ("未匹配文件", report.unmatched),
        ("多重匹配文件", report.multi),
    ] {
        if count > 0 {
            unified.record_business_issue(
                pytools_core::run_logging::BusinessIssue::new(category, "完整性检查发现待处理项")
                    .count(count),
            );
        }
    }
    let (unified_summary, unified_path) = unified.finish(
        false,
        false,
        false,
        std::slice::from_ref(&report.output),
        Some(pytools_core::run_logging::OutputStats {
            workbooks_saved: Some(1),
            sheets_saved: None,
            saved: Some(true),
        }),
        None,
        None,
    );
    let mut output = serde_json::to_value(&report)?;
    output["run_id"] = serde_json::json!(unified_summary.run_id.clone());
    if let Some(path) = unified_path.as_ref() {
        output["run_summary_path"] = serde_json::json!(path);
    }
    if let Some(Ok(path)) = &history_result {
        output["run_history_path"] = serde_json::json!(path);
    }
    if let Some(Err(error)) = &history_result {
        output["run_history_warning"] =
            serde_json::json!(format!("could not write structured run history: {error:#}"));
    }
    if interactive && log_view == "user" {
        emit_unified_terminal_summary_with_outputs(
            &unified_summary,
            std::slice::from_ref(&report.output),
        );
    } else {
        println!("{}", serde_json::to_string_pretty(&output)?);
    }
    Ok(())
}

fn emit_submission_check_event_view(
    log_view: &str,
    event: &SubmissionCheckProgressEvent,
    event_value: &serde_json::Value,
) {
    let current = (event.progress.processed_files > 0).then_some(
        event
            .progress
            .processed_files
            .min(event.progress.total_files),
    );
    emit_batch_event_view(
        log_view,
        event.progress.total_files,
        current,
        &event.status,
        current.is_none() || event.phase == "complete",
        &event.user_message,
        Some(event_value),
        Some(event.total_elapsed_ms),
    );
}

fn write_submission_check_run_history(
    log_dir: &Path,
    report: Option<&SubmissionCheckReport>,
    events: &[serde_json::Value],
) -> Result<PathBuf> {
    fs::create_dir_all(log_dir).with_context(|| {
        format!(
            "create submission-check log directory {}",
            display_path_for_user(&log_dir)
        )
    })?;
    let timestamp = Local::now().format("%Y%m%d_%H%M%S");
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .context("system clock before Unix epoch")?
        .as_nanos();
    let path = log_dir.join(format!("PT-4-8_submission-check_{timestamp}_{nonce}.json"));
    let mut actual_engines = Vec::<String>::new();
    let mut components = Vec::<String>::new();
    for event in events {
        if let Some(engine) = event["actual_engine"].as_str() {
            if !actual_engines.iter().any(|existing| existing == engine) {
                actual_engines.push(engine.to_owned());
            }
        }
        if let Some(event_components) = event["components"].as_array() {
            for component in event_components
                .iter()
                .filter_map(serde_json::Value::as_str)
            {
                if !components.iter().any(|existing| existing == component) {
                    components.push(component.to_owned());
                }
            }
        }
    }
    let summary = report.map(|report| {
        serde_json::json!({
            "rules": report.rules,
            "files": report.files,
            "submitted": report.submitted,
            "missing": report.missing,
            "unmatched": report.unmatched,
            "multi": report.multi,
        })
    });
    let history = serde_json::json!({
        "schema_version": 1,
        "product": "pytools",
        "feature_id": "PT-4-8",
        "function": "submission-check",
        "status": if report.is_some() { "succeeded" } else { "failed" },
        "actual_engines": actual_engines,
        "components": components,
        "summary": summary,
        "total_elapsed_ms": events.last().and_then(|event| event["total_elapsed_ms"].as_u64()),
        "events": events,
    });
    let mut file = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&path)
        .with_context(|| {
            format!(
                "create submission-check run history {}",
                display_path_for_user(&path)
            )
        })?;
    serde_json::to_writer_pretty(&mut file, &history)?;
    file.write_all(b"\n")?;
    Ok(path)
}

fn run_convert_batch_cli(arguments: Vec<String>, excel: bool, interactive: bool) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    if options.positional.is_empty() {
        bail!("provide at least one source file path");
    }
    let engine = options
        .values
        .get("engine")
        .map(String::as_str)
        .unwrap_or("auto");
    let target = options.values.get("to").cloned().unwrap_or_else(|| {
        pytools_core::config::load_global_value(
            &config,
            if excel {
                "excel目的格式"
            } else {
                "word目的格式"
            },
            if excel { "xlsx" } else { "docx" },
        )
        .unwrap_or_else(|_| if excel { "xlsx".into() } else { "docx".into() })
    });
    let log_view = options
        .values
        .get("log-view")
        .map(String::as_str)
        .unwrap_or("user");
    anyhow::ensure!(
        matches!(log_view, "user" | "developer" | "off"),
        "--log-view must be user, developer, or off"
    );
    let log_dir = options
        .values
        .get("log-dir")
        .map(PathBuf::from)
        .or_else(|| load_global(&config).ok().map(|global| global.log_dir))
        .unwrap_or_else(pytools_core::run_logging::default_log_dir);
    let output_dir = options.values.get("output-dir").map(Path::new);
    let feature_id = if excel { "PT-4-2" } else { "PT-4-3" };
    let function = if excel {
        "convert-excel"
    } else {
        "convert-word"
    };
    use pytools_core::run_logging::{EventWriter, FileStatus, Level, RunTracker};
    let mut tracker = RunTracker::new(
        feature_id,
        if excel {
            "Excel 格式转换"
        } else {
            "Word 格式转换"
        },
        &[],
    );
    let mut diagnostic = if log_view == "off" {
        None
    } else {
        match EventWriter::create(&log_dir.join("run-events"), feature_id, &tracker.run_id) {
            Ok(writer) => Some(writer),
            Err(error) => {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("无法创建诊断日志：{error}；业务处理继续")
                    )
                );
                None
            }
        }
    };
    let mut diagnostic_failed = false;
    let mut events = Vec::<serde_json::Value>::new();
    if let Some(writer) = diagnostic.as_mut() {
        if let Err(error) = writer.write(&tracker.event(
            Level::Info,
            "run_started",
            "input_expansion",
            "started",
            "开始转换文件",
        )) {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Warn,
                    &format!("诊断日志写入失败：{error}")
                )
            );
            diagnostic_failed = true;
        }
    }
    let mut on_event = |event: &ConversionProgressEvent| {
        let value = serde_json::to_value(event).unwrap_or(serde_json::Value::Null);
        emit_conversion_event_view(log_view, event, &value);
        if log_view == "developer" {
            events.push(value);
        }
        if let Some(writer) = diagnostic.as_mut() {
            let level = if event.status == "failed" {
                Level::Error
            } else if event.status == "skipped" {
                Level::Warn
            } else {
                Level::Info
            };
            let mut record = tracker.event(
                level,
                &event.phase,
                &event.phase,
                &event.status,
                &event.user_message,
            );
            record.timestamp = event.timestamp.clone();
            record.file_index = event.input_index;
            record.engine = event.actual_engine.clone();
            record.elapsed_ms = event.stage_elapsed_ms;
            if event.status == "failed" {
                record.error_code = Some("FILE_CONVERT_FAILED".to_owned());
            }
            if let Err(error) = writer.write(&record) {
                if !diagnostic_failed {
                    eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            Level::Warn,
                            &format!("诊断日志写入失败：{error}")
                        )
                    );
                    diagnostic_failed = true;
                }
            }
        }
    };
    let result = if excel {
        convert_excel_files_with_progress(
            &options.positional,
            output_dir,
            &target,
            engine,
            &mut on_event,
        )
    } else {
        convert_word_files_with_progress(
            &options.positional,
            output_dir,
            &target,
            engine,
            &mut on_event,
        )
    };
    drop(on_event);
    if let Err(error) = result {
        let mut summary = tracker.finish(
            true,
            false,
            false,
            vec![],
            diagnostic_path_for_log_view(diagnostic.as_ref(), log_view),
        );
        summary.input_stats_available = false;
        if let Some(writer) = diagnostic.as_mut() {
            let _ = writer.write(&tracker.event(
                Level::Error,
                "run_finished",
                "conversion",
                "failed",
                "转换任务中断",
            ));
        }
        if log_view != "off" {
            let path = log_dir.join(format!("{feature_id}_{}.run.json", tracker.run_id));
            if let Err(log_error) = write_json_report_create_new(&path, &summary) {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("运行记录写入失败：{log_error}")
                    )
                );
            }
        }
        return Err(error);
    }
    let result = result.expect("error was returned above");
    let discovered_files = result
        .items
        .iter()
        .map(|item| item.input.clone())
        .collect::<Vec<_>>();
    let _ = tracker.set_inputs(&discovered_files);
    tracker.set_discovered(discovered_files.len());
    for (index, item) in result.items.iter().enumerate() {
        let position = index + 1;
        let _ = tracker.transition(position, FileStatus::InProgress);
        let state = match item.status.as_str() {
            "converted" | "copied" => FileStatus::Succeeded,
            "failed" => FileStatus::Failed,
            _ => FileStatus::Skipped,
        };
        let _ = tracker.transition(position, state);
        if let Some(writer) = diagnostic.as_mut() {
            let level = if state == FileStatus::Failed {
                Level::Error
            } else if state == FileStatus::Skipped {
                Level::Warn
            } else {
                Level::Info
            };
            let conversion_message = match item.status.as_str() {
                "converted" => format!("{}：文件格式转换完成", tracker.feature_name),
                "copied" => format!("{}：格式无需转换，文件已复制", tracker.feature_name),
                "failed" => format!("{}：文件转换失败", tracker.feature_name),
                _ => format!("{}：文件转换已跳过", tracker.feature_name),
            };
            let mut event = tracker.event(
                level,
                "file_result",
                "conversion",
                &item.status,
                conversion_message,
            );
            event.file_index = Some(position);
            event.file_name = Some(pytools_core::run_logging::portable_file_name(&item.input));
            if state == FileStatus::Failed {
                event.error_code = Some("FILE_CONVERT_FAILED".to_owned());
                event.message = format!("转换失败：{}", item.detail);
            }
            if let Err(error) = writer.write(&event) {
                if !diagnostic_failed {
                    eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            Level::Warn,
                            &format!("诊断日志写入失败：{error}")
                        )
                    );
                    diagnostic_failed = true;
                }
            }
        }
    }
    tracker.set_business_metric("成功转换数", serde_json::json!(result.converted));
    tracker.set_business_metric("失败转换数", serde_json::json!(result.failed));
    tracker.set_business_metric("跳过文件数", serde_json::json!(result.skipped));
    let output_paths = result
        .items
        .iter()
        .filter_map(|item| item.output.as_ref())
        .map(|path| path_for_log_view(path, log_view))
        .collect();
    let mut unified = tracker.finish(
        false,
        false,
        result.converted == 0 && result.failed == 0,
        output_paths,
        diagnostic_path_for_log_view(diagnostic.as_ref(), log_view),
    );
    if excel {
        unified.output_stats = Some(pytools_core::run_logging::OutputStats {
            workbooks_saved: Some(result.converted),
            sheets_saved: None,
            saved: Some(result.converted > 0),
        });
    }
    if let Some(writer) = diagnostic.as_mut() {
        if let Err(error) = writer.write(&tracker.event(
            if result.failed > 0 {
                Level::Error
            } else {
                Level::Info
            },
            "run_finished",
            "conversion",
            &format!("{:?}", unified.overall_status),
            format!("{}结束", tracker.feature_name),
        )) {
            if !diagnostic_failed {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("诊断日志写入失败：{error}")
                    )
                );
            }
        }
    }
    let mut unified_report_path = None;
    if log_view != "off" {
        let mut report = serde_json::to_value(&unified)?;
        report["selected_path_count"] = serde_json::json!(options.positional.len());
        report["conversion_errors"] = serde_json::json!(result
            .items
            .iter()
            .filter(|item| item.status == "failed")
            .map(|item| serde_json::json!({
                "file_name": pytools_core::run_logging::portable_file_name(&item.input),
                "reason": item.detail,
            }))
            .collect::<Vec<_>>());
        let path = log_dir.join(format!("{feature_id}_{}.run.json", tracker.run_id));
        match write_json_report_create_new(&path, &report) {
            Ok(()) => {
                if interactive {
                    eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            Level::Info,
                            &format!(
                                "运行编号 {}；运行记录：{}",
                                tracker.run_id,
                                pytools_core::run_logging::path_for_log_view(
                                    &path,
                                    log_view == "developer"
                                )
                            )
                        )
                    );
                }
                unified_report_path = Some(path);
            }
            Err(error) => eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Warn,
                    &format!("运行记录写入失败：{error}")
                )
            ),
        }
    }
    let run_history_path = if log_view == "developer" {
        match write_conversion_run_history(&log_dir, feature_id, function, engine, &result, &events)
        {
            Ok(path) => Some(path),
            Err(error) => {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("历史运行记录写入失败：{error}")
                    )
                );
                None
            }
        }
    } else {
        None
    };
    let mut output = serde_json::to_value(&result)?;
    output["run_id"] = serde_json::json!(tracker.run_id);
    if let Some(path) = unified_report_path {
        output["run_summary_path"] = serde_json::json!(path.to_string_lossy());
    }
    if let Some(path) = run_history_path {
        output["run_history_path"] = serde_json::Value::String(path.to_string_lossy().into_owned());
    }
    for item in result.items.iter().filter(|item| item.status == "failed") {
        let name = item
            .input
            .file_name()
            .unwrap_or(item.input.as_os_str())
            .to_string_lossy();
        let reason = if item.detail.contains("external fidelity engine")
            || item.detail.contains("该功能需要外部引擎")
        {
            "需要可用的转换引擎"
        } else if item.detail.contains("unsupported") {
            "文件格式不受支持"
        } else if item.detail.contains("does not exist") {
            "源文件不存在"
        } else {
            "转换或保存失败，请检查文件及访问权限"
        };
        eprintln!(
            "{}",
            pytools_core::run_logging::terminal_line(
                Level::Error,
                &format!(
                    "转换失败：{name}；{reason}；具体原因：{}；运行编号 {}",
                    item.detail, tracker.run_id
                )
            )
        );
    }
    if !interactive || log_view == "developer" {
        println!("{}", serde_json::to_string_pretty(&output)?);
    } else if log_view == "user" {
        let outputs = result
            .items
            .iter()
            .filter_map(|item| item.output.clone())
            .collect::<Vec<_>>();
        emit_unified_terminal_summary_with_outputs(&unified, &outputs);
    }
    // off 视图：静默。
    if result.failed > 0 {
        bail!("{} conversion(s) failed", result.failed);
    }
    Ok(())
}

fn run_convert_batch_menu<W: Write>(
    arguments: Vec<String>,
    excel: bool,
    output: &mut W,
) -> Result<()> {
    if let Err(error) = run_convert_batch_cli(arguments, excel, true) {
        writeln!(
            output,
            "本次转换未完成：{error:#}。文件处理菜单仍可继续操作。"
        )?;
    }
    Ok(())
}

fn emit_conversion_event_view(
    log_view: &str,
    event: &ConversionProgressEvent,
    event_value: &serde_json::Value,
) {
    let current = event
        .input_index
        .or_else(|| (event.completed_files > 0).then_some(event.completed_files));
    let phase_change = current.is_none()
        || (event.status == "completed" && event.completed_files >= event.total_files);
    emit_batch_event_view(
        log_view,
        event.total_files,
        current,
        &event.status,
        phase_change,
        &event.user_message,
        Some(event_value),
        Some(event.total_elapsed_ms),
    );
}

fn write_conversion_run_history(
    log_dir: &Path,
    feature_id: &str,
    function: &str,
    requested_engine: &str,
    result: &pytools_core::conversion::ConversionRun,
    events: &[serde_json::Value],
) -> Result<PathBuf> {
    fs::create_dir_all(log_dir).with_context(|| {
        format!(
            "create conversion log directory {}",
            display_path_for_user(&log_dir)
        )
    })?;
    let timestamp = Local::now().format("%Y%m%d_%H%M%S");
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .context("system clock before Unix epoch")?
        .as_nanos();
    let path = log_dir.join(format!("{feature_id}_conversion_{timestamp}_{nonce}.json"));
    let mut actual_engines = Vec::<String>::new();
    for event in events {
        if let Some(engine) = event["actual_engine"].as_str() {
            if !actual_engines.iter().any(|existing| existing == engine) {
                actual_engines.push(engine.to_owned());
            }
        }
    }
    let status = if result.failed == 0 {
        "succeeded"
    } else if result.converted > 0 {
        "partial_failure"
    } else {
        "failed"
    };
    let history = serde_json::json!({
        "schema_version": 1,
        "product": "pytools",
        "feature_id": feature_id,
        "function": function,
        "status": status,
        "requested_engine": requested_engine,
        "converted": result.converted,
        "skipped": result.skipped,
        "failed": result.failed,
        "actual_engines": actual_engines,
        "total_elapsed_ms": events.last().map(|event| event["total_elapsed_ms"].clone()),
        "events": events,
    });
    fs::write(&path, serde_json::to_vec_pretty(&history)?).with_context(|| {
        format!(
            "write conversion run history {}",
            display_path_for_user(&path)
        )
    })?;
    Ok(path)
}

fn run_convert_file_cli(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let input = required_path(&options, "input")?;
    let target = options.values.get("to").context("--to is required")?;
    let output = if let Some(value) = options.values.get("output") {
        PathBuf::from(value)
    } else {
        input.with_extension(target)
    };
    let engine = options
        .values
        .get("engine")
        .map(String::as_str)
        .unwrap_or("auto");
    let validation = convert_one(&input, &output, target, engine)?;
    println!("{validation}");
    Ok(())
}

fn run_timeline_extract_cli(arguments: Vec<String>) -> Result<()> {
    let report = run_timeline_extract_cli_report(arguments, None)?;
    println!("{}", serde_json::to_string_pretty(&report)?);
    Ok(())
}

fn run_timeline_extract_cli_report(
    arguments: Vec<String>,
    mut menu_io: Option<(&mut dyn BufRead, &mut dyn Write)>,
) -> Result<serde_json::Value> {
    let (commit_target, arguments) = extract_boolean_flag(arguments, "commit-target")?;
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let log_view = options
        .values
        .get("log-view")
        .map(String::as_str)
        .unwrap_or("user");
    anyhow::ensure!(
        matches!(log_view, "user" | "developer" | "off"),
        "--log-view must be user, developer, or off"
    );
    let capabilities = options.values.get("capabilities").map(PathBuf::from);
    let platform = options.values.get("platform").cloned();
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );

    let run_started = Instant::now();
    let started_at = Local::now().to_rfc3339();
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_nanos())
        .unwrap_or_default();
    let run_id = format!("pytools-pt-1-2-{}-{nonce}", std::process::id());
    let mut events = Vec::<TimelineExtractRunEvent>::new();
    let mut progress = TimelineExtractProgress {
        total_inputs: options.positional.len(),
        ..TimelineExtractProgress::default()
    };
    let configuration_started = Instant::now();
    if log_view != "off" {
        emit_timeline_extract_cli_event(
            log_view,
            &mut events,
            TimelineExtractRunEvent {
                sequence: 1,
                timestamp: Local::now().to_rfc3339(),
                feature_id: "PT-1-2",
                phase: "configuration_read".to_owned(),
                status: "started".to_owned(),
                user_message: "正在读取时序提取配置".to_owned(),
                progress: progress.clone(),
                input_index: None,
                stage_elapsed_ms: None,
                total_elapsed_ms: run_started.elapsed().as_millis(),
                components: vec![
                    "Rust".to_owned(),
                    "pytools-core".to_owned(),
                    "excel-reader".to_owned(),
                    "calamine".to_owned(),
                ],
                actual_engine: None,
                error_type: None,
            },
        );
    }
    let configuration = (|| -> Result<_> {
        let global = load_global(&config)?;
        let rules = pytools_core::config::load_timeline_rules(&config)?;
        let path_maps = pytools_core::config::load_path_maps(&config)?;
        Ok((global, rules, path_maps))
    })();
    let (global, rules, path_maps) = match configuration {
        Ok(configuration) => configuration,
        Err(error) => {
            if log_view != "off" {
                progress.failed_inputs = options.positional.len();
                let sequence = events.len() + 1;
                emit_timeline_extract_cli_event(
                    log_view,
                    &mut events,
                    TimelineExtractRunEvent {
                        sequence,
                        timestamp: Local::now().to_rfc3339(),
                        feature_id: "PT-1-2",
                        phase: "configuration_read".to_owned(),
                        status: "failed".to_owned(),
                        user_message: "时序提取配置读取失败".to_owned(),
                        progress: progress.clone(),
                        input_index: None,
                        stage_elapsed_ms: Some(configuration_started.elapsed().as_millis()),
                        total_elapsed_ms: run_started.elapsed().as_millis(),
                        components: vec![
                            "Rust".to_owned(),
                            "pytools-core".to_owned(),
                            "excel-reader".to_owned(),
                            "calamine".to_owned(),
                        ],
                        actual_engine: Some("calamine".to_owned()),
                        error_type: Some("configuration_read_error".to_owned()),
                    },
                );
                let log_dir = options
                    .values
                    .get("log-dir")
                    .map(PathBuf::from)
                    .unwrap_or_else(pytools_core::run_logging::default_log_dir);
                if let Err(history_error) = write_timeline_extract_run_history(
                    &log_dir,
                    &run_id,
                    &started_at,
                    run_started.elapsed().as_millis(),
                    None,
                    "failed",
                    "configuration_read_error",
                    &events,
                ) {
                    eprintln!(
                        "pytools: warning: could not save PT-1-2 run history: {history_error:#}"
                    );
                }
            }
            return Err(error).context("时序提取配置读取失败；请检查配置文件");
        }
    };
    let log_dir = options
        .values
        .get("log-dir")
        .map(PathBuf::from)
        .unwrap_or_else(|| global.log_dir.clone());
    if log_view != "off" {
        let sequence = events.len() + 1;
        emit_timeline_extract_cli_event(
            log_view,
            &mut events,
            TimelineExtractRunEvent {
                sequence,
                timestamp: Local::now().to_rfc3339(),
                feature_id: "PT-1-2",
                phase: "configuration_read".to_owned(),
                status: "completed".to_owned(),
                user_message: format!(
                    "时序提取配置已读取：{} 条规则、{} 条路径映射",
                    rules.len(),
                    path_maps.len()
                ),
                progress: progress.clone(),
                input_index: None,
                stage_elapsed_ms: Some(configuration_started.elapsed().as_millis()),
                total_elapsed_ms: run_started.elapsed().as_millis(),
                components: vec![
                    "Rust".to_owned(),
                    "pytools-core".to_owned(),
                    "excel-reader".to_owned(),
                    "calamine".to_owned(),
                ],
                actual_engine: Some("calamine".to_owned()),
                error_type: None,
            },
        );
    }

    let source_paths = if options.positional.is_empty() {
        let discovery_started = Instant::now();
        if log_view != "off" {
            let sequence = events.len() + 1;
            emit_timeline_extract_cli_event(
                log_view,
                &mut events,
                TimelineExtractRunEvent {
                    sequence,
                    timestamp: Local::now().to_rfc3339(),
                    feature_id: "PT-1-2",
                    phase: "source_discovery".to_owned(),
                    status: "started".to_owned(),
                    user_message: "正在扫描配置输入目录中的时序源文件".to_owned(),
                    progress: progress.clone(),
                    input_index: None,
                    stage_elapsed_ms: None,
                    total_elapsed_ms: run_started.elapsed().as_millis(),
                    components: vec!["Rust".to_owned(), "std::fs".to_owned()],
                    actual_engine: None,
                    error_type: None,
                },
            );
        }
        match discover_configured_source_files(&global) {
            Ok(sources) => {
                progress.total_inputs = sources.len();
                if log_view != "off" {
                    let sequence = events.len() + 1;
                    emit_timeline_extract_cli_event(
                        log_view,
                        &mut events,
                        TimelineExtractRunEvent {
                            sequence,
                            timestamp: Local::now().to_rfc3339(),
                            feature_id: "PT-1-2",
                            phase: "source_discovery".to_owned(),
                            status: "completed".to_owned(),
                            user_message: format!("已发现 {} 个时序源文件", sources.len()),
                            progress: progress.clone(),
                            input_index: None,
                            stage_elapsed_ms: Some(discovery_started.elapsed().as_millis()),
                            total_elapsed_ms: run_started.elapsed().as_millis(),
                            components: vec!["Rust".to_owned(), "std::fs".to_owned()],
                            actual_engine: None,
                            error_type: None,
                        },
                    );
                }
                sources
            }
            Err(error) => {
                if log_view != "off" {
                    let sequence = events.len() + 1;
                    emit_timeline_extract_cli_event(
                        log_view,
                        &mut events,
                        TimelineExtractRunEvent {
                            sequence,
                            timestamp: Local::now().to_rfc3339(),
                            feature_id: "PT-1-2",
                            phase: "source_discovery".to_owned(),
                            status: "failed".to_owned(),
                            user_message: "扫描时序输入目录失败".to_owned(),
                            progress: progress.clone(),
                            input_index: None,
                            stage_elapsed_ms: Some(discovery_started.elapsed().as_millis()),
                            total_elapsed_ms: run_started.elapsed().as_millis(),
                            components: vec!["Rust".to_owned(), "std::fs".to_owned()],
                            actual_engine: None,
                            error_type: Some("source_discovery_error".to_owned()),
                        },
                    );
                    if let Err(history_error) = write_timeline_extract_run_history(
                        &log_dir,
                        &run_id,
                        &started_at,
                        run_started.elapsed().as_millis(),
                        None,
                        "failed",
                        "source_discovery_error",
                        &events,
                    ) {
                        eprintln!(
                            "pytools: warning: could not save PT-1-2 run history: {history_error:#}"
                        );
                    }
                }
                return Err(error).context("扫描时序提取输入目录失败");
            }
        }
    } else {
        options.positional.clone()
    };

    let output_dir = options
        .values
        .get("output-dir")
        .map(PathBuf::from)
        .unwrap_or_else(|| global.output_dir.clone());
    let mut execution_global = global.clone();
    execution_global.output_dir = output_dir.clone();
    let mut result = if log_view == "off" {
        run_timeline_extract(&source_paths, &rules, &path_maps, &execution_global)
    } else {
        run_timeline_extract_with_progress(
            &source_paths,
            &rules,
            &path_maps,
            &execution_global,
            |mut event| {
                event.sequence = events.len() + 1;
                event.total_elapsed_ms = run_started.elapsed().as_millis();
                progress = event.progress.clone();
                emit_timeline_extract_cli_event(log_view, &mut events, event);
            },
        )
    };
    if result.is_err()
        && log_view != "off"
        && !events.last().is_some_and(|event| event.status == "failed")
    {
        let sequence = events.len() + 1;
        let failure_progress = events
            .last()
            .map(|event| event.progress.clone())
            .unwrap_or_else(|| progress.clone());
        emit_timeline_extract_cli_event(
            log_view,
            &mut events,
            TimelineExtractRunEvent {
                sequence,
                timestamp: Local::now().to_rfc3339(),
                feature_id: "PT-1-2",
                phase: "run".to_owned(),
                status: "failed".to_owned(),
                user_message: "时序提取失败".to_owned(),
                progress: failure_progress,
                input_index: None,
                stage_elapsed_ms: None,
                total_elapsed_ms: run_started.elapsed().as_millis(),
                components: vec!["Rust".to_owned(), "pytools-core".to_owned()],
                actual_engine: None,
                error_type: Some("timeline_extract_error".to_owned()),
            },
        );
    }

    let mut target_execution_error = None;
    let mut target_append_results = Vec::<serde_json::Value>::new();
    if let Ok(extracted) = result.as_mut() {
        let target_progress = events
            .last()
            .map(|event| event.progress.clone())
            .unwrap_or_else(|| progress.clone());
        let existing_ready = extracted
            .target_append_plans
            .iter()
            .filter(|item| item.plan.target_existed && item.plan.appended_rows > 0)
            .collect::<Vec<_>>();
        let mut commit_existing = commit_target;

        if commit_target && !existing_ready.is_empty() && capabilities.is_none() {
            target_execution_error = Some(anyhow::anyhow!(
                "--commit-target requires --capabilities and --platform for existing PT-1-2 target workbooks"
            ));
        }

        if target_execution_error.is_none()
            && !commit_existing
            && !existing_ready.is_empty()
            && capabilities.is_some()
            && menu_io.is_some()
        {
            let (input, output) = menu_io.as_mut().expect("menu IO was checked");
            writeln!(
                output,
                "PT-1-2 将更新以下已存在的目标簿（新增 {} 行）：",
                existing_ready
                    .iter()
                    .map(|item| item.plan.appended_rows)
                    .sum::<usize>()
            )?;
            for item in &existing_ready {
                writeln!(
                    output,
                    "  {}（{} 行）",
                    display_path_for_user(&item.plan.target),
                    item.plan.appended_rows
                )?;
            }
            writeln!(output, "输入 y/yes/是 确认执行，其他输入只保留计划。")?;
            output.flush()?;
            let mut choice = String::new();
            let confirmed = input.read_line(&mut choice)? > 0
                && matches!(
                    choice.trim().to_ascii_lowercase().as_str(),
                    "y" | "yes" | "是"
                );
            commit_existing = confirmed;
        }

        let mut any_deferred = false;
        for (index, item) in extracted.target_append_plans.iter().enumerate() {
            let plan = &item.plan;
            let has_payload = if plan.target_existed {
                plan.edit_plan.is_some() && plan.replacement_workbook.is_some()
            } else {
                plan.new_workbook.is_some()
            };
            if plan.appended_rows == 0 || !has_payload {
                let status = if plan.task_skipped > 0 {
                    "skipped"
                } else {
                    "no_new_rows"
                };
                target_append_results.push(serde_json::json!({
                    "target": plan.target,
                    "target_existed": plan.target_existed,
                    "status": status,
                    "requested_rows": item.requested_rows,
                    "appended_rows": plan.appended_rows,
                    "duplicate_rows_filtered": plan.duplicate_rows_filtered,
                    "issues": plan.issues,
                }));
                continue;
            }

            let plan_path =
                output_dir.join(format!("PT-1-2_{run_id}_target_{:03}.plan.json", index + 1));
            if let Err(error) = write_json_report_create_new(&plan_path, plan) {
                target_execution_error = Some(error.context("save PT-1-2 target append plan"));
                any_deferred = true;
                break;
            }

            let should_commit = !plan.target_existed || commit_existing;
            if !should_commit {
                any_deferred = true;
                target_append_results.push(serde_json::json!({
                    "target": plan.target,
                    "target_existed": true,
                    "status": "planned_only",
                    "reason": "explicit_confirmation_or_capability_required",
                    "plan_file": plan_path,
                    "requested_rows": item.requested_rows,
                    "appended_rows": plan.appended_rows,
                    "duplicate_rows_filtered": plan.duplicate_rows_filtered,
                    "issues": plan.issues,
                }));
                if log_view == "user" {
                    eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            pytools_core::run_logging::Level::Info,
                            &format!(
                                "目标簿计划已保存，未修改既有文件：{}",
                                display_path_for_user(&plan_path)
                            )
                        )
                    );
                }
                continue;
            }

            let mut execute_arguments = vec![
                "--plan".to_owned(),
                plan_path.to_string_lossy().into_owned(),
                "--commit-target".to_owned(),
            ];
            if plan.target_existed {
                let Some(capabilities) = capabilities.as_ref() else {
                    target_execution_error = Some(anyhow::anyhow!(
                        "existing PT-1-2 target requires --capabilities and --platform"
                    ));
                    any_deferred = true;
                    break;
                };
                execute_arguments.push("--capabilities".to_owned());
                execute_arguments.push(capabilities.to_string_lossy().into_owned());
                execute_arguments.push("--platform".to_owned());
                execute_arguments.push(
                    platform
                        .as_deref()
                        .expect("capability/platform pair validated")
                        .to_owned(),
                );
            }
            let execution_started = Instant::now();
            match run_dedup_append_execute_report(execute_arguments) {
                Ok(execution) => {
                    let elapsed_ms = execution_started.elapsed().as_millis();
                    let engine = execution["engine"]
                        .as_str()
                        .or_else(|| execution["candidate_engine"].as_str())
                        .map(str::to_owned);
                    target_append_results.push(serde_json::json!({
                        "target": plan.target,
                        "target_existed": plan.target_existed,
                        "status": "succeeded",
                        "requested_rows": item.requested_rows,
                        "appended_rows": plan.appended_rows,
                        "duplicate_rows_filtered": plan.duplicate_rows_filtered,
                        "elapsed_ms": elapsed_ms,
                        "engine": engine,
                        "target_commit": execution["target_commit"],
                        "plan_file": plan_path,
                        "issues": plan.issues,
                    }));
                    if log_view == "user" {
                        eprintln!(
                            "目标簿写入完成：{}，新增 {} 行，耗时 {:.3} 秒",
                            display_path_for_user(&plan.target),
                            plan.appended_rows,
                            (elapsed_ms) as f64 / 1000.0
                        );
                    }
                    if log_view != "off" {
                        let event = TimelineExtractRunEvent {
                            sequence: events.len() + 1,
                            timestamp: Local::now().to_rfc3339(),
                            feature_id: "PT-1-2",
                            phase: "target_append_execute".to_owned(),
                            status: "completed".to_owned(),
                            user_message: format!(
                                "目标簿写入完成：新增 {} 行，实际引擎 {}",
                                plan.appended_rows,
                                engine.as_deref().unwrap_or("unknown")
                            ),
                            progress: target_progress.clone(),
                            input_index: None,
                            stage_elapsed_ms: Some(elapsed_ms),
                            total_elapsed_ms: run_started.elapsed().as_millis(),
                            components: std::iter::once("Rust".to_owned())
                                .chain(std::iter::once("pytools-core".to_owned()))
                                .chain(engine.iter().cloned())
                                .collect(),
                            actual_engine: engine,
                            error_type: None,
                        };
                        emit_timeline_extract_cli_event(log_view, &mut events, event);
                    }
                }
                Err(error) => {
                    let elapsed_ms = execution_started.elapsed().as_millis();
                    any_deferred = true;
                    target_append_results.push(serde_json::json!({
                        "target": plan.target,
                        "target_existed": plan.target_existed,
                        "status": "failed",
                        "requested_rows": item.requested_rows,
                        "appended_rows": plan.appended_rows,
                        "elapsed_ms": elapsed_ms,
                        "error_type": "target_append_execute_error",
                    }));
                    if log_view != "off" {
                        let event = TimelineExtractRunEvent {
                            sequence: events.len() + 1,
                            timestamp: Local::now().to_rfc3339(),
                            feature_id: "PT-1-2",
                            phase: "target_append_execute".to_owned(),
                            status: "failed".to_owned(),
                            user_message: "目标簿写入失败；已有目标保持原件，候选已验证后再提交"
                                .to_owned(),
                            progress: target_progress.clone(),
                            input_index: None,
                            stage_elapsed_ms: Some(elapsed_ms),
                            total_elapsed_ms: run_started.elapsed().as_millis(),
                            components: vec![
                                "Rust".to_owned(),
                                "pytools-core".to_owned(),
                                "excel-editor".to_owned(),
                            ],
                            actual_engine: None,
                            error_type: Some("target_append_execute_error".to_owned()),
                        };
                        emit_timeline_extract_cli_event(log_view, &mut events, event);
                    }
                    target_execution_error = Some(error.context("PT-1-2 target append failed"));
                    break;
                }
            }
        }
        extracted.target_writes_deferred = any_deferred;
    }

    let history_path = if log_view == "off" {
        None
    } else {
        let status = if target_execution_error.is_some() {
            "failed"
        } else {
            match result.as_ref() {
                Err(_) => "failed",
                Ok(result) if result.failed_files > 0 || result.target_writes_deferred => "partial",
                Ok(result) if result.path.is_none() => "skipped",
                Ok(_) => "succeeded",
            }
        };
        let error_type = if target_execution_error.is_some() {
            "target_append_execute_error"
        } else {
            events
                .iter()
                .rev()
                .find_map(|event| event.error_type.as_deref())
                .unwrap_or("")
        };
        match write_timeline_extract_run_history(
            &log_dir,
            &run_id,
            &started_at,
            run_started.elapsed().as_millis(),
            result.as_ref().ok(),
            status,
            error_type,
            &events,
        ) {
            Ok(path) => Some(path),
            Err(error) => {
                eprintln!("pytools: warning: could not save PT-1-2 run history: {error:#}");
                None
            }
        }
    };

    if let Some(error) = target_execution_error {
        return Err(error);
    }

    match result {
        Ok(result) => {
            let mut output = serde_json::to_value(result)?;
            output["target_append_results"] = serde_json::Value::Array(target_append_results);
            if log_view != "off" {
                output["run_id"] = serde_json::Value::String(run_id);
                if let Some(path) = history_path {
                    output["run_log_path"] =
                        serde_json::Value::String(display_path_for_user(&path).to_string());
                    if log_view == "user" {
                        eprintln!(
                            "{}",
                            pytools_core::run_logging::terminal_line(
                                pytools_core::run_logging::Level::Info,
                                &format!(
                                    "运行记录已保存：{}",
                                    pytools_core::run_logging::path_for_log_view(&path, false)
                                )
                            )
                        );
                    }
                }
            }
            Ok(output)
        }
        Err(error) => Err(error),
    }
}

fn emit_timeline_extract_cli_event(
    log_view: &str,
    events: &mut Vec<TimelineExtractRunEvent>,
    event: TimelineExtractRunEvent,
) {
    match log_view {
        "developer" => match serde_json::to_string(&event) {
            Ok(encoded) => eprintln!("{encoded}"),
            Err(_) => {
                eprintln!("{{\"phase\":\"progress_event\",\"status\":\"serialization_failed\"}}")
            }
        },
        "user" => {
            let level = match event.status.as_str() {
                "failed" => pytools_core::run_logging::Level::Error,
                "skipped" | "partial_failure" => pytools_core::run_logging::Level::Warn,
                _ => pytools_core::run_logging::Level::Info,
            };
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(level, &event.user_message)
            );
        }
        _ => return,
    }
    events.push(event);
}

fn write_timeline_extract_run_history(
    log_dir: &Path,
    run_id: &str,
    started_at: &str,
    total_elapsed_ms: u128,
    result: Option<&TimelineExtractResult>,
    status: &str,
    error_type: &str,
    events: &[TimelineExtractRunEvent],
) -> Result<PathBuf> {
    let mut actual_engines = Vec::<String>::new();
    for event in events {
        if let Some(engine) = event.actual_engine.as_ref() {
            if !actual_engines.contains(engine) {
                actual_engines.push(engine.clone());
            }
        }
    }
    let failure_phase = events
        .iter()
        .rev()
        .find(|event| event.status == "failed")
        .map(|event| event.phase.clone());
    let report = serde_json::json!({
        "schema_version": 1,
        "product": "pytools",
        "feature_id": "PT-1-2",
        "function": "timeline-extract",
        "run_id": run_id,
        "started_at": started_at,
        "status": status,
        "summary": result,
        "progress": events.last().map(|event| &event.progress),
        "total_elapsed_ms": total_elapsed_ms,
        "actual_engines": actual_engines,
        "failure_phase": failure_phase,
        "error_type": if error_type.is_empty() { serde_json::Value::Null } else { serde_json::json!(error_type) },
        "events": events,
    });
    let path = log_dir.join(format!("PT-1-2_{run_id}.run.json"));
    write_json_report_create_new(&path, &report)?;
    Ok(path)
}

fn run_timeline_precheck_cli(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let log_dir = options
        .values
        .get("log-dir")
        .map(PathBuf::from)
        .unwrap_or_else(pytools_core::run_logging::default_log_dir);
    let log_view = options
        .values
        .get("log-view")
        .map(String::as_str)
        .unwrap_or("user");
    anyhow::ensure!(
        matches!(log_view, "user" | "developer" | "off"),
        "--log-view must be user, developer, or off"
    );
    let mut unified = UnifiedRunSession::new(&log_dir, "PT-1-5", "时序配置预校验", log_view, None);
    unified.set_inputs(std::slice::from_ref(&config), None);
    let report = match run_timeline_precheck(&config) {
        Ok(report) => report,
        Err(error) => {
            unified.mark_inputs_unavailable();
            unified.operation_failure(
                "configuration_read",
                "CONFIG_READ_FAILED",
                "时序配置预校验读取失败",
                &error,
            );
            let _ = unified.finish(true, false, false, &[], None, None, None);
            return Err(error);
        }
    };
    unified.set_discovered(1);
    if report.errors > 0 {
        unified.file_result(
            1,
            pytools_core::run_logging::FileStatus::Failed,
            true,
            "configuration_validation",
            "配置预校验发现错误",
            Some("CONFIG_VALIDATION_FAILED"),
            None,
        );
        for issue in report.issues.iter().filter(|issue| issue.level == "error") {
            unified.progress(
                pytools_core::run_logging::Level::Error,
                "configuration_issue",
                "configuration_validation",
                "failed",
                &format!("{}：{}", issue.location, issue.message),
                None,
                None,
                Some(serde_json::json!({"location": issue.location})),
            );
        }
    } else {
        unified.file_result(
            1,
            pytools_core::run_logging::FileStatus::Succeeded,
            true,
            "configuration_validation",
            "配置预校验通过",
            None,
            None,
        );
    }
    let (summary, summary_path) =
        unified.finish(report.errors > 0, false, false, &[], None, None, None);
    let mut value = serde_json::to_value(&report)?;
    value["run_id"] = serde_json::json!(summary.run_id.clone());
    if let Some(path) = summary_path {
        value["run_summary_path"] = serde_json::json!(path);
    }
    println!("{}", serde_json::to_string_pretty(&value)?);
    if report.errors > 0 {
        bail!("时序配置预校验发现 {} 项错误", report.errors);
    }
    Ok(())
}

fn run_interactive_timeline_precheck<W>(
    output: &mut W,
    config_path: &Path,
    global: &GlobalConfig,
    log_view: &str,
) -> Result<()>
where
    W: Write,
{
    let mut diagnostic =
        UnifiedRunSession::new(&global.log_dir, "PT-1-5", "时序配置预校验", log_view, None);
    diagnostic.set_inputs(std::slice::from_ref(&config_path.to_path_buf()), Some(1));
    writeln!(output, "\n=== 预校验报告 ===")?;
    writeln!(
        output,
        "{}",
        pytools_core::run_logging::terminal_line(
            pytools_core::run_logging::Level::Info,
            &format!(
                "开始预校验配置文件 {}",
                pytools_core::run_logging::portable_file_name(config_path)
            )
        )
    )?;
    let report = match run_timeline_precheck(config_path) {
        Ok(report) => report,
        Err(error) => {
            diagnostic.mark_inputs_unavailable();
            diagnostic.operation_failure(
                "configuration_read",
                "CONFIG_READ_FAILED",
                "时序配置预校验读取失败",
                &error,
            );
            let _ = diagnostic.finish(true, false, false, &[], None, None, None);
            return Err(error);
        }
    };
    if report.errors > 0 {
        diagnostic.file_result(
            1,
            pytools_core::run_logging::FileStatus::Failed,
            true,
            "configuration_validation",
            "配置预校验发现错误",
            Some("CONFIG_VALIDATION_FAILED"),
            None,
        );
    } else {
        diagnostic.file_result(
            1,
            pytools_core::run_logging::FileStatus::Succeeded,
            true,
            "configuration_validation",
            "配置预校验通过",
            None,
            None,
        );
    }
    for issue in &report.issues {
        let level = if issue.level == "error" {
            "错误"
        } else {
            "警告"
        };
        let log_level = if issue.level == "error" {
            pytools_core::run_logging::Level::Error
        } else {
            pytools_core::run_logging::Level::Warn
        };
        writeln!(
            output,
            "{}",
            pytools_core::run_logging::terminal_line(
                log_level,
                &format!("[{level}] {}: {}", issue.location, issue.message)
            )
        )?;
        diagnostic.progress(
            if issue.level == "error" {
                pytools_core::run_logging::Level::Error
            } else {
                pytools_core::run_logging::Level::Warn
            },
            "configuration_issue",
            "configuration_validation",
            &issue.level,
            &format!("{}：{}", issue.location, issue.message),
            None,
            None,
            Some(serde_json::json!({"location": issue.location})),
        );
    }
    let (summary, summary_path) =
        diagnostic.finish(report.errors > 0, false, false, &[], None, None, None);
    writeln!(
        output,
        "{}",
        pytools_core::run_logging::terminal_line(
            if report.errors > 0 {
                pytools_core::run_logging::Level::Error
            } else {
                pytools_core::run_logging::Level::Info
            },
            &format!(
                "预校验完成：错误 {} 项，警告 {} 项",
                report.errors, report.warnings
            )
        )
    )?;
    if report.errors == 0 {
        writeln!(
            output,
            "{}",
            pytools_core::run_logging::terminal_line(
                pytools_core::run_logging::Level::Info,
                "时序配置预校验通过。"
            )
        )?;
    } else {
        writeln!(
            output,
            "{}",
            pytools_core::run_logging::terminal_line(
                pytools_core::run_logging::Level::Error,
                &format!("发现 {} 项错误，请按上方位置修正后重试。", report.errors)
            )
        )?;
    }
    writeln!(
        output,
        "{}",
        pytools_core::run_logging::terminal_line(
            pytools_core::run_logging::Level::Info,
            &format!(
                "运行编号 {}{}",
                summary.run_id,
                summary_path
                    .map(|path| format!("；运行记录 {}", display_path_for_user(&path)))
                    .unwrap_or_default()
            )
        )
    )?;
    Ok(())
}

fn run_timeline_compare_cli(arguments: Vec<String>, interactive: bool) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let template = required_path(&options, "template")?;
    let run_started = Instant::now();
    let started_at = Local::now().to_rfc3339();
    let run_id = header_path_compare_run_id();
    let log_view = options
        .values
        .get("log-view")
        .map(String::as_str)
        .unwrap_or("user");
    anyhow::ensure!(
        matches!(log_view, "user" | "developer" | "off"),
        "--log-view must be user, developer, or off"
    );
    let explicit_log_dir = options.values.get("log-dir").map(PathBuf::from);
    let mut log_dir = explicit_log_dir
        .clone()
        .unwrap_or_else(pytools_core::run_logging::default_log_dir);
    let mut progress = HeaderPathCompareProgress {
        total_inputs: options.positional.len(),
        ..HeaderPathCompareProgress::default()
    };
    let mut events = Vec::<HeaderPathCompareRunEvent>::new();
    let configuration_started = Instant::now();
    if log_view != "off" {
        let event = new_header_path_compare_cli_event(
            &events,
            &run_started,
            &progress,
            "configuration_read",
            "started",
            "正在读取表头比对配置".to_owned(),
            &configuration_started,
            &["excel-reader", "calamine"],
            None,
            None,
        );
        emit_header_path_compare_event_view(log_view, &mut events, event);
    }
    let configuration = (|| -> Result<_> {
        let global = load_global(&config)?;
        let rules = pytools_core::config::load_timeline_rules(&config)?;
        Ok((global, rules))
    })();
    let (global, rules) = match configuration {
        Ok(configuration) => configuration,
        Err(error) => {
            let error_type = "configuration_read_error";
            if log_view != "off" {
                let event = new_header_path_compare_cli_event(
                    &events,
                    &run_started,
                    &progress,
                    "configuration_read",
                    "failed",
                    "表头比对配置读取失败".to_owned(),
                    &configuration_started,
                    &["excel-reader", "calamine"],
                    Some("calamine"),
                    Some(error_type),
                );
                emit_header_path_compare_event_view(log_view, &mut events, event);
                let task_event = new_header_path_compare_cli_event(
                    &events,
                    &run_started,
                    &progress,
                    "task",
                    "failed",
                    "表头路径比对任务失败".to_owned(),
                    &configuration_started,
                    &["pytools-core"],
                    None,
                    Some(error_type),
                );
                emit_header_path_compare_event_view(log_view, &mut events, task_event);
            }
            if log_view == "developer" {
                match write_header_path_compare_run_history(
                    &log_dir,
                    &run_id,
                    "failed",
                    &started_at,
                    run_started.elapsed().as_millis(),
                    &progress,
                    error_type,
                    &events,
                ) {
                    Ok(path) => eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            pytools_core::run_logging::Level::Info,
                            &format!("PT-1-1 run history: {}", display_path_for_user(&path))
                        )
                    ),
                    Err(history_error) => {
                        eprintln!("{}", pytools_core::run_logging::terminal_line(pytools_core::run_logging::Level::Warn, &format!("pytools: warning: could not write PT-1-1 run history: {history_error:#}")))
                    }
                }
            }
            let mut failed_run = UnifiedRunSession::new(
                &log_dir,
                "PT-1-1",
                "表头路径比对",
                log_view,
                Some(run_id.clone()),
            );
            failed_run.mark_inputs_unavailable();
            failed_run.operation_failure(
                "configuration_read",
                "CONFIG_READ_FAILED",
                "表头比对配置读取失败",
                &error,
            );
            let _ = failed_run.finish(true, false, false, &[], None, None, None);
            bail!("表头比对配置读取失败；请检查配置文件");
        }
    };
    if explicit_log_dir.is_none() {
        log_dir = global.log_dir.clone();
    }
    use pytools_core::run_logging::{EventWriter, FileStatus, Level, RunTracker, TaskStatus};
    let mut tracker = RunTracker::with_id(
        "PT-1-1",
        "表头路径比对",
        &options.positional,
        run_id.clone(),
    );
    let mut diagnostic = if log_view == "off" {
        None
    } else {
        match EventWriter::create(&log_dir.join("run-events"), "PT-1-1", &run_id) {
            Ok(writer) => Some(writer),
            Err(error) => {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("无法创建诊断日志：{error}")
                    )
                );
                None
            }
        }
    };
    if let Some(writer) = diagnostic.as_mut() {
        if let Err(error) = writer.write(&tracker.event(
            Level::Info,
            "run_started",
            "configuration",
            "started",
            "开始表头路径比对",
        )) {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Warn,
                    &format!("诊断日志写入失败：{error}")
                )
            );
        }
    }
    if log_view != "off" {
        let event = new_header_path_compare_cli_event(
            &events,
            &run_started,
            &progress,
            "configuration_read",
            "completed",
            "表头比对配置读取完成".to_owned(),
            &configuration_started,
            &["excel-reader", "calamine"],
            Some("calamine"),
            None,
        );
        emit_header_path_compare_event_view(log_view, &mut events, event);
    }

    let mut compared_before = 0usize;
    let compare_result = run_header_path_compare_with_progress(
        &template,
        &options.positional,
        &rules,
        &global,
        |core_event| {
            let mut event = core_event.clone();
            if let Some(index) = event.progress.current_input {
                match (
                    event.phase.as_str(),
                    event.status.as_str(),
                    tracker.state(index),
                ) {
                    ("source_processing", "skipped", Some(FileStatus::Pending)) => {
                        let _ = tracker.transition(index, FileStatus::Skipped);
                    }
                    ("source_read", "started", Some(FileStatus::Pending)) => {
                        compared_before = event.progress.compared_sheets;
                        let _ = tracker.transition(index, FileStatus::InProgress);
                    }
                    ("source_read", "failed", Some(FileStatus::InProgress)) => {
                        let _ = tracker.transition(index, FileStatus::Failed);
                    }
                    ("source_compare", "completed" | "partial", Some(FileStatus::InProgress)) => {
                        let state = if event.progress.compared_sheets > compared_before {
                            FileStatus::Succeeded
                        } else {
                            FileStatus::Skipped
                        };
                        let _ = tracker.transition(index, state);
                    }
                    _ => {}
                }
            }
            event.sequence = events.len() + 1;
            event.total_elapsed_ms = run_started.elapsed().as_millis();
            progress = event.progress.clone();
            if let Some(writer) = diagnostic.as_mut() {
                let level = if event.status == "failed" {
                    Level::Error
                } else if matches!(event.status.as_str(), "skipped" | "partial") {
                    Level::Warn
                } else {
                    Level::Info
                };
                let mut record = tracker.event(
                    level,
                    &event.phase,
                    &event.phase,
                    &event.status,
                    "表头路径比对阶段",
                );
                record.timestamp = event.timestamp.clone();
                record.file_index = event.progress.current_input;
                record.file_name = event
                    .progress
                    .current_input
                    .and_then(|index| options.positional.get(index.saturating_sub(1)))
                    .map(|path| pytools_core::run_logging::portable_file_name(path));
                record.engine = event.actual_engine.clone();
                record.elapsed_ms = event.stage_elapsed_ms;
                if event.status == "failed" {
                    record.error_code = Some(
                        if event.phase == "report_write" {
                            "OUTPUT_WRITE_FAILED"
                        } else {
                            "EXCEL_READ_FAILED"
                        }
                        .to_owned(),
                    );
                }
                if let Err(error) = writer.write(&record) {
                    eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            Level::Warn,
                            &format!("诊断日志写入失败：{error}")
                        )
                    );
                }
            }
            emit_header_path_compare_event_view(log_view, &mut events, event);
        },
    );
    let status = match &compare_result {
        Err(_) => "failed",
        Ok(_)
            if progress.failed_inputs > 0
                || progress.unmatched_inputs > 0
                || progress.partially_unmatched_inputs > 0 =>
        {
            "partial"
        }
        Ok(_) => "succeeded",
    };
    let error_type = events
        .last()
        .and_then(|event| event.error_type.as_deref())
        .map(str::to_owned)
        .unwrap_or_else(|| {
            if compare_result.is_err() {
                "comparison_error".to_owned()
            } else {
                String::new()
            }
        });
    if compare_result.is_err() && log_view != "off" {
        let event = new_header_path_compare_cli_event(
            &events,
            &run_started,
            &progress,
            "task",
            "failed",
            "表头路径比对任务失败".to_owned(),
            &run_started,
            &["pytools-core"],
            None,
            (!error_type.is_empty()).then_some(error_type.as_str()),
        );
        emit_header_path_compare_event_view(log_view, &mut events, event);
    }
    let mut compare_error_summary = None;
    if let Err(error) = compare_result.as_ref() {
        let stage = events
            .iter()
            .rev()
            .find(|event| event.status == "failed")
            .map(|event| event.phase.as_str())
            .unwrap_or("comparison");
        let code = if stage == "report_write" {
            "OUTPUT_WRITE_FAILED"
        } else {
            "EXCEL_COMPARE_FAILED"
        };
        compare_error_summary = Some(pytools_core::run_logging::ErrorSummary {
            stage: stage.to_owned(),
            code: code.to_owned(),
            message: "表头路径比对失败".to_owned(),
            file_name: progress
                .current_input
                .and_then(|index| options.positional.get(index.saturating_sub(1)))
                .map(|path| pytools_core::run_logging::portable_file_name(path)),
            sheet_name: None,
        });
        if let Some(writer) = diagnostic.as_mut() {
            let mut detail = tracker.event(
                Level::Error,
                "operation_failed",
                stage,
                "failed",
                "表头路径比对失败",
            );
            detail.error_code = Some(code.to_owned());
            if log_view == "developer" {
                detail.details = Some(serde_json::json!({"error_chain": format!("{error:#}")}));
            }
            if let Err(write_error) = writer.write(&detail) {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("诊断日志写入失败：{write_error}")
                    )
                );
            }
        }
    }
    tracker.set_business_metric("比对工作表数", serde_json::json!(progress.compared_sheets));
    tracker.set_business_metric("未匹配输入数", serde_json::json!(progress.unmatched_inputs));
    tracker.set_business_metric(
        "部分未匹配输入数",
        serde_json::json!(progress.partially_unmatched_inputs),
    );
    if progress.unmatched_inputs > 0 {
        tracker.record_business_issue(
            pytools_core::run_logging::BusinessIssue::new(
                "未匹配工作表",
                "部分输入工作表未匹配模板路径",
            )
            .count(progress.unmatched_inputs),
        );
    }
    if progress.partially_unmatched_inputs > 0 {
        tracker.record_business_issue(
            pytools_core::run_logging::BusinessIssue::new(
                "部分未匹配",
                "工作簿中存在未匹配的工作表",
            )
            .count(progress.partially_unmatched_inputs),
        );
    }
    let paths = compare_result
        .as_ref()
        .ok()
        .map(|path| vec![path_for_log_view(path, log_view)])
        .unwrap_or_default();
    let mut unified = tracker.finish(
        compare_result.is_err(),
        false,
        compare_result.is_ok() && tracker.input_stats().succeeded == 0,
        paths,
        diagnostic_path_for_log_view(diagnostic.as_ref(), log_view),
    );
    unified.error_summary = compare_error_summary;
    if compare_result.is_ok()
        && unified.input_stats.succeeded > 0
        && (progress.unmatched_inputs > 0
            || progress.partially_unmatched_inputs > 0
            || progress.failed_inputs > 0)
    {
        unified.overall_status = TaskStatus::PartialSuccess;
    }
    unified.output_stats = Some(pytools_core::run_logging::OutputStats {
        workbooks_saved: Some(unified.output_paths.len()),
        sheets_saved: None,
        saved: Some(compare_result.is_ok()),
    });
    if let Some(writer) = diagnostic.as_mut() {
        if let Err(error) = writer.write(&tracker.event(
            if compare_result.is_err() {
                Level::Error
            } else {
                Level::Info
            },
            "run_finished",
            "run",
            &format!("{:?}", unified.overall_status),
            "表头路径比对结束",
        )) {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Warn,
                    &format!("诊断日志写入失败：{error}")
                )
            );
        }
    }
    let unified_path = if log_view == "off" {
        None
    } else {
        let path = log_dir
            .join("run-unified")
            .join(format!("PT-1-1_{run_id}.json"));
        match write_json_report_create_new(&path, &unified) {
            Ok(()) => Some(path),
            Err(error) => {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("统一运行记录写入失败：{error}")
                    )
                );
                None
            }
        }
    };
    let history_result = if log_view == "developer" {
        Some(write_header_path_compare_run_history(
            &log_dir,
            &run_id,
            status,
            &started_at,
            run_started.elapsed().as_millis(),
            &progress,
            &error_type,
            &events,
        ))
    } else {
        None
    };
    if interactive && log_view == "user" {
        emit_unified_terminal_summary_with_outputs(&unified, &[]);
    }
    let output = match compare_result {
        Ok(path) => path,
        Err(_error) => {
            if let Some(Ok(history_path)) = history_result.as_ref() {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        pytools_core::run_logging::Level::Info,
                        &format!(
                            "PT-1-1 run history: {}",
                            display_path_for_user(&history_path)
                        )
                    )
                );
            } else if let Some(Err(history_error)) = history_result.as_ref() {
                eprintln!(
                    "pytools: warning: could not write PT-1-1 run history: {history_error:#}"
                );
            }
            bail!("表头路径比对未完成；请检查配置、模板、源工作簿和输出目录");
        }
    };
    let mut report =
        serde_json::json!({"output": output, "run_id": run_id, "run_summary_path": unified_path});
    if let Some(history_result) = history_result {
        match history_result {
            Ok(path) => report["run_history_path"] = serde_json::json!(path),
            Err(error) => report["history_warning"] = serde_json::json!(error.to_string()),
        }
    }
    if interactive && log_view == "user" {
        // The shared runtime logger already printed the final summary.
    } else if interactive && log_view == "off" {
        let stats = &unified.input_stats;
        println!("{}", pytools_core::run_logging::terminal_line(if matches!(unified.overall_status, TaskStatus::Failed | TaskStatus::PartialSuccess) { Level::Warn } else { Level::Info }, &format!("表头比对结束：选择 {} 个，成功 {} 个，失败 {} 个，跳过 {} 个；报告：{}；运行编号 {}", stats.selected, stats.succeeded, stats.failed, stats.skipped, display_path_for_user(&output), tracker.run_id)));
    } else {
        println!("{}", serde_json::to_string(&report)?);
    }
    if log_view == "user" {
        eprintln!(
            "[人工确认] 请核对比对报告中的“原始路径/标准路径”候选；确认后再录入配置工作簿的“路径标准化映射”表，供后续时序汇总使用。本次运行只生成报告，配置和源工作簿保持不变。"
        );
    }
    Ok(())
}

fn new_header_path_compare_cli_event(
    events: &[HeaderPathCompareRunEvent],
    run_started: &Instant,
    progress: &HeaderPathCompareProgress,
    phase: &str,
    status: &str,
    user_message: String,
    stage_started: &Instant,
    components: &[&str],
    actual_engine: Option<&str>,
    error_type: Option<&str>,
) -> HeaderPathCompareRunEvent {
    let mut recorded_components = vec!["Rust".to_owned()];
    recorded_components.extend(components.iter().map(|component| (*component).to_owned()));
    HeaderPathCompareRunEvent {
        sequence: events.len() + 1,
        timestamp: Local::now().to_rfc3339(),
        feature_id: "PT-1-1".to_owned(),
        phase: phase.to_owned(),
        status: status.to_owned(),
        user_message,
        progress: progress.clone(),
        stage_elapsed_ms: Some(stage_started.elapsed().as_millis()),
        total_elapsed_ms: run_started.elapsed().as_millis(),
        components: recorded_components,
        actual_engine: actual_engine.map(str::to_owned),
        error_type: error_type.map(str::to_owned),
    }
}

fn emit_header_path_compare_event_view(
    log_view: &str,
    events: &mut Vec<HeaderPathCompareRunEvent>,
    event: HeaderPathCompareRunEvent,
) {
    let current = event.progress.current_input;
    let value = serde_json::to_value(&event).ok();
    emit_batch_event_view(
        log_view,
        event.progress.total_inputs,
        current,
        &event.status,
        current.is_none(),
        &event.user_message,
        value.as_ref(),
        Some(event.total_elapsed_ms),
    );
    events.push(event);
}

fn write_header_path_compare_run_history(
    log_dir: &Path,
    run_id: &str,
    status: &str,
    started_at: &str,
    total_elapsed_ms: u128,
    progress: &HeaderPathCompareProgress,
    error_type: &str,
    events: &[HeaderPathCompareRunEvent],
) -> Result<PathBuf> {
    let mut actual_engines = Vec::<String>::new();
    for event in events {
        if let Some(engine) = event.actual_engine.as_ref() {
            if !actual_engines.contains(engine) {
                actual_engines.push(engine.clone());
            }
        }
    }
    let report = serde_json::json!({
        "schema_version": 1,
        "product": "pytools",
        "feature_id": "PT-1-1",
        "run_id": run_id,
        "status": status,
        "started_at": started_at,
        "total_elapsed_ms": total_elapsed_ms,
        "progress": progress,
        "actual_engines": actual_engines,
        "error_type": if error_type.is_empty() { serde_json::Value::Null } else { serde_json::json!(error_type) },
        "events": events,
    });
    let path = log_dir.join(format!("PT-1-1_{run_id}.run.json"));
    write_json_report_create_new(&path, &report)?;
    Ok(path)
}

fn header_path_compare_run_id() -> String {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_nanos())
        .unwrap_or_default();
    format!("pytools-pt-1-1-{}-{nonce}", std::process::id())
}

fn run_wide_summary_cli(arguments: Vec<String>, row_suffix: bool) -> Result<()> {
    let report = run_wide_summary_cli_report(arguments, row_suffix)?;
    println!("{}", serde_json::to_string(&report)?);
    Ok(())
}

fn run_wide_summary_cli_report(
    arguments: Vec<String>,
    row_suffix: bool,
) -> Result<serde_json::Value> {
    let options = parse_cli_arguments(arguments)?;
    let run_started = Instant::now();
    let started_at = Local::now().to_rfc3339();
    let config = required_path(&options, "config")?;
    let global = load_global(&config)?;
    let log_view = options
        .values
        .get("log-view")
        .map(String::as_str)
        .unwrap_or("user");
    anyhow::ensure!(
        matches!(log_view, "user" | "developer" | "off"),
        "--log-view must be user, developer, or off"
    );
    let log_dir = options
        .values
        .get("log-dir")
        .map(PathBuf::from)
        .unwrap_or_else(|| global.log_dir.clone());
    if log_view == "user" {
        eprintln!(
            "{}",
            pytools_core::run_logging::terminal_line(
                pytools_core::run_logging::Level::Info,
                "开始宽表汇总，正在读取配置"
            )
        );
    }
    let rules = pytools_core::config::load_timeline_rules(&config)?;
    let maps = pytools_core::config::load_path_maps(&config)?;
    if log_view == "user" {
        eprintln!(
            "{}",
            pytools_core::run_logging::terminal_line(
                pytools_core::run_logging::Level::Info,
                &format!("已启用时序配置：{} 项", rules.len())
            )
        );
    }
    let capabilities = options.values.get("capabilities").map(PathBuf::from);
    let platform = options.values.get("platform").map(String::as_str);
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let feature_id = if row_suffix { "PT-1-3" } else { "1-1" };
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_nanos())
        .unwrap_or_default();
    let run_id = format!(
        "pytools-{}-{}-{nonce}",
        feature_id.to_lowercase(),
        std::process::id()
    );
    let mut events = Vec::<WideSummaryRunEvent>::new();
    let source_paths = if options.positional.is_empty() {
        let discovery_started = Instant::now();
        if log_view != "off" {
            let event = new_wide_summary_cli_event(
                &events,
                &run_started,
                feature_id,
                0,
                "source_discovery",
                "started",
                "正在扫描配置输入目录中的宽表源文件",
                None,
                0,
                &["pytools-core", "std::fs"],
                None,
                None,
            );
            emit_wide_summary_event(log_view, &mut events, event);
        }
        match discover_configured_source_files(&global) {
            Ok(sources) => {
                if log_view != "off" {
                    let event = new_wide_summary_cli_event(
                        &events,
                        &run_started,
                        feature_id,
                        sources.len(),
                        "source_discovery",
                        "completed",
                        format!("已发现 {} 个宽表源文件", sources.len()),
                        Some(discovery_started.elapsed().as_millis()),
                        0,
                        &["pytools-core", "std::fs"],
                        None,
                        None,
                    );
                    emit_wide_summary_event(log_view, &mut events, event);
                }
                Ok(sources)
            }
            Err(error) => {
                if log_view != "off" {
                    let event = new_wide_summary_cli_event(
                        &events,
                        &run_started,
                        feature_id,
                        0,
                        "source_discovery",
                        "failed",
                        "扫描宽表输入目录失败",
                        Some(discovery_started.elapsed().as_millis()),
                        0,
                        &["pytools-core", "std::fs"],
                        None,
                        Some("source_discovery_error"),
                    );
                    emit_wide_summary_event(log_view, &mut events, event);
                }
                Err(error.context("扫描宽表输入目录失败"))
            }
        }
    } else {
        Ok(options.positional.clone())
    };
    use pytools_core::run_logging::{EventWriter, FileStatus, Level, RunTracker, TaskStatus};
    let selected = source_paths.as_ref().ok().cloned().unwrap_or_default();
    let rule_source_matches = rules
        .iter()
        .map(|rule| {
            let matched = selected
                .iter()
                .filter(|path| {
                    let workbook_name = path.file_name().unwrap_or_default().to_string_lossy();
                    pytools_core::rule_engine::pick_rules_for_workbook(
                        &rules,
                        &workbook_name,
                        &global.timeline_rule_match_mode,
                    )
                    .iter()
                    .any(|candidate| candidate.row_no == rule.row_no)
                })
                .count();
            (rule.row_no, matched)
        })
        .collect::<Vec<_>>();
    if log_view == "user" {
        for rule in &rules {
            let rule_name = if rule.name.trim().is_empty() {
                "未命名规则"
            } else {
                rule.name.trim()
            };
            let sheet_keyword = if rule.sheet_keyword.trim().is_empty() {
                "（未限制）"
            } else {
                rule.sheet_keyword.trim()
            };
            let key_columns = if rule.target_dedup_cols.is_empty() {
                "（未配置）".to_owned()
            } else {
                rule.target_dedup_cols.join(", ")
            };
            let matched = rule_source_matches
                .iter()
                .find(|(row_no, _)| *row_no == rule.row_no)
                .map(|(_, matched)| *matched)
                .unwrap_or_default();
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Info,
                    &format!(
                        "正在检查时序提取规则：工作表「时序提取规则」第 {} 行；规则「{}」；工作簿关键字：{}；工作表关键字：{}；所选源文件匹配 {} 个；目标去重列：{}",
                        rule.row_no,
                        rule_name,
                        if rule.wb_keyword.trim().is_empty() { "（未限制）" } else { rule.wb_keyword.trim() },
                        sheet_keyword,
                        matched,
                        key_columns
                    )
                )
            );
        }
    }
    let mut tracker = RunTracker::with_id(feature_id, "时序宽表汇总", &selected, run_id.clone());
    if options.positional.is_empty() {
        tracker.set_discovered(selected.len());
    }
    let mut diagnostic = if log_view == "off" {
        None
    } else {
        match EventWriter::create(&log_dir.join("run-events"), feature_id, &run_id) {
            Ok(writer) => Some(writer),
            Err(error) => {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("无法创建诊断日志：{error}；业务处理继续")
                    )
                );
                None
            }
        }
    };
    let mut diagnostic_failed = false;
    if let Some(writer) = diagnostic.as_mut() {
        if let Err(error) = writer.write(&tracker.event(
            Level::Info,
            "run_started",
            "configuration",
            "started",
            "开始时序宽表汇总",
        )) {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Warn,
                    &format!("诊断日志写入失败：{error}")
                )
            );
            diagnostic_failed = true;
        }
    }
    let mut matched_before = 0usize;
    let wide_result = match source_paths {
        Ok(source_paths) => run_wide_summary_with_progress(
            &source_paths,
            &rules,
            &maps,
            &global,
            row_suffix,
            |mut event| {
                if let Some(index) = event.source_input_index {
                    match (
                        event.phase.as_str(),
                        event.status.as_str(),
                        tracker.state(index),
                    ) {
                        ("source_filter", "skipped", Some(FileStatus::Pending)) => {
                            let _ = tracker.transition(index, FileStatus::Skipped);
                        }
                        ("source_read", "started", Some(FileStatus::Pending)) => {
                            matched_before = event.progress.matched_rule_sheets;
                            let _ = tracker.transition(index, FileStatus::InProgress);
                        }
                        (_, "failed", Some(FileStatus::InProgress)) => {
                            let _ = tracker.transition(index, FileStatus::Failed);
                        }
                        ("rule_aggregation", "completed", Some(FileStatus::InProgress)) => {
                            let state = if event.progress.matched_rule_sheets > matched_before {
                                FileStatus::Succeeded
                            } else {
                                FileStatus::Skipped
                            };
                            let _ = tracker.transition(index, state);
                        }
                        _ => {}
                    }
                }
                event.sequence = events.len() + 1;
                event.total_elapsed_ms = run_started.elapsed().as_millis();
                if let Some(writer) = diagnostic.as_mut() {
                    let record = wide_diagnostic_event(&tracker, &event);
                    if let Err(error) = writer.write(&record) {
                        if !diagnostic_failed {
                            eprintln!(
                                "{}",
                                pytools_core::run_logging::terminal_line(
                                    Level::Warn,
                                    &format!("诊断日志写入失败：{error}")
                                )
                            );
                            diagnostic_failed = true;
                        }
                    }
                }
                emit_wide_summary_event(log_view, &mut events, event);
            },
        ),
        Err(error) => Err(error),
    };
    let total_inputs = events
        .last()
        .map(|event| event.progress.total_inputs)
        .unwrap_or_default();
    let output = wide_result.as_ref().ok().cloned().flatten();
    let result = wide_result.map(|_| ());

    if result.is_err()
        && log_view != "off"
        && !events.last().is_some_and(|event| event.status == "failed")
    {
        let event = new_wide_summary_cli_event(
            &events,
            &run_started,
            feature_id,
            total_inputs,
            "wide_summary",
            "failed",
            "宽表汇总失败",
            None,
            0,
            &["pytools-core"],
            None,
            Some("wide_summary_error"),
        );
        emit_wide_summary_event(log_view, &mut events, event);
    }

    let (target_writes, target_write_status) = match (result.is_ok(), output.as_deref()) {
        (true, Some(summary_path)) if !row_suffix => execute_wide_summary_target_writes(
            summary_path,
            &rules,
            &global,
            capabilities.as_deref(),
            platform,
            log_view,
            &run_started,
            feature_id,
            &mut events,
        ),
        (true, Some(_)) => (Vec::new(), "not_applicable".to_owned()),
        (true, None) => (Vec::new(), "not_run".to_owned()),
        (false, _) => (Vec::new(), "not_run".to_owned()),
    };
    let diagnostic_path = diagnostic_path_for_log_view(diagnostic.as_ref(), log_view);
    let mut full_output_paths = output
        .as_deref()
        .map(Path::to_path_buf)
        .into_iter()
        .collect::<Vec<_>>();
    for item in &target_writes {
        if item["status"] == "succeeded" {
            if let Some(path) = item["target"].as_str() {
                let path = PathBuf::from(path);
                if !full_output_paths.iter().any(|existing| existing == &path) {
                    full_output_paths.push(path);
                }
            }
        }
    }
    let output_paths = full_output_paths
        .iter()
        .map(|path| path_for_log_view(path, log_view))
        .collect::<Vec<_>>();
    let critical_failure =
        result.is_err() || matches!(target_write_status.as_str(), "failed" | "blocked");
    let mut unified = tracker.finish(
        critical_failure,
        false,
        output.is_none() && result.is_ok(),
        output_paths,
        diagnostic_path,
    );
    if target_write_status == "partial" {
        unified.overall_status = TaskStatus::PartialSuccess;
    }
    let all_outputs_saved = output.is_some()
        && !matches!(
            target_write_status.as_str(),
            "partial" | "failed" | "blocked"
        );
    unified.output_stats = Some(pytools_core::run_logging::OutputStats {
        workbooks_saved: Some(unified.output_paths.len()),
        sheets_saved: None,
        saved: Some(all_outputs_saved),
    });
    if critical_failure {
        if let Some(event) = events.iter().rev().find(|event| event.status == "failed") {
            unified.error_summary = Some(pytools_core::run_logging::ErrorSummary {
                stage: event.phase.clone(),
                code: summary_event_error_code(event.error_type.as_deref(), &event.phase),
                message: if event.phase == "target_write" {
                    target_writes
                        .iter()
                        .rev()
                        .find(|item| item["status"] == "failed")
                        .and_then(|item| item["error"].as_str())
                        .map(safe_wide_target_error)
                        .unwrap_or_else(|| {
                            "目标工作簿写入失败；请检查目标文件、工作表和写入权限".to_owned()
                        })
                } else {
                    "宽表汇总未完成；请检查输入文件和配置".to_owned()
                },
                file_name: event
                    .workbook_name
                    .as_deref()
                    .map(|name| pytools_core::run_logging::portable_file_name(Path::new(name))),
                sheet_name: event.worksheet_name.clone(),
            });
        }
    }
    if let Err(error) = result.as_ref() {
        let event = events.iter().rev().find(|event| event.status == "failed");
        let stage = event
            .map(|event| event.phase.as_str())
            .unwrap_or("wide_summary");
        if let Some(writer) = diagnostic.as_mut() {
            let mut failure = tracker.event(
                Level::Error,
                "operation_failed",
                stage,
                "failed",
                "宽表汇总未完成；请检查失败阶段、文件和配置",
            );
            failure.error_code = Some(summary_event_error_code(
                event.and_then(|event| event.error_type.as_deref()),
                stage,
            ));
            failure.file_name = event
                .and_then(|event| event.workbook_name.as_deref())
                .map(|name| pytools_core::run_logging::portable_file_name(Path::new(name)));
            failure.sheet_name = event.and_then(|event| event.worksheet_name.clone());
            if log_view == "developer" {
                failure.details = Some(serde_json::json!({"error_chain": format!("{error:#}")}));
            }
            if let Err(write_error) = writer.write(&failure) {
                if !diagnostic_failed {
                    eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            Level::Warn,
                            &format!("诊断日志写入失败：{write_error}；业务结果仍保留")
                        )
                    );
                    diagnostic_failed = true;
                }
            }
        }
    }
    if let Some(writer) = diagnostic.as_mut() {
        for event in events.iter().filter(|event| event.phase == "target_write") {
            let mut record = wide_diagnostic_event(&tracker, event);
            if log_view == "developer" && event.status == "failed" {
                record.details = Some(serde_json::json!({"diagnostic": event.user_message}));
            }
            if let Err(error) = writer.write(&record) {
                if !diagnostic_failed {
                    eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            Level::Warn,
                            &format!("诊断日志写入失败：{error}")
                        )
                    );
                    diagnostic_failed = true;
                }
                break;
            }
        }
        let event = tracker.event(
            if critical_failure {
                Level::Error
            } else {
                Level::Info
            },
            "run_finished",
            "run",
            &format!("{:?}", unified.overall_status),
            "时序宽表汇总结束",
        );
        if let Err(error) = writer.write(&event) {
            if !diagnostic_failed {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("诊断日志写入失败：{error}")
                    )
                );
            }
        }
    }

    let progress = events
        .last()
        .map(|event| event.progress.clone())
        .unwrap_or_default();
    if log_view == "user" {
        for rule in &rules {
            let completed_sheets = events
                .iter()
                .filter(|event| {
                    event.phase == "worksheet_aggregation"
                        && event.status == "completed"
                        && event.config_row == Some(rule.row_no)
                })
                .count();
            let matched_sources = rule_source_matches
                .iter()
                .find(|(row_no, _)| *row_no == rule.row_no)
                .map(|(_, matched)| *matched)
                .unwrap_or_default();
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    if completed_sheets == 0 { Level::Warn } else { Level::Info },
                    &format!(
                        "配置结果：工作表「时序提取规则」第 {} 行「{}」；匹配源文件 {} 个；完成工作表 {} 个；汇总文件生成并保存：{}；配置目标整体写入：{}",
                        rule.row_no,
                        if rule.name.trim().is_empty() { "未命名规则" } else { rule.name.trim() },
                        matched_sources,
                        completed_sheets,
                        if output.is_some() { "是" } else { "否" },
                        target_write_status
                    )
                )
            );
        }
    }
    let status = if result.is_err() {
        "failed"
    } else if progress.failed_inputs > 0
        || matches!(
            target_write_status.as_str(),
            "partial" | "failed" | "blocked"
        )
    {
        "partial"
    } else if output.is_some() {
        "succeeded"
    } else {
        "skipped"
    };

    let mut history_events = events.clone();
    let mut history_target_writes = target_writes.clone();
    if log_view != "developer" {
        for event in &mut history_events {
            if let Some(name) = event.workbook_name.take() {
                event.workbook_name = Some(pytools_core::run_logging::portable_file_name(
                    Path::new(&name),
                ));
            }
            if event.status == "failed" {
                let message = match event.phase.as_str() {
                    "source_read" => "源工作簿读取失败；请检查文件是否可打开".to_owned(),
                    "summary_write" => "汇总工作簿写入失败；请检查输出目录和权限".to_owned(),
                    "target_write" => "目标工作簿写入失败；请检查目标文件、工作表和权限".to_owned(),
                    _ => "宽表汇总阶段失败；请检查运行配置".to_owned(),
                };
                event.user_message = pytools_core::run_logging::workbook_sheet_context_message(
                    &message,
                    event.workbook_name.as_deref().map(Path::new),
                    event.worksheet_name.as_deref(),
                );
            }
        }
        for item in &mut history_target_writes {
            if let Some(target) = item
                .get("target")
                .and_then(serde_json::Value::as_str)
                .map(str::to_owned)
            {
                item["target"] =
                    serde_json::Value::String(path_for_log_view(Path::new(&target), log_view));
            }
            if item["status"] == "failed" {
                let safe_error = item["error"]
                    .as_str()
                    .map(safe_wide_target_error)
                    .unwrap_or_else(|| "目标工作簿写入失败；请检查文件、工作表和权限".to_owned());
                item["error"] = serde_json::Value::String(safe_error);
            }
        }
    }

    let run_log_path = if log_view == "off" {
        None
    } else {
        let progress = events
            .last()
            .map(|event| event.progress.clone())
            .unwrap_or_else(|| WideSummaryProgress {
                total_inputs,
                ..WideSummaryProgress::default()
            });
        let status = if result.is_err() {
            "failed"
        } else if progress.failed_inputs > 0
            || matches!(
                target_write_status.as_str(),
                "partial" | "failed" | "blocked"
            )
        {
            "partial"
        } else if output.is_some() {
            "succeeded"
        } else {
            "skipped"
        };
        let mut actual_engines = Vec::<String>::new();
        for engine in events
            .iter()
            .filter_map(|event| event.actual_engine.as_ref())
        {
            if !actual_engines.contains(engine) {
                actual_engines.push(engine.clone());
            }
        }
        let failure = result.as_ref().err().map(|_| {
            let event = events.last();
            serde_json::json!({
                "stage": event.map(|event| event.phase.as_str()).unwrap_or("wide_summary"),
                "input_index": event.and_then(|event| event.input_index),
                "error_type": event.and_then(|event| event.error_type.as_deref()).unwrap_or("wide_summary_error")
            })
        });
        let path = log_dir.join(format!("{feature_id}_{run_id}.json"));
        let output_display = output
            .as_deref()
            .map(|path| path_for_log_view(path, log_view));
        let report = serde_json::json!({
            "schema_version": 1,
            "product": "pytools",
            "feature_id": feature_id,
            "function": if row_suffix { "wide-summary-row-suffix" } else { "wide-summary" },
            "run_id": &run_id,
            "started_at": &started_at,
            "status": status,
            "overall_status": unified.overall_status,
            "feature_name": unified.feature_name,
            "ended_at": unified.ended_at,
            "input_stats": unified.input_stats,
            "inputs": unified.inputs,
            "output_stats": unified.output_stats,
            "output_paths": unified.output_paths,
            "diagnostic_log_path": unified.diagnostic_log_path,
            "total_inputs": progress.total_inputs,
            "completed_inputs": progress.completed_inputs,
            "failed_inputs": progress.failed_inputs,
            "matched_rule_sheets": progress.matched_rule_sheets,
            "skipped_inputs": progress.skipped_inputs,
            "skipped_worksheets": progress.skipped_worksheets,
            "extracted_cells": progress.extracted_cells,
            "output_rows": progress.output_rows,
            "output": output_display,
            "target_write_status": &target_write_status,
            "target_writes": &history_target_writes,
            "actual_engines": actual_engines,
            "failure": failure,
            "elapsed_ms": run_started.elapsed().as_millis(),
            "total_elapsed_ms": run_started.elapsed().as_millis(),
            "events": &history_events,
        });
        match write_json_report_create_new(&path, &report) {
            Ok(()) => Some(path),
            Err(error) => {
                eprintln!("pytools: warning: could not save {feature_id} run history: {error:#}");
                None
            }
        }
    };

    match result {
        Ok(()) => {
            let mut result = serde_json::json!({});
            result["output"] = output
                .as_deref()
                .map(display_path_for_user)
                .map(serde_json::Value::String)
                .unwrap_or(serde_json::Value::Null);
            result["target_write_status"] = serde_json::Value::String(target_write_status);
            result["target_writes"] = serde_json::Value::Array(target_writes);
            result["status"] = serde_json::json!(status);
            result["overall_status"] = serde_json::json!(unified.overall_status);
            result["input_stats"] = serde_json::to_value(&unified.input_stats)?;
            result["total_inputs"] = serde_json::json!(progress.total_inputs);
            result["completed_inputs"] = serde_json::json!(progress.completed_inputs);
            result["failed_inputs"] = serde_json::json!(progress.failed_inputs);
            result["skipped_inputs"] = serde_json::json!(progress.skipped_inputs);
            result["total_elapsed_ms"] = serde_json::json!(run_started.elapsed().as_millis());
            if log_view != "off" {
                result["run_id"] = serde_json::Value::String(run_id);
                if let Some(path) = run_log_path {
                    let display_path = display_path_for_user(&path);
                    result["run_log_path"] = serde_json::Value::String(display_path.clone());
                    if log_view == "user" {
                        eprintln!(
                            "{} 运行记录已保存：{display_path}",
                            pytools_core::engine_diag::now_prefix()
                        );
                    }
                }
            }
            Ok(result)
        }
        Err(error) => {
            if log_view == "user" {
                if let Some(path) = run_log_path {
                    eprintln!(
                        "{} 运行记录已保存：{}",
                        pytools_core::engine_diag::now_prefix(),
                        display_path_for_user(&path)
                    );
                }
                let stats = &unified.input_stats;
                eprintln!("{}", pytools_core::run_logging::terminal_line(Level::Error, &format!("宽表汇总失败：选择 {} 个，尝试 {} 个，成功 {} 个，失败 {} 个，跳过 {} 个，未完成 {} 个；运行编号 {run_id}", stats.selected, stats.attempted, stats.succeeded, stats.failed, stats.skipped, stats.pending)));
                let _ = error;
                Err(anyhow::anyhow!(
                    "宽表汇总失败，请检查失败阶段、文件和配置；运行编号 {run_id}"
                ))
            } else {
                Err(error)
            }
        }
    }
}

fn wide_dump_used_bounds(sheet: &excel_reader::SheetDump) -> (u32, u32) {
    let (mut max_row, mut max_column) = (0usize, 0usize);
    for address in sheet.values.keys().chain(sheet.formulas.keys()) {
        if let Some((row, column)) = parse_wide_a1(address) {
            max_row = max_row.max(row);
            max_column = max_column.max(column);
        }
    }
    (max_column as u32, max_row as u32)
}

fn wide_dump_cell_display(sheet: &excel_reader::SheetDump, column: u32, row: u32) -> String {
    let address = pytools_core::rule_engine::cell_address(row as usize, column as usize);
    let Some(value) = sheet.values.get(&address) else {
        return sheet
            .formulas
            .get(&address)
            .map(|formula| format!("={formula}"))
            .unwrap_or_default();
    };
    match value.get("type").and_then(serde_json::Value::as_str) {
        Some("number" | "integer" | "datetime_serial") => value
            .get("value")
            .and_then(serde_json::Value::as_f64)
            .map(|number| {
                if number.fract() == 0.0 {
                    format!("{number:.0}")
                } else {
                    number.to_string()
                }
            })
            .unwrap_or_default(),
        Some("boolean") => value
            .get("value")
            .map(ToString::to_string)
            .unwrap_or_default(),
        _ => value
            .get("value")
            .and_then(serde_json::Value::as_str)
            .unwrap_or_default()
            .trim()
            .to_owned(),
    }
}

fn parse_wide_a1(address: &str) -> Option<(usize, usize)> {
    let split = address
        .char_indices()
        .find_map(|(index, character)| character.is_ascii_digit().then_some(index))?;
    let (letters, digits) = address.split_at(split);
    if letters.is_empty()
        || !letters
            .chars()
            .all(|character| character.is_ascii_alphabetic())
    {
        return None;
    }
    let column = letters.chars().try_fold(0usize, |value, character| {
        value
            .checked_mul(26)?
            .checked_add(character.to_ascii_uppercase() as usize - 'A' as usize + 1)
    })?;
    Some((digits.parse().ok()?, column))
}

/// Append selected rows from the Rust-built 1-1 summary workbook to an
/// existing target through the Openpyxl worker, then publish with a verified
/// snapshot and backup. The source target is never edited in place by Python.
fn commit_wide_summary_existing_target_openpyxl(
    summary_path: &Path,
    target: &Path,
    expected_hash: &str,
    summary_sheet: &str,
    target_sheet: &str,
    summary_max_col: u32,
    write_start_row: u32,
    include_header: bool,
    create_sheet: bool,
    kept_rows: &[u32],
) -> Result<()> {
    use excel_editor::{
        EditOperation, ExcelEditPlan, PreservationRequirement, RangeStyleMode, WorkbookFormat,
    };

    anyhow::ensure!(summary_max_col > 0, "汇总工作表没有可写入的列");
    let profile = excel_editor::inspect_xlsx_profile(target)
        .with_context(|| format!("检查既有目标工作簿失败：{}", target.display()))?;
    anyhow::ensure!(
        profile.inventory_complete,
        "目标工作簿结构清单不完整，拒绝使用 openpyxl 修改"
    );
    let mut operations = Vec::new();
    if create_sheet {
        operations.push(EditOperation::CreateSheet {
            name: target_sheet.to_owned(),
        });
    }
    let mut destination_row = write_start_row;
    let mut push_import = |source_start: u32,
                           source_end: u32,
                           destination_start: u32|
     -> Result<()> {
        let source_start_cell = pytools_core::rule_engine::cell_address(source_start as usize, 1);
        let source_end_cell =
            pytools_core::rule_engine::cell_address(source_end as usize, summary_max_col as usize);
        operations.push(EditOperation::ImportRange {
            source_workbook: summary_path.to_path_buf(),
            source_sheet: summary_sheet.to_owned(),
            source_range: format!("{source_start_cell}:{source_end_cell}"),
            target_sheet: target_sheet.to_owned(),
            target_start: format!("A{destination_start}"),
            copy_values: true,
            copy_formulas: true,
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
            materialize_blank_cells: false,
            hide_zero: false,
        });
        Ok(())
    };
    if include_header {
        push_import(1, 1, destination_row)?;
        destination_row += 1;
    }
    let mut rows = kept_rows.iter().copied();
    if let Some(mut range_start) = rows.next() {
        let mut range_end = range_start;
        for row in rows {
            if row == range_end + 1 {
                range_end = row;
                continue;
            }
            push_import(range_start, range_end, destination_row)?;
            destination_row += range_end - range_start + 1;
            range_start = row;
            range_end = row;
        }
        push_import(range_start, range_end, destination_row)?;
    }
    drop(push_import);
    anyhow::ensure!(!operations.is_empty(), "目标追加计划没有实际写入区域");

    let plan = ExcelEditPlan {
        request_id: format!(
            "rusttools-1-1-openpyxl-{}",
            target.file_stem().unwrap_or_default().to_string_lossy()
        ),
        input: target.to_path_buf(),
        output: target.to_path_buf(),
        format: WorkbookFormat::Xlsx,
        workbook: profile,
        operations,
        preserve: std::collections::BTreeSet::from([
            PreservationRequirement::CellValuesAndTypes,
            PreservationRequirement::FormulaText,
            PreservationRequirement::Styles,
            PreservationRequirement::CommentsAndPositions,
            PreservationRequirement::MergedRanges,
            PreservationRequirement::ConditionalFormatting,
            PreservationRequirement::DataValidations,
            PreservationRequirement::DefinedNames,
            PreservationRequirement::RowAndColumnDimensions,
            PreservationRequirement::PageSettings,
            PreservationRequirement::HeaderFooterText,
            PreservationRequirement::Relationships,
            PreservationRequirement::SheetIdentityAndOrder,
        ]),
        drop_allowed: Default::default(),
    };
    plan.validate().context("校验 1-1 Openpyxl 追加计划失败")?;
    verify_planning_snapshot(target, Some(expected_hash))
        .context("目标工作簿在规划期间发生变化，未执行追加")?;

    let candidate = wide_target_candidate_path(target)?;
    let mut plan = plan;
    plan.output = candidate.path().to_path_buf();
    plan.validate().context("校验 1-1 Openpyxl 候选路径失败")?;
    pytools_core::inplace_edit::execute_candidate(
        &plan,
        pytools_core::inplace_edit::InPlaceEngine::Openpyxl,
    )
    .context("openpyxl 生成并校验目标候选文件失败")?;
    replace_target_from_candidate(candidate.path(), target, expected_hash)
        .context("校验并安全替换目标工作簿失败")?;
    Ok(())
}

fn wide_target_candidate_path(target: &Path) -> Result<TemporaryFile> {
    let parent = target
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let stem = target.file_stem().unwrap_or_default().to_string_lossy();
    let stamp = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
    for index in 0..1000 {
        let path = parent.join(format!(
            ".{stem}.rusttools-wide-{}-{stamp}-{index}.xlsx",
            std::process::id()
        ));
        if !path.exists() {
            return Ok(TemporaryFile::new(path));
        }
    }
    bail!("无法为目标工作簿分配临时候选路径：{}", target.display())
}

fn execute_wide_summary_target_writes(
    summary_path: &Path,
    rules: &[pytools_core::config::TimelineRule],
    _global: &GlobalConfig,
    _capabilities: Option<&Path>,
    _platform: Option<&str>,
    _log_view: &str,
    run_started: &Instant,
    feature_id: &str,
    events: &mut Vec<WideSummaryRunEvent>,
) -> (Vec<serde_json::Value>, String) {
    // Calamine reads source/target values; Openpyxl edits existing targets;
    // Rust XlsxWriter creates new targets.
    let stage_started = Instant::now();
    let total_inputs = events
        .last()
        .map(|event| event.progress.total_inputs)
        .unwrap_or_default();
    let targets: Vec<&pytools_core::config::TimelineRule> = rules
        .iter()
        .filter(|rule| {
            rule.target_write_enabled
                && rule.target_wb_path.is_some()
                && rule
                    .target_sheet
                    .as_deref()
                    .is_some_and(|sheet| !sheet.trim().is_empty())
        })
        .collect();
    if targets.is_empty() {
        let enabled = rules
            .iter()
            .filter(|rule| rule.target_write_enabled)
            .count();
        let mut event = new_wide_summary_cli_event(
            events,
            run_started,
            feature_id,
            total_inputs,
            "target_write",
            "skipped",
            format!("没有可写入的目标：{} 条规则启用了目标写入；请检查目标工作簿路径和目标工作表是否填写，汇总文件已生成", enabled),
            Some(stage_started.elapsed().as_millis()),
            0,
            &["pytools-core"],
            None,
            Some("no_intermediate_target_configured"),
        );
        event.workbook_name = summary_path
            .file_name()
            .map(|name| name.to_string_lossy().into_owned());
        emit_wide_summary_event(_log_view, events, event);
        return (Vec::new(), "not_configured".to_owned());
    }
    let summary_book = match excel_reader::read_workbook(summary_path) {
        Ok(book) => book,
        Err(error) => {
            let mut event = new_wide_summary_cli_event(
                events,
                run_started,
                feature_id,
                total_inputs,
                "target_write",
                "failed",
                format!(
                    "读取汇总文件 {} 失败：{error:#}",
                    display_path_for_user(&summary_path)
                ),
                Some(stage_started.elapsed().as_millis()),
                0,
                &["pytools-core", "calamine"],
                None,
                Some("target_write_error"),
            );
            event.workbook_name = summary_path
                .file_name()
                .map(|name| name.to_string_lossy().into_owned());
            emit_wide_summary_event(_log_view, events, event);
            return (
                vec![serde_json::json!({
                    "target": display_path_for_user(&summary_path).to_string(),
                    "status": "failed",
                    "error": format!("读取汇总文件失败：{error:#}"),
                })],
                "failed".to_owned(),
            );
        }
    };
    let mut results = Vec::new();
    let mut any_ok = false;
    let mut any_failed = false;
    for rule in targets {
        let Some(target) = rule.target_wb_path.as_ref() else {
            continue;
        };
        let Some(target_sheet_name) = rule
            .target_sheet
            .as_deref()
            .map(|sheet| sheet.trim().chars().take(31).collect::<String>())
        else {
            continue;
        };
        let parent = target
            .parent()
            .filter(|parent| !parent.as_os_str().is_empty())
            .map(|parent| parent.to_path_buf());
        if let Some(parent) = &parent {
            if !parent.is_dir() {
                any_failed = true;
                let message = format!("目标路径文件夹不存在：{}", display_path_for_user(&parent));
                let mut event = new_wide_summary_cli_event(
                    events,
                    run_started,
                    feature_id,
                    total_inputs,
                    "target_write",
                    "failed",
                    message.clone(),
                    Some(stage_started.elapsed().as_millis()),
                    0,
                    &["pytools-core"],
                    None,
                    Some("target_folder_missing"),
                );
                event.workbook_name = Some(target.to_string_lossy().into_owned());
                emit_wide_summary_event(_log_view, events, event);
                results.push(serde_json::json!({
                    "target": display_path_for_user(&target).to_string(),
                    "sheet": target_sheet_name,
                    "status": "failed",
                    "error": message,
                }));
                continue;
            }
        }
        let existed = target.is_file();
        let original_hash = if existed {
            match sha256_file(target) {
                Ok(hash) => Some(hash),
                Err(error) => {
                    any_failed = true;
                    let mut event = new_wide_summary_cli_event(
                        events,
                        run_started,
                        feature_id,
                        total_inputs,
                        "target_write",
                        "failed",
                        "读取目标工作簿快照失败；未执行写入".to_owned(),
                        Some(stage_started.elapsed().as_millis()),
                        0,
                        &["pytools-core", "sha2"],
                        None,
                        Some("target_snapshot_error"),
                    );
                    event.workbook_name =
                        Some(pytools_core::run_logging::portable_file_name(target));
                    emit_wide_summary_event(_log_view, events, event);
                    results.push(serde_json::json!({"target": display_path_for_user(&target).to_string(), "status": "failed", "error": format!("读取目标文件快照失败：{error:#}")}));
                    continue;
                }
            }
        } else {
            None
        };
        let target_book = if existed {
            match excel_reader::read_workbook(target) {
                Ok(book) => Some(book),
                Err(error) => {
                    any_failed = true;
                    let message = format!("打开既有目标工作簿失败：{error:#}");
                    let mut event = new_wide_summary_cli_event(
                        events,
                        run_started,
                        feature_id,
                        total_inputs,
                        "target_write",
                        "failed",
                        format!("目标工作簿 {}：{message}", display_path_for_user(&target)),
                        Some(stage_started.elapsed().as_millis()),
                        0,
                        &["pytools-core", "calamine"],
                        None,
                        Some("target_write_error"),
                    );
                    event.workbook_name = Some(target.to_string_lossy().into_owned());
                    emit_wide_summary_event(_log_view, events, event);
                    results.push(serde_json::json!({
                        "target": display_path_for_user(&target).to_string(),
                        "status": "failed",
                        "error": message,
                    }));
                    continue;
                }
            }
        } else {
            None
        };
        let existed_sheet = target_book.as_ref().is_some_and(|book| {
            book.sheets
                .iter()
                .any(|sheet| sheet.name == target_sheet_name)
        });
        // 每条规则只能写入对应的汇总工作表；名称冲突时拒绝猜测。
        let summary_sheet_name = pytools_core::sheet_extract::sanitize_sheet_name(&rule.name);
        let ambiguous = rules
            .iter()
            .filter(|other| {
                pytools_core::sheet_extract::sanitize_sheet_name(&other.name)
                    .eq_ignore_ascii_case(&summary_sheet_name)
            })
            .count()
            > 1;
        if ambiguous
            || !summary_book
                .sheets
                .iter()
                .any(|sheet| sheet.name == summary_sheet_name)
        {
            any_failed = true;
            let mut event = new_wide_summary_cli_event(
                events,
                run_started,
                feature_id,
                total_inputs,
                "target_write",
                "failed",
                "汇总工作表不存在或名称不唯一；未执行写入".to_owned(),
                Some(stage_started.elapsed().as_millis()),
                0,
                &["pytools-core", "calamine"],
                None,
                Some("summary_sheet_missing"),
            );
            event.workbook_name = Some(pytools_core::run_logging::portable_file_name(target));
            event.worksheet_name = Some(target_sheet_name.clone());
            emit_wide_summary_event(_log_view, events, event);
            results.push(serde_json::json!({
                "target": display_path_for_user(&target).to_string(),
                "sheet": target_sheet_name,
                "status": "failed",
                "error": format!("规则 {:?} 的汇总工作表不存在或名称不唯一，未写入目标", rule.name),
            }));
            continue;
        }
        let summary_sheet = summary_book
            .sheets
            .iter()
            .find(|sheet| sheet.name == summary_sheet_name)
            .expect("summary sheet resolved above");
        let (summary_max_col, summary_max_row) = wide_dump_used_bounds(summary_sheet);

        // 既有工作表：完全空白→视为待填充，直接写入表头+数据；
        // 否则先比对表头（列数 + 每列文本）：一致→在末尾追加数据行；
        // 不一致→拒绝写入并说明差异，绝不清空既有数据。
        let mut existing_target_rows: u32 = 0;
        let mut blank_existing_sheet = false;
        let (write_start_row, _first_source_row, mode) = if existed_sheet {
            let (target_max_col, target_max_row, mismatch) = {
                let target_ws = target_book
                    .as_ref()
                    .and_then(|book| {
                        book.sheets
                            .iter()
                            .find(|sheet| sheet.name == target_sheet_name)
                    })
                    .expect("target sheet resolved above");
                let (target_max_col, target_max_row) = wide_dump_used_bounds(target_ws);
                if target_max_col == 0 || target_max_row == 0 {
                    // 空白目标工作表视为待填充。
                    (0u32, 0u32, None)
                } else {
                    let mut mismatch = None;
                    if target_max_col != summary_max_col {
                        mismatch = Some(format!(
                            "表头列数不同：目标 {target_max_col} 列，汇总 {summary_max_col} 列"
                        ));
                    } else {
                        for column in 1..=summary_max_col {
                            let expected = wide_dump_cell_display(target_ws, column, 1);
                            let actual = wide_dump_cell_display(summary_sheet, column, 1);
                            if expected != actual {
                                mismatch = Some(format!(
                                    "第 {column} 列表头不同：目标 {expected:?}，汇总 {actual:?}"
                                ));
                                break;
                            }
                        }
                    }
                    (target_max_col, target_max_row, mismatch)
                }
            };
            let _ = target_max_col;
            if let Some(reason) = mismatch {
                any_failed = true;
                let message = format!(
                    "目标工作表 {target_sheet_name:?} 与汇总表头不一致，已拒绝写入（不清空既有数据）：{reason}"
                );
                let mut event = new_wide_summary_cli_event(
                    events,
                    run_started,
                    feature_id,
                    total_inputs,
                    "target_write",
                    "failed",
                    format!("目标工作簿 {}：{message}", display_path_for_user(&target)),
                    Some(stage_started.elapsed().as_millis()),
                    0,
                    &["pytools-core", "calamine"],
                    None,
                    Some("target_header_mismatch"),
                );
                event.workbook_name = Some(target.to_string_lossy().into_owned());
                event.worksheet_name = Some(target_sheet_name.clone());
                emit_wide_summary_event(_log_view, events, event);
                results.push(serde_json::json!({
                    "target": display_path_for_user(&target).to_string(),
                    "sheet": target_sheet_name,
                    "status": "failed",
                    "error": message,
                }));
                continue;
            }
            if target_max_row == 0 {
                // 空白工作表：在既有空 sheet 上写入表头+数据，其他 sheet 不受影响。
                blank_existing_sheet = true;
                (1u32, 1u32, "空白工作表写入")
            } else {
                // 表头一致：跳过汇总表头行，追加到目标末行之后。
                existing_target_rows = target_max_row;
                (target_max_row + 1, 2u32, "表头一致追加")
            }
        } else {
            (1u32, 1u32, "新建工作表写入")
        };
        // 追加语义仅在“既有非空白工作表”上成立；新表/空白表都由本次写入表头。
        let append_to_existing = existed_sheet && !blank_existing_sheet;

        // 对齐旧 Python：写入前先做两级去重——键列默认“数据日期+集合表头+行头路径/首个动态列”，
        // 可由规则“目标去重列”显式指定；批内重复行与目标既有行都不再写入。
        let summary_headers: Vec<String> = (1..=summary_max_col)
            .map(|column| wide_dump_cell_display(summary_sheet, column, 1))
            .collect();
        let key_names =
            match pytools_core::timeline_ops::wide_target_key_names(rule, &summary_headers) {
                Ok(names) => names,
                Err(error) => {
                    // 与 Python 一致：去重列不存在时跳过目标写入，不做猜测。
                    let message = format!("解析目标去重列失败，已跳过目标写入：{error:#}");
                    let mut event = new_wide_summary_cli_event(
                        events,
                        run_started,
                        feature_id,
                        total_inputs,
                        "target_write",
                        "skipped",
                        format!("目标工作簿 {}：{message}", display_path_for_user(&target)),
                        Some(stage_started.elapsed().as_millis()),
                        0,
                        &["pytools-core"],
                        None,
                        Some("target_dedup_columns_missing"),
                    );
                    event.workbook_name = Some(target.to_string_lossy().into_owned());
                    event.worksheet_name = Some(target_sheet_name.clone());
                    emit_wide_summary_event(_log_view, events, event);
                    results.push(serde_json::json!({
                        "target": display_path_for_user(&target).to_string(),
                        "sheet": target_sheet_name,
                        "status": "skipped",
                        "error": message,
                    }));
                    continue;
                }
            };
        let key_columns: Vec<usize> = key_names
            .iter()
            .map(|name| {
                summary_headers
                    .iter()
                    .position(|header| header == name)
                    .expect("key names resolved against summary headers")
                    + 1
            })
            .collect();
        let row_key = |sheet: &excel_reader::SheetDump, row: u32| -> Vec<String> {
            key_columns
                .iter()
                .map(|&column| wide_dump_cell_display(sheet, column as u32, row))
                .collect()
        };
        let mut seen_keys: std::collections::HashSet<Vec<String>> =
            std::collections::HashSet::new();
        if existed_sheet {
            let target_ws = target_book
                .as_ref()
                .and_then(|book| {
                    book.sheets
                        .iter()
                        .find(|sheet| sheet.name == target_sheet_name)
                })
                .expect("target sheet resolved above");
            for row in 2..=existing_target_rows {
                seen_keys.insert(row_key(target_ws, row));
            }
        }
        let mut batch_keys: std::collections::HashSet<Vec<String>> =
            std::collections::HashSet::new();
        let mut batch_dedup_rows = 0usize;
        let mut existing_filtered_rows = 0usize;
        let mut kept_rows = Vec::<u32>::new();
        // 去重始终只针对数据行（第 2 行起）；表头行由下方写入逻辑单独处理。
        for row in 2..=summary_max_row {
            let key = row_key(summary_sheet, row);
            if !batch_keys.insert(key.clone()) {
                batch_dedup_rows += 1;
                continue;
            }
            if seen_keys.contains(&key) {
                existing_filtered_rows += 1;
                continue;
            }
            seen_keys.insert(key);
            kept_rows.push(row);
        }
        let input_rows = summary_max_row.saturating_sub(1) as usize;
        if append_to_existing && kept_rows.is_empty() {
            // 目标去重后没有新增行：不生成候选、不改写目标文件。
            let message = format!(
                "目标工作簿 {}：表头一致，输入 {input_rows} 行，批内去重 {batch_dedup_rows} 行，批内去重后与目标去重键相同 {existing_filtered_rows} 行；实际追加到目标 0 行，没有新增数据，未修改目标",
                display_path_for_user(&target)
            );
            let mut event = new_wide_summary_cli_event(
                events,
                run_started,
                feature_id,
                total_inputs,
                "target_write",
                "skipped",
                message.clone(),
                Some(stage_started.elapsed().as_millis()),
                0,
                &["pytools-core", "calamine"],
                None,
                Some("no_new_rows_after_dedup"),
            );
            event.workbook_name = Some(target.to_string_lossy().into_owned());
            event.worksheet_name = Some(target_sheet_name.clone());
            emit_wide_summary_event(_log_view, events, event);
            results.push(serde_json::json!({
                "target": display_path_for_user(&target).to_string(),
                "sheet": target_sheet_name,
                "status": "skipped",
                "mode": mode,
                "input_rows": input_rows,
                "batch_dedup_rows": batch_dedup_rows,
                "existing_filtered_rows": existing_filtered_rows,
                "target_duplicate_rows": existing_filtered_rows,
                "written_rows": 0,
                "written_data_rows": 0,
                "cells": 0,
                "error": message,
            }));
            continue;
        }

        // 统计实际导入的非空/公式单元格；既有目标由 Openpyxl 导入区域。
        let mut rows_to_write =
            Vec::with_capacity(kept_rows.len() + usize::from(!append_to_existing));
        if !append_to_existing {
            rows_to_write.push(1);
        }
        rows_to_write.extend(kept_rows.iter().copied());
        let copied = rows_to_write
            .iter()
            .map(|row| {
                (1..=summary_max_col)
                    .filter(|column| {
                        let address = pytools_core::rule_engine::cell_address(
                            *row as usize,
                            *column as usize,
                        );
                        summary_sheet.values.contains_key(&address)
                            || summary_sheet.formulas.contains_key(&address)
                    })
                    .count()
            })
            .sum::<usize>();
        let publish = (|| -> Result<()> {
            if let Some(hash) = original_hash.as_deref() {
                commit_wide_summary_existing_target_openpyxl(
                    summary_path,
                    target,
                    hash,
                    &summary_sheet_name,
                    &target_sheet_name,
                    summary_max_col,
                    write_start_row,
                    !append_to_existing,
                    !existed_sheet,
                    &kept_rows,
                )?;
            } else {
                let candidate = wide_target_candidate_path(target)?;
                let writer_rows =
                    pytools_core::timeline_ops::write_wide_summary_new_target_candidate(
                        summary_path,
                        rule,
                        &target_sheet_name,
                        candidate.path(),
                    )?;
                anyhow::ensure!(
                    writer_rows == kept_rows.len(),
                    "Rust XlsxWriter 与目标去重规划的行数不一致：计划 {} 行，实际 {} 行",
                    kept_rows.len(),
                    writer_rows
                );
                create_target_from_candidate(candidate.path(), target)
                    .context("Rust XlsxWriter 候选验证后发布新目标工作簿失败")?;
            }
            Ok(())
        })();
        if let Err(error) = publish {
            any_failed = true;
            let message = format!("保存目标工作簿失败：{error:#}");
            let mut event = new_wide_summary_cli_event(
                events,
                run_started,
                feature_id,
                total_inputs,
                "target_write",
                "failed",
                format!("目标工作簿 {}：{message}", display_path_for_user(&target)),
                Some(stage_started.elapsed().as_millis()),
                0,
                if existed {
                    &["pytools-core", "openpyxl"]
                } else {
                    &["pytools-core", "rust_xlsxwriter"]
                },
                None,
                Some("target_write_error"),
            );
            event.workbook_name = Some(target.to_string_lossy().into_owned());
            emit_wide_summary_event(_log_view, events, event);
            results.push(serde_json::json!({
                "target": display_path_for_user(&target).to_string(),
                "sheet": target_sheet_name,
                "status": "failed",
                "error": message,
            }));
            continue;
        }
        any_ok = true;
        let mode = if !existed {
            "新建工作簿并写入"
        } else {
            mode
        };
        let written_rows = if append_to_existing {
            kept_rows.len()
        } else {
            kept_rows.len() + 1
        };
        let written_data_rows = kept_rows.len();
        let data_write_label = if append_to_existing {
            "实际追加到目标"
        } else {
            "实际写入目标数据"
        };
        let message = format!(
            "目标工作簿 {}：{mode}工作表 {target_sheet_name:?}（输入 {input_rows} 行，批内去重 {batch_dedup_rows} 行，批内去重后与目标去重键相同 {existing_filtered_rows} 行，{data_write_label} {written_data_rows} 行 / 写入 {copied} 个单元格）",
            display_path_for_user(&target)
        );
        let mut event = new_wide_summary_cli_event(
            events,
            run_started,
            feature_id,
            total_inputs,
            "target_write",
            "succeeded",
            message.clone(),
            Some(stage_started.elapsed().as_millis()),
            0,
            if existed {
                &["pytools-core", "openpyxl"]
            } else {
                &["pytools-core", "rust_xlsxwriter"]
            },
            Some(if existed {
                "openpyxl"
            } else {
                "rust_xlsxwriter"
            }),
            None,
        );
        event.workbook_name = Some(target.to_string_lossy().into_owned());
        event.worksheet_name = Some(target_sheet_name.clone());
        emit_wide_summary_event(_log_view, events, event);
        results.push(serde_json::json!({
            "target": display_path_for_user(&target).to_string(),
            "sheet": target_sheet_name,
            "status": "succeeded",
            "mode": mode,
            "cells": copied,
            "input_rows": input_rows,
            "batch_dedup_rows": batch_dedup_rows,
            "existing_filtered_rows": existing_filtered_rows,
            "target_duplicate_rows": existing_filtered_rows,
            "written_rows": written_rows,
            "written_data_rows": written_data_rows,
        }));
    }
    let status = if any_failed && any_ok {
        "partial".to_owned()
    } else if any_failed {
        "failed".to_owned()
    } else {
        "succeeded".to_owned()
    };
    (results, status)
}

fn display_path_for_user(path: &Path) -> String {
    let value = path.to_string_lossy();
    if let Some(unc_path) = value.strip_prefix(r"\\?\UNC\") {
        format!(r"\\{unc_path}")
    } else if let Some(drive_path) = value.strip_prefix(r"\\?\") {
        drive_path.to_owned()
    } else {
        value.into_owned()
    }
}

fn safe_wide_target_error(error: &str) -> String {
    if error.contains("表头不一致")
        || error.contains("表头列数不同")
        || error.contains("列表头不同")
    {
        "目标工作表表头与汇总不一致，已拒绝写入；既有目标数据未更改".to_owned()
    } else if error.contains("不存在或名称不唯一") || error.contains("汇总工作表") {
        "汇总工作表不存在或名称不唯一，未执行写入".to_owned()
    } else if error.contains("目标路径文件夹不存在") {
        "目标路径文件夹不存在；请检查目标路径配置".to_owned()
    } else if error.contains("快照") {
        "无法读取目标文件快照；未执行写入".to_owned()
    } else {
        "目标工作簿写入失败；请检查目标文件、工作表、文件占用状态和权限".to_owned()
    }
}

fn new_wide_summary_cli_event(
    events: &[WideSummaryRunEvent],
    run_started: &Instant,
    feature_id: &str,
    total_inputs: usize,
    phase: &str,
    status: &str,
    user_message: impl Into<String>,
    stage_elapsed_ms: Option<u128>,
    target_append_plans: usize,
    components: &[&str],
    actual_engine: Option<&str>,
    error_type: Option<&str>,
) -> WideSummaryRunEvent {
    let mut progress = events
        .last()
        .map(|event| event.progress.clone())
        .unwrap_or_default();
    progress.total_inputs = total_inputs;
    progress.current_input = None;
    progress.target_append_plans = target_append_plans;
    let mut recorded_components = vec!["Rust".to_owned()];
    recorded_components.extend(components.iter().map(|component| (*component).to_owned()));
    WideSummaryRunEvent {
        sequence: events.len() + 1,
        timestamp: Local::now().to_rfc3339(),
        feature_id: feature_id.to_owned(),
        phase: phase.to_owned(),
        status: status.to_owned(),
        user_message: user_message.into(),
        progress,
        input_index: None,
        source_input_index: None,
        workbook_name: None,
        worksheet_name: None,
        rule_name: None,
        config_row: None,
        stage_elapsed_ms,
        workbook_elapsed_ms: None,
        total_elapsed_ms: run_started.elapsed().as_millis(),
        components: recorded_components,
        actual_engine: actual_engine.map(str::to_owned),
        error_type: error_type.map(str::to_owned),
    }
}

fn wide_diagnostic_event(
    tracker: &pytools_core::run_logging::RunTracker,
    event: &WideSummaryRunEvent,
) -> pytools_core::run_logging::LogEvent {
    use pytools_core::run_logging::Level;
    let level = match event.status.as_str() {
        "failed" => Level::Error,
        "skipped" => Level::Warn,
        "started" | "running" => Level::Info,
        _ => Level::Debug,
    };
    let phase_message = match event.phase.as_str() {
        "source_discovery" => "查找符合时序规则的源工作簿",
        "source_match" => "匹配启用的时序规则与源工作簿",
        "source_read" => "读取时序源工作簿",
        "summary_write" => "写入时序宽表汇总文件",
        "target_write" => "将汇总数据写入配置指定的目标工作簿",
        "row_write" => "写入时序宽表数据行",
        "run" => "时序宽表汇总",
        other => other,
    };
    let message = if event.status == "failed" {
        format!("{phase_message}失败；请检查对应配置、源/目标工作簿和读写权限")
    } else {
        format!("{phase_message}：{}", event.status)
    };
    let mut record = tracker.event(level, &event.phase, &event.phase, &event.status, message);
    record.timestamp = event.timestamp.clone();
    record.file_index = event.source_input_index;
    record.file_name = event
        .workbook_name
        .as_deref()
        .map(|name| pytools_core::run_logging::portable_file_name(Path::new(name)));
    record.sheet_name = event.worksheet_name.clone();
    record.config_row = event.config_row;
    record.engine = event.actual_engine.clone();
    record.elapsed_ms = event.stage_elapsed_ms;
    if event.status == "failed" {
        record.error_code = Some(
            match event.phase.as_str() {
                "source_read" => "EXCEL_READ_FAILED",
                "summary_write" | "target_write" => "OUTPUT_WRITE_FAILED",
                _ => "INTERNAL_OPERATION_FAILED",
            }
            .to_owned(),
        );
    }
    record
}

fn emit_wide_summary_event(
    log_view: &str,
    events: &mut Vec<WideSummaryRunEvent>,
    event: WideSummaryRunEvent,
) {
    match log_view {
        "developer" => match serde_json::to_string(&event) {
            Ok(encoded) => eprintln!("{encoded}"),
            Err(_) => {
                eprintln!("{{\"phase\":\"progress_event\",\"status\":\"serialization_failed\"}}")
            }
        },
        "user" => {
            if event.status == "failed"
                || matches!(
                    event.phase.as_str(),
                    "source_match" | "summary_write" | "target_write" | "source_discovery"
                )
            {
                let level = match event.status.as_str() {
                    "failed" => pytools_core::run_logging::Level::Error,
                    "skipped" => pytools_core::run_logging::Level::Warn,
                    _ => pytools_core::run_logging::Level::Info,
                };
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(level, &event.user_message)
                );
            }
        }
        _ => {}
    }
    events.push(event);
}

fn run_summary_used_range_cli(arguments: Vec<String>, interactive: bool) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let output_dir = options
        .values
        .get("output-dir")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("output"));
    let log_view = options
        .values
        .get("log-view")
        .map(String::as_str)
        .unwrap_or("user");
    anyhow::ensure!(
        matches!(log_view, "user" | "developer" | "off"),
        "--log-view must be user, developer, or off"
    );
    let log_dir = options
        .values
        .get("log-dir")
        .map(PathBuf::from)
        .unwrap_or_else(pytools_core::run_logging::default_log_dir);
    if options.positional.is_empty() {
        bail!("provide at least one workbook path");
    }

    let run_started = Instant::now();
    let started_at = Local::now().to_rfc3339();
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_nanos())
        .unwrap_or_default();
    let run_id = format!("pytools-pt-6-1-{}-{nonce}", std::process::id());
    use pytools_core::run_logging::{EventWriter, FileStatus, Level, RunTracker, SheetStatus};
    let mut tracker = RunTracker::with_id(
        "PT-6-1",
        "按使用区域快速汇总",
        &options.positional,
        run_id.clone(),
    );
    let mut diagnostic = if log_view == "off" {
        None
    } else {
        match EventWriter::create(&log_dir.join("run-events"), "PT-6-1", &run_id) {
            Ok(writer) => Some(writer),
            Err(error) => {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("无法创建诊断日志：{error}；业务处理继续")
                    )
                );
                None
            }
        }
    };
    let mut diagnostic_failed = false;
    let mut events = Vec::<SummaryRunEvent>::new();
    let mut next_sheet_index = std::collections::HashMap::<usize, usize>::new();
    let result = summarize_used_ranges_with_progress(&options.positional, &output_dir, |event| {
        let input_index = event.progress.current_input_index;
        if let Some(index) = input_index {
            match (
                event.phase.as_str(),
                event.status.as_str(),
                tracker.state(index),
            ) {
                ("input_filter", "skipped", Some(FileStatus::Pending)) => {
                    let _ = tracker.transition(index, FileStatus::Skipped);
                }
                ("extent_scan", "started", Some(FileStatus::Pending)) => {
                    let _ = tracker.transition(index, FileStatus::InProgress);
                }
                ("extent_scan" | "value_read", "failed", Some(FileStatus::InProgress)) => {
                    let _ = tracker.transition(index, FileStatus::Failed);
                }
                ("row_write", "succeeded", Some(FileStatus::InProgress)) => {
                    let _ = tracker.transition(index, FileStatus::Succeeded);
                }
                _ => {}
            }
        }
        if event.phase == "sheet_process" {
            if let (Some(index), Some(sheet_name)) =
                (input_index, event.progress.current_sheet.as_deref())
            {
                match event.status.as_str() {
                    "started" => {
                        let sheet_index = next_sheet_index.entry(index).or_default();
                        *sheet_index += 1;
                        if tracker
                            .discover_sheet(index, Some(*sheet_index), sheet_name)
                            .is_ok()
                        {
                            let _ = tracker.transition_sheet(
                                index,
                                sheet_name,
                                SheetStatus::InProgress,
                                &event.phase,
                                None,
                                None,
                            );
                        }
                    }
                    "succeeded" | "skipped" | "failed" => {
                        if tracker.sheet_state(index, sheet_name) == Some(SheetStatus::InProgress) {
                            let next = match event.status.as_str() {
                                "succeeded" => SheetStatus::Succeeded,
                                "skipped" => SheetStatus::Skipped,
                                _ => SheetStatus::Failed,
                            };
                            let error_code = event
                                .error_type
                                .as_deref()
                                .map(|kind| summary_event_error_code(Some(kind), &event.phase));
                            let _ = tracker.transition_sheet(
                                index,
                                sheet_name,
                                next,
                                &event.phase,
                                None,
                                error_code.as_deref(),
                            );
                        }
                    }
                    _ => {}
                }
            }
        }
        let level = match event.status.as_str() {
            "failed" => Level::Error,
            "skipped" => Level::Warn,
            "started" | "running" => Level::Info,
            "succeeded" if event.phase == "row_write" || event.phase == "run" => Level::Info,
            _ => Level::Debug,
        };
        let mut record = tracker.event(
            level,
            &event.phase,
            &event.phase,
            &event.status,
            &event.user_message,
        );
        record.timestamp = event.timestamp.clone();
        record.file_index = if event.file_name.is_some() {
            input_index
        } else {
            None
        };
        record.file_name = event.file_name.clone();
        record.sheet_name = event.progress.current_sheet.clone();
        record.engine = event.actual_engine.clone();
        record.elapsed_ms = event.stage_elapsed_ms;
        record.error_code = event.error_type.as_ref().map(|_| {
            match event.phase.as_str() {
                "extent_scan" | "value_read" => "EXCEL_READ_FAILED",
                "prepare_output" | "output_save" => "OUTPUT_WRITE_FAILED",
                _ => "INTERNAL_OPERATION_FAILED",
            }
            .to_owned()
        });
        if let Some(writer) = diagnostic.as_mut() {
            if let Err(error) = writer.write(&record) {
                if !diagnostic_failed {
                    eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            Level::Warn,
                            &format!("诊断日志写入失败：{error}；业务处理继续")
                        )
                    );
                    diagnostic_failed = true;
                }
            }
        }
        if log_view == "developer"
            || (log_view == "user"
                && (event.status == "failed"
                    || (event.status == "skipped" && event.phase != "sheet_process")
                    || matches!(
                        (event.phase.as_str(), event.status.as_str()),
                        ("run", _) | ("row_write", "succeeded") | ("output_save", "started")
                    )))
        {
            emit_summary_batch_event_view(log_view, &event);
        }
        events.push(event);
    });

    if matches!(&result, Ok(summary) if !summary.saved) {
        tracker.skip_unfinished();
    } else if result.is_err() {
        if let Some(event) = events.last() {
            if let (Some(index), Some(sheet_name)) = (
                event.progress.current_input_index,
                event.progress.current_sheet.as_deref(),
            ) {
                if tracker.sheet_state(index, sheet_name) == Some(SheetStatus::InProgress) {
                    let code = summary_event_error_code(event.error_type.as_deref(), &event.phase);
                    let _ = tracker.transition_sheet(
                        index,
                        sheet_name,
                        SheetStatus::Failed,
                        &event.phase,
                        None,
                        Some(&code),
                    );
                }
            }
            if event.file_name.is_some() {
                if let Some(index) = event.progress.current_input_index {
                    if tracker.state(index) == Some(FileStatus::InProgress) {
                        let _ = tracker.transition(index, FileStatus::Failed);
                    }
                }
            }
        }
    }
    let diagnostic_path = diagnostic_path_for_log_view(diagnostic.as_ref(), log_view);
    let outputs = result
        .as_ref()
        .ok()
        .and_then(|summary| summary.path.as_ref())
        .map(|path| vec![path_for_log_view(path, log_view)])
        .unwrap_or_default();
    let mut unified = tracker.finish(
        result.is_err(),
        false,
        matches!(&result, Ok(summary) if !summary.saved),
        outputs,
        diagnostic_path.clone(),
    );
    if let Some(last) = events.last() {
        let progress = &last.progress;
        let tracked = tracker.sheet_stats();
        let sheet_stats = pytools_core::run_logging::SheetStats {
            discovered: tracked.discovered.max(progress.sheets_discovered),
            attempted: tracked.attempted.max(progress.sheets_attempted),
            succeeded: tracked.succeeded.max(progress.sheets_succeeded),
            failed: tracked.failed.max(progress.sheets_failed),
            skipped: tracked.skipped.max(progress.sheets_skipped),
        };
        if sheet_stats.discovered > 0 || sheet_stats.attempted > 0 {
            unified.sheet_stats = Some(sheet_stats);
        }
        unified.record_stats = Some(pytools_core::run_logging::RecordStats {
            rows_processed: Some(progress.processed_rows),
            rows_written: result
                .as_ref()
                .ok()
                .filter(|summary| summary.saved)
                .map(|summary| summary.rows),
        });
        if result.is_err() {
            let failed_event = events.iter().rev().find(|event| event.status == "failed");
            unified.error_summary = Some(pytools_core::run_logging::ErrorSummary {
                stage: failed_event
                    .map(|event| event.phase.clone())
                    .unwrap_or_else(|| last.phase.clone()),
                code: summary_event_error_code(
                    failed_event.and_then(|event| event.error_type.as_deref()),
                    failed_event
                        .map(|event| event.phase.as_str())
                        .unwrap_or(last.phase.as_str()),
                ),
                message: failed_event
                    .map(|event| event.user_message.clone())
                    .unwrap_or_else(|| {
                        "按使用区域汇总未完成；请检查输入工作簿和输出目录".to_owned()
                    }),
                file_name: failed_event
                    .and_then(|event| event.file_name.clone())
                    .or_else(|| last.file_name.clone()),
                sheet_name: failed_event
                    .and_then(|event| event.progress.current_sheet.clone())
                    .or_else(|| last.progress.current_sheet.clone()),
            });
        }
    }
    unified.output_stats = Some(pytools_core::run_logging::OutputStats {
        workbooks_saved: Some(unified.output_paths.len()),
        sheets_saved: result
            .as_ref()
            .ok()
            .filter(|summary| summary.saved)
            .map(|_| 1),
        saved: Some(result.as_ref().is_ok_and(|summary| summary.saved)),
    });
    if let Some(execution) = result
        .as_ref()
        .ok()
        .and_then(|summary| summary.execution.as_ref())
    {
        unified.performance_stats = Some(pytools_core::run_logging::PerformanceStats {
            read_elapsed_ms: Some(execution.read_elapsed_ms),
            write_elapsed_ms: Some(execution.write_elapsed_ms),
        });
    }
    if let Err(error) = result.as_ref() {
        let failed_event = events.iter().rev().find(|event| event.status == "failed");
        let stage = failed_event
            .map(|event| event.phase.as_str())
            .unwrap_or("summary");
        let message = failed_event
            .map(|event| event.user_message.as_str())
            .unwrap_or("按使用区域汇总未完成");
        if let Some(writer) = diagnostic.as_mut() {
            let mut failure =
                tracker.event(Level::Error, "operation_failed", stage, "failed", message);
            failure.error_code = Some(summary_event_error_code(
                failed_event.and_then(|event| event.error_type.as_deref()),
                stage,
            ));
            let failure_context = failed_event.or_else(|| {
                events.iter().rev().find(|event| {
                    event.file_name.is_some()
                        && matches!(
                            event.phase.as_str(),
                            "extent_scan" | "value_read" | "sheet_process" | "row_write"
                        )
                })
            });
            if let Some(context) = failure_context {
                failure.file_index = context.progress.current_input_index;
                failure.file_name = context.file_name.clone();
                failure.sheet_name = context.progress.current_sheet.clone();
                if log_view == "developer" {
                    failure.file_path = context
                        .progress
                        .current_input_index
                        .and_then(|index| options.positional.get(index.saturating_sub(1)))
                        .map(|path| path.to_string_lossy().into_owned());
                }
            }
            if log_view == "developer" {
                failure.details = Some(serde_json::json!({"error_chain": format!("{error:#}")}));
            }
            if let Err(write_error) = writer.write(&failure) {
                if !diagnostic_failed {
                    eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            Level::Warn,
                            &format!("诊断日志写入失败：{write_error}；业务结果仍保留")
                        )
                    );
                    diagnostic_failed = true;
                }
            }
        }
    }
    if let Some(writer) = diagnostic.as_mut() {
        let event = tracker.event(
            if result.is_err() {
                Level::Error
            } else {
                Level::Info
            },
            "run_finished",
            "run",
            &format!("{:?}", unified.overall_status),
            format!("{}结束", tracker.feature_name),
        );
        if let Err(error) = writer.write(&event) {
            if !diagnostic_failed {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("诊断日志写入失败：{error}；业务结果仍保留")
                    )
                );
            }
        }
    }

    if result.is_err()
        && log_view != "off"
        && !events.last().is_some_and(|event| event.status == "failed")
    {
        let prior = events.last();
        events.push(SummaryRunEvent {
            sequence: events.len() + 1,
            timestamp: Local::now().to_rfc3339(),
            phase: "summary".to_owned(),
            status: "failed".to_owned(),
            user_message: "按使用区域汇总失败".to_owned(),
            progress: prior
                .map(|event| event.progress.clone())
                .unwrap_or_else(|| SummaryProgress {
                    total_inputs: options.positional.len(),
                    ..SummaryProgress::default()
                }),
            file_name: prior.and_then(|event| event.file_name.clone()),
            stage_elapsed_ms: None,
            total_elapsed_ms: run_started.elapsed().as_millis(),
            components: vec!["pytools-core".to_owned()],
            actual_engine: None,
            error_type: Some("summary_error".to_owned()),
        });
        let event = events.last().expect("failure event was just appended");
        if log_view == "developer" {
            eprintln!("{}", serde_json::to_string(event)?);
        } else {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(Level::Error, &event.user_message)
            );
        }
    }

    if interactive && log_view == "user" {
        emit_unified_terminal_summary_with_outputs(&unified, &[]);
    }

    let run_log_path = if log_view == "off" {
        None
    } else {
        let path = log_dir.join(format!("PT-6-1_{run_id}.json"));
        let status = match &result {
            Ok(summary) if summary.saved => "succeeded",
            Ok(_) => "skipped",
            Err(_) => "failed",
        };
        let failure = result.as_ref().err().map(|_| {
            let event = events.last();
            serde_json::json!({
                "stage": event.map(|event| event.phase.as_str()).unwrap_or("summary"),
                "error_type": event.and_then(|event| event.error_type.as_deref()).unwrap_or("summary_error")
            })
        });
        let report = serde_json::json!({
            "schema_version": 1,
            "product": "pytools",
            "feature_id": "PT-6-1",
            "function": "summary-used-range",
            "run_id": &run_id,
            "started_at": &started_at,
            "status": status,
            "overall_status": unified.overall_status,
            "feature_name": unified.feature_name,
            "ended_at": unified.ended_at,
            "total_elapsed_ms": unified.total_elapsed_ms,
            "input_stats": unified.input_stats,
            "sheet_stats": unified.sheet_stats,
            "record_stats": unified.record_stats,
            "output_stats": unified.output_stats,
            "performance_stats": unified.performance_stats,
            "error_summary": unified.error_summary,
            "inputs": unified.inputs,
            "sheets": unified.sheets,
            "output_paths": unified.output_paths,
            "diagnostic_log_path": unified.diagnostic_log_path,
            "total_inputs": options.positional.len(),
            "completed_files": events.last().map(|event| event.progress.completed_files).unwrap_or(0),
            "failed_files": events.last().map(|event| event.progress.failed_files).unwrap_or(0),
            "summary": result.as_ref().ok(),
            "failure": failure,
            "elapsed_ms": run_started.elapsed().as_millis(),
            // 与开发者 stderr 流使用同一份注入 level/event_name 后的事件，
            // 保证两个通道逐字段一致（历史记录即终端所见）。
            "events": events.iter().map(summary_event_json).collect::<Vec<_>>(),
        });
        match write_json_report_create_new(&path, &report) {
            Ok(()) => Some(path),
            Err(_error) => {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        "无法保存 PT-6-1 运行汇总；业务结果仍保留",
                    )
                );
                None
            }
        }
    };

    match result {
        Ok(result) => {
            let mut output = serde_json::to_value(result)?;
            if log_view != "off" {
                output["run_id"] = serde_json::Value::String(run_id);
                if let Some(path) = run_log_path {
                    output["run_log_path"] =
                        serde_json::Value::String(display_path_for_user(&path).to_string());
                    if log_view == "user" {
                        eprintln!(
                            "{}",
                            pytools_core::run_logging::terminal_line(
                                Level::Info,
                                &format!("运行记录已保存：{}", path_for_log_view(&path, log_view))
                            )
                        );
                    }
                }
            }
            if interactive && log_view == "user" {
                Ok(())
            } else if interactive && log_view == "off" {
                if let Some(path) = output["path"].as_str() {
                    println!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            Level::Info,
                            &format!("汇总文件：{path}")
                        )
                    );
                }
                Ok(())
            } else {
                println!("{}", serde_json::to_string_pretty(&output)?);
                Ok(())
            }
        }
        Err(error) => {
            if log_view == "user" {
                if let Some(path) = run_log_path {
                    eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            Level::Info,
                            &format!("运行记录已保存：{}", path_for_log_view(&path, log_view))
                        )
                    );
                }
            }
            Err(error)
        }
    }
}

fn run_summary_named_ranges_cli(arguments: Vec<String>, interactive: bool) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let template = required_path(&options, "template")?;
    let output_dir = options
        .values
        .get("output-dir")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("output"));
    let log_view = options
        .values
        .get("log-view")
        .map(String::as_str)
        .unwrap_or("user");
    anyhow::ensure!(
        matches!(log_view, "user" | "developer" | "off"),
        "--log-view must be user, developer, or off"
    );
    let log_dir = options
        .values
        .get("log-dir")
        .map(PathBuf::from)
        .unwrap_or_else(pytools_core::run_logging::default_log_dir);
    anyhow::ensure!(
        !options.positional.is_empty(),
        "provide at least one source workbook path"
    );

    let run_started = Instant::now();
    let started_at = Local::now().to_rfc3339();
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_nanos())
        .unwrap_or_default();
    let run_id = format!("pytools-pt-6-2-named-{}-{nonce}", std::process::id());
    use pytools_core::run_logging::{EventWriter, FileStatus, Level, RunTracker};
    let mut tracker = RunTracker::with_id(
        "PT-6-2",
        "按模板命名区域汇总",
        &options.positional,
        run_id.clone(),
    );
    let mut diagnostic = if log_view == "off" {
        None
    } else {
        match EventWriter::create(&log_dir.join("run-events"), "PT-6-2", &run_id) {
            Ok(writer) => Some(writer),
            Err(error) => {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("无法创建诊断日志：{error}")
                    )
                );
                None
            }
        }
    };
    let mut diagnostic_failed = false;
    let mut events = Vec::<SummaryRunEvent>::new();
    let result = summarize_by_named_ranges_with_progress(
        &template,
        &options.positional,
        &output_dir,
        |event| {
            let index = event.progress.current_input_index;
            if let Some(index) = index {
                match (
                    event.phase.as_str(),
                    event.status.as_str(),
                    tracker.state(index),
                ) {
                    ("input_filter", "skipped", Some(FileStatus::Pending)) => {
                        let _ = tracker.transition(index, FileStatus::Skipped);
                    }
                    ("source_read", "started", Some(FileStatus::Pending)) => {
                        let _ = tracker.transition(index, FileStatus::InProgress);
                    }
                    ("source_structure", "skipped", Some(FileStatus::InProgress)) => {
                        let _ = tracker.transition(index, FileStatus::Skipped);
                    }
                    (_, "failed", Some(FileStatus::InProgress)) if event.file_name.is_some() => {
                        let _ = tracker.transition(index, FileStatus::Failed);
                    }
                    ("row_extract", "succeeded", Some(FileStatus::InProgress)) => {
                        let _ = tracker.transition(index, FileStatus::Succeeded);
                    }
                    _ => {}
                }
            }
            let detailed_phase = matches!(
                event.phase.as_str(),
                "template_read"
                    | "source_read"
                    | "sheet_validation"
                    | "sheet_routing"
                    | "summary_region_read"
                    | "cell_field_read"
                    | "template_validation"
            );
            let level = match event.status.as_str() {
                "failed" => Level::Error,
                "skipped" | "warning" => Level::Warn,
                "started" => Level::Info,
                "succeeded" if detailed_phase => Level::Info,
                _ => Level::Debug,
            };
            let safe_message = &event.user_message;
            let mut record = tracker.event(
                level,
                &event.phase,
                &event.phase,
                &event.status,
                safe_message,
            );
            record.timestamp = event.timestamp.clone();
            record.file_index = if event.file_name.is_some() {
                index
            } else {
                None
            };
            record.file_name = event.file_name.clone();
            record.sheet_name = event.progress.current_sheet.clone();
            record.engine = event.actual_engine.clone();
            record.elapsed_ms = event.stage_elapsed_ms;
            if event.status == "failed" {
                record.error_code = Some(summary_event_error_code(
                    event.error_type.as_deref(),
                    &event.phase,
                ));
            }
            if let Some(writer) = diagnostic.as_mut() {
                if let Err(error) = writer.write(&record) {
                    if !diagnostic_failed {
                        eprintln!(
                            "{}",
                            pytools_core::run_logging::terminal_line(
                                Level::Warn,
                                &format!("诊断日志写入失败：{error}")
                            )
                        );
                        diagnostic_failed = true;
                    }
                }
            }
            if log_view == "developer"
                || (log_view == "user"
                    && (matches!(event.status.as_str(), "failed" | "skipped" | "warning")
                        || matches!(
                            event.phase.as_str(),
                            "run"
                                | "row_extract"
                                | "output_save"
                                | "source_read"
                                | "sheet_validation"
                        )))
            {
                emit_summary_batch_event_view(log_view, &event);
            }
            events.push(event);
        },
    );
    if result.is_err() {
        if let Some(last) = events.last() {
            if last.file_name.is_some() {
                if let Some(index) = last.progress.current_input_index {
                    if tracker.state(index) == Some(FileStatus::InProgress) {
                        let _ = tracker.transition(index, FileStatus::Failed);
                    }
                }
            }
        }
    }
    let outputs = result
        .as_ref()
        .ok()
        .and_then(|summary| summary.path.as_ref())
        .map(|path| vec![path_for_log_view(path, log_view)])
        .unwrap_or_default();
    if let Some(summary) = result.as_ref().ok() {
        tracker.set_business_metric("成功处理源工作簿数", serde_json::json!(summary.files));
        tracker.set_business_metric("提取行数", serde_json::json!(summary.rows));
        tracker.set_business_metric(
            "跳过源文件或工作表项",
            serde_json::json!(summary.skipped_sources.len()),
        );
    }
    for event in events.iter().filter(|event| event.status == "warning") {
        tracker.record_business_issue(
            pytools_core::run_logging::BusinessIssue::new("业务提醒", &event.user_message).at(
                event.file_name.as_deref(),
                event.progress.current_sheet.as_deref(),
                None,
            ),
        );
    }
    let mut unified = tracker.finish(
        result.is_err(),
        false,
        matches!(&result, Ok(summary) if !summary.saved),
        outputs,
        diagnostic_path_for_log_view(diagnostic.as_ref(), log_view),
    );
    unified.output_stats = Some(pytools_core::run_logging::OutputStats {
        workbooks_saved: Some(unified.output_paths.len()),
        sheets_saved: None,
        saved: Some(result.as_ref().is_ok_and(|summary| summary.saved)),
    });
    if result.is_err() {
        if let Some(event) = events.iter().rev().find(|event| event.status == "failed") {
            unified.error_summary = Some(pytools_core::run_logging::ErrorSummary {
                stage: event.phase.clone(),
                code: summary_event_error_code(event.error_type.as_deref(), &event.phase),
                message: event.user_message.clone(),
                file_name: event.file_name.clone(),
                sheet_name: event.progress.current_sheet.clone(),
            });
        }
        if let (Err(error), Some(writer)) = (&result, diagnostic.as_mut()) {
            let event = events.iter().rev().find(|event| event.status == "failed");
            let stage = event.map(|event| event.phase.as_str()).unwrap_or("summary");
            let mut failure = tracker.event(
                Level::Error,
                "operation_failed",
                stage,
                "failed",
                "按模板命名区域汇总失败",
            );
            failure.error_code = Some(summary_event_error_code(
                event.and_then(|event| event.error_type.as_deref()),
                stage,
            ));
            if log_view == "developer" {
                failure.details = Some(serde_json::json!({"error_chain": format!("{error:#}")}));
            }
            if let Err(write_error) = writer.write(&failure) {
                if !diagnostic_failed {
                    eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            Level::Warn,
                            &format!("诊断日志写入失败：{write_error}")
                        )
                    );
                    diagnostic_failed = true;
                }
            }
        }
    }
    if let Some(writer) = diagnostic.as_mut() {
        if let Err(error) = writer.write(&tracker.event(
            if result.is_err() {
                Level::Error
            } else {
                Level::Info
            },
            "run_finished",
            "run",
            &format!("{:?}", unified.overall_status),
            "命名区域汇总结束",
        )) {
            if !diagnostic_failed {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("诊断日志写入失败：{error}")
                    )
                );
            }
        }
    }

    if result.is_err()
        && log_view != "off"
        && !events.last().is_some_and(|event| event.status == "failed")
    {
        // 模板/配置级失败（未进入逐文件处理）：没有 file_name 是正常语义，
        // 但必须把根因（缺表头区域、缺汇总区域等）带进失败事件。
        let had_file_failure = events.iter().any(|event| event.status == "failed");
        let error_chain = result
            .as_ref()
            .err()
            .map(|error| format!("{error:#}"))
            .unwrap_or_default();
        let (classified_type, root_hint) = classify_named_range_failure(&error_chain);
        let prior = events.last();
        events.push(SummaryRunEvent {
            sequence: events.len() + 1,
            timestamp: Local::now().to_rfc3339(),
            phase: "summary".to_owned(),
            status: "failed".to_owned(),
            user_message: format!("按模板命名区域汇总失败：{root_hint}"),
            progress: prior
                .map(|event| event.progress.clone())
                .unwrap_or_else(|| SummaryProgress {
                    total_inputs: options.positional.len(),
                    ..SummaryProgress::default()
                }),
            file_name: prior.and_then(|event| event.file_name.clone()),
            stage_elapsed_ms: None,
            total_elapsed_ms: run_started.elapsed().as_millis(),
            components: vec!["pytools-core".to_owned()],
            actual_engine: None,
            error_type: Some(classified_type.to_owned()),
        });
        // 统一摘要此前按“最后一个失败事件”填充；模板级失败时它可能是空的，
        // 用分类结果覆盖，保证运行汇总直接给出根因。若批内已有逐文件失败
        // 事件（带文件归属），保留原有摘要不覆盖。
        if !had_file_failure {
            unified.error_summary = Some(pytools_core::run_logging::ErrorSummary {
                stage: "summary".to_owned(),
                code: summary_event_error_code(Some(classified_type), "summary"),
                message: format!("按模板命名区域汇总失败：{root_hint}"),
                file_name: None,
                sheet_name: None,
            });
        }
        let event = events.last().expect("failure event was just appended");
        if log_view == "developer" {
            eprintln!("{}", serde_json::to_string(event)?);
        } else {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(Level::Error, &event.user_message)
            );
        }
    }

    if interactive && log_view == "user" {
        emit_unified_terminal_summary_with_outputs(&unified, &[]);
    }

    let run_log_path = if log_view == "off" {
        None
    } else {
        let path = log_dir.join(format!("PT-6-2_{run_id}.json"));
        let status = match &result {
            Ok(summary) if summary.saved => "succeeded",
            Ok(_) => "skipped",
            Err(_) => "failed",
        };
        let failure = result.as_ref().err().map(|error| {
            let event = events.last();
            let (error_type, root_hint) = classify_named_range_failure(&format!("{error:#}"));
            let failure_type = event
                .and_then(|event| event.error_type.as_deref())
                .filter(|event_type| *event_type != "template_named_region_invalid")
                .unwrap_or(error_type);
            serde_json::json!({
                "stage": event.map(|event| event.phase.as_str()).unwrap_or("summary"),
                "error_type": failure_type,
                "message": root_hint,
            })
        });
        let report = serde_json::json!({
            "schema_version": 1,
            "product": "pytools",
            "feature_id": "PT-6-2",
            "function": "summary-named-ranges",
            "run_id": &run_id,
            "started_at": &started_at,
            "status": status,
            "overall_status": unified.overall_status,
            "feature_name": unified.feature_name,
            "ended_at": unified.ended_at,
            "input_stats": unified.input_stats,
            "business_metrics": unified.business_metrics,
            "business_issues": unified.business_issues,
            "inputs": unified.inputs,
            "output_stats": unified.output_stats,
            "output_paths": unified.output_paths,
            "diagnostic_log_path": unified.diagnostic_log_path,
            "total_inputs": options.positional.len(),
            "completed_files": events.last().map(|event| event.progress.completed_files).unwrap_or(0),
            "failed_files": events.last().map(|event| event.progress.failed_files).unwrap_or(0),
            "summary": result.as_ref().ok(),
            "failure": failure,
            "elapsed_ms": run_started.elapsed().as_millis(),
            "events": events,
        });
        match write_json_report_create_new(&path, &report) {
            Ok(()) => Some(path),
            Err(error) => {
                eprintln!("运行诊断记录保存失败；汇总结果仍保留：{error:#}");
                None
            }
        }
    };

    match result {
        Ok(result) => {
            let mut output = serde_json::to_value(result)?;
            if log_view != "off" {
                output["run_id"] = serde_json::Value::String(run_id);
                if let Some(path) = run_log_path {
                    output["run_log_path"] =
                        serde_json::Value::String(display_path_for_user(&path).to_string());
                }
            }
            if interactive && log_view == "user" {
                Ok(())
            } else if interactive && log_view == "off" {
                if let Some(path) = output["path"].as_str() {
                    println!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            Level::Info,
                            &format!("汇总文件：{path}")
                        )
                    );
                }
                Ok(())
            } else {
                println!("{}", serde_json::to_string_pretty(&output)?);
                Ok(())
            }
        }
        Err(error) => {
            if interactive && log_view == "user" {
                Err(anyhow::anyhow!(
                    "按模板命名区域汇总失败；请检查上方提示的失败阶段、模板区域与源文件"
                ))
            } else {
                Err(error)
            }
        }
    }
}

fn run_named_validation_extract_cli(arguments: Vec<String>, interactive: bool) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let template = required_path(&options, "template")?;
    anyhow::ensure!(
        !options.positional.is_empty(),
        "provide at least one source workbook path"
    );
    let output_dir = options
        .values
        .get("output-dir")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("output"));
    let log_view = options
        .values
        .get("log-view")
        .map(String::as_str)
        .unwrap_or("user");
    anyhow::ensure!(
        matches!(log_view, "user" | "developer" | "off"),
        "--log-view must be user, developer, or off"
    );
    let log_dir = options
        .values
        .get("log-dir")
        .map(PathBuf::from)
        .unwrap_or_else(pytools_core::run_logging::default_log_dir);

    use pytools_core::run_logging::{EventWriter, FileStatus, Level, RunTracker};
    let run_started = Instant::now();
    let started_at = Local::now().to_rfc3339();
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_nanos())
        .unwrap_or_default();
    let run_id = format!("pytools-pt-6-3-validation-{}-{nonce}", std::process::id());
    let mut tracker = RunTracker::with_id(
        "PT-6-3",
        "按模板校验区域提取校验",
        &options.positional,
        run_id.clone(),
    );
    let mut diagnostic = if log_view == "off" {
        None
    } else {
        match EventWriter::create(&log_dir.join("run-events"), "PT-6-3", &run_id) {
            Ok(writer) => Some(writer),
            Err(error) => {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("无法创建诊断日志：{error}；继续处理")
                    )
                );
                None
            }
        }
    };
    let mut diagnostic_failed = false;
    let mut events = Vec::<SummaryRunEvent>::new();
    let result = pytools_core::summary_ops::extract_named_validation_regions_with_progress(
        &template,
        &options.positional,
        &output_dir,
        |event| {
            if let Some(index) = event.progress.current_input_index {
                match (
                    event.phase.as_str(),
                    event.status.as_str(),
                    tracker.state(index),
                ) {
                    ("input_filter", "skipped", Some(FileStatus::Pending)) => {
                        let _ = tracker.transition(index, FileStatus::Skipped);
                    }
                    ("source_read", "started", Some(FileStatus::Pending)) => {
                        let _ = tracker.transition(index, FileStatus::InProgress);
                    }
                    ("source_validation", "skipped", Some(FileStatus::InProgress)) => {
                        let _ = tracker.transition(index, FileStatus::Skipped);
                    }
                    ("validation_calculation", "succeeded", Some(FileStatus::InProgress)) => {
                        let _ = tracker.transition(index, FileStatus::Succeeded);
                    }
                    (_, "failed", Some(FileStatus::InProgress)) if event.file_name.is_some() => {
                        let _ = tracker.transition(index, FileStatus::Failed);
                    }
                    _ => {}
                }
            }
            let level = match event.status.as_str() {
                "failed" => Level::Error,
                "skipped" | "warning" => Level::Warn,
                "started" | "succeeded" => Level::Info,
                _ => Level::Debug,
            };
            let mut record = tracker.event(
                level,
                &event.phase,
                &event.phase,
                &event.status,
                &event.user_message,
            );
            record.timestamp = event.timestamp.clone();
            record.file_index = if event.file_name.is_some() {
                event.progress.current_input_index
            } else {
                None
            };
            record.file_name = event.file_name.clone();
            record.sheet_name = event.progress.current_sheet.clone();
            record.engine = event.actual_engine.clone();
            record.elapsed_ms = event.stage_elapsed_ms;
            if event.status == "failed" {
                record.error_code = Some(
                    event
                        .error_type
                        .clone()
                        .unwrap_or_else(|| "named_validation_error".to_owned()),
                );
            }
            if let Some(writer) = diagnostic.as_mut() {
                if let Err(error) = writer.write(&record) {
                    if !diagnostic_failed {
                        eprintln!(
                            "{}",
                            pytools_core::run_logging::terminal_line(
                                Level::Warn,
                                &format!("诊断日志写入失败：{error}")
                            )
                        );
                        diagnostic_failed = true;
                    }
                }
            }
            if log_view == "developer"
                || (log_view == "user"
                    && (matches!(event.status.as_str(), "failed" | "skipped" | "warning")
                        || matches!(
                            event.phase.as_str(),
                            "run" | "validation_calculation" | "output_save"
                        )))
            {
                emit_summary_batch_event_view(log_view, &event);
            }
            events.push(event);
        },
    );
    if result.is_err() && !events.last().is_some_and(|event| event.status == "failed") {
        let message = result
            .as_ref()
            .err()
            .map(|error| error.root_cause().to_string())
            .unwrap_or_default();
        let event = tracker.event(
            Level::Error,
            "operation_failed",
            "validation_extract",
            "failed",
            &format!("按模板校验区域提取校验失败：{message}"),
        );
        if let Some(writer) = diagnostic.as_mut() {
            if let Err(error) = writer.write(&event) {
                if !diagnostic_failed {
                    eprintln!("诊断日志写入失败：{error}");
                }
            }
        }
        events.push(SummaryRunEvent {
            sequence: events.len() + 1,
            timestamp: Local::now().to_rfc3339(),
            phase: "operation_failed".to_owned(),
            status: "failed".to_owned(),
            user_message: format!("按模板校验区域提取校验失败：{message}"),
            progress: events
                .last()
                .map(|event| event.progress.clone())
                .unwrap_or_default(),
            file_name: events.last().and_then(|event| event.file_name.clone()),
            stage_elapsed_ms: None,
            total_elapsed_ms: run_started.elapsed().as_millis(),
            components: vec!["pytools-core".to_owned()],
            actual_engine: None,
            error_type: Some("named_validation_extract_error".to_owned()),
        });
        if log_view == "user" {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Error,
                    &format!("按模板校验区域提取校验失败：{message}")
                )
            );
        }
    }
    if let Some(result) = result.as_ref().ok() {
        tracker.set_business_metric("成功读取工作簿数", serde_json::json!(result.files));
        tracker.set_business_metric("提取校验结果数", serde_json::json!(result.rows));
        tracker.set_business_metric(
            "跳过源文件或工作表项",
            serde_json::json!(result.skipped_sources.len()),
        );
        if result.rows > 0 {
            tracker.record_business_issue(
                pytools_core::run_logging::BusinessIssue::new(
                    "校验异常",
                    "提取到需要人工查看的校验结果",
                )
                .count(result.rows),
            );
        }
    }
    for event in events.iter().filter(|event| event.status == "warning") {
        tracker.record_business_issue(
            pytools_core::run_logging::BusinessIssue::new("业务提醒", &event.user_message).at(
                event.file_name.as_deref(),
                event.progress.current_sheet.as_deref(),
                None,
            ),
        );
    }
    let outputs = result
        .as_ref()
        .ok()
        .and_then(|result| result.path.as_ref())
        .map(|path| vec![path_for_log_view(path, log_view)])
        .unwrap_or_default();
    let mut unified = tracker.finish(
        result.is_err(),
        false,
        matches!(&result, Ok(result) if !result.saved),
        outputs,
        diagnostic_path_for_log_view(diagnostic.as_ref(), log_view),
    );
    unified.output_stats = Some(pytools_core::run_logging::OutputStats {
        workbooks_saved: Some(unified.output_paths.len()),
        sheets_saved: Some(1),
        saved: Some(result.as_ref().is_ok_and(|result| result.saved)),
    });
    if result.is_err() {
        if let Some(event) = events.iter().rev().find(|event| event.status == "failed") {
            unified.error_summary = Some(pytools_core::run_logging::ErrorSummary {
                stage: event.phase.clone(),
                code: event
                    .error_type
                    .clone()
                    .unwrap_or_else(|| "named_validation_extract_error".to_owned()),
                message: event.user_message.clone(),
                file_name: event.file_name.clone(),
                sheet_name: event.progress.current_sheet.clone(),
            });
        }
    }
    if let Some(writer) = diagnostic.as_mut() {
        let finished = tracker.event(
            if result.is_err() {
                Level::Error
            } else {
                Level::Info
            },
            "run_finished",
            "run",
            &format!("{:?}", unified.overall_status),
            "按模板校验区域提取校验结束",
        );
        if let Err(error) = writer.write(&finished) {
            if !diagnostic_failed {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("诊断日志写入失败：{error}")
                    )
                );
            }
        }
    }
    let report_path = if log_view == "off" {
        None
    } else {
        let path = log_dir.join(format!("PT-6-3_{run_id}.json"));
        let report = serde_json::json!({
            "schema_version": 1,
            "product": "pytools",
            "feature_id": "PT-6-3",
            "function": "named-validation-extract",
            "run_id": run_id,
            "started_at": started_at,
            "status": if result.is_ok() { "succeeded" } else { "failed" },
            "overall_status": unified.overall_status,
            "feature_name": unified.feature_name,
            "ended_at": unified.ended_at,
            "input_stats": unified.input_stats,
            "output_stats": unified.output_stats,
            "business_metrics": unified.business_metrics,
            "business_issues": unified.business_issues,
            "output_paths": unified.output_paths,
            "diagnostic_log_path": unified.diagnostic_log_path,
            "summary": result.as_ref().ok(),
            "failure": result.as_ref().err().map(|error| error.root_cause().to_string()),
            "elapsed_ms": run_started.elapsed().as_millis(),
            "events": events,
        });
        match write_json_report_create_new(&path, &report) {
            Ok(()) => Some(path),
            Err(error) => {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("无法保存校验运行汇总：{error}")
                    )
                );
                None
            }
        }
    };

    if interactive && log_view == "user" {
        emit_unified_terminal_summary_with_outputs(&unified, &[]);
    }

    match result {
        Ok(result) => {
            if interactive && log_view == "user" {
                Ok(())
            } else if interactive && log_view == "off" {
                if let Some(path) = result.path.as_ref() {
                    println!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            Level::Info,
                            &format!("校验结果：{}", path_for_log_view(path, log_view))
                        )
                    );
                }
                Ok(())
            } else {
                println!(
                    "{}",
                    serde_json::to_string_pretty(
                        &serde_json::json!({"result": result, "run_summary": report_path})
                    )?
                );
                Ok(())
            }
        }
        Err(error) => {
            if interactive {
                Err(anyhow::anyhow!(
                    "按模板校验区域提取校验失败；请检查失败阶段、源工作簿副本和模板配置"
                ))
            } else {
                Err(error)
            }
        }
    }
}

fn run_summary_comments_cli(arguments: Vec<String>, interactive: bool) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let template = required_path(&options, "template")?;
    let output_dir = options
        .values
        .get("output-dir")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("output"));
    let log_view = options
        .values
        .get("log-view")
        .map(String::as_str)
        .unwrap_or("user");
    anyhow::ensure!(
        matches!(log_view, "user" | "developer" | "off"),
        "--log-view must be user, developer, or off"
    );
    let log_dir = options
        .values
        .get("log-dir")
        .map(PathBuf::from)
        .unwrap_or_else(pytools_core::run_logging::default_log_dir);
    if options.positional.is_empty() {
        bail!("provide at least one source workbook path");
    }

    let run_started = Instant::now();
    let started_at = Local::now().to_rfc3339();
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_nanos())
        .unwrap_or_default();
    let run_id = format!("pytools-pt-6-2-{}-{nonce}", std::process::id());
    use pytools_core::run_logging::{
        EventWriter, Level, PerformanceStats, RecordStats, RunTracker,
    };
    let mut tracker =
        RunTracker::with_id("PT-6-2", "按批注汇总", &options.positional, run_id.clone());
    let mut diagnostic = if log_view == "off" {
        None
    } else {
        match EventWriter::create(&log_dir.join("run-events"), "PT-6-2", &run_id) {
            Ok(writer) => Some(writer),
            Err(error) => {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("无法创建诊断日志：{error}；业务处理继续")
                    )
                );
                None
            }
        }
    };
    if let Some(writer) = diagnostic.as_mut() {
        let event = tracker.event(
            Level::Info,
            "run_started",
            "run",
            "started",
            "开始按批注汇总",
        );
        if let Err(error) = writer.write(&event) {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Warn,
                    &format!("诊断日志写入失败：{error}")
                )
            );
        }
    }
    let mut events = Vec::<SummaryRunEvent>::new();
    let mut sheet_stats = pytools_core::run_logging::SheetStats::default();
    let mut log_write_warning_emitted = false;
    let result =
        summarize_by_comments_with_progress(&template, &options.positional, &output_dir, |event| {
            if log_view == "developer" || log_view == "user" {
                if event.status == "failed" {
                    let mut contextual = event.clone();
                    contextual.user_message =
                        format!("按批注汇总阶段 {} 失败；请检查对应文件和模板", event.phase);
                    emit_summary_batch_event_view(log_view, &contextual);
                } else {
                    emit_summary_batch_event_view(log_view, &event);
                }
            }
            record_summary_progress_event(
                &mut tracker,
                diagnostic.as_mut(),
                &event,
                &options.positional,
                "row_extract",
                &mut sheet_stats,
                &mut log_write_warning_emitted,
            );
            events.push(event);
        });

    if result.is_err()
        && log_view != "off"
        && !events.last().is_some_and(|event| event.status == "failed")
    {
        let prior = events.last();
        let mut failure_event = SummaryRunEvent {
            sequence: events.len() + 1,
            timestamp: Local::now().to_rfc3339(),
            phase: "summary".to_owned(),
            status: "failed".to_owned(),
            user_message: "按批注汇总失败".to_owned(),
            progress: prior
                .map(|event| event.progress.clone())
                .unwrap_or_else(|| SummaryProgress {
                    total_inputs: options.positional.len(),
                    ..SummaryProgress::default()
                }),
            file_name: None,
            stage_elapsed_ms: None,
            total_elapsed_ms: run_started.elapsed().as_millis(),
            components: vec!["pytools-core".to_owned()],
            actual_engine: None,
            error_type: Some("summary_error".to_owned()),
        };
        failure_event.progress.current_input_index = None;
        failure_event.progress.current_sheet = None;
        record_summary_progress_event(
            &mut tracker,
            diagnostic.as_mut(),
            &failure_event,
            &options.positional,
            "row_extract",
            &mut sheet_stats,
            &mut log_write_warning_emitted,
        );
        if log_view == "developer" {
            eprintln!("{}", serde_json::to_string(&failure_event)?);
        } else {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Error,
                    "按批注汇总失败；详细阶段见运行记录"
                )
            );
        }
        events.push(failure_event);
    }

    if let Err(error) = result.as_ref() {
        let failed_event = events.iter().rev().find(|event| event.status == "failed");
        let stage = failed_event
            .map(|event| event.phase.as_str())
            .unwrap_or("summary");
        if let Some(writer) = diagnostic.as_mut() {
            let mut failure = tracker.event(
                Level::Error,
                "operation_failed",
                stage,
                "failed",
                "按批注汇总未完成；请检查关联文件和模板",
            );
            failure.error_code = Some(summary_event_error_code(
                failed_event.and_then(|event| event.error_type.as_deref()),
                stage,
            ));
            failure.file_name = failed_event.and_then(|event| event.file_name.clone());
            failure.sheet_name =
                failed_event.and_then(|event| event.progress.current_sheet.clone());
            if log_view == "developer" {
                failure.details = Some(serde_json::json!({"error_chain": format!("{error:#}")}));
            }
            if let Err(write_error) = writer.write(&failure) {
                if !log_write_warning_emitted {
                    eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            Level::Warn,
                            &format!("诊断日志写入失败：{write_error}；业务结果仍保留")
                        )
                    );
                    log_write_warning_emitted = true;
                }
            }
        }
    }

    let input_stats_available;
    let mut unified_summary = {
        let critical_failure = result.is_err()
            && events.last().is_some_and(|event| {
                event.phase.contains("output")
                    || event.phase == "prepare_output"
                    || event.error_type.as_deref().is_some_and(|kind| {
                        kind.contains("output")
                            || kind.contains("template")
                            || kind.contains("config")
                    })
            });
        let no_op = result.as_ref().is_ok_and(|summary| !summary.saved);
        let output_paths = result
            .as_ref()
            .ok()
            .and_then(|summary| summary.path.as_ref())
            .map(|path| vec![path_for_log_view(path, log_view)])
            .unwrap_or_default();
        let diagnostic_path = diagnostic.as_ref().map(|writer| {
            pytools_core::run_logging::diagnostic_path_for_log_view(
                writer.path(),
                log_view == "developer",
            )
        });
        let mut summary = tracker.finish(
            critical_failure,
            false,
            no_op,
            output_paths,
            diagnostic_path,
        );
        input_stats_available = summary.input_stats.pending == 0;
        summary.input_stats_available = input_stats_available;
        if sheet_stats.discovered > 0 || sheet_stats.attempted > 0 {
            summary.sheet_stats = Some(sheet_stats.clone());
        }
        if let Some(result) = result.as_ref().ok() {
            summary.record_stats = Some(RecordStats {
                rows_processed: Some(result.rows),
                rows_written: result.saved.then_some(result.rows),
            });
            summary.output_stats = Some(pytools_core::run_logging::OutputStats {
                workbooks_saved: Some(if result.saved { 1 } else { 0 }),
                sheets_saved: None,
                saved: Some(result.saved),
            });
            if let Some(execution) = &result.execution {
                summary.performance_stats = Some(PerformanceStats {
                    read_elapsed_ms: Some(execution.read_elapsed_ms),
                    write_elapsed_ms: Some(execution.write_elapsed_ms),
                });
            }
        } else {
            summary.output_stats = Some(pytools_core::run_logging::OutputStats {
                workbooks_saved: Some(0),
                sheets_saved: None,
                saved: Some(false),
            });
        }
        if result.is_err() {
            if let Some(event) = events.iter().rev().find(|event| event.status == "failed") {
                let code = summary_event_error_code(event.error_type.as_deref(), &event.phase);
                summary.error_summary = Some(pytools_core::run_logging::ErrorSummary {
                    stage: event.phase.clone(),
                    code,
                    message: "批注汇总未完成；请检查关联文件、模板与运行事件".to_owned(),
                    file_name: event.file_name.clone(),
                    sheet_name: event.progress.current_sheet.clone(),
                });
            }
        }
        summary
    };
    let unified_path = save_unified_run_summary(
        &tracker,
        &mut diagnostic,
        &mut unified_summary,
        &log_dir,
        log_view,
    );
    if interactive && log_view == "user" {
        emit_unified_terminal_summary_with_outputs(&unified_summary, &[]);
    }

    let run_log_path = if log_view == "off" {
        None
    } else {
        let path = log_dir.join(format!("PT-6-2_{run_id}.json"));
        let status = match &result {
            Ok(summary) if summary.saved => "succeeded",
            Ok(_) => "skipped",
            Err(_) => "failed",
        };
        let failure = result.as_ref().err().map(|_| {
            let event = events.last();
            serde_json::json!({
                "stage": event.map(|event| event.phase.as_str()).unwrap_or("summary"),
                "error_type": event.and_then(|event| event.error_type.as_deref()).unwrap_or("summary_error")
            })
        });
        let report = serde_json::json!({
            "schema_version": 1,
            "product": "pytools",
            "feature_id": "PT-6-2",
            "function": "summary-comments",
            "run_id": &run_id,
            "started_at": &started_at,
            "status": status,
            "total_inputs": options.positional.len(),
            "completed_files": events.last().map(|event| event.progress.completed_files).unwrap_or(0),
            "failed_files": events.last().map(|event| event.progress.failed_files).unwrap_or(0),
            "summary": result.as_ref().ok(),
            "failure": failure,
            "elapsed_ms": run_started.elapsed().as_millis(),
            "events": events,
        });
        match write_json_report_create_new(&path, &report) {
            Ok(()) => Some(path),
            Err(error) => {
                eprintln!("pytools: warning: could not save PT-6-2 run history: {error:#}");
                None
            }
        }
    };

    match result {
        Ok(result) => {
            let mut output = serde_json::to_value(result)?;
            if log_view != "off" {
                output["run_id"] = serde_json::Value::String(run_id.clone());
                if let Some(path) = &unified_path {
                    output["run_summary_path"] =
                        serde_json::Value::String(display_path_for_user(&path).to_string());
                }
                if let Some(path) = run_log_path {
                    output["run_log_path"] =
                        serde_json::Value::String(display_path_for_user(&path).to_string());
                    if log_view == "user" {
                        eprintln!(
                            "{}",
                            pytools_core::run_logging::terminal_line(
                                Level::Info,
                                &format!(
                                    "运行记录已保存：{}",
                                    pytools_core::run_logging::path_for_log_view(&path, false)
                                )
                            )
                        );
                    }
                }
            }
            if interactive && log_view == "user" {
                Ok(())
            } else {
                println!("{}", serde_json::to_string_pretty(&output)?);
                Ok(())
            }
        }
        Err(error) => {
            if log_view == "user" {
                if let Some(path) = run_log_path {
                    eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            Level::Info,
                            &format!(
                                "运行记录已保存：{}",
                                pytools_core::run_logging::path_for_log_view(&path, false)
                            )
                        )
                    );
                }
                if let Some(path) = unified_path {
                    eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            Level::Info,
                            &format!(
                                "统一记录：{}",
                                pytools_core::run_logging::path_for_log_view(&path, false)
                            )
                        )
                    );
                }
            }
            Err(error)
        }
    }
}

fn run_archive_cli(arguments: Vec<String>, operation: &str) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let result = execute_archive_operation(&options, operation)?;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn execute_archive_operation(options: &ParsedCli, operation: &str) -> Result<ArchiveResult> {
    let output_dir = options
        .values
        .get("output-dir")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("output"));
    let result = match operation {
        "date-preview" => archive_by_date_preview(
            &options.positional,
            &required_path(&options, "target-root")?,
            &output_dir,
            options.values.get("config").map(Path::new),
        )?,
        "type-preview" => archive_by_type_preview(
            &options.positional,
            &required_path(&options, "target-root")?,
            &output_dir,
            &required_path(&options, "config")?,
        )?,
        "date-copy" => archive_by_date_copy(
            &options.positional,
            &required_path(&options, "target-root")?,
            &output_dir,
            options.values.get("config").map(Path::new),
        )?,
        "type-copy" => archive_by_type_copy(
            &options.positional,
            &required_path(&options, "target-root")?,
            &output_dir,
            &required_path(&options, "config")?,
        )?,
        "plan" => execute_archive_plan(&required_path(&options, "plan")?, &output_dir)?,
        _ => unreachable!(),
    };
    Ok(result)
}

fn run_archive_menu_operation(
    arguments: Vec<String>,
    operation: &str,
    feature_id: &str,
    description: &str,
    input_paths: usize,
    global: &GlobalConfig,
    log_view: &str,
) -> Result<()> {
    let run_started = Instant::now();
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_nanos())
        .unwrap_or_default();
    let run_id = format!(
        "{}_{}_{}",
        Local::now().format("%Y%m%dT%H%M%S"),
        std::process::id(),
        nonce
    );
    use pytools_core::run_logging::{EventWriter, FileStatus, Level, RunTracker};
    let mut tracker = RunTracker::with_id(feature_id, description, &[], run_id.clone());
    let mut diagnostic = if log_view == "off" {
        None
    } else {
        match EventWriter::create(&global.log_dir.join("run-events"), feature_id, &run_id) {
            Ok(writer) => Some(writer),
            Err(error) => {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("无法创建诊断日志：{error}；业务处理继续")
                    )
                );
                None
            }
        }
    };
    if let Some(writer) = diagnostic.as_mut() {
        if let Err(error) = writer.write(&tracker.event(
            Level::Info,
            "run_started",
            "archive",
            "started",
            description,
        )) {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Warn,
                    &format!("诊断日志写入失败：{error}")
                )
            );
        }
    }
    let actual_engines = archive_actual_engines(operation);
    let started = archive_run_event(
        feature_id,
        &run_id,
        1,
        "archive_operation",
        "started",
        &format!("开始：{description}，输入路径 {input_paths} 项"),
        input_paths,
        None,
        0,
        &actual_engines,
        None,
    );
    if log_view != "off" {
        emit_archive_run_event(log_view, &started);
    }

    let options = parse_cli_arguments(arguments)?;
    let result = execute_archive_operation(&options, operation);
    let elapsed_ms = run_started.elapsed().as_millis();
    let result = match result {
        Ok(result) => result,
        Err(error) => {
            let mut summary = tracker.finish(
                true,
                false,
                false,
                vec![],
                diagnostic_path_for_log_view(diagnostic.as_ref(), log_view),
            );
            summary.input_stats_available = false;
            summary.error_summary = Some(pytools_core::run_logging::ErrorSummary {
                stage: "archive_operation".to_owned(),
                code: "ARCHIVE_OPERATION_FAILED".to_owned(),
                message: format!("{description}失败：{error:#}"),
                file_name: None,
                sheet_name: None,
            });
            if let Some(writer) = diagnostic.as_mut() {
                if let Err(log_error) = writer.write(&tracker.event(
                    Level::Error,
                    "run_finished",
                    "archive",
                    "failed",
                    format!("{description}失败"),
                )) {
                    eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            Level::Warn,
                            &format!("诊断日志写入失败：{log_error}")
                        )
                    );
                }
            }
            if log_view != "off" {
                let path = global
                    .log_dir
                    .join("run-unified")
                    .join(format!("{feature_id}_{run_id}.json"));
                if let Err(log_error) = write_json_report_create_new(&path, &summary) {
                    eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            Level::Warn,
                            &format!("统一运行记录写入失败：{log_error}")
                        )
                    );
                }
            }
            let failed = archive_run_event(
                feature_id,
                &run_id,
                2,
                "archive_operation",
                "failed",
                &format!("{description}失败；详细错误见终端错误输出"),
                input_paths,
                None,
                elapsed_ms,
                &actual_engines,
                Some("archive_operation_failed"),
            );
            if log_view != "off" {
                emit_archive_run_event(log_view, &failed);
                let events = [started, failed];
                match write_archive_run_history(
                    &global.log_dir,
                    feature_id,
                    &run_id,
                    "failed",
                    input_paths,
                    elapsed_ms,
                    &actual_engines,
                    None,
                    Some("archive_operation_failed"),
                    &events,
                ) {
                    Ok(path) if log_view == "user" => eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            Level::Info,
                            &format!(
                                "归档运行记录已保存：{}",
                                pytools_core::run_logging::path_for_log_view(&path, false)
                            )
                        )
                    ),
                    Ok(path) => eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            Level::Info,
                            &format!("{feature_id} run history: {}", display_path_for_user(&path))
                        )
                    ),
                    Err(history_error) => {
                        eprintln!("{}", pytools_core::run_logging::terminal_line(Level::Warn, &format!("pytools: could not save {feature_id} run history: {history_error:#}")))
                    }
                }
            }
            if log_view == "user" {
                emit_unified_terminal_summary_with_outputs(&summary, &[]);
            }
            return Err(error);
        }
    };

    let discovered = result
        .rows
        .iter()
        .map(|row| PathBuf::from(&row.source_path))
        .collect::<Vec<_>>();
    let _ = tracker.set_inputs(&discovered);
    tracker.set_discovered(discovered.len());
    for (index, row) in result.rows.iter().enumerate() {
        let position = index + 1;
        let state = if row.status.starts_with("失败") {
            FileStatus::Failed
        } else if row.status.starts_with("跳过") {
            FileStatus::Skipped
        } else {
            FileStatus::Succeeded
        };
        if state != FileStatus::Skipped {
            let _ = tracker.transition(position, FileStatus::InProgress);
        }
        let _ = tracker.transition(position, state);
        if let Some(writer) = diagnostic.as_mut() {
            let level = match state {
                FileStatus::Failed => Level::Error,
                FileStatus::Skipped => Level::Warn,
                _ => Level::Info,
            };
            let status = match state {
                FileStatus::Failed => "failed",
                FileStatus::Skipped => "skipped",
                _ => "succeeded",
            };
            let item_message = match state {
                FileStatus::Succeeded => format!("{description}：文件已归档"),
                FileStatus::Skipped => format!("{description}：文件被跳过"),
                FileStatus::Failed => format!("{description}：文件归档失败"),
                _ => format!("{description}：文件归档状态为 {status}"),
            };
            let mut event = tracker.event(level, "file_result", "archive", status, item_message);
            event.file_index = Some(position);
            event.file_name = Some(row.file_name.clone());
            if state == FileStatus::Failed {
                event.error_code = Some("FILE_ARCHIVE_FAILED".to_owned());
            }
            if let Err(log_error) = writer.write(&event) {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("诊断日志写入失败：{log_error}")
                    )
                );
                break;
            }
        }
    }
    let mut unified = tracker.finish(
        false,
        false,
        result.rows.is_empty(),
        vec![path_for_log_view(&result.output, log_view)],
        diagnostic_path_for_log_view(diagnostic.as_ref(), log_view),
    );
    unified
        .business_metrics
        .insert("发现文件数".to_owned(), serde_json::json!(result.files));
    unified
        .business_metrics
        .insert("无效条目数".to_owned(), serde_json::json!(result.invalid));
    unified
        .business_metrics
        .insert("已复制数".to_owned(), serde_json::json!(result.copied));
    unified
        .business_metrics
        .insert("跳过数".to_owned(), serde_json::json!(result.skipped));
    unified
        .business_metrics
        .insert("失败数".to_owned(), serde_json::json!(result.failed));
    if result.invalid > 0 {
        unified.business_issues.push(
            pytools_core::run_logging::BusinessIssue::new(
                "无效归档项",
                "发现无法按当前规则归档的输入项",
            )
            .count(result.invalid),
        );
    }
    if let Some(row) = result
        .rows
        .iter()
        .find(|row| row.status.starts_with("失败"))
    {
        unified.error_summary = Some(pytools_core::run_logging::ErrorSummary {
            stage: "archive_copy".to_owned(),
            code: "FILE_ARCHIVE_FAILED".to_owned(),
            message: if row.note.trim().is_empty() {
                row.status.clone()
            } else {
                row.note.clone()
            },
            file_name: Some(row.file_name.clone()),
            sheet_name: None,
        });
    }
    if let Some(writer) = diagnostic.as_mut() {
        if let Err(log_error) = writer.write(&tracker.event(
            if result.failed > 0 {
                Level::Error
            } else {
                Level::Info
            },
            "run_finished",
            "archive",
            &format!("{:?}", unified.overall_status),
            format!("{description}结束"),
        )) {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Warn,
                    &format!("诊断日志写入失败：{log_error}")
                )
            );
        }
    }
    if log_view != "off" {
        let mut report = serde_json::to_value(&unified)?;
        report["selected_path_count"] = serde_json::json!(input_paths);
        report["archive_stats"] = serde_json::json!({"files": result.files, "invalid": result.invalid, "copied": result.copied, "skipped": result.skipped, "failed": result.failed});
        let path = global
            .log_dir
            .join("run-unified")
            .join(format!("{feature_id}_{run_id}.json"));
        match write_json_report_create_new(&path, &report) {
            Ok(()) if log_view == "user" => eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Info,
                    &format!(
                        "运行编号 {run_id}；运行记录：{}",
                        pytools_core::run_logging::path_for_log_view(&path, false)
                    )
                )
            ),
            Ok(()) => {}
            Err(log_error) => eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Warn,
                    &format!("统一运行记录写入失败：{log_error}")
                )
            ),
        }
    }

    let status = if result.invalid > 0 || result.skipped > 0 || result.failed > 0 {
        "partial"
    } else {
        "succeeded"
    };
    let result_summary = serde_json::json!({
        "files": result.files,
        "invalid": result.invalid,
        "copied": result.copied,
        "skipped": result.skipped,
        "failed": result.failed,
    });
    let completed = archive_run_event(
        feature_id,
        &run_id,
        2,
        "archive_operation",
        status,
        &format!(
            "{description}完成：文件 {}，无效 {}，已复制 {}，跳过 {}，失败 {}，用时 {:.3} 秒",
            result.files,
            result.invalid,
            result.copied,
            result.skipped,
            result.failed,
            (elapsed_ms) as f64 / 1000.0
        ),
        input_paths,
        Some(&result),
        elapsed_ms,
        &actual_engines,
        None,
    );
    if log_view != "off" {
        if log_view == "developer" {
            emit_archive_run_event(log_view, &completed);
        }
        let events = [started, completed];
        match write_archive_run_history(
            &global.log_dir,
            feature_id,
            &run_id,
            status,
            input_paths,
            elapsed_ms,
            &actual_engines,
            Some(result_summary),
            None,
            &events,
        ) {
            Ok(path) if log_view == "user" => eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Info,
                    &format!(
                        "归档运行记录已保存：{}",
                        pytools_core::run_logging::path_for_log_view(&path, false)
                    )
                )
            ),
            Ok(path) => eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Info,
                    &format!("{feature_id} run history: {}", display_path_for_user(&path))
                )
            ),
            Err(error) => {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("pytools: could not save {feature_id} run history: {error:#}")
                    )
                )
            }
        }
    }

    if log_view == "developer" {
        println!("{}", serde_json::to_string_pretty(&result)?);
    } else {
        if log_view == "user" {
            for row in &result.rows {
                let level = if row.status.starts_with("失败") {
                    Some(Level::Error)
                } else if row.status.starts_with("无效") {
                    Some(Level::Warn)
                } else {
                    None
                };
                if let Some(level) = level {
                    let detail = if row.note.trim().is_empty() {
                        row.status.as_str()
                    } else {
                        row.note.as_str()
                    };
                    eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            level,
                            &format!("{}：{}", row.file_name, detail)
                        )
                    );
                }
            }
        }
        emit_unified_terminal_summary_with_outputs(&unified, std::slice::from_ref(&result.output));
    }
    Ok(())
}

fn archive_actual_engines(operation: &str) -> Vec<&'static str> {
    let mut engines = vec![
        "Rust std::fs",
        "excel-reader/Calamine",
        "excel-writer/rust_xlsxwriter",
    ];
    if matches!(operation, "date-copy" | "type-copy" | "plan") {
        engines.push("Rust std::fs::copy");
    }
    engines
}

fn archive_run_event(
    feature_id: &str,
    run_id: &str,
    sequence: usize,
    phase: &str,
    status: &str,
    user_message: &str,
    input_paths: usize,
    result: Option<&ArchiveResult>,
    elapsed_ms: u128,
    actual_engines: &[&str],
    error_type: Option<&str>,
) -> serde_json::Value {
    serde_json::json!({
        "sequence": sequence,
        "timestamp": Local::now().to_rfc3339(),
        "feature_id": feature_id,
        "run_id": run_id,
        "phase": phase,
        "status": status,
        "user_message": user_message,
        "progress": {
            "input_paths": input_paths,
            "processed_files": result.map(|value| value.files + value.invalid).unwrap_or(0),
            "copied": result.map(|value| value.copied).unwrap_or(0),
            "skipped": result.map(|value| value.skipped).unwrap_or(0),
            "failed": result.map(|value| value.failed).unwrap_or(0),
        },
        "stage_elapsed_ms": if sequence == 1 { None } else { Some(elapsed_ms) },
        "total_elapsed_ms": elapsed_ms,
        "components": ["pytools-core", "excel-reader", "excel-writer"],
        "actual_engines": actual_engines,
        "error_type": error_type,
    })
}

fn emit_archive_run_event(log_view: &str, event: &serde_json::Value) {
    match log_view {
        "developer" => match serde_json::to_string(event) {
            Ok(encoded) => eprintln!("{encoded}"),
            Err(_) => {
                eprintln!("{{\"phase\":\"archive_operation\",\"status\":\"serialization_failed\"}}")
            }
        },
        "user" => eprintln!(
            "{}",
            pytools_core::run_logging::terminal_line(
                if event["status"] == "failed" {
                    pytools_core::run_logging::Level::Error
                } else {
                    pytools_core::run_logging::Level::Info
                },
                event["user_message"].as_str().unwrap_or("归档处理中")
            )
        ),
        _ => {}
    }
}

fn write_archive_run_history(
    log_dir: &Path,
    feature_id: &str,
    run_id: &str,
    status: &str,
    input_paths: usize,
    total_elapsed_ms: u128,
    actual_engines: &[&str],
    result: Option<serde_json::Value>,
    error_type: Option<&str>,
    events: &[serde_json::Value],
) -> Result<PathBuf> {
    fs::create_dir_all(log_dir).with_context(|| {
        format!(
            "create archive log directory {}",
            display_path_for_user(&log_dir)
        )
    })?;
    let report = serde_json::json!({
        "schema_version": 1,
        "product": "pytools",
        "feature_id": feature_id,
        "run_id": run_id,
        "status": status,
        "input_paths": input_paths,
        "total_elapsed_ms": total_elapsed_ms,
        "actual_engines": actual_engines,
        "result": result,
        "error_type": error_type,
        "events": events,
    });
    let path = log_dir.join(format!("{feature_id}_{run_id}.run.json"));
    write_json_report_create_new(&path, &report)?;
    Ok(path)
}

struct ParsedCli {
    values: std::collections::BTreeMap<String, String>,
    positional: Vec<PathBuf>,
}
fn parse_cli_arguments(arguments: Vec<String>) -> Result<ParsedCli> {
    let mut values = std::collections::BTreeMap::new();
    let mut positional = Vec::new();
    let mut args = arguments.into_iter();
    while let Some(value) = args.next() {
        if value == "--help" || value == "-h" {
            print_usage();
            return Ok(ParsedCli { values, positional });
        }
        if let Some(name) = value.strip_prefix("--") {
            let argument = args
                .next()
                .with_context(|| format!("--{name} requires a value"))?;
            values.insert(name.to_owned(), argument);
        } else if value.starts_with('-') {
            bail!("unknown option {value:?}");
        } else {
            positional.push(PathBuf::from(value));
        }
    }
    Ok(ParsedCli { values, positional })
}

fn extract_boolean_flag(arguments: Vec<String>, name: &str) -> Result<(bool, Vec<String>)> {
    let flag = format!("--{name}");
    let inline_prefix = format!("{flag}=");
    anyhow::ensure!(
        !arguments
            .windows(2)
            .any(|pair| { pair[0] == flag && matches!(pair[1].as_str(), "true" | "false") }),
        "{flag} is a boolean switch and does not take a value"
    );
    let mut found = false;
    let mut filtered = Vec::with_capacity(arguments.len());
    for argument in arguments {
        if argument == flag {
            anyhow::ensure!(!found, "{flag} may only be specified once");
            found = true;
        } else if argument.starts_with(&inline_prefix) {
            bail!("{flag} is a boolean switch and does not take a value");
        } else {
            filtered.push(argument);
        }
    }
    Ok((found, filtered))
}
fn required_path(options: &ParsedCli, name: &str) -> Result<PathBuf> {
    options
        .values
        .get(name)
        .map(PathBuf::from)
        .with_context(|| format!("--{name} is required"))
}

fn unique_workbook_candidate(input: &Path, output_dir: &Path, suffix: &str) -> PathBuf {
    use std::hash::{Hash, Hasher};
    let stem = input.file_stem().unwrap_or_default().to_string_lossy();
    let mut hasher = std::collections::hash_map::DefaultHasher::new();
    input.hash(&mut hasher);
    let base = format!("{stem}_{suffix}_{:08x}", hasher.finish() as u32);
    let mut candidate = output_dir.join(format!("{base}.xlsx"));
    let mut index = 1usize;
    while candidate.exists() {
        candidate = output_dir.join(format!("{base}_{index}.xlsx"));
        index += 1;
    }
    candidate
}

fn unique_workbook_candidate_preserving_extension(
    input: &Path,
    output_dir: &Path,
    suffix: &str,
) -> PathBuf {
    use std::hash::{Hash, Hasher};
    let stem = input.file_stem().unwrap_or_default().to_string_lossy();
    let extension = input
        .extension()
        .and_then(|extension| extension.to_str())
        .unwrap_or("xls");
    let mut hasher = std::collections::hash_map::DefaultHasher::new();
    input.hash(&mut hasher);
    let base = format!("{stem}_{suffix}_{:08x}", hasher.finish() as u32);
    let mut candidate = output_dir.join(format!("{base}.{extension}"));
    let mut index = 1usize;
    while candidate.exists() {
        candidate = output_dir.join(format!("{base}_{index}.{extension}"));
        index += 1;
    }
    candidate
}

fn create_rust_report_copy(
    input: &Path,
    candidate: &Path,
    plan: &excel_editor::ExcelEditPlan,
    expected_source_hash: Option<&str>,
) -> Result<serde_json::Value> {
    pytools_core::institution_normalize::create_rust_new_workbook_from_cell_plan(
        input, candidate, plan,
    )?;
    verify_planning_snapshot(input, expected_source_hash)?;
    let output = specialized_report_copy_path(input, "PyTools_专用报表处理结果")?;
    create_target_from_candidate(candidate, &output).with_context(|| {
        format!(
            "publish Rust report copy to {}",
            display_path_for_user(&output)
        )
    })?;
    Ok(serde_json::json!({
        "engine": "rust_xlsxwriter",
        "execution": "created_and_validated_workbook_copy",
        "output": output,
        "operation_count": plan.operations.len()
    }))
}

fn specialized_report_copy_path(input: &Path, directory_name: &str) -> Result<PathBuf> {
    anyhow::ensure!(
        !input
            .components()
            .any(|component| component.as_os_str() == directory_name),
        "refusing to reprocess a workbook already inside {directory_name}"
    );
    let parent = input
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let result_dir = parent.join(directory_name);
    fs::create_dir_all(&result_dir).with_context(|| {
        format!(
            "create report result directory {}",
            display_path_for_user(&result_dir)
        )
    })?;
    let metadata = fs::symlink_metadata(&result_dir).with_context(|| {
        format!(
            "inspect report result directory {}",
            display_path_for_user(&result_dir)
        )
    })?;
    anyhow::ensure!(
        metadata.is_dir() && !metadata.file_type().is_symlink(),
        "report result path must be a real directory: {}",
        display_path_for_user(&result_dir)
    );
    let source_name = input
        .file_name()
        .context("source workbook has no file name")?;
    let stem = input
        .file_stem()
        .context("source workbook has no file stem")?
        .to_string_lossy();
    let extension = input
        .extension()
        .context("source workbook has no extension")?
        .to_string_lossy();
    let first = result_dir.join(source_name);
    if !path_has_directory_entry(&first)? {
        return Ok(first);
    }
    for suffix in 1..=10_000usize {
        let candidate = result_dir.join(format!("{stem}_{suffix}.{extension}"));
        if !path_has_directory_entry(&candidate)? {
            return Ok(candidate);
        }
    }
    anyhow::bail!(
        "could not choose a unique report output name beside {}",
        display_path_for_user(&input)
    )
}

fn path_has_directory_entry(path: &Path) -> Result<bool> {
    match fs::symlink_metadata(path) {
        Ok(_) => Ok(true),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(false),
        Err(error) => {
            Err(error).with_context(|| format!("inspect path {}", display_path_for_user(&path)))
        }
    }
}

fn remove_regular_candidate(path: &Path) -> Result<()> {
    let metadata = match fs::symlink_metadata(path) {
        Ok(metadata) => metadata,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => return Ok(()),
        Err(error) => {
            return Err(error).with_context(|| format!("inspect {}", display_path_for_user(&path)))
        }
    };
    anyhow::ensure!(
        metadata.is_file() && !metadata.file_type().is_symlink(),
        "refusing to remove non-regular candidate {}",
        display_path_for_user(&path)
    );
    fs::remove_file(path).with_context(|| {
        format!(
            "remove temporary candidate {}",
            display_path_for_user(&path)
        )
    })
}

fn run_local_filename_map(arguments: Vec<String>) -> Result<()> {
    let mut arguments = arguments.into_iter();
    let mut config_path = None::<PathBuf>;
    let mut log_dir = None::<PathBuf>;
    let mut log_view = None::<String>;
    let mut files = Vec::new();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--config" => {
                config_path = Some(PathBuf::from(
                    arguments.next().context("--config requires a path")?,
                ));
            }
            "--log-dir" => {
                log_dir = Some(PathBuf::from(
                    arguments.next().context("--log-dir requires a path")?,
                ));
            }
            "--log-view" => {
                log_view = Some(arguments.next().context("--log-view requires a value")?);
            }
            "--help" | "-h" => {
                print_usage();
                return Ok(());
            }
            value if value.starts_with('-') => bail!("unknown option {value:?}"),
            path => files.push(PathBuf::from(path)),
        }
    }
    let config_path = config_path.context("--config is required")?;
    if files.is_empty() {
        bail!("provide at least one file path");
    }
    let log_view = log_view.unwrap_or_else(|| "user".to_owned());
    anyhow::ensure!(
        matches!(log_view.as_str(), "user" | "developer" | "off"),
        "--log-view must be user, developer, or off"
    );
    let log_dir = log_dir.unwrap_or_else(pytools_core::run_logging::default_log_dir);

    let mut events = Vec::<serde_json::Value>::new();
    let mut active_stage = None;
    let mut unified =
        UnifiedRunSession::new(&log_dir, "PT-4-8", "按配置映射文件名", &log_view, None);
    unified.set_inputs(&files, Some(files.len()));
    let operation_result = if log_view == "off" {
        run_local_filename_map_candidate(&config_path, &files)
    } else {
        run_local_filename_map_with_progress(&config_path, &files, |event| {
            let event_value = local_filename_map_event_with_stage_elapsed(event, &mut active_stage);
            let level = match event.status.as_str() {
                "failed" => pytools_core::run_logging::Level::Error,
                "partial_failure" => pytools_core::run_logging::Level::Warn,
                _ => pytools_core::run_logging::Level::Info,
            };
            unified.progress(
                level,
                &event.phase,
                &event.phase,
                &event.status,
                &event.user_message,
                Some(event_value["stage_elapsed_ms"].as_u64().unwrap_or_default() as u128),
                event.actual_engine.as_deref(),
                event
                    .error_type
                    .as_ref()
                    .map(|error| serde_json::json!({"error_type": error})),
            );
            if log_view == "developer" {
                events.push(event_value.clone());
            }
            emit_local_filename_map_event_view(&log_view, event, &event_value);
        })
    };
    let history_result = if log_view == "developer" {
        Some(write_local_filename_map_run_history(&log_dir, &events))
    } else {
        None
    };
    let result = match operation_result {
        Ok(result) => result,
        Err(error) => {
            if let Some(Err(history_error)) = history_result.as_ref() {
                eprintln!("pytools: could not write local-map run history: {history_error:#}");
            }
            unified.mark_inputs_unavailable();
            unified.operation_failure(
                "configuration_or_rename",
                "FILE_RENAME_FAILED",
                "局部文件名映射失败",
                &error,
            );
            let (mut summary, _) = unified.finish(true, false, false, &[], None, None, None);
            summary.input_stats_available = false;
            if log_view == "user" {
                emit_unified_terminal_summary_with_outputs(&summary, &[]);
            }
            return Err(error);
        }
    };
    unified.set_discovered(result.rows.len());
    let mut output_paths = Vec::new();
    for (index, row) in result.rows.iter().enumerate() {
        let renamed = row.status == "已按局部映射改名";
        let failed = row.status == "跳过：改名失败";
        let status = if failed {
            pytools_core::run_logging::FileStatus::Failed
        } else if renamed {
            pytools_core::run_logging::FileStatus::Succeeded
        } else {
            pytools_core::run_logging::FileStatus::Skipped
        };
        unified.file_result(
            index + 1,
            status,
            !row.status.contains("非文件"),
            "rename",
            &row.status,
            failed.then_some("FILE_RENAME_FAILED"),
            None,
        );
        if renamed {
            let source = Path::new(&row.source_path);
            if let Some(parent) = source.parent() {
                output_paths.push(parent.join(&row.new_name));
            }
        }
    }
    unified.set_business_metric("映射规则数", serde_json::json!(result.rules));
    unified.set_business_metric("实际改名数", serde_json::json!(result.ok));
    unified.set_business_metric("业务报告跳过数", serde_json::json!(result.skip));
    let (unified_summary, unified_path) = unified.finish(
        false,
        false,
        result.ok == 0 && result.skip > 0,
        &output_paths,
        Some(pytools_core::run_logging::OutputStats {
            workbooks_saved: Some(result.ok),
            sheets_saved: None,
            saved: Some(result.ok > 0),
        }),
        None,
        None,
    );
    let (history_path, history_warning) = match history_result {
        Some(Ok(path)) => (Some(path), None),
        Some(Err(error)) => (None, Some(error.to_string())),
        None => (None, None),
    };
    let mut output = serde_json::to_value(result)?;
    if let Some(path) = history_path {
        output["run_history_path"] = serde_json::json!(path);
    }
    if let Some(warning) = history_warning {
        output["run_history_warning"] = serde_json::json!(warning);
    }
    output["run_id"] = serde_json::json!(unified_summary.run_id);
    if let Some(path) = unified_path {
        output["run_summary_path"] = serde_json::json!(path);
    }
    if log_view == "user" {
        emit_unified_terminal_summary_with_outputs(&unified_summary, &[]);
    }
    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}

fn local_filename_map_event_with_stage_elapsed(
    event: &LocalFilenameMapProgressEvent,
    active_stage: &mut Option<(String, u128)>,
) -> serde_json::Value {
    let mut value = serde_json::to_value(event).unwrap_or(serde_json::Value::Null);
    match event.status.as_str() {
        "started" => {
            *active_stage = Some((event.phase.clone(), event.elapsed_ms));
            value["stage_elapsed_ms"] = serde_json::json!(0);
        }
        "completed" | "failed" | "partial_failure" => {
            let stage_elapsed_ms = active_stage
                .take()
                .filter(|(phase, _)| phase == &event.phase)
                .map(|(_, started_ms)| event.elapsed_ms.saturating_sub(started_ms))
                .unwrap_or(event.elapsed_ms);
            value["stage_elapsed_ms"] = serde_json::json!(stage_elapsed_ms);
        }
        _ => {}
    }
    value
}

fn emit_local_filename_map_event_view(
    log_view: &str,
    event: &LocalFilenameMapProgressEvent,
    event_value: &serde_json::Value,
) {
    let current =
        (event.completed_files > 0).then_some(event.completed_files.min(event.total_files));
    emit_batch_event_view(
        log_view,
        event.total_files,
        current,
        &event.status,
        current.is_none() || matches!(event.status.as_str(), "completed" | "run_finished"),
        &event.user_message,
        Some(event_value),
        Some(event.elapsed_ms),
    );
}

fn write_local_filename_map_run_history(
    log_dir: &Path,
    events: &[serde_json::Value],
) -> Result<PathBuf> {
    fs::create_dir_all(log_dir)
        .with_context(|| format!("create log directory {}", display_path_for_user(&log_dir)))?;
    let timestamp = Local::now().format("%Y%m%d_%H%M%S");
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .context("system clock before Unix epoch")?
        .as_nanos();
    let path = log_dir.join(format!(
        "PT-4-8_local_filename_map_{timestamp}_{nonce}.run.json"
    ));
    let actual_engines = events
        .iter()
        .filter_map(|event| event["actual_engine"].as_str())
        .collect::<std::collections::BTreeSet<_>>();
    let last_status = events
        .last()
        .and_then(|event| event["status"].as_str())
        .unwrap_or("failed");
    let status = if last_status != "completed" {
        "failed"
    } else if events
        .iter()
        .any(|event| event["error_type"].as_str().is_some())
    {
        "partial_failure"
    } else {
        "succeeded"
    };
    let history = serde_json::json!({
        "schema_version": 1,
        "product": "pytools",
        "feature_id": "PT-4-8",
        "function": "rename-local-map",
        "status": status,
        "actual_engines": actual_engines,
        "total_elapsed_ms": events.last().map(|event| event["elapsed_ms"].clone()),
        "events": events,
    });
    let bytes = serde_json::to_vec_pretty(&history)?;
    let mut file = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&path)
        .with_context(|| format!("create run history {}", display_path_for_user(&path)))?;
    file.write_all(&bytes)?;
    file.write_all(b"\n")?;
    file.flush()?;
    Ok(path)
}

fn run_batch_rename(arguments: Vec<String>, interactive: bool) -> Result<()> {
    let mut arguments = arguments.into_iter();
    let mut config_path = None::<PathBuf>;
    let mut log_dir = None::<PathBuf>;
    let mut log_view = "user".to_owned();
    let mut paths = Vec::new();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--config" => {
                config_path = Some(PathBuf::from(
                    arguments.next().context("--config requires a path")?,
                ));
            }
            "--log-dir" => {
                log_dir = Some(PathBuf::from(
                    arguments.next().context("--log-dir requires a path")?,
                ));
            }
            "--log-view" => {
                log_view = arguments
                    .next()
                    .context("--log-view requires user, developer, or off")?;
            }
            "--help" | "-h" => {
                print_usage();
                return Ok(());
            }
            value if value.starts_with('-') => bail!("unknown option {value:?}"),
            path => paths.push(PathBuf::from(path)),
        }
    }

    let config_path = config_path.context("--config is required")?;
    anyhow::ensure!(
        matches!(log_view.as_str(), "user" | "developer" | "off"),
        "--log-view must be user, developer, or off"
    );
    let log_dir = log_dir.unwrap_or_else(pytools_core::run_logging::default_log_dir);
    let mut legacy_log = match open_batch_rename_log(&log_dir) {
        Ok(value) => Some(value),
        Err(error) => {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    pytools_core::run_logging::Level::Warn,
                    &format!("无法创建重命名日志：{error}；业务处理继续")
                )
            );
            None
        }
    };
    use pytools_core::run_logging::{EventWriter, FileStatus, Level, RunTracker};
    let mut tracker = RunTracker::new("PT-4-1", "按配置批量重命名", &paths);
    let mut diagnostic = if log_view == "off" {
        None
    } else {
        match EventWriter::create(&log_dir.join("run-events"), "PT-4-1", &tracker.run_id) {
            Ok(writer) => Some(writer),
            Err(error) => {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("无法创建诊断日志：{error}")
                    )
                );
                None
            }
        }
    };
    let mut events = Vec::<BatchRenameProgressEvent>::new();
    let mut failure_event = None::<BatchRenameProgressEvent>;
    let run_result = run_batch_rename_business(&config_path, &paths, |event| {
        emit_batch_rename_event_view(&log_view, event);
        if let Some(writer) = diagnostic.as_mut() {
            let level = if event.status == "failed" {
                Level::Error
            } else {
                Level::Info
            };
            let mut record = tracker.event(
                level,
                &event.phase,
                &event.phase,
                &event.status,
                &event.user_message,
            );
            record.timestamp = event.timestamp.clone();
            record.elapsed_ms = Some(event.stage_elapsed_ms);
            record.engine = event.actual_engine.clone();
            if event.status == "failed" {
                record.error_code = Some("CONFIG_READ_FAILED".to_owned());
            }
            if let Err(error) = writer.write(&record) {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("诊断日志写入失败：{error}")
                    )
                );
            }
        }
        if log_view == "developer" {
            events.push(event.clone());
        }
        if event.status == "failed" {
            failure_event = Some(event.clone());
        }
    });
    let history_result = if log_view == "developer" {
        Some(write_batch_rename_run_history(&log_dir, &events))
    } else {
        None
    };
    let result = match run_result {
        Ok(result) => result,
        Err(error) => {
            if let Some((_, file)) = legacy_log.as_mut() {
                if let Err(log_error) = write_batch_rename_failure_log(file, failure_event.as_ref())
                {
                    eprintln!("pytools: could not write batch-rename failure log: {log_error:#}");
                }
            }
            let summary = tracker.finish(
                true,
                false,
                false,
                vec![],
                diagnostic_path_for_log_view(diagnostic.as_ref(), &log_view),
            );
            if interactive && log_view == "user" {
                emit_unified_terminal_summary_with_outputs(&summary, &[]);
            }
            if let Some(writer) = diagnostic.as_mut() {
                let _ = writer.write(&tracker.event(
                    Level::Error,
                    "run_finished",
                    "configuration",
                    "failed",
                    "按配置批量重命名失败",
                ));
            }
            if log_view != "off" {
                let path = log_dir.join(format!("PT-4-1_{}.run.json", tracker.run_id));
                if let Err(log_error) = write_json_report_create_new(&path, &summary) {
                    eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            Level::Warn,
                            &format!("统一运行记录写入失败：{log_error}")
                        )
                    );
                }
            }
            if let Some(Err(history_error)) = history_result.as_ref() {
                eprintln!("pytools: could not write batch-rename run history: {history_error:#}");
            }
            return Err(error);
        }
    };
    let log_warning = legacy_log
        .as_mut()
        .and_then(|(_, file)| write_batch_rename_log(file, &result.rows).err());
    let mut outputs = Vec::new();
    for (index, row) in result.rows.iter().enumerate() {
        let position = index + 1;
        let state = if row.status == "renamed" {
            FileStatus::Succeeded
        } else if row.reason.starts_with("rename_failed")
            || row.reason == "not_exists"
            || row.reason.starts_with("full_name_no_code")
        {
            FileStatus::Failed
        } else {
            FileStatus::Skipped
        };
        let _ = tracker.transition(position, FileStatus::InProgress);
        let _ = tracker.transition(position, state);
        if state == FileStatus::Succeeded {
            if let Some(parent) = paths[index].parent() {
                outputs.push(parent.join(&row.new_name).to_string_lossy().into_owned());
            }
        }
        if let Some(writer) = diagnostic.as_mut() {
            let level = match state {
                FileStatus::Failed => Level::Error,
                FileStatus::Skipped => Level::Warn,
                _ => Level::Info,
            };
            let status = match state {
                FileStatus::Failed => "failed",
                FileStatus::Skipped => "skipped",
                _ => "succeeded",
            };
            let mut event = tracker.event(level, "file_result", "rename", status, "重命名文件结果");
            event.file_index = Some(position);
            event.file_name = Some(row.original_name.clone());
            if state == FileStatus::Failed {
                event.error_code = Some(
                    if row.reason == "not_exists" {
                        "FILE_NOT_FOUND"
                    } else if row.reason.starts_with("full_name_no_code") {
                        "CONFIG_MAPPING_MISSING"
                    } else {
                        "FILE_RENAME_FAILED"
                    }
                    .to_owned(),
                );
            }
            if let Err(error) = writer.write(&event) {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("诊断日志写入失败：{error}")
                    )
                );
                break;
            }
        }
    }
    let stats = tracker.input_stats();
    tracker.set_business_metric("实际改名数", serde_json::json!(stats.succeeded));
    tracker.set_business_metric("未改名条目", serde_json::json!(stats.skipped));
    tracker.set_business_metric("失败条目", serde_json::json!(stats.failed));
    let outputs = outputs
        .iter()
        .map(|path| path_for_log_view(Path::new(path), &log_view))
        .collect();
    let unified = tracker.finish(
        false,
        false,
        stats.succeeded == 0 && stats.failed == 0,
        outputs,
        diagnostic_path_for_log_view(diagnostic.as_ref(), &log_view),
    );
    if let Some(writer) = diagnostic.as_mut() {
        if let Err(error) = writer.write(&tracker.event(
            if stats.failed > 0 {
                Level::Error
            } else {
                Level::Info
            },
            "run_finished",
            "rename",
            &format!("{:?}", unified.overall_status),
            "批量重命名结束",
        )) {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Warn,
                    &format!("诊断日志写入失败：{error}")
                )
            );
        }
    }
    let mut unified_path = None;
    if log_view != "off" {
        let path = log_dir.join(format!("PT-4-1_{}.run.json", tracker.run_id));
        match write_json_report_create_new(&path, &unified) {
            Ok(()) => {
                if interactive {
                    eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            Level::Info,
                            &format!(
                                "运行编号 {}；运行记录：{}",
                                tracker.run_id,
                                pytools_core::run_logging::path_for_log_view(
                                    &path,
                                    log_view == "developer"
                                )
                            )
                        )
                    );
                }
                unified_path = Some(path);
            }
            Err(error) => eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Warn,
                    &format!("统一运行记录写入失败：{error}")
                )
            ),
        }
    }
    let (history_path, history_warning) = match history_result {
        Some(Ok(path)) => (Some(path), None),
        Some(Err(error)) => (None, Some(error.to_string())),
        None => (None, None),
    };
    let mut output = serde_json::to_value(result)?;
    if let Some(object) = output.as_object_mut() {
        object.insert(
            "log_path".to_owned(),
            serde_json::json!(legacy_log.as_ref().map(|(path, _)| path.to_string_lossy())),
        );
        object.insert("run_id".to_owned(), serde_json::json!(tracker.run_id));
        if let Some(path) = unified_path {
            object.insert(
                "run_summary_path".to_owned(),
                serde_json::json!(path.to_string_lossy()),
            );
        }
        if let Some(path) = history_path {
            object.insert(
                "run_history_path".to_owned(),
                serde_json::json!(path.to_string_lossy()),
            );
        }
        if let Some(warning) = log_warning {
            object.insert(
                "log_warning".to_owned(),
                serde_json::json!(warning.to_string()),
            );
        }
        if let Some(warning) = history_warning {
            object.insert("history_warning".to_owned(), serde_json::json!(warning));
        }
    }
    if !interactive || log_view == "developer" {
        println!("{}", serde_json::to_string_pretty(&output)?);
    } else if log_view == "user" {
        emit_unified_terminal_summary_with_outputs(&unified, &[]);
    }
    // off 视图：静默（不输出终端事件，也不输出机器 JSON）。
    Ok(())
}

fn emit_batch_rename_event_view(log_view: &str, event: &BatchRenameProgressEvent) {
    let current =
        (event.completed_files > 0).then_some(event.completed_files.min(event.total_files));
    let value = serde_json::to_value(event).ok();
    emit_batch_event_view(
        log_view,
        event.total_files,
        current,
        &event.status,
        current.is_none() || matches!(event.status.as_str(), "completed" | "run_finished"),
        &event.user_message,
        value.as_ref(),
        Some(event.total_elapsed_ms),
    );
}

fn write_batch_rename_run_history(
    log_dir: &Path,
    events: &[BatchRenameProgressEvent],
) -> Result<PathBuf> {
    fs::create_dir_all(log_dir)
        .with_context(|| format!("create log directory {}", display_path_for_user(&log_dir)))?;
    let timestamp = Local::now().format("%Y%m%d_%H%M%S");
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .context("system clock before Unix epoch")?
        .as_nanos();
    let path = log_dir.join(format!("PT-4-1_batch_rename_{timestamp}_{nonce}.json"));
    let history = serde_json::json!({
        "schema_version": 1,
        "product": "pytools",
        "feature_id": "PT-4-1",
        "total_elapsed_ms": events.last().map(|event| event.total_elapsed_ms),
        "events": events,
    });
    let bytes = serde_json::to_vec_pretty(&history)?;
    let mut file = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&path)
        .with_context(|| format!("create run history {}", display_path_for_user(&path)))?;
    file.write_all(&bytes)?;
    file.write_all(b"\n")?;
    file.flush()?;
    Ok(path)
}

fn open_batch_rename_log(log_dir: &Path) -> Result<(PathBuf, File)> {
    fs::create_dir_all(log_dir).with_context(|| {
        format!(
            "create batch-rename log directory {}",
            display_path_for_user(&log_dir)
        )
    })?;
    let timestamp = Local::now().format("%Y%m%d_%H%M%S");
    let base = format!("3_3_batch_rename_{timestamp}");
    for index in 0usize.. {
        let name = if index == 0 {
            format!("{base}.log")
        } else {
            format!("{base}_{index}.log")
        };
        let path = log_dir.join(name);
        match OpenOptions::new().write(true).create_new(true).open(&path) {
            Ok(file) => return Ok((path, file)),
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => {
                return Err(error)
                    .with_context(|| format!("create log {}", display_path_for_user(&path)))
            }
        }
    }
    unreachable!("usize log suffix range exhausted")
}

fn write_batch_rename_log(
    file: &mut File,
    rows: &[pytools_core::batch_rename::BatchRenameRow],
) -> Result<()> {
    for row in rows {
        let (level, message) = if row.status == "renamed" {
            (
                "INFO",
                format!("ok old={} new={}", row.original_name, row.new_name),
            )
        } else {
            let file = if row.original_name.is_empty() {
                row.source_path.as_str()
            } else {
                row.original_name.as_str()
            };
            let level = if row.reason.starts_with("rename_failed") {
                "WARNING"
            } else {
                "INFO"
            };
            let reason = row.reason.split_whitespace().next().unwrap_or("unknown");
            (
                level,
                format!(
                    "skip file={} reason={reason}",
                    pytools_core::run_logging::portable_file_name(Path::new(file))
                ),
            )
        };
        let line = format!(
            "{} [{level}] {message}",
            Local::now().format("%Y-%m-%d %H:%M:%S")
        );
        writeln!(file, "{line}")?;
    }
    file.flush()?;
    Ok(())
}

fn write_batch_rename_failure_log(
    file: &mut File,
    event: Option<&BatchRenameProgressEvent>,
) -> Result<()> {
    let (phase, error_type) = event
        .map(|event| {
            (
                event.phase.as_str(),
                event.error_type.as_deref().unwrap_or("unknown"),
            )
        })
        .unwrap_or(("task", "unknown"));
    writeln!(
        file,
        "{} [ERROR] phase={phase} error_type={error_type} message=批量重命名失败，请查看终端错误提示",
        Local::now().format("%Y-%m-%d %H:%M:%S")
    )?;
    file.flush()?;
    Ok(())
}

fn run_content_prefix_rename_cli(arguments: Vec<String>, interactive: bool) -> Result<()> {
    run_content_prefix_cli(arguments, true, interactive)
}

fn run_content_prefix_unrename_cli(arguments: Vec<String>, interactive: bool) -> Result<()> {
    run_content_prefix_cli(arguments, false, interactive)
}

fn run_content_prefix_cli(
    arguments: Vec<String>,
    add_prefix: bool,
    interactive: bool,
) -> Result<()> {
    let Some((config_path, files, log_dir, log_view)) =
        parse_content_prefix_cli_arguments(arguments)?
    else {
        return Ok(());
    };
    let (feature_id, function_name) = if add_prefix {
        ("PT-4-6", "rename-by-content")
    } else {
        ("PT-4-7", "unrename-by-content")
    };
    let feature_name = if add_prefix {
        "按工作簿内容添加文件名前缀"
    } else {
        "按工作簿内容取消文件名前缀"
    };
    let mut unified = UnifiedRunSession::new(&log_dir, feature_id, feature_name, &log_view, None);
    unified.set_inputs(&files, Some(files.len()));
    let mut events = Vec::<serde_json::Value>::new();
    let mut active_stage = None;
    let operation_result = if log_view == "off" {
        if add_prefix {
            run_content_prefix_rename_xlsx_candidate(&config_path, &files)
                .map(ContentPrefixOperationReport::Rename)
        } else {
            run_content_prefix_unrename_xlsx_candidate(&config_path, &files)
                .map(ContentPrefixOperationReport::Unrename)
        }
    } else if add_prefix {
        run_content_prefix_rename_xlsx_candidate_with_progress(&config_path, &files, |event| {
            let event_value = file_sort_event_with_stage_elapsed(event, &mut active_stage);
            let level = if event.status == "failed" {
                pytools_core::run_logging::Level::Error
            } else {
                pytools_core::run_logging::Level::Info
            };
            unified.progress(
                level,
                &event.phase,
                &event.phase,
                &event.status,
                &event.user_message,
                Some(event.elapsed_ms),
                event.actual_engine.as_deref(),
                event
                    .error_type
                    .as_ref()
                    .map(|error| serde_json::json!({"error_type": error})),
            );
            if log_view == "developer" {
                events.push(event_value.clone());
            }
            emit_file_sort_event_view(&log_view, event, &event_value);
        })
        .map(ContentPrefixOperationReport::Rename)
    } else {
        run_content_prefix_unrename_xlsx_candidate_with_progress(&config_path, &files, |event| {
            let event_value = file_sort_event_with_stage_elapsed(event, &mut active_stage);
            if log_view == "developer" {
                events.push(event_value.clone());
            }
            emit_file_sort_event_view(&log_view, event, &event_value);
        })
        .map(ContentPrefixOperationReport::Unrename)
    };
    let history_result = if log_view == "developer" {
        write_content_prefix_run_history(&log_dir, feature_id, function_name, &events).map(Some)
    } else {
        Ok(None)
    };
    if let Err(error) = &history_result {
        eprintln!("pytools: warning: could not write structured run history: {error:#}");
    }
    let result = match operation_result {
        Ok(result) => result,
        Err(error) => {
            unified.mark_inputs_unavailable();
            unified.operation_failure(
                "file_processing",
                "FILE_OPERATION_FAILED",
                "文件内容命名处理失败",
                &error,
            );
            let _ = unified.finish(
                true,
                false,
                false,
                &[],
                Some(pytools_core::run_logging::OutputStats {
                    workbooks_saved: None,
                    sheets_saved: None,
                    saved: None,
                }),
                None,
                None,
            );
            return Err(error);
        }
    };
    let (rows, summary_output) = match &result {
        ContentPrefixOperationReport::Rename(report) => (
            report
                .rows
                .iter()
                .map(|row| (row.source_path.clone(), row.status.clone()))
                .collect::<Vec<_>>(),
            (report.ok, report.skip),
        ),
        ContentPrefixOperationReport::Unrename(report) => (
            report
                .rows
                .iter()
                .map(|row| (row.source_path.clone(), row.status.clone()))
                .collect::<Vec<_>>(),
            (report.ok, report.skip),
        ),
    };
    unified.set_discovered(rows.len());
    for (index, (source_path, row_status)) in rows.iter().enumerate() {
        let status = if row_status.starts_with("已") {
            pytools_core::run_logging::FileStatus::Succeeded
        } else if row_status.contains("内容读取失败") || row_status.contains("改名失败") {
            pytools_core::run_logging::FileStatus::Failed
        } else {
            pytools_core::run_logging::FileStatus::Skipped
        };
        let code = match status {
            pytools_core::run_logging::FileStatus::Failed if row_status.contains("读取") => {
                Some("EXCEL_READ_FAILED")
            }
            pytools_core::run_logging::FileStatus::Failed => Some("FILE_RENAME_FAILED"),
            pytools_core::run_logging::FileStatus::Skipped if row_status.contains("格式") => {
                Some("FORMAT_UNSUPPORTED")
            }
            _ => None,
        };
        unified.file_result(
            index + 1,
            status,
            !row_status.contains("非文件"),
            "rename",
            row_status,
            code,
            None,
        );
    }
    unified.set_business_metric("实际改名数", serde_json::json!(summary_output.0));
    unified.set_business_metric("未改名条目", serde_json::json!(summary_output.1));
    unified.set_business_metric(
        "失败条目",
        serde_json::json!(unified.tracker.input_stats().failed),
    );
    let (unified_summary, unified_path) = unified.finish(
        false,
        false,
        summary_output.0 == 0 && summary_output.1 > 0,
        &[],
        None,
        None,
        None,
    );
    let mut output = match result {
        ContentPrefixOperationReport::Rename(report) => serde_json::to_value(report)?,
        ContentPrefixOperationReport::Unrename(report) => serde_json::to_value(report)?,
    };
    if let Ok(Some(path)) = &history_result {
        output["run_history_path"] = serde_json::json!(path);
    }
    output["run_id"] = serde_json::json!(unified_summary.run_id.clone());
    if let Some(path) = unified_path.as_ref() {
        output["run_summary_path"] = serde_json::json!(path);
    }
    if let Err(error) = &history_result {
        output["run_history_warning"] =
            serde_json::json!(format!("could not write structured run history: {error:#}"));
    }
    if interactive && log_view == "user" {
        emit_unified_terminal_summary_with_outputs(&unified_summary, &[]);
    } else {
        println!("{}", serde_json::to_string_pretty(&output)?);
    }
    Ok(())
}

enum ContentPrefixOperationReport {
    Rename(pytools_core::file_sort_prefix::ContentPrefixRenameRunReport),
    Unrename(pytools_core::file_sort_prefix::ContentPrefixUnrenameRunReport),
}

fn parse_content_prefix_cli_arguments(
    arguments: Vec<String>,
) -> Result<Option<(PathBuf, Vec<PathBuf>, PathBuf, String)>> {
    let mut arguments = arguments.into_iter();
    let mut config_path = None::<PathBuf>;
    let mut log_dir = None::<PathBuf>;
    let mut log_view = None::<String>;
    let mut files = Vec::new();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--config" => {
                config_path = Some(PathBuf::from(
                    arguments.next().context("--config requires a path")?,
                ));
            }
            "--log-dir" => {
                log_dir = Some(PathBuf::from(
                    arguments.next().context("--log-dir requires a path")?,
                ));
            }
            "--log-view" => {
                log_view = Some(arguments.next().context("--log-view requires a value")?);
            }
            "--help" | "-h" => {
                print_usage();
                return Ok(None);
            }
            value if value.starts_with('-') => bail!("unknown option {value:?}"),
            path => files.push(PathBuf::from(path)),
        }
    }

    let config_path = config_path.context("--config is required")?;
    if files.is_empty() {
        bail!("provide at least one file path");
    }
    let log_view = log_view.unwrap_or_else(|| "user".to_owned());
    anyhow::ensure!(
        matches!(log_view.as_str(), "user" | "developer" | "off"),
        "--log-view must be user, developer, or off"
    );
    Ok(Some((
        config_path,
        files,
        log_dir.unwrap_or_else(pytools_core::run_logging::default_log_dir),
        log_view,
    )))
}

fn write_content_prefix_run_history(
    log_dir: &Path,
    feature_id: &str,
    function_name: &str,
    events: &[serde_json::Value],
) -> Result<PathBuf> {
    fs::create_dir_all(log_dir)
        .with_context(|| format!("create log directory {}", display_path_for_user(&log_dir)))?;
    let timestamp = Local::now().format("%Y%m%d_%H%M%S");
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .context("system clock before Unix epoch")?
        .as_nanos();
    let path = log_dir.join(format!(
        "{feature_id}_content_prefix_{timestamp}_{nonce}.run.json"
    ));
    let actual_engines = events
        .iter()
        .filter_map(|event| event["actual_engine"].as_str())
        .collect::<std::collections::BTreeSet<_>>();
    let last_status = events
        .last()
        .and_then(|event| event["status"].as_str())
        .unwrap_or("failed");
    let run_status = if last_status != "completed" {
        "failed"
    } else if events
        .iter()
        .any(|event| event["error_type"].as_str().is_some())
    {
        "partial_failure"
    } else {
        "succeeded"
    };
    let history = serde_json::json!({
        "schema_version": 1,
        "product": "pytools",
        "feature_id": feature_id,
        "function": function_name,
        "status": run_status,
        "actual_engines": actual_engines,
        "total_elapsed_ms": events.last().map(|event| event["elapsed_ms"].clone()),
        "events": events,
    });
    let bytes = serde_json::to_vec_pretty(&history)?;
    let mut file = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&path)
        .with_context(|| format!("create run history {}", display_path_for_user(&path)))?;
    file.write_all(&bytes)?;
    file.write_all(b"\n")?;
    file.flush()?;
    Ok(path)
}

fn run_file_sort_prefix(arguments: Vec<String>, interactive: bool) -> Result<()> {
    let mut arguments = arguments.into_iter();
    let mut config_path = None::<PathBuf>;
    let mut output_dir = None::<PathBuf>;
    let mut log_dir = None::<PathBuf>;
    let mut log_view = None::<String>;
    let mut files = Vec::new();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--config" => {
                config_path = Some(PathBuf::from(
                    arguments.next().context("--config requires a path")?,
                ));
            }
            "--output-dir" => {
                output_dir = Some(PathBuf::from(
                    arguments.next().context("--output-dir requires a path")?,
                ));
            }
            "--log-dir" => {
                log_dir = Some(PathBuf::from(
                    arguments.next().context("--log-dir requires a path")?,
                ));
            }
            "--log-view" => {
                log_view = Some(arguments.next().context("--log-view requires a value")?);
            }
            "--help" | "-h" => {
                print_usage();
                return Ok(());
            }
            value if value.starts_with('-') => bail!("unknown option {value:?}"),
            path => files.push(PathBuf::from(path)),
        }
    }

    let config_path = config_path.context("--config is required")?;
    if files.is_empty() {
        bail!("provide at least one file path");
    }
    let output_dir = output_dir.unwrap_or_else(|| PathBuf::from("output"));
    let log_dir = log_dir.unwrap_or_else(pytools_core::run_logging::default_log_dir);
    let log_view = file_sort_log_view(log_view)?;
    let mut unified =
        UnifiedRunSession::new(&log_dir, "PT-4-9", "文件排序前缀添加", &log_view, None);
    unified.set_inputs(&files, Some(files.len()));
    let mut events = Vec::new();
    let mut active_stage = None;
    let operation_result =
        run_file_sort_prefix_with_progress_business(&config_path, &files, &output_dir, |event| {
            let event_value = file_sort_event_with_stage_elapsed(event, &mut active_stage);
            let level = if event.status == "failed" {
                pytools_core::run_logging::Level::Error
            } else {
                pytools_core::run_logging::Level::Info
            };
            unified.progress(
                level,
                &event.phase,
                &event.phase,
                &event.status,
                &event.user_message,
                Some(event.elapsed_ms),
                event.actual_engine.as_deref(),
                event
                    .error_type
                    .as_ref()
                    .map(|error| serde_json::json!({"error_type": error})),
            );
            if log_view == "developer" {
                events.push(event_value.clone());
            }
            emit_file_sort_event_view(&log_view, event, &event_value);
        });
    let history_result = if log_view == "developer" {
        write_file_sort_run_history(&log_dir, "PT-4-9", &events).map(Some)
    } else {
        Ok(None)
    };
    if let Err(error) = &history_result {
        eprintln!("pytools: warning: could not write structured run history: {error:#}");
    }
    let result = match operation_result {
        Ok(result) => result,
        Err(error) => {
            unified.mark_inputs_unavailable();
            unified.operation_failure(
                "file_sort",
                "FILE_SORT_FAILED",
                "文件排序前缀处理失败",
                &error,
            );
            let _ = unified.finish(true, false, false, &[], None, None, None);
            return Err(error);
        }
    };
    unified.set_discovered(result.rows.len());
    for (index, row) in result.rows.iter().enumerate() {
        let status = match row.status.as_str() {
            "失败" => pytools_core::run_logging::FileStatus::Failed,
            "未命中" | "跳过：非文件" | "跳过：无匹配排序前缀" => {
                pytools_core::run_logging::FileStatus::Skipped
            }
            _ => pytools_core::run_logging::FileStatus::Succeeded,
        };
        unified.file_result(
            index + 1,
            status,
            !row.status.contains("非文件"),
            "file_rename",
            &row.status,
            (status == pytools_core::run_logging::FileStatus::Failed)
                .then_some("FILE_RENAME_FAILED"),
            (status == pytools_core::run_logging::FileStatus::Failed).then(|| row.remark.clone()),
        );
    }
    unified.set_business_metric("实际改名数", serde_json::json!(result.renamed));
    unified.set_business_metric("未命中数", serde_json::json!(result.unmatched));
    unified.set_business_metric("跳过条目", serde_json::json!(result.skipped));
    unified.set_business_metric("失败条目", serde_json::json!(result.failed));
    let (unified_summary, unified_path) = unified.finish(
        false,
        false,
        result.renamed == 0 && result.failed == 0,
        std::slice::from_ref(&result.output),
        Some(pytools_core::run_logging::OutputStats {
            workbooks_saved: Some(1),
            sheets_saved: None,
            saved: Some(true),
        }),
        None,
        None,
    );
    let legacy_log_result = if log_view == "developer" {
        write_file_sort_legacy_log(&log_dir, &result.rows)
    } else {
        Ok(None)
    };
    if let Err(error) = &legacy_log_result {
        eprintln!("pytools: warning: could not write complete operation log: {error:#}");
    }
    if interactive && log_view == "user" {
        emit_unified_terminal_summary_with_outputs(&unified_summary, &[]);
        return Ok(());
    }
    let mut output = serde_json::to_value(result)?;
    output["run_id"] = serde_json::json!(unified_summary.run_id.clone());
    if let Some(path) = unified_path.as_ref() {
        output["run_summary_path"] = serde_json::json!(path);
    }
    if let Ok(Some(log_path)) = &legacy_log_result {
        output["log_path"] = serde_json::json!(log_path);
    }
    if let Ok(Some(history_path)) = &history_result {
        output["run_history_path"] = serde_json::json!(history_path);
    }
    if let Err(error) = &legacy_log_result {
        let warning = format!("could not write complete operation log: {error:#}");
        output["log_warning"] = serde_json::json!(warning);
    }
    if let Err(error) = &history_result {
        let warning = format!("could not write structured run history: {error:#}");
        output["run_history_warning"] = serde_json::json!(warning);
    }
    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}

fn file_sort_log_view(log_view: Option<String>) -> Result<String> {
    let log_view = log_view.unwrap_or_else(|| "user".to_owned());
    if !matches!(log_view.as_str(), "user" | "developer" | "off") {
        bail!("--log-view must be user, developer, or off");
    }
    Ok(log_view)
}

fn file_sort_event_with_stage_elapsed(
    event: &FileSortProgressEvent,
    active_stage: &mut Option<(String, u128)>,
) -> serde_json::Value {
    let mut value = serde_json::to_value(event).unwrap_or(serde_json::Value::Null);
    match event.status.as_str() {
        "started" => {
            *active_stage = Some((event.phase.clone(), event.elapsed_ms));
            value["stage_elapsed_ms"] = serde_json::json!(0);
        }
        "completed" | "failed" | "partial_failure" => {
            let stage_elapsed_ms = active_stage
                .take()
                .filter(|(phase, _)| phase == &event.phase)
                .map(|(_, started_ms)| event.elapsed_ms.saturating_sub(started_ms))
                .unwrap_or(event.elapsed_ms);
            value["stage_elapsed_ms"] = serde_json::json!(stage_elapsed_ms);
        }
        _ => {}
    }
    value
}

fn emit_file_sort_event_view(
    log_view: &str,
    event: &FileSortProgressEvent,
    event_value: &serde_json::Value,
) {
    let current =
        (event.completed_files > 0).then_some(event.completed_files.min(event.total_files));
    emit_batch_event_view(
        log_view,
        event.total_files,
        current,
        &event.status,
        current.is_none() || matches!(event.status.as_str(), "completed" | "run_finished"),
        &event.user_message,
        Some(event_value),
        Some(event.elapsed_ms),
    );
}

fn write_file_sort_run_history(
    log_dir: &Path,
    feature_id: &str,
    events: &[serde_json::Value],
) -> Result<PathBuf> {
    fs::create_dir_all(log_dir)
        .with_context(|| format!("create log directory {}", display_path_for_user(&log_dir)))?;
    let timestamp = Local::now().format("%Y%m%d_%H%M%S");
    let nonce = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .context("system clock before Unix epoch")?
        .as_nanos();
    let path = log_dir.join(format!("{feature_id}_file_sort_{timestamp}_{nonce}.json"));
    let history = serde_json::json!({
        "schema_version": 1,
        "product": "pytools",
        "feature_id": feature_id,
        "total_elapsed_ms": events.last().map(|event| event["elapsed_ms"].clone()),
        "events": events,
    });
    let bytes = serde_json::to_vec_pretty(&history)?;
    let mut file = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&path)
        .with_context(|| format!("create run history {}", display_path_for_user(&path)))?;
    file.write_all(&bytes)?;
    file.write_all(b"\n")?;
    file.flush()?;
    Ok(path)
}

fn write_file_sort_legacy_log(
    log_dir: &Path,
    rows: &[pytools_core::file_sort_prefix::FileSortPrefixRow],
) -> Result<Option<PathBuf>> {
    let (log_path, mut log_file) = open_file_sort_log(log_dir)?;
    write_file_sort_log(&mut log_file, rows)?;
    Ok(Some(log_path))
}

fn open_file_sort_log(log_dir: &Path) -> Result<(PathBuf, File)> {
    fs::create_dir_all(log_dir)
        .with_context(|| format!("create log directory {}", display_path_for_user(&log_dir)))?;
    let timestamp = Local::now().format("%Y%m%d_%H%M%S");
    let path = log_dir.join(format!("4_9_file_sort_prefix_{timestamp}.log"));
    let file = OpenOptions::new()
        .create(true)
        .append(true)
        .open(&path)
        .with_context(|| format!("open operation log {}", display_path_for_user(&path)))?;
    Ok((path, file))
}

fn write_file_sort_log(
    file: &mut File,
    rows: &[pytools_core::file_sort_prefix::FileSortPrefixRow],
) -> Result<()> {
    for row in rows {
        let message = if let Some(config_row) = row.config_row {
            Some((
                "INFO",
                format!(
                    "ok old={} new={} row={} prefix={}",
                    row.original_name, row.new_name, config_row, row.sort_prefix
                ),
            ))
        } else if row.status == "失败" {
            Some((
                "WARNING",
                format!("failed file={} err={}", row.source_path, row.remark),
            ))
        } else {
            None
        };
        if let Some((level, message)) = message {
            let line = format!(
                "{} [{level}] {message}",
                Local::now().format("%Y-%m-%d %H:%M:%S")
            );
            writeln!(file, "{line}")?;
            eprintln!("{line}");
        }
    }
    file.flush()?;
    Ok(())
}

/*
 * The removal command uses the same progress contract as addition.
 */
fn write_file_sort_removal_legacy_log(
    log_dir: &Path,
    rows: &[pytools_core::file_sort_prefix::FileSortPrefixRemovalRow],
) -> Result<Option<PathBuf>> {
    let (log_path, mut log_file) = open_file_sort_removal_log(log_dir)?;
    write_file_sort_removal_log(&mut log_file, rows)?;
    Ok(Some(log_path))
}

fn open_file_sort_removal_log(log_dir: &Path) -> Result<(PathBuf, File)> {
    fs::create_dir_all(log_dir)
        .with_context(|| format!("create log directory {}", display_path_for_user(&log_dir)))?;
    let timestamp = Local::now().format("%Y%m%d_%H%M%S");
    let path = log_dir.join(format!("4_10_remove_file_sort_prefix_{timestamp}.log"));
    let file = OpenOptions::new()
        .create(true)
        .append(true)
        .open(&path)
        .with_context(|| format!("open operation log {}", display_path_for_user(&path)))?;
    Ok((path, file))
}

fn write_file_sort_removal_log(
    file: &mut File,
    rows: &[pytools_core::file_sort_prefix::FileSortPrefixRemovalRow],
) -> Result<()> {
    for row in rows {
        let message = if !row.new_name.is_empty() {
            Some((
                "INFO",
                format!("ok old={} new={}", row.original_name, row.new_name),
            ))
        } else if row.status == "失败" {
            Some((
                "WARNING",
                format!("failed file={} err={}", row.source_path, row.remark),
            ))
        } else {
            None
        };
        if let Some((level, message)) = message {
            let line = format!(
                "{} [{level}] {message}",
                Local::now().format("%Y-%m-%d %H:%M:%S")
            );
            writeln!(file, "{line}")?;
            eprintln!("{line}");
        }
    }
    file.flush()?;
    Ok(())
}

fn run_remove_file_sort_prefix(arguments: Vec<String>, interactive: bool) -> Result<()> {
    let mut arguments = arguments.into_iter();
    let mut config_path = None::<PathBuf>;
    let mut output_dir = None::<PathBuf>;
    let mut log_dir = None::<PathBuf>;
    let mut log_view = None::<String>;
    let mut files = Vec::new();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--config" => {
                config_path = Some(PathBuf::from(
                    arguments.next().context("--config requires a path")?,
                ));
            }
            "--output-dir" => {
                output_dir = Some(PathBuf::from(
                    arguments.next().context("--output-dir requires a path")?,
                ));
            }
            "--log-dir" => {
                log_dir = Some(PathBuf::from(
                    arguments.next().context("--log-dir requires a path")?,
                ));
            }
            "--log-view" => {
                log_view = Some(arguments.next().context("--log-view requires a value")?);
            }
            "--help" | "-h" => {
                print_usage();
                return Ok(());
            }
            value if value.starts_with('-') => bail!("unknown option {value:?}"),
            path => files.push(PathBuf::from(path)),
        }
    }

    let config_path = config_path.context("--config is required")?;
    let output_dir = output_dir.unwrap_or_else(|| PathBuf::from("output"));
    let log_dir = log_dir.unwrap_or_else(pytools_core::run_logging::default_log_dir);
    let log_view = file_sort_log_view(log_view)?;
    let mut unified =
        UnifiedRunSession::new(&log_dir, "PT-4-10", "取消文件排序前缀", &log_view, None);
    unified.set_inputs(&files, Some(files.len()));
    let mut events = Vec::new();
    let mut active_stage = None;
    let operation_result = run_remove_file_sort_prefix_with_progress_business(
        &config_path,
        &files,
        &output_dir,
        |event| {
            let event_value = file_sort_event_with_stage_elapsed(event, &mut active_stage);
            let level = if event.status == "failed" {
                pytools_core::run_logging::Level::Error
            } else {
                pytools_core::run_logging::Level::Info
            };
            unified.progress(
                level,
                &event.phase,
                &event.phase,
                &event.status,
                &event.user_message,
                Some(event.elapsed_ms),
                event.actual_engine.as_deref(),
                event
                    .error_type
                    .as_ref()
                    .map(|error| serde_json::json!({"error_type": error})),
            );
            if log_view == "developer" {
                events.push(event_value.clone());
            }
            emit_file_sort_event_view(&log_view, event, &event_value);
        },
    );
    let history_result = if log_view == "developer" {
        write_file_sort_run_history(&log_dir, "PT-4-10", &events).map(Some)
    } else {
        Ok(None)
    };
    if let Err(error) = &history_result {
        eprintln!("pytools: warning: could not write structured run history: {error:#}");
    }
    let result = match operation_result {
        Ok(result) => result,
        Err(error) => {
            unified.mark_inputs_unavailable();
            unified.operation_failure(
                "file_sort",
                "FILE_SORT_FAILED",
                "取消文件排序前缀处理失败",
                &error,
            );
            let _ = unified.finish(true, false, false, &[], None, None, None);
            return Err(error);
        }
    };
    unified.set_discovered(result.rows.len());
    for (index, row) in result.rows.iter().enumerate() {
        let status = match row.status.as_str() {
            "失败" => pytools_core::run_logging::FileStatus::Failed,
            "未命中" | "跳过：非文件" | "跳过：无匹配排序前缀" => {
                pytools_core::run_logging::FileStatus::Skipped
            }
            _ => pytools_core::run_logging::FileStatus::Succeeded,
        };
        unified.file_result(
            index + 1,
            status,
            !row.status.contains("非文件"),
            "file_rename",
            &row.status,
            (status == pytools_core::run_logging::FileStatus::Failed)
                .then_some("FILE_RENAME_FAILED"),
            None,
        );
    }
    let (unified_summary, unified_path) = unified.finish(
        false,
        false,
        result.renamed == 0 && result.failed == 0,
        std::slice::from_ref(&result.output),
        Some(pytools_core::run_logging::OutputStats {
            workbooks_saved: Some(1),
            sheets_saved: None,
            saved: Some(true),
        }),
        None,
        None,
    );
    let legacy_log_result = if log_view == "developer" {
        write_file_sort_removal_legacy_log(&log_dir, &result.rows)
    } else {
        Ok(None)
    };
    if let Err(error) = &legacy_log_result {
        eprintln!("pytools: warning: could not write complete operation log: {error:#}");
    }
    if interactive && log_view == "user" {
        println!("{}", pytools_core::run_logging::terminal_line(if result.failed > 0 { pytools_core::run_logging::Level::Warn } else { pytools_core::run_logging::Level::Info }, &format!("文件排序前缀取消结束：重命名 {} 个，跳过 {} 个，失败 {} 个；报告 {}；运行编号 {}{}", result.renamed, result.skipped, result.failed, display_path_for_user(&result.output), unified_summary.run_id, unified_path.as_ref().map(|path| format!("；运行记录 {}", display_path_for_user(&path))).unwrap_or_default())));
        return Ok(());
    }
    let mut output = serde_json::to_value(result)?;
    output["run_id"] = serde_json::json!(unified_summary.run_id.clone());
    if let Some(path) = unified_path.as_ref() {
        output["run_summary_path"] = serde_json::json!(path);
    }
    if let Ok(Some(log_path)) = &legacy_log_result {
        output["log_path"] = serde_json::json!(log_path);
    }
    if let Ok(Some(history_path)) = &history_result {
        output["run_history_path"] = serde_json::json!(history_path);
    }
    if let Err(error) = &legacy_log_result {
        output["log_warning"] =
            serde_json::json!(format!("could not write complete operation log: {error:#}"));
    }
    if let Err(error) = &history_result {
        output["run_history_warning"] =
            serde_json::json!(format!("could not write structured run history: {error:#}"));
    }
    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}

fn run_extract_sheets(arguments: Vec<String>, interactive: bool) -> Result<()> {
    let mut arguments = arguments.into_iter();
    let mut config_path = None::<PathBuf>;
    let mut output_dir = None::<PathBuf>;
    let mut log_dir = None::<PathBuf>;
    let mut log_view = None::<String>;
    let mut append_existing = false;
    let mut source_workbooks = Vec::new();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--config" => {
                config_path = Some(PathBuf::from(
                    arguments.next().context("--config requires a path")?,
                ));
            }
            "--output-dir" => {
                output_dir = Some(PathBuf::from(
                    arguments.next().context("--output-dir requires a path")?,
                ));
            }
            "--log-dir" => {
                log_dir = Some(PathBuf::from(
                    arguments.next().context("--log-dir requires a path")?,
                ));
            }
            "--log-view" => {
                log_view = Some(arguments.next().context("--log-view requires a value")?);
            }
            // Accepted for compatibility with older invocations. PT-7-6 now
            // selects the bundled openpyxl worker internally.
            "--capabilities" => {
                let _ = arguments.next().context("--capabilities requires a path")?;
            }
            "--platform" => {
                let _ = arguments.next().context("--platform requires an id")?;
            }
            "--append-existing" => append_existing = true,
            "--help" | "-h" => {
                print_usage();
                return Ok(());
            }
            value if value.starts_with('-') => bail!("unknown option {value:?}"),
            path => source_workbooks.push(PathBuf::from(path)),
        }
    }

    let config_path = config_path.context("--config is required")?;
    let log_view_explicit = log_view.is_some();
    if log_view.is_none() {
        log_view = Some("user".to_owned());
    }
    if let Some(log_view) = log_view.as_deref() {
        anyhow::ensure!(
            matches!(log_view, "user" | "developer" | "off"),
            "--log-view must be user, developer, or off"
        );
    }
    if source_workbooks.is_empty() {
        bail!("provide at least one source workbook path");
    }
    let tasks = load_extract_tasks(&config_path).with_context(|| {
        format!(
            "load worksheet-extraction tasks from {}",
            display_path_for_user(&config_path)
        )
    })?;
    // 仅在显式给出 --log-view/--log-dir 或缺少 --output-dir 时才从配置工作簿读取全局设置；
    // 夹具/极简配置可能没有"全局配置"表。
    let global = if output_dir.is_none() || (log_view_explicit && log_dir.is_none()) {
        Some(load_global(&config_path).with_context(|| {
            format!(
                "load output and log directories from {}",
                display_path_for_user(&config_path)
            )
        })?)
    } else {
        None
    };
    let output_dir = output_dir
        .or_else(|| global.as_ref().map(|config| config.output_dir.clone()))
        .context("output directory is unavailable")?;
    let log_dir = log_dir
        .or_else(|| global.map(|config| config.log_dir))
        .unwrap_or_else(pytools_core::run_logging::default_log_dir);
    let run_started = Instant::now();
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_nanos())
        .unwrap_or_default();
    let run_id = format!(
        "{}_{}_{}",
        Local::now().format("%Y%m%dT%H%M%S"),
        std::process::id(),
        nonce
    );
    use pytools_core::run_logging::{EventWriter, Level, RunTracker};
    let view = log_view.as_deref().unwrap_or("user");
    if view == "user" {
        eprintln!(
            "{}",
            pytools_core::run_logging::terminal_line(
                pytools_core::run_logging::Level::Info,
                &format!("已启用工作表提取配置：{} 项", tasks.len())
            )
        );
        for task in &tasks {
            let matching_workbooks = source_workbooks
                .iter()
                .filter(|path| {
                    let name = path.file_name().unwrap_or_default().to_string_lossy();
                    pytools_core::sheet_extract::workbook_matches(&name, task)
                })
                .count();
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    pytools_core::run_logging::Level::Info,
                    &format!(
                        "正在检查配置：工作表「工作表提取」第 {} 行；任务「工作表提取第 {} 行」；工作簿关键字：{}；工作表关键字：{}；匹配源文件 {} 个；提取区域：{}；输出文件名：{}",
                        task.row_no,
                        task.row_no,
                        if task.workbook_keywords.trim().is_empty() { "（未限制）" } else { task.workbook_keywords.as_str() },
                        if task.sheet_keywords.trim().is_empty() { "（未限制）" } else { task.sheet_keywords.as_str() },
                        matching_workbooks,
                        if task.area_spec.trim().is_empty() { "（按整表规则）" } else { task.area_spec.as_str() },
                        if task.output_name.trim().is_empty() { "（自动生成）" } else { task.output_name.as_str() }
                    )
                )
            );
        }
    }
    let mut tracker = RunTracker::with_id(
        "PT-7-6",
        "工作表数据提取",
        &source_workbooks,
        run_id.clone(),
    );
    let mut diagnostic = if view == "off" {
        None
    } else {
        match EventWriter::create(&log_dir.join("run-events"), "PT-7-6", &run_id) {
            Ok(writer) => Some(writer),
            Err(error) => {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("无法创建诊断日志：{error}；业务处理继续")
                    )
                );
                None
            }
        }
    };
    if let Some(writer) = diagnostic.as_mut() {
        let event = tracker.event(
            Level::Info,
            "run_started",
            "run",
            "started",
            "开始工作表提取",
        );
        if let Err(error) = writer.write(&event) {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Warn,
                    &format!("诊断日志写入失败：{error}")
                )
            );
        }
    }
    let mut events = Vec::<SummaryRunEvent>::new();
    let mut sheet_stats = pytools_core::run_logging::SheetStats::default();
    let mut log_write_warning_emitted = false;
    let result = run_extract_sheets_candidate_with_progress(
        &source_workbooks,
        &tasks,
        &output_dir,
        append_existing,
        None,
        |event| {
            if view == "off" {
                events.push(event.clone());
            } else {
                emit_extract_sheets_event(&mut events, view, event.clone());
            }
            record_summary_progress_event(
                &mut tracker,
                diagnostic.as_mut(),
                &event,
                &source_workbooks,
                "source_extract",
                &mut sheet_stats,
                &mut log_write_warning_emitted,
            );
        },
    );

    let result = match result {
        Ok(result) => result,
        Err(error) => {
            let view = log_view.as_deref().unwrap_or("user");
            if !events.last().is_some_and(|event| event.status == "failed") {
                let prior = events.last();
                let mut progress = prior
                    .map(|event| event.progress.clone())
                    .unwrap_or_default();
                let prior_phase = prior.map(|event| event.phase.as_str()).unwrap_or("run");
                let output_failure = prior_phase.contains("output")
                    || prior_phase.contains("publish")
                    || prior_phase.contains("commit");
                if output_failure {
                    progress.current_input_index = None;
                    progress.current_sheet = None;
                    progress.current_config_row = None;
                }
                let mut event = SummaryRunEvent {
                    sequence: events.len() + 1,
                    timestamp: Local::now().to_rfc3339(),
                    phase: if output_failure {
                        "output_write".to_owned()
                    } else {
                        "source_extract".to_owned()
                    },
                    status: "failed".to_owned(),
                    user_message: "工作表提取失败".to_owned(),
                    progress,
                    file_name: if output_failure {
                        None
                    } else {
                        prior.and_then(|event| event.file_name.clone())
                    },
                    stage_elapsed_ms: Some(run_started.elapsed().as_millis()),
                    total_elapsed_ms: run_started.elapsed().as_millis(),
                    components: vec!["pytools-core".to_owned()],
                    actual_engine: None,
                    error_type: Some(
                        if output_failure {
                            "output_write_error"
                        } else {
                            "extraction_failed"
                        }
                        .to_owned(),
                    ),
                };
                if !output_failure && event.progress.current_input_index.is_none() {
                    event.progress.current_input_index =
                        prior.and_then(|prior| prior.progress.current_input_index);
                }
                record_summary_progress_event(
                    &mut tracker,
                    diagnostic.as_mut(),
                    &event,
                    &source_workbooks,
                    "source_extract",
                    &mut sheet_stats,
                    &mut log_write_warning_emitted,
                );
                emit_extract_sheets_event(&mut events, view, event);
            }
            let output_failure = events
                .iter()
                .rev()
                .find(|event| event.status == "failed")
                .is_some_and(|event| event.phase.contains("output"));
            let failed_event = events.iter().rev().find(|event| event.status == "failed");
            let mut summary = tracker.finish(
                output_failure,
                false,
                false,
                Vec::new(),
                diagnostic_path_for_log_view(diagnostic.as_ref(), view),
            );
            summary.input_stats_available = summary.input_stats.pending == 0;
            summary
                .business_metrics
                .insert("启用配置项数".to_owned(), serde_json::json!(tasks.len()));
            summary.business_metrics.insert(
                "失败工作簿数".to_owned(),
                serde_json::json!(summary.input_stats.failed),
            );
            if sheet_stats.discovered > 0 || sheet_stats.attempted > 0 {
                summary.sheet_stats = Some(sheet_stats.clone());
            }
            summary.output_stats = Some(pytools_core::run_logging::OutputStats {
                workbooks_saved: if output_failure { None } else { Some(0) },
                sheets_saved: if output_failure { None } else { Some(0) },
                saved: if output_failure { None } else { Some(false) },
            });
            if let Some(event) = failed_event {
                summary.error_summary = Some(pytools_core::run_logging::ErrorSummary {
                    stage: event.phase.clone(),
                    code: summary_event_error_code(event.error_type.as_deref(), &event.phase),
                    message: "工作表提取未完成；请检查关联文件、工作表和配置行".to_owned(),
                    file_name: event.file_name.clone(),
                    sheet_name: event.progress.current_sheet.clone(),
                });
            }
            if let Some(writer) = diagnostic.as_mut() {
                let stage = failed_event
                    .map(|event| event.phase.as_str())
                    .unwrap_or("source_extract");
                let mut failure = tracker.event(
                    Level::Error,
                    "operation_failed",
                    stage,
                    "failed",
                    "工作表提取未完成；请检查关联文件、工作表和配置行",
                );
                failure.error_code = Some(summary_event_error_code(
                    failed_event.and_then(|event| event.error_type.as_deref()),
                    stage,
                ));
                failure.file_name = failed_event.and_then(|event| event.file_name.clone());
                failure.sheet_name =
                    failed_event.and_then(|event| event.progress.current_sheet.clone());
                if view == "developer" {
                    failure.details =
                        Some(serde_json::json!({"error_chain": format!("{error:#}")}));
                }
                if let Err(write_error) = writer.write(&failure) {
                    if !log_write_warning_emitted {
                        eprintln!(
                            "{}",
                            pytools_core::run_logging::terminal_line(
                                Level::Warn,
                                &format!("诊断日志写入失败：{write_error}")
                            )
                        );
                        log_write_warning_emitted = true;
                    }
                }
            }
            let unified_path =
                save_unified_run_summary(&tracker, &mut diagnostic, &mut summary, &log_dir, view);
            if view != "off" {
                if let Err(history_error) = save_extract_sheets_run_history(
                    &log_dir,
                    &run_id,
                    "failed",
                    tasks.len(),
                    run_started.elapsed().as_millis(),
                    None,
                    Some("extraction_failed"),
                    &events,
                ) {
                    eprintln!(
                        "{}",
                        pytools_core::run_logging::terminal_line(
                            Level::Warn,
                            &format!("无法保存工作表提取记录：{history_error}")
                        )
                    );
                }
                if interactive && view == "user" {
                    // 用户视图的定位指针已由分组摘要的「详细日志」行承载；
                    // run_id 与统一记录属于开发者视图与机器接口。
                    emit_unified_terminal_summary_with_outputs(&summary, &[]);
                }
            }
            return Err(error);
        }
    };

    let status = if result.failed_files > 0 {
        "partial"
    } else {
        "succeeded"
    };
    let view = log_view.as_deref().unwrap_or("user");
    let output_paths = result
        .paths
        .iter()
        .map(|path| path_for_log_view(Path::new(path), view))
        .collect::<Vec<_>>();
    let no_op =
        result.output_files == 0 && result.extracted_sheets == 0 && result.failed_files == 0;
    let mut unified_summary = tracker.finish(
        false,
        false,
        no_op,
        output_paths,
        diagnostic_path_for_log_view(diagnostic.as_ref(), view),
    );
    unified_summary.input_stats_available = unified_summary.input_stats.pending == 0;
    if sheet_stats.discovered > 0 || sheet_stats.attempted > 0 {
        unified_summary.sheet_stats = Some(sheet_stats.clone());
    }
    unified_summary.output_stats = Some(pytools_core::run_logging::OutputStats {
        workbooks_saved: Some(result.output_files),
        sheets_saved: Some(result.extracted_sheets),
        saved: Some(result.output_files > 0),
    });
    unified_summary
        .business_metrics
        .insert("启用配置项数".to_owned(), serde_json::json!(tasks.len()));
    unified_summary.business_metrics.insert(
        "输出工作簿数".to_owned(),
        serde_json::json!(result.output_files),
    );
    unified_summary.business_metrics.insert(
        "提取工作表数".to_owned(),
        serde_json::json!(result.extracted_sheets),
    );
    unified_summary.business_metrics.insert(
        "失败工作簿数".to_owned(),
        serde_json::json!(result.failed_files),
    );
    if result.failed_files > 0 {
        if let Some(event) = events.iter().rev().find(|event| event.status == "failed") {
            unified_summary.error_summary = Some(pytools_core::run_logging::ErrorSummary {
                stage: event.phase.clone(),
                code: summary_event_error_code(event.error_type.as_deref(), &event.phase),
                message: "部分源工作簿未能提取；请检查对应文件和任务配置".to_owned(),
                file_name: event.file_name.clone(),
                sheet_name: event.progress.current_sheet.clone(),
            });
        }
    }
    let unified_path = save_unified_run_summary(
        &tracker,
        &mut diagnostic,
        &mut unified_summary,
        &log_dir,
        view,
    );
    if view != "off" {
        match save_extract_sheets_run_history(
            &log_dir,
            &run_id,
            status,
            tasks.len(),
            run_started.elapsed().as_millis(),
            Some(&result),
            None,
            &events,
        ) {
            Ok(path) if view == "user" => {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Info,
                        &format!(
                            "工作表提取运行记录已保存：{}",
                            pytools_core::run_logging::path_for_log_view(&path, false)
                        )
                    )
                )
            }
            Ok(path) => eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Info,
                    &format!("PT-7-6 run history: {}", display_path_for_user(&path))
                )
            ),
            Err(error) => {
                eprintln!(
                    "{}",
                    pytools_core::run_logging::terminal_line(
                        Level::Warn,
                        &format!("无法保存工作表提取记录：{error}")
                    )
                )
            }
        }
    }

    if interactive && view == "user" {
        // 用户视图的定位指针已由分组摘要的「详细日志」行承载。
        emit_unified_terminal_summary_with_outputs(
            &unified_summary,
            &result.paths.iter().map(PathBuf::from).collect::<Vec<_>>(),
        );
    } else {
        let mut output = serde_json::to_value(&result)?;
        if view != "off" {
            output["run_id"] = serde_json::json!(tracker.run_id);
            if let Some(path) = unified_path {
                output["run_summary_path"] = serde_json::json!(path);
            }
        }
        println!("{}", serde_json::to_string_pretty(&output)?);
    }
    Ok(())
}

fn emit_extract_sheets_event(
    events: &mut Vec<SummaryRunEvent>,
    log_view: &str,
    mut event: SummaryRunEvent,
) {
    if event.status == "failed" {
        let workbook = event.file_name.as_deref().map(Path::new);
        event.user_message = pytools_core::run_logging::workbook_sheet_context_message(
            &format!("工作表提取阶段 {} 失败", event.phase),
            workbook,
            event.progress.current_sheet.as_deref(),
        );
    }
    match log_view {
        "developer" => match serde_json::to_string(&event) {
            Ok(encoded) => eprintln!("{encoded}"),
            Err(_) => {
                eprintln!("{{\"phase\":\"progress_event\",\"status\":\"serialization_failed\"}}")
            }
        },
        "user" => {
            use pytools_core::run_logging::{terminal_line, Level};
            let level = match event.status.as_str() {
                "failed" => Level::Error,
                "skipped" | "warning" | "partial" => Level::Warn,
                _ => Level::Info,
            };
            let message = event.user_message.clone();
            eprintln!("{}", terminal_line(level, &message));
        }
        _ => {}
    }
    events.push(event);
}

/// Bridge existing summary progress into the shared event model without
/// changing the legacy progress payload or inferring unavailable file results.
fn record_summary_progress_event(
    tracker: &mut pytools_core::run_logging::RunTracker,
    writer: Option<&mut pytools_core::run_logging::EventWriter>,
    event: &SummaryRunEvent,
    inputs: &[PathBuf],
    terminal_success_phase: &str,
    sheet_stats: &mut pytools_core::run_logging::SheetStats,
    log_write_warning_emitted: &mut bool,
) {
    use pytools_core::run_logging::{FileStatus, Level};

    tracker.set_discovered(event.progress.total_files);
    sheet_stats.discovered = sheet_stats.discovered.max(event.progress.sheets_discovered);
    sheet_stats.attempted = sheet_stats.attempted.max(event.progress.sheets_attempted);
    sheet_stats.succeeded = sheet_stats.succeeded.max(event.progress.sheets_succeeded);
    sheet_stats.failed = sheet_stats.failed.max(event.progress.sheets_failed);
    sheet_stats.skipped = sheet_stats.skipped.max(event.progress.sheets_skipped);

    let input_scoped_phase = matches!(
        event.phase.as_str(),
        "source_read"
            | "input_filter"
            | "source_structure"
            | "source_extract"
            | "row_extract"
            | "extent_scan"
            | "value_read"
            | "sheet_process"
            | "file_process"
    );
    if input_scoped_phase {
        if let Some(index) = event.progress.current_input_index {
            let current = tracker.state(index);
            let next = match event.status.as_str() {
                "started" | "running" if current == Some(FileStatus::Pending) => {
                    Some(FileStatus::InProgress)
                }
                "failed"
                    if matches!(current, Some(FileStatus::Pending | FileStatus::InProgress)) =>
                {
                    Some(FileStatus::Failed)
                }
                "skipped"
                    if matches!(current, Some(FileStatus::Pending | FileStatus::InProgress)) =>
                {
                    Some(FileStatus::Skipped)
                }
                "succeeded"
                    if event.phase == terminal_success_phase
                        && matches!(
                            current,
                            Some(FileStatus::Pending | FileStatus::InProgress)
                        ) =>
                {
                    Some(FileStatus::Succeeded)
                }
                _ => None,
            };
            if let Some(next) = next {
                let _ = tracker.transition(index, next);
            }
        }
    }

    let Some(writer) = writer else { return };
    let level = match event.status.as_str() {
        "failed" => Level::Error,
        "skipped" | "warning" | "partial" => Level::Warn,
        _ => Level::Info,
    };
    let message = if event.status == "failed" {
        format!("{} 阶段失败", event.phase)
    } else if event.status == "skipped" {
        format!("{} 阶段跳过", event.phase)
    } else {
        event.user_message.clone()
    };
    let mut record = tracker.event(level, &event.phase, &event.phase, &event.status, message);
    record.timestamp = event.timestamp.clone();
    record.file_index = event.progress.current_input_index;
    record.file_name = event.file_name.clone().or_else(|| {
        event
            .progress
            .current_input_index
            .and_then(|index| index.checked_sub(1))
            .and_then(|index| inputs.get(index))
            .map(|path| pytools_core::run_logging::portable_file_name(path))
    });
    record.sheet_name = event.progress.current_sheet.clone();
    record.config_row = event.progress.current_config_row;
    record.engine = event.actual_engine.clone();
    record.elapsed_ms = event.stage_elapsed_ms;
    if event.status == "failed" {
        record.error_code = Some(summary_event_error_code(
            event.error_type.as_deref(),
            &event.phase,
        ));
    }
    if let Err(error) = writer.write(&record) {
        if !*log_write_warning_emitted {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Warn,
                    &format!("诊断日志写入失败：{error}；业务处理继续"),
                )
            );
            *log_write_warning_emitted = true;
        }
    }
}

/// 把命名区域汇总（PT-6-2）的错误链分类为具体错误类型与用户可读根因。
/// 错误链由 anyhow 按“外层上下文在前、根因在后”打印；根因即最末一行，
/// 承载真正的业务失败点（如“模板缺少表头区域_N”“缺少汇总区域”等）。
/// 返回 (error_type, 用户可读根因说明)。
fn classify_named_range_failure(error_chain: &str) -> (&'static str, String) {
    let root = error_chain
        .lines()
        .map(str::trim)
        .rev()
        .find(|line| !line.is_empty())
        .unwrap_or_default();
    let matched = |needle: &str| root.contains(needle) || error_chain.contains(needle);
    // 按业务检查点从细到粗匹配；root 优先，避免被外层上下文抢走。
    if root.contains("has no 表头区域") {
        return ("missing_header_region", root.to_owned());
    }
    if root.contains("配置了汇总区域，但缺少表头区域") {
        return ("missing_header_region", root.to_owned());
    }
    if root.contains("needs 固定行汇总区域") {
        return ("missing_summary_region", root.to_owned());
    }
    if root.contains("没有相交的表头区域") {
        return (
            "summary_region_without_header_intersection",
            root.to_owned(),
        );
    }
    if matched("模板中没有可用命名区域") {
        return (
            "template_has_no_named_regions",
            "模板中没有可用命名区域；请配置 表头区域_N 与 固定行汇总区域_N 或 任意行汇总区域_N"
                .to_owned(),
        );
    }
    if root.contains("模板工作表") && root.contains("不存在") {
        return ("template_sheet_missing", root.to_owned());
    }
    if matched("named-range summary template must be") {
        return (
            "template_format_error",
            "命名区域模板只支持 XLSX/XLSM".to_owned(),
        );
    }
    if matched("没有可汇总的源工作簿") {
        return (
            "no_summarizable_source",
            "没有可汇总的源工作簿；请检查输入文件、工作表名称路由和表结构区域配置".to_owned(),
        );
    }
    if root.is_empty() {
        (
            "summary_error",
            "汇总失败；请查看诊断日志了解失败阶段".to_owned(),
        )
    } else {
        ("summary_error", root.to_owned())
    }
}

fn summary_event_error_code(error_type: Option<&str>, phase: &str) -> String {
    let value = error_type.unwrap_or(phase).to_ascii_lowercase();
    if value.contains("template_workbook_read_error")
        || value.contains("template_defined_names_read_error")
        || value.contains("named_summary_region_read_error")
        || value.contains("named_cell_field_read_error")
        || value.contains("source_read_error")
    {
        "EXCEL_READ_FAILED".to_owned()
    } else if value.contains("config") || value.contains("template") || value.contains("rule") {
        "CONFIG_INVALID".to_owned()
    } else if value.contains("read") || phase.contains("read") || phase.contains("extent") {
        "EXCEL_READ_FAILED".to_owned()
    } else if value.contains("write") || value.contains("output") || phase.contains("output") {
        "OUTPUT_WRITE_FAILED".to_owned()
    } else if value.contains("sheet") || value.contains("structure") {
        "SHEET_OPERATION_FAILED".to_owned()
    } else {
        "INTERNAL_OPERATION_FAILED".to_owned()
    }
}

fn save_unified_run_summary(
    tracker: &pytools_core::run_logging::RunTracker,
    writer: &mut Option<pytools_core::run_logging::EventWriter>,
    summary: &mut pytools_core::run_logging::RunSummary,
    log_dir: &Path,
    log_view: &str,
) -> Option<PathBuf> {
    use pytools_core::run_logging::{Level, TaskStatus};
    if log_view == "off" {
        return None;
    }
    let diagnostic_path = writer.as_ref().map(|event_writer| {
        pytools_core::run_logging::diagnostic_path_for_log_view(
            event_writer.path(),
            log_view == "developer",
        )
    });
    summary.diagnostic_log_path = diagnostic_path;
    if let Some(writer) = writer.as_mut() {
        let status = format!("{:?}", summary.overall_status);
        let level = match summary.overall_status {
            TaskStatus::Failed => Level::Error,
            TaskStatus::PartialSuccess | TaskStatus::Cancelled => Level::Warn,
            TaskStatus::Success | TaskStatus::NoOp => Level::Info,
        };
        let event = tracker.event(
            level,
            "run_finished",
            "run",
            &status,
            format!("{}结束", tracker.feature_name),
        );
        if let Err(error) = writer.write(&event) {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Warn,
                    &format!("诊断日志写入失败：{error}"),
                )
            );
        }
    }
    let path = log_dir
        .join("run-unified")
        .join(format!("{}_{}.json", tracker.feature_id, tracker.run_id));
    match write_json_report_create_new(&path, summary) {
        Ok(()) => Some(path),
        Err(error) => {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Warn,
                    &format!("统一运行记录写入失败：{error}；业务结果仍保留"),
                )
            );
            None
        }
    }
}

fn save_extract_sheets_run_history(
    log_dir: &Path,
    run_id: &str,
    status: &str,
    task_count: usize,
    total_elapsed_ms: u128,
    result: Option<&pytools_core::sheet_extract::ExtractSheetsAreaResult>,
    error_type: Option<&str>,
    events: &[SummaryRunEvent],
) -> Result<PathBuf> {
    fs::create_dir_all(log_dir).with_context(|| {
        format!(
            "create worksheet extraction log directory {}",
            display_path_for_user(&log_dir)
        )
    })?;
    let actual_engines = events
        .iter()
        .filter_map(|event| event.actual_engine.clone())
        .collect::<std::collections::BTreeSet<_>>()
        .into_iter()
        .collect::<Vec<_>>();
    let progress = events
        .last()
        .map(|event| event.progress.clone())
        .unwrap_or_default();
    let result_summary = result
        .map(|result| {
            serde_json::json!({
                "files": result.files,
                "extracted_sheets": result.extracted_sheets,
                "output_files": result.output_files,
                "failed_files": result.failed_files,
                "output_file_names": result.paths.iter().filter_map(|path| {
                    Path::new(path).file_name().map(|name| name.to_string_lossy().into_owned())
                }).collect::<Vec<_>>(),
            })
        })
        .unwrap_or_else(|| {
            serde_json::json!({
                "files": progress.completed_files.saturating_sub(progress.failed_files),
                "extracted_sheets": progress.processed_sheets,
                "output_files": 0,
                "failed_files": progress.failed_files,
            })
        });
    let history = serde_json::json!({
        "schema_version": 1,
        "product": "pytools",
        "feature_id": "PT-7-6",
        "run_id": run_id,
        "status": status,
        "task_count": task_count,
        "total_inputs": progress.total_inputs,
        "processed_files": progress.processed_files,
        "total_elapsed_ms": total_elapsed_ms,
        "actual_engines": actual_engines,
        "result": result_summary,
        "error_type": error_type,
        "events": events,
    });
    let path = log_dir.join(format!("PT-7-6_{run_id}.run.json"));
    fs::write(&path, serde_json::to_vec_pretty(&history)?).with_context(|| {
        format!(
            "write worksheet extraction run history {}",
            display_path_for_user(&path)
        )
    })?;
    Ok(path)
}

fn run_region_sum(arguments: Vec<String>, interactive: bool) -> Result<()> {
    let mut arguments = arguments.into_iter();
    let mut config_path = None::<PathBuf>;
    let mut output_dir = None::<PathBuf>;
    let mut log_dir = None::<PathBuf>;
    let mut log_view = None::<String>;
    let mut source_workbooks = Vec::new();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--config" => {
                config_path = Some(PathBuf::from(
                    arguments.next().context("--config requires a path")?,
                ));
            }
            "--output-dir" => {
                output_dir = Some(PathBuf::from(
                    arguments.next().context("--output-dir requires a path")?,
                ));
            }
            "--log-dir" => {
                log_dir = Some(PathBuf::from(
                    arguments.next().context("--log-dir requires a path")?,
                ));
            }
            "--log-view" => {
                log_view = Some(arguments.next().context("--log-view requires a value")?);
            }
            "--help" | "-h" => {
                print_usage();
                return Ok(());
            }
            value if value.starts_with('-') => bail!("unknown option {value:?}"),
            path => source_workbooks.push(PathBuf::from(path)),
        }
    }

    let view = log_view.as_deref().unwrap_or("user");
    anyhow::ensure!(
        matches!(view, "user" | "developer" | "off"),
        "--log-view must be user, developer, or off"
    );
    if source_workbooks.is_empty() {
        bail!("provide at least one source workbook path");
    }

    let global = match config_path.as_ref() {
        Some(path) => Some(load_global(path).with_context(|| {
            format!(
                "load global configuration from {}",
                display_path_for_user(&path)
            )
        })?),
        None => None,
    };
    let output_dir = output_dir
        .or_else(|| global.as_ref().map(|config| config.output_dir.clone()))
        .unwrap_or_else(|| PathBuf::from("output"));
    let log_dir = log_dir
        .or_else(|| global.as_ref().map(|config| config.log_dir.clone()))
        .unwrap_or_else(pytools_core::run_logging::default_log_dir);

    use pytools_core::run_logging::{EventWriter, Level, RunTracker};
    let mut events = Vec::new();
    let mut tracker = None::<RunTracker>;
    let mut diagnostic = None::<EventWriter>;
    let mut sheet_stats = pytools_core::run_logging::SheetStats::default();
    let mut log_write_warning_emitted = false;
    let source_labels = pytools_core::run_logging::disambiguated_file_names(&source_workbooks);
    let operation_result =
        run_region_sum_check_with_progress(&source_workbooks, &output_dir, &log_dir, |event| {
            if tracker.is_none() {
                tracker = Some(RunTracker::with_id(
                    "PT-7-5",
                    "地区总分校验",
                    &source_workbooks,
                    event.run_id.clone(),
                ));
                if view != "off" {
                    match EventWriter::create(&log_dir.join("run-events"), "PT-7-5", &event.run_id)
                    {
                        Ok(writer) => diagnostic = Some(writer),
                        Err(error) => eprintln!(
                            "{}",
                            pytools_core::run_logging::terminal_line(
                                Level::Warn,
                                &format!("无法创建诊断日志：{error}；业务处理继续")
                            )
                        ),
                    }
                }
            }
            let Some(tracker) = tracker.as_mut() else {
                return;
            };
            record_region_sum_progress_event(
                tracker,
                diagnostic.as_mut(),
                &event,
                &source_workbooks,
                &mut sheet_stats,
                &mut log_write_warning_emitted,
            );
            emit_region_sum_event(view, &event, &source_labels);
            events.push(event);
        });
    let tracker =
        tracker.unwrap_or_else(|| RunTracker::new("PT-7-5", "地区总分校验", &source_workbooks));
    if view != "off" {
        if let Err(error) = write_region_sum_run_history(&log_dir, &events, &operation_result) {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Warn,
                    &format!("无法保存地区总分校验记录：{error}")
                )
            );
        }
    }
    let failed_event = events.iter().rev().find(|event| event.status == "failed");
    let critical_failure = operation_result.is_err()
        && failed_event.is_some_and(|event| {
            matches!(
                event.phase.as_str(),
                "output_setup" | "report_write" | "log_setup"
            )
        });
    let paths = operation_result
        .as_ref()
        .ok()
        .and_then(|result| result.path.as_ref())
        .map(|path| vec![path_for_log_view(path, view)])
        .unwrap_or_default();
    let no_op = operation_result
        .as_ref()
        .is_ok_and(|result| result.sheets == 0 && result.errors == 0);
    let mut unified_summary = tracker.finish(
        critical_failure,
        false,
        no_op,
        paths,
        diagnostic_path_for_log_view(diagnostic.as_ref(), view),
    );
    unified_summary.input_stats_available = unified_summary.input_stats.pending == 0;
    if sheet_stats.discovered > 0 || sheet_stats.attempted > 0 {
        unified_summary.sheet_stats = Some(sheet_stats);
    }
    if let Ok(result) = &operation_result {
        unified_summary
            .business_metrics
            .insert("核对工作表数".to_owned(), serde_json::json!(result.sheets));
        unified_summary
            .business_metrics
            .insert("发现差异数".to_owned(), serde_json::json!(result.errors));
        unified_summary.business_metrics.insert(
            "失败工作簿数".to_owned(),
            serde_json::json!(result.failed_files),
        );
        unified_summary.output_stats = Some(pytools_core::run_logging::OutputStats {
            workbooks_saved: Some(if result.path.is_some() { 1 } else { 0 }),
            sheets_saved: None,
            saved: Some(result.path.is_some()),
        });
        if result.errors > 0 {
            unified_summary.business_issues.push(
                pytools_core::run_logging::BusinessIssue::new("地区总分校验", "校验发现不一致项")
                    .count(result.errors),
            );
        }
    } else {
        unified_summary.output_stats = Some(pytools_core::run_logging::OutputStats {
            workbooks_saved: None,
            sheets_saved: None,
            saved: None,
        });
    }
    if let Some(event) = failed_event {
        unified_summary.error_summary = Some(pytools_core::run_logging::ErrorSummary {
            stage: event.phase.clone(),
            code: summary_event_error_code(event.error_type.as_deref(), &event.phase),
            message: event.user_message.clone(),
            file_name: event
                .input_index
                .and_then(|index| index.checked_sub(1))
                .and_then(|index| source_labels.get(index))
                .cloned(),
            sheet_name: None,
        });
    }
    let unified_path = save_unified_run_summary(
        &tracker,
        &mut diagnostic,
        &mut unified_summary,
        &log_dir,
        view,
    );
    match operation_result {
        Ok(result) => {
            if interactive && view == "user" {
                let outputs = result
                    .path
                    .as_ref()
                    .map(std::slice::from_ref)
                    .unwrap_or(&[]);
                emit_unified_terminal_summary_with_outputs(&unified_summary, outputs);
                // 用户视图的定位指针已由分组摘要的「详细日志」行承载。
            } else {
                let path = result
                    .path
                    .as_ref()
                    .map(|path| path.to_string_lossy().into_owned());
                let mut output = serde_json::json!({
                    "files": result.files,
                    "sheets": result.sheets,
                    "errors": result.errors,
                    "failed_files": result.failed_files,
                    "path": path,
                    "log_path": result.log_path.to_string_lossy(),
                    "run_id": tracker.run_id,
                });
                if let Some(path) = unified_path {
                    output["run_summary_path"] = serde_json::json!(path);
                }
                println!("{}", serde_json::to_string_pretty(&output)?);
            }
            Ok(())
        }
        Err(error) => {
            if view == "user" {
                // 用户视图的定位指针已由分组摘要的「详细日志」行承载。
                emit_unified_terminal_summary_with_outputs(&unified_summary, &[]);
            }
            Err(error)
        }
    }
}

fn record_region_sum_progress_event(
    tracker: &mut pytools_core::run_logging::RunTracker,
    writer: Option<&mut pytools_core::run_logging::EventWriter>,
    event: &RegionSumRunEvent,
    inputs: &[PathBuf],
    sheet_stats: &mut pytools_core::run_logging::SheetStats,
    log_write_warning_emitted: &mut bool,
) {
    use pytools_core::run_logging::{FileStatus, Level};
    tracker.set_discovered(event.progress.total_files);
    sheet_stats.discovered = sheet_stats.discovered.max(event.progress.matched_sheets);
    sheet_stats.attempted = sheet_stats.attempted.max(event.progress.matched_sheets);
    sheet_stats.succeeded = sheet_stats.succeeded.max(event.progress.matched_sheets);

    if let Some(index) = event.input_index {
        let current = tracker.state(index);
        let next = match (event.phase.as_str(), event.status.as_str()) {
            ("workbook_read", "started") if current == Some(FileStatus::Pending) => {
                Some(FileStatus::InProgress)
            }
            ("input_validation" | "workbook_read", "failed")
                if matches!(current, Some(FileStatus::Pending | FileStatus::InProgress)) =>
            {
                Some(FileStatus::Failed)
            }
            ("region_check", "completed")
                if matches!(current, Some(FileStatus::Pending | FileStatus::InProgress)) =>
            {
                Some(FileStatus::Succeeded)
            }
            _ => None,
        };
        if let Some(next) = next {
            let _ = tracker.transition(index, next);
        }
    }

    let Some(writer) = writer else { return };
    let level = match event.status.as_str() {
        "failed" => Level::Error,
        "skipped" | "warning" => Level::Warn,
        _ => Level::Info,
    };
    let message = event.user_message.clone();
    let mut record = tracker.event(level, &event.phase, &event.phase, &event.status, message);
    record.timestamp = event.timestamp.clone();
    record.file_index = event.input_index;
    record.file_name = event
        .input_index
        .and_then(|index| index.checked_sub(1))
        .and_then(|index| inputs.get(index))
        .map(|path| pytools_core::run_logging::portable_file_name(path));
    record.sheet_name = event.worksheet_name.clone();
    record.engine = event.actual_engine.clone();
    record.elapsed_ms = event.stage_elapsed_ms;
    if event.status == "failed" {
        record.error_code = Some(summary_event_error_code(
            event.error_type.as_deref(),
            &event.phase,
        ));
    }
    if let Err(error) = writer.write(&record) {
        if !*log_write_warning_emitted {
            eprintln!(
                "{}",
                pytools_core::run_logging::terminal_line(
                    Level::Warn,
                    &format!("诊断日志写入失败：{error}；业务处理继续")
                )
            );
            *log_write_warning_emitted = true;
        }
    }
}

fn emit_region_sum_event(log_view: &str, event: &RegionSumRunEvent, file_labels: &[String]) {
    match log_view {
        "user" => {
            use pytools_core::run_logging::{BatchEventKind, BatchLogPolicy, Level};
            // 业务模块自带的「完成：成功 N 个…」与统一分组摘要完全重复；
            // 用户只保留一份最终摘要（目标 4.4 避免重复摘要）。
            if event.phase == "run" && event.status == "completed" {
                return;
            }
            let total = event.progress.total_files;
            let current = event.input_index;
            let mut message = event.user_message.clone();
            if let (Some(index), Some(original_name)) = (current, event.workbook_name.as_deref()) {
                if let Some(label) = index
                    .checked_sub(1)
                    .and_then(|index| file_labels.get(index))
                {
                    message = message.replace(original_name, label);
                }
            }
            if event.status == "failed" {
                message = format!("地区总分校验阶段 {} 失败：{message}", event.phase);
            }
            let phase_change = matches!(
                event.phase.as_str(),
                "run" | "output_setup" | "report_write" | "log_setup"
            );
            let policy = BatchLogPolicy::new(total).with_elapsed_ms(event.total_elapsed_ms);
            let kind = if event.status == "failed" {
                BatchEventKind::Failure
            } else if matches!(event.status.as_str(), "warning" | "skipped") {
                BatchEventKind::BusinessIssue
            } else if phase_change {
                BatchEventKind::PhaseChange
            } else {
                BatchEventKind::Routine
            };
            // Large batches show one checkpoint per group of files. Per-sheet
            // business findings and every failure remain visible.
            if total > 8 && kind == BatchEventKind::Routine && event.phase != "region_check" {
                return;
            }
            if policy.should_emit(current, kind) {
                emit_batch_event_view(
                    "user",
                    total,
                    current,
                    &event.status,
                    phase_change,
                    &message,
                    None,
                    Some(event.total_elapsed_ms),
                );
            }
        }
        "developer" => {
            let mut safe = event.clone();
            if safe.status == "failed" {
                safe.user_message = format!("地区总分校验阶段 {} 失败", safe.phase);
            }
            match serde_json::to_string(&safe) {
                Ok(encoded) => eprintln!("{encoded}"),
                Err(_) => eprintln!(
                    "{{\"feature_id\":\"PT-7-5\",\"phase\":\"{}\",\"status\":\"{}\"}}",
                    event.phase, event.status
                ),
            }
        }
        "off" => {}
        _ => unreachable!("region-sum log view was validated"),
    }
}

fn write_region_sum_run_history(
    log_dir: &Path,
    events: &[RegionSumRunEvent],
    operation_result: &Result<pytools_core::region_sum::RegionSumRun>,
) -> Result<PathBuf> {
    let first = events.first().context("PT-7-5 run emitted no events")?;
    let last = events.last().context("PT-7-5 run emitted no events")?;
    let mut actual_engines = Vec::<String>::new();
    for event in events {
        if let Some(engine) = &event.actual_engine {
            if !actual_engines.contains(engine) {
                actual_engines.push(engine.clone());
            }
        }
        for component in &event.components {
            if component != "Rust" && !actual_engines.contains(component) {
                actual_engines.push(component.clone());
            }
        }
    }
    let history = serde_json::json!({
        "schema_version": 1,
        "product": "pytools",
        "feature_id": "PT-7-5",
        "function": "region-sum-check",
        "run_id": &first.run_id,
        "started_at": &first.timestamp,
        "status": if operation_result.is_ok() { "completed" } else { "failed" },
        "total_files": first.progress.total_files,
        "processed_files": last.progress.processed_files,
        "completed_files": last.progress.completed_files,
        "failed_files": last.progress.failed_files,
        "matched_sheets": last.progress.matched_sheets,
        "errors": last.progress.errors,
        "total_elapsed_ms": last.total_elapsed_ms,
        "actual_engines": actual_engines,
        "failure_phase": events.iter().rev().find(|event| event.status == "failed").map(|event| event.phase.clone()),
        "failure_type": events.iter().rev().find(|event| event.status == "failed").and_then(|event| event.error_type.clone()),
        "events": events,
    });
    fs::create_dir_all(log_dir).with_context(|| {
        format!(
            "create PT-7-5 history directory {}",
            display_path_for_user(&log_dir)
        )
    })?;
    let path = log_dir.join(format!("PT-7-5_{}.run.json", first.run_id));
    fs::write(&path, serde_json::to_vec_pretty(&history)?)
        .with_context(|| format!("write PT-7-5 run history {}", display_path_for_user(&path)))?;
    Ok(path)
}

fn print_usage() {
    eprintln!("rustTools {}", build_banner().trim());
    eprintln!(
        "Usage:\n\
         pytools [menu] [--config <config.xlsx>] [--log-dir <directory>] [--log-view user|developer|off]\n\
         no command opens the rustTools menu; available categories: 1 (data extraction and summaries 1-5), 2 (configured deduplication and append 1-4), 3 (static print and precheck 1-2), 4 (file processing 1-8), 5 (special reports 1-8), 6 (archive preview and copy), and 7 (configuration and quick actions 1-5)\n\
         timeline-precheck --config <workbook.xlsx>\n\
         timeline-compare --config <workbook.xlsx> --template <workbook.xlsx> [--log-dir <dir>] [--log-view user|developer|off] <source.xlsx>...\n\
         timeline-extract --config <workbook.xlsx> [--output-dir <dir>] [--log-dir <dir>] [--log-view user|developer|off] [--capabilities <json> --platform <id> [--commit-target]] [<source.xlsx>...]\n\
         note: timeline-extract is a retained standalone candidate; PT-1-2 is excluded from the confirmed migration scope and interactive menu\n\
         wide-summary --config <workbook.xlsx> [<source.xlsx>...]\n\
         wide-summary-row-suffix --config <workbook.xlsx> [<source.xlsx>...]\n\
         wide-summary writes a summary workbook and the configured intermediate target; use menu 2-3 to append that intermediate to historical data\n\
         summary-used-range [--output-dir <dir>] [--log-dir <dir>] [--log-view user|developer|off] <workbook>...\n\
         summary-named-ranges --template <workbook.xlsx|xlsm> [--output-dir <dir>] [--log-dir <dir>] [--log-view user|developer|off] <source.xlsx|xlsm>...\n\
         summary-named-validation --template <workbook.xlsx|xlsm> [--output-dir <dir>] [--log-dir <dir>] [--log-view user|developer|off] <source.xls|xlsx|xlsm>...\n\
         summary-comments --template <workbook.xlsx> [--output-dir <dir>] [--log-dir <dir>] [--log-view user|developer|off] <source.xlsx>...\n\
         excel-open-repair [--recursive] [--log-dir <dir>] [--log-view user|developer|off] <file-or-directory>... (修复 Excel COM 无法打开的 .xlsx/.xlsm；只产出修复副本，原件不变)\n\
         dedup-comment-plan --input <workbook.xlsx> --output <candidate.xlsx> [--action check|mark|delete] [--plan-output <plan.json>] [--capabilities <json> --platform <id>]\n\
         dedup-config-plan --config <workbook.xlsx> [--output-dir <dir>] [--action check|mark|delete] [--capabilities <json> --platform <id>]\n\
         dedup-precheck --config <workbook.xlsx>\n\
         dedup-append-plan --config <workbook.xlsx> [--output-dir <dir>]\n\
         dedup-append-execute --plan <append-plan.json> [--capabilities <json> --platform <id>] [--commit-target]\n\
         print-comment-plan --input <workbook.xlsx> --output <candidate.xlsx> [--drop-comments true] [--hide-zero true] [--plan-output <plan.json>] [--capabilities <json> --platform <id>]\n\
         print-config-precheck --config <workbook.xlsx>\n\
         print-config-plan --config <workbook.xlsx> [--output-dir <plan-report-dir>] [--capabilities <json> --platform <id>]\n\
         apply-edit-plan --plan <plan.json> --capabilities <json> --platform <id>\n\
         config-init --path <new-config.xlsx>\n\
         config-repair-plan --input <config.xlsx> --output <candidate.xlsx> [--plan-output <plan.json>] [--capabilities <json> --platform <id>]\n\
         config-view --config <config.xlsx>\n\
         village-split-plan --current <xlsx> --previous <xlsx> --input <target.xlsx> --output <candidate.xlsx> [--plan-output <json>] [--capabilities <json> --platform <id>]\n\
         village-split --current <xls|xlsx> --previous <xls|xlsx> [--config <config.xlsx>] [--output-dir <directory>] [--log-dir <directory>] [--log-view user|developer|off] [--capabilities <json> --platform <id> [--commit-target]] <target.xls|xlsx>... (Windows .xls uses Excel/WPS COM; .xlsx uses openpyxl)\n\
         remove-foreign-bank-plan --config <config.xlsx> --input <xlsx> --output <candidate.xlsx> [--plan-output <json>] [--capabilities <json> --platform <id>]\n\
         remove-foreign-bank --config <config.xlsx> [--plan-dir <dir>] [--view user|developer|off] [--log-dir <dir>] [--report-output <json>] [--capabilities <json> --platform <id>] <source.xlsx>...\n\
         rural-loan-plan --current <xlsx> --previous <xlsx> --output <candidate.xlsx> [--plan-output <json>] [--capabilities <json> --platform <id>]\n\
         quarterly-rural-loan-plan --current <xlsx> --previous <xlsx> --output <candidate.xlsx> [--plan-output <json>] [--capabilities <json> --platform <id>]\n\
         institution-normalize-plan --config <config.xlsx> --input <xlsx> --output <candidate.xlsx> [--plan-output <json>] [--capabilities <json> --platform <id>]\n\
         institution-normalize --config <config.xlsx> --plan-dir <dir> [--view user|developer|off] [--report-output <json>] <workbook.xlsx>...\n\
         institution-normalize --config <config.xlsx> --apply [--capabilities <json> --platform <id>] [--view user|developer|off] [--report-output <json>] <workbook.xlsx>... (creates verified copies beside sources)\n\
         fx-header <xlsx>... [--output-dir <dir>] [--log-dir <dir>] [--plan-dir <dir>] [--view user|developer|off] [--commit-target --capabilities <json> --platform <id>]\n\
         fx-header-plan --input <xlsx> --output <candidate.xlsx> [--plan-output <json>] [--capabilities <json> --platform <id>]\n\
         rename-sheets-plan --config <config.xlsx> --input <xlsx> --output <candidate.xlsx> [--plan-output <json>] [--capabilities <json> --platform <id>]\n\
         rename-sheets --config <config.xlsx> [--recursive] <file-or-directory>... (.xlsx 使用 openpyxl；Windows .xls 使用专用 Excel/WPS COM 改名)\n\
         submission-check --config <config.xlsx> [--output-dir <dir>] [--log-dir <dir>] [--log-view user|developer|off] <file-or-directory>...\n\
         archive-date-preview --target-root <dir> [--config <config.xlsx>] [--output-dir <dir>] <file-or-directory>...\n\
         archive-type-preview --config <config.xlsx> --target-root <dir> [--output-dir <dir>] <file-or-directory>...\n\
         archive-date-copy --target-root <dir> [--config <config.xlsx>] [--output-dir <dir>] <file-or-directory>...\n\
         archive-type-copy --config <config.xlsx> --target-root <dir> [--output-dir <dir>] <file-or-directory>...\n\
         archive-execute-plan --plan <archive-plan.xlsx> [--output-dir <dir>]\n\
         convert-excel --config <config.xlsx> [--to xlsx] [--engine auto|com|pywin32] [--output-dir <dir>] [--log-dir <dir>] [--log-view user|developer|off] <file>...\n\
         convert-word --config <config.xlsx> [--to docx] [--engine auto|com|pywin32] [--output-dir <dir>] [--log-dir <dir>] [--log-view user|developer|off] <file>...\n\
         convert-file --input <file> --to <extension> [--output <file>] [--engine auto|com|pywin32|native]\n\
         Office 文件转换仅使用 Windows pywin32 + Microsoft Excel/WPS COM；Linux 不执行 LibreOffice 转换\n\
         native = 按值重建（C01，丢弃样式/公式缓存）；Excel 公式文本保留但不重算，CSV 按文本写入\n\
                   ISO 日期/时长、错误值、空文本和检测到的 1904 epoch XLS/XLSM/XLSB 会拒绝转换\n\
         pytools extract-sheets --config <workbook.xlsx> [--output-dir <dir>] [--log-dir <dir>] [--log-view user|developer|off] [--append-existing] <source.xlsx>...\n\
         pytools region-sum-check [--config <workbook.xlsx>] [--output-dir <dir>] [--log-dir <dir>] [--log-view user|developer|off] <source.xlsx>...\n\
         pytools file-sort-prefix --config <workbook.xlsx> [--output-dir <dir>] [--log-dir <dir>] [--log-view user|developer|off] <file>...\n\
         pytools remove-file-sort-prefix --config <workbook.xlsx> [--output-dir <dir>] [--log-dir <dir>] [--log-view user|developer|off] [<file>...]\n\
         pytools rename-by-content --config <workbook.xlsx> [--log-dir <dir>] [--log-view user|developer|off] <file>...\n\
         pytools unrename-by-content --config <workbook.xlsx> [--log-dir <dir>] [--log-view user|developer|off] <file>...\n\
         pytools rename-local-map --config <workbook.xlsx> [--log-dir <dir>] [--log-view user|developer|off] <file>...\n\
         pytools rename-batch --config <workbook.xlsx> [--log-dir <dir>] [--log-view user|developer|off] [<file-or-directory>...]\n\
         The region-sum command runs the Rust non-COM regional total check. Without --config,\n\
         output defaults to ./output. Logs default to the per-user Rust-Tools state directory\n\
         unless --log-dir or the global configuration specifies another path. File-sort-prefix\n\
         and remove-file-sort-prefix write reports under ./output unless --output-dir is supplied;\n\
         accepts an empty file list to produce the source-compatible blank report. rename-by-content reads .xlsx/.xlsm/.xls through Calamine; unrename-by-content accepts .xlsx/.xlsm. Legacy .xlt COM reads are not part of the Rust path."
    );
}

#[cfg(test)]
mod wide_summary_path_tests {
    use super::display_path_for_user;
    use std::path::Path;

    #[test]
    fn strips_windows_extended_prefix_from_user_facing_paths() {
        assert_eq!(
            display_path_for_user(Path::new(r"\\?\D:\PyTools\output\summary.xlsx")),
            r"D:\PyTools\output\summary.xlsx"
        );
        assert_eq!(
            display_path_for_user(Path::new(r"\\?\UNC\server\share\summary.xlsx")),
            r"\\server\share\summary.xlsx"
        );
        assert_eq!(
            display_path_for_user(Path::new(r"D:\PyTools\output\summary.xlsx")),
            r"D:\PyTools\output\summary.xlsx"
        );
    }
}

#[cfg(test)]
mod unified_run_session_tests {
    use super::UnifiedRunSession;
    use pytools_core::run_logging::{FileStatus, OutputStats, TaskStatus};
    use std::{
        fs,
        path::PathBuf,
        time::{SystemTime, UNIX_EPOCH},
    };

    fn temporary_directory(label: &str) -> PathBuf {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let path =
            std::env::temp_dir().join(format!("pytools-{label}-{}-{nonce}", std::process::id()));
        fs::create_dir_all(&path).unwrap();
        path
    }

    #[test]
    fn user_run_summary_correlates_file_results_and_hides_absolute_paths() {
        let root = temporary_directory("user-run-log");
        let mut session = UnifiedRunSession::new(&root, "PT-TEST", "日志适配器测试", "user", None);
        let inputs = [
            PathBuf::from(r"C:\业务\机构甲.xlsx"),
            PathBuf::from(r"C:\业务\机构乙.xlsx"),
        ];
        session.set_inputs(&inputs, Some(2));
        session.file_result(
            1,
            FileStatus::Succeeded,
            true,
            "read_and_write",
            "文件处理完成",
            None,
            None,
        );
        session.file_result(
            2,
            FileStatus::Skipped,
            true,
            "selection",
            "文件不符合筛选条件",
            None,
            None,
        );
        let output = PathBuf::from(r"C:\业务输出\汇总.xlsx");
        let (summary, summary_path) = session.finish(
            false,
            false,
            false,
            std::slice::from_ref(&output),
            Some(OutputStats {
                workbooks_saved: Some(1),
                sheets_saved: None,
                saved: Some(true),
            }),
            None,
            None,
        );
        let summary_path = summary_path.unwrap();
        let persisted: serde_json::Value =
            serde_json::from_slice(&fs::read(summary_path).unwrap()).unwrap();
        assert_eq!(summary.overall_status, TaskStatus::Success);
        assert_eq!(persisted["run_id"], summary.run_id);
        assert_eq!(persisted["input_stats"]["selected"], 2);
        assert_eq!(persisted["input_stats"]["attempted"], 2);
        assert_eq!(persisted["input_stats"]["succeeded"], 1);
        assert_eq!(persisted["input_stats"]["skipped"], 1);
        assert_eq!(persisted["output_paths"][0], "汇总.xlsx");
        assert!(persisted
            .get("diagnostic_log_path")
            .unwrap()
            .as_str()
            .unwrap()
            .ends_with(".jsonl"));
        let _ = fs::remove_dir_all(root);
    }

    #[test]
    fn developer_run_log_keeps_the_error_chain_and_failed_stage() {
        let root = temporary_directory("developer-run-log");
        let input = PathBuf::from(r"C:\业务\机构甲.xlsx");
        let mut session =
            UnifiedRunSession::new(&root, "PT-TEST", "日志适配器测试", "developer", None);
        session.set_inputs(std::slice::from_ref(&input), Some(1));
        let error = anyhow::anyhow!("reader failed: invalid workbook");
        session.operation_failure(
            "workbook_read",
            "EXCEL_READ_FAILED",
            "工作簿读取失败",
            &error,
        );
        let (summary, summary_path) = session.finish(true, false, false, &[], None, None, None);
        let summary_path = summary_path.unwrap();
        assert_eq!(summary.overall_status, TaskStatus::Failed);
        let event_path = root
            .join("run-events")
            .join(format!("PT-TEST_{}.jsonl", summary.run_id));
        let content = fs::read_to_string(event_path).unwrap();
        let events = content
            .lines()
            .map(|line| serde_json::from_str::<serde_json::Value>(line).unwrap())
            .collect::<Vec<_>>();
        let failed = events
            .iter()
            .find(|event| event["event_name"] == "operation_failed")
            .unwrap();
        assert_eq!(failed["stage"], "workbook_read");
        assert_eq!(
            failed["details"]["error_chain"],
            "reader failed: invalid workbook"
        );
        assert!(summary_path.exists());
        let _ = fs::remove_dir_all(root);
    }
}
