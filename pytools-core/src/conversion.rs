//! Explicit same-category conversion requests. Fidelity conversions run the
//! in-process engine (`conversion_engine`): Office/WPS COM on Windows, a
//! local LibreOffice `soffice` CLI elsewhere; no external sidecar executable.

use anyhow::{anyhow, bail, Context, Result};
use chrono::Local;
use encoding_rs::GBK;
use excel_reader::{read_workbook, CellValue};
use excel_writer::Workbook;
use serde::Serialize;
use std::{
    collections::HashSet,
    fs,
    io::Read,
    path::{Path, PathBuf},
    time::Instant,
};

#[derive(Clone, Debug, Serialize)]
pub struct ConversionItem {
    pub input: PathBuf,
    pub output: Option<PathBuf>,
    pub status: String,
    pub detail: String,
}
#[derive(Clone, Debug, Serialize)]
pub struct ConversionRun {
    pub converted: usize,
    pub skipped: usize,
    pub failed: usize,
    pub items: Vec<ConversionItem>,
}

#[derive(Clone, Debug, Serialize)]
pub struct ConversionProgressEvent {
    pub sequence: usize,
    pub timestamp: String,
    pub feature_id: String,
    pub function: String,
    pub phase: String,
    pub status: String,
    pub completed_files: usize,
    pub total_files: usize,
    pub converted_files: usize,
    pub failed_files: usize,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub input_index: Option<usize>,
    pub total_elapsed_ms: u128,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stage_elapsed_ms: Option<u128>,
    pub components: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub actual_engine: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_type: Option<String>,
    pub user_message: String,
}

pub fn convert_excel_files(
    inputs: &[PathBuf],
    output_dir: &Path,
    engine: &str,
) -> Result<ConversionRun> {
    convert_excel_files_to(inputs, output_dir, "xlsx", engine)
}
pub fn convert_excel_files_to(
    inputs: &[PathBuf],
    output_dir: &Path,
    target: &str,
    engine: &str,
) -> Result<ConversionRun> {
    convert_excel_files_with_progress(inputs, Some(output_dir), target, engine, |_| {})
}

pub fn convert_excel_files_with_progress<F>(
    inputs: &[PathBuf],
    output_dir: Option<&Path>,
    target: &str,
    engine: &str,
    on_event: F,
) -> Result<ConversionRun>
where
    F: FnMut(&ConversionProgressEvent),
{
    anyhow::ensure!(
        target
            .trim()
            .trim_start_matches('.')
            .eq_ignore_ascii_case("xlsx"),
        "Excel conversion currently targets .xlsx only"
    );
    convert_files(
        inputs,
        output_dir,
        target,
        engine,
        ConversionCategory::Excel,
        on_event,
    )
}
pub fn convert_excel_files_in_source_folders(
    inputs: &[PathBuf],
    target: &str,
    engine: &str,
) -> Result<ConversionRun> {
    convert_excel_files_with_progress(inputs, None, target, engine, |_| {})
}
pub fn convert_word_files(
    inputs: &[PathBuf],
    output_dir: &Path,
    engine: &str,
) -> Result<ConversionRun> {
    convert_word_files_to(inputs, output_dir, "docx", engine)
}
pub fn convert_word_files_to(
    inputs: &[PathBuf],
    output_dir: &Path,
    target: &str,
    engine: &str,
) -> Result<ConversionRun> {
    convert_word_files_with_progress(inputs, Some(output_dir), target, engine, |_| {})
}

pub fn convert_word_files_with_progress<F>(
    inputs: &[PathBuf],
    output_dir: Option<&Path>,
    target: &str,
    engine: &str,
    on_event: F,
) -> Result<ConversionRun>
where
    F: FnMut(&ConversionProgressEvent),
{
    anyhow::ensure!(
        target
            .trim()
            .trim_start_matches('.')
            .eq_ignore_ascii_case("docx"),
        "Word conversion currently targets .docx only"
    );
    convert_files(
        inputs,
        output_dir,
        target,
        engine,
        ConversionCategory::Word,
        on_event,
    )
}
pub fn convert_word_files_in_source_folders(
    inputs: &[PathBuf],
    target: &str,
    engine: &str,
) -> Result<ConversionRun> {
    convert_word_files_with_progress(inputs, None, target, engine, |_| {})
}
pub fn convert_one(input: &Path, output: &Path, target: &str, engine: &str) -> Result<String> {
    let mut runtime = crate::conversion_engine::ConversionRuntime::default();
    convert_one_with_runtime(input, output, target, engine, &mut runtime)
}

fn convert_one_with_runtime(
    input: &Path,
    output: &Path,
    target: &str,
    engine: &str,
    runtime: &mut crate::conversion_engine::ConversionRuntime,
) -> Result<String> {
    anyhow::ensure!(
        input.is_file(),
        "input file does not exist: {}",
        input.display()
    );
    anyhow::ensure!(
        !output.exists(),
        "refusing to overwrite existing file: {}",
        output.display()
    );
    let target = target.trim().trim_start_matches('.').to_ascii_lowercase();
    let source_ext = input
        .extension()
        .and_then(|ext| ext.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    anyhow::ensure!(
        is_same_category(&source_ext, &target),
        "cross-category conversion is not supported: .{source_ext} -> .{target}"
    );
    anyhow::ensure!(
        matches!(target.as_str(), "xlsx" | "docx"),
        "only .xlsx and .docx conversion targets are supported"
    );
    anyhow::ensure!(
        source_ext != target,
        "input is already in the requested format: .{target}"
    );
    anyhow::ensure!(
        matches!(engine, "auto" | "lo" | "com" | "native"),
        "--engine must be auto, lo, com or native"
    );
    if source_ext == "et" {
        anyhow::ensure!(
            !is_ooxml_spreadsheet_package(input)?,
            "unsupported format: input content is an OOXML Excel workbook with an .et extension; refusing extension-only conversion"
        );
    }
    if engine == "native" {
        anyhow::ensure!(
            target == "xlsx",
            "native value-only conversion is currently available only for Excel to .xlsx"
        );
        if source_ext == "csv" {
            return convert_native_csv_values(input, output);
        }
        return convert_native_excel_values(input, output, &source_ext);
    }
    let (staged_input, _staging_guard) = stage_sidecar_input(input, output, &source_ext)?;
    let staged_output = if source_ext == "et" {
        staged_input.with_extension(&target)
    } else {
        output.to_path_buf()
    };
    let detail = invoke_engine_with_runtime(&staged_input, &staged_output, &target, engine, runtime)?;
    if source_ext == "et" {
        ensure_et_source_format_verified(&detail)?;
        crate::file_commit::create_target_from_candidate(&staged_output, output)
            .context("publish ET conversion only after source format is verified")?;
    }
    Ok(detail)
}

fn ensure_et_source_format_verified(detail: &str) -> Result<()> {
    let Ok(report) = serde_json::from_str::<serde_json::Value>(detail) else {
        bail!("unsupported format: conversion engine did not provide ET source-format evidence");
    };
    let validation = report.get("validation");
    anyhow::ensure!(
        report.get("status").and_then(serde_json::Value::as_str) == Some("converted")
            && report.get("engine").and_then(serde_json::Value::as_str) == Some("com")
            && validation
                .and_then(|value| value.get("source_format_verified"))
                .and_then(serde_json::Value::as_bool)
                == Some(true)
            && validation
                .and_then(|value| value.get("source_format"))
                .and_then(serde_json::Value::as_str)
                == Some("et")
            && validation
                .and_then(|value| value.get("source_file_format"))
                .and_then(serde_json::Value::as_i64)
                == Some(65_521),
        "unsupported format: conversion engine did not verify that the .et source is native ET; output was not published"
    );
    Ok(())
}

fn is_ooxml_spreadsheet_package(path: &Path) -> Result<bool> {
    let file = fs::File::open(path)
        .with_context(|| format!("open .et input for format validation: {}", path.display()))?;
    let mut archive = match zip::ZipArchive::new(file) {
        Ok(archive) => archive,
        Err(_) => return Ok(false),
    };
    let has_workbook_xml = archive.file_names().any(|name| name == "xl/workbook.xml");
    let has_workbook_binary = archive.file_names().any(|name| name == "xl/workbook.bin");
    let mut content_types = match archive.by_name("[Content_Types].xml") {
        Ok(content_types) => content_types,
        Err(_) => return Ok(false),
    };
    let mut content_types_xml = Vec::new();
    content_types
        .by_ref()
        .take(64 * 1024)
        .read_to_end(&mut content_types_xml)
        .context("read .et input content types for format validation")?;
    let content_types_xml = String::from_utf8_lossy(&content_types_xml);

    let xml_spreadsheet_type = [
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
        "application/vnd.ms-excel.template.macroEnabled.main+xml",
    ]
    .iter()
    .any(|content_type| content_types_xml.contains(content_type));
    let binary_spreadsheet_type =
        content_types_xml.contains("application/vnd.ms-excel.sheet.binary.macroEnabled.main");

    Ok((has_workbook_xml && xml_spreadsheet_type)
        || (has_workbook_binary && binary_spreadsheet_type))
}

fn stage_sidecar_input(
    input: &Path,
    output: &Path,
    source_ext: &str,
) -> Result<(PathBuf, Option<TempDirGuard>)> {
    if source_ext != "et" {
        return Ok((input.to_path_buf(), None));
    }

    let output_parent = output
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    fs::create_dir_all(output_parent).with_context(|| {
        format!(
            "create ET conversion staging parent {}",
            output_parent.display()
        )
    })?;
    let staging_dir = unique_sidecar_input_temp_dir(output_parent)?;
    let guard = TempDirGuard(staging_dir.clone());
    let file_name = input
        .file_name()
        .context("ET conversion input has no file name")?;
    let staged_input = staging_dir.join(file_name);
    fs::copy(input, &staged_input).with_context(|| {
        format!(
            "stage ET source {} for external conversion",
            input.display()
        )
    })?;

    Ok((staged_input, Some(guard)))
}

#[cfg(test)]
fn invoke_engine(input: &Path, output: &Path, target: &str, engine: &str) -> Result<String> {
    let mut runtime = crate::conversion_engine::ConversionRuntime::default();
    invoke_engine_with_runtime(input, output, target, engine, &mut runtime)
}

fn invoke_engine_with_runtime(
    input: &Path,
    output: &Path,
    target: &str,
    engine: &str,
    runtime: &mut crate::conversion_engine::ConversionRuntime,
) -> Result<String> {
    crate::conversion_engine::run_fidelity_conversion_with_runtime(input, output, target, engine, runtime).map_err(
        |error| {
            let message = format!("{error:#}");
            if message.starts_with("该功能需要外部引擎") {
                anyhow!("{message}")
            } else if message.contains("install soffice")
                || message.contains("LibreOffice is unavailable")
                || message.contains("LibreOffice is required")
                || message.contains("COM conversion is only available")
                || message.contains("No supported Office/WPS COM engine is available")
                || message.contains("No supported spreadsheet COM engine is available")
            {
                anyhow!("该功能需要外部引擎: {message}")
            } else {
                anyhow!("conversion failed: {}", explain_engine_failure(&message))
            }
        },
    )
}

fn explain_engine_failure(message: &str) -> String {
    if message.contains("Input has .et extension but KET opened FileFormat") {
        return "KET 未将该 .et 文件识别为原生 ET 工作簿，未生成 XLSX。请确认输入是有效的 WPS ET 文件。"
            .to_owned();
    }
    message.trim().to_owned()
}

#[derive(Clone, Copy)]
enum ConversionCategory {
    Excel,
    Word,
}

fn convert_files<F>(
    inputs: &[PathBuf],
    output_dir: Option<&Path>,
    target: &str,
    engine: &str,
    category: ConversionCategory,
    mut on_event: F,
) -> Result<ConversionRun>
where
    F: FnMut(&ConversionProgressEvent),
{
    let run_started = Instant::now();
    let target = target.trim().trim_start_matches('.').to_ascii_lowercase();
    anyhow::ensure!(
        matches!(target.as_str(), "xlsx" | "docx"),
        "conversion target must be xlsx or docx"
    );
    if let Some(output_dir) = output_dir {
        fs::create_dir_all(output_dir)?;
    }
    let mut run = ConversionRun {
        converted: 0,
        skipped: 0,
        failed: 0,
        items: Vec::new(),
    };
    let expansion_started = Instant::now();
    let inputs = expand_conversion_inputs(inputs, category)?;
    let (feature_id, function) = match category {
        ConversionCategory::Excel => ("PT-4-2", "convert-excel"),
        ConversionCategory::Word => ("PT-4-3", "convert-word"),
    };
    let mut sequence = 0usize;
    emit_conversion_event(
        &mut on_event,
        &mut sequence,
        run_started,
        feature_id,
        function,
        "input_expansion",
        "completed",
        0,
        inputs.len(),
        &run,
        None,
        Some(expansion_started.elapsed().as_millis()),
        Vec::new(),
        None,
        None,
        format!("已发现 {} 个待处理文件。", inputs.len()),
    );
    let mut runtime = crate::conversion_engine::ConversionRuntime::default();
    // One LibreOffice session serves the whole batch: stage and convert all
    // LO-routed files up front (single shared profile, one soffice process
    // per stage) so the per-file loop never starts a process. Windows auto
    // routes to COM and never creates a session.
    if engine != "native"
        && std::env::var_os("PYTOOLS_CONVERT_ENGINES_DISABLED").is_none()
    {
        let requests = inputs
            .iter()
            .filter(|input| {
                let Some(ext) = input.extension().and_then(|ext| ext.to_str()) else {
                    return false;
                };
                let ext = ext.to_ascii_lowercase();
                ext != target
                    && crate::conversion_engine::libreoffice_batch_request(&ext, &target, engine)
            })
            .cloned()
            .collect::<Vec<_>>();
        if !requests.is_empty() {
            runtime.prepare_libreoffice_batch(
                &requests
                    .into_iter()
                    .map(|input| (input, target.clone()))
                    .collect::<Vec<_>>(),
            );
        }
    }
    for (index, input) in inputs.iter().enumerate() {
        let input_index = index + 1;
        let item_started = Instant::now();
        let initial_components = vec!["Rust".to_owned(), "pytools-core::conversion".to_owned()];
        emit_conversion_event(
            &mut on_event,
            &mut sequence,
            run_started,
            feature_id,
            function,
            "convert_file",
            "started",
            index,
            inputs.len(),
            &run,
            Some(input_index),
            None,
            initial_components.clone(),
            None,
            None,
            format!("正在处理第 {input_index}/{} 个文件。", inputs.len()),
        );
        if !input.is_file() {
            run.failed += 1;
            run.items.push(ConversionItem {
                input: input.clone(),
                output: None,
                status: "failed".into(),
                detail: "input file does not exist".into(),
            });
            emit_conversion_event(
                &mut on_event,
                &mut sequence,
                run_started,
                feature_id,
                function,
                "convert_file",
                "failed",
                input_index,
                inputs.len(),
                &run,
                Some(input_index),
                Some(item_started.elapsed().as_millis()),
                initial_components,
                None,
                Some("input_not_found".to_owned()),
                format!(
                    "第 {input_index}/{} 个文件失败：输入文件不存在。",
                    inputs.len()
                ),
            );
            continue;
        }
        let ext = input
            .extension()
            .and_then(|x| x.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase();
        let category_ok = match category {
            ConversionCategory::Excel => {
                matches!(
                    ext.as_str(),
                    "csv" | "et" | "xls" | "xlsx" | "xlsm" | "xlsb"
                )
            }
            ConversionCategory::Word => matches!(ext.as_str(), "doc" | "docx"),
        };
        if !category_ok {
            run.failed += 1;
            run.items.push(ConversionItem {
                input: input.clone(),
                output: None,
                status: "failed".into(),
                detail: format!("unsupported same-category conversion source: .{ext}"),
            });
            emit_conversion_event(
                &mut on_event,
                &mut sequence,
                run_started,
                feature_id,
                function,
                "convert_file",
                "failed",
                input_index,
                inputs.len(),
                &run,
                Some(input_index),
                Some(item_started.elapsed().as_millis()),
                initial_components,
                None,
                Some("unsupported_format".to_owned()),
                format!(
                    "第 {input_index}/{} 个文件失败：不支持的文件格式。",
                    inputs.len()
                ),
            );
            continue;
        }
        let stem = input
            .file_stem()
            .map(|s| s.to_string_lossy().into_owned())
            .unwrap_or_else(|| "converted".into());
        let file_output_dir = if let Some(output_dir) = output_dir {
            output_dir.to_path_buf()
        } else {
            let parent = input
                .parent()
                .filter(|path| !path.as_os_str().is_empty())
                .unwrap_or_else(|| Path::new("."));
            parent.join(&target)
        };
        fs::create_dir_all(&file_output_dir)?;
        let mut output = file_output_dir.join(format!("{stem}.{target}"));
        let mut i = 1;
        while output.exists() {
            output = file_output_dir.join(format!("{stem}_{i}.{target}"));
            i += 1;
        }
        let components = conversion_components(&ext, &target, engine);
        let result = if ext == target {
            fs::copy(input, &output)
                .map(|_| "same-format copy; source bytes retained".to_owned())
                .map_err(anyhow::Error::from)
        } else {
            convert_one_with_runtime(input, &output, &target, engine, &mut runtime)
        };
        let (actual_engine, error_type, user_message, event_status) = match result {
            Ok(detail) => {
                run.converted += 1;
                let actual_engine = if ext == target {
                    Some("Rust std::fs::copy".to_owned())
                } else if engine == "native" {
                    Some(if ext == "csv" {
                        "csv + rust_xlsxwriter".to_owned()
                    } else {
                        "calamine + rust_xlsxwriter".to_owned()
                    })
                } else {
                    fidelity_engine(&detail)
                };
                run.items.push(ConversionItem {
                    input: input.clone(),
                    output: Some(output),
                    status: if ext == target {
                        "copied".into()
                    } else {
                        "converted".into()
                    },
                    detail,
                });
                (
                    actual_engine,
                    None,
                    format!("第 {input_index}/{} 个文件处理完成。", inputs.len()),
                    "completed",
                )
            }
            Err(error) => {
                run.failed += 1;
                let detail = format!("{error:#}");
                let error_type = conversion_error_type(&detail);
                let message = match error_type {
                    "external_engine_required" => "需要外部引擎",
                    "unsupported_format" => "不支持的文件格式",
                    "output_exists" => "输出文件已存在",
                    _ => "转换失败",
                };
                run.items.push(ConversionItem {
                    input: input.clone(),
                    output: None,
                    status: "failed".into(),
                    detail,
                });
                (
                    native_engine_for_attempt(&ext, &target, engine),
                    Some(error_type.to_owned()),
                    format!("第 {input_index}/{} 个文件失败：{message}。", inputs.len()),
                    "failed",
                )
            }
        };
        emit_conversion_event(
            &mut on_event,
            &mut sequence,
            run_started,
            feature_id,
            function,
            "convert_file",
            event_status,
            input_index,
            inputs.len(),
            &run,
            Some(input_index),
            Some(item_started.elapsed().as_millis()),
            components,
            actual_engine,
            error_type,
            user_message,
        );
    }
    emit_conversion_event(
        &mut on_event,
        &mut sequence,
        run_started,
        feature_id,
        function,
        "conversion_batch",
        if run.failed == 0 {
            "completed"
        } else {
            "partial_failure"
        },
        inputs.len(),
        inputs.len(),
        &run,
        None,
        Some(run_started.elapsed().as_millis()),
        Vec::new(),
        None,
        if run.failed == 0 {
            None
        } else {
            Some("conversion_failed".to_owned())
        },
        format!("转换结束：成功 {}，失败 {}。", run.converted, run.failed),
    );
    Ok(run)
}

fn emit_conversion_event<F>(
    on_event: &mut F,
    sequence: &mut usize,
    run_started: Instant,
    feature_id: &str,
    function: &str,
    phase: &str,
    status: &str,
    completed_files: usize,
    total_files: usize,
    run: &ConversionRun,
    input_index: Option<usize>,
    stage_elapsed_ms: Option<u128>,
    components: Vec<String>,
    actual_engine: Option<String>,
    error_type: Option<String>,
    user_message: String,
) where
    F: FnMut(&ConversionProgressEvent),
{
    *sequence += 1;
    on_event(&ConversionProgressEvent {
        sequence: *sequence,
        timestamp: Local::now().to_rfc3339(),
        feature_id: feature_id.to_owned(),
        function: function.to_owned(),
        phase: phase.to_owned(),
        status: status.to_owned(),
        completed_files,
        total_files,
        converted_files: run.converted,
        failed_files: run.failed,
        input_index,
        total_elapsed_ms: run_started.elapsed().as_millis(),
        stage_elapsed_ms,
        components,
        actual_engine,
        error_type,
        user_message,
    });
}

fn conversion_components(ext: &str, target: &str, engine: &str) -> Vec<String> {
    if ext == target {
        return vec!["Rust".to_owned(), "std::fs::copy".to_owned()];
    }
    if engine == "native" {
        if ext == "csv" {
            return vec![
                "Rust".to_owned(),
                "pytools-core::conversion".to_owned(),
                "csv".to_owned(),
                "rust_xlsxwriter".to_owned(),
            ];
        }
        return vec![
            "Rust".to_owned(),
            "pytools-core::conversion".to_owned(),
            "excel-reader/calamine".to_owned(),
            "excel-writer/rust_xlsxwriter".to_owned(),
        ];
    }
    vec![
        "Rust".to_owned(),
        "pytools-core::conversion".to_owned(),
        "pytools-core::conversion_engine".to_owned(),
    ]
}

fn fidelity_engine(detail: &str) -> Option<String> {
    let report = serde_json::from_str::<serde_json::Value>(detail).ok()?;
    let engine = report.get("engine")?.as_str()?;
    Some(match engine {
        "lo" => "LibreOffice".to_owned(),
        "com" if report.get("com_bridge").and_then(serde_json::Value::as_str) == Some("rust") => {
            match report.get("com_application").and_then(serde_json::Value::as_str) {
                Some("Excel.Application") => "Microsoft Excel COM (Rust)".to_owned(),
                Some(name) if name.eq_ignore_ascii_case("KET.Application") => "WPS 表格 COM (Rust)".to_owned(),
                _ => "Windows Office/WPS COM (Rust)".to_owned(),
            }
        }
        "com" => "Windows Office/WPS COM".to_owned(),
        other => other.to_owned(),
    })
}

fn native_engine_for_attempt(ext: &str, target: &str, engine: &str) -> Option<String> {
    if engine != "native" || target != "xlsx" {
        return None;
    }
    match ext {
        "csv" => Some("csv".to_owned()),
        "xls" | "xlsm" | "xlsb" => Some("calamine".to_owned()),
        _ => None,
    }
}

fn conversion_error_type(detail: &str) -> &'static str {
    if detail.contains("该功能需要外部引擎") || detail.contains("external fidelity engine") {
        "external_engine_required"
    } else if detail.contains("unsupported") || detail.contains("only .xlsx and .docx") {
        "unsupported_format"
    } else if detail.contains("refusing to overwrite") {
        "output_exists"
    } else {
        "conversion_failed"
    }
}

fn expand_conversion_inputs(
    inputs: &[PathBuf],
    category: ConversionCategory,
) -> Result<Vec<PathBuf>> {
    let mut files = Vec::new();
    let mut seen = HashSet::new();
    for input in inputs {
        if input.is_dir() {
            collect_conversion_directory(input, category, &mut files, &mut seen)?;
        } else if !is_lock_file(input) {
            remember_conversion_file(input, &mut files, &mut seen);
        }
    }
    Ok(files)
}

fn collect_conversion_directory(
    directory: &Path,
    category: ConversionCategory,
    files: &mut Vec<PathBuf>,
    seen: &mut HashSet<String>,
) -> Result<()> {
    let mut entries = fs::read_dir(directory)
        .with_context(|| format!("list conversion input directory {}", directory.display()))?
        .map(|entry| entry.map(|entry| entry.path()))
        .collect::<std::io::Result<Vec<_>>>()?;
    entries.sort_by_key(|path| {
        path.strip_prefix(directory)
            .unwrap_or(path)
            .to_string_lossy()
            .to_lowercase()
    });
    for path in entries {
        let file_type = fs::symlink_metadata(&path)
            .with_context(|| format!("inspect conversion input {}", path.display()))?
            .file_type();
        if file_type.is_dir() {
            collect_conversion_directory(&path, category, files, seen)?;
        } else if file_type.is_file()
            && !is_lock_file(&path)
            && conversion_source_matches_category(&path, category)
        {
            remember_conversion_file(&path, files, seen);
        }
    }
    Ok(())
}

fn remember_conversion_file(path: &Path, files: &mut Vec<PathBuf>, seen: &mut HashSet<String>) {
    let identity = path
        .canonicalize()
        .unwrap_or_else(|_| path.to_path_buf())
        .to_string_lossy()
        .to_lowercase();
    if seen.insert(identity) {
        files.push(path.to_path_buf());
    }
}

fn conversion_source_matches_category(path: &Path, category: ConversionCategory) -> bool {
    let extension = path
        .extension()
        .and_then(|extension| extension.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    match category {
        ConversionCategory::Excel => {
            matches!(
                extension.as_str(),
                "csv" | "et" | "xls" | "xlsx" | "xlsm" | "xlsb"
            )
        }
        ConversionCategory::Word => matches!(extension.as_str(), "doc" | "docx"),
    }
}

fn is_lock_file(path: &Path) -> bool {
    path.file_name()
        .and_then(|name| name.to_str())
        .is_some_and(|name| name.starts_with("~$"))
}

fn is_same_category(source: &str, target: &str) -> bool {
    match target {
        "xlsx" => matches!(source, "csv" | "et" | "xls" | "xlsx" | "xlsm" | "xlsb"),
        "docx" => matches!(source, "doc" | "docx"),
        _ => false,
    }
}

fn convert_native_csv_values(input: &Path, output: &Path) -> Result<String> {
    const MAX_ROWS: usize = 1_048_576;
    const MAX_COLUMNS: usize = 16_384;
    anyhow::ensure!(
        input.is_file(),
        "input file does not exist: {}",
        input.display()
    );
    anyhow::ensure!(
        !output.exists(),
        "refusing to overwrite existing file: {}",
        output.display()
    );
    let bytes = fs::read(input).with_context(|| format!("read CSV {}", input.display()))?;
    let decoded = decode_csv_text(&bytes)?;
    let mut reader = csv::ReaderBuilder::new()
        .has_headers(false)
        .flexible(true)
        .from_reader(decoded.as_bytes());
    let mut rows = Vec::<Vec<String>>::new();
    for record in reader.records() {
        let record = record.with_context(|| format!("parse CSV {}", input.display()))?;
        anyhow::ensure!(
            rows.len() < MAX_ROWS,
            "CSV exceeds Excel's 1,048,576-row limit"
        );
        anyhow::ensure!(
            record.len() <= MAX_COLUMNS,
            "CSV row {} exceeds Excel's 16,384-column limit",
            rows.len() + 1
        );
        rows.push(record.iter().map(str::to_owned).collect());
    }

    let parent = output
        .parent()
        .filter(|path| !path.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    fs::create_dir_all(parent)?;
    let temp_dir = unique_native_temp_dir(parent)?;
    let _guard = TempDirGuard(temp_dir.clone());
    let candidate = temp_dir.join(
        output
            .file_name()
            .context("native CSV output must have a file name")?,
    );
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("Sheet1")?;
    for (row_index, row) in rows.iter().enumerate() {
        let row_index = u32::try_from(row_index).context("CSV exceeds Excel row limit")?;
        for (column_index, value) in row.iter().enumerate() {
            let column_index =
                u16::try_from(column_index).context("CSV exceeds Excel column limit")?;
            if !value.is_empty() {
                worksheet.write_string(row_index, column_index, value)?;
            }
        }
    }
    workbook.save(&candidate).with_context(|| {
        format!(
            "write native CSV conversion candidate {}",
            candidate.display()
        )
    })?;
    validate_native_csv_output(&candidate, &rows)?;
    fs::hard_link(&candidate, output).with_context(|| {
        format!(
            "commit native CSV conversion without replacing an existing file {}",
            output.display()
        )
    })?;
    Ok(format!(
        "按值重建（C01，丢弃样式/公式缓存）；native CSV-to-XLSX conversion; {} rows written as text; no Office type inference or source formatting",
        rows.len()
    ))
}

fn decode_csv_text(bytes: &[u8]) -> Result<String> {
    let bytes = bytes.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(bytes);
    if let Ok(text) = std::str::from_utf8(bytes) {
        return Ok(text.to_owned());
    }
    let (text, _, had_errors) = GBK.decode(bytes);
    anyhow::ensure!(
        !had_errors,
        "CSV is neither valid UTF-8 nor valid GBK; choose an external conversion engine or transcode it first"
    );
    Ok(text.into_owned())
}

fn validate_native_csv_output(candidate: &Path, expected: &[Vec<String>]) -> Result<()> {
    let workbook = read_workbook(candidate)
        .with_context(|| format!("reopen native CSV candidate {}", candidate.display()))?;
    anyhow::ensure!(
        workbook.sheets.len() == 1 && workbook.sheets[0].name == "Sheet1",
        "native CSV candidate must contain exactly one Sheet1 worksheet"
    );
    let sheet = &workbook.sheets[0];
    let mut expected_addresses = std::collections::BTreeSet::new();
    for (row_index, row) in expected.iter().enumerate() {
        for (column_index, value) in row.iter().enumerate() {
            if value.is_empty() {
                continue;
            }
            let address = format!("{}{}", csv_column_name(column_index + 1), row_index + 1);
            expected_addresses.insert(address.clone());
            let actual = sheet
                .values
                .get(&address)
                .with_context(|| format!("native CSV candidate is missing {address}"))?;
            anyhow::ensure!(
                actual.get("type").and_then(serde_json::Value::as_str) == Some("string")
                    && actual.get("value").and_then(serde_json::Value::as_str)
                        == Some(value.as_str()),
                "native CSV value differs at {address}"
            );
        }
    }
    let actual_addresses = sheet
        .values
        .keys()
        .cloned()
        .collect::<std::collections::BTreeSet<_>>();
    anyhow::ensure!(
        actual_addresses == expected_addresses,
        "native CSV candidate contains missing or extra non-empty cells"
    );
    Ok(())
}

fn csv_column_name(mut column: usize) -> String {
    let mut name = String::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        name.insert(0, (b'A' + remainder as u8) as char);
        column = (column - 1) / 26;
    }
    name
}

fn convert_native_excel_values(input: &Path, output: &Path, source_ext: &str) -> Result<String> {
    anyhow::ensure!(
        matches!(source_ext, "xls" | "xlsm" | "xlsb"),
        "native value-only conversion cannot read .{source_ext}; use an external fidelity engine"
    );
    anyhow::ensure!(
        !output.exists(),
        "refusing to overwrite existing file: {}",
        output.display()
    );
    anyhow::ensure!(
        !excel_reader::uses_1904_date_system(input)
            .with_context(|| format!("inspect source date system {}", input.display()))?,
        "native value-only conversion refuses .{source_ext} input using the 1904 date system because the writer cannot preserve its workbook epoch; use an external fidelity engine"
    );

    let source = read_workbook(input).with_context(|| {
        format!(
            "read source workbook for native conversion {}",
            input.display()
        )
    })?;
    validate_native_source_values(&source)?;

    let parent = output
        .parent()
        .filter(|path| !path.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    fs::create_dir_all(parent)?;
    let temp_dir = unique_native_temp_dir(parent)?;
    let guard = TempDirGuard(temp_dir.clone());
    let candidate = temp_dir.join(
        output
            .file_name()
            .context("native conversion output must have a file name")?,
    );

    let mut workbook = Workbook::new();
    for sheet in &source.sheets {
        let worksheet = workbook.add_worksheet();
        worksheet
            .set_name(&sheet.name)
            .with_context(|| format!("create worksheet {:?}", sheet.name))?;
        for (address, formula) in &sheet.formulas {
            let (row, column) = parse_a1(address).with_context(|| {
                format!("invalid formula coordinate {}!{}", sheet.name, address)
            })?;
            let formula = if formula.starts_with('=') {
                formula.clone()
            } else {
                format!("={formula}")
            };
            worksheet
                .write_formula(row, column, formula.as_str())
                .with_context(|| format!("write formula {}!{}", sheet.name, address))?;
        }
        for (address, value) in &sheet.values {
            if sheet.formulas.contains_key(address) {
                continue;
            }
            let (row, column) = parse_a1(address)
                .with_context(|| format!("invalid value coordinate {}!{}", sheet.name, address))?;
            write_native_cell(worksheet, row, column, value)
                .with_context(|| format!("write value {}!{}", sheet.name, address))?;
        }
    }
    workbook
        .save(&candidate)
        .with_context(|| format!("write native conversion candidate {}", candidate.display()))?;

    let converted = read_workbook(&candidate)
        .with_context(|| format!("reopen native conversion candidate {}", candidate.display()))?;
    validate_native_conversion(&source, &converted)?;
    fs::hard_link(&candidate, output).with_context(|| {
        format!(
            "publish validated native conversion without replacing {}",
            output.display()
        )
    })?;
    drop(guard);

    Ok("按值重建（C01，丢弃样式/公式缓存）；native value-only conversion validated: worksheet names/order, stored values and formula text preserved; original formula caches are not preserved; formatting, comments, merges, dimensions, names, page settings, validations, conditional formatting, drawings and macros are not preserved; formulas are not recalculated".to_owned())
}

fn validate_native_source_values(source: &excel_reader::WorkbookDump) -> Result<()> {
    for sheet in &source.sheets {
        for (address, value) in &sheet.values {
            if sheet.formulas.contains_key(address) {
                continue;
            }
            let parsed = CellValue::from_probe_json(value)
                .with_context(|| format!("decode {}!{}", sheet.name, address))?;
            validate_native_cell_value(&sheet.name, address, parsed)?;
        }
    }
    Ok(())
}

fn validate_native_cell_value(sheet_name: &str, address: &str, value: CellValue) -> Result<()> {
    match value {
        CellValue::DateTimeIso(_) | CellValue::DurationIso(_) => bail!(
            "native value-only conversion refuses ISO-typed date/time cell {}!{} because the writer cannot preserve its cell type; use an external fidelity engine",
            sheet_name,
            address
        ),
        CellValue::Text(text) if text.is_empty() => bail!(
            "native value-only conversion cannot preserve an empty-text cell {}!{}; use an external fidelity engine",
            sheet_name,
            address
        ),
        CellValue::Error(_) => bail!(
            "native value-only conversion refuses error cell {}!{} because the writer cannot preserve its typed error value",
            sheet_name,
            address
        ),
        CellValue::Number(number) | CellValue::DateTimeSerial(number) if !number.is_finite() => {
            bail!(
                "native value-only conversion refuses non-finite numeric cell {}!{}",
                sheet_name,
                address
            )
        }
        CellValue::Integer(number) if number.unsigned_abs() > (1_u64 << 53) => {
            bail!(
                "native value-only conversion refuses integer cell {}!{} outside the exact XLSX numeric range",
                sheet_name,
                address
            )
        }
        _ => {}
    }
    Ok(())
}

fn write_native_cell(
    worksheet: &mut excel_writer::Worksheet,
    row: u32,
    column: u16,
    value: &serde_json::Value,
) -> Result<()> {
    match CellValue::from_probe_json(value)? {
        CellValue::Empty => Ok(()),
        CellValue::Text(text) | CellValue::DateTimeIso(text) | CellValue::DurationIso(text) => {
            worksheet.write_string(row, column, text)?;
            Ok(())
        }
        CellValue::Number(number) | CellValue::DateTimeSerial(number) => {
            worksheet.write_number(row, column, number)?;
            Ok(())
        }
        CellValue::Integer(number) => {
            worksheet.write_number(row, column, number as f64)?;
            Ok(())
        }
        CellValue::Boolean(value) => {
            worksheet.write_boolean(row, column, value)?;
            Ok(())
        }
        CellValue::Error(error) => {
            bail!("typed Excel error cells cannot be written natively: {error}")
        }
    }
}

fn validate_native_conversion(
    source: &excel_reader::WorkbookDump,
    converted: &excel_reader::WorkbookDump,
) -> Result<()> {
    let source_names = source
        .sheets
        .iter()
        .map(|sheet| &sheet.name)
        .collect::<Vec<_>>();
    let converted_names = converted
        .sheets
        .iter()
        .map(|sheet| &sheet.name)
        .collect::<Vec<_>>();
    anyhow::ensure!(
        source_names == converted_names,
        "native conversion changed worksheet count, names or order"
    );
    for (source_sheet, output_sheet) in source.sheets.iter().zip(&converted.sheets) {
        anyhow::ensure!(
            source_sheet.formulas.len() == output_sheet.formulas.len(),
            "native conversion changed formula count in worksheet {:?}",
            source_sheet.name
        );
        for (address, formula) in &source_sheet.formulas {
            let output_formula = output_sheet.formulas.get(address).with_context(|| {
                format!(
                    "native conversion removed formula {}!{}",
                    source_sheet.name, address
                )
            })?;
            anyhow::ensure!(
                formula.trim_start_matches('=').trim()
                    == output_formula.trim_start_matches('=').trim(),
                "native conversion changed formula text at {}!{}",
                source_sheet.name,
                address
            );
        }
        let source_cells = source_sheet
            .values
            .keys()
            .filter(|address| !source_sheet.formulas.contains_key(*address))
            .collect::<std::collections::BTreeSet<_>>();
        let output_cells = output_sheet
            .values
            .keys()
            .filter(|address| !output_sheet.formulas.contains_key(*address))
            .collect::<std::collections::BTreeSet<_>>();
        if source_cells != output_cells {
            let missing_addresses = source_cells
                .difference(&output_cells)
                .copied()
                .collect::<Vec<_>>();
            let unexpected_count = output_cells.difference(&source_cells).count();
            let missing = missing_addresses
                .iter()
                .take(12)
                .map(|address| address.as_str())
                .collect::<Vec<_>>();
            let unexpected = output_cells
                .difference(&source_cells)
                .take(12)
                .map(|address| address.as_str())
                .collect::<Vec<_>>();
            let mut missing_types = std::collections::BTreeMap::<&str, usize>::new();
            let mut missing_empty_strings = 0usize;
            for address in &missing_addresses {
                if let Some(value) = source_sheet.values.get(*address) {
                    let value_type = value
                        .get("type")
                        .and_then(serde_json::Value::as_str)
                        .unwrap_or("unknown");
                    *missing_types.entry(value_type).or_default() += 1;
                    if value_type == "string"
                        && value.get("value").and_then(serde_json::Value::as_str) == Some("")
                    {
                        missing_empty_strings += 1;
                    }
                }
            }
            bail!(
                "native conversion changed data cell coordinates in worksheet {:?} (source={}, output={}, missing_in_output={} {:?}, missing_value_types={:?}, missing_empty_strings={}, unexpected_in_output={} {:?})",
                source_sheet.name,
                source_cells.len(),
                output_cells.len(),
                missing_addresses.len(),
                missing,
                missing_types,
                missing_empty_strings,
                unexpected_count,
                unexpected
            );
        }
        for address in source_cells {
            anyhow::ensure!(
                native_values_equal(&source_sheet.values[address], &output_sheet.values[address])?,
                "native conversion changed stored value at {}!{}",
                source_sheet.name,
                address
            );
        }
    }
    Ok(())
}

fn native_values_equal(left: &serde_json::Value, right: &serde_json::Value) -> Result<bool> {
    let left = CellValue::from_probe_json(left)?;
    let right = CellValue::from_probe_json(right)?;
    Ok(match (left, right) {
        (CellValue::Text(left), CellValue::Text(right))
        | (CellValue::DateTimeIso(left), CellValue::DateTimeIso(right))
        | (CellValue::DurationIso(left), CellValue::DurationIso(right)) => left == right,
        (CellValue::Boolean(left), CellValue::Boolean(right)) => left == right,
        (CellValue::Integer(left), CellValue::Integer(right)) => left == right,
        (CellValue::Integer(left), CellValue::Number(right))
        | (CellValue::Number(right), CellValue::Integer(left))
        | (CellValue::Integer(left), CellValue::DateTimeSerial(right))
        | (CellValue::DateTimeSerial(right), CellValue::Integer(left)) => {
            numeric_values_equal(left as f64, right)
        }
        (CellValue::Number(left), CellValue::Number(right))
        | (CellValue::Number(left), CellValue::DateTimeSerial(right))
        | (CellValue::DateTimeSerial(left), CellValue::Number(right))
        | (CellValue::DateTimeSerial(left), CellValue::DateTimeSerial(right)) => {
            numeric_values_equal(left, right)
        }
        (CellValue::Empty, CellValue::Empty) => true,
        _ => false,
    })
}

fn numeric_values_equal(left: f64, right: f64) -> bool {
    if left == right {
        return true;
    }
    left.is_finite()
        && right.is_finite()
        && (left - right).abs() <= 1e-15 * left.abs().max(right.abs())
}

fn parse_a1(address: &str) -> Option<(u32, u16)> {
    let split = address.find(|character: char| character.is_ascii_digit())?;
    let mut column = 0_u32;
    for byte in address[..split].bytes() {
        if !byte.is_ascii_alphabetic() {
            return None;
        }
        column = column.checked_mul(26)?;
        column = column.checked_add((byte.to_ascii_uppercase() - b'A' + 1) as u32)?;
    }
    let row = address[split..].parse::<u32>().ok()?;
    if row == 0 || column == 0 || column > 16_384 {
        return None;
    }
    Some((row - 1, (column - 1) as u16))
}

fn unique_native_temp_dir(parent: &Path) -> Result<PathBuf> {
    use std::time::{SystemTime, UNIX_EPOCH};
    let nanos = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
    for index in 0..100 {
        let path = parent.join(format!(
            ".pytools-native-convert-{}-{nanos}-{index}",
            std::process::id()
        ));
        match fs::create_dir(&path) {
            Ok(()) => return Ok(path),
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => return Err(error.into()),
        }
    }
    bail!("cannot allocate native conversion temporary directory")
}

fn unique_sidecar_input_temp_dir(parent: &Path) -> Result<PathBuf> {
    use std::time::{SystemTime, UNIX_EPOCH};
    let nanos = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
    for index in 0..100 {
        let path = parent.join(format!(
            ".pytools-et-convert-input-{}-{nanos}-{index}",
            std::process::id()
        ));
        match fs::create_dir(&path) {
            Ok(()) => return Ok(path),
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => return Err(error.into()),
        }
    }
    bail!("cannot allocate ET sidecar input staging directory")
}

struct TempDirGuard(PathBuf);

impl Drop for TempDirGuard {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

#[cfg(test)]
mod tests {
    use super::{
        convert_excel_files_in_source_folders, convert_one, ensure_et_source_format_verified,
        explain_engine_failure, invoke_engine, native_values_equal, stage_sidecar_input,
        validate_native_source_values,
    };
    use excel_reader::{read_workbook, SheetDump, WorkbookDump};
    use excel_writer::Workbook;
    use serde_json::json;
    use std::{collections::BTreeMap, fs, path::PathBuf, time::SystemTime};

    struct TempDir(PathBuf);

    impl TempDir {
        fn new(label: &str) -> Self {
            let stamp = SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos();
            let path = std::env::temp_dir().join(format!(
                "pytools-conversion-{label}-{}-{stamp}",
                std::process::id()
            ));
            fs::create_dir_all(&path).unwrap();
            Self(path)
        }
    }

    impl Drop for TempDir {
        fn drop(&mut self) {
            let _ = fs::remove_dir_all(&self.0);
        }
    }

    #[test]
    fn invalid_engine_selection_is_rejected_before_touching_files() {
        let root = TempDir::new("bad-engine");
        let error = invoke_engine(
            &root.0.join("input.xls"),
            &root.0.join("output.xlsx"),
            "xlsx",
            "bogus",
        )
        .expect_err("unknown engine must be rejected");
        assert!(error.to_string().contains("engine must be auto, lo or com"));
    }

    #[test]
    fn et_sidecar_input_is_staged_so_engine_cannot_modify_source() {
        let root = TempDir::new("et-sidecar-stage");
        let input = root.0.join("source.et");
        let output = root.0.join("out/converted.xlsx");
        let original = b"WPS ET source bytes";
        fs::write(&input, original).unwrap();

        let (staged, guard) = stage_sidecar_input(&input, &output, "et").unwrap();
        assert_ne!(staged, input);
        assert!(staged.starts_with(root.0.join("out")));
        assert_eq!(fs::read(&staged).unwrap(), original);

        fs::write(&staged, b"external engine modified its input").unwrap();
        assert_eq!(fs::read(&input).unwrap(), original);
        drop(guard);
        assert!(!staged.exists());
    }

    #[test]
    fn et_sidecar_must_explicitly_verify_native_source_format() {
        let verified = r#"{"status":"converted","engine":"com","validation":{"level":"worksheet_order_names_values_and_formula_text","source_format":"et","source_format_verified":true,"source_file_format":65521}}"#;
        ensure_et_source_format_verified(verified)
            .expect("explicit native ET source evidence should pass");

        for detail in [
            r#"{"status":"converted","engine":"com","validation":{"source_probe":"separate_conversion_using_selected_engine"}}"#,
            r#"{"status":"converted","engine":"com","validation":{"source_format":"xls","source_format_verified":true}}"#,
            r#"{"status":"converted","engine":"com","validation":{"source_format":"et","source_format_verified":true,"source_file_format":56}}"#,
            r#"{"status":"converted","engine":"com","validation":{"source_format":"et","source_format_verified":true}}"#,
            r#"{"status":"converted","engine":"lo","validation":{"source_format":"et","source_format_verified":true,"source_file_format":65521}}"#,
            r#"{"status":"failed","validation":{"source_format":"et","source_format_verified":true}}"#,
            "not json",
        ] {
            assert!(
                ensure_et_source_format_verified(detail).is_err(),
                "sidecar evidence must reject: {detail}"
            );
        }
    }

    #[test]
    fn et_engine_rejection_explains_source_format_without_staging_path() {
        let message = explain_engine_failure(
            r#"powershell stack at F:\staging\convert.ps1: Input has .et extension but KET opened FileFormat 56; expected native ET FileFormat 65521"#,
        );
        assert!(message.contains("未将该 .et 文件识别为原生 ET 工作簿"));
        assert!(message.contains("未生成 XLSX"));
        assert!(!message.contains("F:\\staging"));
        assert!(!message.contains("FileFormat 56"));
    }

    #[test]
    fn explicit_native_conversion_preserves_sheet_data_and_formula_text() {
        let root = TempDir::new("native-values");
        let source_xlsx = root.0.join("source.xlsx");
        let source_xlsm = root.0.join("source.xlsm");
        let output = root.0.join("converted.xlsx");

        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("数据").unwrap();
        sheet.write_string(0, 0, "代码").unwrap();
        sheet.write_number(0, 1, 123.25).unwrap();
        sheet.write_boolean(0, 2, true).unwrap();
        sheet.write_formula(0, 3, "=B1*2").unwrap();
        workbook.add_worksheet().set_name("空工作表").unwrap();
        workbook.save(&source_xlsx).unwrap();
        fs::copy(&source_xlsx, &source_xlsm).unwrap();
        let original_bytes = fs::read(&source_xlsm).unwrap();

        let detail = convert_one(&source_xlsm, &output, "xlsx", "native").unwrap();
        assert!(detail.contains("value-only conversion validated"));
        assert!(detail.contains("macros are not preserved"));
        assert_eq!(fs::read(&source_xlsm).unwrap(), original_bytes);

        let converted = read_workbook(&output).unwrap();
        assert_eq!(
            converted
                .sheets
                .iter()
                .map(|sheet| sheet.name.as_str())
                .collect::<Vec<_>>(),
            vec!["数据", "空工作表"]
        );
        assert_eq!(converted.sheets[0].values["A1"]["value"], "代码");
        assert_eq!(converted.sheets[0].values["B1"]["value"], 123.25);
        assert_eq!(converted.sheets[0].values["C1"]["value"], true);
        assert_eq!(
            converted.sheets[0].formulas["D1"].trim_start_matches('='),
            "B1*2"
        );
    }

    #[test]
    fn native_value_conversion_rejects_empty_text_before_writing() {
        let source = workbook_with_value_and_formula("", None);
        let error = validate_native_source_values(&source)
            .expect_err("empty-text cells cannot be represented by the native writer");
        let message = format!("{error:#}");
        assert!(message.contains("empty-text cell 数据!A1"));
        assert!(message.contains("external fidelity engine"));
        assert!(!message.contains("value="));
    }

    #[test]
    fn native_value_conversion_refuses_iso_date_and_duration_cells_before_writing() {
        let iso_date = json!({"type":"datetime_iso", "value":"2024-02-29T12:30:00"});
        let date_text = json!({"type":"string", "value":"2024-02-29T12:30:00"});
        assert!(!native_values_equal(&iso_date, &date_text).unwrap());

        for (value, expected_kind) in [
            (iso_date, "ISO-typed date/time"),
            (
                json!({"type":"duration_iso", "value":"PT26H"}),
                "ISO-typed date/time",
            ),
        ] {
            let mut source = workbook_with_value_and_formula("", None);
            source.sheets[0].values.insert("A1".to_owned(), value);
            let error = validate_native_source_values(&source)
                .expect_err("ISO-typed values must not be silently converted to text");
            let message = format!("{error:#}");
            assert!(message.contains("数据!A1"), "{message}");
            assert!(message.contains(expected_kind), "{message}");
            assert!(message.contains("cell type"), "{message}");
        }
    }

    #[test]
    fn native_value_conversion_skips_empty_formula_cache_during_preflight() {
        let source = workbook_with_value_and_formula("", Some("=\"\""));
        validate_native_source_values(&source)
            .expect("empty cached formula values are represented by preserved formula text");
    }

    #[test]
    fn native_value_conversion_refuses_et_without_external_engine() {
        let root = TempDir::new("native-et");
        let input = root.0.join("input.et");
        let output = root.0.join("output.xlsx");
        fs::write(&input, b"not an OOXML workbook").unwrap();
        let error = convert_one(&input, &output, "xlsx", "native")
            .expect_err("native value-only reader must reject ET");
        assert!(error.to_string().contains("cannot read .et"));
        assert!(!output.exists());
    }

    fn workbook_with_value_and_formula(value: &str, formula: Option<&str>) -> WorkbookDump {
        let values = BTreeMap::from([("A1".to_owned(), json!({"type":"string", "value":value}))]);
        let formulas = formula
            .map(|formula| BTreeMap::from([("A1".to_owned(), formula.to_owned())]))
            .unwrap_or_default();
        WorkbookDump {
            schema_version: 1,
            reader: "unit-test",
            elapsed_ms: 0,
            sheets: vec![SheetDump {
                name: "数据".to_owned(),
                values,
                formulas,
                formula_status: "available".to_owned(),
                merges: Vec::new(),
                merge_status: "available".to_owned(),
                comments: BTreeMap::new(),
                comments_status: "available".to_owned(),
                conditional_formatting: Vec::new(),
                conditional_formatting_status: "available".to_owned(),
                data_validations: Vec::new(),
                data_validations_status: "available".to_owned(),
            }],
        }
    }

    #[test]
    fn batch_conversion_matches_source_format_folder_and_copies_same_format() {
        let root = TempDir::new("source-folder-layout");
        let input = root.0.join("report.xlsx");
        let output = root.0.join("xlsx/report.xlsx");
        let mut workbook = Workbook::new();
        workbook
            .add_worksheet()
            .write_string(0, 0, "preserve bytes")
            .unwrap();
        workbook.save(&input).unwrap();
        let source_bytes = fs::read(&input).unwrap();

        let result = convert_excel_files_in_source_folders(&[input.clone()], "xlsx", "lo").unwrap();
        assert_eq!(result.converted, 1);
        assert_eq!(result.skipped, 0);
        assert_eq!(result.failed, 0);
        assert_eq!(result.items[0].status, "copied");
        assert_eq!(result.items[0].output.as_deref(), Some(output.as_path()));
        assert_eq!(fs::read(&output).unwrap(), source_bytes);
    }
}
