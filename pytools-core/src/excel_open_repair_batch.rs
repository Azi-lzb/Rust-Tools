//! 批量修复 Excel COM 无法打开的报送工作簿（文件处理 4-7）。
//!
//! 根因与修复语义移植自 3.open-data-audit 的 `excel_open_repair.py`
//! （全角括号表名裸引用 + 冻结窗格 topLeftCell 矛盾），底层由
//! `ooxml_patch::repair_excel_com_open_issues` 实现。本模块负责批处理编排：
//!
//! - **只修副本**：每本工作簿先复制为源文件旁的 `<名称>-Excel修复.<ext>`，
//!   修复只作用于该副本；原始报送文件永不改写。
//! - **无嫌疑不发布**：预检未发现问题的文件不产出副本（字节不变语义），
//!   在报告中记为 unchanged 并给出原因。
//! - **.xls 旧格式**：可经引擎另存为 .xlsx 后再修复（Windows=Excel/WPS COM、
//!   Linux=LibreOffice 无头转换，由调用方按平台传入 `LegacyConversionEngine`）；
//!   未传引擎或另存失败时记为 skipped，不尝试解析。.et 等其它格式仍 skipped。
//!
//! LO 另存可靠性已实测（2026-11-09，wei40.xls 787 公式 + 合成共享公式样本）：
//! 公式文本与 COM 另存零差异、值零差异（布尔缓存值除外：LO 记 1/0、COM 记
//! TRUE/FALSE——仅缓存表示不同，公式语义一致），两个产物 Excel 均可打开。

use anyhow::{Context, Result};
use ooxml_patch::repair_excel_com_open_issues;
use serde::Serialize;
use std::{
    fs,
    path::{Path, PathBuf},
    time::Instant,
};

#[derive(Clone, Debug, Serialize)]
pub struct ExcelOpenRepairFileResult {
    pub input: PathBuf,
    pub output: Option<PathBuf>,
    pub status: String,
    pub quoted: u64,
    pub pane_fixed: u64,
    pub external_refs_stripped: u64,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub missing_sheet_refs: Vec<String>,
    pub external_links_dropped: bool,
    /// shared formula 的 ref 起点归一计数（Excel 打开保存口径）。
    #[serde(default)]
    pub shared_formula_ref_fixed: u64,
    pub reason: Option<String>,
    pub elapsed_ms: u128,
}

#[derive(Clone, Debug, Serialize)]
pub struct ExcelOpenRepairRunReport {
    pub files_scanned: usize,
    pub workbooks: usize,
    pub repaired: usize,
    pub unchanged: usize,
    pub skipped: usize,
    pub quoted_total: u64,
    pub pane_fixed_total: u64,
    pub external_refs_total: u64,
    pub shared_formula_ref_fixed_total: u64,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub missing_sheet_refs: Vec<String>,
    pub external_links_dropped_files: usize,
    pub elapsed_ms: u128,
    pub files: Vec<ExcelOpenRepairFileResult>,
}

/// 批量修复：`inputs` 可为文件或目录（`recursive` 时下钻子目录）。
/// 修复副本发布在源文件旁边，原始文件保持不变。
/// 非 ZIP 的 .xls 进入修复前，用哪个引擎把它另存为 .xlsx。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum LegacyConversionEngine {
    /// Windows：Excel/WPS COM（与 4-4 同一条 pywin32 路由）。
    WindowsCom,
    /// Linux：LibreOffice 无头转换（soffice --convert-to xlsx）。
    LinuxLibreOffice,
}

pub fn run_excel_open_repair_batch(
    inputs: &[PathBuf],
    recursive: bool,
) -> Result<ExcelOpenRepairRunReport> {
    run_excel_open_repair_batch_inner(inputs, recursive, None)
}

/// 带 .xls 另存路由的批量修复：`legacy_engine` 决定非 ZIP 的 .xls 用哪个
/// 引擎先转成 .xlsx（转换发生在修复之前，转换产物就是 `-Excel修复` 副本）。
pub fn run_excel_open_repair_batch_with_legacy_conversion(
    inputs: &[PathBuf],
    recursive: bool,
    legacy_engine: Option<LegacyConversionEngine>,
) -> Result<ExcelOpenRepairRunReport> {
    run_excel_open_repair_batch_inner(inputs, recursive, legacy_engine)
}

fn run_excel_open_repair_batch_inner(
    inputs: &[PathBuf],
    recursive: bool,
    legacy_engine: Option<LegacyConversionEngine>,
) -> Result<ExcelOpenRepairRunReport> {
    let started = Instant::now();
    let files = collect_inputs(inputs, recursive)?;
    let mut report = ExcelOpenRepairRunReport {
        files_scanned: files.len(),
        workbooks: 0,
        repaired: 0,
        unchanged: 0,
        skipped: 0,
        quoted_total: 0,
        pane_fixed_total: 0,
        external_refs_total: 0,
        shared_formula_ref_fixed_total: 0,
        missing_sheet_refs: Vec::new(),
        external_links_dropped_files: 0,
        elapsed_ms: 0,
        files: Vec::with_capacity(files.len()),
    };

    for input in files {
        let file_started = Instant::now();
        if !is_supported_workbook(&input) {
            // 非 ZIP 容器：.xls 可经引擎另存为 .xlsx 后修复；其它格式仍跳过。
            let is_legacy_excel = input
                .extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| extension.eq_ignore_ascii_case("xls"));
            if is_legacy_excel && legacy_engine.is_some() {
                // 另存产物直接落到 -Excel修复 候选名（.xlsx 后缀）。
                let candidate = match publish_candidate_path(&input.with_extension("xlsx")) {
                    Ok(candidate) => candidate,
                    Err(error) => {
                        report.skipped += 1;
                        report.files.push(ExcelOpenRepairFileResult {
                            input: input.clone(),
                            output: None,
                            status: "skipped".to_owned(),
                            quoted: 0,
                            pane_fixed: 0,
                            external_refs_stripped: 0,
                            missing_sheet_refs: Vec::new(),
                            external_links_dropped: false,
                            shared_formula_ref_fixed: 0,
                            reason: Some(format!("无法创建修复副本：{error:#}")),
                            elapsed_ms: file_started.elapsed().as_millis(),
                        });
                        continue;
                    }
                };
                match convert_legacy_excel(
                    &input,
                    &candidate,
                    legacy_engine.expect("checked above"),
                ) {
                    Ok(()) => {
                        report.workbooks += 1;
                        let repair = repair_excel_com_open_issues(&candidate);
                        if repair.changed {
                            report.repaired += 1;
                            report.quoted_total += repair.quoted;
                            report.pane_fixed_total += repair.pane_fixed;
                            report.external_refs_total += repair.external_refs_stripped;
                            report.shared_formula_ref_fixed_total +=
                                repair.shared_formula_ref_fixed;
                            for name in &repair.missing_sheet_refs {
                                if !report.missing_sheet_refs.contains(name) {
                                    report.missing_sheet_refs.push(name.clone());
                                }
                            }
                            report.files.push(ExcelOpenRepairFileResult {
                                input: input.clone(),
                                output: Some(candidate.clone()),
                                status: "repaired".to_owned(),
                                quoted: repair.quoted,
                                pane_fixed: repair.pane_fixed,
                                external_refs_stripped: repair.external_refs_stripped,
                                missing_sheet_refs: repair.missing_sheet_refs.clone(),
                                external_links_dropped: repair.external_links_dropped,
                                shared_formula_ref_fixed: repair.shared_formula_ref_fixed,
                                reason: Some(format!(
                                    "已从 .xls 经 {} 另存为 .xlsx 后修复（原件未改写）",
                                    match legacy_engine.expect("checked above") {
                                        LegacyConversionEngine::WindowsCom => "Excel/WPS COM",
                                        LegacyConversionEngine::LinuxLibreOffice => "LibreOffice",
                                    }
                                )),
                                elapsed_ms: file_started.elapsed().as_millis(),
                            });
                            continue;
                        }
                        // 另存后的 .xlsx 无需修复点：保留候选副本（转换本身
                        // 就是修复价值，删除会浪费一次成功转换）。
                        report.unchanged += 1;
                        report.files.push(ExcelOpenRepairFileResult {
                            input: input.clone(),
                            output: Some(candidate.clone()),
                            status: "unchanged".to_owned(),
                            quoted: 0,
                            pane_fixed: 0,
                            external_refs_stripped: 0,
                            missing_sheet_refs: Vec::new(),
                            external_links_dropped: false,
                            shared_formula_ref_fixed: 0,
                            reason: Some(format!(
                                "已从 .xls 经 {} 另存为 .xlsx；另存产物无需修复点",
                                match legacy_engine.expect("checked above") {
                                    LegacyConversionEngine::WindowsCom => "Excel/WPS COM",
                                    LegacyConversionEngine::LinuxLibreOffice => "LibreOffice",
                                }
                            )),
                            elapsed_ms: file_started.elapsed().as_millis(),
                        });
                        continue;
                    }
                    Err(error) => {
                        report.skipped += 1;
                        report.files.push(ExcelOpenRepairFileResult {
                            input: input.clone(),
                            output: None,
                            status: "skipped".to_owned(),
                            quoted: 0,
                            pane_fixed: 0,
                            external_refs_stripped: 0,
                            missing_sheet_refs: Vec::new(),
                            external_links_dropped: false,
                            shared_formula_ref_fixed: 0,
                            reason: Some(format!(".xls 另存为 .xlsx 失败：{error:#}")),
                            elapsed_ms: file_started.elapsed().as_millis(),
                        });
                        continue;
                    }
                }
            }
            let hint = if is_legacy_excel {
                ".xls 需要引擎另存为 .xlsx 后才能修复；当前未启用转换引擎（Windows 需 Excel/WPS，Linux 需 LibreOffice）"
            } else {
                "仅支持 .xlsx/.xlsm（ZIP 容器）与可另存的 .xls；其它格式不在本修复范围"
            };
            report.skipped += 1;
            report.files.push(ExcelOpenRepairFileResult {
                input: input.clone(),
                output: None,
                status: "skipped".to_owned(),
                quoted: 0,
                pane_fixed: 0,
                external_refs_stripped: 0,
                missing_sheet_refs: Vec::new(),
                external_links_dropped: false,
                shared_formula_ref_fixed: 0,
                reason: Some(hint.to_owned()),
                elapsed_ms: file_started.elapsed().as_millis(),
            });
            continue;
        }
        report.workbooks += 1;
        let candidate = match publish_candidate_path(&input) {
            Ok(candidate) => candidate,
            Err(error) => {
                report.skipped += 1;
                report.files.push(ExcelOpenRepairFileResult {
                    input: input.clone(),
                    output: None,
                    status: "skipped".to_owned(),
                    quoted: 0,
                    pane_fixed: 0,
                    external_refs_stripped: 0,
                    missing_sheet_refs: Vec::new(),
                    external_links_dropped: false,
                    shared_formula_ref_fixed: 0,
                    reason: Some(format!("无法创建修复副本：{error:#}")),
                    elapsed_ms: file_started.elapsed().as_millis(),
                });
                continue;
            }
        };
        let copy_result = fs::copy(&input, &candidate)
            .with_context(|| format!("复制修复副本 {}", candidate.display()));
        if let Err(error) = copy_result {
            report.skipped += 1;
            report.files.push(ExcelOpenRepairFileResult {
                input: input.clone(),
                output: None,
                status: "skipped".to_owned(),
                quoted: 0,
                pane_fixed: 0,
                external_refs_stripped: 0,
                missing_sheet_refs: Vec::new(),
                external_links_dropped: false,
                shared_formula_ref_fixed: 0,
                reason: Some(format!("{error:#}")),
                elapsed_ms: file_started.elapsed().as_millis(),
            });
            continue;
        }
        let repair = repair_excel_com_open_issues(&candidate);
        if repair.changed {
            report.repaired += 1;
            report.quoted_total += repair.quoted;
            report.pane_fixed_total += repair.pane_fixed;
            report.external_refs_total += repair.external_refs_stripped;
            report.shared_formula_ref_fixed_total += repair.shared_formula_ref_fixed;
            if repair.external_links_dropped {
                report.external_links_dropped_files += 1;
            }
            for name in &repair.missing_sheet_refs {
                if !report.missing_sheet_refs.contains(name) {
                    report.missing_sheet_refs.push(name.clone());
                }
            }
            report.files.push(ExcelOpenRepairFileResult {
                input: input.clone(),
                output: Some(candidate),
                status: "repaired".to_owned(),
                quoted: repair.quoted,
                pane_fixed: repair.pane_fixed,
                external_refs_stripped: repair.external_refs_stripped,
                missing_sheet_refs: repair.missing_sheet_refs.clone(),
                external_links_dropped: repair.external_links_dropped,
                shared_formula_ref_fixed: repair.shared_formula_ref_fixed,
                reason: None,
                elapsed_ms: file_started.elapsed().as_millis(),
            });
        } else {
            // 无嫌疑/修复失败都不发布副本；文件字节语义保持不变。
            let _ = fs::remove_file(&candidate);
            report.unchanged += 1;
            report.files.push(ExcelOpenRepairFileResult {
                input: input.clone(),
                output: None,
                status: "unchanged".to_owned(),
                quoted: 0,
                pane_fixed: 0,
                external_refs_stripped: 0,
                missing_sheet_refs: Vec::new(),
                external_links_dropped: false,
                shared_formula_ref_fixed: 0,
                reason: Some(if repair.reason.is_empty() {
                    "未发现需要修复的内容".to_owned()
                } else {
                    repair.reason
                }),
                elapsed_ms: file_started.elapsed().as_millis(),
            });
        }
    }
    report.elapsed_ms = started.elapsed().as_millis();
    Ok(report)
}

fn is_supported_workbook(path: &Path) -> bool {
    path.is_file()
        && path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| {
                extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
            })
}

/// 修复副本发布名：`<stem>-Excel修复.<ext>`，已存在时追加 `_1`、`_2`…
/// 把非 ZIP 的 .xls 另存为 .xlsx（目标即 `-Excel修复` 副本路径）。
///
/// Windows：Excel/WPS COM（与 4-4 同一条 pywin32 路由，经 ConversionRuntime）。
/// Linux：LibreOffice 无头转换（soffice --convert-to xlsx），soffice 缺失时报错。
/// 转换在临时文件上进行，成功后原子改名到目标，失败时清理临时文件。
fn convert_legacy_excel(input: &Path, target: &Path, engine: LegacyConversionEngine) -> Result<()> {
    let parent = input.parent().context("输入文件没有父目录")?;
    let nonce = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|duration| duration.as_nanos())
        .unwrap_or_default();
    let staging = parent.join(format!(
        ".{stem}.repair-convert.{nonce}.xlsx",
        stem = input
            .file_stem()
            .and_then(|stem| stem.to_str())
            .context("输入文件名无法解析")?,
    ));
    let result = (|| -> Result<()> {
        match engine {
            LegacyConversionEngine::WindowsCom => {
                #[cfg(windows)]
                {
                    let mut runtime = crate::conversion_engine::ConversionRuntime::default();
                    crate::conversion::convert_one_with_runtime(
                        input,
                        &staging,
                        "xlsx",
                        "auto",
                        &mut runtime,
                    )?;
                }
                #[cfg(not(windows))]
                {
                    anyhow::bail!("Windows COM 另存仅在 Windows 可用");
                }
            }
            LegacyConversionEngine::LinuxLibreOffice => {
                #[cfg(target_os = "linux")]
                {
                    let soffice =
                        std::env::var("PYTOOLS_SOFFICE").unwrap_or_else(|_| "soffice".to_owned());
                    let output = std::process::Command::new(soffice)
                        .args(["--headless", "--norestore", "--convert-to", "xlsx"])
                        .arg("--outdir")
                        .arg(parent)
                        .arg(input)
                        .output()
                        .map_err(|error| {
                            anyhow::anyhow!("启动 LibreOffice 失败（soffice 不可用？）：{error}")
                        })?;
                    if !output.status.success() {
                        anyhow::bail!(
                            "LibreOffice 转换失败：{}",
                            String::from_utf8_lossy(&output.stderr)
                        );
                    }
                    // soffice 的产物名 = <stem>.xlsx，与暂存名不同则改名。
                    let produced = parent.join(format!(
                        "{stem}.xlsx",
                        stem = input
                            .file_stem()
                            .and_then(|stem| stem.to_str())
                            .context("输入文件名无法解析")?
                    ));
                    if produced != staging {
                        fs::rename(&produced, &staging).with_context(|| {
                            format!("移动 LibreOffice 转换产物 {}", produced.display())
                        })?;
                    }
                }
                #[cfg(not(target_os = "linux"))]
                {
                    anyhow::bail!("LibreOffice 另存仅在 Linux 可用");
                }
            }
        }
        if !staging.is_file() || fs::metadata(&staging)?.len() == 0 {
            anyhow::bail!("另存产物为空或不存在");
        }
        Ok(())
    })();
    match result {
        Ok(()) => {
            // 暂存名换成正式候选名（原子改名，同目录同卷）。
            fs::rename(&staging, target)
                .with_context(|| format!("移动另存产物到修复候选名 {}", target.display()))?;
            Ok(())
        }
        Err(error) => {
            let _ = fs::remove_file(&staging);
            Err(error)
        }
    }
}

fn publish_candidate_path(input: &Path) -> Result<PathBuf> {
    let parent = input.parent().context("输入文件没有父目录")?;
    let stem = input
        .file_stem()
        .and_then(|stem| stem.to_str())
        .context("输入文件名无法解析")?;
    let extension = input
        .extension()
        .and_then(|extension| extension.to_str())
        .context("输入文件没有扩展名")?;
    let base = parent.join(format!("{stem}-Excel修复.{extension}"));
    if !base.exists() {
        return Ok(base);
    }
    for suffix in 1..10_000u32 {
        let candidate = parent.join(format!("{stem}-Excel修复_{suffix}.{extension}"));
        if !candidate.exists() {
            return Ok(candidate);
        }
    }
    anyhow::bail!("无法为 {} 分配唯一的修复副本名称", input.display())
}

fn collect_inputs(inputs: &[PathBuf], recursive: bool) -> Result<Vec<PathBuf>> {
    let mut files = Vec::new();
    for input in inputs {
        if input.is_dir() {
            collect_directory(input, recursive, &mut files)?;
        } else {
            files.push(input.clone());
        }
    }
    let mut unique = std::collections::BTreeMap::new();
    for path in files {
        let key = fs::canonicalize(&path)
            .unwrap_or_else(|_| path.clone())
            .to_string_lossy()
            .to_lowercase();
        unique.entry(key).or_insert(path);
    }
    Ok(unique.into_values().collect())
}

fn collect_directory(root: &Path, recursive: bool, files: &mut Vec<PathBuf>) -> Result<()> {
    let mut pending = vec![root.to_path_buf()];
    while let Some(directory) = pending.pop() {
        let mut entries = fs::read_dir(&directory)
            .with_context(|| format!("scan directory {}", directory.display()))?
            .collect::<std::io::Result<Vec<_>>>()?;
        entries.sort_by_key(|entry| entry.file_name().to_string_lossy().to_lowercase());
        for entry in entries {
            let path = entry.path();
            let metadata = fs::symlink_metadata(&path)
                .with_context(|| format!("inspect {}", path.display()))?;
            if metadata.file_type().is_symlink() {
                continue;
            }
            if metadata.is_file() {
                // Office 锁文件与上一次运行的修复副本都不再处理。
                let name = path
                    .file_name()
                    .map(|name| name.to_string_lossy().into_owned());
                if name.as_deref().is_some_and(|name| name.starts_with("~$")) {
                    continue;
                }
                if name
                    .as_deref()
                    .is_some_and(|name| name.contains("-Excel修复"))
                {
                    continue;
                }
                files.push(path);
            } else if metadata.is_dir() && recursive {
                pending.push(path);
            }
        }
    }
    Ok(())
}
