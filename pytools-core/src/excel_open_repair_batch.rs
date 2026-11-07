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
//! - 非 ZIP 容器、.xls/.et 等其它格式记为 skipped，不尝试解析。

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
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub missing_sheet_refs: Vec<String>,
    pub external_links_dropped_files: usize,
    pub elapsed_ms: u128,
    pub files: Vec<ExcelOpenRepairFileResult>,
}

/// 批量修复：`inputs` 可为文件或目录（`recursive` 时下钻子目录）。
/// 修复副本发布在源文件旁边，原始文件保持不变。
pub fn run_excel_open_repair_batch(
    inputs: &[PathBuf],
    recursive: bool,
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
        missing_sheet_refs: Vec::new(),
        external_links_dropped_files: 0,
        elapsed_ms: 0,
        files: Vec::with_capacity(files.len()),
    };

    for input in files {
        let file_started = Instant::now();
        if !is_supported_workbook(&input) {
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
                reason: Some("仅支持 .xlsx/.xlsm（ZIP 容器）；其它格式不在本修复范围".to_owned()),
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
