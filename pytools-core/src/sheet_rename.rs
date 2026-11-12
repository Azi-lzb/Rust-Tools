//! Plan and safe-copy execution for V2 file menu 4-5 sheet renaming.

use crate::{
    conversion_engine::ConversionRuntime,
    file_commit::{
        create_target_from_candidate, replace_target_from_candidate, sha256_file,
        verify_planning_snapshot, TargetCommitReport,
    },
};
use anyhow::{bail, Context, Result};
use excel_editor::{
    inspect_xlsx_profile, EditExecutionReport, EditOperation, ExcelEditPlan,
    PreservationRequirement, WorkbookFeature, WorkbookFormat,
};
use excel_reader::read_workbook;
use serde::Serialize;
use std::{
    collections::{BTreeMap, BTreeSet},
    fs,
    path::{Path, PathBuf},
    sync::atomic::{AtomicU64, Ordering},
    time::Instant,
};

static CANDIDATE_SEQUENCE: AtomicU64 = AtomicU64::new(0);
pub const SHEET_RENAME_OUTPUT_DIRECTORY: &str = "PyTools_Sheet改名结果";

#[derive(Clone, Debug, Serialize)]
pub struct SheetRenameReport {
    pub input: PathBuf,
    pub output: PathBuf,
    pub workbooks: usize,
    pub renamed: usize,
    pub skipped: usize,
    pub plan: Option<ExcelEditPlan>,
}

#[derive(Clone, Debug, Serialize)]
pub struct SheetRenameFileResult {
    pub input: PathBuf,
    pub output: Option<PathBuf>,
    pub status: String,
    pub engine: Option<String>,
    pub renamed: usize,
    pub skipped: usize,
    pub reason: Option<String>,
    pub planning_ms: u128,
    pub editing_ms: u128,
    pub commit_ms: u128,
    pub elapsed_ms: u128,
    pub execution: Option<EditExecutionReport>,
    pub commit: Option<TargetCommitReport>,
}

#[derive(Clone, Debug, Serialize)]
pub struct SheetRenameRunReport {
    pub files_scanned: usize,
    pub workbooks: usize,
    pub rename_ok: usize,
    pub skip: usize,
    pub elapsed_ms: u128,
    pub files: Vec<SheetRenameFileResult>,
}

pub fn build_sheet_rename_plan(
    config: &Path,
    input: &Path,
    output: &Path,
) -> Result<SheetRenameReport> {
    anyhow::ensure!(
        input.is_file(),
        "input workbook does not exist: {}",
        input.display()
    );
    anyhow::ensure!(
        is_xlsx(input) && is_xlsx(output),
        "sheet rename candidate accepts .xlsx only; .xls must be converted before this operation"
    );
    let mapping = load_mapping(config)?;
    build_sheet_rename_plan_with_mapping(&mapping, input, output)
}

fn build_sheet_rename_plan_with_mapping(
    mapping: &BTreeMap<String, String>,
    input: &Path,
    output: &Path,
) -> Result<SheetRenameReport> {
    anyhow::ensure!(!mapping.is_empty(), "重命名配置 缺少原表名/新表名映射。");
    let dump = read_workbook(input)?;
    let names = dump
        .sheets
        .iter()
        .map(|sheet| sheet.name.clone())
        .collect::<Vec<_>>();
    let (changes, skipped) = proposed_sheet_renames(mapping, &names);
    let mut operations = Vec::with_capacity(changes.len());
    for (old_name, new_name) in changes {
        operations.push(EditOperation::RenameSheet { old_name, new_name });
    }
    let renamed = operations.len();
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
            request_id: "pytools-sheet-rename".into(),
            input: input.into(),
            output: output.into(),
            format: WorkbookFormat::Xlsx,
            workbook: inspect_xlsx_profile(input)?,
            operations,
            preserve,
            drop_allowed: Default::default(),
        };
        plan.validate()?;
        Some(plan)
    };
    Ok(SheetRenameReport {
        input: input.into(),
        output: output.into(),
        workbooks: 1,
        renamed,
        skipped,
        plan,
    })
}

fn proposed_sheet_renames(
    mapping: &BTreeMap<String, String>,
    names: &[String],
) -> (Vec<(String, String)>, usize) {
    let mut changes = Vec::new();
    let mut current_names = names.iter().cloned().collect::<BTreeSet<_>>();
    let mut skipped = 0;
    for old in names {
        let lookup = if mapping.contains_key(old) {
            old.clone()
        } else {
            old.trim().to_owned()
        };
        let Some(new) = mapping.get(&lookup) else {
            continue;
        };
        if old == new {
            skipped += 1;
            continue;
        }
        if !valid_name(new) || has_conflicting_sheet_name(&current_names, old, new) {
            skipped += 1;
            continue;
        }
        changes.push((old.clone(), new.clone()));
        current_names.remove(old);
        current_names.insert(new.clone());
    }
    (changes, skipped)
}

/// Compatibility wrapper for callers that still pass a legacy capability
/// file. The file extension selects the already-approved 4-5 route; the
/// capability is not consulted.
pub fn run_batch_sheet_rename_candidate(
    config: &Path,
    inputs: &[PathBuf],
    recursive: bool,
    capability_path: &Path,
    platform: &str,
) -> Result<SheetRenameRunReport> {
    let _ = (capability_path, platform);
    run_batch_sheet_rename_impl(config, inputs, recursive)
}

/// Use Openpyxl for `.xlsx` and the dedicated Windows Excel/WPS COM rename
/// operation for legacy `.xls`, without requiring a capability record.
pub fn run_batch_sheet_rename(
    config: &Path,
    inputs: &[PathBuf],
    recursive: bool,
) -> Result<SheetRenameRunReport> {
    run_batch_sheet_rename_impl(config, inputs, recursive)
}

/// Compatibility alias retained for existing callers.
pub fn run_batch_sheet_rename_openpyxl(
    config: &Path,
    inputs: &[PathBuf],
    recursive: bool,
) -> Result<SheetRenameRunReport> {
    run_batch_sheet_rename(config, inputs, recursive)
}

fn run_batch_sheet_rename_impl(
    config: &Path,
    inputs: &[PathBuf],
    recursive: bool,
) -> Result<SheetRenameRunReport> {
    let started = Instant::now();
    let mut excel_runtime = ConversionRuntime::default();
    let mapping = load_mapping(config)?;
    anyhow::ensure!(!mapping.is_empty(), "重命名配置 缺少原表名/新表名映射。");
    let files = collect_inputs(inputs, recursive)?;
    let mut report = SheetRenameRunReport {
        files_scanned: files.len(),
        workbooks: 0,
        rename_ok: 0,
        skip: 0,
        elapsed_ms: 0,
        files: Vec::with_capacity(files.len()),
    };

    for input in files {
        let file_started = Instant::now();
        if !input.is_file() {
            report.skip += 1;
            report.files.push(skipped_file_result(
                input,
                "not_found_or_not_a_file",
                file_started.elapsed().as_millis(),
            ));
            continue;
        }
        let input_is_xls = is_xls(&input);
        if !is_xlsx(&input) && !input_is_xls {
            report.skip += 1;
            report.files.push(skipped_file_result(
                input,
                "仅支持 .xlsx 和 Windows Excel/WPS COM 可处理的 .xls 工作簿",
                file_started.elapsed().as_millis(),
            ));
            continue;
        }
        if input_is_xls && !cfg!(windows) {
            report.skip += 1;
            report.files.push(skipped_file_result(
                input,
                ".xls 工作表改名需要 Windows Excel/WPS COM；可先在 Windows 使用 4-4 转成 .xlsx",
                file_started.elapsed().as_millis(),
            ));
            continue;
        }

        let planning_started = Instant::now();
        let input_hash = sha256_file(&input)?;
        let output = sheet_rename_output_path(&input)?;
        let candidate = candidate_path(&input)?;
        let (changes, skipped, plan) = if input_is_xls {
            let dump = read_workbook(&input).with_context(|| {
                format!("读取 .xls 工作簿的工作表名称失败：{}", input.display())
            })?;
            let names = dump
                .sheets
                .iter()
                .map(|sheet| sheet.name.clone())
                .collect::<Vec<_>>();
            let (changes, skipped) = proposed_sheet_renames(&mapping, &names);
            (changes, skipped, None)
        } else {
            let plan_report = build_sheet_rename_plan_with_mapping(&mapping, &input, &candidate)
                .with_context(|| format!("build sheet-rename plan for {}", input.display()))?;
            let changes = plan_report
                .plan
                .as_ref()
                .into_iter()
                .flat_map(|plan| plan.operations.iter())
                .filter_map(|operation| match operation {
                    EditOperation::RenameSheet { old_name, new_name } => {
                        Some((old_name.clone(), new_name.clone()))
                    }
                    _ => None,
                })
                .collect::<Vec<_>>();
            (changes, plan_report.skipped, plan_report.plan)
        };
        let planning_ms = planning_started.elapsed().as_millis();
        report.workbooks += 1;
        report.skip += skipped;

        if changes.is_empty() {
            report.files.push(SheetRenameFileResult {
                input,
                output: None,
                status: "unchanged".to_owned(),
                engine: None,
                renamed: 0,
                skipped,
                reason: None,
                planning_ms,
                editing_ms: 0,
                commit_ms: 0,
                elapsed_ms: file_started.elapsed().as_millis(),
                execution: None,
                commit: None,
            });
            continue;
        }

        let mut candidate_owned = false;
        let result = (|| -> Result<(Option<EditExecutionReport>, TargetCommitReport, u128, String)> {
            verify_planning_snapshot(&input, Some(&input_hash))?;
            let editing_started = Instant::now();
            if let Some(plan) = &plan {
                anyhow::ensure!(
                    plan.workbook.inventory_complete,
                    "工作簿包含未识别的结构，拒绝覆盖源文件"
                );
                anyhow::ensure!(
                    !plan.workbook.features_present.iter().any(|feature| matches!(
                        feature,
                        WorkbookFeature::DigitalSignatures | WorkbookFeature::ActiveX | WorkbookFeature::ExternalLinks | WorkbookFeature::Macros
                    )),
                    "工作簿包含数字签名、ActiveX、外部链接或宏，当前 Openpyxl 改名路径未验证，拒绝覆盖源文件"
                );
            }
            let (execution, engine) = if let (Some(plan), true) = (
                plan.as_ref(),
                std::env::var_os(crate::excel_edit_router::ROUTE_ENV).is_some(),
            ) {
                candidate_owned = true;
                let route_config = crate::excel_edit_router::RouteConfig::from_environment()?;
                let execution = crate::excel_edit_router::execute_candidate_with_route_config(
                    plan,
                    &route_config,
                )
                .with_context(|| format!("ExcelEditPlan 工作表改名失败：{}", input.display()))?;
                let engine = crate::excel_core_diag::engine_label(execution.engine).to_owned();
                (Some(execution), engine)
            } else {
                candidate_owned = true;
                let mut original = fs::File::open(&input)
                    .with_context(|| format!("读取源工作簿失败：{}", input.display()))?;
                let mut temporary = fs::OpenOptions::new()
                    .write(true)
                    .create_new(true)
                    .open(&candidate)
                    .with_context(|| format!("创建改名临时候选失败：{}", candidate.display()))?;
                std::io::copy(&mut original, &mut temporary)?;
                temporary.sync_all()?;
                drop(temporary);
                // Windows cannot replace the source while this read handle is open.
                drop(original);
                let engine = if input_is_xls {
                    excel_runtime
                        .rename_sheets_com_in_candidate(&candidate, &changes)
                        .with_context(|| {
                            format!("Excel/WPS COM 工作表改名失败：{}", input.display())
                        })?
                } else {
                    excel_runtime
                        .rename_sheets_in_candidate(&candidate, &changes)
                        .with_context(|| format!("Openpyxl 工作表改名失败：{}", input.display()))?;
                    "openpyxl".to_owned()
                };
                (None, engine)
            };
            let editing_ms = editing_started.elapsed().as_millis();
            let commit_started = Instant::now();
            verify_planning_snapshot(&input, Some(&input_hash))?;
            // 就地模式：目标就是源文件本身，用带规划哈希校验的替换提交，
            // 若源在规划后被改动则拒绝写入（保持数据保护语义）。
            let mut commit = if output == input {
                replace_target_from_candidate(&candidate, &output, &input_hash).with_context(
                    || format!("safely replace renamed workbook {}", output.display()),
                )?
            } else {
                create_target_from_candidate(&candidate, &output).with_context(|| {
                    format!("safely publish renamed copy to {}", output.display())
                })?
            };
            commit.elapsed_ms = commit_started.elapsed().as_millis();
            Ok((execution, commit, editing_ms, engine))
        })();
        let cleanup = if candidate_owned {
            match fs::remove_file(&candidate) {
                Ok(()) => None,
                Err(error) if error.kind() == std::io::ErrorKind::NotFound => None,
                Err(error) => Some(format!("temporary candidate cleanup failed: {error}")),
            }
        } else {
            None
        };
        let (execution, mut commit, editing_ms, engine) = match result {
            Ok(result) => result,
            Err(error) => {
                return Err(match cleanup {
                    Some(cleanup) => error.context(cleanup),
                    None => error,
                });
            }
        };
        if let Some(cleanup) = cleanup {
            commit.cleanup_warning = Some(match commit.cleanup_warning.take() {
                Some(previous) => format!("{previous}; {cleanup}"),
                None => cleanup,
            });
        }
        report.rename_ok += changes.len();
        report.files.push(SheetRenameFileResult {
            input,
            output: Some(output),
            status: "renamed".to_owned(),
            engine: Some(engine),
            renamed: changes.len(),
            skipped,
            reason: None,
            planning_ms,
            editing_ms,
            commit_ms: commit.elapsed_ms,
            elapsed_ms: file_started.elapsed().as_millis(),
            execution,
            commit: Some(commit),
        });
    }
    report.elapsed_ms = started.elapsed().as_millis();
    Ok(report)
}

fn skipped_file_result(input: PathBuf, reason: &str, elapsed_ms: u128) -> SheetRenameFileResult {
    SheetRenameFileResult {
        input,
        output: None,
        status: "skipped".to_owned(),
        engine: None,
        renamed: 0,
        skipped: 1,
        reason: Some(reason.to_owned()),
        planning_ms: 0,
        editing_ms: 0,
        commit_ms: 0,
        elapsed_ms,
        execution: None,
        commit: None,
    }
}

fn candidate_path(input: &Path) -> Result<PathBuf> {
    let parent = input
        .parent()
        .context("input workbook has no parent directory")?;
    let extension = input
        .extension()
        .and_then(|value| value.to_str())
        .unwrap_or("xlsx");
    for _ in 0..256 {
        let sequence = CANDIDATE_SEQUENCE.fetch_add(1, Ordering::Relaxed);
        let candidate = parent.join(format!(
            ".pytools-sheet-rename-{}-{sequence}.{extension}",
            std::process::id(),
        ));
        if !candidate.exists() {
            return Ok(candidate);
        }
    }
    bail!(
        "could not reserve a unique sheet-rename candidate beside {}",
        input.display()
    )
}

fn sheet_rename_output_path(input: &Path) -> Result<PathBuf> {
    // 4-5 改为就地修改：输出即源文件本身，不再生成结果目录副本。
    // 执行前有 sha256 快照校验，执行后经候选文件验证再替换，源数据仍受保护。
    input
        .file_name()
        .context("input workbook has no file name")?;
    Ok(input.to_path_buf())
}

fn ensure_result_directory(path: &Path, create: bool) -> Result<()> {
    if create {
        fs::create_dir_all(path)
            .with_context(|| format!("create sheet rename result directory {}", path.display()))?;
    }
    match fs::symlink_metadata(path) {
        Ok(metadata) => {
            anyhow::ensure!(
                metadata.is_dir() && !metadata.file_type().is_symlink(),
                "sheet rename result path is not a regular directory: {}",
                path.display()
            );
            Ok(())
        }
        Err(error) if !create && error.kind() == std::io::ErrorKind::NotFound => Ok(()),
        Err(error) => Err(error)
            .with_context(|| format!("inspect sheet rename result directory {}", path.display())),
    }
}

fn path_exists_including_symlink(path: &Path) -> bool {
    fs::symlink_metadata(path).is_ok()
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
    let mut unique = BTreeMap::new();
    for path in files {
        let key_path = fs::canonicalize(&path).unwrap_or_else(|_| path.clone());
        let key = key_path.to_string_lossy().to_lowercase();
        unique.entry(key).or_insert(key_path);
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
                .with_context(|| format!("inspect selected path {}", path.display()))?;
            if metadata.file_type().is_symlink() {
                continue;
            }
            if metadata.is_file() {
                if !path
                    .file_name()
                    .is_some_and(|name| name.to_string_lossy().starts_with("~$"))
                {
                    files.push(path);
                }
            } else if recursive && metadata.is_dir() {
                if path
                    .file_name()
                    .is_some_and(|name| name == SHEET_RENAME_OUTPUT_DIRECTORY)
                {
                    continue;
                }
                pending.push(path);
            }
        }
    }
    Ok(())
}
fn load_mapping(config: &Path) -> Result<BTreeMap<String, String>> {
    let dump =
        read_workbook(config).with_context(|| format!("read config {}", config.display()))?;
    let sheet = dump
        .sheets
        .iter()
        .find(|sheet| sheet.name == "重命名配置")
        .context("重命名配置 缺少工作表")?;
    let headers = sheet
        .values
        .iter()
        .filter_map(|(address, value)| {
            parse_a1(address)
                .filter(|(row, _)| *row == 1)
                .map(|(_, col)| (value_text(value), col))
        })
        .collect::<BTreeMap<_, _>>();
    let old_col = headers.get("原表名").copied().unwrap_or(10);
    let new_col = headers.get("新表名").copied().unwrap_or(11);
    let max_row = sheet
        .values
        .keys()
        .filter_map(|address| parse_a1(address).map(|v| v.0))
        .max()
        .unwrap_or(1);
    let mut result = BTreeMap::new();
    for row in 2..=max_row {
        let old = cell_text(sheet, row, old_col);
        let new = cell_text(sheet, row, new_col);
        if !old.trim().is_empty() && !new.trim().is_empty() {
            result.insert(old.trim().to_owned(), new.trim().to_owned());
        }
    }
    Ok(result)
}
fn valid_name(name: &str) -> bool {
    !name.trim().is_empty()
        && name.chars().count() <= 31
        && !name
            .chars()
            .any(|c| matches!(c, ':' | '\\' | '/' | '?' | '*' | '[' | ']'))
        && !name.starts_with('\'')
        && !name.ends_with('\'')
}
fn has_conflicting_sheet_name(current_names: &BTreeSet<String>, old: &str, new: &str) -> bool {
    let normalized_new = new.to_lowercase();
    current_names
        .iter()
        .any(|existing| existing != old && existing.to_lowercase() == normalized_new)
}
fn is_xlsx(path: &Path) -> bool {
    path.extension()
        .and_then(|e| e.to_str())
        .is_some_and(|e| e.eq_ignore_ascii_case("xlsx"))
}
fn is_xls(path: &Path) -> bool {
    path.extension()
        .and_then(|e| e.to_str())
        .is_some_and(|e| e.eq_ignore_ascii_case("xls"))
}
fn parse_a1(address: &str) -> Option<(usize, usize)> {
    let split = address.find(|c: char| c.is_ascii_digit())?;
    let col = address[..split].bytes().try_fold(0usize, |acc, b| {
        acc.checked_mul(26)?
            .checked_add((b.to_ascii_uppercase() - b'A' + 1) as usize)
    })?;
    Some((address[split..].parse().ok()?, col))
}
fn column_letters(mut col: usize) -> String {
    let mut out = String::new();
    while col > 0 {
        out.insert(0, (b'A' + ((col - 1) % 26) as u8) as char);
        col = (col - 1) / 26;
    }
    out
}
fn cell_text(sheet: &excel_reader::SheetDump, row: usize, col: usize) -> String {
    value_text(
        sheet
            .values
            .get(&format!("{}{}", column_letters(col), row))
            .unwrap_or(&serde_json::Value::Null),
    )
}
fn value_text(value: &serde_json::Value) -> String {
    value
        .get("value")
        .map(|v| {
            v.as_str()
                .map(str::to_owned)
                .unwrap_or_else(|| v.to_string())
        })
        .unwrap_or_default()
}

#[cfg(test)]
mod tests {
    use super::{build_sheet_rename_plan_with_mapping, has_conflicting_sheet_name, read_workbook};
    use std::{
        collections::{BTreeMap, BTreeSet},
        path::PathBuf,
    };

    #[test]
    fn sheet_name_conflicts_are_case_insensitive() {
        let current = BTreeSet::from(["Data".to_owned(), "Summary".to_owned()]);
        assert!(has_conflicting_sheet_name(&current, "Summary", "data"));
    }

    #[test]
    fn sheet_can_keep_its_own_name_or_change_only_its_case() {
        let current = BTreeSet::from(["Data".to_owned(), "Summary".to_owned()]);
        assert!(!has_conflicting_sheet_name(&current, "Data", "data"));
        assert!(!has_conflicting_sheet_name(&current, "Data", "Archive"));
    }

    #[test]
    fn prior_rename_target_is_reserved_for_later_operations() {
        let current = BTreeSet::from(["Summary".to_owned(), "Archive".to_owned()]);
        assert!(has_conflicting_sheet_name(&current, "Summary", "archive"));
    }

    #[test]
    fn plan_skips_case_insensitive_collision_with_another_sheet() -> anyhow::Result<()> {
        let input = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/sheet_rename/input.xlsx");
        let dump = read_workbook(&input)?;
        let first_name = dump
            .sheets
            .first()
            .expect("fixture has at least two sheets")
            .name
            .clone();
        let second_name = dump
            .sheets
            .get(1)
            .expect("fixture has at least two sheets")
            .name
            .clone();
        let mapping = BTreeMap::from([(second_name, first_name.to_lowercase())]);

        let report = build_sheet_rename_plan_with_mapping(
            &mapping,
            &input,
            &input.with_file_name("candidate.xlsx"),
        )?;

        assert_eq!(report.renamed, 0);
        assert_eq!(report.skipped, 1);
        assert!(report.plan.is_none());
        Ok(())
    }
}
