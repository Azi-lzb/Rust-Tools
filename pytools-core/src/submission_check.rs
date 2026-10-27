//! Filename-based submission completeness report (conversion menu 4-8).

use anyhow::{Context, Result};
use chrono::Local;
use excel_reader::{read_value_columns_with_iso_dates, read_workbook, CellValue, SheetDump};
use excel_writer::{Format, Workbook};
use serde::Serialize;
use std::{
    collections::HashMap,
    fs::{self, OpenOptions},
    path::{Path, PathBuf},
    time::{Instant, SystemTime, UNIX_EPOCH},
};

#[derive(Clone, Debug, Serialize)]
pub struct SubmissionCheckReport {
    pub output: PathBuf,
    pub rules: usize,
    pub files: usize,
    pub submitted: usize,
    pub missing: usize,
    pub unmatched: usize,
    pub multi: usize,
}

#[derive(Clone, Debug, Default, Serialize)]
pub struct SubmissionCheckProgress {
    pub total_rules: usize,
    pub processed_rules: usize,
    pub total_files: usize,
    pub processed_files: usize,
    pub submitted_rules: usize,
    pub missing_rules: usize,
    pub unmatched_files: usize,
    pub multi_files: usize,
}

#[derive(Clone, Debug, Serialize)]
pub struct SubmissionCheckProgressEvent {
    pub sequence: usize,
    pub timestamp: String,
    pub product: String,
    pub feature_id: String,
    pub function: String,
    pub phase: String,
    pub status: String,
    pub user_message: String,
    pub progress: SubmissionCheckProgress,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stage_elapsed_ms: Option<u128>,
    pub total_elapsed_ms: u128,
    pub components: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub actual_engine: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_type: Option<String>,
}

struct SubmissionCheckProgressEmitter<'a> {
    started: Instant,
    sequence: usize,
    progress: SubmissionCheckProgress,
    callback: &'a mut dyn FnMut(SubmissionCheckProgressEvent),
}

impl SubmissionCheckProgressEmitter<'_> {
    fn emit(
        &mut self,
        phase: &str,
        status: &str,
        user_message: impl Into<String>,
        stage_started: Option<Instant>,
        components: &[&str],
        actual_engine: Option<&str>,
        error_type: Option<&str>,
    ) {
        self.sequence += 1;
        let mut recorded_components = Vec::with_capacity(components.len() + 1);
        recorded_components.push("Rust".to_owned());
        recorded_components.extend(components.iter().map(|component| (*component).to_owned()));
        (self.callback)(SubmissionCheckProgressEvent {
            sequence: self.sequence,
            timestamp: Local::now().to_rfc3339(),
            product: "pytools".to_owned(),
            feature_id: "PT-4-8".to_owned(),
            function: "submission-check".to_owned(),
            phase: phase.to_owned(),
            status: status.to_owned(),
            user_message: user_message.into(),
            progress: self.progress.clone(),
            stage_elapsed_ms: stage_started.map(|started| started.elapsed().as_millis()),
            total_elapsed_ms: self.started.elapsed().as_millis(),
            components: recorded_components,
            actual_engine: actual_engine.map(str::to_owned),
            error_type: error_type.map(str::to_owned),
        });
    }
}
#[derive(Clone)]
struct Rule {
    row: usize,
    enabled: String,
    date: String,
    area: String,
    institution: String,
    extra: String,
    remark: String,
    groups: Vec<Vec<String>>,
}

pub fn run_submission_check(
    config: &Path,
    inputs: &[PathBuf],
    output_dir: &Path,
) -> Result<SubmissionCheckReport> {
    run_submission_check_with_progress(config, inputs, output_dir, |_| {})
}

pub fn run_submission_check_with_progress<F>(
    config: &Path,
    inputs: &[PathBuf],
    output_dir: &Path,
    mut on_event: F,
) -> Result<SubmissionCheckReport>
where
    F: FnMut(SubmissionCheckProgressEvent),
{
    let mut emitter = SubmissionCheckProgressEmitter {
        started: Instant::now(),
        sequence: 0,
        progress: SubmissionCheckProgress::default(),
        callback: &mut on_event,
    };

    let stage_started = Instant::now();
    emitter.emit(
        "configuration_load",
        "started",
        "正在读取提交检查配置。",
        Some(stage_started),
        &["excel-reader", "calamine"],
        Some("calamine"),
        None,
    );
    let rules = match load_rules(config) {
        Ok(rules) => rules,
        Err(error) => {
            emitter.emit(
                "configuration_load",
                "failed",
                "读取提交检查配置失败。",
                Some(stage_started),
                &["excel-reader", "calamine"],
                Some("calamine"),
                Some("config_read_failed"),
            );
            return Err(error);
        }
    };
    if rules.is_empty() {
        let error = anyhow::anyhow!("提交检查配置 没有启用的检查项。");
        emitter.emit(
            "configuration_load",
            "failed",
            "提交检查配置没有启用的检查项。",
            Some(stage_started),
            &["excel-reader", "calamine"],
            Some("calamine"),
            Some("no_enabled_rules"),
        );
        return Err(error);
    }
    emitter.progress.total_rules = rules.len();
    emitter.emit(
        "configuration_load",
        "completed",
        format!("已读取 {} 条启用检查项。", rules.len()),
        Some(stage_started),
        &["excel-reader", "calamine"],
        Some("calamine"),
        None,
    );

    let stage_started = Instant::now();
    emitter.emit(
        "input_scan",
        "started",
        "正在扫描所选文件及文件夹。",
        Some(stage_started),
        &["std::fs"],
        Some("Rust filesystem"),
        None,
    );
    let files = match collect_files(inputs) {
        Ok(files) => files,
        Err(error) => {
            emitter.emit(
                "input_scan",
                "failed",
                "扫描所选文件夹失败，提交检查已停止。",
                Some(stage_started),
                &["std::fs"],
                Some("Rust filesystem"),
                Some("input_scan_failed"),
            );
            return Err(error);
        }
    };
    emitter.progress.total_files = files.len();
    emitter.emit(
        "input_scan",
        "completed",
        format!("发现 {} 个待检查文件。", files.len()),
        Some(stage_started),
        &["std::fs"],
        Some("Rust filesystem"),
        None,
    );

    let stage_started = Instant::now();
    emitter.emit(
        "filename_match",
        "started",
        "正在按文件名匹配检查项。",
        Some(stage_started),
        &[],
        Some("Rust"),
        None,
    );
    let mut hits = HashMap::<PathBuf, Vec<usize>>::new();
    for file in &files {
        hits.insert(file.clone(), Vec::new());
    }
    let mut rows = Vec::new();
    let progress_step = rules.len().div_ceil(20).max(1);
    for (rule_index, rule) in rules.iter().enumerate() {
        let matched = files
            .iter()
            .filter(|path| {
                matches_rule(
                    path.file_name()
                        .unwrap_or_default()
                        .to_string_lossy()
                        .as_ref(),
                    rule,
                )
            })
            .collect::<Vec<_>>();
        for file in &matched {
            hits.get_mut(*file).expect("known file").push(rule_index);
        }
        rows.push(vec![
            rule.enabled.clone(),
            rule.date.clone(),
            rule.area.clone(),
            rule.institution.clone(),
            rule.extra.clone(),
            matched.len().to_string(),
            matched
                .iter()
                .map(|p| {
                    p.file_name()
                        .unwrap_or_default()
                        .to_string_lossy()
                        .into_owned()
                })
                .collect::<Vec<_>>()
                .join("；"),
            rule.remark.clone(),
            rule.row.to_string(),
        ]);
        emitter.progress.processed_rules = rule_index + 1;
        if emitter.progress.processed_rules % progress_step == 0
            || emitter.progress.processed_rules == rules.len()
        {
            emitter.emit(
                "filename_match",
                "progress",
                format!(
                    "已匹配 {}/{} 条检查项。",
                    emitter.progress.processed_rules,
                    rules.len()
                ),
                Some(stage_started),
                &[],
                Some("Rust"),
                None,
            );
        }
    }
    emitter.progress.processed_files = files.len();
    let missing_rules = rows
        .iter()
        .filter(|row| row[5] == "0")
        .cloned()
        .collect::<Vec<_>>();
    let unmatched = files
        .iter()
        .filter(|path| hits.get(*path).is_some_and(Vec::is_empty))
        .map(|path| {
            vec![
                path.file_name()
                    .unwrap_or_default()
                    .to_string_lossy()
                    .into_owned(),
                path.display().to_string(),
                "未命中任何启用配置".to_owned(),
            ]
        })
        .collect::<Vec<_>>();
    let multi = files
        .iter()
        .filter_map(|path| {
            let matched = hits.get(path)?;
            if matched.len() <= 1 {
                return None;
            }
            Some(vec![
                path.file_name()
                    .unwrap_or_default()
                    .to_string_lossy()
                    .into_owned(),
                path.display().to_string(),
                matched
                    .iter()
                    .map(|rule_index| {
                        let rule = &rules[*rule_index];
                        format!(
                            "第{}行({}|{}|{}|{})",
                            rule.row, rule.date, rule.area, rule.institution, rule.extra
                        )
                    })
                    .collect::<Vec<_>>()
                    .join("；"),
            ])
        })
        .collect::<Vec<_>>();
    let submitted = rows.iter().filter(|row| row[5] != "0").count();
    let missing = rows.len() - submitted;
    emitter.progress.submitted_rules = submitted;
    emitter.progress.missing_rules = missing;
    emitter.progress.unmatched_files = unmatched.len();
    emitter.progress.multi_files = multi.len();
    emitter.emit(
        "filename_match",
        "completed",
        format!(
            "匹配完成：提交 {submitted} 条，缺失 {missing} 条，未识别文件 {} 个，多重命中 {} 个。",
            unmatched.len(),
            multi.len()
        ),
        Some(stage_started),
        &[],
        Some("Rust"),
        None,
    );

    let stage_started = Instant::now();
    emitter.emit(
        "report_write",
        "started",
        "正在生成提交检查工作簿。",
        Some(stage_started),
        &["excel-writer", "rust_xlsxwriter"],
        Some("rust_xlsxwriter"),
        None,
    );
    let output = match write_report(output_dir, &rows, &missing_rules, &unmatched, &multi) {
        Ok(output) => output,
        Err(error) => {
            emitter.emit(
                "report_write",
                "failed",
                "生成提交检查工作簿失败。",
                Some(stage_started),
                &["excel-writer", "rust_xlsxwriter"],
                Some("rust_xlsxwriter"),
                Some("report_write_failed"),
            );
            return Err(error);
        }
    };
    emitter.emit(
        "report_write",
        "completed",
        "已生成提交检查工作簿。",
        Some(stage_started),
        &["excel-writer", "rust_xlsxwriter"],
        Some("rust_xlsxwriter"),
        None,
    );
    emitter.emit(
        "complete",
        "completed",
        "提交检查完成。",
        None,
        &["Rust"],
        None,
        None,
    );
    Ok(SubmissionCheckReport {
        output,
        rules: rules.len(),
        files: files.len(),
        submitted,
        missing,
        unmatched: unmatched.len(),
        multi: multi.len(),
    })
}

fn write_report(
    output_dir: &Path,
    rows: &[Vec<String>],
    missing_rules: &[Vec<String>],
    unmatched: &[Vec<String>],
    multi: &[Vec<String>],
) -> Result<PathBuf> {
    fs::create_dir_all(output_dir)?;
    let mut workbook = Workbook::new();
    write_table(
        &mut workbook,
        "检查结果",
        &[
            "是否启用",
            "日期关键词",
            "县区关键词",
            "机构关键词",
            "备用关键词",
            "命中次数",
            "命中的文件名",
            "备注",
            "配置行",
        ],
        rows,
        &[5, 8],
    )?;
    write_table(
        &mut workbook,
        "缺失清单",
        &[
            "是否启用",
            "日期关键词",
            "县区关键词",
            "机构关键词",
            "备用关键词",
            "命中次数",
            "命中的文件名",
            "备注",
            "配置行",
        ],
        missing_rules,
        &[5, 8],
    )?;
    write_table(
        &mut workbook,
        "未识别文件",
        &["文件名", "源文件路径", "备注"],
        unmatched,
        &[],
    )?;
    write_table(
        &mut workbook,
        "多重命中文件",
        &["文件名", "源文件路径", "命中配置"],
        multi,
        &[],
    )?;
    let candidate = reserve_report_candidate(output_dir)?;
    workbook
        .save(&candidate.0)
        .with_context(|| format!("save submission check candidate {}", candidate.0.display()))?;
    publish_report_candidate(&candidate.0, output_dir)
}

fn load_rules(path: &Path) -> Result<Vec<Rule>> {
    let dump = read_workbook(path)?;
    let sheet = dump
        .sheets
        .iter()
        .find(|sheet| sheet.name == "提交检查配置")
        .context("提交检查配置 缺少工作表")?;
    let headers = header_index(sheet);
    let date_keyword_overrides = date_keyword_overrides(path, sheet, &headers)?;
    let mut rules = Vec::new();
    for row in 2..=used_rows(sheet) {
        let enabled = field(sheet, &headers, row, "是否启用");
        if !matches!(
            enabled.to_lowercase().as_str(),
            "1" | "y" | "yes" | "true" | "是" | "启用"
        ) {
            continue;
        }
        let date_address = headers
            .get("日期关键词")
            .map(|column| format!("{}{}", column_letters(*column), row));
        let date = date_address
            .as_ref()
            .and_then(|address| date_keyword_overrides.get(address))
            .cloned()
            .unwrap_or_else(|| field(sheet, &headers, row, "日期关键词"));
        let area = field(sheet, &headers, row, "县区关键词");
        let institution = field(sheet, &headers, row, "机构关键词");
        let extra = field(sheet, &headers, row, "备用关键词");
        let groups = [&date, &area, &institution, &extra]
            .iter()
            .map(|v| split_tokens(v))
            .collect::<Vec<_>>();
        rules.push(Rule {
            row,
            enabled,
            date,
            area,
            institution,
            extra,
            remark: field(sheet, &headers, row, "备注"),
            groups,
        });
    }
    Ok(rules)
}

fn date_keyword_overrides(
    path: &Path,
    sheet: &SheetDump,
    headers: &HashMap<String, usize>,
) -> Result<HashMap<String, String>> {
    let Some(max_column) = headers.get("日期关键词").copied() else {
        return Ok(HashMap::new());
    };
    let has_serial_dates = sheet.values.iter().any(|(address, value)| {
        parse_a1(address).is_some_and(|(_, column)| column == max_column)
            && value.get("type").and_then(serde_json::Value::as_str) == Some("datetime_serial")
    });
    if !has_serial_dates {
        return Ok(HashMap::new());
    }
    let date_values = read_value_columns_with_iso_dates(path, max_column)?
        .into_iter()
        .find(|candidate| candidate.name == sheet.name)
        .with_context(|| format!("{} sheet disappeared during date projection", sheet.name))?;
    let mut overrides = HashMap::new();
    for (address, value) in date_values.values {
        if let CellValue::DateTimeIso(iso) = value {
            let text = python_datetime_text(&iso)
                .with_context(|| format!("format date keyword cell {}!{address}", sheet.name))?;
            overrides.insert(address, text);
        }
    }
    Ok(overrides)
}

fn python_datetime_text(iso: &str) -> Option<String> {
    let value = chrono::NaiveDateTime::parse_from_str(iso, "%Y-%m-%dT%H:%M:%S%.f").ok()?;
    let whole_seconds = value.format("%Y-%m-%d %H:%M:%S").to_string();
    let microseconds = value.and_utc().timestamp_subsec_micros();
    if microseconds == 0 {
        Some(whole_seconds)
    } else {
        Some(format!("{whole_seconds}.{microseconds:06}"))
    }
}
fn matches_rule(file_name: &str, rule: &Rule) -> bool {
    let name = file_name.to_lowercase();
    let mut has = false;
    for group in &rule.groups {
        if group.is_empty() {
            continue;
        }
        has = true;
        if !group
            .iter()
            .any(|token| name.contains(&token.to_lowercase()))
        {
            return false;
        }
    }
    has
}
fn split_tokens(value: &str) -> Vec<String> {
    value
        .replace('；', ";")
        .replace(',', ";")
        .split(';')
        .map(str::trim)
        .filter(|v| !v.is_empty())
        .map(str::to_owned)
        .collect()
}
fn collect_files(inputs: &[PathBuf]) -> Result<Vec<PathBuf>> {
    let mut out = Vec::new();
    let mut seen = std::collections::BTreeSet::new();
    for path in inputs {
        if path.is_dir() {
            walk(path, &mut out, &mut std::collections::BTreeSet::new())?;
        } else if path.is_file() {
            out.push(path.clone());
        }
    }
    out.retain(|path| {
        !path
            .file_name()
            .unwrap_or_default()
            .to_string_lossy()
            .starts_with("~$")
    });
    out.retain(|path| seen.insert(path.canonicalize().unwrap_or_else(|_| path.clone())));
    Ok(out)
}
fn walk(
    root: &Path,
    out: &mut Vec<PathBuf>,
    active_directories: &mut std::collections::BTreeSet<PathBuf>,
) -> Result<()> {
    let root_identity = fs::canonicalize(root)
        .with_context(|| format!("resolve submission-check directory {}", root.display()))?;
    if !active_directories.insert(root_identity.clone()) {
        return Ok(());
    }
    let entries = fs::read_dir(root)
        .with_context(|| format!("scan submission-check directory {}", root.display()))?;
    for entry in entries {
        let entry = entry.with_context(|| {
            format!(
                "read submission-check directory entry in {}",
                root.display()
            )
        })?;
        let path = entry.path();
        let metadata = match fs::metadata(&path) {
            Ok(metadata) => metadata,
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => continue,
            Err(error) => {
                return Err(error)
                    .with_context(|| format!("inspect submission-check entry {}", path.display()));
            }
        };
        if metadata.is_dir() {
            walk(&path, out, active_directories)?;
        } else if metadata.is_file() {
            out.push(path);
        }
    }
    active_directories.remove(&root_identity);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::{is_report_target_conflict, publish_report_candidate_to_path, walk};
    use anyhow::Result;
    use rust_xlsxwriter::Workbook;
    use std::{
        fs,
        path::PathBuf,
        time::{SystemTime, UNIX_EPOCH},
    };

    #[cfg(windows)]
    use std::process::Command;

    #[cfg(windows)]
    struct TemporaryDirectory(PathBuf);

    #[cfg(windows)]
    impl TemporaryDirectory {
        fn new(name: &str) -> Result<Self> {
            let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
            let path = std::env::temp_dir().join(format!(
                "pytools-submission-{name}-{}-{nonce}",
                std::process::id()
            ));
            fs::create_dir_all(&path)?;
            Ok(Self(path))
        }
    }

    #[cfg(windows)]
    impl Drop for TemporaryDirectory {
        fn drop(&mut self) {
            let _ = fs::remove_dir_all(&self.0);
        }
    }

    #[test]
    fn recursive_scan_reports_read_dir_errors() -> Result<()> {
        let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let path = std::env::temp_dir().join(format!(
            "pytools-submission-scan-file-{}-{nonce}.tmp",
            std::process::id()
        ));
        fs::write(&path, b"not a directory")?;
        let mut files = Vec::<PathBuf>::new();
        let result = walk(&path, &mut files, &mut std::collections::BTreeSet::new());
        fs::remove_file(&path)?;

        let error = result.expect_err("scanning a file as a directory must fail");
        assert!(error
            .to_string()
            .contains("scan submission-check directory"));
        assert!(files.is_empty());
        Ok(())
    }

    #[test]
    fn report_publication_does_not_overwrite_an_existing_target() -> Result<()> {
        let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let dir = std::env::temp_dir().join(format!(
            "pytools-submission-publish-{}-{nonce}",
            std::process::id()
        ));
        fs::create_dir_all(&dir)?;
        let candidate = dir.join("candidate.xlsx");
        let target = dir.join("existing.xlsx");
        let mut workbook = Workbook::new();
        workbook.add_worksheet().write_string(0, 0, "candidate")?;
        workbook.save(&candidate)?;
        fs::write(&target, b"keep existing report")?;

        let error = publish_report_candidate_to_path(&candidate, &target)
            .expect_err("an existing report path must not be replaced");
        assert!(format!("{error:#}").contains("target already exists"));
        assert_eq!(fs::read(&target)?, b"keep existing report");
        fs::remove_dir_all(dir)?;
        Ok(())
    }

    #[test]
    fn report_retry_only_handles_target_name_conflicts() {
        assert!(is_report_target_conflict(&anyhow::anyhow!(
            "target already exists; refusing to overwrite"
        )));
        assert!(is_report_target_conflict(&anyhow::anyhow!(
            std::io::Error::from(std::io::ErrorKind::AlreadyExists)
        )));
        assert!(!is_report_target_conflict(&anyhow::anyhow!(
            "target directory sync failed"
        )));
    }

    #[cfg(windows)]
    #[test]
    fn recursive_scan_follows_directory_junctions_and_stops_cycles() -> Result<()> {
        let root = TemporaryDirectory::new("junction-cycle")?;
        let scan_root = root.0.join("root");
        let child = scan_root.join("child");
        let linked_target = root.0.join("outside-target");
        fs::create_dir_all(&child)?;
        fs::create_dir_all(&linked_target)?;
        let root_file = scan_root.join("root.txt");
        let child_file = child.join("child.txt");
        let linked_file = linked_target.join("outside.txt");
        fs::write(&root_file, b"root")?;
        fs::write(&child_file, b"child")?;
        fs::write(&linked_file, b"outside")?;

        create_junction(&child.join("back-to-root"), &scan_root)?;
        create_junction(&child.join("linked-target"), &linked_target)?;

        let mut files = Vec::new();
        walk(
            &scan_root,
            &mut files,
            &mut std::collections::BTreeSet::new(),
        )?;
        files.sort();
        assert_eq!(
            files,
            vec![
                child_file,
                child.join("linked-target/outside.txt"),
                root_file
            ]
        );
        Ok(())
    }

    #[cfg(windows)]
    fn create_junction(link: &std::path::Path, target: &std::path::Path) -> Result<()> {
        let link_arg = link.to_string_lossy().replace('\'', "''");
        let target_arg = target.to_string_lossy().replace('\'', "''");
        let script = format!(
            "New-Item -ItemType Junction -Path '{link_arg}' -Target '{target_arg}' -ErrorAction Stop | Out-Null"
        );
        let output = Command::new("powershell.exe")
            .args(["-NoProfile", "-NonInteractive", "-Command", script.as_str()])
            .output()?;
        anyhow::ensure!(
            output.status.success(),
            "could not create test directory junction: {}{}",
            String::from_utf8_lossy(&output.stdout),
            String::from_utf8_lossy(&output.stderr)
        );
        Ok(())
    }
}
fn write_table(
    workbook: &mut Workbook,
    name: &str,
    headers: &[&str],
    rows: &[Vec<String>],
    numeric_columns: &[usize],
) -> Result<()> {
    let sheet = workbook.add_worksheet();
    sheet.set_name(name)?;
    let header_format = Format::new().set_bold();
    for (col, header) in headers.iter().enumerate() {
        sheet.write_string_with_format(0, col as u16, *header, &header_format)?;
    }
    for (row, values) in rows.iter().enumerate() {
        for (col, value) in values.iter().enumerate() {
            if numeric_columns.contains(&col) {
                let number = value.parse::<f64>().with_context(|| {
                    format!("submission report column {} must be numeric", col + 1)
                })?;
                sheet.write_number((row + 1) as u32, col as u16, number)?;
            } else {
                sheet.write_string((row + 1) as u32, col as u16, value)?;
            }
        }
    }
    Ok(())
}
fn unique_output(dir: &Path) -> Result<PathBuf> {
    let stamp = Local::now().format("%Y%m%d_%H%M%S");
    for index in 0..10000 {
        let suffix = if index == 0 {
            String::new()
        } else {
            format!("_{index}")
        };
        let path = dir.join(format!("{stamp}_提交缺失检查{suffix}.xlsx"));
        if !path.exists() {
            return Ok(path);
        }
    }
    anyhow::bail!("cannot reserve submission report path")
}

struct ReportCandidate(PathBuf);

impl Drop for ReportCandidate {
    fn drop(&mut self) {
        let _ = fs::remove_file(&self.0);
    }
}

fn reserve_report_candidate(dir: &Path) -> Result<ReportCandidate> {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .context("system clock is before Unix epoch")?
        .as_nanos();
    for index in 0..10_000 {
        let path = dir.join(format!(
            ".pytools-submission-check-{}-{nonce}-{index}.xlsx",
            std::process::id()
        ));
        match OpenOptions::new().write(true).create_new(true).open(&path) {
            Ok(file) => {
                drop(file);
                return Ok(ReportCandidate(path));
            }
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => {
                return Err(error)
                    .with_context(|| format!("reserve report candidate {}", path.display()))
            }
        }
    }
    anyhow::bail!(
        "cannot reserve a unique submission report candidate in {}",
        dir.display()
    )
}

fn publish_report_candidate(candidate: &Path, output_dir: &Path) -> Result<PathBuf> {
    for _ in 0..10_000 {
        let output = unique_output(output_dir)?;
        match publish_report_candidate_to_path(candidate, &output) {
            Ok(()) => return Ok(output),
            Err(error) if is_report_target_conflict(&error) => continue,
            Err(error) => {
                return Err(error).with_context(|| {
                    format!(
                        "publish submission report {} without overwrite",
                        output.display()
                    )
                });
            }
        }
    }
    anyhow::bail!(
        "cannot reserve a unique submission report target in {}",
        output_dir.display()
    )
}

fn publish_report_candidate_to_path(candidate: &Path, output: &Path) -> Result<()> {
    crate::file_commit::create_target_from_candidate(candidate, output)
        .map(|_| ())
        .with_context(|| {
            format!(
                "publish submission report {} without overwrite",
                output.display()
            )
        })
}

fn is_report_target_conflict(error: &anyhow::Error) -> bool {
    error.chain().any(|cause| {
        cause.to_string().contains("target already exists")
            || cause
                .downcast_ref::<std::io::Error>()
                .is_some_and(|error| error.kind() == std::io::ErrorKind::AlreadyExists)
    })
}
fn header_index(sheet: &SheetDump) -> HashMap<String, usize> {
    let mut out = HashMap::new();
    for (address, value) in &sheet.values {
        if parse_a1(address).is_some_and(|(row, _)| row == 1) {
            out.insert(value_text(value), parse_a1(address).expect("checked").1);
        }
    }
    out
}
fn field(sheet: &SheetDump, headers: &HashMap<String, usize>, row: usize, name: &str) -> String {
    headers
        .get(name)
        .map(|col| {
            value_text(
                sheet
                    .values
                    .get(&format!("{}{}", column_letters(*col), row))
                    .unwrap_or(&serde_json::Value::Null),
            )
        })
        .unwrap_or_default()
}
fn value_text(value: &serde_json::Value) -> String {
    match value.get("type").and_then(serde_json::Value::as_str) {
        Some("string" | "datetime_iso" | "duration_iso" | "error") => value
            .get("value")
            .and_then(serde_json::Value::as_str)
            .unwrap_or_default()
            .trim()
            .into(),
        Some("number" | "integer") => value
            .get("value")
            .map(ToString::to_string)
            .unwrap_or_default(),
        Some("boolean") => value
            .get("value")
            .and_then(serde_json::Value::as_bool)
            .map(|v| if v { "True" } else { "False" }.into())
            .unwrap_or_default(),
        _ => String::new(),
    }
}
fn used_rows(sheet: &SheetDump) -> usize {
    sheet
        .values
        .keys()
        .filter_map(|address| parse_a1(address).map(|value| value.0))
        .max()
        .unwrap_or(1)
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
    let mut value = String::new();
    while col > 0 {
        value.insert(0, (b'A' + ((col - 1) % 26) as u8) as char);
        col = (col - 1) / 26;
    }
    value
}
