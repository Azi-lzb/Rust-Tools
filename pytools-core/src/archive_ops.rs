//! Date/type archive preview and copy workflows from the selected pytools scope.

use anyhow::{bail, Context, Result};
use chrono::Local;
use excel_reader::read_workbook;
use excel_writer::{Format, FormatAlign, FormatBorder, Workbook};
use serde::Serialize;
use std::{
    collections::{BTreeSet, HashSet},
    fs,
    path::{Path, PathBuf},
};

#[derive(Clone, Debug, Serialize)]
pub struct ArchiveRow {
    pub source_path: String,
    pub file_name: String,
    pub mode: String,
    pub category: String,
    pub target_dir: String,
    pub planned_target_path: String,
    pub target_path: String,
    pub status: String,
    pub note: String,
}

#[derive(Clone, Debug, Serialize)]
pub struct ArchiveResult {
    pub files: usize,
    pub invalid: usize,
    pub copied: usize,
    pub skipped: usize,
    pub failed: usize,
    pub output: PathBuf,
    pub rows: Vec<ArchiveRow>,
}

#[derive(Clone, Debug, Default)]
struct TypeConfig {
    rules: Vec<(String, String)>,
    include: BTreeSet<String>,
    exclude: BTreeSet<String>,
}

pub fn archive_by_date_preview(
    paths: &[PathBuf],
    target_root: &Path,
    output_dir: &Path,
    config: Option<&Path>,
) -> Result<ArchiveResult> {
    archive(
        paths,
        target_root,
        output_dir,
        config,
        ArchiveMode::Date,
        false,
    )
}
pub fn archive_by_type_preview(
    paths: &[PathBuf],
    target_root: &Path,
    output_dir: &Path,
    config: &Path,
) -> Result<ArchiveResult> {
    archive(
        paths,
        target_root,
        output_dir,
        Some(config),
        ArchiveMode::Type,
        false,
    )
}
pub fn archive_by_date_copy(
    paths: &[PathBuf],
    target_root: &Path,
    output_dir: &Path,
    config: Option<&Path>,
) -> Result<ArchiveResult> {
    archive(
        paths,
        target_root,
        output_dir,
        config,
        ArchiveMode::Date,
        true,
    )
}
pub fn archive_by_type_copy(
    paths: &[PathBuf],
    target_root: &Path,
    output_dir: &Path,
    config: &Path,
) -> Result<ArchiveResult> {
    archive(
        paths,
        target_root,
        output_dir,
        Some(config),
        ArchiveMode::Type,
        true,
    )
}

pub fn execute_archive_plan(plan_path: &Path, output_dir: &Path) -> Result<ArchiveResult> {
    let dump = read_workbook(plan_path)
        .with_context(|| format!("read archive plan {}", plan_path.display()))?;
    let sheet = dump
        .sheets
        .iter()
        .find(|sheet| sheet.name == "归档计划")
        .context("archive plan sheet '归档计划' is missing")?;
    let header = header_index(sheet);
    for name in ["源文件路径", "目标文件路径", "状态"] {
        if !header.contains_key(name) {
            bail!("archive plan is missing required column {name}");
        }
    }
    let max_row = sheet
        .values
        .keys()
        .filter_map(|addr| parse_a1(addr).map(|p| p.0))
        .max()
        .unwrap_or(1);
    let (source_col, target_col, status_col) =
        (header["源文件路径"], header["目标文件路径"], header["状态"]);
    let mut rows = Vec::new();
    for row in 2..=max_row {
        let source = get_cell_text(sheet, row, source_col);
        let target = get_cell_text(sheet, row, target_col);
        let status = get_cell_text(sheet, row, status_col);
        rows.push(ArchiveRow {
            source_path: source.clone(),
            file_name: Path::new(&source)
                .file_name()
                .map(|s| s.to_string_lossy().into_owned())
                .unwrap_or_default(),
            mode: "计划".into(),
            category: String::new(),
            target_dir: Path::new(&target)
                .parent()
                .map(|p| p.to_string_lossy().into_owned())
                .unwrap_or_default(),
            planned_target_path: target.clone(),
            target_path: target,
            status,
            note: String::new(),
        });
    }
    execute_rows(rows, output_dir)
}

#[derive(Clone, Copy)]
enum ArchiveMode {
    Date,
    Type,
}
fn archive(
    paths: &[PathBuf],
    target_root: &Path,
    output_dir: &Path,
    config: Option<&Path>,
    mode: ArchiveMode,
    execute: bool,
) -> Result<ArchiveResult> {
    let kind = match mode {
        ArchiveMode::Date => "按日期",
        ArchiveMode::Type => "按类型",
    };
    let type_config = if matches!(mode, ArchiveMode::Type) {
        load_type_config(config.context("type archive requires --config")?)?
    } else {
        config
            .map(load_type_config)
            .transpose()?
            .unwrap_or_default()
    };
    let (files, invalid) = collect_files(paths)?;
    let mut planned = HashSet::new();
    let mut rows = Vec::new();
    for path in invalid {
        rows.push(ArchiveRow {
            source_path: path.display().to_string(),
            file_name: path
                .file_name()
                .map(|s| s.to_string_lossy().into_owned())
                .unwrap_or_default(),
            mode: kind.into(),
            category: String::new(),
            target_dir: String::new(),
            planned_target_path: String::new(),
            target_path: String::new(),
            status: "跳过：非文件".into(),
            note: String::new(),
        });
    }
    for path in &files {
        let ext = path
            .extension()
            .map(|s| format!(".{}", s.to_string_lossy().to_lowercase()))
            .unwrap_or_default();
        let status = if type_config.exclude.contains(&ext) {
            "跳过：后缀被排除"
        } else if !type_config.include.is_empty() && !type_config.include.contains(&ext) {
            "跳过：后缀不在归档范围"
        } else {
            ""
        };
        if !status.is_empty() {
            rows.push(ArchiveRow {
                source_path: path.display().to_string(),
                file_name: path
                    .file_name()
                    .map(|s| s.to_string_lossy().into_owned())
                    .unwrap_or_default(),
                mode: kind.into(),
                category: String::new(),
                target_dir: String::new(),
                planned_target_path: String::new(),
                target_path: String::new(),
                status: status.into(),
                note: String::new(),
            });
            continue;
        }
        let category = match mode {
            ArchiveMode::Date => {
                archive_month(&path.file_name().unwrap_or_default().to_string_lossy())
                    .unwrap_or_else(|| "未识别日期".into())
            }
            ArchiveMode::Type => type_config
                .rules
                .iter()
                .find(|(keyword, _)| {
                    path.file_name()
                        .unwrap_or_default()
                        .to_string_lossy()
                        .to_lowercase()
                        .contains(&keyword.to_lowercase())
                })
                .map(|(_, folder)| folder.clone())
                .unwrap_or_else(|| "未识别类型".into()),
        };
        let status = if category.starts_with("未识别") {
            category.clone()
        } else {
            "待归档".into()
        };
        let safe = sanitize_folder(&category);
        let target_dir = target_root.join(&safe);
        let raw_target = target_dir.join(path.file_name().unwrap_or_default());
        let target = next_target(&raw_target, &mut planned);
        let note = if target.file_name() != raw_target.file_name() {
            "目标同名：将自动后缀".into()
        } else {
            String::new()
        };
        rows.push(ArchiveRow {
            source_path: path.display().to_string(),
            file_name: path
                .file_name()
                .map(|s| s.to_string_lossy().into_owned())
                .unwrap_or_default(),
            mode: kind.into(),
            category: safe,
            target_dir: target_dir.display().to_string(),
            planned_target_path: target.display().to_string(),
            target_path: target.display().to_string(),
            status,
            note,
        });
    }
    if execute {
        execute_rows(rows, output_dir)
    } else {
        write_archive_report(rows, output_dir, "归档计划")
    }
}

fn execute_rows(mut rows: Vec<ArchiveRow>, output_dir: &Path) -> Result<ArchiveResult> {
    let mut copied = 0;
    let mut skipped = 0;
    let mut failed = 0;
    let mut planned = HashSet::new();
    for row in &mut rows {
        if !matches!(row.status.as_str(), "待归档" | "未识别日期" | "未识别类型") {
            row.note = row.status.clone();
            row.status = "跳过：计划状态不可执行".into();
            skipped += 1;
            continue;
        }
        let source = Path::new(&row.source_path);
        let raw = Path::new(&row.target_path);
        if row.target_path.is_empty() {
            row.note.clear();
            row.status = "跳过：目标路径为空".into();
            skipped += 1;
            continue;
        }
        if !source.is_file() {
            row.note.clear();
            row.status = "跳过：源文件不存在".into();
            skipped += 1;
            continue;
        }
        row.note.clear();
        let target = next_target(raw, &mut planned);
        if row.planned_target_path.is_empty() {
            row.planned_target_path = row.target_path.clone();
        }
        match copy_with_metadata(source, &target) {
            Ok(actual) => {
                if actual != raw {
                    row.note = "目标同名：已自动后缀".into();
                }
                row.target_path = actual.display().to_string();
                row.status = "已复制".into();
                copied += 1;
            }
            Err(error) => {
                row.status = format!("失败：{error}");
                failed += 1;
            }
        }
    }
    let result = write_archive_report(rows, output_dir, "归档执行结果")?;
    Ok(ArchiveResult {
        copied,
        skipped,
        failed,
        ..result
    })
}

fn write_archive_report(
    rows: Vec<ArchiveRow>,
    output_dir: &Path,
    sheet_name: &str,
) -> Result<ArchiveResult> {
    fs::create_dir_all(output_dir)?;
    let stamp = Local::now().format("%Y%m%d_%H%M%S");
    let mut output = output_dir.join(format!("{stamp}_{sheet_name}.xlsx"));
    let mut i = 1;
    while output.exists() {
        output = output_dir.join(format!("{stamp}_{sheet_name}_{i}.xlsx"));
        i += 1;
    }
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name(sheet_name)?;
    sheet.set_margins(0.75, 0.75, 1.0, 1.0, 0.5, 0.5);
    let header_format = Format::new()
        .set_bold()
        .set_border(FormatBorder::Thin)
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::Top);
    let headers = if sheet_name == "归档计划" {
        vec![
            "源文件路径",
            "文件名",
            "分类方式",
            "分类结果",
            "目标文件夹",
            "目标文件路径",
            "状态",
            "备注",
        ]
    } else {
        vec![
            "源文件路径",
            "原计划目标路径",
            "实际目标路径",
            "执行状态",
            "备注",
        ]
    };
    for (c, value) in headers.iter().enumerate() {
        sheet.write_string_with_format(0, c as u16, *value, &header_format)?;
    }
    for (r, row) in rows.iter().enumerate() {
        let values = if sheet_name == "归档计划" {
            vec![
                row.source_path.clone(),
                row.file_name.clone(),
                row.mode.clone(),
                row.category.clone(),
                row.target_dir.clone(),
                row.target_path.clone(),
                row.status.clone(),
                row.note.clone(),
            ]
        } else {
            vec![
                row.source_path.clone(),
                row.planned_target_path.clone(),
                if row.status == "已复制" {
                    row.target_path.clone()
                } else {
                    String::new()
                },
                row.status.clone(),
                row.note.clone(),
            ]
        };
        for (c, value) in values.iter().enumerate() {
            sheet.write_string((r + 1) as u32, c as u16, value)?;
        }
    }
    workbook
        .save(&output)
        .with_context(|| format!("write archive report {}", output.display()))?;
    let files = rows
        .iter()
        .filter(|r| Path::new(&r.source_path).is_file())
        .count();
    let invalid = rows.len().saturating_sub(files);
    let copied = rows.iter().filter(|r| r.status == "已复制").count();
    let skipped = rows.iter().filter(|r| r.status.starts_with("跳过")).count();
    let failed = rows.iter().filter(|r| r.status.starts_with("失败")).count();
    Ok(ArchiveResult {
        files,
        invalid,
        copied,
        skipped,
        failed,
        output,
        rows,
    })
}

fn load_type_config(path: &Path) -> Result<TypeConfig> {
    let wb = read_workbook(path)?;
    let Some(sheet) = wb.sheets.iter().find(|s| s.name == "归档类型配置") else {
        return Ok(TypeConfig::default());
    };
    let headers = header_index(sheet);
    let mut out = TypeConfig::default();
    for row in 2..=sheet
        .values
        .keys()
        .filter_map(|a| parse_a1(a).map(|p| p.0))
        .max()
        .unwrap_or(1)
    {
        let enabled = header(&headers, &["是否启用"])
            .map(|c| get_cell_text(sheet, row, c))
            .unwrap_or_else(|| get_cell_text(sheet, row, 1));
        if !truthy(&enabled) {
            continue;
        }
        let keyword = header(&headers, &["匹配关键词"])
            .map(|c| get_cell_text(sheet, row, c))
            .unwrap_or_default();
        let folder = header(&headers, &["归档文件夹"])
            .map(|c| get_cell_text(sheet, row, c))
            .unwrap_or_default();
        if !keyword.trim().is_empty() {
            out.rules
                .push((keyword.trim().into(), sanitize_folder(&folder)));
        }
    }
    let include = parse_extensions(&cell_value_text(sheet, "E2"));
    let exclude = parse_extensions(&cell_value_text(sheet, "F2"));
    out.include = include;
    out.exclude = exclude;
    Ok(out)
}
fn collect_files(paths: &[PathBuf]) -> Result<(Vec<PathBuf>, Vec<PathBuf>)> {
    let mut files = Vec::new();
    let mut invalid = Vec::new();
    let mut seen = HashSet::new();
    for path in paths {
        let metadata = match fs::metadata(path) {
            Ok(metadata) => metadata,
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => {
                invalid.push(path.clone());
                continue;
            }
            Err(error) => {
                return Err(error)
                    .with_context(|| format!("inspect archive input {}", path.display()));
            }
        };
        if metadata.is_dir() {
            let mut children = Vec::new();
            walk_files(path, &mut children, &mut HashSet::new())?;
            #[cfg(windows)]
            children.sort_by_cached_key(|child| child.to_string_lossy().to_lowercase());
            #[cfg(not(windows))]
            children.sort();
            for child in children {
                let key = child
                    .canonicalize()
                    .with_context(|| format!("resolve archive input {}", child.display()))?;
                if seen.insert(key) {
                    files.push(child);
                }
            }
        } else if metadata.is_file()
            && !path
                .file_name()
                .unwrap_or_default()
                .to_string_lossy()
                .starts_with("~$")
        {
            let key = path
                .canonicalize()
                .with_context(|| format!("resolve archive input {}", path.display()))?;
            if seen.insert(key) {
                files.push(path.clone());
            }
        } else {
            invalid.push(path.clone());
        }
    }
    Ok((files, invalid))
}
fn walk_files(
    root: &Path,
    out: &mut Vec<PathBuf>,
    active_directories: &mut HashSet<PathBuf>,
) -> Result<()> {
    let root_identity = root
        .canonicalize()
        .with_context(|| format!("resolve archive directory {}", root.display()))?;
    if !active_directories.insert(root_identity.clone()) {
        return Ok(());
    }

    let scan_result = (|| -> Result<()> {
        let entries = fs::read_dir(root)
            .with_context(|| format!("scan archive directory {}", root.display()))?;
        for entry in entries {
            let entry = entry.with_context(|| {
                format!("read an entry in archive directory {}", root.display())
            })?;
            let path = entry.path();
            let metadata = fs::metadata(&path)
                .with_context(|| format!("inspect archive input {}", path.display()))?;
            if metadata.is_dir() {
                walk_files(&path, out, active_directories)?;
            } else if metadata.is_file()
                && !path
                    .file_name()
                    .unwrap_or_default()
                    .to_string_lossy()
                    .starts_with("~$")
            {
                out.push(path);
            }
        }
        Ok(())
    })();
    active_directories.remove(&root_identity);
    scan_result
}
fn archive_month(name: &str) -> Option<String> {
    // Match the Python archive patterns in priority order: separated full
    // date, compact full date, separated year-month, compact year-month.
    let stem = Path::new(name).file_stem()?.to_string_lossy();
    let chars = stem.chars().collect::<Vec<_>>();
    for pattern in 0..4 {
        for start in 0..chars.len() {
            if start > 0 && chars[start - 1].is_ascii_digit() {
                continue;
            }
            let Some((year, month)) = archive_match_at(&chars, start, pattern) else {
                continue;
            };
            return Some(format!("{year}{month:02}"));
        }
    }
    None
}

fn archive_match_at(chars: &[char], start: usize, pattern: usize) -> Option<(String, u32)> {
    if chars.get(start..start + 2)? != ['2', '0'] {
        return None;
    }
    let year_text: String = chars.get(start..start + 4)?.iter().collect();
    if !year_text.bytes().all(|byte| byte.is_ascii_digit()) {
        return None;
    }
    let year = year_text.parse::<u32>().ok()?;
    let (month, month_end) = match pattern {
        0 | 2 => {
            if !matches!(*chars.get(start + 4)?, '-' | '_' | '.' | '年') {
                return None;
            }
            let (month, end) = read_month(chars, start + 5)?;
            (month, end)
        }
        1 | 3 => {
            let (month, end) = read_month(chars, start + 4)?;
            if end != start + 6 {
                return None;
            }
            (month, end)
        }
        _ => return None,
    };
    match pattern {
        0 => {
            if !matches!(*chars.get(month_end)?, '-' | '_' | '.' | '月') {
                return None;
            }
            let (day, end) = read_day(chars, month_end + 1)?;
            let end = if chars.get(end) == Some(&'日') {
                end + 1
            } else {
                end
            };
            if !day_valid(day) || chars.get(end).is_some_and(|c| c.is_ascii_digit()) {
                return None;
            }
        }
        1 => {
            if month_end + 2 > chars.len() {
                return None;
            }
            let (day, end) = read_fixed_digits(chars, month_end, 2)?;
            if !(1..=31).contains(&day) || chars.get(end).is_some_and(|c| c.is_ascii_digit()) {
                return None;
            }
        }
        2 => {
            let end = if chars.get(month_end) == Some(&'月') {
                month_end + 1
            } else {
                month_end
            };
            if chars.get(end).is_some_and(|c| c.is_ascii_digit()) {
                return None;
            }
        }
        3 => {
            if chars.get(month_end).is_some_and(|c| c.is_ascii_digit()) {
                return None;
            }
        }
        _ => return None,
    }
    Some((format!("{year:04}"), month))
}

fn read_month(chars: &[char], start: usize) -> Option<(u32, usize)> {
    let (value, end) = read_one_or_two_digits(chars, start)?;
    (1..=12).contains(&value).then_some((value, end))
}
fn read_day(chars: &[char], start: usize) -> Option<(u32, usize)> {
    let (value, end) = read_one_or_two_digits(chars, start)?;
    (1..=31).contains(&value).then_some((value, end))
}
fn read_one_or_two_digits(chars: &[char], start: usize) -> Option<(u32, usize)> {
    let first = *chars.get(start)?;
    if !first.is_ascii_digit() {
        return None;
    }
    if start + 1 < chars.len() && chars[start + 1].is_ascii_digit() {
        let value = format!("{first}{}", chars[start + 1]).parse().ok()?;
        Some((value, start + 2))
    } else {
        Some((first.to_digit(10)?, start + 1))
    }
}
fn read_fixed_digits(chars: &[char], start: usize, count: usize) -> Option<(u32, usize)> {
    let slice = chars.get(start..start + count)?;
    if !slice.iter().all(|character| character.is_ascii_digit()) {
        return None;
    }
    Some((
        slice.iter().collect::<String>().parse().ok()?,
        start + count,
    ))
}
fn day_valid(day: u32) -> bool {
    (1..=31).contains(&day)
}
fn sanitize_folder(text: &str) -> String {
    let filtered = text
        .chars()
        .map(|c| {
            if "<>:\"/\\|?*".contains(c) || c.is_control() {
                '_'
            } else {
                c
            }
        })
        .collect::<String>();
    let trimmed = filtered.trim().trim_matches('.');
    if trimmed.is_empty() {
        "未识别".into()
    } else {
        trimmed.into()
    }
}
fn next_target(raw: &Path, planned: &mut HashSet<PathBuf>) -> PathBuf {
    let mut candidate = raw.to_path_buf();
    let mut i = 1;
    while candidate.exists() || planned.contains(&planned_path_key(&candidate)) {
        let mut candidate_name = format!(
            "{}_{}",
            raw.file_stem().unwrap_or_default().to_string_lossy(),
            i
        );
        if let Some(extension) = raw.extension() {
            candidate_name.push('.');
            candidate_name.push_str(&extension.to_string_lossy());
        }
        candidate = raw.with_file_name(candidate_name);
        i += 1;
    }
    planned.insert(planned_path_key(&candidate));
    candidate
}
fn planned_path_key(path: &Path) -> PathBuf {
    #[cfg(windows)]
    {
        PathBuf::from(path.to_string_lossy().to_lowercase())
    }
    #[cfg(not(windows))]
    {
        path.to_path_buf()
    }
}
fn copy_with_metadata(source: &Path, target: &Path) -> Result<PathBuf> {
    use std::io::Write;
    let parent = target.parent().context("archive target has no parent")?;
    fs::create_dir_all(parent)?;
    let mut candidate = target.to_path_buf();
    let mut suffix = 1usize;
    loop {
        match fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&candidate)
        {
            Ok(mut destination) => {
                let result = (|| -> Result<()> {
                    let mut input = fs::File::open(source)?;
                    std::io::copy(&mut input, &mut destination)?;
                    destination.flush()?;
                    let metadata = fs::metadata(source)?;
                    fs::set_permissions(&candidate, metadata.permissions())?;
                    if let Ok(modified) = metadata.modified() {
                        destination.set_times(std::fs::FileTimes::new().set_modified(modified))?;
                    }
                    Ok(())
                })();
                if let Err(error) = result {
                    let _ = fs::remove_file(&candidate);
                    return Err(error);
                }
                return Ok(candidate);
            }
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => {
                candidate = target.with_file_name(format!(
                    "{}_{}",
                    target.file_stem().unwrap_or_default().to_string_lossy(),
                    suffix
                ));
                if let Some(ext) = target.extension() {
                    candidate.set_extension(ext);
                }
                suffix = suffix
                    .checked_add(1)
                    .context("archive target suffix overflow")?;
            }
            Err(error) => return Err(error.into()),
        }
    }
}
fn parse_extensions(raw: &str) -> BTreeSet<String> {
    raw.split([';', '；', ',', '，', ' ', '\t'])
        .map(str::trim)
        .filter(|s| !s.is_empty())
        .map(|s| format!(".{}", s.trim_start_matches('.').to_lowercase()))
        .collect()
}
fn truthy(raw: &str) -> bool {
    matches!(
        raw.trim().to_lowercase().as_str(),
        "y" | "yes" | "是" | "1" | "true" | "on"
    )
}
fn parse_a1(addr: &str) -> Option<(usize, usize)> {
    let i = addr.find(|c: char| c.is_ascii_digit())?;
    let col = addr[..i].chars().try_fold(0usize, |v, c| {
        v.checked_mul(26)?
            .checked_add(c.to_ascii_uppercase() as usize - 'A' as usize + 1)
    })?;
    Some((addr[i..].parse().ok()?, col))
}
fn header_index(sheet: &excel_reader::SheetDump) -> std::collections::HashMap<String, usize> {
    sheet
        .values
        .iter()
        .filter_map(|(addr, val)| {
            let (row, col) = parse_a1(addr)?;
            (row == 1).then(|| (value_text(val), col))
        })
        .collect()
}
fn header(map: &std::collections::HashMap<String, usize>, names: &[&str]) -> Option<usize> {
    names.iter().find_map(|name| map.get(*name).copied())
}
fn get_cell_text(sheet: &excel_reader::SheetDump, row: usize, col: usize) -> String {
    value_text(
        sheet
            .values
            .get(&format!("{}{}", col_letters(col), row))
            .unwrap_or(&serde_json::Value::Null),
    )
}
fn value_text(v: &serde_json::Value) -> String {
    match v.get("type").and_then(serde_json::Value::as_str) {
        Some("string") | Some("datetime_iso") | Some("duration_iso") | Some("error") => v
            .get("value")
            .and_then(serde_json::Value::as_str)
            .unwrap_or_default()
            .to_owned(),
        Some("number") | Some("integer") => {
            v.get("value").map(ToString::to_string).unwrap_or_default()
        }
        Some("boolean") => v
            .get("value")
            .and_then(serde_json::Value::as_bool)
            .map(|b| if b { "True" } else { "False" })
            .unwrap_or_default()
            .to_owned(),
        _ => String::new(),
    }
}
fn cell_value_text(sheet: &excel_reader::SheetDump, addr: &str) -> String {
    value_text(sheet.values.get(addr).unwrap_or(&serde_json::Value::Null))
}
fn col_letters(mut col: usize) -> String {
    let mut out = String::new();
    while col > 0 {
        let r = (col - 1) % 26;
        out.insert(0, (b'A' + r as u8) as char);
        col = (col - 1) / 26;
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::time::{SystemTime, UNIX_EPOCH};

    struct TestDirectory(PathBuf);

    impl TestDirectory {
        fn new() -> Self {
            let nonce = SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .expect("system clock must be after unix epoch")
                .as_nanos();
            let path = std::env::temp_dir().join(format!(
                "pytools-archive-walk-{}-{nonce}",
                std::process::id()
            ));
            fs::create_dir_all(&path).expect("create archive test directory");
            Self(path)
        }
    }

    impl Drop for TestDirectory {
        fn drop(&mut self) {
            let _ = fs::remove_dir_all(&self.0);
        }
    }

    #[test]
    fn archive_directory_scan_errors_are_propagated() {
        let root = TestDirectory::new();
        let file = root.0.join("not-a-directory.xlsx");
        fs::write(&file, b"source").expect("write test file");
        let mut files = Vec::new();

        let error = walk_files(&file, &mut files, &mut HashSet::new())
            .expect_err("read_dir on a file must fail rather than look like an empty folder");

        assert!(format!("{error:#}").contains("scan archive directory"));
        assert!(files.is_empty());
    }

    #[test]
    fn archive_scan_error_creates_no_target_or_report_directory() {
        let root = TestDirectory::new();
        let invalid_path = PathBuf::from("archive\0input");
        let target_root = root.0.join("archive");
        let output_dir = root.0.join("reports");

        let error = archive_by_date_copy(&[invalid_path], &target_root, &output_dir, None)
            .expect_err("an invalid path must stop the archive before side effects");

        assert!(format!("{error:#}").contains("inspect archive input"));
        assert!(!target_root.exists());
        assert!(!output_dir.exists());
    }

    #[test]
    fn archive_directory_walk_stops_at_an_active_canonical_directory() {
        let root = TestDirectory::new();
        fs::write(root.0.join("source.xlsx"), b"source").expect("write test file");
        let identity = root.0.canonicalize().expect("canonicalize test directory");
        let mut active_directories = HashSet::from([identity.clone()]);
        let mut files = Vec::new();

        walk_files(&root.0, &mut files, &mut active_directories)
            .expect("an ancestor cycle should be skipped safely");

        assert!(files.is_empty());
        assert!(active_directories.contains(&identity));
    }
}
