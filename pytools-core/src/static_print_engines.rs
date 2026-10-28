//! 按配置生成静态打印报表（报表打印菜单：3-1 umya / 3-2 openpyxl；另有
//! rust_xlsxwriter 引擎保留为内部实现）。
//!
//! 三者共享同一份“打印配置”语义，差异只在复制/写出的引擎：
//! - `umya`：umya 打开源工作簿整簿读入，在**同一包内**把源表复制到目标表
//!   （样式表同簿引用，字体/边框/填充/数字格式/行高列宽/合并全部原样保留），
//!   公式写入缓存值转静态。
//! - `openpyxl`：调用本机 Python openpyxl，按原 PyTools `print_310.py`
//!   的 `_copy_block` 语义逐属性复制样式（与旧工具最接近）。
//! - `rust_writer`：calamine 读缓存值 + rust_xlsxwriter 新建目标（外观重置，
//!   保留合并、列宽、页面设置、零值隐藏、批注丢弃）。
//!
//! 各引擎通过 `crate::engine_diag` 推送统一诊断事件（引擎/阶段/结果/上下文/
//! 耗时/原始错误/发生时刻），由菜单层汇总显示。
//!
//! 共通的行为对齐原 Python：不拒绝 Excel 错误值（原工具原样输出）、
//! 允许写入已存在的目标工作簿、段间空一行、公式取缓存值不留公式。

use anyhow::{bail, ensure, Context, Result};
use crate::config::PrintTask;
use crate::engine_diag::{DiagEvent, DiagOutcome, DiagStage};
use excel_reader::SheetDump;
use std::{
    fs,
    path::{Path, PathBuf},
    process::{Command, Stdio},
    time::Instant,
};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StaticPrintEngine {
    RustWriter,
    Umya,
    Openpyxl,
}

impl StaticPrintEngine {
    pub fn label(self) -> &'static str {
        match self {
            Self::RustWriter => "rust_xlsxwriter",
            Self::Umya => "umya-spreadsheet",
            Self::Openpyxl => "openpyxl",
        }
    }
}

#[derive(Debug, Default, Clone)]
pub struct StaticPrintRun {
    pub task_ok: usize,
    pub task_skipped: usize,
    pub task_failed: usize,
    pub outputs: Vec<PathBuf>,
    pub issues: Vec<String>,
    /// 统一引擎诊断事件（含各自发生时刻）；由界面层汇总显示。
    pub events: Vec<DiagEvent>,
}

fn is_xlsx(path: &Path) -> bool {
    path.extension()
        .and_then(|ext| ext.to_str())
        .is_some_and(|ext| matches!(ext.to_ascii_lowercase().as_str(), "xlsx" | "xlsm"))
}

fn used_bounds(sheet: &SheetDump) -> (usize, usize, usize, usize) {
    let mut max_row = 1usize;
    let mut max_col = 1usize;
    for address in sheet.values.keys() {
        if let Some((row, column)) = parse_a1(address) {
            max_row = max_row.max(row);
            max_col = max_col.max(column);
        }
    }
    (1, 1, max_row, max_col)
}

fn parse_a1(address: &str) -> Option<(usize, usize)> {
    let split = address.find(|c: char| c.is_ascii_digit())?;
    let mut column = 0usize;
    for byte in address[..split].bytes() {
        if !byte.is_ascii_alphabetic() {
            return None;
        }
        column = column * 26 + (byte.to_ascii_uppercase() - b'A' + 1) as usize;
    }
    let row: usize = address[split..].parse().ok()?;
    (row > 0 && column > 0).then_some((row, column))
}


/// 解析 A1 区域文本（如 "B2:B3"、"'Sheet'!$A$1:$C$2"）为行列边界。
fn parse_range_bounds(text: &str) -> Option<(usize, usize, usize, usize)> {
    let text = text.split('!').last().unwrap_or(text).replace('$', "");
    let (start, end) = text.split_once(':').unwrap_or((text.as_str(), text.as_str()));
    let (r1, c1) = parse_a1(start.trim())?;
    let (r2, c2) = parse_a1(end.trim())?;
    Some((r1.min(r2), c1.min(c2), r1.max(r2), c1.max(c2)))
}

fn column_letters(mut column: usize) -> String {
    let mut letters = String::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        letters.insert(0, (b'A' + remainder as u8) as char);
        column = (column - 1) / 26;
    }
    letters
}

fn parse_ranges(text: &str) -> Result<Vec<(usize, usize, usize, usize)>> {
    let text = text.trim();
    if text.is_empty() {
        return Ok(Vec::new());
    }
    let mut out = Vec::new();
    for segment in text.replace('；', ";").split(';') {
        let segment = segment.trim();
        if segment.is_empty() {
            continue;
        }
        let (start, end) = segment
            .split_once(':')
            .context(format!("打印区域格式非法: {segment:?}"))?;
        let (r1, c1) = parse_a1(start.trim().replace('$', "").as_str())
            .context(format!("打印区域起点非法: {start:?}"))?;
        let (r2, c2) = parse_a1(end.trim().replace('$', "").as_str())
            .context(format!("打印区域终点非法: {end:?}"))?;
        ensure!(r1 <= r2 && c1 <= c2, "打印区域起止颠倒: {segment:?}");
        out.push((r1, c1, r2, c2));
    }
    Ok(out)
}

/// 原 PyTools `_is_zero_like`：数值 0 与可解析为 0 的文本算零值；布尔不算。
fn is_zero_like(value: &serde_json::Value) -> bool {
    let kind = value.get("type").and_then(serde_json::Value::as_str);
    match kind {
        Some("boolean") => false,
        Some("integer") => value
            .get("value")
            .and_then(serde_json::Value::as_i64)
            .is_some_and(|number| number == 0),
        Some("number") | Some("datetime_serial") => value
            .get("value")
            .and_then(serde_json::Value::as_f64)
            .is_some_and(|number| number == 0.0),
        Some("string") => value
            .get("value")
            .and_then(serde_json::Value::as_str)
            .map(|text| {
                let text = text.trim().replace(',', "");
                !text.is_empty() && text.parse::<f64>().is_ok_and(|number| number == 0.0)
            })
            .unwrap_or(false),
        _ => false,
    }
}

/// 3-3：现有 Rust 写入路径（calamine 读值 → rust_xlsxwriter 新建目标）。
pub fn run_rust_writer_engine(tasks: &[PrintTask]) -> Result<StaticPrintRun> {
    use crate::print_ops::{build_config_print_plans, write_static_print_plans};
    let mut run = StaticPrintRun::default();
    let enabled: Vec<PrintTask> = tasks.iter().filter(|task| task.enabled).cloned().collect();
    if enabled.is_empty() {
        bail!("未找到启用的打印任务，请检查“打印配置”工作表和“是否启用”列（Y/是/1）");
    }
    // 必须一次传入全部启用行：同一目标工作簿通常由多行配置拼装，
    // 逐行建计划会让后一行因“目标已存在”失败。
    let mut plans = Vec::new();
    match build_config_print_plans(&enabled) {
        Ok(reports) => {
            for report in reports {
                match report.plan {
                    Some(plan) => plans.push(plan),
                    None => {
                        run.task_skipped += 1;
                        run.events.push(
                            DiagEvent::new(
                                "rust_xlsxwriter",
                                DiagStage::SaveOutput,
                                DiagOutcome::Skipped,
                            )
                            .with_file(report.output.to_string_lossy().into_owned())
                            .with_message(format!(
                                "未产生计划（命中工作表 {}，标记区域 {}）",
                                report.matching_sheets, report.marked_ranges
                            )),
                        );
                        run.issues.push(format!(
                            "{} 未产生计划（命中工作表 {}，标记区域 {}）",
                            report.output.display(),
                            report.matching_sheets,
                            report.marked_ranges
                        ));
                    }
                }
            }
        }
        Err(error) => {
            run.task_failed += enabled.len();
            run.events.push(
                DiagEvent::new("rust_xlsxwriter", DiagStage::CopySheet, DiagOutcome::Failed)
                    .with_message(format!("打印计划构建失败：{error:#}")),
            );
            run.issues.push(format!("打印计划构建失败：{error:#}"));
        }
    }
    if !plans.is_empty() {
        let write_started = Instant::now();
        let executions = write_static_print_plans(&plans)?;
        for execution in executions {
            run.task_ok += 1;
            run.events.push(
                DiagEvent::new("rust_xlsxwriter", DiagStage::SaveOutput, DiagOutcome::Ok)
                    .with_file(execution.output.to_string_lossy().into_owned())
                    .with_elapsed_ms(write_started.elapsed().as_millis()),
            );
            run.outputs.push(execution.output);
        }
    }
    Ok(run)
}

/// 3-4：umya。同一工作簿内 `new_sheet` 后逐单元格复制（样式随样式表同簿保留），
/// 公式只写缓存值。
pub fn run_umya_engine(tasks: &[PrintTask]) -> Result<StaticPrintRun> {
    let mut run = StaticPrintRun::default();
    let enabled: Vec<&PrintTask> = tasks.iter().filter(|task| task.enabled).collect();
    ensure!(
        !enabled.is_empty(),
        "未找到启用的打印任务，请检查“打印配置”工作表和“是否启用”列（Y/是/1）"
    );

    // 按目标工作簿分组，一个目标只打开/保存一次。
    let mut grouped = std::collections::BTreeMap::<PathBuf, Vec<&PrintTask>>::new();
    for task in enabled {
        let Some(target) = task.target_wb.as_ref() else {
            run.task_skipped += 1;
            run.issues.push(format!("第 {} 行：目标工作簿为空", task.row_no));
            continue;
        };
        grouped.entry(target.clone()).or_default().push(task);
    }

    for (target, target_tasks) in grouped {
        ensure!(is_xlsx(&target), "打印目标必须是 .xlsx/.xlsm：{}", target.display());
        let mut book = if target.is_file() {
            umya_spreadsheet::reader::xlsx::read(&target)
                .with_context(|| format!("umya 打开既有目标 {}", target.display()))?
        } else {
            umya_spreadsheet::new_file()
        };
        // 新建簿默认带一个空 Sheet，目标表创建前移除，避免残留。
        if !target.is_file() {
            let stale: Vec<String> = book
                .get_sheet_collection()
                .iter()
                .map(|sheet| sheet.get_name().to_owned())
                .collect();
            for name in stale {
                let empty = book
                    .get_sheet_by_name(&name)
                    .map(|sheet| sheet.get_highest_row() <= 1 && sheet.get_highest_column() <= 1)
                    .unwrap_or(false);
                if empty {
                    let _ = book.remove_sheet_by_name(&name);
                }
            }
        }

        for task in target_tasks {
            let task_started = Instant::now();
            let Some(source_path) = task.source_wb.as_ref() else {
                run.task_skipped += 1;
                run.events.push(
                    DiagEvent::new("umya-spreadsheet", DiagStage::ReadSource, DiagOutcome::Skipped)
                        .with_row(task.row_no)
                        .with_message("配置的源工作簿路径为空"),
                );
                run.issues.push(format!("第 {} 行：源工作簿为空", task.row_no));
                continue;
            };
            if !source_path.is_file() || !is_xlsx(source_path) {
                run.task_failed += 1;
                run.events.push(
                    DiagEvent::new("umya-spreadsheet", DiagStage::ReadSource, DiagOutcome::Failed)
                        .with_row(task.row_no)
                        .with_file(source_path.to_string_lossy().into_owned())
                        .with_message("源工作簿不存在或不是 xlsx/xlsm"),
                );
                run.issues.push(format!(
                    "第 {} 行：[文件] 源工作簿不存在或不是 xlsx/xlsm：{}",
                    task.row_no,
                    source_path.display()
                ));
                continue;
            }
            let source_book = match umya_spreadsheet::reader::xlsx::read(source_path) {
                Ok(book) => book,
                Err(error) => {
                    run.task_failed += 1;
                    run.events.push(
                        DiagEvent::new(
                            "umya-spreadsheet",
                            DiagStage::ReadSource,
                            DiagOutcome::Failed,
                        )
                        .with_row(task.row_no)
                        .with_file(source_path.to_string_lossy().into_owned())
                        .with_elapsed_ms(task_started.elapsed().as_millis())
                        .with_message(format!("{error:#}")),
                    );
                    run.issues
                        .push(format!("第 {} 行：[读取] umya 读取源工作簿 {} 失败：{error}", task.row_no, source_path.display()));
                    continue;
                }
            };
            let source_lookup = source_book.get_sheet_by_name(&task.source_ws);
            match source_lookup {
                Ok(_) => {}
                Err(_) => {
                    run.task_failed += 1;
                    run.events.push(
                        DiagEvent::new(
                            "umya-spreadsheet",
                            DiagStage::CopySheet,
                            DiagOutcome::Failed,
                        )
                        .with_row(task.row_no)
                        .with_file(source_path.to_string_lossy().into_owned())
                        .with_sheet(task.source_ws.clone())
                        .with_message(format!(
                            "源工作簿中不存在工作表 {:?}",
                            task.source_ws
                        )),
                    );
                    run.issues.push(format!(
                        "第 {} 行：[配置] 源工作簿 {} 中不存在工作表 {:?}",
                        task.row_no, source_path.display(), task.source_ws
                    ));
                    continue;
                }
            }
            let source_sheet = source_book
                .get_sheet_by_name(&task.source_ws)
                .expect("source sheet resolved above");
            // 目标工作表名：空则沿用源名；Excel 限制 31 字符。
            let target_name: String = if task.target_ws.trim().is_empty() {
                task.source_ws.clone()
            } else {
                task.target_ws.trim().chars().take(31).collect()
            };
            if book.get_sheet_by_name(&target_name).is_ok() {
                // 既有目标表按原 Python 语义清空后重写。
                let sheet = book
                    .get_sheet_by_name_mut(&target_name)
                    .expect("target sheet exists");
                let (_, rows) = sheet.get_highest_column_and_row();
                if rows > 0 {
                    sheet.remove_row(1, rows);
                }
            } else {
                book.new_sheet(target_name.clone()).map_err(|error| {
                    anyhow::anyhow!("第 {} 行：[复制] umya 在目标 {} 新建工作表 {target_name:?} 失败：{error}", task.row_no, target.display())
                })?;
            }

            let ranges = if task.source_ranges.trim().is_empty() {
                let (max_col, max_row) = source_sheet.get_highest_column_and_row();
                vec![(1usize, 1usize, max_row as usize, max_col as usize)]
            } else {
                match parse_ranges(&task.source_ranges) {
                    Ok(ranges) => ranges,
                    Err(error) => {
                        run.task_failed += 1;
                        run.events.push(
                            DiagEvent::new(
                                "umya-spreadsheet",
                                DiagStage::ReadConfig,
                                DiagOutcome::Failed,
                            )
                            .with_row(task.row_no)
                            .with_message(format!("打印区域非法：{error:#}")),
                        );
                        run.issues
                            .push(format!("第 {} 行：[配置] 打印区域非法：{error:#}", task.row_no));
                        continue;
                    }
                }
            };

            let mut out_row = 1u32;
            let mut max_cols = 1u32;
            let hide_zero = task.hide_zero;
            let source_sheet_name = task.source_ws.clone();
            for (r1, c1, r2, c2) in ranges {
                let rows = (r2 - r1 + 1) as u32;
                let cols = (c2 - c1 + 1) as u32;
                let dst_start_row_for_segment = out_row;
                for row_offset in 0..rows {
                    for col_offset in 0..cols {
                        let src_row = r1 as u32 + row_offset;
                        let src_col = c1 as u32 + col_offset;
                        let dst_row = out_row + row_offset;
                        let dst_col = 1 + col_offset;
                        let src_address =
                            format!("{}{}", column_letters(src_col as usize), src_row);
                        let Ok(source_ws_ref) = source_book.get_sheet_by_name(&source_sheet_name)
                        else {
                            continue;
                        };
                        let Some(source_cell) = source_ws_ref.get_cell(src_address.as_str())
                        else {
                            continue;
                        };
                        let style = source_cell.get_style().clone();
                        let raw = source_cell.get_raw_value().clone();
                        let Ok(target_ws_ref) = book.get_sheet_by_name_mut(&target_name)
                        else {
                            continue;
                        };
                        let target_cell = target_ws_ref.cell_mut((dst_col, dst_row));
                        // 样式必须先写：即使零值被隐藏，边框/填充/数字格式仍要保留
                        // （跳过整格会让该单元格没有边框，打印出来就是断线）。
                        target_cell.set_style(style);
                        // 零值不输出：数值 0 只写样式不写值，保持真正空白。
                        if hide_zero && raw_value_is_numeric_zero(&raw) {
                            continue;
                        }
                        copy_raw_value_static(&raw, target_cell, source_cell.get_data_type());
                    }
                }
                // 复制工作表外观：umya 的 new_sheet 不带 sheetViews，默认行高也
                // 与源表不同；缺失时 WPS/Excel 会用默认视图渲染，细边框显示不清。
                {
                    let format = match source_book.get_sheet_by_name(&source_sheet_name) {
                        Ok(source_ws_ref) => {
                            let props = source_ws_ref.get_sheet_format_properties();
                            Some((
                                props.default_row_height(),
                                props.default_column_width(),
                            ))
                        }
                        Err(_) => None,
                    };
                    if let Ok(target_ws_ref) = book.get_sheet_by_name_mut(&target_name) {
                        if let Some((row_height, column_width)) = format {
                            let props = target_ws_ref.get_sheet_format_properties_mut();
                            if row_height > 0.0 {
                                props.set_default_row_height(row_height);
                            }
                            if column_width > 0.0 {
                                props.set_default_column_width(column_width);
                            }
                        }
                        // 确保至少存在一个 sheetView，避免输出缺少 <sheetViews>。
                        let views = target_ws_ref.sheet_views_mut().sheet_view_list_mut();
                        if views.is_empty() {
                            views.push(Default::default());
                        }
                    }
                }
                // 复制列宽与行高：umya 的 new_sheet 不会继承尺寸，
                // 缺失会让数值挤在默认列宽里显示不全（用户反馈的“数值还原不好”）。
                {
                    let widths: Vec<(u32, f64)> = match source_book
                        .get_sheet_by_name(&source_sheet_name)
                    {
                        Ok(source_ws_ref) => (0..cols)
                            .filter_map(|offset| {
                                let src_col = c1 as u32 + offset;
                                source_ws_ref
                                    .get_column_dimension_by_number(&src_col)
                                    .map(|column| (1 + offset, column.get_width()))
                            })
                            .collect(),
                        Err(_) => Vec::new(),
                    };
                    let heights: Vec<(u32, f64)> = match source_book
                        .get_sheet_by_name(&source_sheet_name)
                    {
                        Ok(source_ws_ref) => (0..rows)
                            .filter_map(|offset| {
                                let src_row = r1 as u32 + offset;
                                source_ws_ref
                                    .get_row_dimension(&src_row)
                                    .map(|row| (dst_start_row_for_segment + offset, row.get_height()))
                            })
                            .filter(|(_, height)| *height > 0.0)
                            .collect(),
                        Err(_) => Vec::new(),
                    };
                    if let Ok(target_ws_ref) = book.get_sheet_by_name_mut(&target_name) {
                        for (column, width) in widths {
                            target_ws_ref
                                .get_column_dimension_mut(&column_letters(column as usize))
                                .set_width(width);
                        }
                        for (row, height) in heights {
                            target_ws_ref
                                .get_row_dimension_mut(&row)
                                .set_height(height);
                        }
                    }
                }
                // 复制完全落在区域内的合并区（平移坐标），与原 Python 一致。
                let merge_ranges: Vec<String> = match source_book
                    .get_sheet_by_name(&source_sheet_name)
                {
                    Ok(source_ws_ref) => source_ws_ref
                        .get_merge_cells()
                        .iter()
                        .filter_map(|merged| {
                            let parsed = parse_range_bounds(&merged.get_range())?;
                            let (mr1, mc1, mr2, mc2) = parsed;
                            let inside = mr1 >= r1 && mc1 >= c1 && mr2 <= r2 && mc2 <= c2;
                            inside.then(|| {
                                let dr1 = (dst_start_row_for_segment + (mr1 - r1) as u32) as u32;
                                let dr2 = (dst_start_row_for_segment + (mr2 - r1) as u32) as u32;
                                let dc1 = (1 + (mc1 - c1)) as u32;
                                let dc2 = (1 + (mc2 - c1)) as u32;
                                format!(
                                    "{}{}:{}{}",
                                    column_letters(dc1 as usize),
                                    dr1,
                                    column_letters(dc2 as usize),
                                    dr2
                                )
                            })
                        })
                        .collect(),
                    Err(_) => Vec::new(),
                };
                if !merge_ranges.is_empty() {
                    if let Ok(target_ws_ref) = book.get_sheet_by_name_mut(&target_name) {
                        for merged in &merge_ranges {
                            target_ws_ref.add_merge_cells(merged.as_str());
                        }
                    }
                }
                // 段间空一行（与原 Python `out_row += rows + 1` 一致）。
                out_row += rows + 1;
                max_cols = max_cols.max(cols);
            }
            // 页面设置：对齐原 Python `_apply_print_setup`。无论区域多大，
            // 都把打印区域限定为已写入的矩形、按配置方向、FitToPage 缩放到一页。
            {
                let last_row = out_row.saturating_sub(2).max(1);
                let last_col = max_cols.max(1);
                let orientation = match task.orientation.trim() {
                    text if text.contains("横向") => {
                        Some(umya_spreadsheet::OrientationValues::Landscape)
                    }
                    text if text.contains("纵向") => {
                        Some(umya_spreadsheet::OrientationValues::Portrait)
                    }
                    _ => None,
                };
                if let Ok(target_ws_ref) = book.get_sheet_by_name_mut(&target_name) {
                    let page = target_ws_ref.get_page_setup_mut();
                    page.set_fit_to_page(true);
                    page.set_paper_size(9);
                    page.set_scale(0);
                    page.set_fit_to_width(task.fit_wide.max(1) as u32);
                    page.set_fit_to_height(task.fit_tall.max(0) as u32);
                    if let Some(orientation) = orientation {
                        page.set_orientation(orientation);
                    }
                    let options = target_ws_ref.get_print_options_mut();
                    options.set_horizontal_centered(task.center_h);
                    options.set_vertical_centered(task.center_v);
                }
                let print_area = format!(
                    "'{target_name}'!$A$1:${}${last_row}",
                    column_letters(last_col as usize)
                );
                let sheet_index = book
                    .get_sheet_collection()
                    .iter()
                    .position(|sheet| sheet.get_name() == target_name)
                    .unwrap_or(0) as u32;
                let mut defined_name = umya_spreadsheet::DefinedName::default();
                defined_name
                    .set_name("_xlnm.Print_Area")
                    .set_address(&print_area)
                    .set_local_sheet_id(sheet_index);
                book.add_defined_names(defined_name);
            }
            run.task_ok += 1;
            run.events.push(
                DiagEvent::new("umya-spreadsheet", DiagStage::CopySheet, DiagOutcome::Ok)
                    .with_row(task.row_no)
                    .with_file(source_path.to_string_lossy().into_owned())
                    .with_sheet(target_name.clone())
                    .with_elapsed_ms(task_started.elapsed().as_millis()),
            );
        }

        if let Some(parent) = target.parent() {
            if !parent.as_os_str().is_empty() {
                fs::create_dir_all(parent)
                    .with_context(|| format!("[写入] 创建目标目录 {} 失败", parent.display()))?;
            }
        }
        let save_started = Instant::now();
        if let Err(error) = umya_spreadsheet::writer::xlsx::write(&book, &target) {
            run.events.push(
                DiagEvent::new("umya-spreadsheet", DiagStage::SaveOutput, DiagOutcome::Failed)
                    .with_file(target.to_string_lossy().into_owned())
                    .with_elapsed_ms(save_started.elapsed().as_millis())
                    .with_message(format!("{error:#}")),
            );
            return Err(error).with_context(|| {
                format!("[写入] umya 写出目标工作簿 {} 失败", target.display())
            });
        }
        run.events.push(
            DiagEvent::new("umya-spreadsheet", DiagStage::SaveOutput, DiagOutcome::Ok)
                .with_file(target.to_string_lossy().into_owned())
                .with_elapsed_ms(save_started.elapsed().as_millis()),
        );
        run.outputs.push(target);
    }
    Ok(run)
}

fn raw_value_is_numeric_zero(raw: &umya_spreadsheet::CellRawValue) -> bool {
    matches!(raw, umya_spreadsheet::CellRawValue::Numeric(number) if *number == 0.0)
}

fn copy_raw_value_static(
    raw: &umya_spreadsheet::CellRawValue,
    target: &mut umya_spreadsheet::Cell,
    data_type: &str,
) {
    use umya_spreadsheet::CellRawValue;
    match raw {
        CellRawValue::String(text) => {
            // 文本 "0" 属于文本，零值隐藏只针对数值单元格（用户验收口径）。
            target.set_value_string(text.clone());
        }
        CellRawValue::RichText(_) => {
            target.set_value_string(raw.to_string());
        }
        CellRawValue::Numeric(number) => {
            target.set_value_number(*number);
        }
        CellRawValue::Bool(value) => {
            target.set_value_bool(*value);
        }
        CellRawValue::Error(_) => {
            // 原 Python 打印工具原样输出错误值；umya 以文本保留其显示。
            target.set_value_string(raw.to_string());
        }
        _ => {
            if data_type == "e" {
                target.set_value_string(raw.to_string());
            }
        }
    }
}

/// 3-5：openpyxl，逐属性复制样式，行为对齐原 PyTools `print_310.py`。
pub fn run_openpyxl_engine(tasks: &[PrintTask], python_exe: Option<&str>) -> Result<StaticPrintRun> {
    let enabled: Vec<&PrintTask> = tasks.iter().filter(|task| task.enabled).collect();
    ensure!(
        !enabled.is_empty(),
        "未找到启用的打印任务，请检查“打印配置”工作表和“是否启用”列（Y/是/1）"
    );
    let plan = serde_json::json!({
        "tasks": enabled.iter().map(|task| serde_json::json!({
            "row_no": task.row_no,
            "mode": task.mode,
            "source_wb": task.source_wb.as_ref().map(|p| p.to_string_lossy().into_owned()),
            "source_ws": task.source_ws,
            "source_ranges": task.source_ranges,
            "target_wb": task.target_wb.as_ref().map(|p| p.to_string_lossy().into_owned()),
            "target_ws": task.target_ws,
            "fit_wide": task.fit_wide,
            "fit_tall": task.fit_tall,
            "orientation": task.orientation,
            "center_h": task.center_h,
            "center_v": task.center_v,
            "drop_comments": task.drop_comments,
            "hide_zero": task.hide_zero,
        })).collect::<Vec<_>>()
    });
    let plan_path = unique_temp_json(&plan).context("[准备] 创建 openpyxl 临时计划失败")?;
    let script = r#"
import json, sys, copy, time
from pathlib import Path
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.utils.cell import get_column_letter, range_boundaries
from openpyxl.worksheet.properties import PageSetupProperties

plan = json.loads(Path(sys.argv[1]).read_text(encoding='utf-8'))
task_ok = task_skip = task_failed = 0
outputs, issues, events = [], [], []

def record_event(outcome, stage_key, row_no, file_path, message, elapsed_ms):
    # 结构化诊断事件：随 JSON 回传给 Rust 侧统一显示，不单独打印日志。
    events.append({
        'stage': stage_key,
        'outcome': outcome,
        'row': row_no,
        'file': str(file_path) if file_path else None,
        'message': message or '',
        'elapsed_ms': int(elapsed_ms),
        'ts': int(time.time()),
    })

def stage_key_for(text):
    if text.startswith('openpyxl 读取源工作簿'):
        return 'read_source'
    if text.startswith('openpyxl 写入目标工作簿'):
        return 'save'
    if text.startswith('解析打印区域') or text == '校验配置':
        return 'read_config'
    return 'copy'

def used_bounds(ws):
    min_col, min_row, max_col, max_row = range_boundaries(ws.calculate_dimension())
    if max_row < min_row or max_col < min_col:
        return 1, 1, 1, 1
    return min_row, min_col, max_row, max_col

def parse_ranges(text):
    text = (text or '').strip()
    if not text:
        return []
    out = []
    for seg in [p.strip() for p in text.replace('；', ';').split(';') if p.strip()]:
        min_col, min_row, max_col, max_row = range_boundaries(seg)
        out.append((min_row, min_col, max_row, max_col))
    return out

def is_zero_like(v):
    # 只把真正的数值 0 视为零值；文本 "0" 必须保留（用户验收口径）。
    if v is None or isinstance(v, bool):
        return False
    if isinstance(v, (int, float)):
        return abs(float(v)) == 0.0
    return False

def copy_block(src_ws, dst_ws, bounds, dst_start_row, mode, value_ws, hide_zero, drop_comments, row_no):
    min_row, min_col, max_row, max_col = bounds
    rows_count = max_row - min_row + 1
    cols_count = max_col - min_col + 1
    for r in range(rows_count):
        src_r = min_row + r
        dst_r = dst_start_row + r
        if mode == 1:
            h = src_ws.row_dimensions[src_r].height
            if h is not None:
                dst_ws.row_dimensions[dst_r].height = h
        for c in range(cols_count):
            src_c = min_col + c
            dst_c = 1 + c
            src_cell = src_ws.cell(src_r, src_c)
            dst_cell = dst_ws.cell(dst_r, dst_c)
            raw = value_ws.cell(src_r, src_c).value if value_ws is not None else src_cell.value
            dst_cell.value = None if (hide_zero and is_zero_like(raw)) else raw
            if mode in (1, 2):
                if src_cell.has_style:
                    dst_cell.font = copy.copy(src_cell.font)
                    dst_cell.fill = copy.copy(src_cell.fill)
                    dst_cell.border = copy.copy(src_cell.border)
                    dst_cell.alignment = copy.copy(src_cell.alignment)
                    dst_cell.protection = copy.copy(src_cell.protection)
                    dst_cell.number_format = src_cell.number_format
                if (not drop_comments) and src_cell.comment is not None:
                    dst_cell.comment = copy.copy(src_cell.comment)
            else:
                if src_cell.has_style:
                    dst_cell.font = copy.copy(src_cell.font)
                    dst_cell.fill = copy.copy(src_cell.fill)
                    dst_cell.alignment = copy.copy(src_cell.alignment)
                    dst_cell.number_format = src_cell.number_format
    if mode in (1, 2):
        for c in range(cols_count):
            w = src_ws.column_dimensions[get_column_letter(min_col + c)].width
            if w is not None:
                dst_ws.column_dimensions[get_column_letter(1 + c)].width = w
    for merged in list(src_ws.merged_cells.ranges):
        if merged.min_row >= min_row and merged.max_row <= max_row and merged.min_col >= min_col and merged.max_col <= max_col:
            dr1 = dst_start_row + (merged.min_row - min_row)
            dr2 = dst_start_row + (merged.max_row - min_row)
            dc1 = 1 + (merged.min_col - min_col)
            dc2 = 1 + (merged.max_col - min_col)
            try:
                dst_ws.merge_cells(f'{get_column_letter(dc1)}{dr1}:{get_column_letter(dc2)}{dr2}')
            except Exception as e:
                issues.append(f"第 {row_no} 行：[复制] 合并区域 {merged} 复制失败，生成文件可能不完整：{type(e).__name__}: {e}")
    return rows_count, cols_count

def apply_setup(ws, total_rows, total_cols, fit_wide, fit_tall, orientation, center_h, center_v):
    ws.page_setup.orientation = orientation or 'portrait'
    if ws.sheet_properties.pageSetUpPr is None:
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    else:
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = max(1, int(fit_wide or 1))
    ws.page_setup.fitToHeight = max(0, int(fit_tall or 0))
    ws.page_setup.scale = None
    ws.page_setup.paperSize = 9
    ws.print_area = f'A1:{get_column_letter(max(1, total_cols))}{max(1, total_rows)}'
    ws.print_options.horizontalCentered = bool(center_h)
    ws.print_options.verticalCentered = bool(center_v)

def parse_orientation(text):
    text = (text or '').strip()
    if '横向' in text:
        return 'landscape'
    if '纵向' in text:
        return 'portrait'
    return None

cache = {}
for t in plan['tasks']:
    stage = '校验配置'
    task_t0 = time.perf_counter()
    src_path = tgt_path = None
    try:
        src_path = Path(t['source_wb']) if t['source_wb'] else None
        tgt_path = Path(t['target_wb']) if t['target_wb'] else None
        if src_path is None or tgt_path is None or not t['target_ws'].strip():
            task_skip += 1
            missing = []
            if src_path is None:
                missing.append('源工作簿')
            if tgt_path is None:
                missing.append('目标工作簿')
            if not t['target_ws'].strip():
                missing.append('目标工作表')
            detail = f"{'、'.join(missing)}为空，已跳过"
            record_event('skipped', 'read_source', t.get('row_no'), str(src_path or ''), detail, (time.perf_counter() - task_t0) * 1000)
            issues.append(f"第 {t['row_no']} 行：[配置] {detail}")
            continue
        if not src_path.is_file():
            task_failed += 1
            record_event('failed', 'read_source', t.get('row_no'), str(src_path), '源工作簿不存在', (time.perf_counter() - task_t0) * 1000)
            issues.append(f"第 {t['row_no']} 行：[文件] 源工作簿不存在：{src_path}")
            continue
        key = str(src_path)
        if key not in cache:
            stage = f'openpyxl 读取源工作簿 {src_path}'
            cache[key] = (
                load_workbook(src_path, keep_vba=src_path.suffix.lower() == '.xlsm'),
                load_workbook(src_path, keep_vba=src_path.suffix.lower() == '.xlsm', data_only=True),
            )
        src_wb, src_val_wb = cache[key]
        if t['source_ws'] not in src_wb.sheetnames:
            task_failed += 1
            record_event('failed', 'copy', t.get('row_no'), str(src_path), f"源工作簿中不存在工作表 {t['source_ws']!r}", (time.perf_counter() - task_t0) * 1000)
            issues.append(f"第 {t['row_no']} 行：[配置] 源工作簿 {src_path} 中不存在工作表 {t['source_ws']!r}")
            continue
        src_ws = src_wb[t['source_ws']]
        src_val_ws = src_val_wb[t['source_ws']] if t['source_ws'] in src_val_wb.sheetnames else None
        stage = '解析打印区域配置'
        ranges = parse_ranges(t.get('source_ranges')) or [used_bounds(src_ws)]
        stage = f'openpyxl 准备目标工作簿 {tgt_path}'
        if tgt_path.exists() and tgt_path.suffix.lower() in ('.xlsx', '.xlsm'):
            tgt_wb = load_workbook(tgt_path, keep_vba=tgt_path.suffix.lower() == '.xlsm')
        else:
            tgt_wb = Workbook()
            if tgt_wb.active is not None and tgt_wb.active.title == 'Sheet':
                tgt_wb.remove(tgt_wb.active)
        name = (t['target_ws'].strip() or t['source_ws'])[:31]
        if name in tgt_wb.sheetnames:
            ws = tgt_wb[name]
            if ws.max_row > 1 or ws.cell(1, 1).value not in (None, ''):
                ws.delete_rows(1, ws.max_row)
        else:
            ws = tgt_wb.create_sheet(title=name)
        out_row = 1
        max_cols = 1
        issues_before_copy = len(issues)
        stage = f'复制工作表 {t["source_ws"]!r} 到 {name!r}'
        for rg in ranges:
            rows, cols = copy_block(src_ws, ws, rg, out_row, int(t.get('mode', 1)), src_val_ws,
                                    bool(t.get('hide_zero')), bool(t.get('drop_comments')), t['row_no'])
            out_row += rows + 1
            max_cols = max(max_cols, cols)
        stage = f'设置目标工作表 {name!r} 的打印参数'
        apply_setup(ws, out_row - 1, max_cols, t.get('fit_wide'), t.get('fit_tall'),
                    parse_orientation(t.get('orientation')), t.get('center_h'), t.get('center_v'))
        stage = f'openpyxl 写入目标工作簿 {tgt_path}'
        tgt_path.parent.mkdir(parents=True, exist_ok=True)
        tgt_wb.save(tgt_path)
        outputs.append(str(tgt_path))
        if len(issues) > issues_before_copy:
            task_failed += 1
            record_event('failed', 'copy', t.get('row_no'), str(tgt_path), '部分区域复制时出现问题，详见 issues', (time.perf_counter() - task_t0) * 1000)
        else:
            task_ok += 1
            record_event('ok', 'save', t.get('row_no'), str(tgt_path), '', (time.perf_counter() - task_t0) * 1000)
    except Exception as e:
        task_failed += 1
        record_event('failed', stage_key_for(stage), t.get('row_no'), str(tgt_path or src_path or ''), f"{stage}：{type(e).__name__}: {e}", (time.perf_counter() - task_t0) * 1000)
        issues.append(f"第 {t.get('row_no')} 行：[{stage}] 失败：{type(e).__name__}: {e}；源文件：{src_path}；目标文件：{tgt_path}")

sys.stdout.reconfigure(encoding='utf-8')
print(json.dumps({'task_ok': task_ok, 'task_skipped': task_skip, 'task_failed': task_failed, 'outputs': outputs, 'issues': issues, 'events': events}, ensure_ascii=True))
"#;
    let python = python_exe
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("python"));
    let output = Command::new(&python)
        .arg("-c")
        .arg(script)
        .arg(&plan_path)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .output()
        .with_context(|| {
            "该功能需要外部环境: 无法启动 Python（3-5 openpyxl 引擎需要本机 python 与 openpyxl）"
        })?;
    let _ = fs::remove_file(&plan_path);
    ensure!(
        output.status.success(),
        "openpyxl 打印引擎失败：{}",
        String::from_utf8_lossy(&output.stderr).trim()
    );
    let report: serde_json::Value = serde_json::from_slice(&output.stdout)
        .context("解析 openpyxl 打印引擎输出失败")?;
    let mut run = StaticPrintRun {
        task_ok: report["task_ok"].as_u64().unwrap_or(0) as usize,
        task_skipped: report["task_skipped"].as_u64().unwrap_or(0) as usize,
        task_failed: report["task_failed"].as_u64().unwrap_or(0) as usize,
        outputs: report["outputs"]
            .as_array()
            .map(|items| {
                items
                    .iter()
                    .filter_map(|item| item.as_str().map(PathBuf::from))
                    .collect()
            })
            .unwrap_or_default(),
        issues: report["issues"]
            .as_array()
            .map(|items| {
                items
                    .iter()
                    .filter_map(|item| item.as_str().map(str::to_owned))
                    .collect()
            })
            .unwrap_or_default(),
        events: Vec::new(),
    };
    // 子进程以结构化 JSON 回传诊断事件（含各自发生的 Unix 秒），此处转换为
    // 统一事件并还原为本机时间戳；普通输出与事件不混流。
    let mapped: Vec<DiagEvent> = report["events"]
        .as_array()
        .map(|items| {
            items
                .iter()
                .filter_map(|item| {
                    let stage = match item["stage"].as_str()? {
                        "read_config" => DiagStage::ReadConfig,
                        "read_source" => DiagStage::ReadSource,
                        "copy" => DiagStage::CopySheet,
                        "save" => DiagStage::SaveOutput,
                        _ => DiagStage::ModifyWorkbook,
                    };
                    let outcome = match item["outcome"].as_str()? {
                        "ok" => DiagOutcome::Ok,
                        "skipped" => DiagOutcome::Skipped,
                        _ => DiagOutcome::Failed,
                    };
                    let mut event = DiagEvent::new("openpyxl", stage, outcome)
                        .with_message(item["message"].as_str().unwrap_or_default());
                    if let Some(row) = item["row"].as_u64() {
                        event = event.with_row(row as usize);
                    }
                    if let Some(file) = item["file"].as_str() {
                        if !file.is_empty() {
                            event = event.with_file(file);
                        }
                    }
                    if let Some(elapsed_ms) = item["elapsed_ms"].as_u64() {
                        event = event.with_elapsed_ms(u128::from(elapsed_ms));
                    }
                    Some(event.with_epoch(item["ts"].as_i64()))
                })
                .collect()
        })
        .unwrap_or_default();
    run.events = mapped;
    run.task_ok = run.task_ok.max(0);
    // 多行配置可以写入同一个工作簿，最终输出列表按文件去重，任务计数不变。
    let mut seen = std::collections::HashSet::new();
    run.outputs.retain(|path| seen.insert(path.clone()));
    Ok(run)
}

fn unique_temp_json(value: &serde_json::Value) -> Result<PathBuf> {
    let nonce = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)?
        .as_nanos();
    let path = std::env::temp_dir().join(format!(
        "pytools-print-tasks-{}-{:?}-{nonce}.json",
        std::process::id(),
        std::thread::current().id()
    ));
    fs::write(&path, serde_json::to_vec(value)?)?;
    Ok(path)
}

/// 把目标改写到 `<原目录>/<引擎名>/<原文件名>`，三种引擎可并存对比。
fn engine_variant_target(target: &Path, engine: StaticPrintEngine) -> PathBuf {
    let folder = match engine {
        StaticPrintEngine::RustWriter => "3-3-rust_xlsxwriter",
        StaticPrintEngine::Umya => "3-4-umya",
        StaticPrintEngine::Openpyxl => "3-5-openpyxl",
    };
    let parent = target
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .map(|parent| parent.to_path_buf())
        .unwrap_or_else(|| PathBuf::from("."));
    let name = target
        .file_name()
        .map(|name| name.to_owned())
        .unwrap_or_else(|| std::ffi::OsString::from("static-print.xlsx"));
    parent.join(folder).join(name)
}

pub fn probe_python_exe() -> Option<String> {
    std::env::var("PYTOOLS_OPENPYXL_PYTHON")
        .ok()
        .filter(|value| !value.trim().is_empty())
}

/// 读打印配置并执行指定引擎（菜单 3-3/3-4/3-5 共用入口）。
pub fn run_configured_static_print(
    config_path: &Path,
    engine: StaticPrintEngine,
) -> Result<StaticPrintRun> {
    let tasks = crate::config::load_print_tasks(config_path)
        .with_context(|| format!("[配置读取] 读取打印配置 {} 失败（读取引擎：calamine）", config_path.display()))?;
    let enabled = tasks.iter().filter(|task| task.enabled).count();
    let tasks = tasks
        .into_iter()
        .map(|mut task| {
            task.target_wb = task.target_wb.map(|target| engine_variant_target(&target, engine));
            task
        })
        .collect::<Vec<_>>();
    let mut run = match engine {
        StaticPrintEngine::RustWriter => run_rust_writer_engine(&tasks),
        StaticPrintEngine::Umya => run_umya_engine(&tasks),
        StaticPrintEngine::Openpyxl => {
            run_openpyxl_engine(&tasks, probe_python_exe().as_deref())
        }
    }?;
    // 配置读取成功事件排在最前（读取引擎：calamine）。
    let mut events = vec![DiagEvent::new("calamine", DiagStage::ReadConfig, DiagOutcome::Ok)
        .with_file(config_path.to_string_lossy().into_owned())
        .with_message(format!("启用 {enabled} 项打印任务"))];
    events.append(&mut run.events);
    run.events = events;
    Ok(run)
}
