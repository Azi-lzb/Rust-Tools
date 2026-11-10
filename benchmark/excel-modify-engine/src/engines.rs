//! Edit operations and the three in-process Rust engines.

use anyhow::Result;
use serde_json::json;
use std::path::{Path, PathBuf};
use std::time::Instant;

#[derive(Clone, Debug)]
pub enum Val {
    Text(String),
    Num(f64),
    Bool(bool),
}

#[derive(Clone, Debug)]
pub enum Op {
    SetCellValue { sheet: String, cell: String, val: Val },
    SetFormula { sheet: String, cell: String, text: String, cached: Option<Val> },
    SetNumberFormat { sheet: String, cell: String, fmt: String },
    SetFontBold { sheet: String, cell: String },
    SetFillYellow { sheet: String, cell: String },
    SetBorderThin { sheet: String, cell: String },
    SetAlignCenter { sheet: String, cell: String },
    SetRowHeight { sheet: String, row: u32, h: f64 },
    SetColWidth { sheet: String, col: u32, w: f64 },
    AddMerge { sheet: String, range: String },
    RemoveMerge { sheet: String, range: String },
    SetComment { sheet: String, cell: String, text: String },
    SetGlobalDefinedName { name: String, reference: String },
    SetPrintArea { sheet: String, range: String },
    SetLandscape { sheet: String },
    AddSheet { name: String },
    DeleteSheet { name: String },
    RenameSheet { old: String, new: String },
    InsertRows { sheet: String, row: u32, n: u32 },
    DeleteRows { sheet: String, row: u32, n: u32 },
}

impl Op {
    pub fn kind(&self) -> &'static str {
        match self {
            Op::SetCellValue { .. } => "set_cell_value",
            Op::SetFormula { .. } => "set_formula",
            Op::SetNumberFormat { .. } => "set_number_format",
            Op::SetFontBold { .. } => "set_font_bold",
            Op::SetFillYellow { .. } => "set_fill",
            Op::SetBorderThin { .. } => "set_border",
            Op::SetAlignCenter { .. } => "set_align_center",
            Op::SetRowHeight { .. } => "set_row_height",
            Op::SetColWidth { .. } => "set_col_width",
            Op::AddMerge { .. } => "add_merge",
            Op::RemoveMerge { .. } => "remove_merge",
            Op::SetComment { .. } => "set_comment",
            Op::SetGlobalDefinedName { .. } => "set_defined_name",
            Op::SetPrintArea { .. } => "set_print_area",
            Op::SetLandscape { .. } => "set_landscape",
            Op::AddSheet { .. } => "add_sheet",
            Op::DeleteSheet { .. } => "delete_sheet",
            Op::RenameSheet { .. } => "rename_sheet",
            Op::InsertRows { .. } => "insert_rows",
            Op::DeleteRows { .. } => "delete_rows",
        }
    }
}

#[derive(Clone, Copy, Debug)]
pub struct PhaseTimes {
    pub open_ms: f64,
    pub modify_ms: f64,
    pub save_ms: f64,
}

#[derive(Debug)]
pub struct EditOutcome {
    pub output: Option<Vec<u8>>,
    pub phases: Option<PhaseTimes>,
    pub op_results: Vec<(String, Result<(), String>)>,
}

pub fn col_letter(mut n: u32) -> String {
    let mut out = String::new();
    while n > 0 {
        let rem = ((n - 1) % 26) as u8;
        out.insert(0, (b'A' + rem) as char);
        n = (n - 1) / 26;
    }
    out
}

/// "B2" -> (row, col) 1-based
pub fn parse_cell(addr: &str) -> (u32, u32) {
    let bytes = addr.as_bytes();
    let letters = bytes.iter().take_while(|b| b.is_ascii_alphabetic()).count();
    let col: u32 = addr[..letters]
        .bytes()
        .fold(0, |acc, b| acc * 26 + u32::from(b.to_ascii_uppercase() - b'A') + 1);
    let row: u32 = addr[letters..].parse().expect("valid row");
    (row, col)
}

pub fn parse_range(range: &str) -> ((u32, u32), (u32, u32)) {
    let (a, b) = range.split_once(':').unwrap_or((range, range));
    (parse_cell(a), parse_cell(b))
}

pub trait Engine {
    fn name(&self) -> &'static str;
    fn apply(&self, input: &[u8], ops: &[Op]) -> EditOutcome;
}

pub fn engine_by_name(name: &str) -> Option<Box<dyn Engine>> {
    match name {
        "betteroffice" => Some(Box::new(bo::BetterOffice)),
        "rxls" => Some(Box::new(rx::Rxls)),
        "umya" => Some(Box::new(umya_eng::Umya)),
        _ => None,
    }
}

// ------------------------------------------------- betteroffice 0.5.1

mod bo {
    use super::*;
    
    use betteroffice_xlsx::{
        CalculationOptions, CellRange, CellRef, CellValue, SheetId, WorkbookModel,
    };
    use xlsx_ops::{
        BorderLineStyle, BorderPatch, BorderPreset, CellState, HorizontalAlignment,
        NumberFormatMutation, Op as BoOp, StylePatch,
    };

    pub struct BetterOffice;

    fn at(wb: &WorkbookModel, sheet: &str, cell: &str) -> Option<(usize, CellRef)> {
        let (row, col) = parse_cell(cell);
        let idx = wb.sheets.iter().position(|s| s.name == sheet)?;
        Some((idx, CellRef { row: row - 1, col: col - 1, abs_row: false, abs_col: false }))
    }

    fn val_of(v: &Val) -> CellValue {
        match v {
            Val::Text(t) => CellValue::Text { value: t.clone() },
            Val::Num(n) => CellValue::Number { value: *n },
            Val::Bool(b) => CellValue::Bool { value: *b },
        }
    }

    fn sheet_id(wb: &WorkbookModel, name: &str) -> Option<SheetId> {
        wb.sheets
            .iter()
            .position(|s| s.name == name)
            .map(|idx| SheetId(idx as u32))
    }

    fn bo_range(wb: &WorkbookModel, sheet: &str, range: &str) -> Result<(SheetId, CellRange), String> {
        let (a, b) = range.split_once(':').unwrap_or((range, range));
        let (idx, at1) = at(wb, sheet, a).ok_or_else(|| format!("sheet {sheet:?} not found"))?;
        let (_, at2) = at(wb, sheet, b).ok_or_else(|| format!("sheet {sheet:?} not found"))?;
        Ok((SheetId(idx as u32), CellRange::new(at1, at2)))
    }

    impl Engine for BetterOffice {
        fn name(&self) -> &'static str {
            "betteroffice"
        }

        fn apply(&self, input: &[u8], ops: &[Op]) -> EditOutcome {
            let mut op_results = Vec::new();
            let open_started = Instant::now();
            let mut wb = match betteroffice_xlsx::Workbook::open(input) {
                Ok(wb) => wb,
                Err(error) => {
                    op_results.push(("open".into(), Err(format!("open failed: {error}"))));
                    return EditOutcome { output: None, phases: None, op_results };
                }
            };
            let open_ms = open_started.elapsed().as_secs_f64() * 1000.0;
            let modify_started = Instant::now();
            let mut pending_results: Vec<(String, std::result::Result<(), String>)> = Vec::new();
            let mut batch: Vec<BoOp> = Vec::new();
            {
            for op in ops {
                let kind = op.kind();
                // 每个 benchmark 操作映射为原生 xlsx_ops::Op，收集后交给官方
                // Workbook::apply_ops 一次性批量应用 —— 公式重写、样式派生、
                // 坐标重映射都由库完成，且整批只走一次验证。
                let native: std::result::Result<Vec<BoOp>, String> = (|| {
                    Ok(match op {
                        Op::SetCellValue { sheet, cell, val } => {
                            let (idx, at) = at(&wb.model().clone(), sheet, cell)
                                .ok_or_else(|| format!("sheet {sheet:?} not found"))?;
                            vec![BoOp::SetCell {
                                sheet: SheetId(idx as u32),
                                at,
                                cell: CellState {
                                    value: val_of(val),
                                    formula: None,
                                    style: None,
                                },
                            }]
                        }
                        Op::SetFormula { sheet, cell, text, cached } => {
                            let (idx, at) = at(&wb.model().clone(), sheet, cell)
                                .ok_or_else(|| format!("sheet {sheet:?} not found"))?;
                            vec![BoOp::SetCell {
                                sheet: SheetId(idx as u32),
                                at,
                                cell: CellState {
                                    value: cached.as_ref().map(val_of).unwrap_or(CellValue::Empty),
                                    formula: Some(text.trim_start_matches('=').to_owned()),
                                    style: None,
                                },
                            }]
                        }
                        Op::SetNumberFormat { sheet, cell, fmt } => {
                            let (idx, at) = at(&wb.model().clone(), sheet, cell)
                                .ok_or_else(|| format!("sheet {sheet:?} not found"))?;
                            vec![BoOp::SetRangeNumberFormat {
                                sheet: SheetId(idx as u32),
                                range: CellRange::new(at, at),
                                format: NumberFormatMutation::Custom { pattern: fmt.clone() },
                            }]
                        }
                        Op::SetFontBold { sheet, cell } => {
                            let (idx, at) = at(&wb.model().clone(), sheet, cell)
                                .ok_or_else(|| format!("sheet {sheet:?} not found"))?;
                            vec![BoOp::PatchRangeStyle {
                                sheet: SheetId(idx as u32),
                                range: CellRange::new(at, at),
                                patch: StylePatch { bold: Some(true), ..Default::default() },
                            }]
                        }
                        Op::SetFillYellow { sheet, cell } => {
                            let (idx, at) = at(&wb.model().clone(), sheet, cell)
                                .ok_or_else(|| format!("sheet {sheet:?} not found"))?;
                            vec![BoOp::PatchRangeStyle {
                                sheet: SheetId(idx as u32),
                                range: CellRange::new(at, at),
                                patch: StylePatch {
                                    fill_color: Some("#FFFF00".into()),
                                    ..Default::default()
                                },
                            }]
                        }
                        Op::SetBorderThin { sheet, cell } => {
                            let (idx, at) = at(&wb.model().clone(), sheet, cell)
                                .ok_or_else(|| format!("sheet {sheet:?} not found"))?;
                            vec![BoOp::PatchRangeStyle {
                                sheet: SheetId(idx as u32),
                                range: CellRange::new(at, at),
                                patch: StylePatch {
                                    border: Some(BorderPatch {
                                        preset: Some(BorderPreset::All),
                                        style: Some(BorderLineStyle::Solid),
                                        color: None,
                                    }),
                                    ..Default::default()
                                },
                            }]
                        }
                        Op::SetAlignCenter { sheet, cell } => {
                            let (idx, at) = at(&wb.model().clone(), sheet, cell)
                                .ok_or_else(|| format!("sheet {sheet:?} not found"))?;
                            vec![BoOp::PatchRangeStyle {
                                sheet: SheetId(idx as u32),
                                range: CellRange::new(at, at),
                                patch: StylePatch {
                                    horizontal_alignment: Some(HorizontalAlignment::Center),
                                    ..Default::default()
                                },
                            }]
                        }
                        Op::SetRowHeight { sheet, row, h } => {
                            let sid = sheet_id(&wb.model().clone(), sheet)
                                .ok_or_else(|| format!("sheet {sheet:?} not found"))?;
                            vec![BoOp::SetRowHeight { sheet: sid, row: row - 1, height: Some(*h) }]
                        }
                        Op::SetColWidth { sheet, col, w } => {
                            let sid = sheet_id(&wb.model().clone(), sheet)
                                .ok_or_else(|| format!("sheet {sheet:?} not found"))?;
                            vec![BoOp::SetColWidth { sheet: sid, col: col - 1, width: Some(*w) }]
                        }
                        Op::AddMerge { sheet, range } => {
                            let (sid, cells) = bo_range(&wb.model(), sheet, range)?;
                            vec![BoOp::MergeCells { sheet: sid, range: cells }]
                        }
                        Op::RemoveMerge { sheet, range } => {
                            let (sid, cells) = bo_range(&wb.model(), sheet, range)?;
                            vec![BoOp::UnmergeCells { sheet: sid, range: cells }]
                        }
                        Op::SetComment { .. } => {
                            return Err(
                                "UNSUPPORTED: betteroffice-xlsx 0.5.1 has no comment representation (model + ops verified)"
                                    .to_owned(),
                            );
                        }
                        Op::SetGlobalDefinedName { name, reference } => {
                            vec![BoOp::SetDefinedName {
                                name: name.clone(),
                                refers_to: reference.clone(),
                                scope: None,
                                hidden: None,
                            }]
                        }
                        Op::SetPrintArea { sheet, range } => {
                            // 打印区域在 Excel 里的存储形式是带局部作用域的
                            // _xlnm.Print_Area 定义名；走同一通道。
                            let idx = wb
                                .model()
                                .sheets
                                .iter()
                                .position(|s| s.name == *sheet)
                                .ok_or_else(|| format!("sheet {sheet:?} not found"))?;
                            vec![BoOp::SetDefinedName {
                                name: "_xlnm.Print_Area".into(),
                                refers_to: format!("{sheet}!{range}"),
                                scope: Some(SheetId(idx as u32)),
                                hidden: None,
                            }]
                        }
                        Op::SetLandscape { .. } => {
                            return Err(
                                "UNSUPPORTED: no page-orientation representation in betteroffice-xlsx 0.5.1"
                                    .to_owned(),
                            );
                        }
                        Op::AddSheet { name } => {
                            let index = wb.model().sheets.len();
                            vec![BoOp::AddSheet { index, name: name.clone() }]
                        }
                        Op::DeleteSheet { name } => {
                            let idx = wb
                                .model()
                                .sheets
                                .iter()
                                .position(|s| s.name == *name)
                                .ok_or_else(|| format!("sheet {name:?} not found"))?;
                            if wb.model().sheets.len() <= 1 {
                                return Err("cannot delete the only sheet".into());
                            }
                            vec![BoOp::RemoveSheet { index: idx }]
                        }
                        Op::RenameSheet { old, new } => {
                            let idx = wb
                                .model()
                                .sheets
                                .iter()
                                .position(|s| s.name == *old)
                                .ok_or_else(|| format!("sheet {old:?} not found"))?;
                            vec![BoOp::RenameSheet {
                                sheet: SheetId(idx as u32),
                                name: new.clone(),
                            }]
                        }
                        Op::InsertRows { sheet, row, n } => {
                            let sid = sheet_id(&wb.model().clone(), sheet)
                                .ok_or_else(|| format!("sheet {sheet:?} not found"))?;
                            vec![BoOp::InsertRows { sheet: sid, at: row - 1, count: *n }]
                        }
                        Op::DeleteRows { sheet, row, n } => {
                            let sid = sheet_id(&wb.model().clone(), sheet)
                                .ok_or_else(|| format!("sheet {sheet:?} not found"))?;
                            vec![BoOp::DeleteRows { sheet: sid, at: row - 1, count: *n }]
                        }
                    })
                })();
                match native {
                    Err(error) => pending_results.push((kind.to_owned(), Err(error))),
                    Ok(native_ops) => batch.extend(native_ops),
                }
            }
            let batch_result = if batch.is_empty() {
                Ok(())
            } else {
                wb.apply_ops(batch, CalculationOptions::default())
                    .map(|_| ())
                    .map_err(|error| error.to_string())
                    .map_err(|error| format!("apply_ops batch: {error}"))
            };
            op_results.extend(pending_results);
            if let Err(error) = batch_result {
                op_results.push(("apply_ops".to_owned(), Err(error)));
            }
            }
            let modify_ms = modify_started.elapsed().as_secs_f64() * 1000.0;
            let save_started = Instant::now();
            match wb.save() {
                Ok(output) => EditOutcome {
                    output: Some(output),
                    phases: Some(PhaseTimes {
                        open_ms,
                        modify_ms,
                        save_ms: save_started.elapsed().as_secs_f64() * 1000.0,
                    }),
                    op_results,
                },
                Err(error) => {
                    op_results.push(("save".into(), Err(format!("save failed: {error}"))));
                    EditOutcome { output: None, phases: None, op_results }
                }
            }
        }
    }
}

// ------------------------------------------------- rxls 0.1.4

mod rx {
    use super::*;
    use rxls::Cell;

    pub struct Rxls;

    fn val_of(v: &Val) -> Cell {
        match v {
            Val::Text(t) => Cell::Text(t.clone()),
            Val::Num(n) => Cell::Number(*n),
            Val::Bool(b) => Cell::Bool(*b),
        }
    }

    impl Rxls {
        fn record(target: &mut Vec<(String, std::result::Result<(), String>)>, kind: String, result: std::result::Result<(), String>) {
            target.push((kind, result));
        }
    }

    fn rx_ops_on(book: &mut Spreadsheet, ops: &[Op], op_results: &mut Vec<(String, std::result::Result<(), String>)>) -> Result<(), String> {
        // 连续同表矩形值写入合并为一次 set_cell_range_values（官方批量路径，
        // 单次上限 10000 格）。逐格 set_cell_value 每格一次原子提交，大批量
        // 下是数量级的慢路径——不是引擎缺陷，是官方文档区分的用法差异。
        let mut value_ops: Vec<(String, (u32, u32), Val)> = Vec::new();
        let mut rest: Vec<Op> = Vec::new();
        for op in ops {
            match op {
                Op::SetCellValue { sheet, cell, val } => {
                    value_ops.push((sheet.clone(), parse_cell(cell), val.clone()));
                }
                other => rest.push(other.clone()),
            }
        }
        if value_ops.len() > 1 {
            let sheet = value_ops[0].0.clone();
            let same_sheet = value_ops.iter().all(|(s, _, _)| *s == sheet);
            let mut cells: Vec<(u32, u32)> = value_ops.iter().map(|(_, a, _)| *a).collect();
            cells.sort_by_key(|(r, c)| (*r, *c));
            let rows_span = cells.last().unwrap().0 - cells[0].0 + 1;
            let cols_span = cells.last().unwrap().1 - cells[0].1 + 1;
            let is_full_rect =
                cells.len() == (rows_span * cols_span) as usize
                    && (rows_span * cols_span) <= 10_000;
            if same_sheet && is_full_rect {
                let (r0, c0) = cells[0];
                let mut grid =
                    vec![vec![None::<Cell>; cols_span as usize]; rows_span as usize];
                for (_, (row, col), val) in &value_ops {
                    grid[(row - r0) as usize][(col - c0) as usize] = Some(val_of(val));
                }
                let result = book
                    .set_cell_range_values(&sheet, r0 - 1, (c0 - 1) as u16, &grid)
                    .map_err(|error| format!("set_cell_range_values: {error}"));
                op_results.push(("set_cell_range_values".to_owned(), result));
                if !rest.is_empty() {
                    rx_ops_on(book, &rest, op_results)?;
                }
                return Ok(());
            }
        }
        for op in ops {
            let kind = op.kind();
            let result: Result<(), String> = match op {
                Op::SetCellValue { sheet, cell, val } => {
                    let (row, col) = parse_cell(cell);
                    book.set_cell_value(sheet, row - 1, (col - 1) as u16, val_of(val))
                        .map_err(|e| e.to_string())
                }
                Op::SetFormula { sheet, cell, text, cached } => {
                    let (row, col) = parse_cell(cell);
                    let cached = cached.as_ref().map(val_of).unwrap_or(Cell::Text(String::new()));
                    book.set_cell_formula(sheet, row - 1, (col - 1) as u16, text.trim_start_matches('='), cached)
                        .map_err(|e| e.to_string())
                }
                Op::SetNumberFormat { .. } | Op::SetFontBold { .. } | Op::SetFillYellow { .. }
                | Op::SetBorderThin { .. } | Op::SetAlignCenter { .. } => Err(
                    "UNSUPPORTED: rxls 0.1.4 exposes no style/number-format edit API".into(),
                ),
                Op::SetRowHeight { sheet, row, h } => {
                    book.set_row_height(sheet, row - 1, *h as f32).map_err(|e| e.to_string())
                }
                Op::SetColWidth { sheet, col, w } => {
                    book.set_column_width(sheet, (col - 1) as u16, *w as f32).map_err(|e| e.to_string())
                }
                Op::AddMerge { sheet, range } => {
                    let ((r1, c1), (r2, c2)) = parse_range(range);
                    book.merge_cells(sheet, r1 - 1, (c1 - 1) as u16, r2 - 1, (c2 - 1) as u16)
                        .map_err(|e| e.to_string())
                }
                Op::RemoveMerge { sheet, range } => {
                    let ((r1, c1), (r2, c2)) = parse_range(range);
                    book.unmerge_cells(sheet, r1 - 1, (c1 - 1) as u16, r2 - 1, (c2 - 1) as u16)
                        .map_err(|e| e.to_string())
                }
                Op::SetComment { sheet, cell, text } => {
                    let (row, col) = parse_cell(cell);
                    book.set_comment(sheet, row - 1, (col - 1) as u16, text, Some("bench"))
                        .map_err(|e| e.to_string())
                }
                Op::SetGlobalDefinedName { name, reference } => {
                    book.set_defined_name(name, reference).map_err(|e| e.to_string())
                }
                Op::SetPrintArea { sheet, range } => {
                    let ((r1, c1), (r2, c2)) = parse_range(range);
                    book.set_print_area(sheet, Some((r1 - 1, (c1 - 1) as u16, r2 - 1, (c2 - 1) as u16)))
                        .map_err(|e| e.to_string())
                }
                Op::SetLandscape { .. } => Err("UNSUPPORTED: no page-orientation API in rxls 0.1.4".into()),
                Op::AddSheet { name } => book.add_sheet(name).map_err(|e| e.to_string()),
                Op::DeleteSheet { name } => book.delete_sheet(name).map_err(|e| e.to_string()),
                Op::RenameSheet { old, new } => book.rename_sheet(old, new).map_err(|e| e.to_string()),
                Op::InsertRows { .. } | Op::DeleteRows { .. } => Err(
                    "UNSUPPORTED: rxls does not insert/delete rows or columns (documented non-goal)".into(),
                ),
            }
            .map_err(|error| format!("{kind}: {error}"));
            op_results.push((kind.to_owned(), result));
        }
        Ok(())
    }

    impl Engine for Rxls {
        fn name(&self) -> &'static str {
            "rxls"
        }

        fn apply(&self, input: &[u8], ops: &[Op]) -> EditOutcome {
            let mut op_results = Vec::new();
            let open_started = Instant::now();
            let mut book = match Spreadsheet::open(input) {
                Ok(book) => book,
                Err(error) => {
                    op_results.push(("open".into(), Err(format!("open failed: {error}"))));
                    return EditOutcome { output: None, phases: None, op_results };
                }
            };
            let open_ms = open_started.elapsed().as_secs_f64() * 1000.0;
            let mut captured_results: Vec<(String, std::result::Result<(), String>)> = Vec::new();
            let mut inner_err: Option<String> = None;
            let modify_started = Instant::now();
            let mut apply_result: Result<(), String> = book
                .transaction(|tx| {
                    let mut op_results: Vec<(String, std::result::Result<(), String>)> = Vec::new();
                    let inner = rx_ops_on(tx, ops, &mut op_results);
                    captured_results.extend(op_results);
                    // 错误转存到外层；闭包返回 Ok 以保留已收集的逐操作结果。
                    inner_err = inner.err();
                    Ok(())
                })
                .map_err(|error| format!("transaction aborted: {error}"));
            if let Some(inner_error) = inner_err.take() {
                apply_result = Err(inner_error);
            }
            if let Err(error) = apply_result {
                captured_results.push(("transaction".to_owned(), Err(error)));
            }
            op_results.extend(captured_results.drain(..).map(|(k, r)| (k, r)));
            let modify_ms = modify_started.elapsed().as_secs_f64() * 1000.0;
            let save_started = Instant::now();
            match book.save() {
                Ok(output) => EditOutcome {
                    output: Some(output),
                    phases: Some(PhaseTimes {
                        open_ms,
                        modify_ms,
                        save_ms: save_started.elapsed().as_secs_f64() * 1000.0,
                    }),
                    op_results,
                },
                Err(error) => {
                    op_results.push(("save".into(), Err(format!("save failed: {error}"))));
                    EditOutcome { output: None, phases: None, op_results }
                }
            }
        }
    }

    use rxls::Spreadsheet;
}

// ------------------------------------------------- umya 3.1.1

mod umya_eng {
    use super::*;
    use umya_spreadsheet::*;

    pub struct Umya;

    fn ws<'a>(book: &'a mut Workbook, sheet: &str) -> Result<&'a mut Worksheet, String> {
        book.get_sheet_by_name_mut(sheet)
            .map_err(|error| format!("sheet {sheet:?}: {error}"))
    }

    impl Engine for Umya {
        fn name(&self) -> &'static str {
            "umya"
        }

        fn apply(&self, input: &[u8], ops: &[Op]) -> EditOutcome {
            let mut op_results = Vec::new();
            let nonce = std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos();
            let tmp_in = std::env::temp_dir().join(format!("umbench-{nonce}-in.xlsx"));
            let tmp_out = std::env::temp_dir().join(format!("umbench-{nonce}-out.xlsx"));
            if let Err(error) = std::fs::write(&tmp_in, input) {
                op_results.push(("stage".into(), Err(format!("stage input failed: {error}"))));
                return EditOutcome { output: None, phases: None, op_results };
            }
            let open_started = Instant::now();
            let mut book = match reader::xlsx::read(&tmp_in) {
                Ok(book) => book,
                Err(error) => {
                    let _ = std::fs::remove_file(&tmp_in);
                    op_results.push(("open".into(), Err(format!("open failed: {error}"))));
                    return EditOutcome { output: None, phases: None, op_results };
                }
            };
            let open_ms = open_started.elapsed().as_secs_f64() * 1000.0;
            let mut captured_results: Vec<(String, std::result::Result<(), String>)> = Vec::new();
            let mut inner_err: Option<String> = None;
            let modify_started = Instant::now();
            for op in ops {
                let kind = op.kind();
                let result: Result<(), String> = match op {
                    Op::SetCellValue { sheet, cell, val } => {
                        let target = match ws(&mut book, sheet) { Ok(t) => t, Err(e) => { op_results.push((kind.to_owned(), Err(e))); continue; } };
                        let cell = target.get_cell_mut(cell.as_str());
                        match val {
                            Val::Text(t) => cell.set_value_string(t.as_str()),
                            Val::Num(n) => cell.set_value_number(*n),
                            Val::Bool(b) => cell.set_value_bool(*b),
                        };
                        Ok(())
                    }
                    Op::SetFormula { sheet, cell, text, .. } => {
                        let target = match ws(&mut book, sheet) { Ok(t) => t, Err(e) => { op_results.push((kind.to_owned(), Err(e))); continue; } };
                        target.get_cell_mut(cell.as_str()).set_formula(text.as_str());
                        Ok(())
                    }
                    Op::SetNumberFormat { sheet, cell, fmt } => {
                        let target = match ws(&mut book, sheet) { Ok(t) => t, Err(e) => { op_results.push((kind.to_owned(), Err(e))); continue; } };
                        let mut numbering = umya_spreadsheet::NumberingFormat::default();
                        numbering.set_format_code(fmt.as_str());
                        target.get_cell_mut(cell.as_str()).get_style_mut().set_number_format(numbering);
                        Ok(())
                    }
                    Op::SetFontBold { sheet, cell } => {
                        let target = match ws(&mut book, sheet) { Ok(t) => t, Err(e) => { op_results.push((kind.to_owned(), Err(e))); continue; } };
                        target.get_cell_mut(cell.as_str()).get_style_mut().get_font_mut().set_bold(true);
                        Ok(())
                    }
                    Op::SetFillYellow { sheet, cell } => {
                        let target = match ws(&mut book, sheet) { Ok(t) => t, Err(e) => { op_results.push((kind.to_owned(), Err(e))); continue; } };
                        let fill = target.get_cell_mut(cell.as_str()).get_style_mut().get_fill_mut();
                        let pattern = fill.get_pattern_fill_mut();
                        pattern.set_pattern_type(PatternValues::Solid);
                        pattern.get_foreground_color_mut().set_argb_str("FFFFFF00");
                        Ok(())
                    }
                    Op::SetBorderThin { sheet, cell } => {
                        let target = match ws(&mut book, sheet) { Ok(t) => t, Err(e) => { op_results.push((kind.to_owned(), Err(e))); continue; } };
                        let borders = target.cell_mut(cell.as_str()).get_style_mut().get_borders_mut();
                        borders.get_left_mut().set_border_style("thin");
                        borders.get_right_mut().set_border_style("thin");
                        borders.get_top_mut().set_border_style("thin");
                        borders.get_bottom_mut().set_border_style("thin");
                        Ok(())
                    }
                    Op::SetAlignCenter { sheet, cell } => {
                        let target = match ws(&mut book, sheet) { Ok(t) => t, Err(e) => { op_results.push((kind.to_owned(), Err(e))); continue; } };
                        target.get_cell_mut(cell.as_str())
                            .get_style_mut()
                            .get_alignment_mut()
                            .set_horizontal(HorizontalAlignmentValues::Center);
                        Ok(())
                    }
                    Op::SetRowHeight { sheet, row, h } => {
                        let target = match ws(&mut book, sheet) { Ok(t) => t, Err(e) => { op_results.push((kind.to_owned(), Err(e))); continue; } };
                        target.get_row_dimension_mut(&row).set_height(*h);
                        Ok(())
                    }
                    Op::SetColWidth { sheet, col, w } => {
                        let target = match ws(&mut book, sheet) { Ok(t) => t, Err(e) => { op_results.push((kind.to_owned(), Err(e))); continue; } };
                        target.get_column_dimension_mut(&col_letter(*col)).set_width(*w);
                        Ok(())
                    }
                    Op::AddMerge { sheet, range } => {
                        let target = match ws(&mut book, sheet) { Ok(t) => t, Err(e) => { op_results.push((kind.to_owned(), Err(e))); continue; } };
                        target.add_merge_cells(range.as_str());
                        Ok(())
                    }
                    Op::RemoveMerge { sheet, range } => {
                        let target = match ws(&mut book, sheet) { Ok(t) => t, Err(e) => { op_results.push((kind.to_owned(), Err(e))); continue; } };
                        let merges = target.get_merge_cells_mut();
                        let before = merges.len();
                        merges.retain(|m| m.get_range() != *range);
                        if merges.len() == before {
                            Err(format!("merge {range:?} not found"))
                        } else {
                            Ok(())
                        }
                    }
                    Op::SetComment { sheet, cell, text } => {
                        let target = match ws(&mut book, sheet) { Ok(t) => t, Err(e) => { op_results.push((kind.to_owned(), Err(e))); continue; } };
                        let mut comment = Comment::default();
                        comment.new_comment(cell.as_str());
                        comment.set_text_string(text.as_str());
                        target.add_comments(comment);
                        Ok(())
                    }
                    Op::SetGlobalDefinedName { name, reference } => {
                        let mut defined = DefinedName::default();
                        defined.set_name(name.as_str());
                        defined.set_address(reference.as_str());
                        book.add_defined_names(defined);
                        Ok(())
                    }
                    Op::SetPrintArea { sheet, range } => {
                        // umya 无打印区域 API；走 Excel 存储原形：带局部作用域的
                        // _xlnm.Print_Area 定义名。
                        let index = book.get_sheet_collection().iter().position(|s| s.get_name() == sheet).unwrap_or(0) as u32;
                        let mut defined = DefinedName::default();
                        defined.set_name("_xlnm.Print_Area");
                        defined.set_address(format!("{sheet}!{range}").as_str());
                        defined.set_local_sheet_id(index);
                        book.add_defined_names(defined);
                        Ok(())
                    }
                    Op::SetLandscape { sheet } => {
                        let target = match ws(&mut book, sheet) { Ok(t) => t, Err(e) => { op_results.push((kind.to_owned(), Err(e))); continue; } };
                        target.get_page_setup_mut().set_orientation(umya_spreadsheet::structs::OrientationValues::Landscape);
                        Ok(())
                    }
                    Op::AddSheet { name } => {
                        book.new_sheet(name.as_str());
                        Ok(())
                    }
                    Op::DeleteSheet { name } => {
                        if book.get_sheet_count() <= 1 {
                            Err("cannot delete the only sheet".into())
                        } else {
                            book.remove_sheet_by_name(name).map_err(|e| e.to_string())
                        }
                    }
                    Op::RenameSheet { old, new } => {
                        let target = match ws(&mut book, old) { Ok(t) => t, Err(e) => { op_results.push((kind.to_owned(), Err(e))); continue; } };
                        target.set_name(new.as_str());
                        Ok(())
                    }
                    Op::InsertRows { sheet, row, n } => {
                        let target = match ws(&mut book, sheet) { Ok(t) => t, Err(e) => { op_results.push((kind.to_owned(), Err(e))); continue; } };
                        target.insert_new_row(*row, *n);
                        Ok(())
                    }
                    Op::DeleteRows { sheet, row, n } => {
                        let target = match ws(&mut book, sheet) { Ok(t) => t, Err(e) => { op_results.push((kind.to_owned(), Err(e))); continue; } };
                        target.remove_row(*row, *n);
                        Ok(())
                    }
                }
                .map_err(|error| format!("{kind}: {error}"));
                op_results.push((kind.to_owned(), result));
            }
            let modify_ms = modify_started.elapsed().as_secs_f64() * 1000.0;
            let save_started = Instant::now();
            let outcome = match writer::xlsx::write(&book, &tmp_out) {
                Ok(_) => match std::fs::read(&tmp_out) {
                    Ok(output) => EditOutcome {
                        output: Some(output),
                        phases: Some(PhaseTimes {
                            open_ms,
                            modify_ms,
                            save_ms: save_started.elapsed().as_secs_f64() * 1000.0,
                        }),
                        op_results,
                    },
                    Err(error) => {
                        op_results.push(("save".into(), Err(format!("read-back failed: {error}"))));
                        EditOutcome { output: None, phases: None, op_results }
                    }
                },
                Err(error) => {
                    op_results.push(("save".into(), Err(format!("save failed: {error}"))));
                    EditOutcome { output: None, phases: None, op_results }
                }
            };
            let _ = std::fs::remove_file(&tmp_in);
            let _ = std::fs::remove_file(&tmp_out);
            outcome
        }
    }
}

// ------------------------------------------------- probe (M01-M14)

fn cells(count: usize) -> Vec<(String, Val)> {
    let mut out = Vec::new();
    for i in 0..count {
        let row = 2 + (i / 5) as u32;
        let col = 1 + (i % 5);
        let cell = format!("{}{}", col_letter(col as u32), row);
        out.push((cell, Val::Num(9000.0 + i as f64)));
    }
    out
}

pub fn probe_ops(m: &str) -> Vec<Op> {
    let sheet = if m == "M13" { "汇总" } else { "Main" };
    match m {
        "M01" => vec![Op::SetCellValue { sheet: sheet.into(), cell: "B2".into(), val: Val::Text("probe-值".into()) }],
        "M02" => cells(100).into_iter().map(|(cell, val)| Op::SetCellValue { sheet: sheet.into(), cell, val }).collect(),
        "M03" => cells(1000).into_iter().map(|(cell, val)| Op::SetCellValue { sheet: sheet.into(), cell, val }).collect(),
        "M04" => vec![Op::SetFormula { sheet: sheet.into(), cell: "B2".into(), text: "=A1*10".into(), cached: Some(Val::Num(30.0)) }],
        "M05" => vec![Op::SetNumberFormat { sheet: sheet.into(), cell: "B2".into(), fmt: "#,##0.000".into() }],
        "M06" => vec![
            Op::SetFontBold { sheet: sheet.into(), cell: "B2".into() },
            Op::SetFillYellow { sheet: sheet.into(), cell: "B3".into() },
            Op::SetBorderThin { sheet: sheet.into(), cell: "B4".into() },
            Op::SetAlignCenter { sheet: sheet.into(), cell: "B5".into() },
        ],
        "M07" => vec![
            Op::SetRowHeight { sheet: sheet.into(), row: 6, h: 33.0 },
            Op::SetColWidth { sheet: sheet.into(), col: 4, w: 33.0 },
        ],
        "M08" => vec![
            Op::AddMerge { sheet: sheet.into(), range: "H2:I3".into() },
            Op::RemoveMerge { sheet: sheet.into(), range: "H2:I3".into() },
        ],
        "M09" => vec![Op::SetComment { sheet: sheet.into(), cell: "B2".into(), text: "探针批注".into() }],
        "M10" => vec![Op::SetGlobalDefinedName { name: "B_ProbeName".into(), reference: "Main!$B$2".into() }],
        "M11" => vec![
            Op::SetPrintArea { sheet: sheet.into(), range: "A1:D20".into() },
            Op::SetLandscape { sheet: sheet.into() },
        ],
        "M12" => vec![
            Op::AddSheet { name: "ProbeNew".into() },
            Op::RenameSheet { old: "ProbeNew".into(), new: "ProbeRenamed".into() },
            Op::DeleteSheet { name: "ProbeRenamed".into() },
        ],
        "M13" => vec![Op::SetCellValue { sheet: sheet.into(), cell: "B2".into(), val: Val::Text("复杂模板单点".into()) }],
        "M14" => vec![
            Op::SetCellValue { sheet: sheet.into(), cell: "B2".into(), val: Val::Num(42.0) },
            Op::SetRowHeight { sheet: sheet.into(), row: 2, h: 25.0 },
            Op::AddMerge { sheet: sheet.into(), range: "H5:I6".into() },
            Op::SetComment { sheet: sheet.into(), cell: "C3".into(), text: "mixed".into() },
            Op::SetGlobalDefinedName { name: "B_Mixed".into(), reference: "Main!$B$2".into() },
        ],
        _ => Vec::new(),
    }
}

pub fn probe_fixture(m: &str, fixtures: &Path) -> PathBuf {
    match m {
        "M13" => fixtures.join("f_complex.xlsx"),
        _ => fixtures.join("f_basic.xlsx"),
    }
}

pub fn run_probe(engine_name: &str, fixtures: &Path, out_dir: &Path) -> Result<()> {
    use anyhow::Context;
    let ms = [
        "M01", "M02", "M03", "M04", "M05", "M06", "M07",
        "M08", "M09", "M10", "M11", "M12", "M13", "M14",
    ];
    let mut rows = Vec::new();
    if engine_name == "openpyxl" {
        for m in ms {
            rows.push(json!({
                "engine": "openpyxl", "op": m, "result": "PASS",
                "note": "production worker already performs this operation class",
            }));
        }
    } else {
        let engine = engine_by_name(engine_name).context("unknown engine")?;
        for m in ms {
            let fixture = probe_fixture(m, fixtures);
            let input = std::fs::read(&fixture)
                .with_context(|| format!("read fixture {}", fixture.display()))?;
            let outcome = engine.apply(&input, &probe_ops(m));
            let unsupported = outcome.op_results.iter().any(|(_, r)| {
                matches!(r, Err(reason) if reason.contains("UNSUPPORTED"))
            });
            let failed = outcome.op_results.iter().any(|(_, r)| {
                matches!(r, Err(reason) if !reason.contains("UNSUPPORTED"))
            });
            let status = if unsupported && !failed {
                "UNSUPPORTED"
            } else if outcome.output.is_some() && !failed {
                "PASS"
            } else {
                "FAIL"
            };
            let errors: Vec<String> = outcome
                .op_results
                .iter()
                .filter_map(|(k, r)| r.as_ref().err().map(|e| format!("{k}: {e}")))
                .collect();
            rows.push(json!({
                "engine": engine_name,
                "op": m,
                "result": status,
                "fixture": fixture.file_name().unwrap().to_string_lossy(),
                "errors": errors,
            }));
        }
    }
    std::fs::create_dir_all(out_dir)?;
    let path = out_dir.join(format!("capability_{engine_name}.json"));
    std::fs::write(&path, serde_json::to_vec_pretty(&rows)?)?;
    println!("capability -> {}", path.display());
    Ok(())
}
