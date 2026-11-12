//! Independent-reader comparison of converted XLSX output against the source
//! XLS. Values/types/formulas are read with calamine; structure (merges,
//! dimensions, print setup, feature counts, styles) is read with rxls.
//! Both readers are applied to both formats so no backend grades itself.

use crate::common::{sha256_hex, TOOL_VERSION};
use calamine::{Data, Reader, Xls, Xlsx};
use serde::Serialize;
use std::collections::BTreeMap;
use std::path::{Path, PathBuf};
use std::str::FromStr as _;

pub const MAX_SAMPLES_DEFAULT: usize = 12;

/// calamine cell coordinates: `(u32, u32)`.
type CalKey = (u32, u32);
/// rxls cell coordinates: `(u32, u16)`.
type RxKey = (u32, u16);
type BufReaderType = std::io::BufReader<std::fs::File>;

/// One comparison category for a file.
#[derive(Debug, Default, Serialize)]
pub struct Cat {
    pub total: u32,
    pub ok: u32,
    pub mismatch: u32,
    /// Present in source, absent in converted output.
    pub missing: u32,
    /// Present in converted output, absent in source.
    pub extra: u32,
    /// Values equal but the reader-visible type tag differs (e.g. date serial
    /// stored without a date format, error value written as text). Counted
    /// separately from hard mismatches because the stored value survived.
    pub type_only: u32,
    pub samples: Vec<String>,
}

impl Cat {
    fn note_ok(&mut self) {
        self.total += 1;
        self.ok += 1;
    }
    fn note_mismatch(&mut self, cap: usize, sample: String) {
        self.total += 1;
        self.mismatch += 1;
        if self.samples.len() < cap {
            self.samples.push(sample);
        }
    }
    fn note_missing(&mut self, cap: usize, sample: String) {
        self.total += 1;
        self.missing += 1;
        if self.samples.len() < cap {
            self.samples.push(sample);
        }
    }
    fn note_extra(&mut self, cap: usize, sample: String) {
        self.total += 1;
        self.extra += 1;
        if self.samples.len() < cap {
            self.samples.push(sample);
        }
    }
    fn note_type_only(&mut self, cap: usize, sample: String) {
        self.total += 1;
        self.type_only += 1;
        if self.samples.len() < cap {
            self.samples.push(sample);
        }
    }
    fn clean(&self) -> bool {
        self.mismatch == 0 && self.missing == 0 && self.extra == 0
    }
}

#[derive(Debug, Default, Serialize)]
pub struct FileCompare {
    pub file: String,
    /// `pass` | `partial` | `fail` | `failed` (output missing/unreadable)
    pub verdict: String,
    pub sheets: Cat,
    pub dims: Cat,
    pub data: Cat,
    pub formulas: Cat,
    pub structure: Cat,
    pub print: Cat,
    pub style: Cat,
    pub notes: Vec<String>,
}

#[derive(Debug, Serialize)]
pub struct BackendCompare {
    pub label: String,
    pub files: Vec<FileCompare>,
    pub pass_count: usize,
    pub partial_count: usize,
    pub fail_count: usize,
}

#[derive(Debug, Serialize)]
pub struct CompareReport {
    pub tool_version: String,
    pub source: String,
    pub backends: Vec<BackendCompare>,
}

/// Cell style reduced to five comparable components (each the Debug rendering
/// of the corresponding rxls model type).
#[derive(Debug, Default, PartialEq, Eq, Clone)]
struct StyleKey {
    font: String,
    fill: String,
    border: String,
    num_fmt: String,
    align: String,
}

impl StyleKey {
    fn from(style: &rxls::CellStyle) -> Self {
        let fill = match (&style.fill, &style.pattern_fill) {
            (Some(color), Some(pattern)) => format!("{color:?}/{pattern:?}"),
            (Some(color), None) => format!("{color:?}"),
            (None, Some(pattern)) => format!("{pattern:?}"),
            (None, None) => String::new(),
        };
        StyleKey {
            font: style.font.as_ref().map(|f| format!("{f:?}")).unwrap_or_default(),
            fill,
            border: style.border.as_ref().map(|b| format!("{b:?}")).unwrap_or_default(),
            num_fmt: style.num_fmt.clone().unwrap_or_default(),
            align: style.align.as_ref().map(|a| format!("{a:?}")).unwrap_or_default(),
        }
    }

    fn render(&self) -> String {
        format!(
            "font[{}] fill[{}] border[{}] num_fmt[{}] align[{}]",
            self.font, self.fill, self.border, self.num_fmt, self.align
        )
    }
}

#[derive(Debug, Default)]
struct SheetCal {
    values: BTreeMap<CalKey, Data>,
    formulas: BTreeMap<CalKey, String>,
    start: Option<CalKey>,
    end: Option<CalKey>,
}

#[derive(Debug, Default)]
struct CalamineBook {
    names: Vec<String>,
    sheets: BTreeMap<String, SheetCal>,
}

struct RxlsSheetInfo {
    name: String,
    is_worksheet: bool,
    visible: String,
    merges: Vec<(u32, u16, u32, u16)>,
    col_widths: BTreeMap<u16, f32>,
    row_heights: BTreeMap<u32, f32>,
    hidden_rows: Vec<u32>,
    hidden_cols: Vec<u16>,
    freeze: Option<(u32, u16)>,
    page_setup: Option<String>,
    print_gridlines: bool,
    print_headings: bool,
    counts: String,
    style_fidelity: String,
}

#[derive(Debug, Default)]
struct SheetCounts {
    comments: usize,
    data_validations: usize,
    conditional_formats: usize,
    hyperlinks: usize,
    images: usize,
    charts: usize,
    tables: usize,
}

struct RxlsBookInfo {
    defined_names: Vec<String>,
    local_defined_names: Vec<String>,
    sheets: Vec<RxlsSheetInfo>,
    style_cells: BTreeMap<String, BTreeMap<RxKey, StyleKey>>,
}

pub fn run(
    source: &Path,
    converted: Vec<(String, PathBuf)>,
    out_path: &Path,
    samples_cap: usize,
    with_styles: bool,
    values_reader: &str,
) -> anyhow::Result<()> {
    let sources: Vec<(String, PathBuf)> = if source.is_dir() {
        let mut list = Vec::new();
        for entry in std::fs::read_dir(source)? {
            let entry = entry?;
            let path = entry.path();
            let is_xls = path
                .extension()
                .map(|e| e.eq_ignore_ascii_case("xls"))
                .unwrap_or(false);
            if path.is_file() && is_xls {
                list.push((
                    path.file_stem()
                        .unwrap_or_default()
                        .to_string_lossy()
                        .into_owned(),
                    path,
                ));
            }
        }
        list.sort();
        list
    } else {
        vec![(
            source
                .file_stem()
                .unwrap_or_default()
                .to_string_lossy()
                .into_owned(),
            source.to_path_buf(),
        )]
    };

    let mut backends = Vec::new();
    for (label, conv_path) in &converted {
        let mut files = Vec::new();
        for (stem, src) in &sources {
            let candidate: PathBuf = if conv_path.is_dir() {
                conv_path.join(format!("{stem}.xlsx"))
            } else {
                conv_path.clone()
            };
            files.push(compare_one(
                stem,
                src,
                &candidate,
                samples_cap,
                with_styles,
                label.contains("values"),
                values_reader,
            ));
        }
        let pass_count = files.iter().filter(|f| f.verdict == "pass").count();
        let partial_count = files.iter().filter(|f| f.verdict == "partial").count();
        let fail_count = files
            .iter()
            .filter(|f| f.verdict != "pass" && f.verdict != "partial")
            .count();
        backends.push(BackendCompare {
            label: label.clone(),
            files,
            pass_count,
            partial_count,
            fail_count,
        });
    }

    let report = CompareReport {
        tool_version: TOOL_VERSION.to_owned(),
        source: source.to_string_lossy().into_owned(),
        backends,
    };
    if let Some(parent) = out_path.parent() {
        std::fs::create_dir_all(parent)?;
    }
    let json = serde_json::to_string_pretty(&report)?;
    std::fs::write(out_path, json)?;
    eprintln!(
        "compare report written: {} ({} backends)",
        out_path.display(),
        report.backends.len()
    );
    Ok(())
}

fn compare_one(
    stem: &str,
    source: &Path,
    converted: &Path,
    cap: usize,
    with_styles: bool,
    values_only: bool,
    values_reader: &str,
) -> FileCompare {
    let mut fc = FileCompare {
        file: stem.to_owned(),
        verdict: "fail".to_owned(),
        ..Default::default()
    };
    if !converted.exists() {
        fc.verdict = "failed".to_owned();
        fc.notes
            .push(format!("converted output missing: {}", converted.display()));
        return fc;
    }
    let source_bytes = match std::fs::read(source) {
        Ok(bytes) => bytes,
        Err(e) => {
            fc.verdict = "failed".to_owned();
            fc.notes.push(format!("source unreadable: {e}"));
            return fc;
        }
    };
    let converted_bytes = match std::fs::read(converted) {
        Ok(bytes) => bytes,
        Err(e) => {
            fc.verdict = "failed".to_owned();
            fc.notes.push(format!("converted unreadable: {e}"));
            return fc;
        }
    };
    fc.notes.push(format!("sha256 source {}", sha256_hex(&source_bytes)));
    fc.notes
        .push(format!("sha256 converted {}", sha256_hex(&converted_bytes)));

    let (source_cal, converted_cal) = if values_reader == "rxls" {
        // rxls-view: both sides read through the rxls model. Used for legacy
        // codepage sources where the calamine XLS reader mis-decodes strings
        // (e.g. BIFF5 GBK names), while rxls decodes them correctly.
        let source = rxls_values(&source_bytes);
        let converted = rxls_values(&converted_bytes);
        match (source, converted) {
            (Ok(a), Ok(b)) => (a, b),
            (Err(e), _) | (_, Err(e)) => {
                fc.verdict = "failed".to_owned();
                fc.notes.push(e);
                return fc;
            }
        }
    } else {
        let source_cal = read_calamine_xls(source);
        let converted_cal = read_calamine_xlsx(converted);
        match (source_cal, converted_cal) {
            (Ok(a), Ok(b)) => (a, b),
            (Err(e), _) | (_, Err(e)) => {
                fc.verdict = "failed".to_owned();
                fc.notes.push(e);
                return fc;
            }
        }
    };
    let source_rx = read_rxls(&source_bytes, with_styles);
    let converted_rx = read_rxls(&converted_bytes, with_styles);
    let (source_rx, converted_rx) = match (source_rx, converted_rx) {
        (Ok(a), Ok(b)) => (a, b),
        (Err(e), _) | (_, Err(e)) => {
            fc.verdict = "failed".to_owned();
            fc.notes.push(e);
            return fc;
        }
    };

    // Chartsheets have no grid; the XLSX reader may or may not surface them,
    // so they are excluded from the sheet/data comparison and only tracked in
    // the structure category.
    let is_grid = |book: &RxlsBookInfo, name: &str| {
        book.sheets
            .iter()
            .find(|s| s.name == name)
            .map(|s| s.is_worksheet)
            .unwrap_or(true)
    };
    let chartsheet_names: Vec<String> = source_rx
        .sheets
        .iter()
        .filter(|s| !s.is_worksheet)
        .map(|s| s.name.clone())
        .collect();
    if !chartsheet_names.is_empty() {
        fc.notes.push(format!(
            "chartsheets excluded from grid comparison: {:?}",
            chartsheet_names
        ));
    }
    let grid_sheets: Vec<String> = source_cal
        .names
        .iter()
        .filter(|name| is_grid(&source_rx, name) && is_grid(&converted_rx, name))
        .cloned()
        .collect();

    compare_sheets(
        &source_cal,
        &converted_cal,
        cap,
        &mut fc.sheets,
        &chartsheet_names,
    );

    for name in &grid_sheets {
        let (Some(src_sheet), Some(conv_sheet)) = (
            source_cal.sheets.get(name),
            converted_cal.sheets.get(name),
        ) else {
            continue;
        };
        compare_dims(src_sheet, conv_sheet, cap, &mut fc.dims);
        compare_data(src_sheet, conv_sheet, cap, &mut fc.data);
        compare_formulas(src_sheet, conv_sheet, cap, &mut fc.formulas);
    }

    compare_structure(&source_rx, &converted_rx, cap, &mut fc.structure);
    compare_print(&source_rx, &converted_rx, cap, &mut fc.print);
    if with_styles {
        compare_styles(&source_rx, &converted_rx, cap, &mut fc.style);
    } else {
        fc.notes
            .push("style comparison skipped (--styles not set)".to_owned());
    }

    let hard_bad = if values_only {
        // A values-only backend does not claim formula migration; the formula
        // category stays informational there (counts remain in the report).
        !fc.sheets.clean() || !fc.data.clean()
    } else {
        !fc.sheets.clean() || !fc.data.clean() || !fc.formulas.clean()
    };
    if values_only && fc.formulas.total > 0 {
        fc.notes
            .push("formulas not migrated by design (values-only backend)".to_owned());
    }
    let soft_bad =
        !fc.dims.clean() || !fc.structure.clean() || !fc.print.clean() || !fc.style.clean();
    fc.verdict = if hard_bad {
        "fail".to_owned()
    } else if soft_bad {
        "partial".to_owned()
    } else {
        "pass".to_owned()
    };
    fc
}

/// Build the sheet values/formulas view from the rxls model (both sides).
/// Dates map to their serial float; formula cells contribute their cached
/// value plus the decompiled formula text.
fn rxls_values(bytes: &[u8]) -> Result<CalamineBook, String> {
    let book = rxls::Workbook::open(bytes).map_err(|e| format!("rxls open failed: {e}"))?;
    let mut out = CalamineBook::default();
    for sheet in &book.sheets {
        let mut view = SheetCal::default();
        for (row, col, cell) in sheet.cells() {
            let (value, formula): (Option<Data>, Option<String>) = match cell {
                rxls::Cell::Text(t) => (Some(Data::String(t.clone())), None),
                rxls::Cell::Number(n) => (Some(Data::Float(*n)), None),
                rxls::Cell::Date(serial) => (Some(Data::Float(*serial)), None),
                rxls::Cell::Bool(b) => (Some(Data::Bool(*b)), None),
                rxls::Cell::Error(e) => (
                    Some(Data::Error(
                        calamine::CellErrorType::from_str(e)
                            .unwrap_or(calamine::CellErrorType::NA),
                    )),
                    None,
                ),
                rxls::Cell::Formula { formula, cached } => {
                    let cached_value = match cached.as_ref() {
                        rxls::Cell::Text(t) => Data::String(t.clone()),
                        rxls::Cell::Number(n) => Data::Float(*n),
                        rxls::Cell::Date(serial) => Data::Float(*serial),
                        rxls::Cell::Bool(b) => Data::Bool(*b),
                        rxls::Cell::Error(e) => {
                            Data::Error(
                                calamine::CellErrorType::from_str(e)
                                    .unwrap_or(calamine::CellErrorType::NA),
                            )
                        }
                        _ => Data::Empty,
                    };
                    (Some(cached_value), Some(formula.clone()))
                }
            };
            if let Some(value) = value {
                if !matches!(value, Data::Empty) {
                    view.values.insert((row, col as u32), value);
                }
            }
            if let Some(formula) = formula {
                view.formulas.insert((row, col as u32), formula);
            }
        }
        out.names.push(sheet.name.clone());
        out.sheets.insert(sheet.name.clone(), view);
    }
    Ok(out)
}

fn read_calamine_xls(path: &Path) -> Result<CalamineBook, String> {
    let reader = make_reader(path);
    let book = Xls::<BufReaderType>::new(reader)
        .map_err(|e| format!("calamine XLS open failed: {e}"))?;
    finish_calamine(book)
}

fn read_calamine_xlsx(path: &Path) -> Result<CalamineBook, String> {
    let book: Xlsx<BufReaderType> = calamine::open_workbook(path)
        .map_err(|e| format!("calamine XLSX open failed: {e}"))?;
    finish_calamine(book)
}

fn make_reader(path: &Path) -> BufReaderType {
    match std::fs::File::open(path) {
        Ok(file) => BufReaderType::new(file),
        Err(e) => unreachable!("file presence verified before read: {e}"),
    }
}

fn finish_calamine<R>(mut book: R) -> Result<CalamineBook, String>
where
    R: calamine::Reader<BufReaderType>,
    R::Error: std::fmt::Debug,
{
    let mut out = CalamineBook::default();
    for name in book.sheet_names() {
        let values = book
            .worksheet_range(&name)
            .map_err(|e| format!("calamine range '{name}' failed: {e:?}"))?;
        let formulas = book
            .worksheet_formula(&name)
            .map_err(|e| format!("calamine formula '{name}' failed: {e:?}"))?;
        let mut sheet = SheetCal {
            start: values.start(),
            end: values.end(),
            ..Default::default()
        };
        for (row, col, value) in values.used_cells() {
            sheet.values.insert((row as u32, col as u32), value.clone());
        }
        for (row, col, text) in formulas.used_cells() {
            if !text.is_empty() {
                sheet.formulas.insert((row as u32, col as u32), text.clone());
            }
        }
        out.names.push(name.clone());
        out.sheets.insert(name, sheet);
    }
    Ok(out)
}

fn read_rxls(bytes: &[u8], with_styles: bool) -> Result<RxlsBookInfo, String> {
    let book = rxls::Workbook::open(bytes).map_err(|e| format!("rxls open failed: {e}"))?;
    let mut out = RxlsBookInfo {
        defined_names: book
            .defined_names()
            .iter()
            .map(|(name, refers)| format!("{name}={refers}"))
            .collect(),
        local_defined_names: book
            .local_defined_names()
            .iter()
            .map(|n| format!("{}@{}={}", n.name, n.sheet, n.refers_to))
            .collect(),
        sheets: Vec::new(),
        style_cells: BTreeMap::new(),
    };
    for sheet in &book.sheets {
        let default_width = sheet.default_column_width();
        let counts = format!(
            "comments={} dv={} cf={} links={} img={} chart={} table={}",
            sheet.comments().len(),
            sheet.data_validations().len(),
            sheet.conditional_formats().len(),
            sheet.hyperlinks().len(),
            sheet.images().len(),
            sheet.charts().len(),
            sheet.tables().len(),
        );
        out.sheets.push(RxlsSheetInfo {
            name: sheet.name.clone(),
            is_worksheet: sheet.is_worksheet,
            visible: format!("{:?}", sheet.visible()),
            merges: sheet
                .merged_ranges()
                .iter()
                .map(|(r0, c0, r1, c1)| (*r0, *c0, *r1, *c1))
                .collect(),
            col_widths: sheet
                .column_widths()
                .iter()
                .filter(|(_, width)| {
                    // The XLSX writer may emit an explicit default-width <col>
                    // for every column; skip those so they don't read as extra.
                    match default_width {
                        Some(default) => (*width - default).abs() > 0.01,
                        None => true,
                    }
                })
                .map(|(col, width)| (*col, *width))
                .collect(),
            row_heights: sheet.row_heights().clone(),
            hidden_rows: sheet.hidden_rows().iter().copied().collect(),
            hidden_cols: sheet.hidden_columns().iter().copied().collect(),
            freeze: sheet.sheet_view().freeze,
            page_setup: sheet.page_setup().map(|p| format!("{p:?}")),
            // XLSX omits <printOptions> for "false"; treat None as false so a
            // BIFF8 explicit false does not read as a mismatch.
            print_gridlines: sheet.print_metadata().print_gridlines().unwrap_or(false),
            print_headings: sheet.print_metadata().print_headings().unwrap_or(false),
            counts,
            style_fidelity: format!("{:?}", sheet.style_fidelity()),
        });
        if with_styles {
            let mut coords: std::collections::BTreeSet<(u32, u16)> = sheet
                .blank_cell_styles()
                .keys()
                .copied()
                .collect();
            for (row, col, _cell) in sheet.cells() {
                coords.insert((row, col));
            }
            let mut cells = BTreeMap::new();
            for (row, col) in coords {
                let style = sheet.resolved_cell_style(row, col);
                cells.insert(
                    (row, col),
                    StyleKey::from(style.as_ref().unwrap_or(&rxls::CellStyle::default())),
                );
            }
            out.style_cells.insert(sheet.name.clone(), cells);
        }
    }
    Ok(out)
}

fn compare_sheets(
    src: &CalamineBook,
    conv: &CalamineBook,
    cap: usize,
    cat: &mut Cat,
    chartsheet_names: &[String],
) {
    let src_set: std::collections::BTreeSet<&String> = src.names.iter().collect();
    let conv_set: std::collections::BTreeSet<&String> = conv.names.iter().collect();
    for name in &src.names {
        if chartsheet_names.contains(name) {
            continue;
        }
        if !conv_set.contains(name) {
            cat.note_missing(cap, format!("sheet '{name}' missing"));
        }
    }
    for name in &conv.names {
        if chartsheet_names.contains(name) {
            continue;
        }
        if !src_set.contains(name) {
            cat.note_extra(cap, format!("sheet '{name}' extra"));
        }
    }
    let grid_src: Vec<&String> = src
        .names
        .iter()
        .filter(|n| !chartsheet_names.contains(n))
        .collect();
    let grid_conv: Vec<&String> = conv
        .names
        .iter()
        .filter(|n| !chartsheet_names.contains(n))
        .collect();
    if grid_src == grid_conv {
        cat.note_ok();
    } else {
        cat.note_mismatch(
            cap,
            format!(
                "sheet names/order differ: source {:?} vs converted {:?}",
                grid_src, grid_conv
            ),
        );
    }
}

fn compare_dims(src: &SheetCal, conv: &SheetCal, cap: usize, cat: &mut Cat) {
    let render = |start: &Option<CalKey>, end: &Option<CalKey>| {
        format!(
            "{}:{}-{}:{}",
            start.map(|s| s.0).unwrap_or(0),
            start.map(|s| s.1).unwrap_or(0),
            end.map(|e| e.0).unwrap_or(0),
            end.map(|e| e.1).unwrap_or(0)
        )
    };
    let src_str = render(&src.start, &src.end);
    let conv_str = render(&conv.start, &conv.end);
    if src_str == conv_str {
        cat.note_ok();
    } else {
        cat.note_mismatch(
            cap,
            format!("used range differs: source {src_str} vs converted {conv_str}"),
        );
    }
}

fn type_tag(data: &Data) -> &'static str {
    match data {
        Data::Empty => "empty",
        Data::Int(_) | Data::Float(_) => "number",
        Data::String(_) => "text",
        Data::Bool(_) => "bool",
        Data::DateTime(_) => "datetime",
        Data::DateTimeIso(_) => "datetime-iso",
        Data::DurationIso(_) => "duration-iso",
        Data::Error(_) => "error",
    }
}

fn num_close(a: f64, b: f64) -> bool {
    (a - b).abs() <= 1e-9 * a.abs().max(1.0)
}

fn values_equal(a: &Data, b: &Data) -> bool {
    match (a, b) {
        (Data::Empty, Data::Empty) => true,
        (Data::Int(x), Data::Int(y)) => x == y,
        (Data::Int(x), Data::Float(y)) => num_close(*x as f64, *y),
        (Data::Float(x), Data::Int(y)) => num_close(*x, *y as f64),
        (Data::Float(x), Data::Float(y)) => num_close(*x, *y),
        (Data::String(x), Data::String(y)) => x == y,
        (Data::Bool(x), Data::Bool(y)) => x == y,
        (Data::DateTime(x), Data::DateTime(y)) => num_close(x.as_f64(), y.as_f64()),
        (Data::DateTime(x), Data::Float(y)) => num_close(x.as_f64(), *y),
        (Data::Float(x), Data::DateTime(y)) => num_close(*x, y.as_f64()),
        (Data::DateTimeIso(x), Data::DateTimeIso(y)) => x == y,
        (Data::DurationIso(x), Data::DurationIso(y)) => x == y,
        (Data::Error(x), Data::Error(y)) => x.to_string() == y.to_string(),
        // Error value re-materialized as its display text: value survived,
        // storage type did not.
        (Data::Error(x), Data::String(y)) => x.to_string() == *y,
        (Data::String(x), Data::Error(y)) => *x == y.to_string(),
        _ => false,
    }
}

fn compare_data(src: &SheetCal, conv: &SheetCal, cap: usize, cat: &mut Cat) {
    let mut keys: Vec<CalKey> = src.values.keys().copied().collect();
    keys.extend(conv.values.keys().copied());
    keys.sort();
    keys.dedup();
    for key in keys {
        let a = src.values.get(&key);
        let b = conv.values.get(&key);
        match (a, b) {
            (Some(a), Some(b)) => {
                let same_type = type_tag(a) == type_tag(b);
                let equal = values_equal(a, b);
                if equal && same_type {
                    cat.note_ok();
                } else if equal && !same_type {
                    cat.note_type_only(
                        cap,
                        format!(
                            "cell {key:?} value ok, type tag differs: source {} ({a:?}) vs converted {} ({b:?})",
                            type_tag(a),
                            type_tag(b)
                        ),
                    );
                } else {
                    cat.note_mismatch(
                        cap,
                        format!("cell {key:?} value differs: source {a:?} vs converted {b:?}"),
                    );
                }
            }
            (Some(a), None) => {
                if !matches!(a, Data::Empty) {
                    cat.note_missing(cap, format!("cell {key:?} missing ({a:?})"));
                }
            }
            (None, Some(b)) => {
                if !matches!(b, Data::Empty) {
                    cat.note_extra(cap, format!("cell {key:?} extra ({b:?})"));
                }
            }
            (None, None) => {}
        }
    }
}

fn normalize_formula(text: &str) -> String {
    text.trim()
        .replace(' ', "")
        .replace('=', "")
        .to_uppercase()
}

fn compare_formulas(src: &SheetCal, conv: &SheetCal, cap: usize, cat: &mut Cat) {
    let mut keys: Vec<CalKey> = src.formulas.keys().copied().collect();
    keys.extend(conv.formulas.keys().copied());
    keys.sort();
    keys.dedup();
    for key in keys {
        let a = src.formulas.get(&key);
        let b = conv.formulas.get(&key);
        match (a, b) {
            (Some(a), Some(b)) => {
                if normalize_formula(a) == normalize_formula(b) {
                    cat.note_ok();
                } else {
                    cat.note_mismatch(
                        cap,
                        format!("formula {key:?} differs: source '{a}' vs converted '{b}'"),
                    );
                }
            }
            (Some(a), None) => {
                cat.note_missing(cap, format!("formula {key:?} missing: source '{a}'"));
            }
            (None, Some(b)) => {
                cat.note_extra(cap, format!("formula {key:?} extra: '{b}'"));
            }
            (None, None) => {}
        }
    }
}

fn compare_structure(src: &RxlsBookInfo, conv: &RxlsBookInfo, cap: usize, cat: &mut Cat) {
    if src.defined_names.len() == conv.defined_names.len() {
        cat.note_ok();
    } else {
        cat.note_mismatch(
            cap,
            format!(
                "defined names count differs: source {} vs converted {}",
                src.defined_names.len(),
                conv.defined_names.len()
            ),
        );
    }
    if src.local_defined_names.len() == conv.local_defined_names.len() {
        cat.note_ok();
    } else {
        cat.note_mismatch(
            cap,
            format!(
                "sheet-scoped defined names count differs: source {} vs converted {}",
                src.local_defined_names.len(),
                conv.local_defined_names.len()
            ),
        );
    }
    for (src_sheet, conv_sheet) in src.sheets.iter().zip(conv.sheets.iter()) {
        let name = &src_sheet.name;
        let label = |what: &str| format!("{name}/{what}");

        if src_sheet.visible == conv_sheet.visible {
            cat.note_ok();
        } else {            cat.note_mismatch(
                cap,
                format!(
                    "{} visibility differs: source {} vs converted {}",
                    label("sheet"),
                    src_sheet.visible,
                    conv_sheet.visible
                ),
            );
        }

        if src_sheet.merges.len() == conv_sheet.merges.len() {
            cat.note_ok();
        } else {
            cat.note_mismatch(
                cap,
                format!(
                    "{} merge count differs: source {} vs converted {}",
                    label("merges"),
                    src_sheet.merges.len(),
                    conv_sheet.merges.len()
                ),
            );
        }
        let mut src_merges = src_sheet.merges.clone();
        let mut conv_merges = conv_sheet.merges.clone();
        src_merges.sort();
        conv_merges.sort();
        if src_merges == conv_merges {
            cat.note_ok();
        } else {
            cat.note_mismatch(
                cap,
                format!(
                    "{} merge ranges differ: source {:?} vs converted {:?}",
                    label("merges"),
                    src_merges,
                    conv_merges
                ),
            );
        }

        if src_sheet.hidden_rows == conv_sheet.hidden_rows {
            cat.note_ok();
        } else {
            cat.note_mismatch(
                cap,
                format!(
                    "{} differ: source {:?} vs converted {:?}",
                    label("hidden-rows"),
                    src_sheet.hidden_rows,
                    conv_sheet.hidden_rows
                ),
            );
        }
        if src_sheet.hidden_cols == conv_sheet.hidden_cols {
            cat.note_ok();
        } else {
            cat.note_mismatch(
                cap,
                format!(
                    "{} differ: source {:?} vs converted {:?}",
                    label("hidden-cols"),
                    src_sheet.hidden_cols,
                    conv_sheet.hidden_cols
                ),
            );
        }

        if src_sheet.freeze == conv_sheet.freeze {
            cat.note_ok();
        } else {
            cat.note_mismatch(
                cap,
                format!(
                    "{} differ: source {:?} vs converted {:?}",
                    label("freeze"),
                    src_sheet.freeze,
                    conv_sheet.freeze
                ),
            );
        }

        if src_sheet.counts == conv_sheet.counts {
            cat.note_ok();
        } else {
            cat.note_mismatch(
                cap,
                format!(
                    "{} feature counts differ: source [{}] vs converted [{}]",
                    label("counts"),
                    src_sheet.counts,
                    conv_sheet.counts
                ),
            );
        }
        // Chartsheets carry no meaningful grid metrics; the XLSX writer emits
        // a full-width default column strip for them, which would otherwise
        // flood the comparison with false extras.
        if src_sheet.is_worksheet && conv_sheet.is_worksheet {
            map_eq_tol(
                &src_sheet.row_heights,
                &conv_sheet.row_heights,
                0.51,
                cap,
                cat,
                &label("row-height"),
                |v| *v as f64,
            );
            map_eq_tol(
                &width_keyed(&src_sheet.col_widths),
                &width_keyed(&conv_sheet.col_widths),
                0.26,
                cap,
                cat,
                &label("col-width"),
                |v| *v as f64,
            );
        }
    }
}

fn map_eq_tol<T: Copy + PartialEq>(
    src: &BTreeMap<u32, T>,
    conv: &BTreeMap<u32, T>,
    tol: f64,
    cap: usize,
    cat: &mut Cat,
    label: &str,
    get: impl Fn(&T) -> f64,
) {
    let mut keys: Vec<u32> = src.keys().copied().collect();
    keys.extend(conv.keys().copied());
    keys.sort();
    keys.dedup();
    for key in keys {
        match (src.get(&key), conv.get(&key)) {
            (Some(a), Some(b)) => {
                if (get(a) - get(b)).abs() <= tol {
                    cat.note_ok();
                } else {
                    cat.note_mismatch(
                        cap,
                        format!(
                            "{label} {key} differs: source {} vs converted {}",
                            get(a),
                            get(b)
                        ),
                    );
                }
            }
            (Some(a), None) => {
                cat.note_missing(cap, format!("{label} {key} missing (source {})", get(a)))
            }
            (None, Some(b)) => {
                cat.note_extra(cap, format!("{label} {key} extra ({})", get(b)))
            }
            (None, None) => {}
        }
    }
}

fn width_keyed(map: &BTreeMap<u16, f32>) -> BTreeMap<u32, f32> {
    map.iter().map(|(k, v)| (*k as u32, *v)).collect()
}

fn compare_print(src: &RxlsBookInfo, conv: &RxlsBookInfo, cap: usize, cat: &mut Cat) {
    for (src_sheet, conv_sheet) in src.sheets.iter().zip(conv.sheets.iter()) {
        let name = &src_sheet.name;
        match (&src_sheet.page_setup, &conv_sheet.page_setup) {
            (Some(a), Some(b)) => {
                if a == b {
                    cat.note_ok();
                } else {
                    cat.note_mismatch(
                        cap,
                        format!("{name}/page-setup differs:\n  source {a}\n  converted {b}"),
                    );
                }
            }
            (Some(_), None) => cat.note_missing(cap, format!("{name}/page-setup missing")),
            (None, Some(_)) => cat.note_extra(cap, format!("{name}/page-setup extra")),
            (None, None) => cat.note_ok(),
        }
        if src_sheet.print_gridlines == conv_sheet.print_gridlines {
            cat.note_ok();
        } else {
            cat.note_mismatch(
                cap,
                format!(
                    "{name}/print-gridlines differs: {:?} vs {:?}",
                    src_sheet.print_gridlines, conv_sheet.print_gridlines
                ),
            );
        }
        if src_sheet.print_headings == conv_sheet.print_headings {
            cat.note_ok();
        } else {
            cat.note_mismatch(
                cap,
                format!(
                    "{name}/print-headings differs: {:?} vs {:?}",
                    src_sheet.print_headings, conv_sheet.print_headings
                ),
            );
        }
    }
}

fn compare_styles(src: &RxlsBookInfo, conv: &RxlsBookInfo, cap: usize, cat: &mut Cat) {
    for (src_sheet, conv_sheet) in src.sheets.iter().zip(conv.sheets.iter()) {
        if src_sheet.style_fidelity != "Retained" && src_sheet.style_fidelity != "Authored" {
            cat.samples.push(format!(
                "{}: source style fidelity {} (reader may under-report styles)",
                src_sheet.name, src_sheet.style_fidelity
            ));
        }
        if conv_sheet.style_fidelity != "Retained" && conv_sheet.style_fidelity != "Authored" {
            cat.samples.push(format!(
                "{}: converted style fidelity {} (reader may under-report styles)",
                conv_sheet.name, conv_sheet.style_fidelity
            ));
        }
        let (Some(src_cells), Some(conv_cells)) = (
            src.style_cells.get(&src_sheet.name),
            conv.style_cells.get(&conv_sheet.name),
        ) else {
            continue;
        };
        let mut keys: Vec<RxKey> = src_cells.keys().copied().collect();
        keys.extend(conv_cells.keys().copied());
        keys.sort();
        keys.dedup();
        for key in keys {
            match (src_cells.get(&key), conv_cells.get(&key)) {
                (Some(a), Some(b)) => {
                    if a == b {
                        cat.note_ok();
                    } else {
                        cat.note_mismatch(
                            cap,
                            format!(
                                "style {key:?} differs:\n  source {}\n  converted {}",
                                a.render(),
                                b.render()
                            ),
                        );
                    }
                }
                (Some(a), None) => {
                    cat.note_missing(cap, format!("style {key:?} missing: {}", a.render()))
                }
                (None, Some(b)) => {
                    cat.note_extra(cap, format!("style {key:?} extra: {}", b.render()))
                }
                (None, None) => {}
            }
        }
    }
}
