//! Backend C: calamine (XLS reader) + rust_xlsxwriter (XLSX writer).
//!
//! Two modes:
//! - values: cell values only (no styles, no formats) — pure data conversion.
//! - formulas: values + formula text with the reader-provided cached result.
//!
//! Known, documented limitations of this backend: cell styles, number formats,
//! merges, print setup and drawings are not migrated; dates are written with a
//! default writer date format (source number format is unavailable through
//! calamine); error cells are emitted as text.

use crate::common::{elapsed_ms, write_and_commit, FileOutcome, Validate};
use calamine::{Data, Reader, Xls};
use rust_xlsxwriter::{Format, Formula, Workbook, Worksheet};
use std::collections::BTreeMap;
use std::io::{BufReader, Read};
use std::path::{Path, PathBuf};
use std::time::Instant;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CalamineMode {
    Values,
    Formulas,
}

impl CalamineMode {
    pub fn id(self) -> &'static str {
        match self {
            CalamineMode::Values => "values",
            CalamineMode::Formulas => "values+formulas",
        }
    }
}

pub fn convert_one(
    input: &Path,
    out_dir: &Path,
    mode: CalamineMode,
    level: Validate,
) -> anyhow::Result<FileOutcome> {
    let stem = stem_of(input);
    let mut outcome = FileOutcome::new(input, &out_dir.join(format!("{stem}.xlsx")));

    let t0 = Instant::now();
    let file = match std::fs::File::open(input) {
        Ok(f) => f,
        Err(e) => {
            outcome.error = Some(format!("open failed: {e}"));
            outcome.timing.total_ms = elapsed_ms(t0);
            return Ok(outcome);
        }
    };
    let mut raw = Vec::new();
    let mut reader_handle = BufReader::new(file);
    let read_ok = reader_handle.read_to_end(&mut raw).is_ok();
    outcome.hash_input(&raw);
    outcome.timing.read_ms = Some(elapsed_ms(t0));

    let t1 = Instant::now();
    let mut book = match Xls::<BufReader<std::fs::File>>::new(BufReader::new(
        std::fs::File::open(input).map_err(|e| anyhow::anyhow!("open failed: {e}"))?,
    )) {
        Ok(b) => b,
        Err(e) => {
            outcome.error = Some(format!("calamine open failed: {e}"));
            outcome.timing.total_ms = elapsed_ms(t0);
            return Ok(outcome);
        }
    };
    let _ = read_ok;
    let names = book.sheet_names();
    let mut sheet_data: Vec<(String, calamine::Range<Data>)> = Vec::with_capacity(names.len());
    for name in &names {
        match book.worksheet_range(name) {
            Ok(range) => sheet_data.push((name.clone(), range)),
            Err(e) => {
                outcome.error = Some(format!("calamine range '{name}' failed: {e:?}"));
                outcome.timing.total_ms = elapsed_ms(t0);
                return Ok(outcome);
            }
        }
    }
    let mut sheet_formulas: BTreeMap<String, calamine::Range<String>> = BTreeMap::new();
    if mode == CalamineMode::Formulas {
        for name in &names {
            match book.worksheet_formula(name) {
                Ok(range) => {
                    sheet_formulas.insert(name.clone(), range);
                }
                Err(e) => {
                    outcome.error = Some(format!("calamine formula '{name}' failed: {e:?}"));
                    outcome.timing.total_ms = elapsed_ms(t0);
                    return Ok(outcome);
                }
            }
        }
    }
    let parse_ms = elapsed_ms(t1);
    outcome.timing.parse_ms = Some(parse_ms);

    let t2 = Instant::now();
    let mut output = Workbook::new();
    for (name, range) in &sheet_data {
        let sheet: &mut Worksheet = output.add_worksheet();
        if let Err(e) = sheet.set_name(name) {
            outcome.error = Some(format!("set sheet name '{name}' failed: {e}"));
            outcome.timing.total_ms = elapsed_ms(t0);
            return Ok(outcome);
        }
        let formulas = sheet_formulas.get(name);
        for (row, col, data) in range.used_cells() {
            let row = row as u32;
            let col = col as u16;
            let formula = formulas
                .and_then(|f| f.get_value((row, col as u32)))
                .map(|s| s.to_owned());
            if let Some(formula_text) = formula.filter(|f| !f.trim().is_empty()) {
                let mut formula = Formula::new(formula_text);
                if let Some(result) = cached_result_string(data) {
                    formula = formula.set_result(result);
                }
                if let Err(e) = sheet.write_formula(row, col, &formula) {
                    outcome.error = Some(format!("write formula at {row},{col} failed: {e}"));
                    outcome.timing.total_ms = elapsed_ms(t0);
                    return Ok(outcome);
                }
                continue;
            }
            if let Err(e) = write_data(sheet, row, col, data) {
                outcome.error = Some(format!("write cell at {row},{col} failed: {e}"));
                outcome.timing.total_ms = elapsed_ms(t0);
                return Ok(outcome);
            }
        }
    }
    let bytes = match output.save_to_buffer() {
        Ok(bytes) => bytes,
        Err(e) => {
            outcome.error = Some(format!("xlsx serialization failed: {e}"));
            outcome.timing.total_ms = elapsed_ms(t0);
            return Ok(outcome);
        }
    };
    let serialize_ms = elapsed_ms(t2);
    outcome.timing.serialize_ms = Some(serialize_ms);

    match write_and_commit(out_dir, &stem, &bytes, level) {
        Ok((final_path, write_ms, validate_ms)) => {
            outcome.timing.write_ms = Some(write_ms);
            outcome.timing.validate_ms = Some(validate_ms);
            outcome.output = final_path.to_string_lossy().into_owned();
            outcome.hash_output(&bytes);
            outcome.status = "ok".to_owned();
        }
        Err(reason) => {
            outcome.error = Some(reason);
        }
    }
    outcome.timing.total_ms = elapsed_ms(t0);
    Ok(outcome)
}

/// Render a calamine value as a formula's cached-result string.
fn cached_result_string(data: &Data) -> Option<String> {
    match data {
        Data::Empty => None,
        Data::Int(v) => Some(v.to_string()),
        Data::Float(v) => Some(format_float(*v)),
        Data::String(s) => Some(s.clone()),
        Data::Bool(b) => Some(if *b { "TRUE" } else { "FALSE" }.to_owned()),
        Data::DateTime(dt) => Some(format_float(dt.as_f64())),
        Data::DateTimeIso(s) | Data::DurationIso(s) => Some(s.clone()),
        Data::Error(e) => Some(e.to_string()),
    }
}

fn format_float(v: f64) -> String {
    if v == v.trunc() && v.abs() < 1e15 {
        format!("{}", v as i64)
    } else {
        format!("{v}")
    }
}

fn write_data(
    sheet: &mut Worksheet,
    row: u32,
    col: u16,
    data: &Data,
) -> Result<(), rust_xlsxwriter::XlsxError> {
    match data {
        Data::Empty => Ok(()),
        Data::Int(v) => sheet.write_number(row, col, *v as f64).map(|_| ()),
        Data::Float(v) => sheet.write_number(row, col, *v).map(|_| ()),
        Data::String(s) => sheet.write_string(row, col, s).map(|_| ()),
        Data::Bool(b) => sheet.write_boolean(row, col, *b).map(|_| ()),
        Data::DateTime(dt) => {
            let serial = dt.as_f64();
            match rust_xlsxwriter::ExcelDateTime::from_serial_datetime(serial) {
                Ok(parsed) => sheet.write_datetime(row, col, &parsed).map(|_| ()),
                Err(_) => {
                    // Outside the writer's supported serial range: keep the
                    // underlying number so no data is lost.
                    sheet.write_number(row, col, serial).map(|_| ())
                }
            }
        }
        Data::DateTimeIso(s) | Data::DurationIso(s) => {
            sheet.write_string(row, col, s).map(|_| ())
        }
        Data::Error(e) => {
            // rust_xlsxwriter has no static error-cell write; keep the display
            // text so the value survives (type loss is reported elsewhere).
            let text = e.to_string();
            let format = Format::new();
            sheet
                .write_string_with_format(row, col, &text, &format)
                .map(|_| ())
        }
    }
}

pub fn convert_batch(
    inputs: &[PathBuf],
    out_dir: &Path,
    mode: CalamineMode,
    level: Validate,
) -> anyhow::Result<Vec<FileOutcome>> {
    let mut outcomes = Vec::with_capacity(inputs.len());
    for input in inputs {
        outcomes.push(convert_one(input, out_dir, mode, level)?);
    }
    Ok(outcomes)
}

fn stem_of(path: &Path) -> String {
    path.file_stem()
        .map(|s| s.to_string_lossy().into_owned())
        .unwrap_or_else(|| "workbook".to_owned())
}
