//! Open converted XLSX files in real Excel via COM and probe Office-side
//! compatibility: open success, sheet count, print page counts, PDF export,
//! and a SaveAs round-trip diff (Excel silently repairing content would show
//! up as value/formula differences between the round-trip and the original).

use crate::common::now_epoch_ms;
use crate::combase::{pump_messages, sta_call, Apartment, ComApp, ComWorkbook};
use anyhow::{Context, Result};
use serde::Serialize;
use std::path::{Path, PathBuf};
use std::time::Instant;

#[derive(Debug, Default, Serialize)]
pub struct OpenCheckEntry {
    pub file: String,
    /// `ok` | `failed`
    pub status: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error: Option<String>,
    pub open_ms: f64,
    pub sheets_count_com: Option<usize>,
    pub print_pages_com: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pdf_path: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pdf_bytes: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub roundtrip_path: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub roundtrip_value_diffs: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub roundtrip_formula_diffs: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub roundtrip_note: Option<String>,
}

#[derive(Debug, Serialize)]
pub struct OpenCheckReport {
    pub started_at_epoch_ms: u64,
    pub excel_prog_id: String,
    pub wps_available: Option<bool>,
    /// Excel exposes no direct "repair prompt" API; this run used open +
    /// round-trip diffing as the compatibility signal.
    pub repair_prompt_detection: String,
    pub files: Vec<OpenCheckEntry>,
}

pub fn run(
    input: &Path,
    out_dir: &Path,
    with_pdf: bool,
    with_roundtrip: bool,
    wps_files: usize,
) -> Result<()> {
    let mut files = Vec::new();
    for entry in std::fs::read_dir(input)? {
        let path = entry?.path();
        let is_xlsx = path
            .extension()
            .map(|e| e.eq_ignore_ascii_case("xlsx"))
            .unwrap_or(false);
        if path.is_file() && is_xlsx {
            files.push(path);
        }
    }
    files.sort();
    anyhow::ensure!(
        !files.is_empty(),
        "no .xlsx files found under {}",
        input.display()
    );
    let pdf_dir = out_dir.join("pdf");
    let roundtrip_dir = out_dir.join("roundtrip");
    if with_pdf {
        std::fs::create_dir_all(&pdf_dir)?;
    }
    if with_roundtrip {
        std::fs::create_dir_all(&roundtrip_dir)?;
    }

    let owned: Vec<PathBuf> = files.clone();
    let pdf_dir_owned = pdf_dir.clone();
    let roundtrip_dir_owned = roundtrip_dir.clone();

    let entries = sta_call(move || -> Result<Vec<OpenCheckEntry>> {
        let mut entries = Vec::with_capacity(owned.len());
        let apartment = Apartment::initialize()?;
        let mut app = ComApp::create_and_configure("Excel.Application", false)
            .context("create Excel instance for open checks")?;
        for (index, file) in owned.iter().enumerate() {
            let mut entry = OpenCheckEntry {
                file: file.to_string_lossy().into_owned(),
                ..Default::default()
            };
            let t_open = Instant::now();
            match app.open_workbook(&com_path(file)) {
                Ok(dispatch) => {
                    entry.open_ms = crate::common::elapsed_ms(t_open);
                    let mut workbook = ComWorkbook::from_dispatch(dispatch);
                    entry.sheets_count_com = workbook.sheet_count().ok();
                    entry.print_pages_com = workbook.total_print_pages().ok();
                    if with_pdf {
                        let pdf_path = pdf_dir_owned
                            .join(file.file_stem().unwrap_or_default().to_string_lossy().as_ref())
                            .with_extension("pdf");
                        match workbook.export_pdf(&com_path(&pdf_path)) {
                            Ok(()) => {
                                entry.pdf_path = Some(pdf_path.to_string_lossy().into_owned());
                                entry.pdf_bytes =
                                    std::fs::metadata(&pdf_path).map(|m| m.len() as u64).ok();
                            }
                            Err(e) => {
                                entry.error =
                                    Some(format!("pdf export failed: {e:#}"));
                            }
                        }
                    }
                    if with_roundtrip {
                        let roundtrip_path = roundtrip_dir_owned.join(format!(
                            "{}.resave.xlsx",
                            file.file_stem().unwrap_or_default().to_string_lossy()
                        ));
                        match workbook.save_as_xlsx(&com_path(&roundtrip_path)) {
                            Ok(()) => {
                                entry.roundtrip_path =
                                    Some(roundtrip_path.to_string_lossy().into_owned());
                                let (values, formulas, note) =
                                    roundtrip_diff(file, &roundtrip_path);
                                entry.roundtrip_value_diffs = Some(values);
                                entry.roundtrip_formula_diffs = Some(formulas);
                                entry.roundtrip_note = note;
                            }
                            Err(e) => {
                                entry.error = Some(format!(
                                    "{} roundtrip save failed: {e:#}",
                                    entry.error.clone().unwrap_or_default()
                                ));
                            }
                        }
                    }
                    match workbook.close() {
                        Ok(()) => {
                            if entry.error.is_some() {
                                entry.status = "failed".to_owned();
                            } else {
                                entry.status = "ok".to_owned();
                            }
                        }
                        Err(e) => {
                            entry.error = Some(format!(
                                "{} close failed: {e:#}",
                                entry.error.clone().unwrap_or_default()
                            ));
                            entry.status = "failed".to_owned();
                        }
                    }
                }
                Err(e) => {
                    entry.open_ms = crate::common::elapsed_ms(t_open);
                    entry.status = "failed".to_owned();
                    entry.error = Some(format!("Excel open failed: {e:#}"));
                }
            }
            entries.push(entry);
            pump_messages();
            let _ = index;
        }
        app.shutdown()?;
        drop(app);
        drop(apartment);
        Ok(entries)
    })?;

    let wps_available = if wps_files > 0 {
        Some(wps_probe(&files, wps_files))
    } else {
        None
    };

    let report = OpenCheckReport {
        started_at_epoch_ms: now_epoch_ms(),
        excel_prog_id: "Excel.Application".to_owned(),
        wps_available,
        repair_prompt_detection: "none available in Excel COM API; open success + SaveAs round-trip value/formula diff used as the compatibility signal".to_owned(),
        files: entries,
    };
    let report_path = out_dir.join("opencheck-report.json");
    std::fs::create_dir_all(out_dir)?;
    std::fs::write(&report_path, serde_json::to_string_pretty(&report)?)?;
    eprintln!("opencheck report written: {}", report_path.display());
    Ok(())
}

/// Try opening the first `limit` files with WPS ET (Ket.Application). Returns
/// whether the engine was creatable/openable; per-file WPS verdicts are a
/// separate manual exercise.
fn wps_probe(files: &[PathBuf], limit: usize) -> bool {
    let owned: Vec<PathBuf> = files.to_vec();
    let probe = sta_call(move || -> Result<bool> {
        let apartment = Apartment::initialize()?;
        let mut app = match ComApp::create_and_configure("Ket.Application", false) {
            Ok(app) => app,
            Err(_) => {
                drop(apartment);
                return Ok(false);
            }
        };
        let mut ok = false;
        for file in owned.iter().take(limit) {
            if app.open_workbook(&com_path(file)).is_ok() {
                ok = true;
            }
        }
        let _ = app.shutdown();
        drop(app);
        drop(apartment);
        Ok(ok)
    });
    probe.unwrap_or(false)
}

fn roundtrip_diff(original: &Path, resave: &Path) -> (usize, usize, Option<String>) {
    let orig_bytes = match std::fs::read(original) {
        Ok(b) => b,
        Err(e) => return (0, 0, Some(format!("original unreadable: {e}"))),
    };
    let resave_bytes = match std::fs::read(resave) {
        Ok(b) => b,
        Err(e) => return (0, 0, Some(format!("resave unreadable: {e}"))),
    };
    let orig = match rxls::Workbook::open(&orig_bytes) {
        Ok(wb) => wb,
        Err(e) => return (0, 0, Some(format!("original re-parse failed: {e}"))),
    };
    let resave = match rxls::Workbook::open(&resave_bytes) {
        Ok(wb) => wb,
        Err(e) => return (0, 0, Some(format!("resave re-parse failed: {e}"))),
    };
    let mut value_diffs = 0usize;
    let mut formula_diffs = 0usize;
    for (src_sheet, dst_sheet) in orig.sheets.iter().zip(resave.sheets.iter()) {
        let mut src_cells: std::collections::BTreeMap<(u32, u16), &rxls::Cell> =
            std::collections::BTreeMap::new();
        for (row, col, cell) in src_sheet.cells() {
            src_cells.insert((row, col), cell);
        }
        let mut dst_cells: std::collections::BTreeMap<(u32, u16), &rxls::Cell> =
            std::collections::BTreeMap::new();
        for (row, col, cell) in dst_sheet.cells() {
            dst_cells.insert((row, col), cell);
        }
        let keys: std::collections::BTreeSet<(u32, u16)> = src_cells
            .keys()
            .copied()
            .chain(dst_cells.keys().copied())
            .collect();
        for key in keys {
            let a = src_cells.get(&key);
            let b = dst_cells.get(&key);
            let formula_diff = match (a, b) {
                (
                    Some(rxls::Cell::Formula { formula: fa, .. }),
                    Some(rxls::Cell::Formula { formula: fb, .. }),
                ) => fa.trim().to_uppercase() != fb.trim().to_uppercase(),
                (Some(rxls::Cell::Formula { .. }), _) | (_, Some(rxls::Cell::Formula { .. })) => true,
                _ => false,
            };
            let value_diff = match (a, b) {
                (Some(a), Some(b)) => !cells_value_equal(a, b),
                (Some(_), None) | (None, Some(_)) => true,
                (None, None) => false,
            };
            if value_diff {
                value_diffs += 1;
            }
            if formula_diff {
                formula_diffs += 1;
            }
        }
    }
    let note = if orig.sheets.len() != resave.sheets.len() {
        Some(format!(
            "sheet count changed: {} -> {}",
            orig.sheets.len(),
            resave.sheets.len()
        ))
    } else {
        None
    };
    (value_diffs, formula_diffs, note)
}

fn cells_value_equal(a: &rxls::Cell, b: &rxls::Cell) -> bool {
    use rxls::Cell;
    match (a, b) {
        (Cell::Formula { cached: a, .. }, Cell::Formula { cached: b, .. }) => {
            cells_value_equal(a, b)
        }
        (Cell::Formula { cached, .. }, other) | (other, Cell::Formula { cached, .. }) => {
            cells_value_equal(cached, other)
        }
        (Cell::Text(a), Cell::Text(b)) => a == b,
        (Cell::Number(a), Cell::Number(b)) => (a - b).abs() <= 1e-9 * a.abs().max(1.0),
        (Cell::Date(a), Cell::Date(b)) => (a - b).abs() <= 1e-9,
        (Cell::Bool(a), Cell::Bool(b)) => a == b,
        (Cell::Error(a), Cell::Error(b)) => a == b,
        _ => false,
    }
}

fn com_path(path: &Path) -> String {
    let path = if path.is_absolute() {
        path.to_path_buf()
    } else {
        std::env::current_dir()
            .unwrap_or_else(|_| PathBuf::from("."))
            .join(path)
    };
    path.to_string_lossy().replace('/', "\\")
}
