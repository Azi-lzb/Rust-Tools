//! Backend A: Excel COM conversion via native IDispatch (Windows only).
//!
//! Two lifecycle modes mirror the production baseline:
//! - batch: one STA thread + one owned Excel instance convert the whole list
//! - per-file: every input gets a fresh STA thread + Excel instance
//!
//! Excel is configured exactly like the repository's existing isolated
//! benchmark (visible=false, alerts/events/screen-updates off, manual
//! calculation configured through a transient blank workbook, open read-only,
//! SaveAs FileFormat=51, Close(false), PID-verified Quit).

use crate::common::{elapsed_ms, ComMeta, FileOutcome};
use crate::combase::{
    pump_messages, sta_call, Apartment, ComApp, ComWorkbook,
};
use anyhow::{Context as _, Result};
use std::path::{Path, PathBuf};
use std::time::Instant;

/// Normalize a path for COM (absolute, no verbatim `\\?\` prefix).
fn com_path(path: &Path) -> String {
    let path = if path.is_absolute() {
        path.to_path_buf()
    } else {
        std::env::current_dir()
            .unwrap_or_else(|_| PathBuf::from("."))
            .join(path)
    };
    // Skip std canonicalize: it returns `\\?\`-prefixed paths on Windows,
    // which some Excel SaveAs paths reject. Components are already real.
    path.to_string_lossy().replace('/', "\\")
}

pub fn convert(
    pairs: &[(PathBuf, PathBuf)],
    per_file: bool,
) -> Result<(Vec<FileOutcome>, ComMeta)> {
    for (_, output) in pairs {
        if let Some(parent) = output.parent() {
            std::fs::create_dir_all(parent)?;
        }
        if output.exists() {
            anyhow::bail!("refusing to overwrite existing output: {}", output.display());
        }
    }
    if per_file {
        convert_per_file(pairs)
    } else {
        convert_batch(pairs)
    }
}

fn convert_batch(pairs: &[(PathBuf, PathBuf)]) -> Result<(Vec<FileOutcome>, ComMeta)> {
    let owned: Vec<(PathBuf, PathBuf)> = pairs.to_vec();
    let (mut outcomes, mut meta) = sta_call(move || -> Result<(Vec<FileOutcome>, ComMeta)> {
        let mut outcomes = Vec::with_capacity(owned.len());
        let mut meta = ComMeta {
            excel_instances: 1,
            ..Default::default()
        };
        let t_sta = Instant::now();
        let apartment = Apartment::initialize()?;
        meta.sta_init_ms = elapsed_ms(t_sta);

        let t_create = Instant::now();
        let mut app = ComApp::create_and_configure("Excel.Application", true)
            .context("create/configure native Excel COM instance")?;
        meta.excel_create_config_ms = elapsed_ms(t_create);
        meta.excel_pid = Some(app.process_id);

        for (input, output) in &owned {
            outcomes.push(convert_one_with(&app, input, output, false));
            pump_messages();
        }

        let (peak, cpu) = app.memory()?;
        meta.excel_peak_ws_bytes = Some(peak);
        meta.excel_cpu_ms = Some(cpu);

        let t_quit = Instant::now();
        let confirmed = app.shutdown()?;
        meta.excel_quit_ms = elapsed_ms(t_quit);
        meta.excel_exit_confirmed = confirmed;
        drop(app);

        let t_release = Instant::now();
        drop(apartment);
        meta.sta_release_ms = elapsed_ms(t_release);
        Ok((outcomes, meta))
    })?;
    if !meta.excel_exit_confirmed {
        if let Some(pid) = meta.excel_pid {
            anyhow::bail!("owned Excel PID {pid} did not exit cleanly");
        }
    }
    Ok((outcomes, meta))
}

fn convert_per_file(pairs: &[(PathBuf, PathBuf)]) -> Result<(Vec<FileOutcome>, ComMeta)> {
    let mut outcomes = Vec::with_capacity(pairs.len());
    let mut meta = ComMeta::default();
    for (input, output) in pairs {
        let input = input.clone();
        let output = output.clone();
        let (outcome, file_meta) = sta_call(move || -> Result<(FileOutcome, ComMeta)> {
            let mut meta = ComMeta {
                excel_instances: 1,
                ..Default::default()
            };
            let t_all = Instant::now();
            let t_sta = Instant::now();
            let apartment = Apartment::initialize()?;
            meta.sta_init_ms = elapsed_ms(t_sta);

            let t_create = Instant::now();
            let mut app = ComApp::create_and_configure("Excel.Application", true)
                .context("create/configure native Excel COM instance")?;
            meta.excel_create_config_ms = elapsed_ms(t_create);
            meta.excel_pid = Some(app.process_id);

            let outcome = convert_one_with(&app, &input, &output, true);

            let (peak, cpu) = app.memory()?;
            meta.excel_peak_ws_bytes = Some(peak);
            meta.excel_cpu_ms = Some(cpu);

            let t_quit = Instant::now();
            let confirmed = app.shutdown()?;
            meta.excel_quit_ms = elapsed_ms(t_quit);
            meta.excel_exit_confirmed = confirmed;
            drop(app);

            let t_release = Instant::now();
            drop(apartment);
            meta.sta_release_ms = elapsed_ms(t_release);

            let mut outcome = outcome;
            outcome.timing.total_ms = elapsed_ms(t_all);
            Ok((outcome, meta))
        })?;
        meta.excel_instances += 1;
        meta.sta_init_ms += file_meta.sta_init_ms;
        meta.excel_create_config_ms += file_meta.excel_create_config_ms;
        meta.excel_quit_ms += file_meta.excel_quit_ms;
        meta.sta_release_ms += file_meta.sta_release_ms;
        meta.excel_exit_confirmed &= file_meta.excel_exit_confirmed;
        meta.excel_pid = file_meta.excel_pid;
        if let (Some(acc), Some(next)) = (meta.excel_peak_ws_bytes, file_meta.excel_peak_ws_bytes) {
            meta.excel_peak_ws_bytes = Some(acc.max(next));
        } else {
            meta.excel_peak_ws_bytes = file_meta.excel_peak_ws_bytes;
        }
        if let (Some(acc), Some(next)) = (meta.excel_cpu_ms, file_meta.excel_cpu_ms) {
            meta.excel_cpu_ms = Some(acc + next);
        } else {
            meta.excel_cpu_ms = file_meta.excel_cpu_ms;
        }
        outcomes.push(outcome);
        if !file_meta.excel_exit_confirmed {
            anyhow::bail!("owned Excel instance did not exit cleanly");
        }
    }
    Ok((outcomes, meta))
}

/// Convert one workbook through an existing `ComApp`.
///
/// `full_lifecycle` (per-file mode) leaves `timing.total_ms` to the caller,
/// which measures the whole instance lifecycle.
fn convert_one_with(
    app: &ComApp,
    input: &Path,
    output: &Path,
    _full_lifecycle: bool,
) -> FileOutcome {
    let mut outcome = FileOutcome::new(input, output);
    if let Ok(source) = std::fs::read(input) {
        outcome.hash_input(&source);
    }
    let t_open = Instant::now();
    let workbook = match app.open_workbook(&com_path(input)) {
        Ok(workbook) => workbook,
        Err(e) => {
            outcome.error = Some(format!("{e:#}"));
            outcome.timing.total_ms = elapsed_ms(t_open);
            return outcome;
        }
    };
    outcome.timing.com_open_ms = Some(elapsed_ms(t_open));
    let mut workbook = ComWorkbook::from_dispatch(workbook);

    let t_save = Instant::now();
    let save_result = workbook.save_as_xlsx(&com_path(output));
    outcome.timing.com_save_as_ms = Some(elapsed_ms(t_save));

    let t_close = Instant::now();
    let close_result = workbook.close();
    outcome.timing.com_close_ms = Some(elapsed_ms(t_close));
    outcome.timing.total_ms = elapsed_ms(t_open);

    match (save_result, close_result) {
        (Ok(()), Ok(())) => {
            if let Ok(output_bytes) = std::fs::read(output) {
                outcome.hash_output(&output_bytes);
            }
            outcome.status = "ok".to_owned();
        }
        (Err(save_error), close) => {
            outcome.error = Some(format!("{save_error:#}"));
            if let Err(close_error) = close {
                outcome.error = Some(format!(
                    "{save_error:#}; close also failed: {close_error:#}"
                ));
            }
        }
        (Ok(()), Err(close_error)) => {
            outcome.error = Some(format!("{close_error:#}"));
        }
    }
    outcome
}
