//! Backend B: rXLS direct `.xls` → `.xlsx` conversion (`Workbook::open` +
//! `to_xlsx_checked`, with the documented sanitized `to_xlsx` fallback).

use crate::common::{elapsed_ms, write_and_commit, FileOutcome, Validate};
use serde_json::json;
use std::path::{Path, PathBuf};
use std::time::Instant;

/// Convert one XLS with rxls. `with_diag` adds the rxls workbook diagnostics
/// (counts, feature counters, style/print loss kinds) to the outcome.
pub fn convert_one(
    input: &Path,
    out_dir: &Path,
    level: Validate,
    with_diag: bool,
) -> anyhow::Result<FileOutcome> {
    let stem = stem_of(input);
    let mut outcome = FileOutcome::new(input, &out_dir.join(format!("{stem}.xlsx")));

    let t0 = Instant::now();
    let source = std::fs::read(input).map_err(|e| anyhow::anyhow!("read failed: {e}"))?;
    outcome.hash_input(&source);
    outcome.timing.read_ms = Some(elapsed_ms(t0));

    let t1 = Instant::now();
    let workbook = match rxls::Workbook::open(&source) {
        Ok(wb) => wb,
        Err(e) => {
            outcome.error = Some(format!("rxls open failed: {e}"));
            outcome.timing.total_ms = elapsed_ms(t0);
            return Ok(outcome);
        }
    };
    let parse_ms = elapsed_ms(t1);
    outcome.timing.parse_ms = Some(parse_ms);

    if with_diag {
        outcome.rxls_diag = Some(build_diag(&workbook));
    }

    let t2 = Instant::now();
    // Prefer the checked writer: it reports what the sanitized writer would
    // silently paper over. On a checked error we record it verbatim and fall
    // back to the documented best-effort path instead of failing outright —
    // the fallback is rxls' explicitly intended `.xls → .xlsx` conversion
    // path, and the committed file is still validated independently.
    let (bytes, checked_error) = match workbook.to_xlsx_checked() {
        Ok(bytes) => (bytes, None),
        Err(e) => {
            let message = format!("{e:?}");
            let bytes = workbook.to_xlsx();
            (bytes, Some(message))
        }
    };
    outcome.timing.serialize_ms = Some(elapsed_ms(t2));
    outcome.checked_error = checked_error;

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

/// Convert a whole directory of inputs sequentially (resident-process mode).
pub fn convert_batch(
    inputs: &[PathBuf],
    out_dir: &Path,
    level: Validate,
    with_diag: bool,
) -> anyhow::Result<Vec<FileOutcome>> {
    let mut outcomes = Vec::with_capacity(inputs.len());
    for input in inputs {
        outcomes.push(convert_one(input, out_dir, level, with_diag)?);
    }
    Ok(outcomes)
}

fn stem_of(path: &Path) -> String {
    path.file_stem()
        .map(|s| s.to_string_lossy().into_owned())
        .unwrap_or_else(|| "workbook".to_owned())
}

/// Compact rxls diagnostics: counts, feature totals, evaluation distribution,
/// and typed style/print loss kinds per sheet.
fn build_diag(workbook: &rxls::Workbook) -> serde_json::Value {
    let report = rxls::WorkbookReport::from_workbook("xls", workbook);
    let sheets: Vec<serde_json::Value> = workbook
        .sheets
        .iter()
        .map(|sheet| {
            let style_losses = sheet
                .style_losses()
                .iter()
                .map(|loss| json!({ "kind": format!("{:?}", loss.kind), "occurrences": loss.occurrences }))
                .collect::<Vec<_>>();
            let print_losses = sheet
                .print_metadata()
                .losses()
                .iter()
                .map(|loss| json!({ "kind": format!("{:?}", loss.kind), "occurrences": loss.occurrences }))
                .collect::<Vec<_>>();
            json!({
                "name": sheet.name,
                "style_fidelity": format!("{:?}", sheet.style_fidelity()),
                "style_losses": style_losses,
                "print_losses": print_losses,
            })
        })
        .collect();
    json!({
        "report_schema": report.schema_version,
        "cells": report.stats.cells,
        "formulas": report.stats.formulas,
        "defined_names": report.defined_names_count,
        "local_defined_names": report.local_defined_names_count,
        "features": {
            "comments": report.features.comments,
            "data_validations": report.features.data_validations,
            "tables": report.features.tables,
            "merged_ranges": report.features.merged_ranges,
            "hyperlinks": report.features.hyperlinks,
            "images": report.features.images,
            "charts": report.features.charts,
        },
        "evaluation": {
            "computed": report.evaluation.computed,
            "errors": report.evaluation.errors,
            "cached": report.evaluation.cached,
            "unsupported": report.evaluation.unsupported,
            "truncated": report.evaluation.truncated,
        },
        "warnings_count": report.warnings.len(),
        "sheets": sheets,
    })
}
