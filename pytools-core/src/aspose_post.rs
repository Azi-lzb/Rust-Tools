//! Read-only verification and append-only evidence for Aspose Worker outputs.
//! Object cleanup belongs to the isolated Worker. Rust never deletes candidate
//! worksheets based on names, emptiness, or input inventories. Legacy outcome
//! fields remain available for reading historical records.

use anyhow::{Context, Result};
use std::{
    fs,
    path::{Path, PathBuf},
};

pub const EVALUATION_SHEET_NAME: &str = "Evaluation Warning";

/// One related input workbook with independently read metadata.
#[derive(Clone, Debug)]
pub struct InputWorkbookTrace {
    pub path: PathBuf,
    pub sha256: Option<String>,
    /// Sheet names read via calamine; `None` means the list could not be
    /// reliably obtained (recorded as unknown, never as an empty list).
    pub sheets: Option<Vec<String>>,
}

/// Read input workbook metadata without trusting the worker's report.
pub fn trace_input_workbook(path: &Path) -> InputWorkbookTrace {
    let sha256 = crate::file_commit::sha256_file(path).ok();
    let sheets = excel_reader::read_workbook(path)
        .ok()
        .map(|dump| dump.sheets.into_iter().map(|sheet| sheet.name).collect());
    InputWorkbookTrace {
        path: path.to_path_buf(),
        sha256,
        sheets,
    }
}

/// A preserved raw worker output and its hash.
pub struct RawSdkArchive {
    pub path: PathBuf,
    pub sha256: String,
}

/// Copy the raw candidate into `<diagnostics>/raw-sdk/` without ever
/// overwriting an earlier archive.
pub fn archive_raw_candidate(candidate: &Path, diagnostics_dir: &Path) -> Result<RawSdkArchive> {
    let dir = diagnostics_dir.join("raw-sdk");
    fs::create_dir_all(&dir).context("create raw SDK archive directory")?;
    let stem = candidate
        .file_stem()
        .map(|value| value.to_string_lossy().into_owned())
        .unwrap_or_else(|| "candidate".to_owned());
    for attempt in 0..256u32 {
        let suffix = if attempt == 0 {
            String::new()
        } else {
            format!("-{attempt}")
        };
        let target = dir.join(format!("{stem}{suffix}.xlsx"));
        match crate::file_commit::create_target_from_candidate(candidate, &target) {
            Ok(report) => {
                return Ok(RawSdkArchive {
                    path: target,
                    sha256: report.sha256_after,
                })
            }
            Err(error)
                if error
                    .root_cause()
                    .downcast_ref::<std::io::Error>()
                    .is_some_and(|e| e.kind() == std::io::ErrorKind::AlreadyExists) =>
            {
                continue
            }
            Err(error) => {
                return Err(error).context("archive raw Aspose SDK output without overwrite")
            }
        }
    }
    anyhow::bail!("could not reserve a non-overwriting raw SDK archive path")
}

/// Per-entry outcome of the evaluation-sheet post-processing.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum EvaluationPostOutcome {
    /// The exact-name sheet was removed and the processed file re-verified.
    Processed,
    /// No extra system sheet needs processing; no Rust mutation was performed.
    NotPresent,
    /// A related input already has an exact-named worksheet; not removed.
    ConflictSkipped { input: PathBuf },
    /// A related input's worksheet list could not be reliably obtained.
    InputUnknown { input: PathBuf },
    /// No related input workbook is known; creation source not confirmed here.
    NoInputSourceUnverified,
    /// The output is not a safe candidate for removal (structure references).
    UnsafeSkipped { reason: String },
    /// Removal was attempted but verification failed; candidate untouched.
    Failed { reason: String },
}

impl EvaluationPostOutcome {
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Processed => "processed",
            Self::NotPresent => "not_present",
            Self::ConflictSkipped { .. } => "conflict_skipped",
            Self::InputUnknown { .. } => "input_unknown",
            Self::NoInputSourceUnverified => "no_input_source_unverified",
            Self::UnsafeSkipped { .. } => "unsafe_skipped",
            Self::Failed { .. } => "failed",
        }
    }
}

/// Rewrite `key="value"` inside an attribute string, returning the new full
/// element attributes text. Returns `None` when the key is absent.
#[cfg(test)]
fn replace_attr(attrs: &str, key: &str, value: &str) -> Option<String> {
    let pattern = format!("{key}=\"");
    let found = attrs.find(&pattern)?;
    let rest = &attrs[found + pattern.len()..];
    let end = rest.find('"')?;
    let escaped = value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;");
    Some(format!(
        "{}{}{}",
        &attrs[..found + pattern.len()],
        escaped,
        &rest[end..]
    ))
}

/// Verify the roster already checked independently against the request plan
/// by excel-aspose. This is a second read-only check, not permission to delete.
pub fn verify_sanitized_business_sheets(
    candidate: &Path,
    expected: &[String],
) -> Result<EvaluationPostOutcome> {
    let saved = excel_reader::read_workbook(candidate)?;
    let names: Vec<String> = saved.sheets.into_iter().map(|s| s.name).collect();
    anyhow::ensure!(
        !expected.is_empty() && names == expected,
        "Aspose candidate differs from the verified business worksheet roster"
    );
    Ok(EvaluationPostOutcome::NotPresent)
}

/// Compatibility API retained for existing callers and historical diagnostics.
/// It is read-only: a name, even an exact match, never proves system ownership.
pub fn remove_evaluation_warning_sheet(
    candidate: &Path,
    inputs: &[InputWorkbookTrace],
) -> EvaluationPostOutcome {
    if inputs.is_empty() {
        return EvaluationPostOutcome::NoInputSourceUnverified;
    }
    for input in inputs {
        match &input.sheets {
            None => return EvaluationPostOutcome::InputUnknown { input: input.path.clone() },
            Some(names) if names.iter().any(|s| s == EVALUATION_SHEET_NAME) => {
                return EvaluationPostOutcome::ConflictSkipped { input: input.path.clone() };
            }
            Some(_) => {}
        }
    }
    match excel_reader::read_workbook(candidate) {
        Ok(book) if book.sheets.iter().any(|s| s.name == EVALUATION_SHEET_NAME) => {
            EvaluationPostOutcome::UnsafeSkipped {
                reason: "Rust worksheet deletion is disabled; Worker must supply a verified business roster".to_owned(),
            }
        }
        Ok(_) => EvaluationPostOutcome::NotPresent,
        Err(error) => EvaluationPostOutcome::Failed { reason: format!("{error:#}") },
    }
}

/// Write the traceability record next to a raw archive.
pub fn write_trace_record(
    archive: &RawSdkArchive,
    run_id: &str,
    request_id: &str,
    engine: &str,
    evaluation_output: bool,
    inputs: &[InputWorkbookTrace],
    outcome: &EvaluationPostOutcome,
) -> Result<PathBuf> {
    let record_path = archive.path.with_extension("json");
    let mut file = fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&record_path)
        .with_context(|| format!("create trace record {}", record_path.display()))?;
    let inputs_json: Vec<serde_json::Value> = inputs
        .iter()
        .map(|input| {
            serde_json::json!({
                "path": input.path,
                "sha256": input.sha256,
                "sheets": input.sheets,
            })
        })
        .collect();
    serde_json::to_writer_pretty(
        &mut file,
        &serde_json::json!({
            "schema_version": 1,
            "event_name": "aspose_raw_sdk_output_archived",
            // Legacy event/field names retained; this is not the SDK Save stage.
            "archive_stage": "worker_output_after_worker_preservation_before_rust_postprocessing",
            "input_inventory_capture_phase": "after_worker_response",
            "logical_business_inputs_complete": false,
            "timestamp": chrono::Local::now().to_rfc3339(),
            "engine": engine,
            "run_id": run_id,
            "request_id": request_id,
            "archive": archive.path,
            "archive_sha256": archive.sha256,
            "evaluation_output": evaluation_output,
            "inputs": inputs_json,
            "evaluation_sheet_post_processing": outcome.as_str(),
            "rust_candidate_policy": "verify_business_roster_without_mutation",
        }),
    )?;
    use std::io::Write;
    file.flush()?;
    file.sync_all()?;
    Ok(record_path)
}

#[cfg(test)]
mod tests {
    use super::replace_attr;

    #[test]
    fn attribute_replacement_retains_closing_quote_and_other_attributes() {
        for (attributes, key, value) in [
            ("activeTab=\"2\" firstSheet=\"0\"", "activeTab", "1"),
            ("name=\"Local\" localSheetId=\"2\"", "localSheetId", "1"),
            ("r=\"A1\" i=\"3\" s=\"1\"", "i", "2"),
            ("name=\"Old\" localSheetId=\"0\"", "name", "A&B"),
        ] {
            let fixed = replace_attr(attributes, key, value).unwrap();
            let xml = format!("<element {fixed}/>");
            let doc = roxmltree::Document::parse(&xml).unwrap();
            assert_eq!(doc.root_element().attribute(key), Some(value));
            assert_eq!(doc.root_element().attributes().len(), 2 + usize::from(key == "i"));
        }
    }
}
