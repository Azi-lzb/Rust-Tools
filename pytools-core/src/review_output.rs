//! Retain readable candidates for manual review without approving or committing them.
use anyhow::{Context, Result};
use excel_editor::{ExcelEditPlan, ValidationFailureObserver, ValidationReviewOutput};
use std::{
    cell::RefCell,
    fs,
    io::Write,
    path::{Path, PathBuf},
};

#[derive(Clone)]
struct Origin {
    path: PathBuf,
    expected_hash: String,
}

thread_local! {
    static ORIGIN: RefCell<Option<Origin>> = const { RefCell::new(None) };
}

pub(crate) struct ReviewOriginScope(Option<Origin>, std::marker::PhantomData<std::rc::Rc<()>>);
impl ReviewOriginScope {
    pub(crate) fn enter(source: &Path, expected_hash: &str) -> Self {
        Self(
            ORIGIN.with(|v| {
                v.replace(Some(Origin {
                    path: source.to_path_buf(),
                    expected_hash: expected_hash.to_owned(),
                }))
            }),
            std::marker::PhantomData,
        )
    }
}
impl Drop for ReviewOriginScope {
    fn drop(&mut self) {
        ORIGIN.with(|v| {
            v.replace(self.0.take());
        });
    }
}

pub(crate) struct ReviewPublisher;
impl ValidationFailureObserver for ReviewPublisher {
    fn preserve_for_review(
        &self,
        plan: &ExcelEditPlan,
        candidate: &Path,
        issue: &str,
    ) -> Result<Option<ValidationReviewOutput>> {
        preserve_failed_candidate(&plan.input, candidate, &plan.request_id, issue)
    }
}

fn concise_validation_issue(issue: &str) -> String {
    // Keep the invariant and cell/sheet location, without duplicating literal
    // cell contents from an "expected ..., got ..." diagnostic into the new log.
    let line = issue
        .lines()
        .next()
        .unwrap_or("business validation mismatch");
    let end = [
        ": expected ",
        "; expected ",
        " expected ",
        ", got ",
        ", actual ",
        "，期望",
        "，实际",
    ]
    .iter()
    .filter_map(|marker| line.find(marker))
    .min()
    .unwrap_or(line.len());
    line[..end].chars().take(512).collect()
}

/// Only call at a business-validation failure boundary, never for arbitrary
/// engine exceptions or incomplete saves. This deliberately returns no approval.
pub fn preserve_failed_candidate(
    source: &Path,
    candidate: &Path,
    request_id: &str,
    issue: &str,
) -> Result<Option<ValidationReviewOutput>> {
    match fs::symlink_metadata(candidate) {
        Ok(value) if value.is_file() && !value.file_type().is_symlink() && value.len() > 0 => {}
        _ => return Ok(None),
    };
    // A broken/unreadable container is not a usable review workbook.
    if excel_reader::read_workbook(candidate).is_err() {
        return Ok(None);
    }
    let issue = concise_validation_issue(issue);
    let origin = ORIGIN.with(|v| v.borrow().clone());
    let source = origin.as_ref().map_or(source, |v| v.path.as_path());
    let expected_hash = origin.as_ref().map(|v| v.expected_hash.clone());
    let observed_hash = crate::file_commit::sha256_file(source).ok();
    let source_changed = expected_hash
        .as_ref()
        .zip(observed_hash.as_ref())
        .map(|(a, b)| a != b);
    for _ in 0..256 {
        let output = crate::file_commit::next_result_copy_path(source, "待核对_业务校验不符");
        let publication = match crate::file_commit::create_target_from_candidate(candidate, &output)
        {
            Ok(value) => value,
            Err(error)
                if error
                    .root_cause()
                    .downcast_ref::<std::io::Error>()
                    .is_some_and(|e| e.kind() == std::io::ErrorKind::AlreadyExists) =>
            {
                continue
            }
            Err(error) => {
                return Err(error).context("save review copy without replacing the source")
            }
        };
        let record_path = output.with_file_name(format!(
            "{}.run.json",
            output.file_name().unwrap().to_string_lossy()
        ));
        let record_result = (|| -> Result<()> {
            let mut file = fs::OpenOptions::new()
                .write(true)
                .create_new(true)
                .open(&record_path)?;
            serde_json::to_writer_pretty(
                &mut file,
                &serde_json::json!({
                    "schema_version": 1,
                    "timestamp": chrono::Local::now().to_rfc3339(),
                    "event_name": "business_validation_review_copy_created",
                    "request_id": request_id,
                    "status": "review_required",
                    "business_validation_passed": false,
                    "original_target_replaced": false,
                    "source": source,
                    "review_copy": output,
                    "review_copy_sha256": publication.sha256_after,
                    "expected_source_sha256": expected_hash,
                    "observed_source_sha256": observed_hash,
                    "source_changed": source_changed,
                "validation_issue": &issue,
                    "new_edits_merged": false,
                }),
            )?;
            file.flush()?;
            file.sync_all()?;
            Ok(())
        })();
        let record_path = match record_result {
            Ok(()) => Some(record_path),
            Err(error) => {
                eprintln!("待核对副本已保存，但独立运行记录写入失败：{error:#}");
                None
            }
        };
        eprintln!("业务校验与预期不符，原文件未覆盖；待核对副本已保存：{}。此副本未通过业务校验，请自行核对。", output.display());
        return Ok(Some(ValidationReviewOutput {
            path: output,
            validation_issue: issue.clone(),
            record_path,
        }));
    }
    anyhow::bail!("could not reserve a non-overwriting review-copy path")
}

#[derive(Debug)]
struct ReviewSaved(ValidationReviewOutput);
impl std::fmt::Display for ReviewSaved {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            f,
            "业务校验未通过，原文件未覆盖；待核对副本：{}",
            self.0.path.display()
        )
    }
}
impl std::error::Error for ReviewSaved {}

/// Preserve a domain-specific validation failure before its caller cleans up.
pub fn retain_validation_error(
    source: &Path,
    candidate: &Path,
    request_id: &str,
    error: anyhow::Error,
) -> anyhow::Error {
    if error.downcast_ref::<ReviewSaved>().is_some()
        || error
            .downcast_ref::<excel_editor::ExcelModifierError>()
            .is_some_and(|e| e.review_output.is_some())
    {
        return error;
    }
    match preserve_failed_candidate(source, candidate, request_id, &error.to_string()) {
        Ok(Some(review)) => error.context(ReviewSaved(review)),
        Ok(None) => error,
        Err(retention_error) => error.context(format!("待核对副本保存失败：{retention_error:#}")),
    }
}
