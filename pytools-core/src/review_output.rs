//! Retain readable candidates for manual review without approving or committing them.
use anyhow::{Context, Result};
use excel_editor::{ExcelEditPlan, ValidationFailureObserver, ValidationReviewOutput};
use std::{
    cell::RefCell,
    fs,
    io::{Read, Write},
    path::{Path, PathBuf},
};

#[derive(Clone)]
struct Origin {
    path: PathBuf,
    expected_hash: String,
}

thread_local! {
    static ORIGIN: RefCell<Option<Origin>> = const { RefCell::new(None) };
    static DESTINATION: RefCell<Option<PathBuf>> = const { RefCell::new(None) };
}

/// Keep review artifacts beside the user-facing output, never a private seed.
pub(crate) struct ReviewDestinationScope(
    Option<PathBuf>,
    std::marker::PhantomData<std::rc::Rc<()>>,
);
impl ReviewDestinationScope {
    pub(crate) fn enter_if_unset(target: &Path) -> Self {
        let target = DESTINATION
            .with(|v| v.borrow().clone())
            .unwrap_or_else(|| target.to_path_buf());
        Self::enter(&target)
    }
    pub(crate) fn enter(target: &Path) -> Self {
        Self(
            DESTINATION.with(|v| v.replace(Some(target.to_path_buf()))),
            std::marker::PhantomData,
        )
    }
}
impl Drop for ReviewDestinationScope {
    fn drop(&mut self) {
        DESTINATION.with(|v| {
            v.replace(self.0.take());
        });
    }
}

fn is_document(path: &Path) -> bool {
    path.extension()
        .is_some_and(|e| e.eq_ignore_ascii_case("docx"))
}

/// A readable review document is still unapproved. Never remove evaluation content.
fn validate_readable_document(path: &Path) -> Result<()> {
    let mut zip = zip::ZipArchive::new(fs::File::open(path)?)?;
    for part in ["[Content_Types].xml", "_rels/.rels", "word/document.xml"] {
        let mut xml = String::new();
        zip.by_name(part)?.read_to_string(&mut xml)?;
        let doc = roxmltree::Document::parse(&xml)?;
        if part == "word/document.xml" {
            anyhow::ensure!(
                doc.root_element().has_tag_name((
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                    "document"
                )) && doc.descendants().any(|n| n.has_tag_name((
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                    "body"
                ))),
                "DOCX candidate has no document body"
            );
        }
    }
    // Read every member (CRC included), and parse auxiliary XML too. A broken
    // header/media member is a damaged package, not a readable review result.
    for i in 0..zip.len() {
        let mut part = zip.by_index(i)?;
        if part.name().ends_with(".xml") || part.name().ends_with(".rels") {
            let mut xml = String::new();
            part.read_to_string(&mut xml)?;
            roxmltree::Document::parse(&xml)?;
        } else {
            std::io::copy(&mut part, &mut std::io::sink())?;
        }
    }
    Ok(())
}

pub(crate) struct ReviewOriginScope(Option<Origin>, std::marker::PhantomData<std::rc::Rc<()>>);
impl ReviewOriginScope {
    pub(crate) fn enter_if_unset(source: &Path, expected_hash: &str) -> Self {
        match ORIGIN.with(|v| v.borrow().clone()) {
            Some(origin) => Self::enter(&origin.path, &origin.expected_hash),
            None => Self::enter(source, expected_hash),
        }
    }
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
    let document = is_document(candidate);
    let readable = if document {
        validate_readable_document(candidate)
    } else {
        excel_reader::verify_readable_workbook(candidate).map(|_| ())
    };
    if readable.is_err() {
        return Ok(None);
    }
    let issue = concise_validation_issue(issue);
    let origin = ORIGIN.with(|v| v.borrow().clone());
    let source = origin.as_ref().map_or(source, |v| v.path.as_path());
    let destination = DESTINATION.with(|v| v.borrow().clone());
    let target = destination.as_deref().unwrap_or(source);
    let expected_hash = origin.as_ref().map(|v| v.expected_hash.clone());
    let observed_hash = crate::file_commit::sha256_file(source).ok();
    let source_changed = expected_hash
        .as_ref()
        .zip(observed_hash.as_ref())
        .map(|(a, b)| a != b);
    for _ in 0..256 {
        let mut base = target.to_path_buf();
        // XLS→XLSX and DOC→DOCX candidates must retain their actual output format.
        if let Some(extension) = candidate.extension() {
            base.set_extension(extension);
        }
        let output = crate::file_commit::next_result_copy_path(&base, "待核对_业务校验不符");
        let publication = if document {
            crate::file_commit::create_target_from_candidate_with_validator(
                candidate,
                &output,
                |path, expected, _| {
                    anyhow::ensure!(
                        crate::file_commit::sha256_file(path)? == expected,
                        "DOCX review candidate changed during publication"
                    );
                    validate_readable_document(path)
                },
            )
        } else {
            crate::file_commit::create_target_from_candidate(candidate, &output)
        };
        let publication = match publication {
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
                    "planned_target": target,
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

/// Record the real source and destination for a validated-readable but
/// business-rejected candidate from a dedicated legacy workbook route.
pub fn preserve_failed_candidate_at(
    source: &Path,
    expected_source_hash: &str,
    destination: &Path,
    candidate: &Path,
    request_id: &str,
    issue: &str,
) -> Result<Option<ValidationReviewOutput>> {
    let _origin = ReviewOriginScope::enter(source, expected_source_hash);
    let _destination = ReviewDestinationScope::enter(destination);
    preserve_failed_candidate(source, candidate, request_id, issue)
}

#[derive(Debug)]
pub(crate) struct ReviewSaved(ValidationReviewOutput);
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn review_destination_scopes_restore_outer_destination_and_origin() -> Result<()> {
        let root = std::env::temp_dir().join(format!(
            "review-destination-{}-{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)?
                .as_nanos()
        ));
        fs::create_dir_all(root.join("private"))?;
        fs::create_dir_all(root.join("output"))?;
        let source = root.join("original.xls");
        fs::write(&source, b"original untouched XLS fixture placeholder")?;
        let hash = crate::file_commit::sha256_file(&source)?;
        let candidate = root.join("private/candidate.xlsx");
        let mut book = rust_xlsxwriter::Workbook::new();
        book.add_worksheet().write_string(0, 0, "review")?;
        book.save(&candidate)?;
        let _origin = ReviewOriginScope::enter(&source, &hash);
        let target = root.join("output/converted.xlsx");
        {
            let _destination = ReviewDestinationScope::enter(&target);
            {
                let _nested = ReviewDestinationScope::enter_if_unset(&candidate);
                let _nested_origin =
                    ReviewOriginScope::enter_if_unset(&candidate, "private-seed-hash");
                assert_eq!(
                    DESTINATION.with(|v| v.borrow().clone()),
                    Some(target.clone())
                );
                assert_eq!(
                    ORIGIN.with(|v| v.borrow().as_ref().unwrap().path.clone()),
                    source
                );
            }
            {
                let inner = root.join("output/inner.xlsx");
                let _inner = ReviewDestinationScope::enter(&inner);
                assert_eq!(DESTINATION.with(|v| v.borrow().clone()), Some(inner));
            }
            let review = preserve_failed_candidate(
                &candidate,
                &candidate,
                "conversion",
                "Data!A1 mismatch",
            )?
            .unwrap();
            assert_eq!(review.path.parent(), target.parent());
            assert!(review
                .path
                .file_name()
                .unwrap()
                .to_string_lossy()
                .starts_with("converted_待核对"));
            assert_eq!(review.path.extension().unwrap(), "xlsx");
            let record: serde_json::Value =
                serde_json::from_slice(&fs::read(review.record_path.unwrap())?)?;
            assert_eq!(record["source"], source.to_string_lossy().as_ref());
            assert_eq!(record["planned_target"], target.to_string_lossy().as_ref());
            assert_eq!(record["source_changed"], false);
        }
        assert!(DESTINATION.with(|v| v.borrow().is_none()));
        assert_eq!(crate::file_commit::sha256_file(&source)?, hash);
        fs::remove_dir_all(root)?;
        Ok(())
    }
}
