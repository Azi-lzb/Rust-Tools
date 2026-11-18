use anyhow::Result;
use pytools_core::{
    file_commit::{sha256_file, snapshot_workbook},
    review_output::preserve_failed_candidate,
};
use rust_xlsxwriter::Workbook;
use std::{
    fs,
    path::{Path, PathBuf},
};

struct Dir(PathBuf);
impl Dir {
    fn new() -> Result<Self> {
        let stamp = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)?
            .as_nanos();
        let path =
            std::env::temp_dir().join(format!("review-output-{}-{stamp}", std::process::id()));
        fs::create_dir(&path)?;
        Ok(Self(path))
    }
    fn p(&self, name: &str) -> PathBuf {
        self.0.join(name)
    }
}
impl Drop for Dir {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}
fn book(path: &Path, text: &str) -> Result<()> {
    let mut b = Workbook::new();
    b.add_worksheet().write_string(0, 0, text)?;
    b.save(path)?;
    Ok(())
}

fn document(path: &Path, broken_header: bool) -> Result<()> {
    use std::io::Write;
    let mut zip = zip::ZipWriter::new(fs::File::create(path)?);
    let options = zip::write::SimpleFileOptions::default();
    for (name, text) in [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#,
        ),
        (
            "_rels/.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#,
        ),
        (
            "word/document.xml",
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Evaluation Only. Created with Aspose.Words.</w:t></w:r></w:p></w:body></w:document>"#,
        ),
        (
            "word/header1.xml",
            if broken_header {
                "<broken"
            } else {
                "<header/>"
            },
        ),
    ] {
        zip.start_file(name, options)?;
        zip.write_all(text.as_bytes())?;
    }
    zip.finish()?;
    Ok(())
}

#[test]
fn doc_conversion_review_preserves_bytes_and_evaluation_in_actual_docx_format() -> Result<()> {
    let d = Dir::new()?;
    let source = d.p("original.doc");
    fs::write(&source, b"synthetic DOC source placeholder")?;
    let before = fs::read(&source)?;
    let candidate = d.p("candidate.docx");
    document(&candidate, false)?;
    let review =
        preserve_failed_candidate(&source, &candidate, "words", "Word paragraph count changed")?
            .unwrap();
    assert_eq!(review.path.extension().unwrap(), "docx");
    assert_eq!(fs::read(&source)?, before);
    assert_eq!(fs::read(&review.path)?, fs::read(&candidate)?);
    let record: serde_json::Value =
        serde_json::from_slice(&fs::read(review.record_path.unwrap())?)?;
    assert_eq!(record["business_validation_passed"], false);
    assert_eq!(record["original_target_replaced"], false);
    assert_eq!(record["validation_issue"], "Word paragraph count changed");
    Ok(())
}

#[test]
fn malformed_docx_auxiliary_xml_is_not_published_as_review() -> Result<()> {
    let d = Dir::new()?;
    let source = d.p("source.docx");
    let candidate = d.p("candidate.docx");
    document(&source, false)?;
    document(&candidate, true)?;
    assert!(
        preserve_failed_candidate(&source, &candidate, "broken", "header cannot be read")?
            .is_none()
    );
    assert_eq!(fs::read_dir(&d.0)?.count(), 2);
    Ok(())
}

#[test]
fn semantic_failure_keeps_both_original_and_pending_copy_with_explicit_status() -> Result<()> {
    let d = Dir::new()?;
    let source = d.p("report.xlsx");
    let candidate = d.p("candidate.xlsx");
    book(&source, "original")?;
    book(&candidate, "unexpected header")?;
    let before = fs::read(&source)?;
    let result = fs::read(&candidate)?;
    let review = preserve_failed_candidate(
        &source,
        &candidate,
        "run-1",
        "Data!A1: literal was not preserved",
    )?
    .unwrap();
    assert_eq!(fs::read(&source)?, before);
    assert_eq!(fs::read(&candidate)?, result);
    assert_eq!(fs::read(&review.path)?, result);
    assert!(review
        .path
        .file_name()
        .unwrap()
        .to_string_lossy()
        .contains("待核对_业务校验不符"));
    let record: serde_json::Value =
        serde_json::from_slice(&fs::read(review.record_path.unwrap())?)?;
    assert_eq!(record["status"], "review_required");
    assert_eq!(record["business_validation_passed"], false);
    assert_eq!(record["original_target_replaced"], false);
    assert_eq!(record["request_id"], "run-1");
    assert_eq!(record["review_copy_sha256"], sha256_file(&candidate)?);
    Ok(())
}

#[test]
fn review_record_keeps_location_without_copying_expected_or_actual_cell_contents() -> Result<()> {
    let d = Dir::new()?;
    let source = d.p("report.xlsx");
    let candidate = d.p("candidate.xlsx");
    book(&source, "original")?;
    book(&candidate, "unexpected")?;
    let review = preserve_failed_candidate(
        &source,
        &candidate,
        "privacy",
        "unmodified cell value changed at Data!A1: expected SECRET_EXPECTED, got SECRET_ACTUAL",
    )?
    .unwrap();
    let record = fs::read_to_string(review.record_path.unwrap())?;
    assert!(record.contains("Data!A1"));
    assert!(!record.contains("SECRET_EXPECTED"));
    assert!(!record.contains("SECRET_ACTUAL"));
    assert_eq!(
        review.validation_issue,
        "unmodified cell value changed at Data!A1"
    );
    Ok(())
}

#[test]
fn repeated_review_outputs_never_replace_previous_copies_or_records() -> Result<()> {
    let d = Dir::new()?;
    let source = d.p("report.xlsx");
    let candidate = d.p("candidate.xlsx");
    book(&source, "keep")?;
    book(&candidate, "first")?;
    let first = preserve_failed_candidate(&source, &candidate, "first", "mismatch")?.unwrap();
    let data = fs::read(&first.path)?;
    let record = fs::read(first.record_path.as_ref().unwrap())?;
    book(&candidate, "second")?;
    let second = preserve_failed_candidate(&source, &candidate, "second", "mismatch")?.unwrap();
    assert_ne!(first.path, second.path);
    assert_eq!(fs::read(&first.path)?, data);
    assert_eq!(fs::read(first.record_path.as_ref().unwrap())?, record);
    assert_eq!(fs::read(second.path)?, fs::read(&candidate)?);
    Ok(())
}

#[test]
fn unreadable_empty_and_missing_candidates_are_not_published_as_review_workbooks() -> Result<()> {
    let d = Dir::new()?;
    let source = d.p("report.xlsx");
    let candidate = d.p("candidate.xlsx");
    book(&source, "keep")?;
    let before = fs::read(&source)?;
    for contents in [b"broken zip".as_slice(), b"".as_slice()] {
        fs::write(&candidate, contents)?;
        assert!(preserve_failed_candidate(&source, &candidate, "bad", "unreadable")?.is_none());
    }
    fs::remove_file(&candidate)?;
    assert!(preserve_failed_candidate(&source, &candidate, "missing", "missing")?.is_none());
    assert_eq!(fs::read(&source)?, before);
    assert_eq!(fs::read_dir(&d.0)?.count(), 1);
    Ok(())
}

#[test]
fn extra_business_validation_retains_one_copy_before_candidate_cleanup_and_keeps_error_chain(
) -> Result<()> {
    let d = Dir::new()?;
    let source = d.p("report.xlsx");
    let candidate = d.p("candidate.xlsx");
    book(&source, "keep")?;
    book(&candidate, "review")?;
    let before = fs::read(&source)?;
    let error = pytools_core::review_output::retain_validation_error(
        &source,
        &candidate,
        "extra-contract",
        anyhow::anyhow!("extra append contract mismatch"),
    );
    let error = pytools_core::review_output::retain_validation_error(
        &source,
        &candidate,
        "extra-contract",
        error,
    );
    assert!(format!("{error:#}").contains("extra append contract mismatch"));
    assert!(error.to_string().contains("待核对副本"));
    let copies = fs::read_dir(&d.0)?
        .filter_map(|p| p.ok())
        .map(|p| p.path())
        .filter(|p| {
            p.extension().is_some_and(|e| e == "xlsx")
                && p.file_name().unwrap().to_string_lossy().contains("待核对")
        })
        .collect::<Vec<_>>();
    assert_eq!(copies.len(), 1);
    fs::remove_file(candidate)?;
    assert_eq!(excel_reader::read_workbook(&copies[0])?.sheets.len(), 1);
    assert_eq!(fs::read(source)?, before);
    Ok(())
}

#[test]
fn snapshot_origin_records_business_source_and_concurrent_edits_without_merging() -> Result<()> {
    let d = Dir::new()?;
    let source = d.p("report.xlsx");
    let candidate = d.p("candidate.xlsx");
    book(&source, "old")?;
    let expected = sha256_file(&source)?;
    let snapshot = snapshot_workbook(&source, &expected)?;
    book(&source, "new user data")?;
    let newer = fs::read(&source)?;
    book(&candidate, "unverified processing result")?;
    let review =
        preserve_failed_candidate(snapshot.path(), &candidate, "snapshot-run", "mismatch")?
            .unwrap();
    assert!(review
        .path
        .file_name()
        .unwrap()
        .to_string_lossy()
        .starts_with("report_"));
    let record: serde_json::Value =
        serde_json::from_slice(&fs::read(review.record_path.unwrap())?)?;
    assert_eq!(record["source"], source.to_string_lossy().as_ref());
    assert_eq!(record["source_changed"], true);
    assert_eq!(record["expected_source_sha256"], expected);
    assert_eq!(record["new_edits_merged"], false);
    assert_eq!(fs::read(&source)?, newer);
    Ok(())
}

#[test]
fn long_unicode_names_remain_openable_and_preserve_extension() -> Result<()> {
    let d = Dir::new()?;
    let source = d.p(&format!("{}.xlsx", "合成报表".repeat(18)));
    let candidate = d.p("candidate.xlsx");
    book(&source, "keep")?;
    book(&candidate, "review")?;
    let review = preserve_failed_candidate(&source, &candidate, "long-name", "mismatch")?.unwrap();
    assert_eq!(review.path.extension().unwrap(), "xlsx");
    assert!(review.path.file_name().unwrap().to_string_lossy().len() <= 255);
    assert!(
        review
            .record_path
            .as_ref()
            .unwrap()
            .file_name()
            .unwrap()
            .to_string_lossy()
            .len()
            <= 255
    );
    assert_eq!(excel_reader::read_workbook(&review.path)?.sheets.len(), 1);
    Ok(())
}
