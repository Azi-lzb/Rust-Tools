use anyhow::Result;
use excel_reader::read_workbook;
use pytools_core::file_commit::{
    commit_validated_candidate, create_target_from_candidate, replace_target_from_candidate,
    sha256_file, verify_planning_snapshot,
};
use rust_xlsxwriter::Workbook;
use std::{
    fs,
    path::{Path, PathBuf},
};

struct TestDir(PathBuf);

impl TestDir {
    fn new() -> Self {
        let path = std::env::temp_dir().join(format!(
            "pytools-file-commit-{}-{:?}-{}",
            std::process::id(),
            std::thread::current().id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        fs::create_dir_all(&path).unwrap();
        Self(path)
    }

    fn path(&self, name: &str) -> PathBuf {
        self.0.join(name)
    }
}

impl Drop for TestDir {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

fn write_book(path: &Path, value: &str) -> Result<()> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Data")?;
    sheet.write_string(0, 0, value)?;
    workbook.save(path)?;
    Ok(())
}

#[test]
fn validated_commit_replaces_unchanged_target_without_conflict_record() -> Result<()> {
    let dir = TestDir::new();
    let target = dir.path("report.xlsx");
    let candidate = dir.path("candidate.xlsx");
    write_book(&target, "original")?;
    write_book(&candidate, "processed")?;
    let report = commit_validated_candidate(&candidate, &target, &sha256_file(&target)?)?;
    assert_eq!(report.operation, "replaced");
    assert_eq!(report.target, target);
    assert!(!report.is_conflict_copy());
    assert_eq!(fs::read(&target)?, fs::read(&candidate)?);
    assert_eq!(fs::read_dir(&dir.0)?.count(), 2);
    Ok(())
}

#[test]
fn validated_commit_preserves_new_edits_and_saves_result_with_audit_record() -> Result<()> {
    let dir = TestDir::new();
    let target = dir.path("report.xlsx");
    let candidate = dir.path("candidate.xlsx");
    write_book(&target, "planning snapshot")?;
    let expected = sha256_file(&target)?;
    write_book(&candidate, "processed snapshot")?;
    write_book(&target, "user's new edits")?;
    let newer = fs::read(&target)?;
    let candidate_bytes = fs::read(&candidate)?;
    let report = commit_validated_candidate(&candidate, &target, &expected)?;
    assert_eq!(report.operation, "conflict_copy_created");
    assert!(report.is_conflict_copy());
    assert_ne!(report.target, target);
    assert_eq!(report.target.parent(), target.parent());
    assert_eq!(fs::read(&target)?, newer);
    assert_eq!(fs::read(&candidate)?, candidate_bytes);
    assert_eq!(fs::read(&report.target)?, candidate_bytes);
    assert_eq!(
        read_workbook(&report.target)?.sheets[0].values["A1"]["value"],
        "processed snapshot"
    );
    let conflict = report.conflict.as_ref().unwrap();
    assert_eq!(conflict.requested_target, target);
    assert_eq!(conflict.expected_sha256, expected);
    assert_eq!(conflict.observed_sha256, sha256_file(&target)?);
    let audit: serde_json::Value =
        serde_json::from_slice(&fs::read(conflict.record_path.as_ref().unwrap())?)?;
    assert_eq!(audit["event_name"], "target_changed_result_saved");
    assert_eq!(audit["original_target_replaced"], false);
    assert_eq!(audit["new_edits_merged"], false);
    assert_eq!(
        audit["commit"]["target"],
        report.target.to_string_lossy().as_ref()
    );
    assert!(report.conflict_notice().unwrap().contains("未合并"));
    Ok(())
}

#[test]
fn repeated_conflicts_create_distinct_results_and_keep_previous_results() -> Result<()> {
    let dir = TestDir::new();
    let target = dir.path("report.xlsx");
    let candidate = dir.path("candidate.xlsx");
    write_book(&target, "old")?;
    let expected = sha256_file(&target)?;
    write_book(&target, "new user data")?;
    write_book(&candidate, "first result")?;
    let first = commit_validated_candidate(&candidate, &target, &expected)?;
    let first_bytes = fs::read(&first.target)?;
    let first_record = fs::read(
        first
            .conflict
            .as_ref()
            .unwrap()
            .record_path
            .as_ref()
            .unwrap(),
    )?;
    write_book(&candidate, "second result")?;
    let second = commit_validated_candidate(&candidate, &target, &expected)?;
    assert_ne!(first.target, second.target);
    assert_eq!(fs::read(&first.target)?, first_bytes);
    assert_eq!(
        fs::read(
            first
                .conflict
                .as_ref()
                .unwrap()
                .record_path
                .as_ref()
                .unwrap()
        )?,
        first_record
    );
    assert_eq!(fs::read(&second.target)?, fs::read(&candidate)?);
    assert_eq!(
        read_workbook(&target)?.sheets[0].values["A1"]["value"],
        "new user data"
    );
    Ok(())
}

#[test]
fn conflict_does_not_publish_invalid_candidate_or_candidate_alias() -> Result<()> {
    let dir = TestDir::new();
    let target = dir.path("report.xlsx");
    let candidate = dir.path("invalid.xlsx");
    write_book(&target, "old")?;
    let expected = sha256_file(&target)?;
    write_book(&target, "new")?;
    let newer = fs::read(&target)?;
    fs::write(&candidate, b"broken ZIP")?;
    assert!(commit_validated_candidate(&candidate, &target, &expected).is_err());
    assert!(commit_validated_candidate(&target, &target, &expected).is_err());
    assert_eq!(fs::read(&target)?, newer);
    assert_eq!(fs::read_dir(&dir.0)?.count(), 2);
    Ok(())
}

#[test]
fn long_unicode_target_names_keep_extension_and_fit_component_limits() -> Result<()> {
    let dir = TestDir::new();
    let target = dir.path(&format!("{}.xlsx", "合成报表".repeat(18)));
    let candidate = dir.path("candidate.xlsx");
    write_book(&target, "old")?;
    let expected = sha256_file(&target)?;
    write_book(&target, "new")?;
    write_book(&candidate, "result")?;
    let report = commit_validated_candidate(&candidate, &target, &expected)?;
    assert_eq!(report.target.extension().unwrap(), "xlsx");
    for p in [
        &report.target,
        report
            .conflict
            .as_ref()
            .unwrap()
            .record_path
            .as_ref()
            .unwrap(),
    ] {
        let name = p.file_name().unwrap().to_string_lossy();
        assert!(name.len() <= 255);
        assert!(name.encode_utf16().count() <= 255);
    }
    assert_eq!(
        read_workbook(&target)?.sheets[0].values["A1"]["value"],
        "new"
    );
    Ok(())
}

#[test]
fn create_target_publishes_valid_workbook_without_overwrite() -> Result<()> {
    let dir = TestDir::new();
    let candidate = dir.path("candidate.xlsx");
    let target = dir.path("target.xlsx");
    write_book(&candidate, "candidate")?;
    let candidate_hash = sha256_file(&candidate)?;

    let report = create_target_from_candidate(&candidate, &target)?;
    assert_eq!(report.operation, "created");
    assert_eq!(report.sha256_after, candidate_hash);
    assert!(report.sha256_before.is_none());
    assert!(report.elapsed_ms < 60_000);
    assert_eq!(fs::read(&target)?, fs::read(&candidate)?);
    assert_eq!(read_workbook(&target)?.sheets[0].name, "Data");

    let target_before = fs::read(&target)?;
    let error = create_target_from_candidate(&candidate, &target)
        .expect_err("existing target must never be overwritten");
    assert!(format!("{error:#}").contains("target already exists"));
    assert_eq!(fs::read(&target)?, target_before);
    Ok(())
}

#[test]
fn replace_target_reopens_candidate_and_reports_before_after_hashes() -> Result<()> {
    let dir = TestDir::new();
    let candidate = dir.path("candidate.xlsx");
    let target = dir.path("target.xlsx");
    write_book(&candidate, "new")?;
    write_book(&target, "old")?;
    let target_before = sha256_file(&target)?;
    let candidate_after = sha256_file(&candidate)?;

    let report = replace_target_from_candidate(&candidate, &target, &target_before)?;
    assert_eq!(report.operation, "replaced");
    assert_eq!(
        report.sha256_before.as_deref(),
        Some(target_before.as_str())
    );
    assert_eq!(report.sha256_after, candidate_after);
    assert_eq!(fs::read(&target)?, fs::read(&candidate)?);
    let workbook = read_workbook(&target)?;
    assert_eq!(workbook.sheets[0].values["A1"]["value"], "new");
    Ok(())
}

#[test]
fn replace_target_refuses_stale_hash_without_touching_newer_target() -> Result<()> {
    let dir = TestDir::new();
    let candidate = dir.path("candidate.xlsx");
    let target = dir.path("target.xlsx");
    write_book(&candidate, "candidate")?;
    write_book(&target, "planned")?;
    let planned_hash = sha256_file(&target)?;
    write_book(&target, "newer")?;
    let newer_bytes = fs::read(&target)?;

    let error = replace_target_from_candidate(&candidate, &target, &planned_hash)
        .expect_err("stale plan must be rejected");
    assert!(format!("{error:#}").contains("changed after planning"));
    assert_eq!(fs::read(&target)?, newer_bytes);
    assert_eq!(
        read_workbook(&target)?.sheets[0].values["A1"]["value"],
        "newer"
    );
    Ok(())
}

#[test]
fn invalid_candidate_is_rejected_before_existing_target_is_changed() -> Result<()> {
    let dir = TestDir::new();
    let candidate = dir.path("candidate.xlsx");
    let target = dir.path("target.xlsx");
    fs::write(&candidate, b"not an xlsx")?;
    write_book(&target, "original")?;
    let original_bytes = fs::read(&target)?;

    let error = replace_target_from_candidate(&candidate, &target, &sha256_file(&target)?)
        .expect_err("invalid candidate must be rejected");
    assert!(format!("{error:#}").contains("validate candidate workbook"));
    assert_eq!(fs::read(&target)?, original_bytes);
    assert_eq!(
        fs::read_dir(&dir.0)?.count(),
        2,
        "failed candidate validation must not leave staging or backup files"
    );
    Ok(())
}

#[test]
fn planning_snapshot_rejects_changed_or_new_targets() -> Result<()> {
    let dir = TestDir::new();
    let target = dir.path("target.xlsx");

    verify_planning_snapshot(&target, None)?;
    write_book(&target, "planned")?;
    let planned_hash = sha256_file(&target)?;
    verify_planning_snapshot(&target, Some(&planned_hash))?;

    write_book(&target, "changed")?;
    let error = verify_planning_snapshot(&target, Some(&planned_hash))
        .expect_err("changed target must invalidate the plan");
    assert!(format!("{error:#}").contains("changed while the plan was being built"));

    let absent_target = dir.path("absent.xlsx");
    write_book(&absent_target, "appeared")?;
    let error = verify_planning_snapshot(&absent_target, None)
        .expect_err("target created during planning must invalidate the plan");
    assert!(format!("{error:#}").contains("target appeared while the plan was being built"));
    Ok(())
}
