//! Safe publication of validated workbook candidates to configured targets.

use anyhow::{bail, ensure, Context, Result};
use excel_reader::verify_readable_workbook;
use serde::Serialize;
use sha2::{Digest, Sha256};
use std::{
    fs::{self, File, OpenOptions},
    io::{self, Read, Write},
    path::{Path, PathBuf},
    sync::atomic::{AtomicU64, Ordering},
    time::Instant,
};

static SIBLING_SEQUENCE: AtomicU64 = AtomicU64::new(0);

/// Probe an existing target before spending time generating a candidate.
/// Publication still verifies the snapshot and handles races independently.
pub fn ensure_target_writable(path: &Path) -> Result<()> {
    let mut options = OpenOptions::new();
    options.read(true).write(true);
    #[cfg(windows)]
    {
        use std::os::windows::fs::OpenOptionsExt;
        options.share_mode(0);
    }
    let _probe = options.open(path).with_context(|| {
        format!(
            "目标工作簿被占用或没有写入权限：{}；请关闭 Excel/WPS 中的目标文件并检查写入权限后重试",
            path.display()
        )
    })?;
    Ok(())
}

#[cfg(all(test, windows))]
mod target_access_tests {
    use super::*;
    use std::os::windows::fs::OpenOptionsExt;

    #[test]
    fn occupied_target_is_rejected_before_candidate_generation() {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let path = std::env::temp_dir().join(format!(
            "rusttools-target-probe-{}-{nonce}.xlsx",
            std::process::id()
        ));
        fs::write(&path, b"probe fixture").unwrap();
        let held = OpenOptions::new()
            .read(true)
            .write(true)
            .share_mode(0)
            .open(&path)
            .unwrap();
        let error = ensure_target_writable(&path).unwrap_err();
        assert!(error.to_string().contains("被占用"));
        drop(held);
        ensure_target_writable(&path).unwrap();
        assert_eq!(fs::read(&path).unwrap(), b"probe fixture");
        fs::remove_file(&path).unwrap();
    }

    struct LongFixture(PathBuf);
    impl LongFixture {
        fn new(label: &str) -> Self {
            let stamp = std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos();
            let root =
                std::env::temp_dir().join(format!("commit-{label}-{}-{stamp}", std::process::id()));
            fs::create_dir_all(&root).unwrap();
            Self(root)
        }
        fn directory(&self) -> PathBuf {
            let mut path = self.0.clone();
            for _ in 0..5 {
                path.push("业务 Unicode space long-path-regression-0123456789");
            }
            fs::create_dir_all(&path).unwrap();
            assert!(path.as_os_str().len() > 280);
            path
        }
    }
    impl Drop for LongFixture {
        fn drop(&mut self) {
            let _ = fs::remove_dir_all(&self.0);
        }
    }
    fn workbook(path: &Path, value: &str) {
        let mut book = rust_xlsxwriter::Workbook::new();
        book.add_worksheet().write_string(0, 0, value).unwrap();
        book.save(path).unwrap();
    }

    #[test]
    fn long_paths_publish_replace_and_reject_existing_or_stale_targets() {
        let fixture = LongFixture::new("publish");
        let directory = fixture.directory();
        let candidate = directory.join("候选.xlsx");
        let target = directory.join("交付.xlsx");
        workbook(&candidate, "original");
        let original = sha256_file(&candidate).unwrap();
        create_target_from_candidate(&candidate, &target).unwrap();
        assert_eq!(sha256_file(&target).unwrap(), original);
        assert!(create_target_from_candidate(&candidate, &target).is_err());
        let replacement = directory.join("replacement.xlsx");
        workbook(&replacement, "changed");
        assert!(replace_target_from_candidate(&replacement, &target, "stale snapshot").is_err());
        assert_eq!(sha256_file(&target).unwrap(), original);
        replace_target_from_candidate(&replacement, &target, &original).unwrap();
        assert_eq!(
            sha256_file(&target).unwrap(),
            sha256_file(&replacement).unwrap()
        );
        assert_eq!(sha256_file(&candidate).unwrap(), original);

        // Exercise the actual no-replace Win32 boundary after a target appears.
        let source = directory.join("race-source.tmp");
        let appeared = directory.join("race-target.tmp");
        fs::write(&source, b"candidate").unwrap();
        fs::write(&appeared, b"user data").unwrap();
        assert!(publish_new_file(&source, &appeared).is_err());
        assert_eq!(fs::read(&source).unwrap(), b"candidate");
        assert_eq!(fs::read(&appeared).unwrap(), b"user data");
    }

    #[test]
    fn long_paths_keep_originals_on_validation_failure_and_locked_replacement() {
        let fixture = LongFixture::new("failure");
        let directory = fixture.directory();
        let candidate = directory.join("candidate.xlsx");
        let target = directory.join("target.xlsx");
        workbook(&candidate, "candidate");
        let candidate_hash = sha256_file(&candidate).unwrap();
        let rejected =
            create_target_from_candidate_with_validator(&candidate, &target, |path, hash, role| {
                verify_workbook_file(path, hash, role)?;
                ensure!(
                    role != "published target",
                    "injected post-publication validation failure"
                );
                Ok(())
            });
        assert!(rejected.is_err());
        assert!(!target.exists());
        assert_eq!(sha256_file(&candidate).unwrap(), candidate_hash);

        workbook(&target, "user original");
        let original = fs::read(&target).unwrap();
        let original_hash = sha256_file(&target).unwrap();
        let held = OpenOptions::new()
            .read(true)
            .write(true)
            .share_mode(0)
            .open(&target)
            .unwrap();
        assert!(replace_target_from_candidate(&candidate, &target, &original_hash).is_err());
        drop(held);
        assert_eq!(fs::read(&target).unwrap(), original);
        assert_eq!(sha256_file(&candidate).unwrap(), candidate_hash);
    }
}

#[derive(Clone, Debug, Serialize)]
pub struct TargetCommitReport {
    pub operation: String,
    pub target: PathBuf,
    pub sha256_before: Option<String>,
    pub sha256_after: String,
    pub backup_path: Option<PathBuf>,
    pub cleanup_warning: Option<String>,
    pub elapsed_ms: u128,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub conflict: Option<TargetConflict>,
}

#[derive(Clone, Debug, Serialize)]
pub struct TargetConflict {
    pub requested_target: PathBuf,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub changed_source: Option<PathBuf>,
    pub expected_sha256: String,
    pub observed_sha256: String,
    pub reason: String,
    pub record_path: Option<PathBuf>,
}

impl TargetCommitReport {
    pub fn is_conflict_copy(&self) -> bool {
        self.conflict.is_some()
    }

    pub fn conflict_notice(&self) -> Option<String> {
        self.conflict.as_ref().map(|conflict| {
            format!(
                "{} 在运行期间已被修改，保留当前文件；基于处理时数据生成的结果已另存为 {}，未合并运行中的新修改，请自行核对。",
                conflict.changed_source.as_ref().unwrap_or(&conflict.requested_target).display(),
                self.target.display()
            )
        })
    }
}

#[derive(Debug)]
struct TargetSnapshotConflict {
    expected: String,
    observed: String,
}

impl std::fmt::Display for TargetSnapshotConflict {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "configured target changed after planning or during candidate preparation; refusing to replace")
    }
}

impl std::error::Error for TargetSnapshotConflict {}

fn check_replacement_snapshot(expected: &str, observed: &str) -> Result<()> {
    if expected != observed {
        return Err(TargetSnapshotConflict {
            expected: expected.to_owned(),
            observed: observed.to_owned(),
        }
        .into());
    }
    Ok(())
}

pub fn sha256_file(path: &Path) -> Result<String> {
    let mut file =
        File::open(path).with_context(|| format!("open {} for SHA-256", path.display()))?;
    let mut digest = Sha256::new();
    let mut buffer = [0_u8; 64 * 1024];
    loop {
        let count = file
            .read(&mut buffer)
            .with_context(|| format!("read {} for SHA-256", path.display()))?;
        if count == 0 {
            break;
        }
        digest.update(&buffer[..count]);
    }
    Ok(format!("{:x}", digest.finalize()))
}

/// Ensure a target still matches the snapshot used while a plan was built.
/// A missing expected hash means the target was absent when planning began.
pub fn verify_planning_snapshot(path: &Path, expected_sha256: Option<&str>) -> Result<()> {
    match expected_sha256 {
        Some(expected) => {
            let metadata = fs::symlink_metadata(path).with_context(|| {
                format!(
                    "target changed while the plan was being built: {}",
                    path.display()
                )
            })?;
            ensure!(
                !metadata.file_type().is_symlink() && metadata.is_file(),
                "target changed or is not a regular file while the plan was being built: {}",
                path.display()
            );
            let actual = sha256_file(path)?;
            ensure!(
                actual == expected,
                "target changed while the plan was being built: {}",
                path.display()
            );
        }
        None => match fs::symlink_metadata(path) {
            Ok(_) => bail!(
                "target appeared while the plan was being built: {}",
                path.display()
            ),
            Err(error) if error.kind() == io::ErrorKind::NotFound => {}
            Err(error) => {
                return Err(error)
                    .with_context(|| format!("inspect target snapshot {}", path.display()))
            }
        },
    }
    Ok(())
}

/// Publish a validated candidate as a new target without replacing a path
/// which appeared after planning.
pub fn create_target_from_candidate(candidate: &Path, target: &Path) -> Result<TargetCommitReport> {
    create_target_from_candidate_with_validator(candidate, target, |path, hash, role| {
        if role == "published target" {
            // These are the same bytes already reopened as the candidate.
            verify_file_bytes(path, hash, role)
        } else {
            verify_workbook_file(path, hash, role)
        }
    })
}

/// Same snapshot/no-overwrite/durability protocol for separately validated
/// document containers. Callers must validate the actual container at each
/// stage; workbook publication keeps its original fixed validator above.
pub(crate) fn create_target_from_candidate_with_validator(
    candidate: &Path,
    target: &Path,
    validate: impl Fn(&Path, &str, &str) -> Result<()>,
) -> Result<TargetCommitReport> {
    let started = Instant::now();
    ensure_distinct_paths(candidate, target)?;
    ensure_regular_nonsymlink(candidate, "candidate")?;
    ensure_target_absent(target)?;
    let candidate_hash = sha256_file(candidate)?;
    validate(candidate, &candidate_hash, "candidate")
        .context("validate candidate workbook before publication")?;
    let parent = target_parent(target);
    fs::create_dir_all(parent)
        .with_context(|| format!("create target directory {}", parent.display()))?;
    ensure_target_absent(target)?;

    let staging = copy_to_sibling(candidate, target, "stage", &candidate_hash)?;
    if let Err(error) = publish_new_file(&staging, target) {
        let _ = fs::remove_file(&staging);
        return Err(error).with_context(|| {
            format!(
                "publish new workbook target {} without overwrite",
                target.display()
            )
        });
    }
    let cleanup_warning = remove_sibling_warning(&staging);
    if let Err(error) = validate(target, &candidate_hash, "published target") {
        if sha256_file(target).ok().as_deref() == Some(candidate_hash.as_str()) {
            let _ = fs::remove_file(target);
            let _ = sync_parent(target);
        }
        return Err(error).context("validate newly published workbook target");
    }
    if let Err(error) = sync_parent(target) {
        let rollback = match sha256_file(target) {
            Ok(hash) if hash == candidate_hash => {
                fs::remove_file(target).and_then(|()| sync_parent(target).map_err(io::Error::other))
            }
            Ok(_) => Err(io::Error::other(
                "published target changed before durability rollback",
            )),
            Err(error) => Err(io::Error::other(error.to_string())),
        };
        return match rollback {
            Ok(()) => Err(error).context("target publication was rolled back after directory sync failed"),
            Err(rollback_error) => Err(error).context(format!(
                "directory sync failed and target rollback also failed ({rollback_error}); inspect {}",
                target.display()
            )),
        };
    }
    Ok(TargetCommitReport {
        operation: "created".to_owned(),
        target: target.to_path_buf(),
        sha256_before: None,
        sha256_after: candidate_hash,
        backup_path: None,
        cleanup_warning,
        elapsed_ms: started.elapsed().as_millis(),
        conflict: None,
    })
}

/// Copy a validated workbook to a new path without overwriting an existing target.
/// The copy is first staged beside the target, reopened, and published atomically.
pub fn copy_workbook_to_new_target(source: &Path, target: &Path) -> Result<TargetCommitReport> {
    let started = Instant::now();
    ensure_regular_nonsymlink(source, "source")?;
    ensure_target_absent(target)?;
    let source_hash = sha256_file(source)?;
    verify_workbook_file(source, &source_hash, "source")?;
    fs::create_dir_all(target_parent(target)).with_context(|| {
        format!(
            "create target directory {}",
            target_parent(target).display()
        )
    })?;
    ensure_target_absent(target)?;

    let candidate = copy_to_sibling(source, target, "copy-candidate", &source_hash)?;
    let commit = create_target_from_candidate(&candidate, target);
    let cleanup_warning = remove_sibling_warning(&candidate);
    match commit {
        Ok(mut report) => {
            report.operation = "copied".to_owned();
            report.elapsed_ms = started.elapsed().as_millis();
            if let Some(warning) = cleanup_warning {
                report.cleanup_warning = Some(match report.cleanup_warning {
                    Some(existing) => format!("{existing}; {warning}"),
                    None => warning,
                });
            }
            Ok(report)
        }
        Err(error) => match cleanup_warning {
            Some(warning) => {
                Err(error.context(format!("copy candidate cleanup also failed: {warning}")))
            }
            None => Err(error),
        },
    }
}

/// Replace an existing target only if it still has the hash captured when its
/// append plan was built. A verified sibling backup is kept until the
/// replacement has been reopened successfully.
pub fn replace_target_from_candidate(
    candidate: &Path,
    target: &Path,
    expected_sha256: &str,
) -> Result<TargetCommitReport> {
    let started = Instant::now();
    ensure_distinct_paths(candidate, target)?;
    ensure_regular_nonsymlink(candidate, "candidate")?;
    ensure_regular_nonsymlink(target, "existing target")?;

    let original_hash = sha256_file(target)?;
    check_replacement_snapshot(expected_sha256, &original_hash)?;
    let candidate_hash = sha256_file(candidate)?;
    verify_workbook_file(candidate, &candidate_hash, "candidate")
        .context("validate candidate workbook before replacement")?;
    let staging = copy_to_sibling(candidate, target, "stage", &candidate_hash)?;
    let permissions = fs::metadata(target)
        .with_context(|| format!("read permissions for {}", target.display()))?
        .permissions();
    if let Err(error) = fs::set_permissions(&staging, permissions.clone()) {
        let _ = fs::remove_file(&staging);
        return Err(error)
            .with_context(|| format!("preserve target permissions on {}", staging.display()));
    }

    let backup = match copy_to_sibling(target, target, "backup", &original_hash) {
        Ok(path) => path,
        Err(error) => {
            let _ = fs::remove_file(&staging);
            if let Ok(current) = sha256_file(target) {
                check_replacement_snapshot(&original_hash, &current)?;
            }
            return Err(error).context("create verified backup before target replacement");
        }
    };
    if let Err(error) = fs::set_permissions(&backup, permissions) {
        let _ = fs::remove_file(&staging);
        let _ = fs::remove_file(&backup);
        return Err(error).with_context(|| {
            format!("preserve target permissions on backup {}", backup.display())
        });
    }
    let current_hash = match sha256_file(target) {
        Ok(hash) => hash,
        Err(error) => {
            let _ = fs::remove_file(&staging);
            let _ = fs::remove_file(&backup);
            return Err(error).context("recheck configured target before replacement");
        }
    };
    if current_hash != original_hash {
        let _ = fs::remove_file(&staging);
        let _ = fs::remove_file(&backup);
        check_replacement_snapshot(&original_hash, &current_hash)?;
    }

    if let Err(error) = replace_file(&staging, target) {
        let _ = fs::remove_file(&staging);
        let _ = fs::remove_file(&backup);
        return Err(error).with_context(|| format!("replace target {}", target.display()));
    }
    let committed_hash = match verify_file_bytes(target, &candidate_hash, "published target") {
        Ok(()) => candidate_hash.clone(),
        Err(error) => {
            return match rollback_if_candidate(target, &backup, &candidate_hash, &original_hash) {
                Ok(()) => Err(error).context("replacement validation failed; original target restored"),
                Err(rollback_error) => Err(error).context(format!(
                    "replacement validation and rollback failed ({rollback_error}); inspect backup {}",
                    backup.display()
                )),
            };
        }
    };
    if let Err(error) = sync_parent(target) {
        return match rollback_if_candidate(target, &backup, &candidate_hash, &original_hash) {
            Ok(()) => Err(error).context("directory sync failed; original target restored"),
            Err(rollback_error) => Err(error).context(format!(
                "directory sync and rollback failed ({rollback_error}); inspect backup {}",
                backup.display()
            )),
        };
    }
    let (backup_path, cleanup_warning) = match fs::remove_file(&backup) {
        Ok(()) => (None, None),
        Err(error) => (
            Some(backup.clone()),
            Some(format!("could not remove verified backup: {error}")),
        ),
    };
    let cleanup_warning = match sync_parent(target) {
        Ok(()) => cleanup_warning,
        Err(error) => Some(match cleanup_warning {
            Some(previous) => format!("{previous}; final directory sync failed: {error:#}"),
            None => format!("final directory sync failed after verified commit: {error:#}"),
        }),
    };
    Ok(TargetCommitReport {
        operation: "replaced".to_owned(),
        target: target.to_path_buf(),
        sha256_before: Some(original_hash),
        sha256_after: committed_hash,
        backup_path,
        cleanup_warning,
        elapsed_ms: started.elapsed().as_millis(),
        conflict: None,
    })
}

/// Publish an already business-validated result. A changed destination is never
/// overwritten: save a separately named result and report its actual path.
/// Invalid candidates, permission failures and arbitrary IO errors still fail.
/// Internal working copies must continue using the strict replacement primitive.
pub fn commit_validated_candidate(
    candidate: &Path,
    target: &Path,
    expected_sha256: &str,
) -> Result<TargetCommitReport> {
    let started = Instant::now();
    match replace_target_from_candidate(candidate, target, expected_sha256) {
        Ok(report) => Ok(report),
        Err(error) => {
            let Some(conflict) = error.downcast_ref::<TargetSnapshotConflict>() else {
                return Err(error);
            };
            let expected = conflict.expected.clone();
            let observed = conflict.observed.clone();
            publish_conflict_copy(candidate, target, None, &expected, &observed, started)
        }
    }
}

/// Publish to a new path from the validated snapshot of a separate source.
/// If the source was edited during processing, retain a conflict copy instead.
pub fn create_target_from_source_snapshot(
    candidate: &Path,
    target: &Path,
    source: &Path,
    expected_source_hash: &str,
) -> Result<TargetCommitReport> {
    let started = Instant::now();
    ensure_distinct_paths(source, target)?;
    ensure_regular_nonsymlink(source, "source")?;
    let observed = sha256_file(source)?;
    if observed == expected_source_hash {
        return create_target_from_candidate(candidate, target);
    }
    publish_conflict_copy(
        candidate,
        target,
        Some(source),
        expected_source_hash,
        &observed,
        started,
    )
}

fn conflict_copy_path(target: &Path, sequence: u64) -> PathBuf {
    result_copy_path(target, "处理结果_源文件已变更", sequence)
}

pub(crate) fn next_result_copy_path(target: &Path, label: &str) -> PathBuf {
    result_copy_path(
        target,
        label,
        SIBLING_SEQUENCE.fetch_add(1, Ordering::Relaxed),
    )
}

fn result_copy_path(target: &Path, label: &str, sequence: u64) -> PathBuf {
    // Bound both UTF-8 bytes (Linux) and UTF-16 units (Windows), without splitting
    // a character. A long original basename must not make recovery impossible.
    let raw = target.file_stem().unwrap_or_default().to_string_lossy();
    let mut stem = String::new();
    for c in raw.chars() {
        if stem.len() + c.len_utf8() > 100 || stem.encode_utf16().count() + c.len_utf16() > 70 {
            break;
        }
        stem.push(c);
    }
    if stem.is_empty() {
        stem.push_str("workbook");
    }
    let extension = target.extension().unwrap_or_default().to_string_lossy();
    let stamp = chrono::Local::now().format("%Y%m%d_%H%M%S_%3f");
    target_parent(target).join(format!("{stem}_{label}_{stamp}_{sequence}.{extension}"))
}

fn publish_conflict_copy(
    candidate: &Path,
    target: &Path,
    changed_source: Option<&Path>,
    expected: &str,
    observed: &str,
    started: Instant,
) -> Result<TargetCommitReport> {
    for _ in 0..256 {
        let copy = conflict_copy_path(target, SIBLING_SEQUENCE.fetch_add(1, Ordering::Relaxed));
        match fs::symlink_metadata(&copy) {
            Ok(_) => continue,
            Err(error) if error.kind() == io::ErrorKind::NotFound => {}
            Err(error) => return Err(error).context("inspect conflict-copy destination"),
        }
        let mut report =
            match create_target_from_candidate(candidate, &copy) {
                Ok(report) => report,
                Err(error)
                    if error
                        .root_cause()
                        .downcast_ref::<io::Error>()
                        .is_some_and(|e| e.kind() == io::ErrorKind::AlreadyExists) =>
                {
                    continue
                }
                Err(error) => return Err(error).context(
                    "save validated result after target conflict; current target was not replaced",
                ),
            };
        report.operation = "conflict_copy_created".to_owned();
        let record_path = copy.with_file_name(format!(
            "{}.run.json",
            copy.file_name().unwrap().to_string_lossy()
        ));
        report.conflict = Some(TargetConflict {
            requested_target: target.to_path_buf(),
            changed_source: changed_source.map(Path::to_path_buf),
            expected_sha256: expected.to_owned(),
            observed_sha256: observed.to_owned(),
            reason: if changed_source.is_some() {
                "source_changed_after_planning; result_uses_processing_snapshot; new_edits_not_merged"
            } else {
                "target_changed_after_planning; result_uses_processing_snapshot; new_edits_not_merged"
            }.to_owned(),
            record_path: Some(record_path.clone()),
        });
        report.elapsed_ms = started.elapsed().as_millis();
        let record = (|| -> Result<()> {
            let mut file = OpenOptions::new()
                .write(true)
                .create_new(true)
                .open(&record_path)?;
            serde_json::to_writer_pretty(
                &mut file,
                &serde_json::json!({
                    "schema_version": 1,
                    "timestamp": chrono::Local::now().to_rfc3339(),
                    "event_name": if changed_source.is_some() { "source_changed_result_saved" } else { "target_changed_result_saved" },
                    "original_target_replaced": false,
                    "new_edits_merged": false,
                    "commit": &report,
                    "message": report.conflict_notice(),
                }),
            )?;
            file.flush()?;
            file.sync_all()?;
            sync_parent(&record_path)
        })();
        if let Err(error) = record {
            report.conflict.as_mut().unwrap().record_path = None;
            let warning = format!("结果副本已保存，但独立运行记录写入失败：{error:#}");
            report.cleanup_warning = Some(match report.cleanup_warning.take() {
                Some(previous) => format!("{previous}; {warning}"),
                None => warning,
            });
        }
        if let Some(notice) = report.conflict_notice() {
            eprintln!("{notice}");
        }
        if let Some(warning) = &report.cleanup_warning {
            eprintln!("{warning}");
        }
        return Ok(report);
    }
    bail!(
        "could not reserve a unique conflict-copy path beside {}; original target was not replaced",
        target.display()
    )
}

fn verify_workbook_file(path: &Path, expected_hash: &str, role: &str) -> Result<()> {
    verify_file_bytes(path, expected_hash, role)?;
    verify_readable_workbook(path)
        .with_context(|| format!("open {role} workbook {}", path.display()))?;
    Ok(())
}

fn verify_file_bytes(path: &Path, expected_hash: &str, role: &str) -> Result<()> {
    ensure_regular_nonsymlink(path, role)?;
    let actual_hash = sha256_file(path)?;
    ensure!(
        actual_hash == expected_hash,
        "{role} bytes do not match the expected SHA-256: {}",
        path.display()
    );
    Ok(())
}

fn rollback_if_candidate(
    target: &Path,
    backup: &Path,
    candidate_hash: &str,
    original_hash: &str,
) -> Result<()> {
    let current_hash = sha256_file(target)
        .with_context(|| format!("inspect target before rollback {}", target.display()))?;
    ensure!(
        current_hash == candidate_hash,
        "target changed after replacement; original backup retained at {}",
        backup.display()
    );
    replace_file(backup, target).with_context(|| {
        format!(
            "restore original target from backup {}; backup retained",
            backup.display()
        )
    })?;
    ensure!(
        sha256_file(target)? == original_hash,
        "restored target hash differs from the original; inspect {}",
        target.display()
    );
    sync_parent(target)
}

fn copy_to_sibling(
    source: &Path,
    target: &Path,
    role: &str,
    expected_hash: &str,
) -> Result<PathBuf> {
    let extension = source
        .extension()
        .and_then(|value| value.to_str())
        .unwrap_or("xlsx");
    copy_to_sibling_as(source, target, role, expected_hash, extension)
}

pub struct InputWorkbookSnapshot {
    path: PathBuf,
    _review_scope: crate::review_output::ReviewOriginScope,
}

impl InputWorkbookSnapshot {
    pub fn path(&self) -> &Path {
        &self.path
    }
}

impl Drop for InputWorkbookSnapshot {
    fn drop(&mut self) {
        let _ = fs::remove_file(&self.path);
    }
}

pub fn snapshot_workbook(source: &Path, expected_hash: &str) -> Result<InputWorkbookSnapshot> {
    ensure_regular_nonsymlink(source, "processing source")?;
    let extension = source
        .extension()
        .and_then(|v| v.to_str())
        .unwrap_or("xlsx");
    let placeholder = source.with_file_name("rusttools-snapshot.xlsx");
    copy_to_sibling_as(source, &placeholder, "input", expected_hash, extension).map(|path| {
        InputWorkbookSnapshot {
            path,
            _review_scope: crate::review_output::ReviewOriginScope::enter(source, expected_hash),
        }
    })
}

fn copy_to_sibling_as(
    source: &Path,
    target: &Path,
    role: &str,
    expected_hash: &str,
    extension: &str,
) -> Result<PathBuf> {
    let parent = target_parent(target);
    let stem = target.file_stem().unwrap_or_default().to_string_lossy();
    for _ in 0..256 {
        let sequence = SIBLING_SEQUENCE.fetch_add(1, Ordering::Relaxed);
        let path = parent.join(format!(
            ".{stem}.pytools-{role}-{}-{sequence}.{extension}",
            std::process::id()
        ));
        let mut destination = match OpenOptions::new().write(true).create_new(true).open(&path) {
            Ok(file) => file,
            Err(error) if error.kind() == io::ErrorKind::AlreadyExists => continue,
            Err(error) => return Err(error).with_context(|| format!("reserve {}", path.display())),
        };
        let result = (|| -> Result<()> {
            let mut input = File::open(source)
                .with_context(|| format!("open candidate/source {}", source.display()))?;
            io::copy(&mut input, &mut destination)
                .with_context(|| format!("copy {} to {}", source.display(), path.display()))?;
            destination
                .flush()
                .with_context(|| format!("flush {}", path.display()))?;
            destination
                .sync_all()
                .with_context(|| format!("sync {}", path.display()))?;
            Ok(())
        })();
        drop(destination);
        if let Err(error) = result {
            let _ = fs::remove_file(&path);
            return Err(error);
        }
        let copied_hash = match sha256_file(&path) {
            Ok(hash) => hash,
            Err(error) => {
                let _ = fs::remove_file(&path);
                return Err(error);
            }
        };
        if copied_hash != expected_hash {
            let _ = fs::remove_file(&path);
            bail!(
                "copy verification failed for {}; expected SHA-256 {expected_hash}, found {copied_hash}",
                path.display()
            );
        }
        return Ok(path);
    }
    bail!(
        "could not reserve a unique {role} file beside {}",
        target.display()
    )
}

fn ensure_regular_nonsymlink(path: &Path, role: &str) -> Result<()> {
    let metadata =
        fs::symlink_metadata(path).with_context(|| format!("inspect {role} {}", path.display()))?;
    ensure!(
        !metadata.file_type().is_symlink() && metadata.is_file(),
        "{role} must be a regular, non-symlink file: {}",
        path.display()
    );
    Ok(())
}

fn ensure_target_absent(target: &Path) -> Result<()> {
    match fs::symlink_metadata(target) {
        Ok(_) => Err(io::Error::new(
            io::ErrorKind::AlreadyExists,
            format!(
                "target already exists; refusing to overwrite {}",
                target.display()
            ),
        )
        .into()),
        Err(error) if error.kind() == io::ErrorKind::NotFound => Ok(()),
        Err(error) => Err(error).with_context(|| format!("inspect target {}", target.display())),
    }
}

fn ensure_distinct_paths(candidate: &Path, target: &Path) -> Result<()> {
    ensure!(
        path_identity(candidate)? != path_identity(target)?,
        "candidate and configured target resolve to the same path"
    );
    Ok(())
}

fn path_identity(path: &Path) -> Result<PathBuf> {
    if path.exists() {
        return fs::canonicalize(path).with_context(|| format!("canonicalize {}", path.display()));
    }
    let parent = target_parent(path);
    let file_name = path
        .file_name()
        .context("target path has no final component")?;
    let canonical_parent = if parent.exists() {
        fs::canonicalize(parent).with_context(|| format!("canonicalize {}", parent.display()))?
    } else {
        parent.to_path_buf()
    };
    Ok(canonical_parent.join(file_name))
}

fn target_parent(target: &Path) -> &Path {
    target
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."))
}

fn remove_sibling_warning(path: &Path) -> Option<String> {
    match fs::remove_file(path) {
        Ok(()) => None,
        Err(error) if error.kind() == io::ErrorKind::NotFound => None,
        Err(error) => Some(format!(
            "could not remove temporary file {}: {error}",
            path.display()
        )),
    }
}

#[cfg(unix)]
fn publish_new_file(source: &Path, target: &Path) -> Result<()> {
    fs::hard_link(source, target).with_context(|| {
        format!(
            "publish new workbook {} without overwriting an existing path",
            target.display()
        )
    })
}

#[cfg(windows)]
fn publish_new_file(source: &Path, target: &Path) -> Result<()> {
    use std::os::windows::ffi::OsStrExt;
    use windows_sys::Win32::Storage::FileSystem::{MoveFileExW, MOVEFILE_WRITE_THROUGH};

    let source = excel_editor::path_for_external_io(source)?
        .as_os_str()
        .encode_wide()
        .chain(Some(0))
        .collect::<Vec<_>>();
    let target = excel_editor::path_for_external_io(target)?
        .as_os_str()
        .encode_wide()
        .chain(Some(0))
        .collect::<Vec<_>>();
    let result = unsafe { MoveFileExW(source.as_ptr(), target.as_ptr(), MOVEFILE_WRITE_THROUGH) };
    if result == 0 {
        Err(io::Error::last_os_error()).context("MoveFileExW publish new workbook")
    } else {
        Ok(())
    }
}

#[cfg(unix)]
fn replace_file(source: &Path, target: &Path) -> Result<()> {
    fs::rename(source, target).with_context(|| format!("atomically replace {}", target.display()))
}

#[cfg(windows)]
fn replace_file(source: &Path, target: &Path) -> Result<()> {
    use std::{io, os::windows::ffi::OsStrExt};
    use windows_sys::Win32::Storage::FileSystem::{
        MoveFileExW, MOVEFILE_REPLACE_EXISTING, MOVEFILE_WRITE_THROUGH,
    };

    let source = excel_editor::path_for_external_io(source)?
        .as_os_str()
        .encode_wide()
        .chain(Some(0))
        .collect::<Vec<_>>();
    let target = excel_editor::path_for_external_io(target)?
        .as_os_str()
        .encode_wide()
        .chain(Some(0))
        .collect::<Vec<_>>();
    let flags = MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH;
    let result = unsafe { MoveFileExW(source.as_ptr(), target.as_ptr(), flags) };
    if result == 0 {
        Err(io::Error::last_os_error()).context("MoveFileExW replace target")
    } else {
        Ok(())
    }
}

#[cfg(not(any(unix, windows)))]
fn publish_new_file(_source: &Path, _target: &Path) -> Result<()> {
    bail!("safe no-clobber workbook publication is unsupported on this platform")
}

#[cfg(not(any(unix, windows)))]
fn replace_file(_source: &Path, _target: &Path) -> Result<()> {
    bail!("safe workbook replacement is unsupported on this platform")
}

#[cfg(unix)]
fn sync_parent(target: &Path) -> Result<()> {
    File::open(target_parent(target))
        .and_then(|directory| directory.sync_all())
        .with_context(|| format!("sync directory entry for {}", target.display()))
}

#[cfg(not(unix))]
fn sync_parent(_target: &Path) -> Result<()> {
    Ok(())
}
