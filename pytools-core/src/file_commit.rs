//! Safe publication of validated workbook candidates to configured targets.

use anyhow::{bail, ensure, Context, Result};
use excel_reader::read_workbook;
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

#[derive(Clone, Debug, Serialize)]
pub struct TargetCommitReport {
    pub operation: String,
    pub target: PathBuf,
    pub sha256_before: Option<String>,
    pub sha256_after: String,
    pub backup_path: Option<PathBuf>,
    pub cleanup_warning: Option<String>,
    pub elapsed_ms: u128,
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
    let started = Instant::now();
    ensure_distinct_paths(candidate, target)?;
    ensure_regular_nonsymlink(candidate, "candidate")?;
    ensure_target_absent(target)?;
    let candidate_hash = sha256_file(candidate)?;
    verify_workbook_file(candidate, &candidate_hash, "candidate")
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
    if let Err(error) = verify_workbook_file(target, &candidate_hash, "published target") {
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
    ensure!(
        original_hash == expected_sha256,
        "configured target changed after planning; refusing to replace {}",
        target.display()
    );
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
        bail!(
            "configured target changed during candidate preparation; refusing to replace {}",
            target.display()
        );
    }

    if let Err(error) = replace_file(&staging, target) {
        let _ = fs::remove_file(&staging);
        let _ = fs::remove_file(&backup);
        return Err(error).with_context(|| format!("replace target {}", target.display()));
    }
    let committed_hash = match verify_workbook_file(target, &candidate_hash, "published target") {
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
    })
}

fn verify_workbook_file(path: &Path, expected_hash: &str, role: &str) -> Result<()> {
    ensure_regular_nonsymlink(path, role)?;
    let actual_hash = sha256_file(path)?;
    ensure!(
        actual_hash == expected_hash,
        "{role} bytes do not match the expected SHA-256: {}",
        path.display()
    );
    read_workbook(path).with_context(|| format!("open {role} workbook {}", path.display()))?;
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
    let parent = target_parent(target);
    let stem = target.file_stem().unwrap_or_default().to_string_lossy();
    for _ in 0..256 {
        let sequence = SIBLING_SEQUENCE.fetch_add(1, Ordering::Relaxed);
        let path = parent.join(format!(
            ".{stem}.pytools-{role}-{}-{sequence}.xlsx",
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
        Ok(_) => bail!(
            "target already exists; refusing to overwrite {}",
            target.display()
        ),
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

    let source = source
        .as_os_str()
        .encode_wide()
        .chain(Some(0))
        .collect::<Vec<_>>();
    let target = target
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

    let source = source
        .as_os_str()
        .encode_wide()
        .chain(Some(0))
        .collect::<Vec<_>>();
    let target = target
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
