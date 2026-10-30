#![cfg(windows)]

use anyhow::{Context, Result};
use pytools_core::conversion::convert_excel_files_with_progress;
use serde_json::Value;
use std::{
    fs,
    path::{Path, PathBuf},
    time::{SystemTime, UNIX_EPOCH},
};

struct TempDir(PathBuf);
impl TempDir {
    fn new() -> Result<Self> {
        let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let path = std::env::temp_dir().join(format!(
            "rust-tools-com-integration-{}-{nonce}",
            std::process::id()
        ));
        fs::create_dir(&path)?;
        Ok(Self(path))
    }
    fn path(&self) -> &Path {
        &self.0
    }
}
impl Drop for TempDir {
    fn drop(&mut self) {
        let parent = std::env::temp_dir().canonicalize();
        let target = self.0.canonicalize();
        if let (Ok(parent), Ok(target)) = (parent, target) {
            let is_ours = target.parent() == Some(parent.as_path())
                && target.file_name().is_some_and(|name| {
                    name.to_string_lossy()
                        .starts_with("rust-tools-com-integration-")
                });
            if is_ours {
                let _ = fs::remove_dir_all(target);
            }
        }
    }
}

#[test]
#[ignore = "requires locally installed Excel and WPS COM; uses temporary synthetic workbooks"]
fn spreadsheet_com_reuses_session_preserves_sources_and_recovers_after_bad_et() -> Result<()> {
    let root = TempDir::new()?;
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/timeline/wide_summary_elapsed_duration.xls");
    let first = root.path().join("first.xls");
    let bad_et = root.path().join("not-native.et");
    let second = root.path().join("second.xls");
    for path in [&first, &bad_et, &second] {
        fs::copy(&fixture, path)?;
    }
    let before = [&first, &bad_et, &second]
        .iter()
        .map(|path| fs::read(path))
        .collect::<std::io::Result<Vec<_>>>()?;
    let mut events = Vec::new();
    let output = root.path().join("output");
    let run = convert_excel_files_with_progress(
        &[first.clone(), bad_et.clone(), second.clone()],
        Some(&output),
        "xlsx",
        "com",
        |event| events.push(event.clone()),
    )?;
    assert_eq!(run.converted, 2, "{:?}", run.items);
    assert_eq!(run.failed, 1, "{:?}", run.items);
    assert_eq!(run.items[1].status, "failed");
    assert!(
        run.items[1].detail.contains("原生 ET") || run.items[1].detail.contains("FileFormat"),
        "{}",
        run.items[1].detail
    );
    let first_detail: Value = serde_json::from_str(&run.items[0].detail)?;
    let second_detail: Value = serde_json::from_str(&run.items[2].detail)?;
    assert_eq!(first_detail["engine"], "com");
    assert_eq!(first_detail["com_bridge"], "rust");
    assert_eq!(first_detail["com_application"], "Excel.Application");
    let first_pid = first_detail["com_process_id"]
        .as_u64()
        .context("first COM PID missing")?;
    assert_eq!(second_detail["com_process_id"].as_u64(), Some(first_pid));
    assert!(run.items[0]
        .output
        .as_ref()
        .is_some_and(|path| path.is_file()));
    assert!(run.items[2]
        .output
        .as_ref()
        .is_some_and(|path| path.is_file()));
    assert!(!output.join("not-native.xlsx").exists());
    assert_eq!(fs::read_dir(&output)?.count(), 2);
    assert!(events.iter().any(|event| event.status == "failed"));
    assert!(events
        .iter()
        .any(|event| event.actual_engine.as_deref() == Some("Microsoft Excel COM (Rust)")));
    assert_eq!(
        [&first, &bad_et, &second]
            .iter()
            .map(|path| fs::read(path))
            .collect::<std::io::Result<Vec<_>>>()?,
        before
    );
    Ok(())
}
