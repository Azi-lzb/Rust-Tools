//! Explicit OS integration for the selected quick actions.

use anyhow::{bail, Context, Result};
use std::{
    fs,
    path::{Path, PathBuf},
    process::Command,
    time::SystemTime,
};

pub fn open_config(path: &Path) -> Result<()> {
    open_path(path)
}
pub fn open_output_directory(path: &Path) -> Result<()> {
    fs::create_dir_all(path)
        .with_context(|| format!("create output directory {}", path.display()))?;
    open_path(path)
}
pub fn open_latest_output(path: &Path) -> Result<Option<PathBuf>> {
    if !path.exists() {
        return Ok(None);
    }
    let mut files = Vec::new();
    collect_files(path, &mut files);
    files.retain(|file| {
        !file
            .file_name()
            .unwrap_or_default()
            .to_string_lossy()
            .starts_with("~$")
    });
    let latest = files
        .into_iter()
        .filter_map(|file| {
            fs::metadata(&file)
                .ok()
                .and_then(|metadata| metadata.modified().ok().map(|time| (file, time)))
        })
        .max_by_key(|(_, time)| *time)
        .map(|(file, _)| file);
    if let Some(file) = latest {
        open_path(&file)?;
        Ok(Some(file))
    } else {
        Ok(None)
    }
}

fn collect_files(path: &Path, output: &mut Vec<PathBuf>) {
    let Ok(entries) = fs::read_dir(path) else {
        return;
    };
    for entry in entries.flatten() {
        let child = entry.path();
        if child.is_dir() {
            collect_files(&child, output);
        } else if child.is_file() {
            output.push(child);
        }
    }
}
fn open_path(path: &Path) -> Result<()> {
    if !path.exists() {
        bail!("path does not exist: {}", path.display());
    }
    let status = if cfg!(windows) {
        Command::new("explorer.exe").arg(path).status()
    } else if cfg!(target_os = "macos") {
        Command::new("open").arg(path).status()
    } else {
        Command::new("xdg-open").arg(path).status()
    }
    .with_context(|| format!("start system opener for {}", path.display()))?;
    // PyCode opens files with os.startfile (ShellExecute), which never surfaces
    // an exit status. explorer.exe routinely exits with code 1 after a successful
    // hand-off to an existing process, so on Windows only codes beyond 0/1 are
    // treated as failures; existence is validated above like the source menu.
    let opener_failed = if cfg!(windows) {
        !matches!(status.code(), Some(0 | 1))
    } else {
        !status.success()
    };
    if opener_failed {
        bail!("system opener exited with {status}");
    }
    Ok(())
}
