//! Configured source-file discovery shared by PyTools batch workflows.

use crate::config::GlobalConfig;
use anyhow::{Context, Result};
use std::{collections::HashSet, fs, path::PathBuf};

/// Match the original PyCode `list_source_files` default scan: recurse below
/// the configured input directory, accept configured extensions, skip Excel
/// lock files, and return paths in deterministic order.
pub fn discover_configured_source_files(global: &GlobalConfig) -> Result<Vec<PathBuf>> {
    if !global.input_dir.is_dir() {
        return Ok(Vec::new());
    }

    let supported_extensions = global
        .source_exts
        .iter()
        .map(|value| {
            let extension = value.trim().to_lowercase();
            if extension.starts_with('.') {
                extension
            } else {
                format!(".{extension}")
            }
        })
        .collect::<HashSet<_>>();
    let mut pending_directories = vec![global.input_dir.clone()];
    let mut sources = Vec::new();

    while let Some(directory) = pending_directories.pop() {
        for entry in fs::read_dir(&directory)
            .with_context(|| format!("scan configured input directory {}", directory.display()))?
        {
            let entry = entry.with_context(|| {
                format!(
                    "read entry in configured input directory {}",
                    directory.display()
                )
            })?;
            let path = entry.path();
            let file_type = entry
                .file_type()
                .with_context(|| format!("inspect configured input path {}", path.display()))?;

            if file_type.is_dir() {
                pending_directories.push(path);
                continue;
            }

            let file_name = path
                .file_name()
                .map(|name| name.to_string_lossy())
                .unwrap_or_default();
            if file_name.starts_with("~$") || !path.is_file() {
                continue;
            }

            let extension = path
                .extension()
                .map(|value| format!(".{}", value.to_string_lossy().to_lowercase()));
            if extension
                .as_ref()
                .is_some_and(|extension| supported_extensions.contains(extension))
            {
                sources.push(path);
            }
        }
    }

    sources.sort();
    Ok(sources)
}
