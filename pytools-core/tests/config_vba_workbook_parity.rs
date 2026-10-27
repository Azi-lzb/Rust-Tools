//! Loader parity against a real business configuration workbook.
//!
//! The fixture is the main workbook of the user's public VBA toolbox
//! (https://github.com/Azi-lzb/VBA_Excel_Tools, commit cea6dfa). Its
//! configuration sheets carry real rule data (15 timeline rules, 146 path
//! maps, 17 dedup tasks, ...). `expected_loaders.json` was produced by the
//! read-only PyCode loaders over the same workbook, so this test is a
//! field-by-field Python/Rust parity check over real configuration values.

use anyhow::{Context, Result};
use serde_json::Value;
use std::{fs, path::Path};

fn fixture(relative: &str) -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/vba_config")
        .join(relative)
}

fn expected_section(name: &str) -> Result<Value> {
    let raw = fs::read_to_string(fixture("expected_loaders.json"))
        .context("read expected_loaders.json")?;
    let value: Value = serde_json::from_str(&raw)?;
    Ok(value
        .get(name)
        .cloned()
        .with_context(|| format!("expected_loaders.json misses section {name}"))?)
}

fn actual(name: &str, value: impl serde::Serialize) -> (String, Value) {
    (
        name.to_owned(),
        serde_json::to_value(value).expect("serialize loaders"),
    )
}

#[test]
fn real_vba_workbook_loader_parity() -> Result<()> {
    let workbook = fixture("VbaV2.4.2.xlsm");

    let sections = vec![
        actual(
            "timeline_rules",
            pytools_core::config::load_timeline_rules(&workbook)?,
        ),
        actual(
            "path_maps",
            pytools_core::config::load_path_maps(&workbook)?,
        ),
        actual(
            "dedup_tasks",
            pytools_core::config::load_dedup_tasks(&workbook)?,
        ),
        actual(
            "print_tasks",
            pytools_core::config::load_print_tasks(&workbook)?,
        ),
        actual(
            "extract_tasks",
            pytools_core::config::load_extract_tasks(&workbook)?,
        ),
        actual(
            "foreign_banks",
            pytools_core::config::load_foreign_banks(&workbook)?,
        ),
        actual(
            "institution_mapping",
            pytools_core::config::load_institution_mapping(&workbook)?,
        ),
    ];

    for (name, actual_value) in sections {
        let expected = expected_section(&name)?;
        let expected_count = match &expected {
            Value::Array(items) => Some(items.len()),
            Value::Object(map) => Some(map.len()),
            _ => None,
        };
        assert_eq!(
            expected, actual_value,
            "loader parity failed for {name} against the real VBA workbook (expected {:?} entries)",
            expected_count
        );
    }
    Ok(())
}
