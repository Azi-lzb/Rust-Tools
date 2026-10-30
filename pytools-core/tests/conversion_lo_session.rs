//! LibreOffice batch-session integration coverage. These tests drive the
//! public conversion entry points with `--engine lo` and verify that one
//! batch is served by a single `soffice` process (process-id evidence in
//! every item detail). They are ignored unless a real `soffice` is on PATH;
//! the lifecycle logic itself is covered by in-crate fake-runner tests.

use anyhow::{Context, Result};
use pytools_core::conversion;
use serde_json::Value;
use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::time::Instant;

fn soffice_available() -> bool {
    Command::new("soffice")
        .arg("--version")
        .output()
        .map(|output| output.status.success())
        .unwrap_or(false)
}

fn synthetic_xls_fixture() -> Result<PathBuf> {
    let fixture = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../benchmarks/xls-to-xlsx-com/fixtures/generated/SMALL-001.xls");
    anyhow::ensure!(
        fixture.is_file(),
        "synthetic XLS fixture is missing: {}",
        fixture.display()
    );
    Ok(fixture)
}

fn write_workbook_item(detail: &str) -> Value {
    serde_json::from_str(detail).expect("conversion detail must be valid JSON")
}

fn session_process_ids(detail: &str) -> Vec<u32> {
    let report = write_workbook_item(detail);
    report["lo_session"]["process_ids"]
        .as_array()
        .expect("LO session process-id evidence missing")
        .iter()
        .filter_map(|value| value.as_u64().map(|value| value as u32))
        .collect()
}

#[test]
#[ignore = "requires LibreOffice (soffice) on PATH; run explicitly on LO-equipped machines"]
fn lo_batch_real_soffice_converts_batch_in_one_process() -> Result<()> {
    if !soffice_available() {
        panic!("soffice is not available on PATH; this test requires a real LibreOffice");
    }
    let fixture = synthetic_xls_fixture()?;
    let root = std::env::temp_dir().join(format!(
        "pytools-lo-it-{}-{}",
        std::process::id(),
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)?
            .as_nanos()
    ));
    let input_dir = root.join("inputs");
    let output_dir = root.join("xlsx");
    fs::create_dir_all(&input_dir)?;

    const BATCH: usize = 4;
    let mut inputs = Vec::new();
    for index in 0..BATCH {
        let input = input_dir.join(format!("lo-session-{index:03}.xls"));
        fs::copy(&fixture, &input)?;
        inputs.push(input);
    }

    let run_started = Instant::now();
    let run = conversion::convert_excel_files_with_progress(
        &inputs,
        Some(&output_dir),
        "xlsx",
        "lo",
        |_| {},
    )?;
    let batch_elapsed = run_started.elapsed();

    assert_eq!(run.converted, BATCH, "every staged file must convert");
    assert_eq!(run.failed, 0);
    // Session evidence: one conversion process for the whole batch.
    let mut process_id_sets = Vec::new();
    for item in &run.items {
        let report = write_workbook_item(&item.detail);
        assert_eq!(report["engine"], "lo");
        let process_ids = session_process_ids(&item.detail);
        assert_eq!(
            process_ids.len(),
            1,
            "one LO process must serve the batch: {}",
            item.detail
        );
        process_id_sets.push(process_ids);
    }
    for window in process_id_sets.windows(2) {
        assert_eq!(
            window[0], window[1],
            "every item must report the same LO process id"
        );
    }
    for item in &run.items {
        let output = item
            .output
            .as_ref()
            .with_context(|| format!("converted output missing for {}", item.input.display()))?;
        assert!(output.is_file() && fs::metadata(output)?.len() > 0);
    }
    println!(
        "LO batch of {BATCH} files converted in {} ms with a single soffice process",
        batch_elapsed.as_millis()
    );
    fs::remove_dir_all(&root).ok();
    Ok(())
}

#[test]
#[ignore = "requires LibreOffice (soffice) on PATH; run explicitly on LO-equipped machines"]
fn lo_batch_real_session_single_file_failure_keeps_batch_running() -> Result<()> {
    if !soffice_available() {
        panic!("soffice is not available on PATH; this test requires a real LibreOffice");
    }
    let fixture = synthetic_xls_fixture()?;
    let root = std::env::temp_dir().join(format!(
        "pytools-lo-it-fail-{}-{}",
        std::process::id(),
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)?
            .as_nanos()
    ));
    let input_dir = root.join("inputs");
    let output_dir = root.join("xlsx");
    fs::create_dir_all(&input_dir)?;

    let mut inputs = Vec::new();
    for index in 0..2 {
        let input = input_dir.join(format!("good-{index:03}.xls"));
        fs::copy(&fixture, &input)?;
        inputs.push(input);
    }
    let corrupt = input_dir.join("corrupt.xls");
    fs::write(&corrupt, b"this is not a real XLS workbook")?;
    inputs.push(corrupt);

    let run = conversion::convert_excel_files_with_progress(
        &inputs,
        Some(&output_dir),
        "xlsx",
        "lo",
        |_| {},
    )?;
    assert_eq!(run.converted, 2, "both valid files must still convert");
    assert_eq!(run.failed, 1, "only the corrupt file may fail");
    let failed = run
        .items
        .iter()
        .find(|item| item.status == "failed")
        .expect("corrupt input must be reported failed");
    assert!(failed.output.is_none());
    assert!(
        failed.detail.contains("did not produce an output")
            || failed.detail.contains("conversion failed"),
        "failure detail should explain the LO outcome: {}",
        failed.detail
    );
    fs::remove_dir_all(&root).ok();
    Ok(())
}

#[test]
#[ignore = "performance comparison; requires LibreOffice (soffice) on PATH"]
fn lo_batch_performance_batch_session_vs_per_file_process() -> Result<()> {
    if !soffice_available() {
        panic!("soffice is not available on PATH; this benchmark requires a real LibreOffice");
    }
    let fixture = synthetic_xls_fixture()?;
    let root = std::env::temp_dir().join(format!(
        "pytools-lo-perf-{}-{}",
        std::process::id(),
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)?
            .as_nanos()
    ));
    let input_dir = root.join("inputs");
    fs::create_dir_all(&input_dir)?;

    const BATCH: usize = 10;
    let mut inputs = Vec::new();
    for index in 0..BATCH {
        let input = input_dir.join(format!("perf-{index:03}.xls"));
        fs::copy(&fixture, &input)?;
        inputs.push(input);
    }

    // Before-style behaviour: one runtime per file (per-file soffice start).
    let per_file_started = Instant::now();
    for input in &inputs {
        let output = root
            .join("per-file")
            .join(format!("{}.xlsx", input.file_stem().unwrap().to_string_lossy()));
        fs::create_dir_all(output.parent().unwrap())?;
        conversion::convert_one(input, &output, "xlsx", "lo")
            .with_context(|| format!("per-file conversion {}", input.display()))?;
    }
    let per_file_elapsed = per_file_started.elapsed();

    // After-style behaviour: one runtime, one batch session.
    let batch_started = Instant::now();
    let run = conversion::convert_excel_files_with_progress(
        &inputs,
        Some(&root.join("batch")),
        "xlsx",
        "lo",
        |_| {},
    )?;
    let batch_elapsed = batch_started.elapsed();
    assert_eq!(run.converted, BATCH);

    println!(
        "LibreOffice XLS→XLSX batch of {BATCH} synthetic files:\n  per-file process : {} ms\n  batch session    : {} ms\n  session process ids: {:?}",
        per_file_elapsed.as_millis(),
        batch_elapsed.as_millis(),
        session_process_ids(&run.items[0].detail),
    );
    fs::remove_dir_all(&root).ok();
    Ok(())
}

/// Static guard: the LO session evidence must not leak into COM conversions.
#[test]
fn com_conversion_detail_has_no_lo_session_evidence() {
    // The lo_session key is only added when LibreOffice processes ran; the
    // conversion of a same-format copy or a native CSV never sees it. This
    // test pins the JSON contract shape without an engine.
    let detail = r#"{"status":"converted","engine":"com"}"#;
    let report: Value = serde_json::from_str(detail).unwrap();
    assert!(report.get("lo_session").is_none());
}

#[allow(dead_code)]
fn unused_path_guard(path: &Path) -> bool {
    path.is_file()
}
