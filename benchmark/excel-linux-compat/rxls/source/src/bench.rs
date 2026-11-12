//! Batch performance benchmark: parent schedules engine × round × size runs,
//! each executed by a fresh child process (`bench-run`) so memory/CPU stats
//! stay per-run and engine order rotates across rounds.

use crate::common::{Engine, Validate};
use anyhow::{Context, Result};
use serde::{Deserialize, Serialize};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::time::Instant;

#[derive(Debug, Clone, Serialize)]
struct ScheduledRun {
    round: u32,
    warmup: bool,
    engine: Engine,
    size: usize,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct ChildReport {
    engine: String,
    size: usize,
    round: u32,
    warmup: bool,
    init_ms: f64,
    conv_total_ms: f64,
    #[serde(default)]
    excel_session_ms: f64,
    ok: usize,
    failed: usize,
    #[serde(default)]
    peak_ws_bytes: u64,
    #[serde(default)]
    cpu_ms: f64,
    #[serde(default)]
    excel_peak_ws_bytes: Option<u64>,
    #[serde(default)]
    excel_cpu_ms: Option<f64>,
    #[serde(default)]
    input_total_bytes: u64,
    #[serde(default)]
    output_total_bytes: u64,
    #[serde(default)]
    files: Vec<serde_json::Value>,
}

#[derive(Debug, Serialize)]
struct SummaryRow {
    engine: String,
    size: usize,
    rounds: usize,
    median_conv_ms: f64,
    p95_conv_ms: f64,
    min_conv_ms: f64,
    max_conv_ms: f64,
    median_e2e_ms: f64,
    /// COM batch mode only: median of summed per-book open+save+close.
    #[serde(skip_serializing_if = "Option::is_none")]
    median_books_ms: Option<f64>,
    files_per_sec_median: f64,
    median_peak_ws_mb: f64,
    median_cpu_s: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    median_excel_peak_ws_mb: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    median_excel_cpu_s: Option<f64>,
    ok_total: usize,
    failed_total: usize,
    input_mb: f64,
    output_mb: f64,
}

pub fn run(
    fixtures_dir: &Path,
    sizes: Vec<usize>,
    rounds: u32,
    warmup: u32,
    engines: Vec<Engine>,
    out_dir: &Path,
    single_file: Option<&Path>,
    batches_dir: Option<&Path>,
) -> Result<()> {
    // Per-size file lists. With --batches-dir the canonical balanced subsets
    // (B10/B50/B100 from the existing dataset manifest) are used and their
    // SHA-256 entries are verified; otherwise the first N files in sorted
    // order are taken.
    let mut files_for_size: std::collections::BTreeMap<usize, Vec<PathBuf>> =
        std::collections::BTreeMap::new();
    match single_file {
        Some(path) => {
            for size in &sizes {
                files_for_size.insert(*size, vec![path.to_path_buf()]);
            }
        }
        None => match batches_dir {
            Some(batches_dir) => {
                for size in &sizes {
                    let batch_path = batches_dir.join(format!("B{size}.json"));
                    let raw = std::fs::read_to_string(&batch_path)
                        .with_context(|| format!("read batch {}", batch_path.display()))?;
                    #[derive(serde::Deserialize)]
                    struct BatchFile {
                        name: String,
                        sha256: String,
                    }
                    #[derive(serde::Deserialize)]
                    struct Batch {
                        files: Vec<BatchFile>,
                    }
                    let batch: Batch = serde_json::from_str(&raw)
                        .with_context(|| format!("parse batch {}", batch_path.display()))?;
                    let mut list = Vec::with_capacity(batch.files.len());
                    for entry in &batch.files {
                        let path = fixtures_dir.join(&entry.name);
                        let bytes = std::fs::read(&path)
                            .with_context(|| format!("read fixture {}", path.display()))?;
                        let digest = crate::common::sha256_hex(&bytes);
                        anyhow::ensure!(
                            digest == entry.sha256.to_lowercase(),
                            "fixture {} does not match batch sha256 ({} != {})",
                            path.display(),
                            digest,
                            entry.sha256
                        );
                        list.push(path);
                    }
                    anyhow::ensure!(
                        list.len() == *size,
                        "batch B{size} declares {} files, expected {size}",
                        list.len()
                    );
                    files_for_size.insert(*size, list);
                }
            }
            None => {
                let mut files: Vec<PathBuf> = Vec::new();
                for entry in std::fs::read_dir(fixtures_dir)? {
                    let path = entry?.path();
                    let is_xls = path
                        .extension()
                        .map(|e| e.eq_ignore_ascii_case("xls"))
                        .unwrap_or(false);
                    if path.is_file() && is_xls {
                        files.push(path);
                    }
                }
                files.sort();
                anyhow::ensure!(
                    !files.is_empty(),
                    "no .xls fixtures under {}",
                    fixtures_dir.display()
                );
                for size in &sizes {
                    if *size <= files.len() {
                        files_for_size.insert(*size, files[..*size].to_vec());
                    }
                }
            }
        },
    }

    std::fs::create_dir_all(out_dir)?;
    let mut manifest_paths: std::collections::BTreeMap<usize, PathBuf> =
        std::collections::BTreeMap::new();
    for (size, list) in &files_for_size {
        let manifest_path = out_dir.join(format!("files-B{size}.json"));
        let manifest_json = serde_json::to_string(
            &list
                .iter()
                .map(|p| p.to_string_lossy().into_owned())
                .collect::<Vec<_>>(),
        )?;
        std::fs::write(&manifest_path, manifest_json)?;
        manifest_paths.insert(*size, manifest_path);
    }
    let raw_path = out_dir.join("raw.jsonl");
    let summary_path = out_dir.join("summary.json");
    let csv_path = out_dir.join("summary.csv");
    let exe = std::env::current_exe()?;

    let mut schedule = Vec::new();
    let total_rounds = warmup + rounds;
    for round in 0..total_rounds {
        let warmup = round < warmup;
        let mut rotation = engines.clone();
        for shift in 0..(round % engines.len() as u32) {
            rotation.rotate_left(1);
            let _ = shift;
        }
        for engine in &rotation {
            for size in &sizes {
                if files_for_size.contains_key(size) {
                    schedule.push(ScheduledRun {
                        round,
                        warmup,
                        engine: *engine,
                        size: *size,
                    });
                }
            }
        }
    }
    eprintln!(
        "benchmark schedule: {} runs ({} rounds incl. warmup, engines: {:?}, sizes: {:?})",
        schedule.len(),
        total_rounds,
        engines.iter().map(|e| e.id()).collect::<Vec<_>>(),
        sizes
    );

    let mut raw_file = std::fs::File::create(&raw_path)?;
    let mut child_reports: Vec<ChildReport> = Vec::new();
    for (index, scheduled) in schedule.iter().enumerate() {
        let tmp_parent = out_dir.join("tmp");
        let tmp = crate::common::TempDir::create(&tmp_parent, "bench")?;
        let report_path = tmp.0.join("report.json");
        let label = format!(
            "[{}/{}] round={} {} {} size={}",
            index + 1,
            schedule.len(),
            scheduled.round,
            if scheduled.warmup { "warmup" } else { "timed" },
            scheduled.engine.id(),
            scheduled.size
        );
        eprintln!("{label} running...");
        let mut command_args = vec![
            "bench-run".to_owned(),
            "--engine".to_owned(),
            scheduled.engine.id().to_owned(),
            "--files".to_owned(),
            manifest_paths[&scheduled.size].to_string_lossy().into_owned(),
            "--size".to_owned(),
            scheduled.size.to_string(),
            "--round".to_owned(),
            scheduled.round.to_string(),
            "--out-dir".to_owned(),
            tmp.0.to_string_lossy().into_owned(),
            "--report".to_owned(),
            report_path.to_string_lossy().into_owned(),
        ];
        if scheduled.warmup {
            command_args.push("--warmup".to_owned());
        }
        let output = Command::new(&exe)
            .args(&command_args)
            .output()
            .context("spawn bench-run child")?;
        if !output.status.success() {
            eprintln!(
                "{label} FAILED: exit {:?}\nstdout: {}\nstderr: {}",
                output.status.code(),
                String::from_utf8_lossy(&output.stdout),
                String::from_utf8_lossy(&output.stderr)
            );
            continue;
        }
        let report: ChildReport = match serde_json::from_slice(&output.stdout) {
            Ok(report) => report,
            Err(e) => {
                eprintln!("{label} report parse failed: {e}; stdout tail: {}", {
                    let text = String::from_utf8_lossy(&output.stdout);
                    text.chars().rev().take(400).collect::<String>().chars().rev().collect::<String>()
                });
                continue;
            }
        };
        writeln!(raw_file, "{}", serde_json::to_string(&report)?)?;
        child_reports.push(report);
        // TempDir drop cleans converted outputs; timings are already captured.
    }
    raw_file.flush()?;

    let mut rows: Vec<SummaryRow> = Vec::new();
    let mut keys: Vec<(String, usize)> = child_reports
        .iter()
        .filter(|r| !r.warmup)
        .map(|r| (r.engine.clone(), r.size))
        .collect();
    keys.sort();
    keys.dedup();
    for (engine, size) in &keys {
        let timed: Vec<&ChildReport> = child_reports
            .iter()
            .filter(|r| !r.warmup && &r.engine == engine && &r.size == size)
            .collect();
        if timed.is_empty() {
            continue;
        }
        let mut convs: Vec<f64> = timed.iter().map(|r| r.conv_total_ms).collect();
        let mut e2es: Vec<f64> = timed
            .iter()
            .map(|r| r.init_ms + r.conv_total_ms)
            .collect();
        let books: Option<Vec<f64>> = {
            let has_com_files = timed.iter().any(|r| {
                r.files
                    .iter()
                    .any(|f| f.get("timing").and_then(|t| t.get("com_open_ms")).is_some())
            });
            if has_com_files {
                Some(
                    timed
                        .iter()
                        .map(|r| {
                            r.files
                                .iter()
                                .filter_map(|f| f.get("timing"))
                                .map(|t| {
                                    let get = |k: &str| {
                                        t.get(k)
                                            .and_then(|v| v.as_f64())
                                            .unwrap_or(0.0)
                                    };
                                    get("com_open_ms") + get("com_save_as_ms") + get("com_close_ms")
                                })
                                .sum()
                        })
                        .collect(),
                )
            } else {
                None
            }
        };
        let peaks: Vec<f64> = timed
            .iter()
            .map(|r| r.peak_ws_bytes as f64 / 1024.0 / 1024.0)
            .collect();
        let cpus: Vec<f64> = timed.iter().map(|r| r.cpu_ms / 1000.0).collect();
        let excel_peaks: Vec<f64> = timed
            .iter()
            .filter_map(|r| r.excel_peak_ws_bytes.map(|b| b as f64 / 1024.0 / 1024.0))
            .collect();
        let excel_cpus: Vec<f64> = timed
            .iter()
            .filter_map(|r| r.excel_cpu_ms.map(|m| m / 1000.0))
            .collect();
        let median = |values: &mut Vec<f64>| -> f64 {
            values.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
            if values.is_empty() {
                0.0
            } else {
                values[values.len() / 2]
            }
        };
        let p95 = |values: &mut Vec<f64>| -> f64 {
            values.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
            if values.is_empty() {
                0.0
            } else {
                let index = (((values.len() as f64) * 0.95).ceil() as usize)
                    .saturating_sub(1)
                    .min(values.len() - 1);
                values[index]
            }
        };
        let median_conv = median(&mut convs.clone());
        let median_books = books.as_ref().map(|b| median(&mut b.clone()));
        rows.push(SummaryRow {
            engine: engine.clone(),
            size: *size,
            rounds: timed.len(),
            median_conv_ms: median_conv,
            p95_conv_ms: p95(&mut convs.clone()),
            min_conv_ms: convs.iter().cloned().fold(f64::INFINITY, f64::min),
            max_conv_ms: convs.iter().cloned().fold(0.0, f64::max),
            median_e2e_ms: median(&mut e2es),
            median_books_ms: median_books,
            files_per_sec_median: if median_conv > 0.0 {
                *size as f64 / (median_conv / 1000.0)
            } else {
                0.0
            },
            median_peak_ws_mb: median(&mut peaks.clone()),
            median_cpu_s: median(&mut cpus.clone()),
            median_excel_peak_ws_mb: if excel_peaks.is_empty() {
                None
            } else {
                Some(median(&mut excel_peaks.clone()))
            },
            median_excel_cpu_s: if excel_cpus.is_empty() {
                None
            } else {
                Some(median(&mut excel_cpus.clone()))
            },
            ok_total: timed.iter().map(|r| r.ok).sum(),
            failed_total: timed.iter().map(|r| r.failed).sum(),
            input_mb: timed
                .iter()
                .map(|r| r.input_total_bytes)
                .sum::<u64>() as f64
                / 1024.0
                / 1024.0
                / timed.len() as f64,
            output_mb: timed
                .iter()
                .map(|r| r.output_total_bytes)
                .sum::<u64>() as f64
                / 1024.0
                / 1024.0
                / timed.len() as f64,
        });
    }

    std::fs::write(
        &summary_path,
        serde_json::to_string_pretty(&rows)?,
    )?;
    let mut csv = std::fs::File::create(&csv_path)?;
    writeln!(csv, "engine,size,rounds,median_conv_ms,p95_conv_ms,min_conv_ms,max_conv_ms,median_e2e_ms,median_books_ms,files_per_sec_median,median_peak_ws_mb,median_cpu_s,median_excel_peak_ws_mb,median_excel_cpu_s,ok_total,failed_total,input_mb,output_mb")?;
    for row in &rows {
        writeln!(
            csv,
            "{},{},{},{:.1},{:.1},{:.1},{:.1},{:.1},{},{:.2},{:.1},{:.2},{},{},{},{},{:.2},{:.2}",
            row.engine,
            row.size,
            row.rounds,
            row.median_conv_ms,
            row.p95_conv_ms,
            row.min_conv_ms,
            row.max_conv_ms,
            row.median_e2e_ms,
            row.median_books_ms
                .map(|v| format!("{v:.1}"))
                .unwrap_or_default(),
            row.files_per_sec_median,
            row.median_peak_ws_mb,
            row.median_cpu_s,
            row.median_excel_peak_ws_mb
                .map(|v| format!("{v:.1}"))
                .unwrap_or_default(),
            row.median_excel_cpu_s
                .map(|v| format!("{v:.2}"))
                .unwrap_or_default(),
            row.ok_total,
            row.failed_total,
            row.input_mb,
            row.output_mb
        )?;
    }
    eprintln!("benchmark summary: {}", summary_path.display());
    eprintln!("benchmark csv: {}", csv_path.display());
    Ok(())
}

/// Child process entry: run one engine over `size` files once.
pub fn run_child(
    engine: Engine,
    files_path: &Path,
    size: usize,
    round: u32,
    warmup: bool,
    out_dir: &Path,
    report_path: &Path,
) -> Result<()> {
    let raw = std::fs::read_to_string(files_path)?;
    let all: Vec<String> = serde_json::from_str(&raw)?;
    let inputs: Vec<PathBuf> = all.into_iter().map(PathBuf::from).collect();
    std::fs::create_dir_all(out_dir)?;

    let t_start = Instant::now();
    let mut init_ms = 0.0;
    let mut conv_total_ms = 0.0;
    let mut excel_session_ms = 0.0;
    let mut outcomes: Vec<crate::common::FileOutcome>;
    let mut com_meta: Option<crate::common::ComMeta> = None;

    match engine {
        Engine::Rxls => {
            init_ms = t_start.elapsed().as_secs_f64() * 1000.0;
            let t_conv = Instant::now();
            outcomes = crate::bx_rxls::convert_batch(&inputs, out_dir, Validate::Quick, false)?;
            conv_total_ms = t_conv.elapsed().as_secs_f64() * 1000.0;
        }
        Engine::CalamineValues | Engine::CalamineFormulas => {
            let mode = if engine == Engine::CalamineFormulas {
                crate::bx_calamine::CalamineMode::Formulas
            } else {
                crate::bx_calamine::CalamineMode::Values
            };
            init_ms = t_start.elapsed().as_secs_f64() * 1000.0;
            let t_conv = Instant::now();
            outcomes = crate::bx_calamine::convert_batch(&inputs, out_dir, mode, Validate::Quick)?;
            conv_total_ms = t_conv.elapsed().as_secs_f64() * 1000.0;
        }
        Engine::ComBatch | Engine::ComPerFile => {
            #[cfg(windows)]
            {
                let pairs: Vec<(PathBuf, PathBuf)> = inputs
                    .iter()
                    .map(|input| {
                        let stem = input
                            .file_stem()
                            .map(|s| s.to_string_lossy().into_owned())
                            .unwrap_or_else(|| "workbook".to_owned());
                        (input.clone(), out_dir.join(format!("{stem}.xlsx")))
                    })
                    .collect();
                let t_conv = Instant::now();
                let (file_outcomes, meta) =
                    crate::bx_com::convert(&pairs, engine == Engine::ComPerFile)?;
                conv_total_ms = t_conv.elapsed().as_secs_f64() * 1000.0;
                excel_session_ms = meta.sta_init_ms
                    + meta.excel_create_config_ms
                    + meta.excel_quit_ms
                    + meta.sta_release_ms;
                com_meta = Some(meta);
                outcomes = file_outcomes;
            }
            #[cfg(not(windows))]
            {
                anyhow::bail!("COM engines require Windows");
            }
        }
    }

    let mut ok = 0usize;
    let mut failed = 0usize;
    let mut input_total = 0u64;
    let mut output_total = 0u64;
    for outcome in &outcomes {
        match outcome.status.as_str() {
            "ok" => ok += 1,
            "skipped" => {}
            _ => failed += 1,
        }
        input_total += outcome.input_bytes.unwrap_or(0);
        output_total += outcome.output_bytes.unwrap_or(0);
    }

    #[cfg(windows)]
    let (peak_ws_bytes, cpu_ms) = {
        let pid = std::process::id();
        match crate::combase::memory_and_cpu_of_pid(pid) {
            Ok((peak, cpu)) => (peak, cpu),
            Err(_) => (0, 0.0),
        }
    };
    #[cfg(not(windows))]
    let (peak_ws_bytes, cpu_ms) = (0u64, 0.0f64);

    let report = serde_json::json!({
        "engine": engine.id(),
        "size": size,
        "round": round,
        "warmup": warmup,
        "init_ms": init_ms,
        "conv_total_ms": conv_total_ms,
        "excel_session_ms": excel_session_ms,
        "ok": ok,
        "failed": failed,
        "peak_ws_bytes": peak_ws_bytes,
        "cpu_ms": cpu_ms,
        "excel_peak_ws_bytes": com_meta.as_ref().and_then(|m| m.excel_peak_ws_bytes),
        "excel_cpu_ms": com_meta.as_ref().and_then(|m| m.excel_cpu_ms),
        "input_total_bytes": input_total,
        "output_total_bytes": output_total,
        "files": outcomes,
    });
    std::fs::write(report_path, serde_json::to_string_pretty(&report)?)?;
    // Parent parses this from stdout; keep it single-line compact.
    println!("{}", serde_json::to_string(&report)?);
    Ok(())
}
