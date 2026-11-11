//! Standalone XLS→XLSX conversion feasibility benchmark.
//!
//! Backends under test:
//! - A  `com-batch`   Excel COM, one instance per batch (production-like)
//! - A' `com-perfile` Excel COM, one instance per file
//! - B  `rxls`        rXLS direct conversion (no Office runtime)
//! - C  `calamine-values`    calamine + rust_xlsxwriter, values only
//! - C  `calamine-formulas`  calamine + rust_xlsxwriter, values + formulas
//!
//! This crate is excluded from the pytools workspace and does not touch any
//! production conversion code.

mod bench;
mod bx_calamine;
mod bx_rxls;
mod common;
mod compare;
mod mkcorrupt;
#[cfg(windows)]
mod bx_com;
#[cfg(windows)]
mod combase;
#[cfg(windows)]
mod fxgen;
#[cfg(windows)]
mod opencheck;

use anyhow::{bail, Result};
use clap::{Parser, Subcommand};
use common::{Engine, Validate};
use std::path::PathBuf;

#[derive(Parser)]
#[command(
    name = "xls2x-bench",
    version,
    about = "XLS to XLSX conversion feasibility benchmark (excel_com vs rxls vs calamine+rust_xlsxwriter)"
)]
struct Cli {
    #[command(subcommand)]
    command: Command,
}

#[derive(Subcommand)]
enum Command {
    /// Convert XLS files to XLSX with one backend.
    Convert {
        /// Input .xls files or directories.
        inputs: Vec<PathBuf>,
        /// Output directory for .xlsx files.
        #[arg(short, long)]
        out_dir: PathBuf,
        /// Backend engine.
        #[arg(short, long)]
        backend: String,
        /// calamine mode (only for --backend calamine-values|calamine-formulas).
        #[arg(long)]
        calamine_mode: Option<String>,
        /// COM per-file lifecycle (only for --backend com-batch).
        #[arg(long)]
        com_per_file: bool,
        /// Recurse into input directories.
        #[arg(long)]
        recursive: bool,
        /// Output validation: none | quick | full (default full).
        #[arg(long, default_value = "full")]
        validate: String,
        /// Write a JSON run report here.
        #[arg(long)]
        report: Option<PathBuf>,
        /// Include rxls diagnostics per file (rxls backend only).
        #[arg(long)]
        with_diag: bool,
    },
    /// Compare converted XLSX output against the source XLS with independent readers.
    Compare {
        /// Source .xls file or directory.
        #[arg(short, long)]
        source: PathBuf,
        /// Converted location as label=path (repeatable; path is a dir of
        /// <stem>.xlsx or a single .xlsx file).
        #[arg(short = 'c', long = "converted")]
        converted: Vec<String>,
        /// Output JSON report path.
        #[arg(short, long)]
        out: PathBuf,
        /// Max mismatch samples per category.
        #[arg(long, default_value_t = compare::MAX_SAMPLES_DEFAULT)]
        samples: usize,
        /// Include per-cell style comparison (rxls style model).
        #[arg(long)]
        styles: bool,
        /// Reader for the values/formulas view: calamine (independent reader)
        /// or rxls (correct for legacy codepage sources calamine mis-decodes).
        #[arg(long, default_value = "calamine")]
        values_reader: String,
    },
    /// Open converted XLSX files in real Excel (COM) and probe compatibility.
    Opencheck {
        /// Directory containing converted .xlsx files.
        #[arg(short, long)]
        input: PathBuf,
        /// Output directory for reports, PDFs and round-trip files.
        #[arg(short, long)]
        out_dir: PathBuf,
        /// Export PDFs for every file.
        #[arg(long)]
        pdf: bool,
        /// Save a re-saved copy of every file and diff it against the original.
        #[arg(long)]
        roundtrip: bool,
        /// Also probe WPS (Ket.Application) with the first N files.
        #[arg(long, default_value_t = 0)]
        wps_files: usize,
    },
    /// Run the batch performance benchmark (parent orchestrator).
    Bench {
        /// Directory with .xls fixtures (sorted; first N used per batch size).
        #[arg(short, long)]
        fixtures_dir: PathBuf,
        /// Batch sizes, comma separated (default 10,50,100).
        #[arg(long, default_value = "10,50,100")]
        sizes: String,
        /// Timed rounds per combination (>=1).
        #[arg(long, default_value_t = 5)]
        rounds: u32,
        /// Untimed warmup rounds per combination.
        #[arg(long, default_value_t = 1)]
        warmup: u32,
        /// Engines, comma separated (default all five).
        #[arg(long)]
        engines: Option<String>,
        /// Output directory for raw/summary results.
        #[arg(short, long)]
        out_dir: PathBuf,
        /// Single-file fixed-cost test instead of the fixture set.
        #[arg(long)]
        single_file: Option<PathBuf>,
        /// Directory with canonical batch subsets (B<size>.json from the
        /// existing dataset manifest); SHA-256 entries are verified.
        #[arg(long)]
        batches_dir: Option<PathBuf>,
    },
    /// Internal: run one benchmark round in this process (used by `bench`).
    BenchRun {
        #[arg(long)]
        engine: String,
        #[arg(long)]
        files: PathBuf,
        #[arg(long)]
        size: usize,
        #[arg(long)]
        round: u32,
        #[arg(long)]
        warmup: bool,
        #[arg(long)]
        out_dir: PathBuf,
        #[arg(long)]
        report: PathBuf,
    },
    /// Generate fidelity test fixtures via Excel COM (synthetic data only).
    GenFidelity {
        /// Output directory (default benchmark/xls-to-xlsx-rxls/fixtures/fidelity).
        #[arg(short, long)]
        out_dir: Option<PathBuf>,
    },
    /// Generate failure-safety fixtures from a valid source XLS.
    Mkcorrupt {
        /// Valid source .xls.
        #[arg(short, long)]
        source: PathBuf,
        /// Output directory.
        #[arg(short, long)]
        out_dir: PathBuf,
    },
}

fn main() {
    let cli = Cli::parse();
    let code = match dispatch(cli.command) {
        Ok(()) => 0,
        Err(error) => {
            eprintln!("error: {error:#}");
            1
        }
    };
    std::process::exit(code);
}

fn dispatch(command: Command) -> Result<()> {
    match command {
        Command::Convert {
            inputs,
            out_dir,
            backend,
            calamine_mode,
            com_per_file,
            recursive,
            validate,
            report,
            with_diag,
        } => {
            let engine = Engine::parse(&backend)
                .ok_or_else(|| anyhow::anyhow!("unknown backend '{backend}'"))?;
            let level = Validate::parse(&validate)
                .ok_or_else(|| anyhow::anyhow!("unknown validate '{validate}'"))?;
            run_convert(
                engine,
                inputs,
                out_dir,
                calamine_mode,
                com_per_file,
                recursive,
                level,
                report,
                with_diag,
            )
        }
        Command::Compare {
            source,
            converted,
            out,
            samples,
            styles,
            values_reader,
        } => {
            if converted.is_empty() {
                bail!("at least one --converted label=path is required");
            }
            anyhow::ensure!(
                values_reader == "calamine" || values_reader == "rxls",
                "values-reader must be calamine or rxls"
            );
            let mut entries = Vec::new();
            for item in converted {
                let (label, path) = match item.split_once('=') {
                    Some(pair) => pair,
                    None => bail!("--converted must be label=path, got '{item}'"),
                };
                entries.push((label.to_owned(), PathBuf::from(path)));
            }
            compare::run(
                &source,
                entries,
                &out,
                samples,
                styles,
                &values_reader,
            )
        }
        Command::Opencheck {
            input,
            out_dir,
            pdf,
            roundtrip,
            wps_files,
        } => {
            #[cfg(windows)]
            {
                opencheck::run(&input, &out_dir, pdf, roundtrip, wps_files)
            }
            #[cfg(not(windows))]
            {
                let _ = (input, out_dir, pdf, roundtrip, wps_files);
                bail!("opencheck requires Windows (Excel COM)");
            }
        }
        Command::Bench {
            fixtures_dir,
            sizes,
            rounds,
            warmup,
            engines,
            out_dir,
            single_file,
            batches_dir,
        } => {
            let sizes: Vec<usize> = sizes
                .split(',')
                .map(|s| {
                    s.trim()
                        .parse::<usize>()
                        .map_err(|_| anyhow::anyhow!("bad size '{s}'"))
                })
                .collect::<Result<Vec<_>>>()?;
            anyhow::ensure!(!sizes.is_empty(), "at least one size required");
            anyhow::ensure!(rounds >= 1, "rounds must be >= 1");
            let engines: Vec<Engine> = match engines {
                Some(list) => list
                    .split(',')
                    .map(|s| {
                        Engine::parse(s.trim())
                            .ok_or_else(|| anyhow::anyhow!("unknown engine '{s}'"))
                    })
                    .collect::<Result<Vec<_>>>()?,
                None => vec![
                    Engine::ComBatch,
                    Engine::ComPerFile,
                    Engine::Rxls,
                    Engine::CalamineValues,
                    Engine::CalamineFormulas,
                ],
            };
            bench::run(
                &fixtures_dir,
                sizes,
                rounds,
                warmup,
                engines,
                &out_dir,
                single_file.as_deref(),
                batches_dir.as_deref(),
            )
        }
        Command::BenchRun {
            engine,
            files,
            size,
            round,
            warmup,
            out_dir,
            report,
        } => {
            let engine = Engine::parse(&engine)
                .ok_or_else(|| anyhow::anyhow!("unknown engine '{engine}'"))?;
            bench::run_child(engine, &files, size, round, warmup, &out_dir, &report)
        }
        Command::GenFidelity { out_dir } => {
            #[cfg(windows)]
            {
                let out_dir = out_dir.unwrap_or_else(|| {
                    std::env::current_dir()
                        .unwrap_or_else(|_| PathBuf::from("."))
                        .join("fixtures/fidelity")
                });
                fxgen::run(&out_dir)
            }
            #[cfg(not(windows))]
            {
                let _ = out_dir;
                bail!("genfidelity requires Windows (Excel COM)");
            }
        }
        Command::Mkcorrupt { source, out_dir } => mkcorrupt::run(&source, &out_dir),
    }
}

fn run_convert(
    engine: Engine,
    inputs: Vec<PathBuf>,
    out_dir: PathBuf,
    calamine_mode: Option<String>,
    com_per_file: bool,
    recursive: bool,
    validate: Validate,
    report: Option<PathBuf>,
    with_diag: bool,
) -> Result<()> {
    use common::BatchSummary;
    let files = common::list_xls_files(&inputs, recursive)?;
    anyhow::ensure!(!files.is_empty(), "no .xls inputs found");
    std::fs::create_dir_all(&out_dir)?;

    let mode = match engine {
        Engine::CalamineValues | Engine::CalamineFormulas => Some(
            calamine_mode
                .unwrap_or_else(|| "values".to_owned()),
        ),
        Engine::ComBatch => {
            if com_per_file {
                Some("per-file".to_owned())
            } else {
                Some("batch".to_owned())
            }
        }
        _ => None,
    };
    let mut summary = BatchSummary::new(engine, mode);
    let t_start = std::time::Instant::now();

    // Pre-check outputs so every backend behaves identically: never overwrite.
    let mut pending: Vec<PathBuf> = Vec::new();
    for input in &files {
        let stem = input
            .file_stem()
            .map(|s| s.to_string_lossy().into_owned())
            .unwrap_or_else(|| "workbook".to_owned());
        let output = out_dir.join(format!("{stem}.xlsx"));
        if output.exists() {
            let mut outcome = common::FileOutcome::new(input, &output);
            outcome.status = "skipped".to_owned();
            outcome.error = Some(format!("output_exists: {}", output.display()));
            summary.files.push(outcome);
        } else {
            pending.push(input.clone());
        }
    }

    match engine {
        Engine::Rxls => {
            let outcomes = bx_rxls::convert_batch(&pending, &out_dir, validate, with_diag)?;
            summary.files.extend(outcomes);
        }
        Engine::CalamineValues | Engine::CalamineFormulas => {
            let mode = if engine == Engine::CalamineFormulas {
                bx_calamine::CalamineMode::Formulas
            } else {
                bx_calamine::CalamineMode::Values
            };
            let outcomes = bx_calamine::convert_batch(&pending, &out_dir, mode, validate)?;
            summary.files.extend(outcomes);
        }
        Engine::ComBatch | Engine::ComPerFile => {
            #[cfg(windows)]
            {
                let pairs: Vec<(PathBuf, PathBuf)> = pending
                    .iter()
                    .map(|input| {
                        let stem = input
                            .file_stem()
                            .map(|s| s.to_string_lossy().into_owned())
                            .unwrap_or_else(|| "workbook".to_owned());
                        (input.clone(), out_dir.join(format!("{stem}.xlsx")))
                    })
                    .collect();
                let (outcomes, meta) = bx_com::convert(&pairs, engine == Engine::ComPerFile)?;
                summary.com = Some(meta);
                summary.files.extend(outcomes);
            }
            #[cfg(not(windows))]
            {
                let _ = (pending, com_per_file);
                bail!("COM engines require Windows");
            }
        }
    }

    summary.total_ms = t_start.elapsed().as_secs_f64() * 1000.0;
    summary.finalize();
    let report_path = report.unwrap_or_else(|| {
        out_dir.join(format!(
            "convert-report-{}-{}.json",
            engine.id(),
            common::now_epoch_ms()
        ))
    });
    if let Some(parent) = report_path.parent() {
        std::fs::create_dir_all(parent)?;
    }
    std::fs::write(&report_path, serde_json::to_string_pretty(&summary)?)?;
    eprintln!(
        "convert done: engine={} ok={} failed={} skipped={} total_ms={:.0} report={}",
        engine.id(),
        summary.ok_count,
        summary.failed_count,
        summary.skipped_count,
        summary.total_ms,
        report_path.display()
    );
    if summary.failed_count > 0 {
        bail!("{} file(s) failed to convert", summary.failed_count);
    }
    Ok(())
}
