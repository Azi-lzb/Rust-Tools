//! Standalone XLSX modify-engine benchmark harness.
//!
//! Engines under test: betteroffice-xlsx 0.5.1 (vendored patches), rxls 0.1.4,
//! umya-spreadsheet 3.1.1, openpyxl 3.1.5 (via python runners).
//! Milestone-2 harness only; does not touch rustTools business code.

use anyhow::{bail, Context, Result};
use serde_json::{json, Value};
use std::collections::BTreeMap;
use std::path::{Path, PathBuf};
use std::time::Instant;

mod engines;
mod fidelity;
mod perf;
pub use engines::{parse_cell, parse_range, probe_ops, probe_fixture, run_probe, Val, Op, Engine};

fn main() -> Result<()> {
    let args: Vec<String> = std::env::args().collect();
    match args.get(1).map(String::as_str) {
        Some("probe") => {
            let engine = arg(&args, 2)?;
            let fixtures = PathBuf::from(arg(&args, 3)?);
            let out_dir = PathBuf::from(arg(&args, 4)?);
            run_probe(&engine, &fixtures, &out_dir)
        }
        Some("fidelity-run") => {
            // fidelity-run --engine E --fixture F --scenario S --out-dir D --results R.jsonl
            let mut engine = None;
            let mut fixture = None;
            let mut scenario = None;
            let mut out_dir = None;
            let mut results = None;
            let mut i = 2;
            while i < args.len() {
                match args[i].as_str() {
                    "--engine" => { engine = Some(args[i + 1].clone()); i += 2; }
                    "--fixture" => { fixture = Some(args[i + 1].clone()); i += 2; }
                    "--scenario" => { scenario = Some(args[i + 1].clone()); i += 2; }
                    "--out-dir" => { out_dir = Some(args[i + 1].clone()); i += 2; }
                    "--results" => { results = Some(args[i + 1].clone()); i += 2; }
                    other => bail!("unknown flag {other}"),
                }
            }
            fidelity::run_fidelity(
                &engine.context("--engine")?,
                Path::new(&fixture.context("--fixture")?),
                &scenario.context("--scenario")?,
                Path::new(&out_dir.context("--out-dir")?),
                Path::new(&results.context("--results")?),
            )
        }
        Some("worker") => perf::worker_main(&args[2..]),
        Some("perf-single") => perf::perf_single(&args[2..]),
        Some("perf-batch") => perf::perf_batch(&args[2..]),
        _ => {
            eprintln!("commands: probe | fidelity-run | worker | perf-single | perf-batch");
            bail!("no command given");
        }
    }
}

fn arg(args: &[String], idx: usize) -> Result<String> {
    args.get(idx).cloned().context("missing argument")
}

pub fn read_json(path: &Path) -> Result<Value> {
    Ok(serde_json::from_slice(&std::fs::read(path)?)?)
}

pub fn sha256_hex(data: &[u8]) -> String {
    use sha2::{Digest, Sha256};
    let mut hasher = Sha256::new();
    hasher.update(data);
    let out = hasher.finalize();
    out.iter().map(|b| format!("{b:02x}")).collect()
}
