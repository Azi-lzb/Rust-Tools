use anyhow::{Context, Result, bail};
use betteroffice_xlsx::{CalculationOptions, CellRef, CellValue, Workbook as BetterWorkbook};
use formualizer_workbook::{
    LoadStrategy, SpreadsheetReader, UmyaAdapter, Workbook as FormualizerWorkbook, WorkbookConfig,
    XlsxRecalculateOptions, recalculate_xlsx_file,
};
use serde::{Deserialize, Serialize};
use std::{
    collections::BTreeMap,
    env, fs,
    path::{Path, PathBuf},
    time::Instant,
};

#[derive(Debug, Deserialize)]
struct FixtureSpec {
    targets: Vec<Target>,
    mutations: Vec<Mutation>,
}

#[derive(Debug, Deserialize)]
struct Target {
    sheet: String,
    row: u32,
    col: u32,
    address: String,
}

#[derive(Debug, Deserialize)]
struct Mutation {
    sheet: String,
    row: u32,
    col: u32,
    value: f64,
}

#[derive(Debug, Serialize)]
struct EngineRun {
    engine: String,
    load_ms: f64,
    calculate_ms: f64,
    extract_ms: f64,
    save_ms: f64,
    incremental_modify_ms: Option<f64>,
    incremental_recalculate_ms: Option<f64>,
    incremental_extract_ms: Option<f64>,
    formula_targets: usize,
    engine_reported_vertices: Option<usize>,
    formula_results: BTreeMap<String, String>,
    incremental_results: BTreeMap<String, String>,
    output_path: String,
    status: String,
    diagnostics: Vec<String>,
}

fn ms(start: Instant) -> f64 {
    start.elapsed().as_secs_f64() * 1000.0
}

fn read_spec(path: &Path) -> Result<FixtureSpec> {
    let bytes = fs::read(path).with_context(|| format!("read fixture spec {}", path.display()))?;
    Ok(serde_json::from_slice(&bytes)?)
}

fn better_value(value: &CellValue) -> String {
    match value {
        CellValue::Empty => "EMPTY".into(),
        CellValue::Number { value } => format!("NUM:{value:.15}"),
        CellValue::Text { value } => format!("TEXT:{value}"),
        CellValue::Bool { value } => format!("BOOL:{value}"),
        CellValue::Error { value } => format!("ERROR:{}", value.as_str()),
    }
}

fn form_value(value: Option<formualizer_workbook::LiteralValue>) -> String {
    use formualizer_workbook::LiteralValue as V;
    match value.unwrap_or(V::Empty) {
        V::Empty => "EMPTY".into(),
        V::Number(n) => format!("NUM:{n:.15}"),
        V::Int(n) => format!("NUM:{n}"),
        V::Text(s) => format!("TEXT:{s}"),
        V::Boolean(b) => format!("BOOL:{b}"),
        V::Error(e) => format!("ERROR:{}", e.kind),
        V::Date(d) => format!("DATE:{d}"),
        V::DateTime(d) => format!("DATETIME:{d}"),
        V::Time(t) => format!("TIME:{t}"),
        V::Duration(d) => format!("DURATION:{d}"),
        V::Array(a) => format!("ARRAY:{a:?}"),
        V::Pending => "PENDING".into(),
    }
}

fn better_results(wb: &BetterWorkbook, spec: &FixtureSpec) -> Result<BTreeMap<String, String>> {
    let mut out = BTreeMap::new();
    for t in &spec.targets {
        let sid = wb
            .sheet_id(&t.sheet)
            .with_context(|| format!("missing sheet {}", t.sheet))?;
        let sheet = wb.sheet(sid)?;
        let cell = sheet.cell(CellRef::new(t.row, t.col));
        out.insert(
            t.address.clone(),
            cell.map(|c| better_value(&c.value))
                .unwrap_or_else(|| "EMPTY".into()),
        );
    }
    Ok(out)
}

fn form_results(wb: &FormualizerWorkbook, spec: &FixtureSpec) -> BTreeMap<String, String> {
    spec.targets
        .iter()
        .map(|t| {
            // Formualizer's public cell coordinates are 1-based; the fixture
            // manifest stays zero-based so BetterOffice can consume it directly.
            (
                t.address.clone(),
                form_value(wb.get_value(&t.sheet, t.row + 1, t.col + 1)),
            )
        })
        .collect()
}

fn run_better(input: &Path, output: &Path, spec: &FixtureSpec) -> Result<EngineRun> {
    let load = Instant::now();
    let bytes = fs::read(input)?;
    let mut wb = BetterWorkbook::open(&bytes)?;
    let load_ms = ms(load);
    let calc = Instant::now();
    let mut diagnostics = Vec::new();
    // Matches rustTools' production bo_calc policy: a second pass handles dynamic
    // cross-sheet references which may settle only after the first pass.
    for round in 1..=2 {
        let result = wb.recalculate_all(CalculationOptions {
            now_serial: Some(46_000.0),
        });
        diagnostics.push(format!(
            "round {round}: changed_value_cells={}",
            result.changed.len()
        ));
        if !result.cycle_cells.is_empty() {
            diagnostics.push(format!(
                "round {round}: cycles={}",
                result.cycle_cells.len()
            ));
        }
        if !result.limited_cells.is_empty() {
            diagnostics.push(format!(
                "round {round}: limited={}",
                result.limited_cells.len()
            ));
        }
    }
    let calculate_ms = ms(calc);
    let extract = Instant::now();
    let formula_results = better_results(&wb, spec)?;
    let extract_ms = ms(extract);
    let save = Instant::now();
    fs::write(output, wb.save()?)?;
    let save_ms = ms(save);

    let mut incremental_modify_ms = None;
    let mut incremental_recalculate_ms = None;
    let mut incremental_extract_ms = None;
    let mut incremental_results = BTreeMap::new();
    if let Some(m) = spec.mutations.first() {
        let sid = wb
            .sheet_id(&m.sheet)
            .context("incremental input sheet absent")?;
        let start = Instant::now();
        let _mutation = wb.edit_cell(
            sid,
            CellRef::new(m.row, m.col),
            &m.value.to_string(),
            CalculationOptions {
                now_serial: Some(46_000.0),
            },
        )?;
        incremental_modify_ms = Some(ms(start));
        incremental_recalculate_ms = Some(0.0); // edit_cell includes dirty dependency recalc.
        let start = Instant::now();
        incremental_results = better_results(&wb, spec)?;
        incremental_extract_ms = Some(ms(start));
    }
    Ok(EngineRun {
        engine: "BetterOffice 0.5.1 / production recalculate_all x2".into(),
        load_ms,
        calculate_ms,
        extract_ms,
        save_ms,
        incremental_modify_ms,
        incremental_recalculate_ms,
        incremental_extract_ms,
        formula_targets: spec.targets.len(),
        engine_reported_vertices: None,
        formula_results,
        incremental_results,
        output_path: output.display().to_string(),
        status: "EXECUTED".into(),
        diagnostics,
    })
}

fn run_formualizer(input: &Path, output: &Path, spec: &FixtureSpec) -> Result<EngineRun> {
    let load = Instant::now();
    let adapter = UmyaAdapter::open_path(input).context("Formualizer UmyaAdapter open_path")?;
    let mut wb = FormualizerWorkbook::from_reader(
        adapter,
        LoadStrategy::EagerAll,
        WorkbookConfig::interactive(),
    )?;
    let load_ms = ms(load);
    let calc = Instant::now();
    let eval = wb.evaluate_all()?;
    let calculate_ms = ms(calc);
    let extract = Instant::now();
    let formula_results = form_results(&wb, spec);
    let extract_ms = ms(extract);
    let save = Instant::now();
    fs::write(output, wb.to_xlsx_bytes()?)?;
    let save_ms = ms(save);

    let mut incremental_modify_ms = None;
    let mut incremental_recalculate_ms = None;
    let mut incremental_extract_ms = None;
    let mut incremental_results = BTreeMap::new();
    let mut diagnostics = Vec::new();
    if let Some(m) = spec.mutations.first() {
        let start = Instant::now();
        wb.set_value(
            &m.sheet,
            m.row + 1,
            m.col + 1,
            formualizer_workbook::LiteralValue::Number(m.value),
        )?;
        incremental_modify_ms = Some(ms(start));
        let start = Instant::now();
        let (_eval, delta) = wb.evaluate_all_with_delta()?;
        incremental_recalculate_ms = Some(ms(start));
        diagnostics.push(format!(
            "incremental_changed_cells={}",
            delta.changed_cells.len()
        ));
        let start = Instant::now();
        incremental_results = form_results(&wb, spec);
        incremental_extract_ms = Some(ms(start));
    }
    Ok(EngineRun {
        engine: "Formualizer 0.14.0 / UmyaAdapter resident workbook".into(),
        load_ms,
        calculate_ms,
        extract_ms,
        save_ms,
        incremental_modify_ms,
        incremental_recalculate_ms,
        incremental_extract_ms,
        formula_targets: spec.targets.len(),
        engine_reported_vertices: Some(eval.computed_vertices),
        formula_results,
        incremental_results,
        output_path: output.display().to_string(),
        status: "EXECUTED".into(),
        diagnostics,
    })
}

fn run_cache(input: &Path, output: &Path, require_complete: bool) -> Result<serde_json::Value> {
    let start = Instant::now();
    let mut options = XlsxRecalculateOptions::workbook();
    options.require_complete = require_complete;
    let result = match recalculate_xlsx_file(input, Some(output), options) {
        Ok(result) => result,
        Err(error) => {
            return Ok(serde_json::json!({
                "engine": "Formualizer 0.14.0 / source-aware cache recalculation",
                "require_complete": require_complete,
                "total_ms": ms(start),
                "status": "REFUSED",
                "error": error.to_string(),
                "output_path": output.display().to_string(),
                "output_published": output.exists(),
            }));
        }
    };
    let summary = &result.summary;
    let sheets: BTreeMap<_, _> = summary
        .sheets
        .iter()
        .map(|(name, item)| {
            (
                name.clone(),
                serde_json::json!({ "evaluated": item.evaluated, "errors": item.errors }),
            )
        })
        .collect();
    Ok(serde_json::json!({
        "engine": "Formualizer 0.14.0 / source-aware cache recalculation",
        "require_complete": require_complete,
        "total_ms": ms(start),
        "formula_cells": result.formula_cells,
        "cache_cells_changed": result.cache_cells_changed,
        "worksheet_parts_changed": result.worksheet_parts_changed,
        "summary": {
            "status": summary.status.as_str(),
            "evaluated": summary.evaluated,
            "errors": summary.errors,
            "sheets": sheets,
            "unknown_functions": summary.unknown_functions,
        },
        "output_path": output.display().to_string(),
        "status": "EXECUTED"
    }))
}

fn as_excel_number(value: &str) -> Option<f64> {
    if let Some(raw) = value.strip_prefix("NUM:") {
        return raw.parse().ok();
    }
    let date = value.strip_prefix("DATE:")?;
    let mut parts = date.split('-').map(str::parse::<i32>);
    let (year, month, day) = (
        parts.next()?.ok()?,
        parts.next()?.ok()?,
        parts.next()?.ok()?,
    );
    Some(days_from_civil(year, month, day) - days_from_civil(1899, 12, 30))
}

// Howard Hinnant's civil date conversion, with days since 1970-01-01.
fn days_from_civil(year: i32, month: i32, day: i32) -> f64 {
    let y = year - i32::from(month <= 2);
    let era = y.div_euclid(400);
    let yoe = y - era * 400;
    let mp = month + if month > 2 { -3 } else { 9 };
    let doy = (153 * mp + 2) / 5 + day - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    (era * 146097 + doe - 719468) as f64
}

fn values_equal(actual: &str, expected: &str) -> bool {
    match (as_excel_number(actual), as_excel_number(expected)) {
        (Some(a), Some(b)) => (a - b).abs() <= 1e-9_f64.max(1e-9 * a.abs().max(b.abs())),
        (Some(_), None) | (None, Some(_)) => false,
        (None, None) => actual == expected,
    }
}

fn compare_results(
    actual: &BTreeMap<String, String>,
    oracle: &BTreeMap<String, String>,
) -> serde_json::Value {
    let mut mismatches = Vec::new();
    let mut matches = 0usize;
    for (address, expected) in oracle {
        match actual.get(address) {
            Some(value) if values_equal(value, expected) => matches += 1,
            Some(value) => mismatches.push(
                serde_json::json!({"address": address, "actual": value, "expected": expected}),
            ),
            None => mismatches.push(
                serde_json::json!({"address": address, "actual": null, "expected": expected}),
            ),
        }
    }
    serde_json::json!({
        "status": if mismatches.is_empty() { "MATCH" } else { "MISMATCH" },
        "matched": matches,
        "total": oracle.len(),
        "incorrect": mismatches.len(),
        "mismatches": mismatches,
    })
}

fn main() -> Result<()> {
    let args: Vec<String> = env::args().collect();
    if args.len() < 4 {
        bail!("usage: formualizer-vs-betteroffice run <input.xlsx> <output-dir> [cases.json]");
    }
    if args[1] != "run" {
        bail!("unknown command: {}", args[1]);
    }
    let input = PathBuf::from(&args[2]);
    let out_dir = PathBuf::from(&args[3]);
    let spec_path = args.get(4).map(PathBuf::from).unwrap_or_else(|| {
        input
            .parent()
            .unwrap_or_else(|| Path::new("."))
            .join("cases.json")
    });
    fs::create_dir_all(&out_dir)?;
    let spec = read_spec(&spec_path)?;
    let better_path = out_dir.join("betteroffice.xlsx");
    let form_path = out_dir.join("formualizer.xlsx");
    let cache_path = out_dir.join("formualizer-cache.xlsx");
    let better = run_better(&input, &better_path, &spec)?;
    let form = run_formualizer(&input, &form_path, &spec)?;
    let cache = run_cache(&input, &cache_path, false)?;
    let strict_cache = run_cache(&input, &out_dir.join("formualizer-cache-strict.xlsx"), true)?;
    let fixture_dir = spec_path.parent().unwrap_or_else(|| Path::new("."));
    let oracle_path = args
        .get(5)
        .map(PathBuf::from)
        .unwrap_or_else(|| fixture_dir.join("excel-com-oracle.json"));
    let incremental_oracle_path = args
        .get(6)
        .map(PathBuf::from)
        .unwrap_or_else(|| fixture_dir.join("excel-com-incremental-oracle.json"));
    let oracle: BTreeMap<String, String> = serde_json::from_slice(&fs::read(&oracle_path)?)?;
    let incremental_oracle: BTreeMap<String, String> =
        serde_json::from_slice(&fs::read(&incremental_oracle_path)?)?;
    let correctness = serde_json::json!({
        "reference": "Excel COM 16.0 CalculateFullRebuild",
        "absolute_tolerance": 1e-9,
        "relative_tolerance": 1e-9,
        "betteroffice_full": compare_results(&better.formula_results, &oracle),
        "formualizer_full": compare_results(&form.formula_results, &oracle),
        "betteroffice_after_input_change": compare_results(&better.incremental_results, &incremental_oracle),
        "formualizer_after_input_change": compare_results(&form.incremental_results, &incremental_oracle),
    });
    let result = serde_json::json!({ "input": input, "betteroffice": better, "formualizer": form, "formualizer_cache": cache, "formualizer_cache_strict": strict_cache, "correctness": correctness });
    let report = out_dir.join("engine-runs.json");
    fs::write(&report, serde_json::to_vec_pretty(&result)?)?;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}
