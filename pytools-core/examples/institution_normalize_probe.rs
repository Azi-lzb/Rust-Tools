use anyhow::{bail, Context, Result};
use excel_reader::read_workbook;
use std::{env, fs, path::PathBuf};

fn run() -> Result<()> {
    let args = env::args_os().skip(1).collect::<Vec<_>>();
    if args.len() != 3 {
        bail!("usage: institution_normalize_probe <input.xlsx> <mapping-cases.json> <dump.json>");
    }
    let input = PathBuf::from(&args[0]);
    let cases = PathBuf::from(&args[1]);
    let output = PathBuf::from(&args[2]);
    let cases_value: serde_json::Value = serde_json::from_slice(
        &fs::read(&cases).with_context(|| format!("read cases {}", cases.display()))?,
    )
    .with_context(|| format!("parse cases {}", cases.display()))?;
    let mapping = serde_json::from_value::<Vec<(String, String)>>(cases_value["mapping"].clone())
        .context("cases mapping must be an array of [key, standard] pairs")?;

    let output_candidate = output.with_extension("candidate.xlsx");
    let profile = excel_editor::inspect_xlsx_profile(&input)?;
    let business_stats = pytools_core::institution_normalize::build_institution_normalize_plan(
        &input,
        &output_candidate,
        profile,
        "institution-normalize-probe",
        &mapping,
    )?;
    let dump = read_workbook(&input).with_context(|| format!("dump {}", input.display()))?;
    let mut json = serde_json::to_value(dump)?;
    json["business_stats"] = serde_json::json!({
        "branch_sheets": business_stats.branch_sheets,
        "mapped_sheets": business_stats.mapped_sheets,
        "mapped_rows": business_stats.mapped_rows,
        "mapping_empty": business_stats.mapping_empty,
        "plan": business_stats.plan,
    });
    if let Some(parent) = output
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
    {
        fs::create_dir_all(parent)
            .with_context(|| format!("create output directory {}", parent.display()))?;
    }
    fs::write(&output, serde_json::to_vec_pretty(&json)?)
        .with_context(|| format!("write dump {}", output.display()))?;
    println!(
        "branch_sheets={} mapped_sheets={} mapped_rows={} plan_missing={}",
        business_stats.branch_sheets,
        business_stats.mapped_sheets,
        business_stats.mapped_rows,
        usize::from(business_stats.plan.is_none() && !business_stats.mapping_empty)
    );
    Ok(())
}

fn main() {
    if let Err(error) = run() {
        eprintln!("{error:#}");
        std::process::exit(2);
    }
}
