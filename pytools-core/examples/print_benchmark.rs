//! Test-only resident API driver. Production menus and defaults are unchanged.
use anyhow::{Context, Result};
use pytools_core::static_print_engines::{run_configured_static_print, StaticPrintEngine};
use serde_json::{json, Value};
use std::{
    io::{self, BufRead, Write},
    path::PathBuf,
    time::Instant,
};

fn main() -> Result<()> {
    println!("{}", json!({"ready":true,"pid":std::process::id()}));
    io::stdout().flush()?;
    for line in io::stdin().lock().lines() {
        let q: Value = serde_json::from_str(&line?)?;
        let id = q["run_id"].as_str().context("run_id missing")?;
        pytools_core::aspose_edit::bind_run_id(id);
        let engine = match q["engine"].as_str() {
            Some("openpyxl") => StaticPrintEngine::Openpyxl,
            Some("rust") => StaticPrintEngine::RustWriter,
            Some("aspose") => StaticPrintEngine::Aspose,
            _ => anyhow::bail!("unknown benchmark engine"),
        };
        let config = PathBuf::from(q["config"].as_str().context("config missing")?);
        let started = Instant::now();
        let result = run_configured_static_print(&config, engine);
        let elapsed = started.elapsed().as_secs_f64() * 1000.0;
        let reply = match result {
            Ok(report) => json!({"run_id":id,"ok":report.task_failed==0,"pipeline_ms":elapsed,
                "task_ok":report.task_ok,"task_failed":report.task_failed,"task_skipped":report.task_skipped,
                "outputs":report.outputs,"issues":report.issues,"events":report.events}),
            Err(error) => {
                json!({"run_id":id,"ok":false,"pipeline_ms":elapsed,"error":format!("{error:#}")})
            }
        };
        println!("{reply}");
        io::stdout().flush()?;
    }
    Ok(())
}
