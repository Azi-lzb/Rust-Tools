use aspose_worker_host::*;
use std::{
    fs,
    path::PathBuf,
    time::{Duration, Instant},
};

fn backend(dotnet: &str, dll: &str, logs: PathBuf) -> AsposeWorkerBackend {
    AsposeWorkerBackend::new(WorkerConfig {
        executable: dotnet.into(),
        arguments: vec![dll.into(), "--stdio".into()],
        diagnostics_dir: logs,
        startup_timeout: Duration::from_secs(30),
        request_timeout: Duration::from_secs(60),
    })
}
fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = std::env::args().collect();
    if args.len() < 6 {
        return Err("usage: worker_case DOTNET DLL INPUT OUTPUT_DIR COUNT [fresh]".into());
    }
    let input = fs::canonicalize(&args[3])?;
    let root = PathBuf::from(&args[4]);
    fs::create_dir(&root)?;
    let root = fs::canonicalize(root)?;
    let count: usize = args[5].parse()?;
    let fresh = args.get(6).is_some_and(|s| s == "fresh");
    let mut engine = backend(&args[1], &args[2], root.join("logs"));
    let started_at = chrono::Utc::now().to_rfc3339();
    let run_id = format!(
        "aspose-{}-{}",
        std::process::id(),
        chrono::Utc::now().timestamp_nanos_opt().unwrap()
    );
    let start = Instant::now();
    let mut records = Vec::new();
    for i in 0..count {
        if fresh && i > 0 {
            engine.shutdown();
        }
        let request = WorkbookRequest {
            run_id: run_id.clone(),
            input: input.clone(),
            candidate: root.join(format!("candidate-{i:03}.xlsx")),
        };
        let record = if i == 0 {
            engine.calculate(&request)?
        } else {
            engine.modify(&ModifyRequest {
                workbook: request,
                operations: vec![
                    WorkbookOperation::SetValue {
                        sheet: "Data".into(),
                        cell: "D12".into(),
                        value: CellValue::Number(9999.0),
                    },
                    WorkbookOperation::SetStyle {
                        sheet: "Data".into(),
                        cell: "D12".into(),
                        style: CellStyle {
                            bold: Some(true),
                            fill_rgb: Some("FFE699".into()),
                            ..Default::default()
                        },
                    },
                ],
            })?
        };
        records.push(record);
    }
    engine.shutdown();
    let report = serde_json::json!({"records":records,"count":count,"fresh":fresh,
        "generations":engine.generations(),"total_including_shutdown_ms":start.elapsed().as_secs_f64()*1000.0,
        "continued_in_rust":true,"evaluation_output_accepted":true,"menu_route_changed":false,
        "run_id":run_id,"started_at":started_at,"finished_at":chrono::Utc::now().to_rfc3339()});
    fs::write(
        root.join("host-report.json"),
        serde_json::to_vec_pretty(&report)?,
    )?;
    println!(
        "{}",
        serde_json::json!({"count":count,"fresh":fresh,"generations":engine.generations(),"total_ms":report["total_including_shutdown_ms"],"report":root.join("host-report.json")})
    );
    Ok(())
}
