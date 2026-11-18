//! Targeted SDK archive smoke; no PyTools post-processing or menu changes.
use aspose_worker_host::*;
use serde_json::json;
use sha2::{Digest, Sha256};
use std::{fs, path::PathBuf, time::Duration};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<_> = std::env::args().collect();
    if args.len() != 5 && !(args.len() == 6 && args[5] == "legacy") {
        return Err("archive_case DOTNET WORKER_DLL SYNTHETIC_XLSX NEW_OUTPUT_DIRECTORY".into());
    }
    let input = fs::canonicalize(&args[3])?;
    let before = Sha256::digest(fs::read(&input)?);
    fs::create_dir(&args[4])?;
    let root = fs::canonicalize(&args[4])?;
    let imported = root.join("import-source.xlsx");
    fs::copy(&input, &imported)?;
    let imported_before = Sha256::digest(fs::read(&imported)?);
    let mut backend = AsposeWorkerBackend::new(WorkerConfig {
        executable: PathBuf::from(&args[1]),
        arguments: vec![args[2].clone(), "--stdio".into()],
        diagnostics_dir: root.join("diagnostics"),
        startup_timeout: Duration::from_secs(30),
        request_timeout: Duration::from_secs(90),
    });
    let mut records = Vec::new();
    for method in ["Calculate", "Modify", "Convert", "Import"] {
        let request = WorkbookRequest {
            run_id: "raw-archive-host-smoke".into(),
            input: input.clone(),
            candidate: root.join(format!("{}.xlsx", method.to_lowercase())),
        };
        let result = match method {
            "Calculate" => backend.calculate(&request)?,
            "Convert" => backend.convert(&request)?,
            "Import" => backend.modify(&ModifyRequest {
                workbook: request,
                operations: vec![WorkbookOperation::EditPlanOperation { operation: json!({
                    "operation":"import_range", "source_workbook":imported, "source_sheet":"Data",
                    "source_range":"A1:A2", "target_sheet":"Data", "target_start":"D3",
                    "copy_values":true, "copy_formulas":false,"style_mode":"none","copy_comments":false,
                    "copy_merges":false, "copy_row_heights":false,"copy_column_widths":false,
                    "copy_row_hidden":false,"copy_column_hidden":false,"copy_outline_levels":false,
                    "copy_all_row_dimensions":false,"copy_all_column_dimensions":false,
                    "materialize_blank_cells":true,"hide_zero":false
                }) }],
            })?,
            _ => backend.modify(&ModifyRequest {
                workbook: request,
                operations: vec![WorkbookOperation::SetValue {
                    sheet: "Data".into(),
                    cell: "C3".into(),
                    value: CellValue::Text("updated".into()),
                }],
            })?,
        };
        assert!(result
            .output_sheets
            .iter()
            .any(|s| s == "Evaluation Warning"));
        assert_eq!(
            result.sdk_save_archive.is_some(),
            args.len() == 5,
            "Archive availability must match the actual Worker capability"
        );
        assert_eq!(result.timings.raw_sdk_archive_ms.is_some(), args.len() == 5);
        let trace: serde_json::Value = serde_json::from_slice(&fs::read(
            result.input_trace.as_ref().ok_or("missing input audit")?,
        )?)?;
        assert_eq!(trace["logical_business_inputs_complete"], false);
        assert_eq!(trace["inputs"][0]["sheets"], json!(["Data"]));
        assert_eq!(
            trace["inputs"].as_array().unwrap().len(),
            if method == "Import" { 2 } else { 1 }
        );
        if method == "Import" {
            assert_eq!(trace["inputs"][1]["sheets"], json!(["Data"]));
        }
        records.push(result);
    }
    backend.shutdown();
    assert_eq!(before, Sha256::digest(fs::read(input)?));
    assert_eq!(imported_before, Sha256::digest(fs::read(imported)?));
    let report = json!({"status":"PASS", "cases":records, "source_unchanged":true,
        "evaluation_retained":true, "production_post_processor_invoked":false});
    fs::write(
        root.join("RESULT.json"),
        serde_json::to_vec_pretty(&report)?,
    )?;
    println!("{}", json!({"status":"PASS", "cases":4, "output":root}));
    Ok(())
}
