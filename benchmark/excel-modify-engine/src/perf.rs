//! Performance runner: phase timings + peak memory via Windows Job Objects.
//!
//! Modes:
//!   worker <engine> <fixture> <ops.json> <out.xlsx>   — one edit, prints phase JSON
//!   worker-batch <engine> <ops.json> <files.json>     — loop files in one process
//!   perf-single --engine E --fixture F --workload W --reps N --results R.csv
//!   perf-batch --engine E --count N --mode cold|resident --reps N --results R.csv
//!
//! openpyxl runs through python scripts (native) or the production worker
//! (bridge); both are spawned as children under a Job Object so peak memory
//! is measured uniformly.

use anyhow::{bail, Context, Result};
use serde_json::{json, Value};
use std::path::{Path, PathBuf};
use std::time::Instant;

use crate::{engines::engine_by_name, read_json, Op, Val};

pub fn ops_for_workload(workload: &str, fixture: &Path) -> Vec<Op> {
    let sheet = "Main";
    let _ = fixture;
    match workload {
        "w1" => vec![Op::SetCellValue { sheet: sheet.into(), cell: "B2".into(), val: Val::Num(1.0) }],
        "w100" => (0..100)
            .map(|i| Op::SetCellValue {
                sheet: sheet.into(),
                cell: format!("{}{}", col_letter(1 + i % 10), 2 + (i / 10) as u32),
                val: Val::Num(100.0 + i as f64),
            })
            .collect(),
        "w1000" => (0..1000)
            .map(|i| Op::SetCellValue {
                sheet: sheet.into(),
                cell: format!("{}{}", col_letter(1 + i % 20), 2 + (i / 20) as u32),
                val: Val::Num(1000.0 + i as f64),
            })
            .collect(),
        "w-multi" => {
            let mut ops = Vec::new();
            for sheet_idx in 1..=5 {
                let sheet = format!("S{sheet_idx}");
                for i in 0..10 {
                    ops.push(Op::SetCellValue {
                        sheet: sheet.clone(),
                        cell: format!("B{}", 2 + i as u32),
                        val: Val::Num(42.0 + i as f64),
                    });
                }
            }
            ops
        }
        "w-formula" => (0..10)
            .map(|i| Op::SetFormula {
                sheet: sheet.into(),
                cell: format!("H{}", 2 + i as u32),
                text: "=A1*2".into(),
                cached: Some(Val::Num(4.0)),
            })
            .collect(),
        "w-mixed" => {
            let mut ops: Vec<Op> = (0..100)
                .map(|i| Op::SetCellValue {
                    sheet: sheet.into(),
                    cell: format!("{}{}", col_letter(1 + i % 10), 2 + (i / 10) as u32),
                    val: Val::Num(700.0 + i as f64),
                })
                .collect();
            ops.push(Op::SetRowHeight { sheet: sheet.into(), row: 40, h: 30.0 });
            ops.push(Op::AddMerge { sheet: sheet.into(), range: "L2:M3".into() });
            ops.push(Op::SetComment { sheet: sheet.into(), cell: "L2".into(), text: "mixed-op".into() });
            ops
        }
        _ => Vec::new(),
    }
}

fn col_letter(mut n: u32) -> String {
    let mut out = String::new();
    while n > 0 {
        let rem = ((n - 1) % 26) as u8;
        out.insert(0, (b'A' + rem) as char);
        n = (n - 1) / 26;
    }
    out
}

fn parse_flag<'a>(args: &'a [String], flag: &str) -> Option<&'a str> {
    args.iter().position(|a| a == flag).and_then(|i| args.get(i + 1)).map(String::as_str)
}

fn ops_to_json(ops: &[Op]) -> Value {
    json!({
        "ops": ops.iter().map(|op| {
            let mut obj = json!({"kind": op.kind()});
            match op {
                Op::SetCellValue { sheet, cell, val } => {
                    obj["sheet"] = json!(sheet); obj["cell"] = json!(cell);
                    obj["value"] = match val { Val::Text(t) => json!({"t":"text","v":t}),
                        Val::Num(n) => json!({"t":"num","v":n}), Val::Bool(b) => json!({"t":"bool","v":b}) };
                }
                Op::SetFormula { sheet, cell, text, .. } => {
                    obj["sheet"] = json!(sheet); obj["cell"] = json!(cell); obj["text"] = json!(text);
                }
                Op::SetNumberFormat { sheet, cell, fmt } => {
                    obj["sheet"] = json!(sheet); obj["cell"] = json!(cell); obj["fmt"] = json!(fmt);
                }
                Op::SetRowHeight { sheet, row, h } => {
                    obj["sheet"] = json!(sheet); obj["row"] = json!(row); obj["h"] = json!(h);
                }
                Op::SetColWidth { sheet, col, w } => {
                    obj["sheet"] = json!(sheet); obj["col"] = json!(col); obj["w"] = json!(w);
                }
                Op::AddMerge { sheet, range } | Op::RemoveMerge { sheet, range } => {
                    obj["sheet"] = json!(sheet); obj["range"] = json!(range);
                }
                Op::SetComment { sheet, cell, text } => {
                    obj["sheet"] = json!(sheet); obj["cell"] = json!(cell); obj["text"] = json!(text);
                }
                Op::SetGlobalDefinedName { name, reference } => {
                    obj["name"] = json!(name); obj["reference"] = json!(reference);
                }
                Op::SetPrintArea { sheet, range } => {
                    obj["sheet"] = json!(sheet); obj["range"] = json!(range);
                }
                Op::SetLandscape { sheet } => { obj["sheet"] = json!(sheet); }
                Op::AddSheet { name } | Op::DeleteSheet { name } => { obj["name"] = json!(name); }
                Op::RenameSheet { old, new } => { obj["old"] = json!(old); obj["new"] = json!(new); }
                Op::InsertRows { sheet, row, n } | Op::DeleteRows { sheet, row, n } => {
                    obj["sheet"] = json!(sheet); obj["row"] = json!(row); obj["n"] = json!(n);
                }
                Op::SetFontBold { sheet, cell } | Op::SetFillYellow { sheet, cell }
                | Op::SetBorderThin { sheet, cell } | Op::SetAlignCenter { sheet, cell } => {
                    obj["sheet"] = json!(sheet); obj["cell"] = json!(cell);
                }
            }
            obj
        }).collect::<Vec<_>>()
    })
}

fn ops_from_json(value: &Value) -> Vec<Op> {
    value["ops"]
        .as_array()
        .map(|ops| {
            ops.iter()
                .filter_map(|op| match op["kind"].as_str()? {
                    "set_cell_value" => Some(Op::SetCellValue {
                        sheet: op["sheet"].as_str()?.into(),
                        cell: op["cell"].as_str()?.into(),
                        val: json_to_val(&op["value"]),
                    }),
                    "set_formula" => Some(Op::SetFormula {
                        sheet: op["sheet"].as_str()?.into(),
                        cell: op["cell"].as_str()?.into(),
                        text: op["text"].as_str()?.into(),
                        cached: None,
                    }),
                    "set_row_height" => Some(Op::SetRowHeight {
                        sheet: op["sheet"].as_str()?.into(),
                        row: op["row"].as_u64()? as u32,
                        h: op["h"].as_f64()?,
                    }),
                    "set_col_width" => Some(Op::SetColWidth {
                        sheet: op["sheet"].as_str()?.into(),
                        col: op["col"].as_u64()? as u32,
                        w: op["w"].as_f64()?,
                    }),
                    "add_merge" => Some(Op::AddMerge {
                        sheet: op["sheet"].as_str()?.into(),
                        range: op["range"].as_str()?.into(),
                    }),
                    "set_comment" => Some(Op::SetComment {
                        sheet: op["sheet"].as_str()?.into(),
                        cell: op["cell"].as_str()?.into(),
                        text: op["text"].as_str()?.into(),
                    }),
                    _ => None,
                })
                .collect::<Vec<_>>()
        })
        .unwrap_or_default()
}

fn json_to_val(value: &Value) -> Val {
    match value["t"].as_str() {
        Some("text") => Val::Text(value["v"].as_str().unwrap_or_default().into()),
        Some("bool") => Val::Bool(value["v"].as_bool().unwrap_or(false)),
        _ => Val::Num(value["v"].as_f64().unwrap_or_default()),
    }
}

fn python_exe() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../dist/python/python.exe")
        .canonicalize()
        .expect("bundled python runtime must exist")
}

// ------------------------------------------------- job object peak memory

#[cfg(windows)]
mod job {
    use anyhow::Result;
    use std::os::windows::io::AsRawHandle;
    use std::process::{Child, Command};
    use anyhow::bail;
    use windows_sys::Win32::System::JobObjects::{
        AssignProcessToJobObject, CreateJobObjectW, QueryInformationJobObject,
        JOBOBJECT_EXTENDED_LIMIT_INFORMATION,
    };

    pub struct JobChild {
        child: Child,
        job: isize,
    }

    impl JobChild {
        pub fn spawn(command: &mut Command) -> Result<Self> {
            use std::os::windows::ffi::OsStrExt;
            let wide: Vec<u16> = std::ffi::OsStr::new("bench-job")
                .encode_wide()
                .chain(std::iter::once(0))
                .collect();
            let job = unsafe { CreateJobObjectW(std::ptr::null(), wide.as_ptr()) };
            if job.is_null() {
                bail!("CreateJobObjectW failed");
            }
            let mut child = command.spawn()?;
            if unsafe { AssignProcessToJobObject(job, child.as_raw_handle() as _) } == 0 {
                bail!("AssignProcessToJobObject failed");
            }
            Ok(JobChild { child, job: job as isize })
        }

        pub fn wait_with_peak_kb(mut self) -> Result<(std::process::ExitStatus, u64)> {
            let status = self.child.wait()?;
            let mut info: JOBOBJECT_EXTENDED_LIMIT_INFORMATION = unsafe { std::mem::zeroed() };
            let mut returned = 0u32;
            let ok = unsafe {
                QueryInformationJobObject(
                    self.job as _,
                    9, // JobObjectExtendedLimitInformation
                    &mut info as *mut _ as *mut core::ffi::c_void,
                    std::mem::size_of::<JOBOBJECT_EXTENDED_LIMIT_INFORMATION>() as u32,
                    &mut returned,
                )
            };
            let peak = if ok != 0 { info.PeakJobMemoryUsed } else { 0 };
            unsafe { windows_sys::Win32::Foundation::CloseHandle(self.job as _) };
            Ok((status, peak as u64 / 1024))
        }
    }
}

#[cfg(windows)]
use job::JobChild;

fn spawn_measured(command: &mut std::process::Command) -> Result<(std::process::Output, u64)> {
    #[cfg(windows)]
    {
        let child = JobChild::spawn(command)?;
        let (status, peak_kb) = child.wait_with_peak_kb()?;
        // Output was piped inside Command by caller; re-take via wait_with_output is
        // not possible after wait(), so callers must not use piped output with this
        // helper. Instead they redirect to files. We return only peak here.
        Ok((std::process::Output { status, stdout: Vec::new(), stderr: Vec::new() }, peak_kb))
    }
    #[cfg(not(windows))]
    {
        let output = command.output()?;
        Ok((output, 0))
    }
}

// ------------------------------------------------- worker modes

pub fn worker_main(args: &[String]) -> Result<()> {
    match args.first().map(String::as_str) {
        Some("single") => {
            // worker single <engine> <fixture> <ops.json> <out.xlsx>
            let engine_name = &args[1];
            let fixture = Path::new(&args[2]);
            let ops = ops_from_json(&crate::read_json(Path::new(&args[3]))?);
            let out_path = Path::new(&args[4]);
            let input = std::fs::read(fixture)?;
            let total_started = Instant::now();
            if engine_name == "openpyxl" {
                bail!("openpyxl runs through python scripts, not worker single");
            }
            let engine = engine_by_name(engine_name).context("unknown engine")?;
            let outcome = engine.apply(&input, &ops);
            let wrote_output = outcome.output.is_some();
            if let Some(output) = outcome.output.clone() {
                std::fs::write(out_path, &output)?;
            }
            let phases = outcome.phases.unwrap_or(crate::engines::PhaseTimes { open_ms: 0.0, modify_ms: 0.0, save_ms: 0.0 });
            let record = json!({
                "open_ms": phases.open_ms,
                "modify_ms": phases.modify_ms,
                "save_ms": phases.save_ms,
                "total_ms": total_started.elapsed().as_secs_f64() * 1000.0,
                "ok": wrote_output,
                "op_results": outcome.op_results.iter().map(|(k, r)| json!({
                    "op": k,
                    "ok": r.is_ok(),
                    "error": r.as_ref().err(),
                })).collect::<Vec<_>>(),
            });
            println!("{record}");
            Ok(())
        }
        Some("batch") => {
            // worker batch <engine> <ops.json> <files.json>
            let engine_name = &args[1];
            let ops = ops_from_json(&crate::read_json(Path::new(&args[2]))?);
            let files: Vec<String> = crate::read_json(Path::new(&args[3]))?
                .as_array()
                .map(|a| a.iter().filter_map(|v| v.as_str().map(str::to_owned)).collect())
                .unwrap_or_default();
            if engine_name == "openpyxl" {
                bail!("openpyxl runs through python scripts, not worker batch");
            }
            let engine = engine_by_name(engine_name).context("unknown engine")?;
            let mut rows = Vec::new();
            for file in &files {
                let started = Instant::now();
                let input = std::fs::read(file)?;
                let open_started = Instant::now();
                let outcome = engine.apply(&input, &ops);
                let open_ms = open_started.elapsed().as_secs_f64() * 1000.0;
                if let Some(output) = &outcome.output {
                    let out_path = Path::new(file).with_extension("out.xlsx");
                    std::fs::write(&out_path, output)?;
                }
                let phases = outcome.phases.unwrap_or(crate::engines::PhaseTimes { open_ms: 0.0, modify_ms: 0.0, save_ms: 0.0 });
                rows.push(json!({
                    "file": file,
                    "open_ms": phases.open_ms,
                    "modify_ms": phases.modify_ms,
                    "save_ms": phases.save_ms,
                    "total_ms": started.elapsed().as_secs_f64() * 1000.0,
                    "ok": outcome.output.is_some(),
                }));
            }
            println!("{}", serde_json::to_string(&json!({"rows": rows}))?);
            Ok(())
        }
        _ => bail!("worker subcommands: single | batch"),
    }
}

fn run_measured(command: &mut std::process::Command, stdout_path: &Path) -> Result<(Value, u64)> {
    let stdout_file = std::fs::File::create(stdout_path)?;
    command.stdout(std::process::Stdio::from(stdout_file.try_clone()?));
    command.stderr(std::process::Stdio::null());
    let (output, peak_kb) = spawn_measured(command)?;
    if !output.status.success() {
        bail!("child failed with {}", output.status);
    }
    let value: Value = serde_json::from_slice(&std::fs::read(stdout_path)?)?;
    Ok((value, peak_kb))
}

// ------------------------------------------------- perf-single

pub fn perf_single(args: &[String]) -> Result<()> {
    let engine = parse_flag(args, "--engine").context("--engine")?.to_owned();
    let fixture = PathBuf::from(parse_flag(args, "--fixture").context("--fixture")?);
    let workload = parse_flag(args, "--workload").context("--workload")?.to_owned();
    let reps: usize = parse_flag(args, "--reps").and_then(|v| v.parse().ok()).unwrap_or(5);
    let results = PathBuf::from(parse_flag(args, "--results").context("--results")?);
    if let Some(parent) = results.parent() {
        std::fs::create_dir_all(parent)?;
    }

    let ops_json = results.parent().unwrap().join(format!("ops_{workload}.json"));
    std::fs::write(&ops_json, serde_json::to_vec(&ops_to_json(&ops_for_workload(&workload, &fixture)))?)?;

    let mut samples: Vec<Value> = Vec::new();
    for rep in 0..reps {
        let stamp = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH)?.as_nanos();
        let out_path = results.parent().unwrap().join(format!("single-out-{stamp}.xlsx"));
        let stdout_path = results.parent().unwrap().join(format!("single-out-{stamp}.json"));
        let (record, peak_kb) = match engine.as_str() {
            "openpyxl-native" => {
                let mut command = std::process::Command::new(python_exe());
                command
                    .arg("-I")
                    .arg(PathBuf::from("scripts/py_engine.py").canonicalize()?)
                    .arg("bench")
                    .arg(fixture.canonicalize()?)
                    .arg(&out_path)
                    .arg(&workload);
                run_measured(&mut command, &stdout_path)?
            }
            "openpyxl-bridge" => {
                // production-shape: rust writes request file, spawns bundled
                // python with the worker, reads response file.
                let staged_in = results.parent().unwrap().join(format!("bridge-in-{stamp}.xlsx"));
                std::fs::write(&staged_in, std::fs::read(&fixture)?)?;
                let request = json!({
                    "schema_version": 1,
                    "request_id": format!("bench-{stamp}"),
                    "format": "xlsx",
                    "input": staged_in.to_string_lossy(),
                    "output": out_path.to_string_lossy(),
                    "plan": {
                        "workbook": {"inventory_complete": true},
                        "operations": ops_to_json(&ops_for_workload(&workload, &fixture))["ops"],
                    },
                });
                let request_path = results.parent().unwrap().join(format!("bridge-req-{stamp}.json"));
                let response_path = results.parent().unwrap().join(format!("bridge-resp-{stamp}.json"));
                std::fs::write(&request_path, serde_json::to_vec(&request)?)?;
                let worker = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../crates/excel-editor/python/openpyxl_worker.py").canonicalize()?;
                let started = Instant::now();
                let mut command = std::process::Command::new(python_exe());
                command
                    .arg("-I")
                    .arg(&worker)
                    .arg("--request")
                    .arg(request_path.canonicalize()?)
                    .arg("--response")
                    .arg(response_path.canonicalize()?);
                let peak_kb = {
                    #[cfg(windows)]
                    {
                        let child = JobChild::spawn(&mut command)?;
                        let (status, peak) = child.wait_with_peak_kb()?;
                        if !status.success() {
                            bail!("openpyxl bridge worker failed");
                        }
                        peak
                    }
                    #[cfg(not(windows))]
                    {
                        command.output()?;
                        0
                    }
                };
                let total_ms = started.elapsed().as_secs_f64() * 1000.0;
                let response = crate::read_json(&response_path)?;
                let _ = (response_path, request_path, staged_in);
                (
                    json!({
                        "open_ms": 0.0,
                        "modify_ms": 0.0,
                        "save_ms": 0.0,
                        "total_ms": total_ms,
                        "ok": response["status"] == "applied",
                        "phases_note": "bridge: phases not separable without protocol change",
                    }),
                    peak_kb,
                )
            }
            _ => {
                let exe = std::env::current_exe()?;
                let mut command = std::process::Command::new(exe);
                command
                    .arg("worker")
                    .arg("single")
                    .arg(&engine)
                    .arg(fixture.canonicalize()?)
                    .arg(ops_json.canonicalize()?)
                    .arg(canon_future(&out_path)?);
                run_measured(&mut command, &stdout_path)?
            }
        };
        let out_size = out_path.is_file().then(|| std::fs::metadata(&out_path).map(|m| m.len()).unwrap_or(0)).unwrap_or(0);
        let _ = std::fs::remove_file(&out_path);
        let _ = std::fs::remove_file(&stdout_path);
        let mut sample = record.clone();
        if let Some(obj) = sample.as_object_mut() {
            obj.insert("rep".into(), json!(rep));
            obj.insert("peak_kb".into(), json!(peak_kb));
            obj.insert("out_bytes".into(), json!(out_size));
        }
        samples.push(sample);
    }
    for sample in &samples {
        use std::io::Write;
        let mut file = std::fs::OpenOptions::new().create(true).append(true).open(&results)?;
        let mut row = json!({
            "engine": engine, "workload": workload,
            "fixture": fixture.file_name().unwrap().to_string_lossy(),
        });
        if let (Some(base), Some(extra)) = (row.as_object_mut(), sample.as_object()) {
            for (k, v) in extra {
                base.insert(k.clone(), v.clone());
            }
        }
        writeln!(file, "{}", serde_json::to_string(&row)?)?;
    }
    println!("perf-single {engine}/{workload}: {} reps recorded", samples.len());
    Ok(())
}

// ------------------------------------------------- perf-batch

pub fn perf_batch(args: &[String]) -> Result<()> {
    let engine = parse_flag(args, "--engine").context("--engine")?.to_owned();
    let count: usize = parse_flag(args, "--count").context("--count")?.parse()?;
    let mode = parse_flag(args, "--mode").context("--mode")?.to_owned();
    let reps: usize = parse_flag(args, "--reps").and_then(|v| v.parse().ok()).unwrap_or(5);
    let results = PathBuf::from(parse_flag(args, "--results").context("--results")?);
    let fixture = PathBuf::from(parse_flag(args, "--fixture").unwrap_or("fixtures/perf_small_1k.xlsx"));
    std::fs::create_dir_all(results.parent().unwrap())?;

    let workload = "w100";
    let ops = ops_for_workload(workload, &fixture);
    let ops_json = results.parent().unwrap().join("ops_batch.json");
    std::fs::write(&ops_json, serde_json::to_vec(&ops_to_json(&ops))?)?;

    let mut samples = Vec::new();
    for rep in 0..reps {
        let stamp = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH)?.as_nanos();
        let work = results.parent().unwrap().join(format!("batch-{stamp}"));
        std::fs::create_dir_all(&work)?;
        let mut files = Vec::new();
        for i in 0..count {
            let path = work.join(format!("book-{i}.xlsx"));
            std::fs::copy(&fixture, &path)?;
            files.push(path.to_string_lossy().into_owned());
        }
        let started = Instant::now();
        let (rows, peak_kb): (Vec<Value>, u64) = match (engine.as_str(), mode.as_str()) {
            ("openpyxl-native", "resident") => {
                let files_path = work.join("files.json");
                std::fs::write(&files_path, serde_json::to_vec(&files)?)?;
                let stdout_path = work.join("out.json");
                let mut command = std::process::Command::new(python_exe());
                command
                    .arg("-I")
                    .arg(PathBuf::from("scripts/py_engine.py").canonicalize()?)
                    .arg("bench-batch")
                    .arg(&ops_json.canonicalize()?)
                    .arg(files_path.canonicalize()?);
                let (value, peak) = run_measured(&mut command, &stdout_path)?;
                let rows = value
                    .get("rows")
                    .and_then(Value::as_array)
                    .cloned()
                    .unwrap_or_default();
                (rows, peak)
            }
            ("openpyxl-native", "cold") | ("openpyxl-bridge", "cold") => {
                let mut rows = Vec::new();
                let mut peak = 0u64;
                for file in &files {
                    let out_path = Path::new(file).with_extension("out.xlsx");
                    let (row, peak_kb) = run_one_cold(&engine, file, &out_path, &ops)?;
                    peak = peak.max(peak_kb);
                    rows.push(row);
                    let _ = std::fs::remove_file(&out_path);
                }
                (rows, peak)
            }
            ("openpyxl-bridge", "resident") => bail!(
                "production bridge is per-request by design; resident-bridge has no counterpart"
            ),
            (_, "cold") => {
                let mut rows = Vec::new();
                let mut peak = 0u64;
                for file in &files {
                    let out_path = Path::new(file).with_extension("out.xlsx");
                    let exe = std::env::current_exe()?;
                    let stdout_path = work.join(format!("one-{}.json", rows.len()));
                    let mut command = std::process::Command::new(exe);
                    command
                        .arg("worker")
                        .arg("single")
                        .arg(&engine)
                        .arg(file)
                        .arg(ops_json.canonicalize()?)
                        .arg(&canon_future(out_path)?);
                    let (record, peak_kb) = run_measured(&mut command, &stdout_path)?;
                    peak = peak.max(peak_kb);
                    rows.push(record);
                    let _ = std::fs::remove_file(&out_path);
                    let _ = std::fs::remove_file(&stdout_path);
                }
                (rows, peak)
            }
            (_, "resident") => {
                let files_path = work.join("files.json");
                std::fs::write(&files_path, serde_json::to_vec(&files)?)?;
                let stdout_path = work.join("out.json");
                let exe = std::env::current_exe()?;
                let mut command = std::process::Command::new(exe);
                command
                    .arg("worker")
                    .arg("batch")
                    .arg(&engine)
                    .arg(ops_json.canonicalize()?)
                    .arg(files_path.canonicalize()?);
                let (value, peak) = run_measured(&mut command, &stdout_path)?;
                let rows = value
                    .get("rows")
                    .and_then(Value::as_array)
                    .cloned()
                    .unwrap_or_default();
                (rows, peak)
            }
            _ => bail!("unknown engine/mode {engine}/{mode}"),
        };
        let wall_ms = started.elapsed().as_secs_f64() * 1000.0;
        let _ = std::fs::remove_dir_all(&work);
        samples.push(json!({
            "rep": rep,
            "wall_ms": wall_ms,
            "peak_kb": peak_kb,
            "files": count,
            "rows": rows,
        }));
    }
    for sample in &samples {
        use std::io::Write;
        let mut file = std::fs::OpenOptions::new().create(true).append(true).open(&results)?;
        let mut row = json!({
            "engine": engine, "mode": mode, "kind": "batch",
            "fixture": fixture.file_name().unwrap().to_string_lossy(),
        });
        if let (Some(base), Some(extra)) = (row.as_object_mut(), sample.as_object()) {
            for (k, v) in extra {
                base.insert(k.clone(), v.clone());
            }
        }
        writeln!(file, "{}", serde_json::to_string(&row)?)?;
    }
    println!("perf-batch {engine}/{mode}/{count}: {} reps recorded", samples.len());
    Ok(())
}

/// canonicalize 对尚不存在的目标文件会报错；规范化其父目录后拼回文件名。
fn canon_future(path: &Path) -> Result<PathBuf> {
    let parent = path.parent().unwrap_or(Path::new(".")).canonicalize()?;
    Ok(parent.join(path.file_name().unwrap_or(path.as_os_str())))
}

fn run_one_cold(
    engine: &str,
    file: &str,
    out_path: &Path,
    ops: &[Op],
) -> Result<(Value, u64)> {
    let stamp = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH)?.as_nanos();
    let started = Instant::now();
    match engine {
        "openpyxl-native" => {
            let stdout_path = PathBuf::from(format!(".cold-{stamp}.json"));
            let mut command = std::process::Command::new(python_exe());
            command
                .arg("-I")
                .arg(PathBuf::from("scripts/py_engine.py").canonicalize()?)
                .arg("bench")
                .arg(file)
                .arg(out_path)
                .arg("w100");
            let (record, peak) = run_measured(&mut command, &stdout_path)?;
            let _ = std::fs::remove_file(&stdout_path);
            Ok((record, peak))
        }
        "openpyxl-bridge" => {
            let request = json!({
                "schema_version": 1,
                "request_id": format!("bench-{stamp}"),
                "format": "xlsx",
                "input": file,
                "output": out_path.to_string_lossy(),
                "plan": {
                    "workbook": {"inventory_complete": true},
                    "operations": ops_to_json(ops)["ops"],
                },
            });
            let request_path = PathBuf::from(format!(".cold-req-{stamp}.json"));
            let response_path = PathBuf::from(format!(".cold-resp-{stamp}.json"));
            std::fs::write(&request_path, serde_json::to_vec(&request)?)?;
            let worker = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../crates/excel-editor/python/openpyxl_worker.py").canonicalize()?;
            let mut command = std::process::Command::new(python_exe());
            command
                .arg("-I")
                .arg(&worker)
                .arg("--request")
                .arg(request_path.canonicalize()?)
                .arg("--response")
                .arg(response_path.canonicalize()?);
            let peak = {
                #[cfg(windows)]
                {
                    let child = JobChild::spawn(&mut command)?;
                    let (status, peak) = child.wait_with_peak_kb()?;
                    if !status.success() {
                        bail!("bridge worker failed");
                    }
                    peak
                }
                #[cfg(not(windows))]
                {
                    command.output()?;
                    0
                }
            };
            let total_ms = started.elapsed().as_secs_f64() * 1000.0;
            let _ = std::fs::remove_file(&request_path);
            let _ = std::fs::remove_file(&response_path);
            Ok((json!({"total_ms": total_ms, "ok": true}), peak))
        }
        other => bail!("unexpected cold engine {other}"),
    }
}
