use std::{env, fs, path::PathBuf, process::Command, time::Instant};

fn json_escape(value: &str) -> String {
    let mut out = String::with_capacity(value.len() + 8);
    for c in value.chars() {
        match c {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            c if c < ' ' => out.push_str(&format!("\\u{:04x}", c as u32)),
            c => out.push(c),
        }
    }
    out
}

fn number_field(json: &str, name: &str) -> f64 {
    let key = format!("\"{name}\":");
    json.find(&key)
        .and_then(|i| {
            let rest = &json[i + key.len()..];
            let end = rest.find(|c: char| !(c.is_ascii_digit() || c == '.'))?;
            rest[..end].parse().ok()
        })
        .unwrap_or(f64::NAN)
}

fn main() -> Result<(), String> {
    let a: Vec<String> = env::args().collect();
    if a.len() != 5 {
        return Err("usage: openpyxl-worker-bridge PYTHON WORKER.py INPUT.xlsx OUTPUT_DIR".into());
    }
    let python = &a[1];
    let worker = &a[2];
    let input = &a[3];
    let dir = PathBuf::from(&a[4]);
    fs::create_dir_all(&dir).map_err(|e| e.to_string())?;
    let mut csv = String::from("rows,payload_bytes,serialize_and_file_ms,process_roundtrip_ms,end_to_end_ms,worker_ms,output_bytes,status\n");
    for n in [10_000usize, 100_000usize] {
        let e2e = Instant::now();
        let output = dir.join(format!("openpyxl-worker-{n}.xlsx"));
        let request_path = dir.join(format!("request-{n}.json"));
        let response_path = dir.join(format!("response-{n}.json"));
        let mut json = format!("{{\"request_id\":\"openpyxl-{n}\",\"input_path\":\"{}\",\"output_path\":\"{}\",\"sheet_name\":\"Data\",\"append_rows\":[", json_escape(input), json_escape(&output.to_string_lossy()));
        for i in 0..n {
            if i != 0 {
                json.push(',');
            }
            json.push_str(&format!("[\"batch-{i}\",{i}]"));
        }
        json.push_str("]}");
        let payload_bytes = json.len();
        fs::write(&request_path, json).map_err(|e| e.to_string())?;
        let ipc_file_ms = e2e.elapsed().as_secs_f64() * 1000.0;
        let process = Command::new(python)
            .arg("-I")
            .arg(worker)
            .arg(&request_path)
            .arg(&response_path)
            .output()
            .map_err(|e| format!("spawn Python worker: {e}"))?;
        let process_ms = e2e.elapsed().as_secs_f64() * 1000.0 - ipc_file_ms;
        let response = fs::read_to_string(&response_path).map_err(|e| e.to_string())?;
        if !process.status.success() || !response.contains("\"ok\":true") {
            return Err(format!(
                "worker failure stdout={} stderr={} response={response}",
                String::from_utf8_lossy(&process.stdout),
                String::from_utf8_lossy(&process.stderr)
            ));
        }
        let e2e_ms = e2e.elapsed().as_secs_f64() * 1000.0;
        let worker_ms = number_field(&response, "worker_total_ms");
        let bytes = fs::metadata(&output).map_err(|e| e.to_string())?.len();
        csv.push_str(&format!("{n},{payload_bytes},{ipc_file_ms:.3},{process_ms:.3},{e2e_ms:.3},{worker_ms:.3},{bytes},PASS\n"));
        println!("PYTHON_BATCH rows={n} bytes={payload_bytes} request_file_ms={ipc_file_ms:.3} process_roundtrip_ms={process_ms:.3} e2e_ms={e2e_ms:.3} worker_ms={worker_ms:.3} output_bytes={bytes}");
        let _ = fs::remove_file(request_path);
        let _ = fs::remove_file(response_path);
    }
    fs::write(dir.join("bridge-openpyxl.csv"), csv).map_err(|e| e.to_string())?;
    Ok(())
}
