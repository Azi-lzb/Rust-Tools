use std::{
    env,
    io::{BufRead, BufReader, Write},
    process::{Command, Stdio},
    time::Instant,
};

fn send_line<W: Write>(w: &mut W, body: &str) -> Result<(), String> {
    w.write_all(body.as_bytes()).map_err(|e| e.to_string())?;
    w.write_all(b"\n").map_err(|e| e.to_string())?;
    w.flush().map_err(|e| e.to_string())
}

fn read_ok<R: BufRead>(r: &mut R) -> Result<String, String> {
    let mut line = String::new();
    r.read_line(&mut line).map_err(|e| e.to_string())?;
    if line.is_empty() {
        return Err("worker exited without a response".into());
    }
    if !line.contains("\"ok\":true") {
        return Err(format!("worker error: {}", line.trim()));
    }
    Ok(line)
}

fn worker_ms(reply: &str) -> f64 {
    let key = "\"worker_total_ms\":";
    reply
        .find(key)
        .and_then(|i| {
            let rest = &reply[i + key.len()..];
            let end = rest.find(|c: char| !(c.is_ascii_digit() || c == '.'))?;
            rest[..end].parse().ok()
        })
        .unwrap_or(f64::NAN)
}

fn main() -> Result<(), String> {
    let args: Vec<String> = env::args().collect();
    if args.len() != 4 {
        return Err("usage: aspose-dotnet-worker-bridge WORKER INPUT.xlsx OUTPUT_DIR".into());
    }
    let output_dir = &args[3];
    std::fs::create_dir_all(output_dir).map_err(|e| e.to_string())?;
    let start = Instant::now();
    let mut child = Command::new(&args[1])
        .arg("--worker")
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::inherit())
        .spawn()
        .map_err(|e| format!("start worker: {e}"))?;
    let startup_spawn_ms = start.elapsed().as_secs_f64() * 1000.0;
    let mut input = child.stdin.take().ok_or("missing worker stdin")?;
    let stdout = child.stdout.take().ok_or("missing worker stdout")?;
    let mut output = BufReader::new(stdout);
    let handshake = Instant::now();
    send_line(&mut input, r#"{"request_id":"hello","op":"ping"}"#)?;
    read_ok(&mut output)?;
    let startup_ready_ms = handshake.elapsed().as_secs_f64() * 1000.0;
    println!(
        "WORKER_READY spawn_ms={startup_spawn_ms:.3} ready_roundtrip_ms={startup_ready_ms:.3}"
    );

    let mut csv = String::from("rows,payload_bytes,serialization_ms,roundtrip_ms,end_to_end_ms,worker_ms,bridge_overhead_ms,output_bytes,status\n");
    for rows in [10_000usize, 100_000usize] {
        let end_to_end = Instant::now();
        let id = format!("append-{rows}");
        let target = format!("{output_dir}/aspose-worker-{rows}.xlsx");
        let mut request = format!("{{\"request_id\":\"{id}\",\"input_path\":\"{}\",\"output_path\":\"{target}\",\"append_rows\":[", args[2]);
        for i in 0..rows {
            if i != 0 {
                request.push(',');
            }
            request.push_str(&format!("[\"batch-{i}\",{}]", i as f64));
        }
        request.push_str("]}");
        let payload_bytes = request.len();
        let serialization_ms = end_to_end.elapsed().as_secs_f64() * 1000.0;
        let task = Instant::now();
        send_line(&mut input, &request)?;
        let reply = read_ok(&mut output)?;
        let roundtrip_ms = task.elapsed().as_secs_f64() * 1000.0;
        let end_to_end_ms = end_to_end.elapsed().as_secs_f64() * 1000.0;
        let engine_ms = worker_ms(&reply);
        let output_bytes = std::fs::metadata(&target).map_err(|e| e.to_string())?.len();
        let bridge_ms = end_to_end_ms - engine_ms;
        csv.push_str(&format!("{rows},{payload_bytes},{serialization_ms:.3},{roundtrip_ms:.3},{end_to_end_ms:.3},{engine_ms:.3},{bridge_ms:.3},{output_bytes},PASS\n"));
        println!("BATCH rows={rows} bytes={payload_bytes} serialize_ms={serialization_ms:.3} roundtrip_ms={roundtrip_ms:.3} e2e_ms={end_to_end_ms:.3} worker_ms={engine_ms:.3} bridge_ms={bridge_ms:.3} output_bytes={output_bytes}");
    }
    drop(input);
    let status = child.wait().map_err(|e| e.to_string())?;
    if !status.success() {
        return Err(format!("worker exited with {status}"));
    }
    std::fs::write(format!("{output_dir}/bridge-aspose-dotnet.csv"), csv)
        .map_err(|e| e.to_string())?;
    Ok(())
}
