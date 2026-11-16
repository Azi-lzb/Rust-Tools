use calamine::{open_workbook_auto, Data, Reader};
use rust_xlsxwriter::Workbook;
use serde_json::{json, Value};
use sha2::{Digest, Sha256};
use std::{
    collections::HashMap,
    error::Error,
    io::{self, BufRead},
    time::Instant,
};
type Matrix = Vec<Vec<Value>>;
type Result<T> = std::result::Result<T, Box<dyn Error>>;

fn digest(data: &Matrix) -> String {
    let mut h = Sha256::new();
    for row in data {
        for v in row {
            let s = match v {
                Value::Null => "E\n".into(),
                Value::Bool(b) => format!("B{}\n", u8::from(*b)),
                Value::Number(n) => format!("N{}\n", n),
                Value::String(s) => format!("S{s}\n"),
                _ => panic!("unsupported fixture type"),
            };
            h.update(s.as_bytes());
        }
    }
    format!("{:x}", h.finalize())
}
fn read(path: &str, rows: usize, cols: usize) -> Result<Matrix> {
    let mut wb = open_workbook_auto(path)?;
    let range = wb.worksheet_range("Data")?;
    let mut out = Vec::with_capacity(rows);
    for r in 0..rows {
        let mut row = Vec::with_capacity(cols);
        for c in 0..cols {
            row.push(
                match range
                    .get_value((r as u32, c as u32))
                    .unwrap_or(&Data::Empty)
                {
                    Data::Empty => Value::Null,
                    Data::String(s) => json!(s),
                    Data::Int(n) => json!(n),
                    Data::Float(n) if n.fract() == 0.0 => json!(*n as i64),
                    Data::Bool(b) => json!(b),
                    other => {
                        return Err(format!("unsupported fixture cell {r},{c}: {other:?}").into())
                    }
                },
            );
        }
        out.push(row);
    }
    Ok(out)
}
fn write(path: &str, data: &Matrix) -> Result<()> {
    if std::path::Path::new(path).exists() {
        return Err("refuse existing output".into());
    }
    let mut wb = Workbook::new();
    let ws = wb.add_worksheet();
    ws.set_name("Data")?;
    for (r, row) in data.iter().enumerate() {
        for (c, v) in row.iter().enumerate() {
            match v {
                Value::Null => {}
                Value::Bool(b) => {
                    ws.write_boolean(r as u32, c as u16, *b)?;
                }
                Value::Number(n) => {
                    ws.write_number(r as u32, c as u16, n.as_f64().ok_or("number")?)?;
                }
                Value::String(s) => {
                    ws.write_string(r as u32, c as u16, s)?;
                }
                _ => return Err("fixture type".into()),
            }
        }
    }
    wb.save(path)?;
    Ok(())
}
fn run(q: &Value, sets: &mut HashMap<String, Matrix>) -> Result<Value> {
    let op = q["op"].as_str().ok_or("op")?;
    if op == "init" {
        for x in q["fixtures"].as_array().ok_or("fixtures")? {
            sets.insert(
                x["id"].as_str().ok_or("id")?.into(),
                serde_json::from_reader(std::fs::File::open(x["data"].as_str().ok_or("data")?)?)?,
            );
        }
        return Ok(json!({"ok":true}));
    }
    let path = q["input"].as_str().unwrap_or("");
    let output = q["output"].as_str().unwrap_or("");
    let data = sets.get(q["id"].as_str().ok_or("id")?).ok_or("dataset")?;
    let mut owned = None;
    let start = Instant::now();
    match op {
        "read" => owned = Some(read(path, data.len(), data[0].len())?),
        "write" => write(output, data)?,
        "roundtrip" => {
            let d = read(path, data.len(), data[0].len())?;
            write(output, &d)?;
            owned = Some(d);
        }
        _ => return Err("unknown operation".into()),
    }
    let ms = start.elapsed().as_secs_f64() * 1000.0;
    let h = digest(owned.as_ref().unwrap_or(data));
    Ok(json!({"ok":true,"processing_ms":ms,"digest":h,"rows":data.len(),"cols":data[0].len()}))
}
fn main() {
    println!(
        "{}",
        json!({"ready":true,"engine":"rust","calamine":"0.37.0","rust_xlsxwriter":"0.101.0","pid":std::process::id()})
    );
    let mut sets = HashMap::new();
    for line in io::stdin().lock().lines() {
        let response = (|| -> Result<Value> { run(&serde_json::from_str(&line?)?, &mut sets) })();
        println!(
            "{}",
            match response {
                Ok(v) => v,
                Err(e) => json!({"ok":false,"error":e.to_string()}),
            }
        );
    }
}
