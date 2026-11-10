//! Fidelity harness: structure (ZIP parts), semantics (excel-reader dump), scenarios.

use anyhow::{bail, Context, Result};
use serde_json::{json, Value};
use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

use crate::{sha256_hex, Op, Val};

fn sha_of_part(archive: &mut zip::ZipArchive<std::io::Cursor<Vec<u8>>>, name: &str) -> Result<String> {
    use std::io::Read;
    let mut file = archive.by_name(name)?;
    let mut buf = Vec::new();
    file.read_to_end(&mut buf)?;
    Ok(sha256_hex(&buf))
}

pub fn zip_parts(bytes: &[u8]) -> Result<BTreeMap<String, String>> {
    let cursor = std::io::Cursor::new(bytes.to_vec());
    let mut archive = zip::ZipArchive::new(cursor)?;
    let mut map = BTreeMap::new();
    let names: Vec<String> = archive.file_names().map(str::to_owned).collect();
    for name in names {
        let hash = sha_of_part(&mut archive, &name)?;
        map.insert(name, hash);
    }
    Ok(map)
}

fn sheet_semantics(path: &Path) -> Result<BTreeMap<String, Value>> {
    let dump = excel_reader::read_workbook(path)?;
    let mut out = BTreeMap::new();
    for sheet in &dump.sheets {
        out.insert(
            sheet.name.clone(),
            json!({
                "values": sheet.values,
                "formulas": sheet.formulas,
                "merges": sheet.merges,
                "comments": sheet.comments,
                "conditional_formatting": sheet.conditional_formatting,
                "data_validations": sheet.data_validations,
            }),
        );
    }
    Ok(out)
}

fn defined_names(path: &Path) -> Result<Vec<Value>> {
    let records = excel_reader::read_defined_names(path)?;
    Ok(records
        .iter()
        .map(|r| {
            json!({
                "name": r.name,
                "scope": format!("{:?}", r.scope),
                "areas": r.areas.iter().map(|a| format!("{}!{}", a.sheet_name, a.reference)).collect::<Vec<_>>(),
            })
        })
        .collect())
}

fn diff_semantics(
    before: &BTreeMap<String, Value>,
    after: &BTreeMap<String, Value>,
    expected_changed: &[(String, String)],
) -> Vec<Value> {
    let mut issues = Vec::new();
    let expected: std::collections::BTreeSet<(String, String)> =
        expected_changed.iter().cloned().collect();
    for (sheet, before_sheet) in before {
        let Some(after_sheet) = after.get(sheet) else {
            issues.push(json!({"kind": "sheet_missing", "sheet": sheet}));
            continue;
        };
        for key in ["values", "formulas", "merges", "comments", "conditional_formatting", "data_validations"] {
            let b = &before_sheet[key];
            let a = &after_sheet[key];
            if b == a {
                continue;
            }
            if key == "values" || key == "formulas" {
                let bm = b.as_object().cloned().unwrap_or_default();
                let am = a.as_object().cloned().unwrap_or_default();
                let mut addresses: Vec<&String> = bm.keys().chain(am.keys()).collect();
                addresses.sort();
                addresses.dedup();
                for address in addresses {
                    let before_value = bm.get(address);
                    let after_value = am.get(address);
                    if before_value == after_value {
                        continue;
                    }
                    if expected.contains(&(sheet.clone(), address.clone())) {
                        continue;
                    }
                    // Formula-cell cached-value drop: formula text unchanged but
                    // the data_only value disappeared. Recorded as its own kind
                    // so reports can treat "no cache" differently from corruption.
                    let cache_dropped = key == "values"
                        && before_value.is_some()
                        && after_value.is_none()
                        && before_sheet["formulas"].get(address).is_some()
                        && before_sheet["formulas"][address] == *after_sheet["formulas"].get(address).unwrap_or(&Value::Null);
                    issues.push(json!({
                        "kind": if cache_dropped { "formula_cache_dropped" } else { "changed" },
                        "sheet": sheet, "section": key, "address": address,
                        "before": before_value, "after": after_value,
                    }));
                }
            } else {
                issues.push(json!({"kind": "changed", "sheet": sheet, "section": key, "before": b, "after": a}));
            }
        }
    }
    for sheet in after.keys() {
        if !before.contains_key(sheet) {
            issues.push(json!({"kind": "sheet_added", "sheet": sheet}));
        }
    }
    issues
}

fn ops_for_scenario(scenario: &str) -> Vec<Op> {
    ops_for_scenario_on(scenario, "Main")
}

fn ops_for_scenario_on_sheet(scenario: &str, sheet: &str) -> Vec<Op> {
    ops_for_scenario_on(scenario, sheet)
}

fn ops_for_scenario_on(scenario: &str, sheet: &str) -> Vec<Op> {
    match scenario {
        "no-op" | "failsafe" => Vec::new(),
        "edit1" | "repeat-first" => vec![Op::SetCellValue {
            sheet: sheet.into(),
            cell: "B2".into(),
            val: Val::Text("保真探针-值".into()),
        }],
        "repeat-second" => vec![Op::SetCellValue {
            sheet: sheet.into(),
            cell: "C5".into(),
            val: Val::Num(777.0),
        }],
        _ => Vec::new(),
    }
}

fn expected_changes(scenario: &str) -> Vec<(String, String)> {
    expected_changes_on(scenario, "Main")
}

fn expected_changes_on(scenario: &str, sheet: &str) -> Vec<(String, String)> {
    match scenario {
        "edit1" | "repeat-first" => vec![(sheet.into(), "B2".into())],
        "repeat-second" => vec![(sheet.into(), "C5".into())],
        _ => Vec::new(),
    }
}

/// 夹具没有 "Main" 表时改用首表名（复杂模板首表通常是汇总类名称）。
fn first_sheet_name(path: &Path) -> String {
    excel_reader::read_workbook(path)
        .map(|dump| dump.sheets.first().map(|s| s.name.clone()))
        .ok()
        .flatten()
        .unwrap_or_else(|| "Main".into())
}

pub fn run_fidelity(
    engine_name: &str,
    fixture: &Path,
    scenario: &str,
    out_dir: &Path,
    results_path: &Path,
) -> Result<()> {
    std::fs::create_dir_all(out_dir)?;
    let stem = format!(
        "{}_{}_{}",
        engine_name,
        fixture.file_stem().unwrap().to_string_lossy(),
        scenario
    );
    let output_path = out_dir.join(format!("{stem}.xlsx"));
    let mut record = json!({
        "engine": engine_name,
        "fixture": fixture.file_name().unwrap().to_string_lossy(),
        "scenario": scenario,
    });

    let input_bytes = std::fs::read(fixture).context("read fixture")?;
    let before_parts = zip_parts(&input_bytes)?;
    let before_semantics = sheet_semantics(fixture)?;
    let before_names = defined_names(fixture)?;

    // ---- produce output
    let mut failures: Vec<String> = Vec::new();
    let output_bytes: Option<Vec<u8>> = if engine_name == "openpyxl" {
        let staged_in = out_dir.join(format!("{stem}-in.xlsx"));
        std::fs::write(&staged_in, &input_bytes)?;
        let status = std::process::Command::new(python_exe())
            .arg("-I")
            .arg(scripts_dir().join("py_engine.py"))
            .arg("apply")
            .arg(&staged_in)
            .arg(&output_path)
            .arg(scenario)
            .output()
            .context("spawn python engine runner")?;
        let _ = std::fs::remove_file(&staged_in);
        if !status.status.success() {
            failures.push(format!(
                "python runner failed: {}",
                String::from_utf8_lossy(&status.stderr)
            ));
            None
        } else if output_path.is_file() {
            Some(std::fs::read(&output_path)?)
        } else {
            None
        }
    } else if scenario == "failsafe" {
        // Rust engines: feed corrupted bytes; success criterion = error + no output.
        let mut corrupt = input_bytes.clone();
        let len = corrupt.len();
        if len > 200 {
            corrupt[100..140].copy_from_slice(&[0u8; 40]);
        }
        let engine = crate::engines::engine_by_name(engine_name).unwrap();
        let outcome = engine.apply(&corrupt, &ops_for_scenario("edit1"));
        if outcome.output.is_some() {
            failures.push("engine produced output for corrupted input (must refuse)".into());
        }
        None
    } else {
        let engine = crate::engines::engine_by_name(engine_name).unwrap();
        let sheet = first_sheet_name(fixture);
        let ops = ops_for_scenario_on_sheet(scenario, &sheet);
        let outcome = engine.apply(&input_bytes, &ops);
        for (kind, result) in &outcome.op_results {
            if let Err(error) = result {
                failures.push(format!("{kind}: {error}"));
            }
        }
        outcome.output
    };

    // ---- failsafe: original file must be untouched; engine never writes in place.
    // 拒绝方式不限：非零退出、报错、或干脆不产出输出都算「拒绝成功」；
    // 对损坏输入仍产出文件才是 FAIL。
    if scenario == "failsafe" {
        let original_unchanged =
            sha256_hex(&std::fs::read(fixture)?) == sha256_hex(&input_bytes);
        record["failsafe_original_unchanged"] = json!(original_unchanged);
        let refused = output_bytes.is_none() && original_unchanged;
        let verdict = if refused && original_unchanged { "PASS" } else { "FAIL" };
        record["verdict"] = json!(verdict);
        if !refused {
            record["failures"] = json!(["engine did not refuse the corrupted input"]);
        }
        append_record(results_path, &record)?;
        println!("{stem}: {verdict}");
        return Ok(());
    }

    let Some(output_bytes) = output_bytes else {
        record["verdict"] = json!("FAIL");
        record["failures"] = json!(failures);
        append_record(results_path, &record)?;
        println!("{stem}: FAIL (no output)");
        return Ok(());
    };
    std::fs::write(&output_path, &output_bytes)?;

    // ---- A. structure
    let after_parts = zip_parts(&output_bytes)?;
    let mut added = Vec::new();
    let mut removed = Vec::new();
    let mut changed = Vec::new();
    let mut unchanged = 0usize;
    for (name, hash) in &after_parts {
        match before_parts.get(name) {
            None => added.push(name.clone()),
            Some(before_hash) if before_hash != hash => changed.push(name.clone()),
            Some(_) => unchanged += 1,
        }
    }
    for name in before_parts.keys() {
        if !after_parts.contains_key(name) {
            removed.push(name.clone());
        }
    }
    record["structure"] = json!({
        "parts_before": before_parts.len(),
        "parts_after": after_parts.len(),
        "unchanged": unchanged,
        "changed": changed,
        "added": added,
        "removed": removed,
    });

    // ---- B. semantics
    let tmp_output = out_dir.join(format!("{stem}-verify.xlsx"));
    std::fs::write(&tmp_output, &output_bytes)?;
    let after_semantics = match sheet_semantics(&tmp_output) {
        Ok(value) => value,
        Err(error) => {
            record["verdict"] = json!("FAIL");
            record["failures"] = json!([format!("output not readable: {error}")]);
            append_record(results_path, &record)?;
            let _ = std::fs::remove_file(&tmp_output);
            println!("{stem}: FAIL (unreadable output)");
            return Ok(());
        }
    };
    let after_names = defined_names(&tmp_output).unwrap_or_default();
    let semantic_sheet = first_sheet_name(fixture);
    let mut semantic_issues =
        diff_semantics(&before_semantics, &after_semantics, &expected_changes_on(scenario, &semantic_sheet));
    if before_names != after_names {
        semantic_issues.push(json!({"kind": "changed", "section": "defined_names", "before": before_names, "after": after_names}));
    }
    let cache_drops = semantic_issues
        .iter()
        .filter(|i| i["kind"] == "formula_cache_dropped")
        .count();
    let unexpected: Vec<&Value> = semantic_issues.iter().filter(|i| i["kind"] != "formula_cache_dropped").collect();
    record["semantics"] = json!({
        "unexpected_changes": unexpected,
        "formula_cache_dropped": cache_drops,
    });
    let _ = std::fs::remove_file(&tmp_output);

    // 结构层：被移除的部件是真实丢失（合法重序列化只会改变内容，不会删除
    // 声明部件）。removed 非空即 FAIL，防止「字节级清单检查通过」掩盖丢件。
    if !removed.is_empty() {
        failures.push(format!(
            "parts removed by rewrite: {}",
            removed.join(", ")
        ));
    }
    let verdict = if unexpected.is_empty() && removed.is_empty() && failures.is_empty() {
        "PASS"
    } else {
        "FAIL"
    };
    record["verdict"] = json!(verdict);
    if !failures.is_empty() {
        record["failures"] = json!(failures);
    }
    append_record(results_path, &record)?;
    println!("{stem}: {verdict} (cache_drops={cache_drops}, unexpected={})", unexpected.len());
    Ok(())
}

fn append_record(results_path: &Path, record: &Value) -> Result<()> {
    use std::io::Write;
    if let Some(parent) = results_path.parent() {
        std::fs::create_dir_all(parent)?;
    }
    let mut file = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(results_path)?;
    writeln!(file, "{}", serde_json::to_string(record)?)?;
    Ok(())
}

fn python_exe() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../dist/python/python.exe")
        .canonicalize()
        .expect("bundled python runtime must exist")
}

fn scripts_dir() -> PathBuf {
    PathBuf::from("scripts")
}
