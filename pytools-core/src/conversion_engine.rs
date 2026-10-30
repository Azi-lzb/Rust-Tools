//! Fidelity conversion engine (former `lo-convert` sidecar).
//! Windows spreadsheet conversion uses Rust COM on an isolated STA thread;
//! Word retains its PowerShell COM bridge. LibreOffice uses the local
//! `soffice` CLI. No separate `lo-convert` executable is required.

use anyhow::{anyhow, bail, ensure, Context, Result};
use encoding_rs::GBK;
use excel_reader::{
    read_value_columns_with_iso_dates, read_workbook, CellValue as ReaderCellValue, WorkbookDump,
};
use std::{
    fs,
    path::{Path, PathBuf},
    process::{Command, Stdio},
    time::{SystemTime, UNIX_EPOCH},
};

/// State scoped to one user conversion request or batch.
#[derive(Default)]
pub(crate) struct ConversionRuntime {
    #[cfg(windows)]
    spreadsheet_com: Option<crate::conversion_com::ComWorker>,
}

impl ConversionRuntime {
    #[cfg(windows)]
    fn spreadsheet_com(
        &mut self,
        input: &Path,
        output: &Path,
        source_ext: &str,
    ) -> Result<ComExecution> {
        if self.spreadsheet_com.is_none() {
            self.spreadsheet_com = Some(crate::conversion_com::ComWorker::start()?);
        }
        let converted = self.spreadsheet_com
            .as_mut()
            .expect("spreadsheet COM worker initialized")
            .convert(input, output, source_ext)?;
        Ok(ComExecution {
            source_format: converted.source_format,
            bridge: "rust",
            application: Some(converted.application),
            process_id: Some(converted.process_id),
        })
    }
}

struct ComExecution {
    source_format: Option<i32>,
    bridge: &'static str,
    application: Option<&'static str>,
    process_id: Option<u32>,
}

/// Run one same-category fidelity conversion and retain the existing JSON
/// detail contract. The caller may reuse the runtime across a batch.
pub(crate) fn run_fidelity_conversion_with_runtime(
    input: &Path,
    output: &Path,
    target: &str,
    engine: &str,
    runtime: &mut ConversionRuntime,
) -> Result<String> {
    if std::env::var_os("PYTOOLS_CONVERT_ENGINES_DISABLED").is_some() {
        // Test override: simulate an environment with no external engine.
        bail!("该功能需要外部引擎: no conversion engine is available (PYTOOLS_CONVERT_ENGINES_DISABLED)");
    }
    ensure!(
        engine == "auto" || engine == "lo" || engine == "com",
        "engine must be auto, lo or com"
    );
    ensure!(
        input.is_file(),
        "input file does not exist: {}",
        input.display()
    );
    let source_ext = extension(input)?;
    ensure!(
        (matches!(target, "xlsx")
            && matches!(source_ext.as_str(), "xls" | "xlsb" | "xlsm" | "et"))
            || (target == "docx" && source_ext == "doc"),
        "unsupported conversion: .{source_ext} -> .{target}"
    );
    if let Some(parent) = output.parent() {
        fs::create_dir_all(parent)?;
    }
    ensure!(
        !output.exists(),
        "refusing to overwrite existing output: {}",
        output.display()
    );

    let temp_guard = TempDirGuard(unique_temp_dir(
        output.parent().unwrap_or_else(|| Path::new(".")),
    )?);
    let temp_dir = &temp_guard.0;
    let converted = temp_dir.join(format!("{}.{}", file_stem(input), target));
    let selected = if engine == "auto" {
        if cfg!(windows) {
            "com"
        } else {
            "lo"
        }
    } else {
        engine
    };
    let com_execution = match selected {
        "lo" => {
            run_libreoffice(input, temp_dir, target)?;
            None
        }
        "com" if cfg!(windows) => Some(run_com(input, &converted, target, &source_ext, runtime)?),
        "com" => bail!(
            "该功能需要外部引擎: COM conversion is only available on Windows; install LibreOffice or use --engine lo"
        ),
        _ => unreachable!(),
    };
    let source_file_format = com_execution.as_ref().and_then(|execution| execution.source_format);
    let candidate = if converted.exists() {
        converted
    } else {
        locate_converted(temp_dir, target)?
    };
    ensure!(
        candidate.is_file() && fs::metadata(&candidate)?.len() > 0,
        "conversion engine did not create a non-empty output file"
    );
    let validation = validate_conversion(
        input,
        &candidate,
        &source_ext,
        target,
        selected,
        temp_dir,
        source_file_format,
        runtime,
    )?;
    let validation = serde_json::from_str::<serde_json::Value>(&validation)
        .context("decode conversion validation details")?;
    let mut report = serde_json::json!({
        "status": "converted",
        "engine": selected,
        "input": input.display().to_string(),
        "output": output.display().to_string(),
        "validation": validation,
    });
    if let Some(execution) = com_execution {
        report["com_bridge"] = serde_json::json!(execution.bridge);
        if let Some(application) = execution.application {
            report["com_application"] = serde_json::json!(application);
        }
        if let Some(process_id) = execution.process_id {
            report["com_process_id"] = serde_json::json!(process_id);
        }
    }
    fs::rename(&candidate, output)
        .with_context(|| format!("commit validated output {}", output.display()))?;
    Ok(report.to_string())
}

fn run_libreoffice(input: &Path, temp_dir: &Path, target: &str) -> Result<()> {
    let profile = temp_dir.join("lo-profile");
    let profile_argument = libreoffice_user_installation_arg(&profile)?;
    let mut command = Command::new("soffice");
    command
        .arg("--headless")
        .arg("--nologo")
        .arg("--nodefault")
        .arg("--nolockcheck")
        .arg(profile_argument)
        .arg("--convert-to")
        .arg(if target == "xlsx" {
            "xlsx:Calc MS Excel 2007 XML"
        } else {
            "docx:Office Open XML Text"
        })
        .arg("--outdir")
        .arg(temp_dir)
        .arg(input)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped());
    let output = command.output().with_context(|| {
        "该功能需要外部引擎: LibreOffice is unavailable; install soffice to use this conversion"
    })?;
    ensure!(
        output.status.success(),
        "LibreOffice conversion failed ({}): {}",
        output.status,
        String::from_utf8_lossy(&output.stderr)
    );
    Ok(())
}

fn run_com(
    input: &Path,
    output: &Path,
    target: &str,
    source_ext: &str,
    runtime: &mut ConversionRuntime,
) -> Result<ComExecution> {
    let output_parent = output
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    fs::create_dir_all(output_parent)?;
    let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
    let com_input = output_parent.join(format!(
        ".pytools-convert-source-{}-{nonce}.{source_ext}",
        std::process::id()
    ));
    // Some WPS COM failure paths can alter the workbook passed to Open even
    // when ReadOnly=true. External automation must never receive the caller's
    // source path; it only opens this disposable copy.
    fs::copy(input, &com_input).with_context(|| {
        format!(
            "stage a private source copy for Office/WPS COM conversion from {}",
            input.display()
        )
    })?;
    if target == "xlsx" {
        #[cfg(windows)]
        let result = runtime.spreadsheet_com(&com_input, output, source_ext);
        #[cfg(not(windows))]
        let result = {
            let _ = runtime;
            Err(anyhow!("COM conversion is only available on Windows"))
        };
        let _ = fs::remove_file(&com_input);
        return result;
    }
    let app_names: &[&str] = if target == "docx" {
        &[
            "Word.Application",
            "kwps.Application",
            "KWPS.Application",
            "wps.Application",
        ]
    } else if source_ext == "et" {
        // ET is a WPS-native workbook format. Do not route it through Excel
        // COM, which may open a renamed XLS and make the extension look valid.
        &["KET.Application"]
    } else {
        &["Excel.Application", "ket.Application", "KET.Application"]
    };
    let format_code = if target == "docx" { 16 } else { 51 };
    let mut script = String::from("$ErrorActionPreference='Stop';\n");
    script.push_str("$app=$null; $appName=''; $lastError='';\n");
    for app_name in app_names {
        script.push_str(&format!(
            "if ($null -eq $app) {{ try {{ $app=New-Object -ComObject '{}'; $appName='{}' }} catch {{ $lastError=$_.Exception.Message }} }}\n",
            app_name, app_name
        ));
    }
    script.push_str("if ($null -eq $app) { throw \"No supported Office/WPS COM engine is available: $lastError\" }\n");
    script.push_str("$doc=$null; $book=$null;\ntry {\n$app.Visible=$false; ");
    script.push_str(com_display_alerts_setting(target));
    script.push('\n');
    if target == "docx" {
        script.push_str(&format!("$doc=$app.Documents.Open('{}',$false,$true); try {{ $doc.SaveAs2('{}',{}) }} catch {{ $doc.SaveAs('{}',{}) }};\n", ps_quote(&com_input.to_string_lossy()), ps_quote(&output.to_string_lossy()), format_code, ps_quote(&output.to_string_lossy()), format_code));
    } else {
        script.push_str(&format!(
            "$book=$app.Workbooks.Open('{}',0,$true);\n",
            ps_quote(&com_input.to_string_lossy())
        ));
        if source_ext == "et" {
            // Inspect the format WPS actually opened before publishing an
            // XLSX. A CFB/XLS file renamed to .et reports 56; native ET is
            // identified by KET as 65521.
            script.push_str(
                "$sourceFormat=[int]$book.FileFormat; if ($sourceFormat -ne 65521) { throw \"Input has .et extension but KET opened FileFormat $sourceFormat; expected native ET FileFormat 65521\" }; [Console]::Out.WriteLine(\"PYTOOLS_SOURCE_FILE_FORMAT=$sourceFormat\");\n",
            );
        }
        script.push_str(&format!(
            "$book.SaveAs('{}',{});\n",
            ps_quote(&output.to_string_lossy()),
            format_code
        ));
    }
    script.push_str("} finally { if ($null -ne $doc) { $doc.Close($false) }; if ($null -ne $book) { $book.Close($false) }; if ($null -ne $app) { $app.Quit() } }\n");
    let script_path = output.with_extension("pytools-convert.ps1");
    let result = (|| {
        write_powershell_script(&script_path, &script)?;
        Command::new("powershell.exe")
            .args([
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
            ])
            .arg(&script_path)
            .stdout(Stdio::piped())
            .stderr(Stdio::piped())
            .output()
            .with_context(|| "start Office/WPS COM automation")
    })();
    let _ = fs::remove_file(script_path);
    let _ = fs::remove_file(com_input);
    let result = result?;
    if !result.status.success() {
        bail!(
            "Office/WPS COM conversion failed: {}",
            String::from_utf8_lossy(&result.stderr)
        );
    }
    if source_ext == "et" {
        let file_format = String::from_utf8_lossy(&result.stdout)
            .lines()
            .find_map(|line| line.trim().strip_prefix("PYTOOLS_SOURCE_FILE_FORMAT="))
            .context("KET did not return source FileFormat evidence for .et input")?
            .parse::<i32>()
            .context("KET returned invalid source FileFormat evidence")?;
        ensure!(
            file_format == 65521,
            "KET opened .et input as FileFormat {file_format}; expected native ET FileFormat 65521"
        );
        Ok(ComExecution {
            source_format: Some(file_format),
            bridge: "powershell",
            application: None,
            process_id: None,
        })
    } else {
        Ok(ComExecution {
            source_format: None,
            bridge: "powershell",
            application: None,
            process_id: None,
        })
    }
}

fn com_display_alerts_setting(target: &str) -> &'static str {
    if target == "docx" {
        // Word expects the WdAlertLevel enum (wdAlertsNone = 0), not a Boolean.
        "$app.DisplayAlerts=0;"
    } else {
        "$app.DisplayAlerts=$false;"
    }
}

fn write_powershell_script(path: &Path, script: &str) -> Result<()> {
    let mut encoded = Vec::with_capacity(script.len() + 3);
    encoded.extend_from_slice(&[0xEF, 0xBB, 0xBF]);
    encoded.extend_from_slice(script.as_bytes());
    fs::write(path, encoded).with_context(|| format!("write PowerShell script {}", path.display()))
}

fn validate_conversion(
    input: &Path,
    output: &Path,
    source_ext: &str,
    target: &str,
    engine: &str,
    temp_dir: &Path,
    source_file_format: Option<i32>,
    runtime: &mut ConversionRuntime,
) -> Result<String> {
    if target == "docx" {
        let bytes = fs::read(output)?;
        ensure!(
            bytes.starts_with(b"PK\x03\x04"),
            "converted DOCX is not a valid OOXML ZIP package"
        );
        let (source_text, output_text) = match engine {
            "lo" => (
                extract_word_text_lo(input, &temp_dir.join("source-text"))?,
                extract_word_text_lo(output, &temp_dir.join("output-text"))?,
            ),
            "com" if cfg!(windows) => (
                extract_word_text_com(input, &temp_dir.join("source.txt"))?,
                extract_word_text_com(output, &temp_dir.join("output.txt"))?,
            ),
            _ => bail!("unsupported text-validation engine {engine:?}"),
        };
        ensure!(
            normalize_document_text(&source_text) == normalize_document_text(&output_text),
            "converted DOCX text content differs from the source document"
        );
        return Ok("{\"level\":\"docx_package_and_text_content\"}".into());
    }
    let converted = read_comparable_workbook(output)
        .with_context(|| format!("validate converted workbook {}", output.display()))?;
    ensure!(
        !converted.sheets.is_empty(),
        "converted workbook has no worksheets"
    );
    if source_ext == "csv" {
        validate_csv_conversion(input, &converted)?;
        return Ok("{\"level\":\"one_worksheet_and_csv_values\",\"numeric_tolerance\":1e-15}".into());
    }
    let source_path = if source_ext == "et" {
        // ET is a CFB container, not necessarily an OOXML ZIP. Convert the
        // original source separately with the selected engine, then compare
        // that workbook's sheets, values and formulas with the candidate.
        // Never disguise the source by changing its extension and asking
        // Calamine to read it as XLSX.
        let probe_dir = temp_dir.join("source-probe");
        fs::create_dir_all(&probe_dir)?;
        match engine {
            "lo" => {
                run_libreoffice(input, &probe_dir, target)
                    .context("convert ET source for preservation validation with LibreOffice")?;
                locate_converted(&probe_dir, target)
                    .context("locate separately converted ET source for validation")?
            }
            "com" if cfg!(windows) => {
                let probe = probe_dir.join(format!("source.{target}"));
                let probe_file_format = run_com(input, &probe, target, source_ext, runtime)
                    .context("convert ET source for preservation validation with COM")?
                    .source_format;
                ensure!(
                    source_file_format == probe_file_format,
                    "ET source format verification differed between candidate and preservation probe"
                );
                ensure!(
                    probe.is_file() && fs::metadata(&probe)?.len() > 0,
                    "COM did not create a non-empty ET source validation workbook"
                );
                probe
            }
            _ => bail!("unsupported ET source-validation engine {engine:?}"),
        }
    } else {
        input.to_path_buf()
    };
    let source = read_comparable_workbook(&source_path).with_context(|| {
        format!(
            "read source workbook for preservation validation {}",
            input.display()
        )
    })?;
    compare_workbooks(&source, &converted)?;
    if source_ext == "et" {
        let (source_format_verified, source_format, source_file_format) =
            match (engine, source_file_format) {
                ("com", Some(65521)) => (true, "et", 65521),
                _ => (false, "unknown", 0),
            };
        Ok(format!("{{\"level\":\"worksheet_order_names_values_and_formula_text\",\"source_probe\":\"separate_conversion_using_selected_engine\",\"numeric_tolerance\":1e-15,\"source_format_verified\":{source_format_verified},\"source_format\":\"{source_format}\",\"source_file_format\":{source_file_format}}}"))
    } else {
        Ok(
            "{\"level\":\"worksheet_order_names_values_and_formula_text\",\"numeric_tolerance\":1e-15}"
                .into(),
        )
    }
}

fn extract_word_text_lo(input: &Path, output_dir: &Path) -> Result<String> {
    fs::create_dir_all(output_dir)?;
    let profile = output_dir.join("lo-profile");
    let profile_argument = libreoffice_user_installation_arg(&profile)?;
    let command = Command::new("soffice")
        .arg("--headless")
        .arg("--nologo")
        .arg("--nodefault")
        .arg("--nolockcheck")
        .arg(profile_argument)
        .arg("--convert-to")
        .arg("txt:Text")
        .arg("--outdir")
        .arg(output_dir)
        .arg(input)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .output()
        .with_context(|| "该功能需要外部引擎: LibreOffice is required to validate converted Word text")?;
    ensure!(
        command.status.success(),
        "LibreOffice text extraction failed: {}",
        String::from_utf8_lossy(&command.stderr)
    );
    let text_path = output_dir.join(format!("{}.txt", file_stem(input)));
    fs::read_to_string(&text_path)
        .with_context(|| format!("read extracted Word text {}", text_path.display()))
}

fn extract_word_text_com(input: &Path, text_output: &Path) -> Result<String> {
    if let Some(parent) = text_output.parent() {
        fs::create_dir_all(parent)?;
    }
    let script_path = text_output.with_extension("extract-word-text.ps1");
    let script = format!(
        "$ErrorActionPreference='Stop';\n$app=$null; $appName=''; $lastError='';\nforeach ($name in @('Word.Application','kwps.Application','KWPS.Application','wps.Application')) {{ if ($null -eq $app) {{ try {{ $app=New-Object -ComObject $name; $appName=$name }} catch {{ $lastError=$_.Exception.Message }} }} }}\nif ($null -eq $app) {{ throw \"No supported Office/WPS COM engine is available: $lastError\" }}\n$doc=$null;\ntry {{ $app.Visible=$false; $app.DisplayAlerts=0; $doc=$app.Documents.Open('{}',$false,$true); $text=$doc.Content.Text; [System.IO.File]::WriteAllText('{}',$text,[System.Text.UTF8Encoding]::new($false)) }} finally {{ if ($null -ne $doc) {{ $doc.Close($false) }}; if ($null -ne $app) {{ $app.Quit() }} }}\n",
        ps_quote(&input.to_string_lossy()),
        ps_quote(&text_output.to_string_lossy())
    );
    write_powershell_script(&script_path, &script)?;
    let result = Command::new("powershell.exe")
        .args([
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
        ])
        .arg(&script_path)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .output()
        .with_context(|| "start Word COM text validation")?;
    let _ = fs::remove_file(script_path);
    ensure!(
        result.status.success(),
        "Word COM text extraction failed: {}",
        String::from_utf8_lossy(&result.stderr)
    );
    fs::read_to_string(text_output)
        .with_context(|| format!("read extracted Word text {}", text_output.display()))
}

fn normalize_document_text(text: &str) -> String {
    text.replace("\r\n", "\n")
        .replace('\r', "\n")
        .trim_end_matches(['\n', '\u{000c}'])
        .to_owned()
}

fn read_comparable_workbook(path: &Path) -> Result<WorkbookDump> {
    let mut workbook = read_workbook(path)?;
    let has_date_serial = workbook.sheets.iter().any(|sheet| {
        sheet.values.values().any(|value| {
            value.get("type").and_then(serde_json::Value::as_str) == Some("datetime_serial")
        })
    });
    if !has_date_serial {
        return Ok(workbook);
    }
    let max_columns = workbook
        .sheets
        .iter()
        .flat_map(|sheet| {
            sheet
                .values
                .keys()
                .filter_map(|address| column_number(address))
        })
        .max()
        .unwrap_or(1)
        .max(1);
    for sheet in read_value_columns_with_iso_dates(path, max_columns)? {
        let Some(target) = workbook
            .sheets
            .iter_mut()
            .find(|candidate| candidate.name == sheet.name)
        else {
            continue;
        };
        for (address, value) in sheet.values {
            if let ReaderCellValue::DateTimeIso(value) = value {
                target.values.insert(
                    address,
                    serde_json::json!({"type":"datetime_iso","value":value}),
                );
            }
        }
    }
    Ok(workbook)
}

fn compare_workbooks(source: &WorkbookDump, converted: &WorkbookDump) -> Result<()> {
    let source_names = source
        .sheets
        .iter()
        .map(|sheet| &sheet.name)
        .collect::<Vec<_>>();
    let converted_names = converted
        .sheets
        .iter()
        .map(|sheet| &sheet.name)
        .collect::<Vec<_>>();
    ensure!(
        source_names == converted_names,
        "conversion changed worksheet count, names or order"
    );
    for (source_sheet, output_sheet) in source.sheets.iter().zip(&converted.sheets) {
        ensure!(
            source_sheet.formulas.len() == output_sheet.formulas.len(),
            "conversion changed formula count in worksheet {:?}",
            source_sheet.name
        );
        for (address, formula) in &source_sheet.formulas {
            let Some(output_formula) = output_sheet.formulas.get(address) else {
                bail!(
                    "conversion removed formula {}!{}",
                    source_sheet.name,
                    address
                );
            };
            ensure!(
                formula.trim_start_matches('=').trim()
                    == output_formula.trim_start_matches('=').trim(),
                "conversion changed formula text at {}!{}: {:?} -> {:?}",
                source_sheet.name,
                address,
                formula,
                output_formula
            );
        }
        let source_cells = source_sheet
            .values
            .keys()
            .filter(|address| !source_sheet.formulas.contains_key(*address))
            .collect::<std::collections::BTreeSet<_>>();
        let output_cells = output_sheet
            .values
            .keys()
            .filter(|address| !output_sheet.formulas.contains_key(*address))
            .collect::<std::collections::BTreeSet<_>>();
        ensure!(
            source_cells == output_cells,
            "conversion changed non-formula data cell coordinates in worksheet {:?}",
            source_sheet.name
        );
        for address in source_cells {
            let source_value = &source_sheet.values[address];
            let output_value = &output_sheet.values[address];
            ensure!(
                cell_values_equal(source_value, output_value),
                "conversion changed stored value at {}!{}: {} -> {}",
                source_sheet.name,
                address,
                source_value,
                output_value
            );
        }
    }
    Ok(())
}

fn cell_values_equal(left: &serde_json::Value, right: &serde_json::Value) -> bool {
    let numeric = |kind: &str| matches!(kind, "number" | "integer" | "datetime_serial");
    let left_kind = cell_kind(left);
    let right_kind = cell_kind(right);
    if numeric(left_kind) && numeric(right_kind) {
        if left_kind == "integer" && right_kind == "integer" {
            return left.get("value").and_then(serde_json::Value::as_i64)
                == right.get("value").and_then(serde_json::Value::as_i64);
        }
        let Some(left) = left.get("value").and_then(serde_json::Value::as_f64) else {
            return left == right;
        };
        let Some(right) = right.get("value").and_then(serde_json::Value::as_f64) else {
            return false;
        };
        return left == right
            || (left.is_finite()
                && right.is_finite()
                && left != 0.0
                && right != 0.0
                && (left - right).abs() <= 1e-15 * left.abs().max(right.abs()));
    }
    left == right
}

fn cell_kind(value: &serde_json::Value) -> &str {
    value
        .get("type")
        .and_then(serde_json::Value::as_str)
        .unwrap_or_default()
}

fn column_number(address: &str) -> Option<usize> {
    let split = address.find(|character: char| character.is_ascii_digit())?;
    address[..split].bytes().try_fold(0usize, |value, byte| {
        value
            .checked_mul(26)?
            .checked_add((byte.to_ascii_uppercase().checked_sub(b'A')? + 1) as usize)
    })
}

fn locate_converted(temp: &Path, target: &str) -> Result<PathBuf> {
    let mut entries = fs::read_dir(temp)?
        .filter_map(|entry| entry.ok().map(|e| e.path()))
        .filter(|p| {
            p.extension()
                .and_then(|e| e.to_str())
                .is_some_and(|e| e.eq_ignore_ascii_case(target))
        })
        .collect::<Vec<_>>();
    entries.sort();
    entries
        .into_iter()
        .next()
        .ok_or_else(|| anyhow!("conversion produced no .{target} output"))
}

fn unique_temp_dir(parent: &Path) -> Result<PathBuf> {
    let nanos = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
    for i in 0..100 {
        let path = parent.join(format!(
            ".pytools-convert-{}-{:?}-{nanos}-{i}",
            std::process::id(),
            std::thread::current().id()
        ));
        if !path.exists() {
            fs::create_dir(&path)?;
            return Ok(path);
        }
    }
    bail!("cannot allocate conversion work directory")
}

fn extension(path: &Path) -> Result<String> {
    path.extension()
        .and_then(|s| s.to_str())
        .map(|s| s.to_ascii_lowercase())
        .ok_or_else(|| anyhow!("file has no extension: {}", path.display()))
}

fn file_stem(path: &Path) -> String {
    path.file_stem()
        .map(|s| s.to_string_lossy().into_owned())
        .unwrap_or_else(|| "converted".into())
}

fn libreoffice_user_installation_arg(profile: &Path) -> Result<String> {
    fs::create_dir_all(profile).context("create LibreOffice isolated profile directory")?;
    Ok(format!("-env:UserInstallation={}", file_uri(profile)?))
}

fn file_uri(path: &Path) -> Result<String> {
    ensure!(
        path.is_absolute(),
        "LibreOffice profile path must be absolute"
    );
    let path = path.to_string_lossy().replace('\\', "/");
    #[cfg(windows)]
    let path = if path.starts_with('/') {
        path
    } else {
        format!("/{path}")
    };
    let mut encoded = String::with_capacity(path.len());
    for byte in path.bytes() {
        if byte.is_ascii_alphanumeric() || matches!(byte, b'/' | b':' | b'-' | b'.' | b'_' | b'~')
        {
            encoded.push(byte as char);
        } else {
            encoded.push_str(&format!("%{byte:02X}"));
        }
    }
    Ok(format!("file://{encoded}"))
}

fn validate_csv_conversion(input: &Path, converted: &WorkbookDump) -> Result<()> {
    let bytes = fs::read(input).with_context(|| format!("read CSV {}", input.display()))?;
    let bytes = bytes.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(&bytes);
    let decoded = match std::str::from_utf8(bytes) {
        Ok(text) => text.to_owned(),
        Err(_) => {
            let (text, _, had_errors) = GBK.decode(bytes);
            ensure!(
                !had_errors,
                "CSV is neither valid UTF-8 nor valid GBK; cannot validate converted values"
            );
            text.into_owned()
        }
    };
    let mut reader = csv::ReaderBuilder::new()
        .has_headers(false)
        .flexible(true)
        .from_reader(decoded.as_bytes());
    let rows = reader
        .records()
        .collect::<std::result::Result<Vec<_>, _>>()
        .context("parse CSV for conversion validation")?;
    ensure!(
        converted.sheets.len() == 1,
        "CSV conversion must produce exactly one worksheet"
    );
    let sheet = &converted.sheets[0];
    let mut expected_cells = std::collections::BTreeSet::new();
    for (row_index, row) in rows.iter().enumerate() {
        for (column_index, raw) in row.iter().enumerate() {
            if raw.is_empty() {
                continue;
            }
            let address = format!("{}{}", csv_column_name(column_index + 1), row_index + 1);
            expected_cells.insert(address.clone());
            if let Some(formula) = sheet.formulas.get(&address) {
                ensure!(
                    raw.starts_with('=')
                        && raw.trim_start_matches('=').trim()
                            == formula.trim_start_matches('=').trim(),
                    "CSV conversion changed formula text at {}!{address}",
                    sheet.name
                );
                continue;
            }
            let value = sheet
                .values
                .get(&address)
                .with_context(|| format!("CSV conversion lost data at {}!{address}", sheet.name))?;
            ensure!(
                csv_text_matches_probe(raw, value),
                "CSV conversion changed data at {}!{address}: source={raw:?}, output={value}",
                sheet.name
            );
        }
    }
    for address in sheet.values.keys().chain(sheet.formulas.keys()) {
        let non_empty = sheet
            .values
            .get(address)
            .and_then(|value| value.get("value"))
            .is_some_and(|value| {
                !value.is_null() && value.as_str().is_none_or(|text| !text.is_empty())
            });
        if non_empty || sheet.formulas.contains_key(address) {
            ensure!(
                expected_cells.contains(address),
                "CSV conversion introduced data at {}!{address}",
                sheet.name
            );
        }
    }
    Ok(())
}

fn csv_text_matches_probe(source: &str, actual: &serde_json::Value) -> bool {
    let Some(kind) = actual.get("type").and_then(serde_json::Value::as_str) else {
        return false;
    };
    let Some(value) = actual.get("value") else {
        return false;
    };
    match kind {
        "string" | "error" => value.as_str() == Some(source),
        "number" | "integer" | "datetime_serial" => {
            let Some(left) = source.trim().parse::<f64>().ok() else {
                return false;
            };
            let Some(right) = value.as_f64() else {
                return false;
            };
            left.is_finite()
                && right.is_finite()
                && (left - right).abs() <= 1e-15 * left.abs().max(right.abs()).max(1.0)
        }
        "boolean" => {
            let source = source.trim();
            match value.as_bool() {
                Some(true) => source.eq_ignore_ascii_case("true") || source == "1",
                Some(false) => source.eq_ignore_ascii_case("false") || source == "0",
                None => false,
            }
        }
        "datetime_iso" => {
            let source = source.trim().replace('/', "-").replace(' ', "T");
            let actual = value.as_str().unwrap_or_default();
            source == actual
                || (source.len() == 10 && actual.starts_with(&source))
                || chrono::NaiveDate::parse_from_str(&source, "%Y-%m-%d")
                    .ok()
                    .is_some_and(|date| actual.starts_with(&date.format("%Y-%m-%d").to_string()))
        }
        _ => false,
    }
}

fn csv_column_name(mut column: usize) -> String {
    let mut name = String::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        name.insert(0, (b'A' + remainder as u8) as char);
        column = (column - 1) / 26;
    }
    name
}

fn ps_quote(value: &str) -> String {
    value.replace('\'', "''")
}

struct TempDirGuard(PathBuf);

impl Drop for TempDirGuard {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

#[cfg(test)]
mod tests {
    use super::{com_display_alerts_setting, file_uri, libreoffice_user_installation_arg};

    #[test]
    fn com_display_alerts_setting_uses_host_application_type() {
        assert_eq!(com_display_alerts_setting("docx"), "$app.DisplayAlerts=0;");
        assert_eq!(
            com_display_alerts_setting("xlsx"),
            "$app.DisplayAlerts=$false;"
        );
    }

    #[test]
    fn libreoffice_profile_path_is_encoded_as_a_file_uri() {
        let root = std::env::temp_dir().join("pytools engine profile # 中文");
        let uri = file_uri(&root).unwrap();
        assert!(uri.starts_with("file://"));
        assert!(uri.contains("%20"));
        assert!(uri.contains("%23"));
        assert!(!uri.contains(' '));
    }

    #[test]
    fn libreoffice_user_installation_argument_creates_profile() {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let profile = std::env::temp_dir().join(format!(
            "pytools_engine_profile_{}_{nonce}",
            std::process::id()
        ));
        let _ = std::fs::remove_dir_all(&profile);
        let argument = libreoffice_user_installation_arg(&profile).unwrap();
        assert!(profile.is_dir());
        assert!(argument.starts_with("-env:UserInstallation=file://"));
        std::fs::remove_dir_all(profile).ok();
    }
}
