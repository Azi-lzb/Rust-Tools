//! Explicit Office/LibreOffice conversion sidecar. The product binary never
//! links or imports this crate; it starts this executable only for a requested
//! same-category conversion.

use anyhow::{anyhow, bail, ensure, Context, Result};
use encoding_rs::GBK;
use excel_reader::{
    read_value_columns_with_iso_dates, read_workbook, CellValue as ReaderCellValue, WorkbookDump,
};
use std::{
    collections::BTreeMap,
    env, fs,
    path::{Path, PathBuf},
    process::{Command, Stdio},
    time::{SystemTime, UNIX_EPOCH},
};
use zip::ZipArchive;

const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

fn main() {
    if let Err(error) = run() {
        eprintln!("lo-convert: {error:#}");
        std::process::exit(2);
    }
}

fn run() -> Result<()> {
    let mut args = env::args().skip(1);
    let Some(command) = args.next() else {
        print_usage();
        bail!("missing command");
    };
    if command == "--help" || command == "-h" {
        print_usage();
        return Ok(());
    }
    if command == "recalculate" {
        return run_recalculate(args.collect());
    }
    ensure!(command == "convert", "unsupported command {command:?}");
    let mut values = BTreeMap::new();
    while let Some(key) = args.next() {
        ensure!(key.starts_with("--"), "expected an option, got {key:?}");
        let name = key.trim_start_matches('-').to_owned();
        let value = args
            .next()
            .with_context(|| format!("{key} requires a value"))?;
        values.insert(name, value);
    }
    let input = PathBuf::from(values.get("input").context("--input is required")?);
    let output = PathBuf::from(values.get("output").context("--output is required")?);
    let target = values
        .get("to")
        .context("--to is required")?
        .to_ascii_lowercase();
    let mode = values.get("mode").map(String::as_str).unwrap_or("fidelity");
    let engine = values.get("engine").map(String::as_str).unwrap_or("auto");
    ensure!(mode == "fidelity", "only fidelity conversion is implemented; native value-only conversion is not available yet");
    ensure!(
        engine == "auto" || engine == "lo" || engine == "com",
        "--engine must be auto, lo or com"
    );
    ensure!(
        input.is_file(),
        "input file does not exist: {}",
        input.display()
    );
    let source_ext = extension(&input)?;
    ensure!(
        same_category(&source_ext, &target),
        "conversion must stay within the same document category"
    );
    ensure!(
        (matches!(target.as_str(), "xlsx")
            && matches!(source_ext.as_str(), "csv" | "xls" | "xlsb" | "xlsm" | "et"))
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
    let converted = temp_dir.join(format!("{}.{}", file_stem(&input), target));
    let selected = if engine == "auto" {
        if cfg!(windows) {
            "com"
        } else {
            "lo"
        }
    } else {
        engine
    };
    let source_file_format = match selected {
        "lo" => {
            run_libreoffice(&input, temp_dir, &target)?;
            None
        }
        "com" if cfg!(windows) => run_com(&input, &converted, &target, &source_ext)?,
        "com" => bail!(
            "COM conversion is only available on Windows; install LibreOffice or use --engine lo"
        ),
        _ => unreachable!(),
    };
    let candidate = if converted.exists() {
        converted
    } else {
        locate_converted(temp_dir, &target)?
    };
    ensure!(
        candidate.is_file() && fs::metadata(&candidate)?.len() > 0,
        "conversion engine did not create a non-empty output file"
    );
    let validation = validate_conversion(
        &input,
        &candidate,
        &source_ext,
        &target,
        selected,
        temp_dir,
        source_file_format,
    )?;
    fs::rename(&candidate, &output)
        .with_context(|| format!("commit validated output {}", output.display()))?;
    println!("{{\"status\":\"converted\",\"engine\":\"{selected}\",\"input\":{},\"output\":{},\"validation\":{}}}", json_string(&input.display().to_string()), json_string(&output.display().to_string()), validation);
    Ok(())
}

fn run_recalculate(arguments: Vec<String>) -> Result<()> {
    if arguments
        .iter()
        .any(|argument| argument == "--help" || argument == "-h")
    {
        eprintln!("Usage: lo-convert recalculate --input <xlsx> --output <new.xlsx> [--engine auto|lo|com]");
        return Ok(());
    }
    let mut values = BTreeMap::new();
    let mut args = arguments.into_iter();
    while let Some(key) = args.next() {
        ensure!(key.starts_with("--"), "expected an option, got {key:?}");
        let name = key.trim_start_matches('-').to_owned();
        ensure!(
            matches!(name.as_str(), "input" | "output" | "engine"),
            "unknown recalculate option --{name}"
        );
        let value = args
            .next()
            .with_context(|| format!("{key} requires a value"))?;
        values.insert(name, value);
    }
    let input = PathBuf::from(values.get("input").context("--input is required")?);
    let output = PathBuf::from(values.get("output").context("--output is required")?);
    ensure!(
        input.is_file(),
        "input workbook does not exist: {}",
        input.display()
    );
    ensure!(
        input
            .extension()
            .and_then(|ext| ext.to_str())
            .is_some_and(|ext| ext.eq_ignore_ascii_case("xlsx")),
        "formula recalculation currently supports .xlsx only"
    );
    ensure!(
        output
            .extension()
            .and_then(|ext| ext.to_str())
            .is_some_and(|ext| ext.eq_ignore_ascii_case("xlsx")),
        "formula recalculation output must be .xlsx"
    );
    let canonical_input = fs::canonicalize(&input).context("resolve recalculation input")?;
    ensure!(
        output != input && (!output.exists() || fs::canonicalize(&output)? != canonical_input),
        "recalculation requires a distinct output path"
    );
    if let Some(parent) = output
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
    {
        fs::create_dir_all(parent)?;
    }
    ensure!(
        !output.exists(),
        "refusing to overwrite existing recalculation output: {}",
        output.display()
    );
    let engine = values.get("engine").map(String::as_str).unwrap_or("auto");
    ensure!(
        engine == "auto" || engine == "lo" || engine == "com",
        "--engine must be auto, lo or com"
    );
    let selected = if engine == "auto" {
        if cfg!(windows) {
            "com"
        } else {
            "lo"
        }
    } else {
        engine
    };
    if selected == "com" && !cfg!(windows) {
        bail!("COM recalculation is only available on Windows; use --engine lo");
    }

    let guard = TempDirGuard(unique_temp_dir(
        output.parent().unwrap_or_else(|| Path::new(".")),
    )?);
    let input_dir = guard.0.join("input");
    let engine_output_dir = guard.0.join("engine-output");
    fs::create_dir_all(&input_dir)?;
    fs::create_dir_all(&engine_output_dir)?;
    let cacheless_input = input_dir.join("formula-cache-cleared.xlsx");
    let recalc_input = input_dir.join("recalculation-request.xlsx");
    let cache_report =
        ooxml_patch::clear_formula_caches_for_recalculation(&input, &cacheless_input)
            .context("remove prior formula caches from the recalculation request copy")?;
    ooxml_patch::patch_workbook_calculation_properties(
        &cacheless_input,
        &recalc_input,
        ooxml_patch::WorkbookCalculationProperties {
            full_calc_on_load: Some(true),
            force_full_calc: Some(true),
        },
    )
    .context("prepare explicit full recalculation request")?;
    let (actual_engine, candidate) = match selected {
        "lo" => {
            let expected = engine_output_dir.join("recalculation-request.xlsx");
            ensure!(
                !expected.exists(),
                "refusing to accept a pre-existing recalculation output"
            );
            let actual_engine = run_libreoffice_recalculate(&recalc_input, &engine_output_dir)?;
            (actual_engine, expected)
        }
        "com" => {
            let expected = engine_output_dir.join("recalculated.xlsx");
            ensure!(
                !expected.exists(),
                "refusing to accept a pre-existing recalculation output"
            );
            let actual_engine = run_com_recalculate(&recalc_input, &expected)?;
            (actual_engine, expected)
        }
        _ => unreachable!(),
    };
    require_engine_output(&candidate)?;
    let validation = validate_recalculated_workbook(&input, &candidate)?;
    fs::rename(&candidate, &output)
        .with_context(|| format!("publish recalculated workbook {}", output.display()))?;
    println!("{{\"status\":\"recalculated\",\"engine\":{},\"input\":{},\"output\":{},\"formula_cache_values_cleared_before_engine\":{},\"validation\":{}}}", json_string(&actual_engine), json_string(&input.display().to_string()), json_string(&output.display().to_string()), cache_report.cache_values_removed, validation);
    Ok(())
}

fn run_libreoffice_recalculate(input: &Path, temp_dir: &Path) -> Result<String> {
    let profile = temp_dir.join("lo-profile");
    let profile_argument = libreoffice_user_installation_arg(&profile)?;
    let command = Command::new("soffice")
        .arg("--headless")
        .arg("--nologo")
        .arg("--nodefault")
        .arg("--nolockcheck")
        .arg("--nofirststartwizard")
        .arg(profile_argument)
        .arg("--convert-to")
        .arg("xlsx:Calc MS Excel 2007 XML")
        .arg("--outdir")
        .arg(temp_dir)
        .arg(file_uri(input)?)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .output()
        .with_context(|| {
            "LibreOffice sidecar is unavailable; install soffice to recalculate formulas"
        })?;
    ensure!(
        command.status.success(),
        "LibreOffice recalculation failed ({}): {}",
        command.status,
        String::from_utf8_lossy(&command.stderr)
    );
    Ok("LibreOffice".to_owned())
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
        if byte.is_ascii_alphanumeric() || matches!(byte, b'/' | b':' | b'-' | b'.' | b'_' | b'~') {
            encoded.push(byte as char);
        } else {
            encoded.push_str(&format!("%{byte:02X}"));
        }
    }
    Ok(format!("file://{encoded}"))
}

fn libreoffice_user_installation_arg(profile: &Path) -> Result<String> {
    fs::create_dir_all(profile).context("create LibreOffice isolated profile directory")?;
    Ok(format!("-env:UserInstallation={}", file_uri(profile)?))
}

fn require_engine_output(path: &Path) -> Result<()> {
    ensure!(
        path.is_file() && fs::metadata(path)?.len() > 0,
        "calculation engine did not create its expected non-empty output workbook: {}",
        path.display()
    );
    Ok(())
}

fn run_com_recalculate(input: &Path, output: &Path) -> Result<String> {
    let script_path = output.with_extension("recalculate.ps1");
    let script = format!(
        "$ErrorActionPreference='Stop';\n$app=$null; $appName=''; $lastError='';\nforeach ($name in @('Excel.Application','ket.Application','KET.Application','kwps.Application','KWPS.Application','wps.Application')) {{ if ($null -eq $app) {{ try {{ $app=New-Object -ComObject $name; $appName=$name }} catch {{ $lastError=$_.Exception.Message }} }} }}\nif ($null -eq $app) {{ throw \"No supported spreadsheet COM engine is available: $lastError\" }}\n$book=$null;\ntry {{ $app.Visible=$false; $app.DisplayAlerts=$false; $book=$app.Workbooks.Open('{}',$false,$false); $app.CalculateFullRebuild(); $book.SaveAs('{}',51); $book.Close($true) }} finally {{ if ($null -ne $book) {{ $book.Close($false) }}; if ($null -ne $app) {{ $app.Quit() }} }}\nWrite-Output $appName\n",
        ps_quote(&input.to_string_lossy()),
        ps_quote(&output.to_string_lossy())
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
        .with_context(|| "start spreadsheet COM recalculation")?;
    let _ = fs::remove_file(&script_path);
    ensure!(
        result.status.success(),
        "spreadsheet COM recalculation failed: {}",
        String::from_utf8_lossy(&result.stderr)
    );
    let engine = String::from_utf8_lossy(&result.stdout).trim().to_owned();
    ensure!(
        !engine.is_empty(),
        "COM recalculation did not report the selected application"
    );
    Ok(engine)
}

fn validate_recalculated_workbook(input: &Path, output: &Path) -> Result<String> {
    let source = excel_reader::read_workbook_with_unknown_error_tokens_as_text(input)
        .context("read source workbook before recalculation")?
        .workbook;
    let result = excel_reader::read_workbook_with_unknown_error_tokens_as_text(output)
        .context("read recalculated workbook")?
        .workbook;
    let source_sheets = source
        .sheets
        .iter()
        .map(|sheet| &sheet.name)
        .collect::<Vec<_>>();
    let output_sheets = result
        .sheets
        .iter()
        .map(|sheet| &sheet.name)
        .collect::<Vec<_>>();
    ensure!(
        source_sheets == output_sheets,
        "recalculation changed worksheet identity/order"
    );
    let mut formula_count = 0usize;
    for (before, after) in source.sheets.iter().zip(&result.sheets) {
        ensure!(
            before.formulas == after.formulas,
            "recalculation changed formulas in worksheet {:?}",
            before.name
        );
        formula_count += before.formulas.len();
    }
    ensure!(
        formula_count > 0,
        "workbook contains no formulas to recalculate"
    );
    let cache_count = formula_cache_cell_count(output)?;
    ensure!(cache_count == formula_count, "recalculated workbook has formula/cache count mismatch: {formula_count} formulas, {cache_count} cached formula cells");
    Ok(format!("{{\"formula_text_preserved\":true,\"formula_count\":{formula_count},\"formula_cache_cells\":{cache_count},\"cache_nodes_complete\":true,\"scope\":\"cache_presence_and_formula_preservation; result correctness requires parity comparison\"}}"))
}

fn formula_cache_cell_count(path: &Path) -> Result<usize> {
    let mut zip =
        ZipArchive::new(fs::File::open(path).context("open recalculated OOXML package")?)?;
    let root_rels = zip_text(&mut zip, "_rels/.rels")?;
    let rels = roxmltree::Document::parse(&root_rels)?;
    let target = rels
        .descendants()
        .find(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(PACKAGE_REL_NS)
                && node.tag_name().name() == "Relationship"
                && node
                    .attribute("Type")
                    .is_some_and(|kind| kind.ends_with("/officeDocument"))
                && !node
                    .attribute("TargetMode")
                    .is_some_and(|mode| mode.eq_ignore_ascii_case("External"))
        })
        .and_then(|node| node.attribute("Target"))
        .context("package has no internal workbook relationship")?;
    let workbook_part = normalize_part("", target);
    let workbook_xml = zip_text(&mut zip, &workbook_part)?;
    let workbook = roxmltree::Document::parse(&workbook_xml)?;
    let workbook_rels_part = relationships_part(&workbook_part);
    let workbook_rels_xml = zip_text(&mut zip, &workbook_rels_part)?;
    let workbook_rels = roxmltree::Document::parse(&workbook_rels_xml)?;
    let targets = workbook_rels
        .descendants()
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(PACKAGE_REL_NS)
                && node.tag_name().name() == "Relationship"
        })
        .filter_map(|node| {
            Some((
                node.attribute("Id")?.to_owned(),
                (
                    node.attribute("Type")?.to_owned(),
                    node.attribute("Target")?.to_owned(),
                ),
            ))
        })
        .collect::<BTreeMap<_, _>>();
    let parent = workbook_part
        .rsplit_once('/')
        .map(|(parent, _)| parent)
        .unwrap_or("");
    let mut formula_cells = 0usize;
    for sheet in workbook.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(MAIN_NS)
            && node.tag_name().name() == "sheet"
    }) {
        let relationship_id = sheet
            .attribute((REL_NS, "id"))
            .context("sheet has no relationship id")?;
        let (kind, target) = targets
            .get(relationship_id)
            .context("sheet relationship is unresolved")?;
        if !kind.ends_with("/worksheet") {
            continue;
        }
        let part = normalize_part(parent, target);
        let xml = zip_text(&mut zip, &part)?;
        let document = roxmltree::Document::parse(&xml)?;
        for cell in document.descendants().filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(MAIN_NS)
                && node.tag_name().name() == "c"
                && node.children().any(|child| {
                    child.is_element()
                        && child.tag_name().namespace() == Some(MAIN_NS)
                        && child.tag_name().name() == "f"
                })
        }) {
            ensure!(
                cell.children().any(|child| child.is_element()
                    && child.tag_name().namespace() == Some(MAIN_NS)
                    && child.tag_name().name() == "v"),
                "formula cell {:?} on sheet {:?} has no cached value node after recalculation",
                cell.attribute("r"),
                sheet.attribute("name")
            );
            formula_cells += 1;
        }
    }
    Ok(formula_cells)
}

fn zip_text(zip: &mut ZipArchive<fs::File>, name: &str) -> Result<String> {
    let mut entry = zip
        .by_name(name)
        .with_context(|| format!("missing OOXML part {name}"))?;
    let mut text = String::new();
    use std::io::Read;
    entry.read_to_string(&mut text)?;
    Ok(text)
}

fn normalize_part(base: &str, target: &str) -> String {
    let combined = if target.starts_with('/') {
        target.trim_start_matches('/').to_owned()
    } else if base.is_empty() {
        target.to_owned()
    } else {
        format!("{base}/{target}")
    };
    let mut parts = Vec::new();
    for part in combined.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            _ => parts.push(part),
        }
    }
    parts.join("/")
}

fn relationships_part(part: &str) -> String {
    let (parent, file) = part.rsplit_once('/').unwrap_or(("", part));
    if parent.is_empty() {
        format!("_rels/{file}.rels")
    } else {
        format!("{parent}/_rels/{file}.rels")
    }
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
        "LibreOffice sidecar is unavailable; install soffice to use this conversion"
    })?;
    if !output.status.success() {
        bail!(
            "LibreOffice conversion failed ({}): {}",
            output.status,
            String::from_utf8_lossy(&output.stderr)
        );
    }
    Ok(())
}

fn run_com(input: &Path, output: &Path, target: &str, source_ext: &str) -> Result<Option<i32>> {
    let output_parent = output
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    fs::create_dir_all(output_parent)?;
    let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
    let com_input = output_parent.join(format!(
        ".lo-convert-source-{}-{nonce}.{source_ext}",
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
                "$sourceFormat=[int]$book.FileFormat; if ($sourceFormat -ne 65521) { throw \"Input has .et extension but KET opened FileFormat $sourceFormat; expected native ET FileFormat 65521\" }; [Console]::Out.WriteLine(\"LO_CONVERT_SOURCE_FILE_FORMAT=$sourceFormat\");\n",
            );
        }
        script.push_str(&format!(
            "$book.SaveAs('{}',{});\n",
            ps_quote(&output.to_string_lossy()),
            format_code
        ));
    }
    script.push_str("} finally { if ($null -ne $doc) { $doc.Close($false) }; if ($null -ne $book) { $book.Close($false) }; if ($null -ne $app) { $app.Quit() } }\n");
    let script_path = output.with_extension("lo-convert.ps1");
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
            .find_map(|line| line.trim().strip_prefix("LO_CONVERT_SOURCE_FILE_FORMAT="))
            .context("KET did not return source FileFormat evidence for .et input")?
            .parse::<i32>()
            .context("KET returned invalid source FileFormat evidence")?;
        ensure!(
            file_format == 65521,
            "KET opened .et input as FileFormat {file_format}; expected native ET FileFormat 65521"
        );
        Ok(Some(file_format))
    } else {
        Ok(None)
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
        return Ok(
            "{\"level\":\"one_worksheet_and_csv_values\",\"numeric_tolerance\":1e-15}".into(),
        );
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
                let probe_file_format = run_com(input, &probe, target, source_ext)
                    .context("convert ET source for preservation validation with COM")?;
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
        .with_context(|| "LibreOffice is required to validate converted Word text")?;
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
        "$ErrorActionPreference='Stop';\n$app=$null; $appName=''; $lastError='';\nforeach ($name in @('Word.Application','kwps.Application','KWPS.Application','wps.Application')) {{ if ($null -eq $app) {{ try {{ $app=New-Object -ComObject $name; $appName=$name }} catch {{ $lastError=$_.Exception.Message }} }} }}\nif ($null -eq $app) {{ throw \"No supported Office/WPS COM engine is available: $lastError\" }}\n$doc=$null;\ntry {{ $app.Visible=$false; $app.DisplayAlerts=$false; $doc=$app.Documents.Open('{}',$false,$true); $text=$doc.Content.Text; [System.IO.File]::WriteAllText('{}',$text,[System.Text.UTF8Encoding]::new($false)) }} finally {{ if ($null -ne $doc) {{ $doc.Close($false) }}; if ($null -ne $app) {{ $app.Quit() }} }}\n",
        ps_quote(&input.to_string_lossy()),
        ps_quote(&text_output.to_string_lossy())
    );
    let script = script.replace(
        "$app.DisplayAlerts=$false;",
        com_display_alerts_setting("docx"),
    );
    ensure!(
        script.contains(com_display_alerts_setting("docx")),
        "could not configure Word COM alert level"
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
        let path = parent.join(format!(".lo-convert-{}-{nanos}-{i}", std::process::id()));
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
fn same_category(source: &str, target: &str) -> bool {
    let excel = ["csv", "xls", "xlsx", "xlsm", "xlsb", "et"];
    let word = ["doc", "docx"];
    (excel.contains(&source) && target == "xlsx") || (word.contains(&source) && target == "docx")
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
fn json_string(value: &str) -> String {
    serde_json::to_string(value).unwrap_or_else(|_| "\"\"".into())
}
fn print_usage() {
    eprintln!("Usage: lo-convert convert --input <file> --output <file> --to <xlsx|docx> [--mode fidelity] [--engine auto|lo|com]\n       lo-convert recalculate --input <xlsx> --output <new.xlsx> [--engine auto|lo|com]");
}

struct TempDirGuard(PathBuf);

impl Drop for TempDirGuard {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

#[cfg(test)]
#[path = "../tests/support/et_fixture.rs"]
mod et_fixture_support;

#[cfg(test)]
mod et_com_source_safety_tests {
    use super::{et_fixture_support, run_com};
    use anyhow::{ensure, Result};
    use std::{env, fs, path::Path};

    #[test]
    fn com_rejects_excel8_source_named_et_without_mutating_it() -> Result<()> {
        if !cfg!(windows) {
            eprintln!("skipping ET COM failure test: COM is only available on Windows");
            return Ok(());
        }
        if env::var("LO_CONVERT_RUN_WPS_ET_TESTS").as_deref() != Ok("1") {
            eprintln!("skipping WPS ET failure test: set LO_CONVERT_RUN_WPS_ET_TESTS=1 to enable");
            return Ok(());
        }
        let Some(fixture) = env::var_os("LO_CONVERT_CFB_ET_NAMED_XLS_FIXTURE") else {
            eprintln!("skipping WPS CFB source-safety test: fixture path is not set");
            return Ok(());
        };
        let root = et_fixture_support::temp_test_dir("lo_convert_et_com_failure")?;
        let input = et_fixture_support::copy_and_seed_excel8_et_named_fixture(
            &root.0,
            Path::new(&fixture),
        )?;
        let source_bytes = fs::read(&input)?;

        let output = root.0.join("rejected.xlsx");
        let error = run_com(&input, &output, "xlsx", "et")
            .expect_err("Excel 8 content renamed to .et must be rejected");
        ensure!(
            error
                .to_string()
                .contains("expected native ET FileFormat 65521"),
            "expected native ET format rejection, got: {error:#}"
        );
        ensure!(
            fs::read(&input)? == source_bytes,
            "COM failure modified the caller's ET source"
        );
        ensure!(!output.exists(), "rejected .et input produced an output");
        Ok(())
    }
}

#[cfg(test)]
mod recalculation_tests {
    use super::{
        com_display_alerts_setting, file_uri, formula_cache_cell_count,
        libreoffice_user_installation_arg, require_engine_output, write_powershell_script, MAIN_NS,
        PACKAGE_REL_NS, REL_NS,
    };
    use anyhow::Result;
    use std::{
        env,
        fs::{self, File},
        io::Write,
        path::PathBuf,
        time::{SystemTime, UNIX_EPOCH},
    };
    use zip::{write::SimpleFileOptions, ZipWriter};

    #[test]
    fn com_display_alerts_setting_uses_host_application_type() {
        assert_eq!(com_display_alerts_setting("docx"), "$app.DisplayAlerts=0;");
        assert_eq!(
            com_display_alerts_setting("xlsx"),
            "$app.DisplayAlerts=$false;"
        );
    }

    #[test]
    fn powershell_scripts_are_written_as_utf8_with_bom() -> Result<()> {
        let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let path = env::temp_dir().join(format!(
            "lo_convert_script_{}_{}.ps1",
            std::process::id(),
            nonce
        ));
        write_powershell_script(&path, "$input='中文路径';")?;
        let bytes = fs::read(&path)?;
        assert!(bytes.starts_with(&[0xEF, 0xBB, 0xBF]));
        assert_eq!(std::str::from_utf8(&bytes[3..])?, "$input='中文路径';");
        fs::remove_file(path)?;
        Ok(())
    }

    fn fixture(with_cache: bool) -> Result<PathBuf> {
        let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let path = env::temp_dir().join(format!(
            "lo_recalc_fixture_{}_{}.xlsx",
            std::process::id(),
            nonce
        ));
        let zip = ZipWriter::new(File::create(&path)?);
        let workbook = format!(
            r#"<workbook xmlns="{MAIN_NS}" xmlns:r="{REL_NS}"><sheets><sheet name="Formula" sheetId="1" r:id="rId1"/></sheets></workbook>"#
        );
        let rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#
        );
        let root_rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
        );
        let cache = if with_cache { "<v>2</v>" } else { "" };
        let sheet = format!(
            r#"<worksheet xmlns="{MAIN_NS}"><sheetData><row r="1"><c r="A1"><f>1+1</f>{cache}</c></row></sheetData></worksheet>"#
        );
        let mut writer = zip;
        for (name, content) in [
            ("_rels/.rels", root_rels),
            ("xl/workbook.xml", workbook),
            ("xl/_rels/workbook.xml.rels", rels),
            ("xl/worksheets/sheet1.xml", sheet),
        ] {
            writer.start_file(name, SimpleFileOptions::default())?;
            writer.write_all(content.as_bytes())?;
        }
        writer.finish()?;
        Ok(path)
    }

    #[test]
    fn formula_cache_counter_requires_a_value_node_for_every_formula() -> Result<()> {
        let valid = fixture(true)?;
        assert_eq!(formula_cache_cell_count(&valid)?, 1);
        fs::remove_file(valid)?;

        let invalid = fixture(false)?;
        let error =
            formula_cache_cell_count(&invalid).expect_err("formula without cache must fail");
        assert!(error.to_string().contains("has no cached value node"));
        fs::remove_file(invalid)?;
        Ok(())
    }

    #[test]
    fn missing_engine_output_is_rejected_even_when_request_copy_exists() -> Result<()> {
        let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let root =
            env::temp_dir().join(format!("lo_recalc_output_{}_{}", std::process::id(), nonce));
        let input_dir = root.join("input");
        let output_dir = root.join("engine-output");
        fs::create_dir_all(&input_dir)?;
        fs::create_dir_all(&output_dir)?;
        let request_copy = input_dir.join("recalculation-request.xlsx");
        fs::write(&request_copy, b"input copy")?;
        let expected_output = output_dir.join("recalculation-request.xlsx");

        let error = require_engine_output(&expected_output)
            .expect_err("missing engine output must not select the input copy");
        assert!(error.to_string().contains("did not create"));
        assert!(request_copy.is_file());
        assert!(!expected_output.exists());

        fs::remove_dir_all(root)?;
        Ok(())
    }

    #[test]
    fn libreoffice_profile_path_is_encoded_as_a_file_uri() -> Result<()> {
        let root = env::temp_dir().join("lo profile # 中文");
        let uri = file_uri(&root)?;
        assert!(uri.starts_with("file://"));
        assert!(uri.contains("%20"));
        assert!(uri.contains("%23"));
        assert!(!uri.contains(' '));
        assert!(!uri.contains('中'));
        Ok(())
    }

    #[test]
    fn libreoffice_user_installation_argument_creates_profile_and_uses_file_uri() -> Result<()> {
        let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let profile = env::temp_dir().join(format!(
            "lo_profile_arg_{}_{}_# space",
            std::process::id(),
            nonce
        ));
        assert!(!profile.exists());

        let argument = libreoffice_user_installation_arg(&profile)?;

        assert!(profile.is_dir());
        assert!(argument.starts_with("-env:UserInstallation=file://"));
        assert!(argument.contains("%20"));
        assert!(argument.contains("%23"));
        fs::remove_dir_all(profile)?;
        Ok(())
    }
}
