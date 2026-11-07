use anyhow::{Context, Result};
use excel_editor::{
    BackendCapability, CapabilityEvidence, CapabilityStatus, EditOperation, EngineId, ExcelEditPlan,
};
use excel_reader::read_workbook;
use pytools_core::sheet_rename::{build_sheet_rename_plan, SHEET_RENAME_OUTPUT_DIRECTORY};
use serde_json::{json, Value};
use std::{
    collections::BTreeSet,
    fs::{self, File},
    io::{Read, Write},
    path::{Path, PathBuf},
    process::Command,
    sync::atomic::{AtomicU64, Ordering},
};
use zip::{write::SimpleFileOptions, ZipArchive, ZipWriter};

const PLATFORM: &str = "pt45-synthetic-probe";
static TEMP_SEQUENCE: AtomicU64 = AtomicU64::new(0);
const OFFICE_REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

struct TestDir(PathBuf);

impl TestDir {
    fn new(label: &str) -> Result<Self> {
        let path = std::env::temp_dir().join(format!(
            "pytools-pt45-{label}-{}-{}",
            std::process::id(),
            TEMP_SEQUENCE.fetch_add(1, Ordering::Relaxed)
        ));
        fs::create_dir_all(&path)?;
        Ok(Self(path))
    }

    fn path(&self, name: &str) -> PathBuf {
        self.0.join(name)
    }
}

impl Drop for TestDir {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

fn fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/sheet_rename")
        .join(name)
}

fn capability_for(plan: &ExcelEditPlan, status: CapabilityStatus) -> BackendCapability {
    BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status,
        format: plan.format,
        operations: plan.operations.iter().map(EditOperation::kind).collect(),
        supported_features: plan.workbook.features_present.clone(),
        preserved: plan.preserve.clone(),
        platform: PLATFORM.to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: "synthetic-pt-4-4-sheet-rename".to_owned(),
            report_path: "pytools-core/tests/sheet_rename_execution.rs".to_owned(),
            adapter_version: "umya-spreadsheet-3.1.1".to_owned(),
            engine_version: "umya-spreadsheet-3.1.1".to_owned(),
            platform: PLATFORM.to_owned(),
        }),
    }
}

fn write_capability(path: &Path, plan: &ExcelEditPlan, status: CapabilityStatus) -> Result<()> {
    fs::write(
        path,
        serde_json::to_vec_pretty(&capability_for(plan, status))?,
    )?;
    Ok(())
}

fn configure_source_openpyxl_worker(command: &mut Command) {
    command.env(
        "PYTOOLS_COM_WORKER",
        Path::new(env!("CARGO_MANIFEST_DIR")).join("python/pywin32_excel_conversion_worker.py"),
    );
}

fn plan_for(config: &Path, input: &Path, output: &Path) -> Result<ExcelEditPlan> {
    build_sheet_rename_plan(config, input, output)?
        .plan
        .context("fixture should produce a non-empty sheet-rename plan")
}

fn write_legacy_sheet_mapping(path: &Path) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("重命名配置")?;
    worksheet.write_string(0, 9, "原表名")?;
    worksheet.write_string(0, 10, "新表名")?;
    worksheet.write_string(1, 9, "人民币")?;
    worksheet.write_string(1, 10, "1人民币")?;
    worksheet.write_string(2, 9, "外币")?;
    worksheet.write_string(2, 10, "2外币")?;
    workbook.save(path)?;
    Ok(())
}

fn run_cli(config: &Path, capability: &Path, input: &Path) -> Result<std::process::Output> {
    let mut command = Command::new(env!("CARGO_BIN_EXE_RustTools"));
    configure_source_openpyxl_worker(&mut command);
    Ok(command
        .args(["rename-sheets", "--config"])
        .arg(config)
        .args(["--capabilities"])
        .arg(capability)
        .args(["--platform", PLATFORM])
        .arg(input)
        .output()?)
}

/// Default rename path (no capability file): the bundled Openpyxl worker
/// updates cross-sheet references and named ranges while renaming.
fn run_cli_excel(config: &Path, input: &Path) -> Result<std::process::Output> {
    let mut command = Command::new(env!("CARGO_BIN_EXE_RustTools"));
    configure_source_openpyxl_worker(&mut command);
    Ok(command
        .args(["rename-sheets", "--config"])
        .arg(config)
        .arg(input)
        .output()?)
}

fn fixture_with_formula_cache(source: &Path, output: &Path) -> Result<()> {
    let mut archive = ZipArchive::new(File::open(source)?)?;
    let mut writer = ZipWriter::new(File::create(output)?);
    let mut formula_cache_injected = false;
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index)?;
        let name = entry.name().to_owned();
        let compression = entry.compression();
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes)?;
        if name == "xl/worksheets/sheet2.xml" {
            let mut xml = String::from_utf8(bytes)?;
            let original = r#"<c r="B1"><f>'A'!A1</f><v /></c>"#;
            let cached = r#"<c r="B1" t="str"><f>'A'!A1</f><v>cell-a</v></c>"#;
            anyhow::ensure!(
                xml.contains(original),
                "formula cache fixture cell B1 missing"
            );
            xml = xml.replacen(original, cached, 1);
            bytes = xml.into_bytes();
            formula_cache_injected = true;
        }
        writer.start_file(
            name,
            SimpleFileOptions::default().compression_method(compression),
        )?;
        writer.write_all(&bytes)?;
    }
    anyhow::ensure!(
        formula_cache_injected,
        "formula cache fixture sheet missing"
    );
    writer.finish()?;
    Ok(())
}

fn worksheet_cell_xml(path: &Path, part: &str, address: &str) -> Result<String> {
    let mut archive = ZipArchive::new(File::open(path)?)?;
    let mut sheet = String::new();
    archive.by_name(part)?.read_to_string(&mut sheet)?;
    let address_marker = format!("r=\"{address}\"");
    let address_index = sheet
        .find(&address_marker)
        .with_context(|| format!("cell {address} missing from {part}"))?;
    let cell_start = sheet[..address_index]
        .rfind("<c ")
        .context("cell start tag missing")?;
    let cell_end = sheet[address_index..]
        .find("</c>")
        .map(|offset| address_index + offset + "</c>".len())
        .context("cell end tag missing")?;
    Ok(sheet[cell_start..cell_end].to_owned())
}

fn run_cli_excel_recursive(config: &Path, input: &Path) -> Result<std::process::Output> {
    let mut command = Command::new(env!("CARGO_BIN_EXE_RustTools"));
    configure_source_openpyxl_worker(&mut command);
    Ok(command
        .args(["rename-sheets", "--config"])
        .arg(config)
        .arg("--recursive")
        .arg(input)
        .output()?)
}

fn baseline_report() -> Result<Value> {
    Ok(serde_json::from_str(include_str!(
        "../probe/reports/PT-4-4_python_baseline_20261023.json"
    ))?)
}

#[test]
fn sheet_rename_cli_rewrites_references_and_protects_source_snapshot() -> Result<()> {
    let dir = TestDir::new("execution")?;
    let config = fixture("config.xlsx");
    let input = dir.path("input.xlsx");
    fs::copy(fixture("input.xlsx"), &input)?;
    let original_hash = pytools_core::file_commit::sha256_file(&input)?;
    let config_hash = pytools_core::file_commit::sha256_file(&config)?;
    let output = run_cli_excel(&config, &input)?;
    anyhow::ensure!(
        output.status.success(),
        "rename-sheets CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let actual: Value = serde_json::from_slice(&output.stdout)?;
    let baseline = baseline_report()?;
    assert_eq!(actual["workbooks"], baseline["result"]["workbooks"]);
    assert_eq!(actual["rename_ok"], baseline["result"]["rename_ok"]);
    assert_eq!(actual["skip"], baseline["result"]["skip"]);
    assert_eq!(actual["files"][0]["status"], "renamed");
    // Openpyxl 改名路径不使用通用 ExcelEditPlan executor，execution 计划报告为空。
    assert!(actual["files"][0]["execution"].is_null());
    // 核心契约：跨表公式与命名区域引用随改名自动更新。
    let renamed_path = PathBuf::from(
        actual["files"][0]["output"]
            .as_str()
            .context("renamed workbook path missing")?,
    );
    let renamed = excel_reader::read_workbook(&renamed_path)?;
    // 4-5 改为就地修改：源工作簿被改写，配置保持不变。
    assert_ne!(
        pytools_core::file_commit::sha256_file(&input)?,
        original_hash,
        "in-place sheet rename must rewrite the source workbook"
    );
    assert_eq!(
        pytools_core::file_commit::sha256_file(&config)?,
        config_hash
    );

    let mut expected_projection = baseline["projection"].clone();
    // 打印区域里简单表名的带引号/不带引号写法等价，比较前统一规范化。
    if let Some(print_area) = expected_projection["renamed_a1"]["print_area"].as_str() {
        expected_projection["renamed_a1"]["print_area"] =
            normalize_sheet_address(print_area).into();
    }
    // 旧 Python 基线把公式和命名区域留在旧表名（'A'!），改名后它们实际指向
    // 新表 A（原 B 表），属于用户要求修复的依赖失效。当前路径按改名映射改写：
    // 原 A 表已改为 RenamedA，因此引用应为 RenamedA。
    expected_projection["formula"] = json!("=RenamedA!A1");
    for defined_name in expected_projection["defined_names"]
        .as_array_mut()
        .context("baseline defined_names missing")?
    {
        if defined_name["name"] == "TargetCell" {
            defined_name["reference"] = json!("RenamedA!$A$1");
        }
    }
    // 就地模式：输出路径即源工作簿本身。
    let output_path = renamed_path.clone();
    let canonical_input = fs::canonicalize(&input)?;
    let canonical_output = fs::canonicalize(&output_path)?;
    assert_eq!(
        canonical_output, canonical_input,
        "in-place rename must target the source"
    );
    assert!(output_path.is_file());
    // 公式依赖随改名更新：旧表名不再出现在任何工作表公式中。
    let stale = renamed
        .sheets
        .iter()
        .flat_map(|sheet| sheet.formulas.values())
        .any(|formula| formula.contains("OldSheet"));
    assert!(!stale, "stale OldSheet reference remains after rename");
    assert_ne!(
        pytools_core::file_commit::sha256_file(&output_path)?,
        original_hash
    );
    let actual_projection = workbook_projection(&output_path)?;
    assert_eq!(actual_projection, expected_projection);

    let leaked_candidates = fs::read_dir(output_path.parent().context("output parent missing")?)?
        .filter_map(std::result::Result::ok)
        .filter(|entry| {
            entry
                .file_name()
                .to_string_lossy()
                .starts_with(".pytools-sheet-rename-")
        })
        .count();
    assert_eq!(leaked_candidates, 0);

    if let Some(report_path) = std::env::var_os("PYTOOLS_PT45_REPORT_PATH") {
        let report_path = PathBuf::from(report_path);
        if let Some(parent) = report_path.parent() {
            fs::create_dir_all(parent)?;
        }
        let report = json!({
            "suite": "pytools.pt_4_5_sheet_rename",
            "comparison_status": "passed_synthetic_projection",
            "python_baseline_report": "pytools-core/probe/reports/PT-4-4_python_baseline_20261023.json",
            "sample_class": "synthetic",
            "engine": "openpyxl",
            "capability_scope": "bundled Openpyxl worker; no production capability gate",
            "input_sha256_before": original_hash,
            "output_sha256_after": pytools_core::file_commit::sha256_file(&output_path)?,
            "comparison_note": "Print-area sheet quoting is normalized because quoted and unquoted simple sheet names are equivalent OOXML references.",
            "scope_limit": "No real business workbook, XLSM, Windows-native Rust, target Linux, or deb validation.",
            "rust_cli_result": actual,
        });
        fs::write(report_path, serde_json::to_vec_pretty(&report)?)?;
    }
    Ok(())
}

#[test]
fn sheet_rename_preserves_saved_formula_cache() -> Result<()> {
    let dir = TestDir::new("formula-cache")?;
    let config = fixture("config.xlsx");
    let input = dir.path("input.xlsx");
    fixture_with_formula_cache(&fixture("input.xlsx"), &input)?;
    let input_formula_cache_cell = worksheet_cell_xml(&input, "xl/worksheets/sheet2.xml", "B1")?;
    assert!(
        input_formula_cache_cell.contains("<v>cell-a</v>"),
        "test fixture did not contain a saved formula cache: {input_formula_cache_cell}"
    );
    let before_hash = pytools_core::file_commit::sha256_file(&input)?;

    let output = run_cli_excel(&config, &input)?;
    anyhow::ensure!(
        output.status.success(),
        "rename-sheets CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let actual: Value = serde_json::from_slice(&output.stdout)?;
    assert_eq!(actual["files"][0]["status"], "renamed");
    let output_path = PathBuf::from(
        actual["files"][0]["output"]
            .as_str()
            .context("renamed workbook path missing")?,
    );
    assert_eq!(
        fs::canonicalize(&output_path)?,
        fs::canonicalize(&input)?,
        "sheet rename must modify the selected source copy in place"
    );
    assert_ne!(pytools_core::file_commit::sha256_file(&input)?, before_hash);

    let formula_cache_cell = worksheet_cell_xml(&output_path, "xl/worksheets/sheet2.xml", "B1")?;
    assert!(
        formula_cache_cell.contains("<f>RenamedA!A1</f>"),
        "formula reference was not rewritten: {formula_cache_cell}"
    );
    assert!(
        formula_cache_cell.contains("<v>cell-a</v>"),
        "saved formula cache was not retained: {formula_cache_cell}"
    );
    Ok(())
}

#[test]
fn legacy_capability_does_not_route_sheet_rename_away_from_openpyxl() -> Result<()> {
    let dir = TestDir::new("legacy-capability")?;
    let config = fixture("config.xlsx");
    let input = dir.path("input.xlsx");
    fs::copy(fixture("input.xlsx"), &input)?;
    let original_hash = pytools_core::file_commit::sha256_file(&input)?;
    let plan = plan_for(&config, &input, &dir.path("plan-output.xlsx"))?;
    let capability = dir.path("passed-capability.json");
    write_capability(&capability, &plan, CapabilityStatus::Passed)?;

    let output = run_cli(&config, &capability, &input)?;
    anyhow::ensure!(
        output.status.success(),
        "legacy capability was not ignored: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let actual: Value = serde_json::from_slice(&output.stdout)?;
    assert_eq!(actual["files"][0]["status"], "renamed");
    assert_ne!(
        pytools_core::file_commit::sha256_file(&input)?,
        original_hash
    );
    let renamed = excel_reader::read_workbook(&input)?;
    assert!(renamed
        .sheets
        .iter()
        .flat_map(|sheet| sheet.formulas.values())
        .any(|formula| formula.contains("RenamedA")));
    assert!(!renamed
        .sheets
        .iter()
        .flat_map(|sheet| sheet.formulas.values())
        .any(|formula| formula.contains("'A'!")));
    assert_eq!(
        fs::read_dir(&dir.0)?
            .filter_map(std::result::Result::ok)
            .filter(|entry| {
                entry
                    .file_name()
                    .to_string_lossy()
                    .starts_with(".pytools-sheet-rename-")
            })
            .count(),
        0
    );
    assert!(!dir.path(SHEET_RENAME_OUTPUT_DIRECTORY).exists());
    Ok(())
}

#[test]
fn recursive_directory_batch_reports_unsupported_inputs_without_touching_them() -> Result<()> {
    let dir = TestDir::new("recursive")?;
    let config = fixture("config.xlsx");
    let input_dir = dir.path("input");
    let nested = input_dir.join("nested");
    fs::create_dir_all(&nested)?;
    let first = input_dir.join("a.xlsx");
    let second = nested.join("b.xlsx");
    let unsupported = input_dir.join("old.xlsm");
    fs::copy(fixture("input.xlsx"), &first)?;
    fs::copy(fixture("input.xlsx"), &second)?;
    let first_hash = pytools_core::file_commit::sha256_file(&first)?;
    let second_hash = pytools_core::file_commit::sha256_file(&second)?;
    fs::write(&unsupported, b"not an xls workbook")?;
    fs::write(input_dir.join("~$lock.xlsx"), b"ignored office lock")?;
    let unsupported_hash = pytools_core::file_commit::sha256_file(&unsupported)?;

    let output = run_cli_excel_recursive(&config, &input_dir)?;
    anyhow::ensure!(
        output.status.success(),
        "recursive rename-sheets CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report: Value = serde_json::from_slice(&output.stdout)?;
    assert_eq!(report["files_scanned"], 3);
    assert_eq!(report["workbooks"], 2);
    assert_eq!(report["rename_ok"], 10);
    assert_eq!(report["skip"], 11);
    let files = report["files"].as_array().context("file results missing")?;
    let status_for = |suffix: &str| -> Result<&str> {
        files
            .iter()
            .find(|file| {
                file["input"]
                    .as_str()
                    .is_some_and(|input| input.replace('\\', "/").ends_with(suffix))
            })
            .and_then(|file| file["status"].as_str())
            .with_context(|| format!("result for {suffix} missing"))
    };
    assert_eq!(status_for("a.xlsx")?, "renamed");
    assert_eq!(status_for("nested/b.xlsx")?, "renamed");
    assert_eq!(status_for("old.xlsm")?, "skipped");
    // 就地模式：两个 xlsx 都被改写，不支持的 xlsm 保持原样。
    assert_ne!(pytools_core::file_commit::sha256_file(&first)?, first_hash);
    assert_ne!(
        pytools_core::file_commit::sha256_file(&second)?,
        second_hash
    );
    assert_eq!(
        pytools_core::file_commit::sha256_file(&unsupported)?,
        unsupported_hash
    );

    let second_output = run_cli_excel_recursive(&config, &input_dir)?;
    anyhow::ensure!(
        second_output.status.success(),
        "repeat recursive rename-sheets CLI failed: {}",
        String::from_utf8_lossy(&second_output.stderr)
    );
    let second_report: Value = serde_json::from_slice(&second_output.stdout)?;
    assert_eq!(second_report["files_scanned"], 3);
    assert_eq!(second_report["workbooks"], 2);
    // 就地模式：源文件在首轮已被改写，此处只要求它们仍存在且可读。
    assert!(first.is_file());
    assert!(second.is_file());
    let second_files = second_report["files"]
        .as_array()
        .context("repeat results missing")?;
    assert!(second_files.iter().all(|file| {
        file["input"]
            .as_str()
            .is_some_and(|input| !input.contains(SHEET_RENAME_OUTPUT_DIRECTORY))
    }));
    // 就地模式：第二轮可重复执行且不破坏源文件（每次改名幂等或按映射再处理）。
    for file in second_files {
        let status = file["status"].as_str().unwrap_or_default();
        assert!(
            matches!(status, "renamed" | "unchanged" | "skipped"),
            "unexpected repeat status: {status}"
        );
        if let Some(output) = file["output"].as_str() {
            assert!(
                Path::new(output).is_file(),
                "reported output must exist: {output}"
            );
        }
    }
    Ok(())
}

#[cfg(not(windows))]
#[test]
fn legacy_xls_sheet_rename_requires_windows_com() -> Result<()> {
    let dir = TestDir::new("xls-convert-first")?;
    let config = dir.path("config.xlsx");
    write_legacy_sheet_mapping(&config)?;
    let input = dir.path("input.xls");
    fs::copy(fixture("com_input.xls"), &input)?;
    let original_hash = pytools_core::file_commit::sha256_file(&input)?;

    let output = run_cli_excel(&config, &input)?;
    anyhow::ensure!(
        output.status.success(),
        ".xls should be reported as skipped without Windows COM: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report: Value = serde_json::from_slice(&output.stdout)?;
    assert_eq!(report["files"][0]["status"], "skipped");
    let reason = report["files"][0]["reason"]
        .as_str()
        .context("skip reason missing")?;
    assert!(
        reason.contains("Windows Excel/WPS COM"),
        "unexpected skip reason: {reason}"
    );
    assert_eq!(
        pytools_core::file_commit::sha256_file(&input)?,
        original_hash
    );
    Ok(())
}

#[cfg(windows)]
#[test]
#[ignore = "requires an installed Excel or WPS COM server"]
fn legacy_xls_sheet_rename_uses_dedicated_com_route() -> Result<()> {
    let dir = TestDir::new("xls-com-rename")?;
    let config = dir.path("config.xlsx");
    write_legacy_sheet_mapping(&config)?;
    let input = dir.path("input.xls");
    fs::copy(fixture("com_input.xls"), &input)?;
    let original_hash = pytools_core::file_commit::sha256_file(&input)?;

    let output = run_cli_excel(&config, &input)?;
    anyhow::ensure!(
        output.status.success(),
        ".xls COM rename failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report: Value = serde_json::from_slice(&output.stdout)?;
    assert_eq!(report["files"][0]["status"], "renamed");
    assert_eq!(report["files"][0]["renamed"], 2);
    let engine = report["files"][0]["engine"]
        .as_str()
        .context("COM application name missing")?;
    assert!(
        engine.contains("Application"),
        "unexpected COM engine: {engine}"
    );
    assert_ne!(
        pytools_core::file_commit::sha256_file(&input)?,
        original_hash,
        "successful in-place rename should update the isolated source copy"
    );
    let renamed = read_workbook(&input)?;
    let names = renamed
        .sheets
        .iter()
        .map(|sheet| sheet.name.as_str())
        .collect::<Vec<_>>();
    assert!(names.contains(&"1人民币"));
    assert!(names.contains(&"2外币"));
    Ok(())
}

fn workbook_projection(path: &Path) -> Result<Value> {
    let dump = read_workbook(path)?;
    let sheet_names = dump
        .sheets
        .iter()
        .map(|sheet| sheet.name.clone())
        .collect::<Vec<_>>();
    let a1_values = dump
        .sheets
        .iter()
        .filter_map(|sheet| {
            sheet
                .values
                .get("A1")
                .and_then(|value| value.get("value"))
                .map(|value| (sheet.name.clone(), value.clone()))
        })
        .collect::<serde_json::Map<_, _>>();
    let formula = dump
        .sheets
        .iter()
        .find(|sheet| sheet.name == "A")
        .and_then(|sheet| sheet.formulas.get("B1"))
        .map(|formula| format!("={}", formula.trim_start_matches('=')));
    let renamed = dump
        .sheets
        .iter()
        .find(|sheet| sheet.name == "RenamedA")
        .context("RenamedA sheet missing")?;
    let (workbook_xml, first_sheet_xml, styles_xml) = (
        read_part(path, "xl/workbook.xml")?,
        read_worksheet_part(path, "RenamedA")?,
        read_part(path, "xl/styles.xml")?,
    );
    let workbook_doc = roxmltree::Document::parse(&workbook_xml)?;
    let sheet_doc = roxmltree::Document::parse(&first_sheet_xml)?;
    let styles_doc = roxmltree::Document::parse(&styles_xml)?;
    let style_id = sheet_doc
        .descendants()
        .find(|node| {
            node.is_element() && node.tag_name().name() == "c" && node.attribute("r") == Some("A1")
        })
        .and_then(|node| node.attribute("s"))
        .unwrap_or("0")
        .parse::<usize>()?;
    let cell_xfs = styles_doc
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "cellXfs")
        .context("cellXfs missing")?;
    let style = cell_xfs
        .children()
        .filter(roxmltree::Node::is_element)
        .nth(style_id)
        .context("A1 style index missing")?;
    let font_id = style.attribute("fontId").unwrap_or("0").parse::<usize>()?;
    let fill_id = style.attribute("fillId").unwrap_or("0").parse::<usize>()?;
    let font = styles_doc
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "fonts")
        .and_then(|node| {
            node.children()
                .filter(roxmltree::Node::is_element)
                .nth(font_id)
        })
        .context("A1 font missing")?;
    let fill = styles_doc
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "fills")
        .and_then(|node| {
            node.children()
                .filter(roxmltree::Node::is_element)
                .nth(fill_id)
        })
        .context("A1 fill missing")?;
    let color = fill
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "fgColor")
        .and_then(|node| node.attribute("rgb"))
        .map(str::to_owned);
    let bold = font
        .descendants()
        .any(|node| node.is_element() && node.tag_name().name() == "b");
    let comments = renamed.comments.get("A1").cloned();
    let merges = renamed.merges.iter().cloned().collect::<BTreeSet<_>>();
    let row_height = sheet_doc
        .descendants()
        .find(|node| {
            node.is_element() && node.tag_name().name() == "row" && node.attribute("r") == Some("1")
        })
        .and_then(|node| node.attribute("ht"))
        .map(str::parse::<f64>)
        .transpose()?;
    let column_width = sheet_doc
        .descendants()
        .find(|node| {
            node.is_element()
                && node.tag_name().name() == "col"
                && node
                    .attribute("min")
                    .and_then(|value| value.parse::<u32>().ok())
                    .is_some_and(|min| min <= 2)
                && node
                    .attribute("max")
                    .and_then(|value| value.parse::<u32>().ok())
                    .is_some_and(|max| max >= 2)
        })
        .and_then(|node| node.attribute("width"))
        .map(str::parse::<f64>)
        .transpose()?;
    let orientation = sheet_doc
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "pageSetup")
        .and_then(|node| node.attribute("orientation"))
        .map(str::to_owned);
    let sheet_index = sheet_names
        .iter()
        .position(|name| name == "RenamedA")
        .context("RenamedA index missing")?;
    let print_area = workbook_doc
        .descendants()
        .find(|node| {
            node.is_element()
                && node.tag_name().name() == "definedName"
                && node.attribute("name") == Some("_xlnm.Print_Area")
                && node.attribute("localSheetId") == Some(sheet_index.to_string().as_str())
        })
        .and_then(|node| node.text())
        .map(normalize_sheet_address);
    let defined_names = workbook_doc
        .descendants()
        .filter(|node| {
            node.is_element()
                && node.tag_name().name() == "definedName"
                && node.attribute("localSheetId").is_none()
                && node.attribute("name") != Some("_xlnm.Print_Area")
        })
        .map(|node| {
            json!({
                "name": node.attribute("name").unwrap_or_default(),
                "reference": node.text().unwrap_or_default(),
            })
        })
        .collect::<Vec<_>>();

    Ok(json!({
        "sheet_names": sheet_names,
        "a1_values": a1_values,
        "formula": formula,
        "defined_names": defined_names,
        "renamed_a1": {
            "fill": color,
            "bold": bold,
            "comment": comments,
            "merged_ranges": merges,
            "row_1_height": row_height,
            "column_b_width": column_width,
            "orientation": orientation,
            "print_area": print_area,
        }
    }))
}

fn read_part(path: &Path, name: &str) -> Result<String> {
    let mut archive = ZipArchive::new(File::open(path)?)?;
    let mut part = archive
        .by_name(name)
        .with_context(|| format!("part {name} missing"))?;
    let mut xml = String::new();
    part.read_to_string(&mut xml)?;
    Ok(xml)
}

fn normalize_sheet_address(address: &str) -> String {
    let Some((sheet, range)) = address.split_once('!') else {
        return address.to_owned();
    };
    let sheet = sheet
        .strip_prefix('\'')
        .and_then(|sheet| sheet.strip_suffix('\''))
        .unwrap_or(sheet)
        .replace("''", "'");
    let needs_quotes = sheet
        .chars()
        .any(|character| !character.is_ascii_alphanumeric() && !matches!(character, '_' | '.'));
    let rendered_sheet = if needs_quotes {
        format!("'{}'", sheet.replace('\'', "''"))
    } else {
        sheet
    };
    format!("{rendered_sheet}!{range}")
}

fn read_worksheet_part(path: &Path, sheet_name: &str) -> Result<String> {
    let workbook_xml = read_part(path, "xl/workbook.xml")?;
    let relationships_xml = read_part(path, "xl/_rels/workbook.xml.rels")?;
    let workbook = roxmltree::Document::parse(&workbook_xml)?;
    let relationships = roxmltree::Document::parse(&relationships_xml)?;
    let sheet = workbook
        .descendants()
        .find(|node| {
            node.is_element()
                && node.tag_name().name() == "sheet"
                && node.attribute("name") == Some(sheet_name)
        })
        .context("target worksheet record missing")?;
    let relationship_id = sheet
        .attribute((OFFICE_REL_NS, "id"))
        .context("worksheet relationship id missing")?;
    let target = relationships
        .descendants()
        .find(|node| {
            node.is_element()
                && node.tag_name().name() == "Relationship"
                && node.attribute("Id") == Some(relationship_id)
        })
        .and_then(|node| node.attribute("Target"))
        .context("worksheet relationship target missing")?;
    let part_name = if target.starts_with('/') {
        target.trim_start_matches('/').to_owned()
    } else if target.starts_with("xl/") {
        target.to_owned()
    } else {
        format!("xl/{target}")
    };
    read_part(path, &part_name)
}
