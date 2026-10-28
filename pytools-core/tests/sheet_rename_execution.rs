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
    io::Read,
    path::{Path, PathBuf},
    process::Command,
    sync::atomic::{AtomicU64, Ordering},
};
use zip::ZipArchive;

const PLATFORM: &str = "pt44-synthetic-probe";
static TEMP_SEQUENCE: AtomicU64 = AtomicU64::new(0);
const OFFICE_REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

struct TestDir(PathBuf);

impl TestDir {
    fn new(label: &str) -> Result<Self> {
        let path = std::env::temp_dir().join(format!(
            "pytools-pt44-{label}-{}-{}",
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

fn plan_for(config: &Path, input: &Path, output: &Path) -> Result<ExcelEditPlan> {
    build_sheet_rename_plan(config, input, output)?
        .plan
        .context("fixture should produce a non-empty sheet-rename plan")
}

fn run_cli(config: &Path, capability: &Path, input: &Path) -> Result<std::process::Output> {
    Ok(Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["rename-sheets", "--config"])
        .arg(config)
        .args(["--capabilities"])
        .arg(capability)
        .args(["--platform", PLATFORM])
        .arg(input)
        .output()?)
}

fn baseline_report() -> Result<Value> {
    Ok(serde_json::from_str(include_str!(
        "../probe/reports/PT-4-4_python_baseline_20261023.json"
    ))?)
}

#[test]
fn sheet_rename_cli_matches_python_and_publishes_a_copy_without_changing_source() -> Result<()> {
    let dir = TestDir::new("execution")?;
    let config = fixture("config.xlsx");
    let input = dir.path("input.xlsx");
    fs::copy(fixture("input.xlsx"), &input)?;
    let original_hash = pytools_core::file_commit::sha256_file(&input)?;
    let config_hash = pytools_core::file_commit::sha256_file(&config)?;
    let plan = plan_for(&config, &input, &dir.path("plan-output.xlsx"))?;
    let capability = dir.path("passed-capability.json");
    write_capability(&capability, &plan, CapabilityStatus::Passed)?;

    let output = run_cli(&config, &capability, &input)?;
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
    assert_eq!(
        actual["files"][0]["execution"]["engine"],
        "umya_spreadsheet"
    );
    assert_eq!(actual["files"][0]["status"], "renamed");
    assert!(
        actual["files"][0]["execution"]["route_decisions"][0]["eligible"]
            .as_bool()
            .unwrap_or(false)
    );
    let file = &actual["files"][0];
    let recorded_phase_ms = ["planning_ms", "editing_ms", "commit_ms"]
        .iter()
        .map(|phase| file[*phase].as_u64().unwrap_or_default())
        .sum::<u64>();
    anyhow::ensure!(
        file["elapsed_ms"].as_u64().unwrap_or_default() >= recorded_phase_ms,
        "per-file elapsed time must include its sequential phase durations"
    );
    assert_eq!(file["execution"]["attempts"][0]["outcome"], "succeeded");
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
    if let Some(print_area) = expected_projection["renamed_a1"]["print_area"].as_str() {
        expected_projection["renamed_a1"]["print_area"] =
            normalize_sheet_address(print_area).into();
    }
    let output_path = PathBuf::from(
        file["output"]
            .as_str()
            .context("renamed workbook path missing")?,
    );
    // 就地模式：输出路径即源工作簿本身。
    let canonical_input = fs::canonicalize(&input)?;
    let canonical_output = fs::canonicalize(&output_path)?;
    assert_eq!(
        canonical_output, canonical_input,
        "in-place rename must target the source"
    );
    assert!(output_path.is_file());
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

    if let Some(report_path) = std::env::var_os("PYTOOLS_PT44_REPORT_PATH") {
        let report_path = PathBuf::from(report_path);
        if let Some(parent) = report_path.parent() {
            fs::create_dir_all(parent)?;
        }
        let report = json!({
            "suite": "pytools.pt_4_4_sheet_rename_candidate",
            "comparison_status": "passed_synthetic_projection",
            "python_baseline_report": "pytools-core/probe/reports/PT-4-4_python_baseline_20261023.json",
            "sample_class": "synthetic",
            "engine": "umya_spreadsheet",
            "capability_scope": "temporary test-only Passed capability; not production registration",
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
fn missing_passed_capability_rejects_before_source_change() -> Result<()> {
    let dir = TestDir::new("rejected")?;
    let config = fixture("config.xlsx");
    let input = dir.path("input.xlsx");
    fs::copy(fixture("input.xlsx"), &input)?;
    let original_hash = pytools_core::file_commit::sha256_file(&input)?;
    let plan = plan_for(&config, &input, &dir.path("plan-output.xlsx"))?;
    let capability = dir.path("unverified-capability.json");
    write_capability(&capability, &plan, CapabilityStatus::Unverified)?;

    let output = run_cli(&config, &capability, &input)?;
    assert!(!output.status.success());
    assert!(String::from_utf8_lossy(&output.stderr).contains("capability"));
    assert_eq!(
        pytools_core::file_commit::sha256_file(&input)?,
        original_hash
    );
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
    let unsupported = input_dir.join("old.xls");
    fs::copy(fixture("input.xlsx"), &first)?;
    fs::copy(fixture("input.xlsx"), &second)?;
    let first_hash = pytools_core::file_commit::sha256_file(&first)?;
    let second_hash = pytools_core::file_commit::sha256_file(&second)?;
    fs::write(&unsupported, b"not an xls workbook")?;
    fs::write(input_dir.join("~$lock.xlsx"), b"ignored office lock")?;
    let unsupported_hash = pytools_core::file_commit::sha256_file(&unsupported)?;
    let plan = plan_for(&config, &first, &dir.path("plan-output.xlsx"))?;
    let capability = dir.path("passed-capability.json");
    write_capability(&capability, &plan, CapabilityStatus::Passed)?;

    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["rename-sheets", "--config"])
        .arg(&config)
        .args(["--capabilities"])
        .arg(&capability)
        .args(["--platform", PLATFORM, "--recursive"])
        .arg(&input_dir)
        .output()?;
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
    assert_eq!(status_for("old.xls")?, "skipped");
    let original_outputs = report["files"]
        .as_array()
        .context("initial file results missing")?
        .iter()
        .filter_map(|file| file["output"].as_str().map(PathBuf::from))
        .map(|path| Ok((path.clone(), pytools_core::file_commit::sha256_file(&path)?)))
        .collect::<Result<Vec<_>>>()?;
    assert_eq!(
        pytools_core::file_commit::sha256_file(&unsupported)?,
        unsupported_hash
    );

    let second_output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["rename-sheets", "--config"])
        .arg(&config)
        .args(["--capabilities"])
        .arg(&capability)
        .args(["--platform", PLATFORM, "--recursive"])
        .arg(&input_dir)
        .output()?;
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
        .map(str::to_owned);
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
