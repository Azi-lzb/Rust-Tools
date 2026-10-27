use anyhow::{Context, Result};
use excel_editor::{
    BackendCapability, CapabilityEvidence, CapabilityStatus, EditOperation, EngineId,
    WorkbookFormat,
};
use excel_reader::{read_header_inventory, read_workbook, HeaderVariant};
use rust_xlsxwriter::Workbook;
use serde_json::Value;
use std::{
    fs,
    io::Write,
    path::{Path, PathBuf},
    process::{Command, Output, Stdio},
    sync::atomic::{AtomicU64, Ordering},
    time::{SystemTime, UNIX_EPOCH},
};

const PLATFORM: &str = "synthetic-pt-7-4-menu-test";
static NEXT_DIRECTORY_ID: AtomicU64 = AtomicU64::new(0);

struct TestDirectory(PathBuf);

impl TestDirectory {
    fn new() -> Result<Self> {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .context("system clock before Unix epoch")?
            .as_nanos();
        let serial = NEXT_DIRECTORY_ID.fetch_add(1, Ordering::Relaxed);
        let path = std::env::temp_dir().join(format!(
            "pytools-fx-header-menu-{}-{nonce}-{serial}",
            std::process::id()
        ));
        fs::create_dir_all(&path)?;
        Ok(Self(path))
    }

    fn path(&self) -> &Path {
        &self.0
    }
}

impl Drop for TestDirectory {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

#[test]
fn fx_header_cli_command_defaults_to_plan_only() -> Result<()> {
    let root = TestDirectory::new()?;

    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/fx_header/source.xlsx");
    let target = root.path().join("cli-target.xlsx");
    fs::copy(&fixture, &target)?;
    let original = fs::read(&target)?;
    let output_dir = root.path().join("cli/output");
    let log_dir = root.path().join("cli/logs");
    let plan_dir = root.path().join("cli/plans");

    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("fx-header")
        .arg(&target)
        .arg("--output-dir")
        .arg(&output_dir)
        .arg("--log-dir")
        .arg(&log_dir)
        .arg("--plan-dir")
        .arg(&plan_dir)
        .args(["--view", "user"])
        .output()?;
    assert!(
        output.status.success(),
        "fx-header CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    let result = parse_result(&stdout, "PT-7-4")?;
    assert_eq!(result["mode"], "plan_only");
    assert_eq!(result["summary"]["total_files"], 1);
    assert_eq!(result["summary"]["planned_files"], 1);
    assert!(result["events"]
        .as_array()
        .is_some_and(|events| !events.is_empty()));
    assert_eq!(
        fs::read(&target)?,
        original,
        "plan-only CLI changed its input"
    );
    let plan_path = find_file_recursively(&plan_dir, ".plan.json")?
        .context("fx-header CLI plan was not saved")?;
    let plan: Value = serde_json::from_slice(&fs::read(plan_path)?)?;
    assert_eq!(plan["operations"][0]["operation"], "replace_header_text");
    Ok(())
}

#[test]
fn fx_header_menu_creates_validated_copy_without_capability_and_skips_unverified_xls() -> Result<()>
{
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/fx_header/source.xlsx");
    let target = root.path().join("target.xlsx");
    fs::copy(&fixture, &target)?;
    let original = fs::read(&target)?;
    let unsupported = root.path().join("legacy.xls");
    fs::write(&unsupported, b"synthetic legacy workbook placeholder")?;
    let unsupported_original = fs::read(&unsupported)?;

    let output = run_menu(
        &config,
        &format!("5\n4\n{};{}\n0\n", target.display(), unsupported.display()),
        &[],
    )?;
    assert!(
        output.status.success(),
        "PT-7-4 copy menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stdout.contains("4  修改外汇工作表页眉单位"));
    assert!(stdout.contains("PT-7-4 当前仅验收 .xlsx"));
    assert!(stdout.contains("[新建副本]"));
    assert_eq!(fs::read(&target)?, original);
    assert_eq!(fs::read(&unsupported)?, unsupported_original);

    let result = parse_result(&stdout, "PT-7-4")?;
    assert_eq!(result["mode"], "create_copy");
    assert_eq!(result["summary"]["total_files"], 1);
    assert_eq!(result["summary"]["planned_files"], 0);
    assert_eq!(result["summary"]["created_copies"], 1);
    assert_eq!(result["summary"]["failed_files"], 0);
    assert!(result["events"]
        .as_array()
        .is_some_and(|events| !events.is_empty()));
    assert!(!stdout.contains("actual_engine"));
    assert!(!stderr.contains("actual_engine"));

    let output_copy = PathBuf::from(
        result["outputs"][0]
            .as_str()
            .context("PT-7-4 output copy path missing from user report")?,
    );
    assert!(output_copy.is_file());
    assert_eq!(
        output_copy.parent().unwrap().file_name().unwrap(),
        "PyTools_专用报表处理结果"
    );
    let baseline: Value = serde_json::from_slice(&fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR")).join("probe/reports/fx_header_python_baseline.json"),
    )?)?;
    assert_headers_match_baseline(&output_copy, &baseline)?;
    assert!(!root.path().join("app/output/plans/PT-7-4").exists());

    let history_path = PathBuf::from(
        result["history_report"]
            .as_str()
            .context("user run history path missing")?,
    );
    let history: Value = serde_json::from_slice(&fs::read(history_path)?)?;
    assert_eq!(history["feature_id"], "PT-7-4");
    assert_eq!(history["summary"]["planned_files"], 0);
    assert_eq!(history["summary"]["created_copies"], 1);
    let history_text = serde_json::to_string(&history)?;
    assert!(!history_text.contains("target.xlsx"));
    assert!(!history_text.contains("万元"));
    assert!(!history_text.contains("万美元"));
    Ok(())
}

#[test]
fn fx_header_menu_commits_only_with_matching_capability_and_matches_python_baseline() -> Result<()>
{
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/fx_header/source.xlsx");
    let target = root.path().join("target.xlsx");
    fs::copy(&fixture, &target)?;
    let original = fs::read(&target)?;
    let candidate = root.path().join("capability-candidate.xlsx");
    let profile = excel_editor::inspect_xlsx_profile(&target)?;
    let plan = pytools_core::fx_header::build_fx_header_plan(
        &target,
        &candidate,
        profile,
        "synthetic-pt-7-4-menu-capability",
    )?
    .plan
    .context("FX header edit plan missing")?;
    let capability_path = root.path().join("temporary-test-capability.json");
    fs::write(
        &capability_path,
        serde_json::to_vec_pretty(&fx_header_capability(&plan))?,
    )?;

    let output = run_menu(
        &config,
        &format!("5\n4\n{}\n0\n", target.display()),
        &[
            "--capabilities".to_owned(),
            capability_path.to_string_lossy().into_owned(),
            "--platform".to_owned(),
            PLATFORM.to_owned(),
            "--log-view".to_owned(),
            "developer".to_owned(),
        ],
    )?;
    assert!(
        output.status.success(),
        "PT-7-4 capability menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stdout.contains("[执行模式]"));
    assert!(stderr.contains("\"phase\":\"safe_commit\""));
    assert!(stderr.contains("\"actual_engine\":\"ooxml_patch\""));

    let result = parse_result(&stdout, "PT-7-4")?;
    assert_eq!(result["mode"], "commit_target");
    assert_eq!(result["status"], "succeeded");
    assert_eq!(result["completed_files"], 1);
    assert_eq!(result["failed_files"], 0);
    assert_eq!(result["committed_files"], 1);
    assert_eq!(result["fx_sheets"], 2);
    assert_eq!(result["modified_sheets"], 1);
    assert_eq!(result["items"][0]["execution"]["engine"], "ooxml_patch");
    assert_ne!(fs::read(&target)?, original, "target was not committed");
    let target_commit = &result["items"][0]["target_commit"];
    assert_eq!(target_commit["operation"], "replaced");
    assert!(
        target_commit["backup_path"].is_null(),
        "successful safe commit should remove its temporary backup"
    );
    assert!(
        target_commit["cleanup_warning"].is_null(),
        "successful safe commit should clean its temporary backup"
    );

    let baseline: Value = serde_json::from_slice(&fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR")).join("probe/reports/fx_header_python_baseline.json"),
    )?)?;
    assert_eq!(
        result["fx_sheets"],
        baseline["python_baseline_stats"]["fx_sheets"]
    );
    assert_eq!(
        result["modified_sheets"],
        baseline["python_baseline_stats"]["modified_sheets"]
    );
    assert_headers_match_baseline(&target, &baseline)?;

    let workbook = read_workbook(&target)?;
    let fx = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "外汇报表")
        .context("FX worksheet missing after commit")?;
    assert_eq!(
        fx.values["B2"]["value"],
        baseline["preserved_workbook_properties"]["cell_b2"]
    );
    assert_eq!(
        fx.formulas["C2"],
        baseline["preserved_workbook_properties"]["formula_c2"]
            .as_str()
            .context("baseline formula missing")?
            .trim_start_matches('=')
    );
    assert_eq!(
        serde_json::to_value(&fx.merges)?,
        baseline["preserved_workbook_properties"]["merged_ranges"]
    );
    assert_eq!(
        fx.comments.get("A1").map(String::as_str),
        baseline["preserved_workbook_properties"]["comment_a1"].as_str()
    );
    assert_eq!(
        fx.conditional_formatting.len(),
        baseline["preserved_workbook_properties"]["conditional_formatting_count"]
            .as_u64()
            .unwrap() as usize
    );
    assert_eq!(
        fx.data_validations.len(),
        baseline["preserved_workbook_properties"]["data_validation_count"]
            .as_u64()
            .unwrap() as usize
    );

    let history_path = PathBuf::from(
        result["history_report"]
            .as_str()
            .context("developer run history path missing")?,
    );
    let history: Value = serde_json::from_slice(&fs::read(history_path)?)?;
    assert!(history["events"].as_array().is_some_and(|events| {
        events.iter().any(|event| {
            event["phase"] == "engine_execution" && event["actual_engine"] == "ooxml_patch"
        }) && events.iter().any(|event| event["phase"] == "safe_commit")
    }));
    let history_text = serde_json::to_string(&history)?;
    assert!(!history_text.contains("target.xlsx"));
    assert!(!history_text.contains("万元"));
    assert!(!history_text.contains("万美元"));
    Ok(())
}

#[test]
fn fx_header_menu_off_view_saves_no_history_or_events() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/fx_header/source.xlsx");
    let target = root.path().join("target.xlsx");
    fs::copy(&fixture, &target)?;
    let original = fs::read(&target)?;

    let output = run_menu(
        &config,
        &format!("5\n4\n{}\n0\n", target.display()),
        &["--log-view".to_owned(), "off".to_owned()],
    )?;
    assert!(
        output.status.success(),
        "PT-7-4 off-view menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains(r#""created_copies": 1"#));
    assert!(!stdout.contains("history_report"));
    assert!(!stdout.contains("\"events\""));
    assert!(String::from_utf8_lossy(&output.stderr).is_empty());
    assert_eq!(fs::read(&target)?, original);
    let result = parse_result(&stdout, "PT-7-4")?;
    let copy = PathBuf::from(result["outputs"][0].as_str().context("copy path missing")?);
    assert!(copy.is_file());
    assert_eq!(result["mode"], "create_copy");
    let log_dir = root.path().join("app/logs");
    assert!(!log_dir.exists() || fs::read_dir(log_dir)?.next().is_none());
    Ok(())
}

fn fx_header_capability(plan: &excel_editor::ExcelEditPlan) -> BackendCapability {
    BackendCapability {
        engine: EngineId::OoxmlPatch,
        status: CapabilityStatus::Passed,
        format: WorkbookFormat::Xlsx,
        operations: plan.operations.iter().map(EditOperation::kind).collect(),
        supported_features: plan.workbook.features_present.clone(),
        preserved: plan.preserve.clone(),
        platform: PLATFORM.to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: "synthetic-pt-7-4-menu".to_owned(),
            report_path: "pytools-core/tests/fx_header_menu_cli.rs".to_owned(),
            adapter_version: "excel-editor/0.1.0 + ooxml-patch/0.1.0".to_owned(),
            engine_version: "ooxml-patch/0.1.0 + zip/6".to_owned(),
            platform: PLATFORM.to_owned(),
        }),
    }
}

fn create_config(root: &Path) -> Result<PathBuf> {
    let app = root.join("app");
    let config = app.join("config/config.xlsx");
    fs::create_dir_all(config.parent().context("config parent missing")?)?;
    fs::create_dir_all(app.join("input"))?;
    let mut workbook = Workbook::new();
    let global = workbook.add_worksheet();
    global.set_name("全局配置")?;
    global.write_string(0, 0, "配置项")?;
    global.write_string(0, 1, "值")?;
    for (index, (key, value)) in [
        ("输入目录", app.join("input").to_string_lossy().into_owned()),
        (
            "输出目录",
            app.join("output").to_string_lossy().into_owned(),
        ),
        ("日志目录", app.join("logs").to_string_lossy().into_owned()),
    ]
    .iter()
    .enumerate()
    {
        global.write_string((index + 1) as u32, 0, *key)?;
        global.write_string((index + 1) as u32, 1, value)?;
    }
    let mapping = workbook.add_worksheet();
    mapping.set_name("机构映射表")?;
    mapping.write_string(0, 0, "原始机构")?;
    mapping.write_string(0, 1, "标准机构")?;
    workbook.save(&config)?;
    Ok(config)
}

fn run_menu(config: &Path, input: &str, extra: &[String]) -> Result<Output> {
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("menu")
        .arg("--config")
        .arg(config)
        .args(extra)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()?;
    child
        .stdin
        .take()
        .context("menu stdin missing")?
        .write_all(input.as_bytes())?;
    Ok(child.wait_with_output()?)
}

fn parse_result(output: &str, feature_id: &str) -> Result<Value> {
    let marker = format!("\"feature_id\": \"{feature_id}\"");
    let marker_offset = output
        .find(&marker)
        .with_context(|| format!("result marker {feature_id} missing from CLI output"))?;
    let bytes = output.as_bytes();
    for start in (0..=marker_offset)
        .rev()
        .filter(|offset| bytes[*offset] == b'{')
    {
        let mut depth = 0usize;
        let mut in_string = false;
        let mut escaped = false;
        for (offset, byte) in bytes[start..].iter().copied().enumerate() {
            if in_string {
                if escaped {
                    escaped = false;
                } else if byte == b'\\' {
                    escaped = true;
                } else if byte == b'"' {
                    in_string = false;
                }
                continue;
            }
            match byte {
                b'"' => in_string = true,
                b'{' => depth += 1,
                b'}' => {
                    depth -= 1;
                    if depth == 0 {
                        if let Ok(value) =
                            serde_json::from_slice::<Value>(&bytes[start..=start + offset])
                        {
                            if value["feature_id"] == feature_id {
                                return Ok(value);
                            }
                        }
                        break;
                    }
                }
                _ => {}
            }
        }
    }
    anyhow::bail!("could not find result object for feature {feature_id}")
}

fn assert_headers_match_baseline(path: &Path, baseline: &Value) -> Result<()> {
    let actual = read_header_inventory(path)?;
    for (sheet_name, variants) in baseline["headers_after"]
        .as_object()
        .context("baseline headers missing")?
    {
        for variant in ["odd", "even", "first"] {
            let Some(parts) = variants[variant].as_object() else {
                anyhow::bail!("baseline header variant missing: {sheet_name}/{variant}");
            };
            let mut expected = String::new();
            for (section, marker) in [("left", "&L"), ("center", "&C"), ("right", "&R")] {
                if let Some(text) = parts[section].as_str() {
                    expected.push_str(marker);
                    expected.push_str(text);
                }
            }
            if expected.is_empty() {
                continue;
            }
            let expected_variant = match variant {
                "odd" => HeaderVariant::Odd,
                "even" => HeaderVariant::Even,
                _ => HeaderVariant::First,
            };
            assert!(
                actual.iter().any(|sheet| {
                    sheet.sheet_name == sheet_name.as_str()
                        && sheet.headers.iter().any(|header| {
                            header.variant == expected_variant && header.text == expected
                        })
                }),
                "header mismatch at {sheet_name}/{variant}"
            );
        }
    }
    Ok(())
}

fn find_file_recursively(root: &Path, suffix: &str) -> Result<Option<PathBuf>> {
    for entry in fs::read_dir(root).with_context(|| format!("read directory {}", root.display()))? {
        let path = entry?.path();
        if path.is_dir() {
            if let Some(found) = find_file_recursively(&path, suffix)? {
                return Ok(Some(found));
            }
        } else if path
            .file_name()
            .and_then(|name| name.to_str())
            .is_some_and(|name| name.ends_with(suffix))
        {
            return Ok(Some(path));
        }
    }
    Ok(None)
}
