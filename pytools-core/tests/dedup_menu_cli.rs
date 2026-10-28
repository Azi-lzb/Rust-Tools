use anyhow::{Context, Result};
use excel_editor::{
    BackendCapability, CapabilityEvidence, CapabilityStatus, EditOperation, EngineId, ExcelEditPlan,
};
use excel_reader::read_workbook;
use pytools_core::dedup_ops::AppendTargetPlan;
use serde::Deserialize;
use serde_json::Value;
use std::{
    fs,
    io::Write,
    path::{Path, PathBuf},
    process::{Command, Output, Stdio},
    sync::atomic::{AtomicU64, Ordering},
    time::{SystemTime, UNIX_EPOCH},
};

static NEXT_TEST_DIRECTORY_ID: AtomicU64 = AtomicU64::new(0);

struct TestDirectory(PathBuf);

impl TestDirectory {
    fn new() -> Result<Self> {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .context("system clock before Unix epoch")?
            .as_nanos();
        let serial = NEXT_TEST_DIRECTORY_ID.fetch_add(1, Ordering::Relaxed);
        let path = std::env::temp_dir().join(format!(
            "pytools-dedup-menu-{}-{nonce}-{serial}",
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
fn interactive_dedup_menu_routes_four_v2_actions_without_committing_edits() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_dedup_config(root.path())?;
    let source = root.path().join("source.xlsx");
    let target = root.path().join("target.xlsx");
    create_pure_data_duplicate_source(&source)?;
    let source_before = fs::read(&source)?;

    let menu = run_menu(&config, "2\n0\n0\n")?;
    assert!(menu.status.success());
    let menu_text = String::from_utf8_lossy(&menu.stdout);
    assert!(menu_text.contains("按配置检查重复"));
    assert!(!menu_text.contains("按批注"));

    let source_dump = read_workbook(&source)?;
    let config_check = run_menu(&config, "2\n1\n0\n")?;
    assert!(config_check.status.success());
    let config_check_stdout = String::from_utf8_lossy(&config_check.stdout);
    assert!(
        config_check_stdout.contains("处理完成：任务成功 1 项，跳过 0 项，失败 0 项"),
        "{config_check_stdout}"
    );
    assert!(
        config_check_stdout.contains("标记重复行 1 行"),
        "{config_check_stdout}"
    );
    assert!(config_check_stdout.contains("已就地标记"), "{config_check_stdout}");
    assert!(
        !config_check_stdout.contains("\"execution\""),
        "menu must not print JSON"
    );
    // 标记直接写入源工作簿：值不变、重复行整行标红、数字格式保留。
    let checked = read_workbook(&source)?;
    assert_eq!(checked.sheets.len(), 1);
    assert_eq!(checked.sheets[0].name, "Commented");
    assert_eq!(checked.sheets[0].values, source_dump.sheets[0].values);
    for address in ["A3", "B3", "C3"] {
        assert!(
            cell_fill(&source, address)?.is_some_and(|color| color.ends_with("FF0000")),
            "duplicate row cell {address} should have red fill in the source"
        );
    }
    assert_eq!(cell_number_format(&source, "B2")?.as_deref(), Some("000000"));
    assert_eq!(cell_number_format(&source, "B3")?.as_deref(), Some("000000"));
    assert!(cell_fill(&source, "A2")?.is_none());
    assert!(cell_fill(&source, "D3")?.is_none());
    assert_ne!(
        fs::read(&source)?,
        source_before,
        "in-place marking must modify the source workbook"
    );

    let config_delete = run_menu(&config, "2\n2\n0\n")?;
    assert!(config_delete.status.success());
    let delete_stdout = String::from_utf8_lossy(&config_delete.stdout);
    assert!(
        delete_stdout.contains("处理完成：任务成功 1 项，跳过 0 项，失败 0 项"),
        "{delete_stdout}"
    );
    assert!(delete_stdout.contains("删除重复行 1 行"), "{delete_stdout}");
    assert!(delete_stdout.contains("已就地删除"), "{delete_stdout}");
    assert!(
        !delete_stdout.contains("\"execution\""),
        "menu must not print JSON"
    );
    // 就地删除：源工作簿被直接修改，重复行移除、保留行上移。
    let deleted = read_workbook(&source)?;
    assert_eq!(deleted.sheets.len(), 1);
    assert_eq!(deleted.sheets[0].name, "Commented");
    assert_eq!(
        deleted.sheets[0].values.len(),
        source_dump.sheets[0].values.len() - 3
    );
    for address in ["A1", "B1", "C1", "A2", "B2", "C2"] {
        assert_eq!(
            deleted.sheets[0].values.get(address),
            source_dump.sheets[0].values.get(address),
            "nonduplicate row cell {address} should be preserved"
        );
    }
    for (after_address, before_address) in [("A3", "A4"), ("B3", "B4"), ("C3", "C4")] {
        assert_eq!(
            deleted.sheets[0].values.get(after_address),
            source_dump.sheets[0].values.get(before_address),
            "surviving rows should shift upward after in-place deletion"
        );
    }
    assert_eq!(
        cell_number_format(&source, "B2")?.as_deref(),
        Some("000000")
    );
    assert!(!deleted.sheets[0].values.contains_key("A4"));
    // 不再产生副本目录或临时文件。
    let parent = source.parent().context("source parent missing")?;
    assert!(!fs::read_dir(parent)?.filter_map(|entry| entry.ok()).any(|entry| {
        entry
            .file_name()
            .to_string_lossy()
            .contains(".pytools-dedup-")
    }));

    // 菜单不再打印 JSON；追加/预校验的机器可读结果通过 CLI 子命令验证。
    let append_output_dir = root.path().join("append-out");
    let append_plan = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-append-plan", "--config"])
        .arg(&config)
        .args(["--output-dir"])
        .arg(&append_output_dir)
        .output()?;
    assert!(
        append_plan.status.success(),
        "append plan failed: {}",
        String::from_utf8_lossy(&append_plan.stderr)
    );
    let append_json = output_json(&append_plan)?;
    assert_eq!(append_json[0]["target_existed"], false);
    let append_plan_file = PathBuf::from(
        append_json[0]["plan_file"]
            .as_str()
            .context("append plan path missing")?,
    );
    assert!(append_plan_file.is_file());
    assert!(!target.exists());

    let precheck = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-precheck", "--config"])
        .arg(&config)
        .output()?;
    assert!(precheck.status.success());
    let precheck_json = output_json(&precheck)?;
    assert_eq!(precheck_json["fail"], 0);
    assert_eq!(precheck_json["ok"], 1);

    // 2-1/2-2 已改为就地修改：源工作簿应被标记/删除改写，
    // 而配置与目标工作簿保持不变。
    assert_ne!(
        fs::read(&source)?,
        source_before,
        "in-place dedup must modify the source workbook"
    );
    let log_dir = root.path().join("logs");
    assert!(!log_dir.exists() || fs::read_dir(log_dir)?.next().is_none());
    Ok(())
}

#[test]
fn configured_duplicate_check_marks_formula_workbook_in_place() -> Result<()> {
    // 2026-10-28 起 2-1 直接在源工作簿上标红标记重复行，不再拒绝公式工作簿。
    let root = TestDirectory::new()?;
    let config = create_dedup_config(root.path())?;
    let source = root.path().join("source.xlsx");
    create_duplicate_source_with_formula(&source)?;
    let source_before = fs::read(&source)?;

    let output = run_menu(&config, "2\n1\n0\n")?;
    assert!(output.status.success());
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("处理完成：任务成功"), "{stdout}");
    assert!(!stdout.contains("\"execution\""), "menu must not print JSON");
    assert_ne!(
        fs::read(&source)?,
        source_before,
        "in-place marking must modify the source workbook"
    );
    Ok(())
}

#[test]
fn configured_duplicate_check_marks_comment_workbook_in_place() -> Result<()> {
    // 带批注的工作簿同样就地标记；检查本身不删除批注或内容。
    let root = TestDirectory::new()?;
    let config = create_dedup_config(root.path())?;
    let source = root.path().join("source.xlsx");
    create_duplicate_source(&source)?;
    let source_before = fs::read(&source)?;

    let output = run_menu(&config, "2\n1\n0\n")?;
    assert!(output.status.success());
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("处理完成：任务成功"), "{stdout}");
    assert_ne!(fs::read(&source)?, source_before);
    Ok(())
}

#[test]
fn configured_duplicate_delete_removes_duplicates_in_place() -> Result<()> {
    // 2026-10-28 起 2-2 直接在源工作簿上删除重复行（保留首次出现）。
    let root = TestDirectory::new()?;
    let config = create_dedup_config(root.path())?;
    let source = root.path().join("source.xlsx");
    create_duplicate_source_with_formula(&source)?;
    let source_before = fs::read(&source)?;

    let output = run_menu(&config, "2\n2\n0\n")?;
    assert!(output.status.success());
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("处理完成：任务成功"), "{stdout}");
    assert!(stdout.contains("已就地删除"), "{stdout}");
    assert!(!stdout.contains("\"execution\""), "menu must not print JSON");
    assert_ne!(
        fs::read(&source)?,
        source_before,
        "in-place deletion must modify the source workbook"
    );
    Ok(())
}


fn cell_fill(path: &Path, address: &str) -> Result<Option<String>> {
    let workbook = excel_editor::open_xlsx(path)?;
    let sheet = workbook
        .sheet_by_name("Commented")
        .context("duplicate check worksheet missing")?;
    Ok(sheet
        .cell(address)
        .and_then(|cell| cell.style().fill())
        .and_then(|fill| fill.pattern_fill())
        .and_then(|fill| fill.foreground_color())
        .map(|color| color.argb_str().to_ascii_uppercase()))
}

fn cell_number_format(path: &Path, address: &str) -> Result<Option<String>> {
    let workbook = excel_editor::open_xlsx(path)?;
    let sheet = workbook
        .sheet_by_name("Commented")
        .context("duplicate result worksheet missing")?;
    Ok(sheet
        .cell(address)
        .and_then(|cell| cell.style().number_format())
        .map(|format| format.format_code().to_owned()))
}

#[test]
fn interactive_dedup_append_requires_confirmation_then_creates_a_new_target() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_dedup_config(root.path())?;
    let source = root.path().join("source.xlsx");
    let target = root.path().join("target.xlsx");
    create_reviewed_intermediate(&source)?;
    let source_before = fs::read(&source)?;

    let menu_input = "2\n3\nyes\n0\n";
    let output = run_menu_with_view(&config, &menu_input, "developer")?;
    assert!(
        output.status.success(),
        "confirmed append failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(!stdout.contains("选择已人工检查的时序宽表中间工作簿"));
    let report = output_json(&output)?;
    assert_eq!(report[0]["status"], "succeeded");
    assert_eq!(report[0]["execution"], "committed");
    assert_eq!(report[0]["target_existed"], false);
    assert_eq!(report[0]["appended_rows"], 2);
    assert_eq!(report[0]["duplicate_rows_filtered"], 1);
    assert_eq!(
        report[0]["execution_result"]["target_commit"]["operation"],
        "created"
    );
    assert!(target.is_file());
    let candidate = PathBuf::from(
        report[0]["candidate"]
            .as_str()
            .context("append candidate path missing")?,
    );
    assert!(candidate.is_file());
    assert_eq!(fs::read(&target)?, fs::read(&candidate)?);
    let appended = read_workbook(&target)?;
    let appended_sheet = appended
        .sheets
        .iter()
        .find(|sheet| sheet.name == "Append")
        .context("append target worksheet missing")?;
    assert_eq!(appended_sheet.values["B2"]["value"], "checked-1");
    assert_eq!(appended_sheet.values["B3"]["value"], "checked-2");
    assert_eq!(fs::read(&source)?, source_before);

    let log_dir = root.path().join("logs");
    let histories = fs::read_dir(&log_dir)?
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .filter(|path| {
            path.extension()
                .is_some_and(|extension| extension == "json")
        })
        .collect::<Vec<_>>();
    assert_eq!(histories.len(), 1);
    let history_text = fs::read_to_string(&histories[0])?;
    let history: Value = serde_json::from_str(&history_text)?;
    assert_eq!(history["feature_id"], "PT-2-5");
    assert_eq!(history["status"], "succeeded");
    assert_eq!(history["excel_writer_invoked"], true);
    assert_eq!(history["excel_modifier_invoked"], false);
    assert_eq!(history["plan_only"], false);
    assert!(history["actual_engines"]
        .as_array()
        .is_some_and(|engines| engines.iter().any(|engine| engine == "rust_xlsxwriter")));
    assert!(history["stages"].as_array().is_some_and(|stages| stages
        .iter()
        .any(|stage| stage["name"] == "new_workbook_write_and_reopen")));
    assert!(!history_text.contains("source.xlsx"));
    assert!(!history_text.contains(&root.path().to_string_lossy().to_string()));
    Ok(())
}

#[test]
fn interactive_dedup_append_keeps_existing_target_plan_only_without_capability() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_dedup_config(root.path())?;
    let source = root.path().join("source.xlsx");
    let target = root.path().join("target.xlsx");
    create_duplicate_source(&source)?;
    create_initial_target(&target)?;
    let source_before = fs::read(&source)?;
    let target_before = fs::read(&target)?;

    let menu_input = "2\n3\n0\n";
    let output = run_menu_with_view(&config, menu_input, "developer")?;
    assert!(
        output.status.success(),
        "plan-only append failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert!(String::from_utf8_lossy(&output.stdout).contains("PT-2 受阻"));
    let report = output_json(&output)?;
    assert_eq!(report[0]["target_existed"], true);
    assert_eq!(report[0]["status"], "planned_only");
    assert_eq!(
        report[0]["execution"],
        "requires ExcelModifier capability and platform"
    );
    assert!(PathBuf::from(
        report[0]["plan_file"]
            .as_str()
            .context("existing-target append plan path missing")?
    )
    .is_file());
    assert!(!PathBuf::from(
        report[0]["candidate"]
            .as_str()
            .context("append candidate path missing")?
    )
    .exists());
    assert_eq!(fs::read(&target)?, target_before);
    assert_eq!(fs::read(&source)?, source_before);
    let history_path = fs::read_dir(root.path().join("logs"))?
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .find(|path| {
            path.file_name()
                .and_then(|name| name.to_str())
                .is_some_and(|name| name.starts_with("pytools-dedup-config_append-"))
        })
        .context("blocked append history missing")?;
    let history: Value = serde_json::from_slice(&fs::read(history_path)?)?;
    assert_eq!(history["status"], "blocked");
    assert_eq!(history["error_type"], "excel_modifier_capability_required");
    assert_eq!(history["plan_only"], true);
    Ok(())
}

#[test]
fn interactive_dedup_append_rejects_whitespace_header_mismatch() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_dedup_config(root.path())?;
    let source = root.path().join("source.xlsx");
    let target = root.path().join("target.xlsx");
    create_duplicate_source(&source)?;
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Append")?;
    sheet.write_string(0, 0, " 源数据")?;
    sheet.write_string(0, 1, "ID")?;
    sheet.write_string(0, 2, "Value")?;
    sheet.write_string(1, 0, "existing")?;
    workbook.save(&target)?;
    let source_before = fs::read(&source)?;
    let target_before = fs::read(&target)?;

    let menu_input = "2\n3\n0\n";
    // 运行记录只在 developer 视图产出；菜单本身不再打印 JSON。
    let output = run_menu_with_view(&config, &menu_input, "developer")?;
    assert!(
        output.status.success(),
        "header mismatch should skip the append task: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report = output_json(&output)?;
    assert_eq!(report[0]["appended_rows"], 0);
    assert_eq!(report[0]["task_ok"], 0);
    assert!(report[0]["issues"].as_array().is_some_and(|issues| {
        issues.iter().any(|issue| {
            issue
                .as_str()
                .is_some_and(|message| message.contains("header_mismatch"))
        })
    }));
    assert_eq!(fs::read(&source)?, source_before);
    assert_eq!(fs::read(&target)?, target_before);
    Ok(())
}

#[test]
fn interactive_dedup_append_commits_existing_target_with_matching_capability() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_dedup_config(root.path())?;
    let source = root.path().join("source.xlsx");
    let target = root.path().join("target.xlsx");
    let capability_path = root.path().join("test-capability.json");
    create_duplicate_source(&source)?;
    create_initial_target(&target)?;
    let source_before = fs::read(&source)?;
    let target_before = fs::read(&target)?;

    let planned = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-append-plan", "--config"])
        .arg(&config)
        .output()
        .context("plan existing-target append for its synthetic capability")?;
    assert!(
        planned.status.success(),
        "planning failed: {}",
        String::from_utf8_lossy(&planned.stderr)
    );
    let plan_report: Value = serde_json::from_slice(&planned.stdout)?;
    let plan_path = PathBuf::from(
        plan_report[0]["plan_file"]
            .as_str()
            .context("append plan path missing")?,
    );
    let plan: AppendTargetPlan = serde_json::from_slice(&fs::read(plan_path)?)?;
    let edit_plan = plan.edit_plan.as_ref().context("edit plan missing")?;
    fs::write(
        &capability_path,
        serde_json::to_vec_pretty(&capability_for(edit_plan))?,
    )?;

    let menu_input = "2\n3\nyes\n0\n";
    let output = run_menu_with_capability(
        &config,
        menu_input,
        &capability_path,
        "synthetic-pt-2-5-menu-test",
    )?;
    assert!(
        output.status.success(),
        "confirmed existing-target append failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report = output_json(&output)?;
    assert_eq!(report[0]["target_existed"], true);
    assert_eq!(report[0]["status"], "succeeded");
    assert_eq!(report[0]["execution"], "committed");
    assert_eq!(report[0]["execution_result"]["engine"], "umya_spreadsheet");
    assert_eq!(
        report[0]["execution_result"]["target_commit"]["operation"],
        "replaced"
    );
    let candidate = PathBuf::from(
        report[0]["candidate"]
            .as_str()
            .context("append candidate path missing")?,
    );
    assert!(candidate.is_file());
    assert_eq!(fs::read(&target)?, fs::read(candidate)?);
    assert_ne!(fs::read(&target)?, target_before);
    assert_eq!(fs::read(&source)?, source_before);
    Ok(())
}

#[test]
fn interactive_dedup_menu_reports_in_place_marking_without_json() -> Result<()> {
    // 2-1 就地标记：输出可读行，不再产出旧 PT-2 计划/JSON 输出。
    let root = TestDirectory::new()?;
    let config = create_dedup_config(root.path())?;
    let source = root.path().join("source.xlsx");
    create_pure_data_duplicate_source(&source)?;

    let developer = run_menu_with_view(&config, "2\n1\n0\n", "developer")?;
    assert!(
        developer.status.success(),
        "developer menu failed: {}",
        String::from_utf8_lossy(&developer.stderr)
    );
    let stdout = String::from_utf8_lossy(&developer.stdout);
    assert!(stdout.contains("处理完成：任务成功 1 项"), "{stdout}");
    assert!(stdout.contains("标记重复行 1 行"), "{stdout}");
    assert!(!stdout.contains("\"execution\""), "{stdout}");

    let off = run_menu_with_view(&config, "2\n1\n0\n", "off")?;
    assert!(off.status.success());
    Ok(())
}


fn create_duplicate_source(path: &Path) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Commented")?;
    sheet.write_string(0, 0, "源数据")?;
    sheet.write_string(0, 1, "ID")?;
    sheet.write_string(0, 2, "Value")?;
    sheet.insert_note(0, 0, &rust_xlsxwriter::Note::new("本表为源数据"))?;
    sheet.insert_note(0, 1, &rust_xlsxwriter::Note::new("标识列"))?;
    for (row, id, value) in [(1, "A", 10.0), (2, "A", 20.0), (3, "B", 30.0)] {
        sheet.write_string(row, 0, id)?;
        sheet.write_string(row, 1, id)?;
        sheet.write_number(row, 2, value)?;
    }
    workbook.save(path)?;
    Ok(())
}

fn create_reviewed_intermediate(path: &Path) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Commented")?;
    for (column, header) in ["数据来源", "机构编号", "金额"].iter().enumerate() {
        sheet.write_string(0, column as u16, *header)?;
    }
    for (row, id, amount) in [
        (1, "checked-1", 10.0),
        (2, "checked-1", 20.0),
        (3, "checked-2", 30.0),
    ] {
        sheet.write_string(row, 0, "人工检查后的中间表")?;
        sheet.write_string(row, 1, id)?;
        sheet.write_number(row, 2, amount)?;
    }
    workbook.save(path)?;
    Ok(())
}

fn create_pure_data_duplicate_source(path: &Path) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Commented")?;
    sheet.write_string(0, 0, "源数据")?;
    sheet.write_string(0, 1, "ID")?;
    sheet.write_string(0, 2, "Value")?;
    let code_format = rust_xlsxwriter::Format::new().set_num_format("000000");
    for (row, id, value) in [(1, 123.0, 10.0), (2, 123.0, 20.0), (3, 456.0, 30.0)] {
        sheet.write_string(row, 0, format!("row-{row}"))?;
        sheet.write_number_with_format(row, 1, id, &code_format)?;
        sheet.write_number(row, 2, value)?;
    }
    workbook.save(path)?;
    Ok(())
}

fn create_duplicate_source_with_formula(path: &Path) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Commented")?;
    sheet.write_string(0, 0, "源数据")?;
    sheet.write_string(0, 1, "ID")?;
    sheet.write_string(0, 2, "Value")?;
    for (row, id, formula, result) in [
        (1, "A", "5+5", "10"),
        (2, "A", "10+10", "20"),
        (3, "B", "15+15", "30"),
    ] {
        sheet.write_string(row, 0, id)?;
        sheet.write_string(row, 1, id)?;
        sheet.write_formula(
            row,
            2,
            rust_xlsxwriter::Formula::new(formula).set_result(result),
        )?;
    }
    workbook.save(path)?;
    Ok(())
}

fn create_initial_target(path: &Path) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Append")?;
    sheet.write_string(0, 0, "源数据")?;
    sheet.write_string(0, 1, "ID")?;
    sheet.write_string(0, 2, "Value")?;
    sheet.write_string(1, 0, "existing")?;
    sheet.write_string(1, 1, "existing")?;
    workbook.save(path)?;
    Ok(())
}

fn capability_for(plan: &ExcelEditPlan) -> BackendCapability {
    BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Passed,
        format: plan.format,
        operations: plan.operations.iter().map(EditOperation::kind).collect(),
        supported_features: plan.workbook.features_present.clone(),
        preserved: plan.preserve.clone(),
        platform: "synthetic-pt-2-5-menu-test".to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: "synthetic-pt-2-5-existing-target".to_owned(),
            report_path: "pytools-core/tests/dedup_menu_cli.rs".to_owned(),
            adapter_version: "umya-spreadsheet-3.1.1".to_owned(),
            engine_version: "umya-spreadsheet-3.1.1".to_owned(),
            platform: "synthetic-pt-2-5-menu-test".to_owned(),
        }),
    }
}

fn create_dedup_config(root: &Path) -> Result<PathBuf> {
    let config = root.join("config.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    {
        let global = workbook.add_worksheet();
        global.set_name("全局配置")?;
        global.write_string(0, 0, "配置项")?;
        global.write_string(0, 1, "值")?;
        for (row, (key, value)) in [
            ("输入目录", root.to_string_lossy().into_owned()),
            (
                "输出目录",
                root.join("output").to_string_lossy().into_owned(),
            ),
            ("日志目录", root.join("logs").to_string_lossy().into_owned()),
            ("错误策略", "fail_fast".to_owned()),
        ]
        .iter()
        .enumerate()
        {
            global.write_string((row + 1) as u32, 0, *key)?;
            global.write_string((row + 1) as u32, 1, value)?;
        }
    }
    {
        let tasks = workbook.add_worksheet();
        tasks.set_name("去重追加数据配置")?;
        let headers = [
            "是否启用",
            "任务名",
            "源数据工作簿",
            "源数据工作表",
            "标识列序号",
            "目标工作簿",
            "目标工作表",
            "执行模式",
        ];
        for (column, header) in headers.iter().enumerate() {
            tasks.write_string(0, column as u16, *header)?;
        }
        let values = [
            "是".to_owned(),
            "DedupMenu".to_owned(),
            root.join("source.xlsx").to_string_lossy().into_owned(),
            "Commented".to_owned(),
            "2".to_owned(),
            root.join("target.xlsx").to_string_lossy().into_owned(),
            "Append".to_owned(),
            "追加".to_owned(),
        ];
        for (column, value) in values.iter().enumerate() {
            tasks.write_string(1, column as u16, value.as_str())?;
        }
    }
    workbook.save(&config)?;
    Ok(config)
}

fn run_menu(config: &Path, input: &str) -> Result<Output> {
    run_menu_with_view(config, input, "user")
}

fn run_menu_with_view(config: &Path, input: &str, log_view: &str) -> Result<Output> {
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("menu")
        .arg("--config")
        .arg(config)
        .arg("--log-view")
        .arg(log_view)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .context("start interactive pytools menu")?;
    child
        .stdin
        .take()
        .context("interactive stdin missing")?
        .write_all(input.as_bytes())
        .context("send interactive menu selections")?;
    child
        .wait_with_output()
        .context("wait for interactive menu")
}

fn run_menu_with_capability(
    config: &Path,
    input: &str,
    capability: &Path,
    platform: &str,
) -> Result<Output> {
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("menu")
        .arg("--config")
        .arg(config)
        .arg("--capabilities")
        .arg(capability)
        .arg("--platform")
        .arg(platform)
        // 菜单默认（user 视图）不打印 JSON；运行记录与机器可读结果走 developer 视图。
        .arg("--log-view")
        .arg("developer")
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .context("start interactive pytools menu with capability evidence")?;
    child
        .stdin
        .take()
        .context("interactive stdin missing")?
        .write_all(input.as_bytes())
        .context("send interactive menu selections")?;
    child
        .wait_with_output()
        .context("wait for interactive pytools menu")
}

fn output_json(output: &Output) -> Result<Value> {
    let start = output
        .stdout
        .iter()
        .enumerate()
        .filter(|(_, byte)| matches!(byte, b'{' | b'['))
        .map(|(index, _)| index)
        .min()
        .context("CLI JSON output missing")?;
    let mut deserializer = serde_json::Deserializer::from_slice(&output.stdout[start..]);
    Value::deserialize(&mut deserializer).context("decode CLI JSON output")
}

/// 菜单不再打印 JSON：追加计划通过 CLI 子命令获取（仍输出数组 JSON）。
fn append_plan_json(config: &Path, output_dir: &Path) -> Result<Value> {
    let run = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-append-plan", "--config"])
        .arg(config)
        .args(["--output-dir"])
        .arg(output_dir)
        .output()?;
    anyhow::ensure!(
        run.status.success(),
        "dedup-append-plan failed: {}",
        String::from_utf8_lossy(&run.stderr)
    );
    output_json(&run)
}

/// 菜单不再打印 JSON：从运行记录目录读取指定前缀的最新 JSON（developer 视图产出）。
fn menu_history_json(root: &Path, prefix: &str) -> Result<Value> {
    let log_dir = root.join("logs");
    let mut candidates: Vec<PathBuf> = fs::read_dir(&log_dir)?
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .filter(|path| {
            path.file_name()
                .and_then(|name| name.to_str())
                .is_some_and(|name| name.starts_with(prefix))
        })
        .collect();
    candidates.sort();
    let path = candidates
        .pop()
        .with_context(|| format!("run history starting with {prefix} missing"))?;
    let bytes = fs::read(&path)?;
    serde_json::from_slice(&bytes).with_context(|| format!("decode {}", path.display()))
}
