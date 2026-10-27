use anyhow::{Context, Result};
use excel_editor::{
    BackendCapability, CapabilityEvidence, CapabilityStatus, EditOperation, EngineId,
};
use excel_reader::{read_defined_names, read_workbook};
use pytools_core::batch_ops::{
    build_quarterly_rural_loan_plan, build_remove_foreign_plan, build_rural_loan_plan,
    build_village_split_plan,
};
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
            "pytools-batch-menu-{}-{nonce}-{serial}",
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
fn interactive_batch_menu_runs_region_sum_check_through_real_cli() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let fixture_dir = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/region_sum");
    let fixture_names = [
        "balanced.xlsx",
        "mismatch.xlsx",
        "no_target.xlsx",
        "non_numeric_missing.xlsx",
        "broken.xlsx",
    ];
    let inputs = fixture_names
        .iter()
        .map(|name| fixture_dir.join(name))
        .collect::<Vec<_>>();
    let original_inputs = inputs
        .iter()
        .map(|path| Ok((path.clone(), fs::read(path)?)))
        .collect::<Result<Vec<_>>>()?;
    let unsupported_xls = root.path().join("unsupported.xls");
    fs::write(&unsupported_xls, b"synthetic legacy workbook placeholder")?;
    let original_unsupported_xls = fs::read(&unsupported_xls)?;
    let input_list = inputs
        .iter()
        .map(|path| path.to_string_lossy())
        .collect::<Vec<_>>()
        .join(";")
        + ";"
        + &unsupported_xls.to_string_lossy();

    let output = run_menu(&config, &format!("5\n5\n1\n{input_list}\n0\n"))?;
    assert!(
        output.status.success(),
        "batch menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("5  专用报表处理"));
    assert!(stdout.contains("5  地区总分校验"));
    assert!(stdout.contains("已选择文件 6 个"));
    assert!(stdout.contains("先显式转换为 .xlsx"));
    assert!(stdout.contains(r#""files": 4"#));
    assert!(stdout.contains(r#""sheets": 3"#));
    assert!(stdout.contains(r#""errors": 1"#));
    assert!(stdout.contains(r#""failed_files": 1"#));

    let output_dir = root.path().join("app/output");
    let report_path = fs::read_dir(&output_dir)?
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .find(|path| {
            path.file_name()
                .and_then(|name| name.to_str())
                .is_some_and(|name| {
                    name.starts_with("地区数值核对结果_") && name.ends_with(".xlsx")
                })
        })
        .context("region sum result workbook was not created")?;
    let report = read_workbook(&report_path)?;
    let sheet = report
        .sheets
        .iter()
        .find(|sheet| sheet.name == "核对结果")
        .context("region sum result sheet missing")?;
    assert_eq!(sheet.values["B2"]["value"], "balanced.xlsx");
    assert_eq!(sheet.values["D3"]["value"], "错误");
    assert_eq!(sheet.values["B5"]["value"], "已检查文件数: 4");

    let log_dir = root.path().join("app/logs");
    assert!(fs::read_dir(log_dir)?.next().is_some());
    for (path, original) in original_inputs {
        assert_eq!(
            fs::read(&path)?,
            original,
            "input changed: {}",
            path.display()
        );
    }
    assert_eq!(fs::read(&unsupported_xls)?, original_unsupported_xls);
    Ok(())
}

#[test]
fn interactive_region_sum_menu_records_user_developer_and_off_views() -> Result<()> {
    let fixture_dir = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/region_sum");
    let balanced = fixture_dir.join("balanced.xlsx");
    let broken = fixture_dir.join("broken.xlsx");
    let input = format!(
        "5\n5\n1\n{};{}\n0\n0\n",
        balanced.display(),
        broken.display()
    );
    let mut expected_projection = None::<Value>;

    for view in ["user", "developer", "off"] {
        let root = TestDirectory::new()?;
        let config = create_config(root.path())?;
        let output =
            run_menu_with_arguments(&config, &input, &["--log-view".to_owned(), view.to_owned()])?;
        assert!(
            output.status.success(),
            "PT-7-5 {view} menu failed: {}",
            String::from_utf8_lossy(&output.stderr)
        );
        let stdout = String::from_utf8_lossy(&output.stdout);
        let stderr = String::from_utf8_lossy(&output.stderr);
        assert!(stdout.contains(r#""files": 1"#));
        assert!(stdout.contains(r#""failed_files": 1"#));

        let output_dir = root.path().join("app/output");
        let report_path = fs::read_dir(&output_dir)?
            .filter_map(|entry| entry.ok().map(|entry| entry.path()))
            .find(|path| {
                path.file_name()
                    .and_then(|name| name.to_str())
                    .is_some_and(|name| {
                        name.starts_with("地区数值核对结果_") && name.ends_with(".xlsx")
                    })
            })
            .context("PT-7-5 result workbook was not created")?;
        let report = read_workbook(&report_path)?;
        let sheet = report
            .sheets
            .iter()
            .find(|sheet| sheet.name == "核对结果")
            .context("PT-7-5 result worksheet is missing")?;
        assert_eq!(sheet.values["B2"]["value"], "balanced.xlsx");
        let projection = serde_json::to_value(&sheet.values)?;
        if let Some(expected) = &expected_projection {
            assert_eq!(&projection, expected, "output changed in {view} view");
        } else {
            expected_projection = Some(projection);
        }

        let log_dir = root.path().join("app/logs");
        let history_paths = fs::read_dir(&log_dir)?
            .filter_map(|entry| entry.ok().map(|entry| entry.path()))
            .filter(|path| {
                path.file_name()
                    .and_then(|name| name.to_str())
                    .is_some_and(|name| name.starts_with("PT-7-5_") && name.ends_with(".run.json"))
            })
            .collect::<Vec<_>>();
        if view == "off" {
            assert!(history_paths.is_empty(), "off view must not save history");
            assert!(!stderr.contains("workbook_read"));
            assert!(!stderr.contains("calamine"));
            continue;
        }

        assert_eq!(history_paths.len(), 1, "{view} history count");
        let history_bytes = fs::read(&history_paths[0])?;
        let history_text = String::from_utf8_lossy(&history_bytes);
        let history: Value = serde_json::from_slice(&history_bytes)?;
        assert_eq!(history["product"], "pytools");
        assert_eq!(history["feature_id"], "PT-7-5");
        assert_eq!(history["function"], "region-sum-check");
        assert_eq!(history["total_files"], 2);
        assert_eq!(history["completed_files"], 1);
        assert_eq!(history["failed_files"], 1);
        assert!(history["actual_engines"]
            .as_array()
            .context("actual_engines")?
            .iter()
            .any(|engine| engine == "calamine"));
        for engine in ["rust_xlsxwriter", "ooxml-patch"] {
            assert!(history["actual_engines"]
                .as_array()
                .context("actual_engines")?
                .iter()
                .any(|actual| actual == engine));
        }
        let events = history["events"].as_array().context("run events")?;
        assert_eq!(events.first().context("start event")?["sequence"], 1);
        assert_eq!(
            events.last().context("completion event")?["status"],
            "completed"
        );
        assert!(events.iter().any(|event| {
            event["phase"] == "report_write" && event["stage_elapsed_ms"].is_number()
        }));
        assert!(history_text.contains("workbook_read_failed"));
        assert!(!history_text.contains("balanced.xlsx"));
        assert!(!history_text.contains(&balanced.to_string_lossy().to_string()));

        if view == "user" {
            assert!(stderr.contains("开始地区总分校验"));
            assert!(stderr.contains("第 2 个工作簿读取失败"));
            assert!(!stderr.contains("calamine"));
            assert!(!stderr.contains("rust_xlsxwriter"));
        } else {
            assert!(stderr.contains(r#""actual_engine":"calamine""#));
            assert!(stderr.contains(r#""actual_engine":"rust_xlsxwriter""#));
            assert!(stderr.contains(r#""error_type":"workbook_read_failed""#));
        }
    }
    Ok(())
}

#[test]
fn interactive_village_split_menu_creates_validated_copy_without_modifying_target() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let current = root.path().join("current.xlsx");
    let previous = root.path().join("previous.xlsx");
    let target = root.path().join("target.xlsx");
    make_village_period(&current, true)?;
    make_village_period(&previous, false)?;
    make_village_target(&target)?;
    let original_target = excel_editor::open_xlsx(&target)?;
    assert!(original_target
        .defined_names()
        .iter()
        .chain(
            original_target
                .sheet_collection()
                .iter()
                .flat_map(|sheet| sheet.defined_names()),
        )
        .any(|name| name.name() == "VillageTail"));
    let current_bytes = fs::read(&current)?;
    let previous_bytes = fs::read(&previous)?;
    let target_bytes = fs::read(&target)?;

    let input = village_menu_input(&current, &previous, &target);
    let output = run_menu(&config, &input)?;
    assert!(
        output.status.success(),
        "batch menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("7-1"));
    assert!(stdout.contains("[新建模式]"));
    assert!(stdout.contains(r#""feature_id": "PT-7-1""#));
    assert!(stdout.contains(r#""created_copies": 1"#));
    assert!(stdout.contains(r#""committed_files": 0"#));
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("已使用 Rust 新建并校验专用报表副本"));
    assert_eq!(fs::read(&current)?, current_bytes);
    assert_eq!(fs::read(&previous)?, previous_bytes);
    assert_eq!(
        fs::read(&target)?,
        target_bytes,
        "new-copy mode must be read-only"
    );

    let output_copy = root
        .path()
        .join("PyTools_专用报表处理结果")
        .join("target.xlsx");
    let output_book = read_workbook(&output_copy)?;
    let branch = output_book
        .sheets
        .iter()
        .find(|sheet| sheet.name == "广东分机构本外币")
        .context("village split result sheet missing")?;
    assert_eq!(branch.values["A3"]["value"], "博罗长江村镇银行");
    assert_eq!(branch.values["A4"]["value"], "惠东惠民村镇银行");
    assert_eq!(branch.values["A5"]["value"], "尾行");
    assert_eq!(branch.formulas["J5"], "A5");
    assert!(branch.merges.contains(&"K5:L5".to_owned()));
    assert!(branch.comments.contains_key("M5"));
    let untouched = output_book
        .sheets
        .iter()
        .find(|sheet| sheet.name == "非目标工作表")
        .context("untouched worksheet missing from output")?;
    assert_eq!(untouched.formulas["B2"], "'广东分机构本外币'!A5");
    assert!(untouched.merges.contains(&"D4:E4".to_owned()));
    assert!(untouched.comments.contains_key("F2"));
    assert_eq!(
        read_defined_names(&output_copy)?
            .iter()
            .find(|name| name.name == "VillageTail")
            .map(|name| name.raw_reference.as_str()),
        Some("'广东分机构本外币'!$J$5")
    );
    assert!(find_file_recursively(&root.path().join("app/output"), ".plan.json")?.is_none());

    let history_path = find_file_recursively(&root.path().join("app/logs"), ".run.json")?
        .context("village split run history was not written")?;
    let history_bytes = fs::read(&history_path)?;
    let history: Value = serde_json::from_slice(&history_bytes)?;
    assert_eq!(history["product"], "pytools");
    assert_eq!(history["feature_id"], "PT-7-1");
    assert_eq!(history["status"], "succeeded");
    assert!(history["events"].as_array().is_some_and(|events| {
        events
            .iter()
            .any(|event| event["phase"] == "planning" && event["actual_engine"] == "calamine")
            && events.iter().any(|event| {
                event["phase"] == "engine_execution"
                    && event["status"] == "completed"
                    && event["actual_engine"] == "rust_xlsxwriter"
            })
    }));
    assert!(!String::from_utf8_lossy(&history_bytes).contains("target.xlsx"));
    Ok(())
}

#[test]
fn village_split_numeric_text_matches_legacy_python_float_conversion() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let current = root.path().join("current.xlsx");
    let previous = root.path().join("previous.xlsx");
    let target = root.path().join("target.xlsx");

    let current_bytes = {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("村镇银行本外币")?;
        sheet.write_string(4, 2, "1,000")?;
        sheet.write_number(4, 86, 200.0)?;
        workbook.save(&current)?;
        fs::read(&current)?
    };
    let previous_bytes = {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("村镇银行本外币")?;
        sheet.write_number(4, 2, 500.0)?;
        sheet.write_number(4, 86, 100.0)?;
        workbook.save(&previous)?;
        fs::read(&previous)?
    };
    make_village_target(&target)?;
    let target_bytes = fs::read(&target)?;

    let output = run_menu(&config, &village_menu_input(&current, &previous, &target))?;
    assert!(
        output.status.success(),
        "village split menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert!(
        String::from_utf8_lossy(&output.stdout).contains(r#""created_copies": 1"#),
        "village output: {}\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    let output_copy = root
        .path()
        .join("PyTools_专用报表处理结果")
        .join("target.xlsx");
    let workbook = read_workbook(&output_copy)?;
    let branch = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "广东分机构本外币")
        .context("village split result sheet missing")?;
    assert_eq!(branch.values["B3"]["value"].as_f64(), Some(0.0));
    assert_eq!(branch.values["E3"]["value"].as_f64(), Some(-100.0));
    assert!(find_file_recursively(&root.path().join("app/output"), ".plan.json")?.is_none());

    assert_eq!(fs::read(&current)?, current_bytes);
    assert_eq!(fs::read(&previous)?, previous_bytes);
    assert_eq!(fs::read(&target)?, target_bytes);
    Ok(())
}

#[test]
fn interactive_village_split_menu_off_view_skips_run_history() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let current = root.path().join("current.xlsx");
    let previous = root.path().join("previous.xlsx");
    let target = root.path().join("target.xlsx");
    make_village_period(&current, true)?;
    make_village_period(&previous, false)?;
    make_village_target(&target)?;
    let target_bytes = fs::read(&target)?;

    let output = run_menu_with_arguments(
        &config,
        &village_menu_input(&current, &previous, &target),
        &["--log-view".to_owned(), "off".to_owned()],
    )?;
    assert!(
        output.status.success(),
        "batch menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert!(!String::from_utf8_lossy(&output.stderr).contains("运行记录已保存"));
    assert_eq!(fs::read(&target)?, target_bytes);
    assert!(root
        .path()
        .join("PyTools_专用报表处理结果/target.xlsx")
        .exists());
    let log_dir = root.path().join("app/logs");
    assert!(!log_dir.exists() || fs::read_dir(log_dir)?.next().is_none());
    Ok(())
}

#[test]
fn interactive_village_split_menu_executes_capability_and_safely_commits_target() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let current = root.path().join("current.xlsx");
    let previous = root.path().join("previous.xlsx");
    let target = root.path().join("target.xlsx");
    let capability_path = root.path().join("test-capability.json");
    make_village_period(&current, true)?;
    make_village_period(&previous, false)?;
    make_village_target(&target)?;
    let original_target = excel_editor::open_xlsx(&target)?;
    let original_branch_row_height = original_target
        .sheet_by_name("广东分机构本外币")?
        .row_dimension(4)
        .map(|row| row.height());
    let original_untouched_row_height = original_target
        .sheet_by_name("非目标工作表")?
        .row_dimension(5)
        .map(|row| row.height());
    let current_bytes = fs::read(&current)?;
    let previous_bytes = fs::read(&previous)?;
    let target_bytes = fs::read(&target)?;
    let plan = build_village_split_plan(
        &current,
        &previous,
        &target,
        &root.path().join("capability-candidate.xlsx"),
    )?
    .plan
    .context("village split plan missing")?;
    let capability = village_split_capability(&plan);
    fs::write(&capability_path, serde_json::to_vec_pretty(&capability)?)?;

    let input = village_menu_input(&current, &previous, &target);
    let output = run_menu_with_capability(
        &config,
        &input,
        &capability_path,
        "synthetic-pt-7-1-umya-menu-test",
        "developer",
    )?;
    assert!(
        output.status.success(),
        "batch menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stdout.contains("[执行模式]"));
    assert!(stdout.contains(r#""committed_files": 1"#));
    assert!(stdout.contains("umya_spreadsheet"));
    assert!(stderr.contains("\"actual_engine\":\"umya_spreadsheet\""));
    assert!(stderr.contains("\"phase\":\"safe_commit\""));
    assert_eq!(fs::read(&current)?, current_bytes);
    assert_eq!(fs::read(&previous)?, previous_bytes);
    assert_ne!(
        fs::read(&target)?,
        target_bytes,
        "target should be committed"
    );

    let workbook = read_workbook(&target)?;
    let branch = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "广东分机构本外币")
        .context("branch sheet missing after village split")?;
    assert_eq!(branch.values["A3"]["value"], "博罗长江村镇银行");
    assert_eq!(branch.values["A4"]["value"], "惠东惠民村镇银行");
    assert_eq!(branch.values["B3"]["value"], 100.0);
    assert_eq!(branch.values["B4"]["value"], 110.0);
    let untouched = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "非目标工作表")
        .context("untouched worksheet missing")?;
    assert_eq!(untouched.values["A2"]["value"], "保持不变");

    let preserved = excel_editor::open_xlsx(&target)?;
    let branch = preserved.sheet_by_name("广东分机构本外币")?;
    assert_eq!(branch.cell("J5").map(|cell| cell.formula()), Some("A5"));
    assert!(branch
        .merge_cells()
        .iter()
        .any(|range| range.range() == "K5:L5"));
    assert_eq!(
        branch.row_dimension(5).map(|row| row.height()),
        original_branch_row_height
    );
    assert!(branch
        .comments()
        .iter()
        .any(|comment| comment.coordinate().to_string() == "M5"));
    let defined_name = preserved
        .defined_names()
        .iter()
        .chain(
            preserved
                .sheet_collection()
                .iter()
                .flat_map(|sheet| sheet.defined_names()),
        )
        .find(|name| name.name() == "VillageTail")
        .context("village tail defined name missing")?;
    assert!(defined_name.address().ends_with("$J$5"));

    let untouched = preserved.sheet_by_name("非目标工作表")?;
    assert_eq!(
        untouched.cell("B2").map(|cell| cell.formula()),
        Some("'广东分机构本外币'!A5"),
        "cross-sheet formula references must follow the target row"
    );
    assert!(untouched
        .merge_cells()
        .iter()
        .any(|range| range.range() == "D4:E4"));
    assert_eq!(
        untouched.row_dimension(5).map(|row| row.height()),
        original_untouched_row_height
    );
    assert!(untouched
        .comments()
        .iter()
        .any(|comment| comment.coordinate().to_string() == "F2"));
    Ok(())
}

#[test]
fn interactive_foreign_bank_menu_creates_result_copies_and_user_history() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_foreign_bank_config(root.path())?;
    let source = root.path().join("foreign.xlsx");
    let no_change_source = root.path().join("no-change.xlsx");
    let unsupported = root.path().join("legacy.xls");
    make_foreign_bank_source(&source, true)?;
    make_foreign_bank_source(&no_change_source, false)?;
    fs::write(&unsupported, b"synthetic legacy workbook placeholder")?;
    let source_bytes = fs::read(&source)?;
    let no_change_bytes = fs::read(&no_change_source)?;
    let unsupported_bytes = fs::read(&unsupported)?;

    let input = foreign_bank_menu_input(&[&source, &no_change_source, &unsupported]);
    let output = run_menu(&config, &input)?;
    assert!(
        output.status.success(),
        "foreign-bank menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("------ 专用报表处理 ------"));
    assert!(stdout.contains("删除分机构表中的外资行"));
    assert!(stdout.contains("PT-7-3 当前只接受 .xlsx"));
    assert!(stdout.contains(r#""mode": "execute_copy""#));
    assert!(
        stdout.contains(r#""completed_files": 2"#),
        "foreign-bank output: {stdout}"
    );
    assert!(stdout.contains(r#""planned_branch_sheets": 4"#));
    assert!(stdout.contains(r#""planned_deleted_rows": 3"#));
    assert_eq!(
        fs::read(&source)?,
        source_bytes,
        "copy mode must preserve its source"
    );
    assert_eq!(
        fs::read(&no_change_source)?,
        no_change_bytes,
        "copy mode must preserve its source"
    );
    assert_eq!(fs::read(&unsupported)?, unsupported_bytes);
    let result_dir = root.path().join("PyTools_专用报表处理结果");
    let foreign_copy = result_dir.join("foreign(金融局).xlsx");
    let no_change_copy = result_dir.join("no-change(金融局).xlsx");
    assert!(foreign_copy.exists());
    assert!(no_change_copy.exists());
    assert!(find_file_recursively(&root.path().join("app/output"), ".plan.json")?.is_none());
    let copied = read_workbook(&foreign_copy)?;
    let branch = copied
        .sheets
        .iter()
        .find(|sheet| sheet.name == "甲分机构本外币")
        .context("foreign-bank copied branch sheet missing")?;
    assert_eq!(branch.values["A2"]["value"], "普通甲");
    assert_eq!(branch.values["A3"]["value"], "普通乙");
    assert_eq!(branch.values["A4"]["value"], "普通丙");
    assert_eq!(branch.values["A5"]["value"], "尾行");
    assert_eq!(branch.formulas["H1"], "SUM(B2:B5)");
    assert!(branch.comments.contains_key("C3"));
    assert!(branch.comments.contains_key("C5"));
    assert!(branch.merges.contains(&"E7:F8".to_owned()));
    assert_eq!(
        read_defined_names(&foreign_copy)?
            .iter()
            .find(|name| name.name == "BranchValues")
            .map(|name| name.raw_reference.as_str()),
        Some("'甲分机构本外币'!$B$2:$B$5")
    );

    let history_path = find_file_recursively(&root.path().join("app/logs"), ".json")?
        .context("PT-7-3 user run history was not saved")?;
    let history: Value = serde_json::from_slice(&fs::read(history_path)?)?;
    assert_eq!(history["feature_id"], "PT-7-3");
    assert_eq!(history["mode"], "create_copy");
    assert_eq!(history["summary"]["completed_files"], 2);
    assert_eq!(history["summary"]["planned_deleted_rows"], 3);
    assert_eq!(history["summary"]["deleted_rows"], 3);
    assert!(history["events"]
        .as_array()
        .is_some_and(|events| !events.is_empty()));
    let history_text = serde_json::to_string(&history)?;
    assert!(!history_text.contains("foreign.xlsx"));
    assert!(!history_text.contains("外资甲"));
    assert!(String::from_utf8_lossy(&output.stderr).contains("已完成 2"));
    Ok(())
}

#[test]
fn interactive_foreign_bank_menu_off_view_has_no_events_or_history() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_foreign_bank_config(root.path())?;
    let source = root.path().join("foreign.xlsx");
    make_foreign_bank_source(&source, true)?;
    let source_bytes = fs::read(&source)?;

    let output = run_menu_with_arguments(
        &config,
        &foreign_bank_menu_input(&[&source]),
        &["--log-view".to_owned(), "off".to_owned()],
    )?;
    assert!(
        output.status.success(),
        "foreign-bank menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains(r#""planned_deleted_rows": 3"#));
    assert!(!stdout.contains("history_report"));
    assert!(!stdout.contains("\"events\""));
    assert!(String::from_utf8_lossy(&output.stderr).is_empty());
    let log_dir = root.path().join("app/logs");
    assert!(!log_dir.exists() || fs::read_dir(log_dir)?.next().is_none());
    assert_eq!(fs::read(&source)?, source_bytes);
    assert!(root
        .path()
        .join("PyTools_专用报表处理结果/foreign(金融局).xlsx")
        .exists());
    Ok(())
}

#[test]
fn interactive_foreign_bank_menu_stops_before_file_selection_when_mapping_is_empty() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let output = run_menu(&config, "5\n3\n0\n0\n")?;
    assert!(
        output.status.success(),
        "foreign-bank menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("外资行配置为空"));
    assert!(!stdout.contains("选择要生成外发用副本的工作簿"));
    assert!(!root.path().join("app/output").exists());
    assert!(!root.path().join("app/logs").exists());
    Ok(())
}

#[test]
fn interactive_foreign_bank_menu_executes_only_with_synthetic_capability() -> Result<()> {
    const PLATFORM: &str = "synthetic-pt-7-3-menu-test";
    let root = TestDirectory::new()?;
    let config = create_foreign_bank_config(root.path())?;
    let source = root.path().join("foreign.xlsx");
    let capability_path = root.path().join("test-capability.json");
    make_foreign_bank_source(&source, true)?;
    let source_bytes = fs::read(&source)?;
    let expected_plan = build_remove_foreign_plan(
        &source,
        &root.path().join("foreign(金融局).xlsx"),
        &pytools_core::config::load_foreign_banks(&config)?,
    )?
    .plan
    .context("foreign-bank edit plan missing")?;
    fs::write(
        &capability_path,
        serde_json::to_vec_pretty(&foreign_bank_capability(&expected_plan, PLATFORM))?,
    )?;

    let output = run_menu_with_capability(
        &config,
        &foreign_bank_menu_input(&[&source]),
        &capability_path,
        PLATFORM,
        "developer",
    )?;
    assert!(
        output.status.success(),
        "foreign-bank menu apply failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains(r#""mode": "execute""#));
    assert!(stdout.contains(r#""ok_files": 1"#));
    assert!(stdout.contains(r#""failed_files": 0"#));
    assert!(stdout.contains(r#""engine": "umya_spreadsheet""#));

    let output_path = root.path().join("foreign(金融局).xlsx");
    assert!(output_path.exists());
    assert_eq!(fs::read(&source)?, source_bytes, "source remains unchanged");
    let workbook = read_workbook(&output_path)?;
    let branch = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "甲分机构本外币")
        .context("foreign-bank output branch sheet missing")?;
    assert_eq!(branch.values["A2"]["value"], "普通甲");
    assert_eq!(branch.values["A3"]["value"], "普通乙");
    assert_eq!(branch.values["A4"]["value"], "普通丙");
    assert_eq!(branch.values["A5"]["value"], "尾行");
    assert_eq!(branch.formulas["H1"], "SUM(B2:B5)");

    let history_path = find_file_recursively(&root.path().join("app/logs"), ".json")?
        .context("PT-7-3 developer run history was not saved")?;
    let history: Value = serde_json::from_slice(&fs::read(history_path)?)?;
    assert!(history["events"].as_array().is_some_and(|events| {
        events.iter().any(|event| {
            event["phase"] == "engine_execution" && event["actual_engine"] == "umya_spreadsheet"
        })
    }));
    Ok(())
}

fn create_foreign_bank_config(root: &Path) -> Result<PathBuf> {
    let app = root.join("app");
    let config = app.join("config/config.xlsx");
    fs::create_dir_all(config.parent().context("config parent missing")?)?;
    fs::create_dir_all(app.join("input"))?;
    let mut workbook = Workbook::new();
    let global = workbook.add_worksheet();
    global.set_name("全局配置")?;
    global.write_string(0, 0, "配置项")?;
    global.write_string(0, 1, "值")?;
    for (index, (name, value)) in [
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
        global.write_string((index + 1) as u32, 0, *name)?;
        global.write_string((index + 1) as u32, 1, value)?;
    }
    let mapping = workbook.add_worksheet();
    mapping.set_name("机构映射表")?;
    for (column, header) in ["序号", "机构名称", "是否为外资行"].iter().enumerate() {
        mapping.write_string(0, column as u16, *header)?;
    }
    for (row, (name, is_foreign)) in [("外资甲", 1.0), ("外资乙", 1.0), ("普通甲", 0.0)]
        .iter()
        .enumerate()
    {
        mapping.write_number((row + 1) as u32, 0, (row + 1) as f64)?;
        mapping.write_string((row + 1) as u32, 1, *name)?;
        mapping.write_number((row + 1) as u32, 2, *is_foreign)?;
    }
    workbook.save(&config)?;
    Ok(config)
}

fn make_foreign_bank_source(path: &Path, include_foreign: bool) -> Result<()> {
    let mut workbook = Workbook::new();
    let branch = workbook.add_worksheet();
    branch.set_name("甲分机构本外币")?;
    branch.write_string(0, 0, "机构")?;
    branch.write_string(0, 1, "余额")?;
    for (row, name, amount) in [
        (1, "普通甲", 10.0),
        (
            2,
            if include_foreign {
                "外资甲"
            } else {
                "普通己"
            },
            20.0,
        ),
        (3, "普通乙", 30.0),
        (
            4,
            if include_foreign {
                "外资乙"
            } else {
                "普通庚"
            },
            40.0,
        ),
        (5, "普通丙", 50.0),
        (6, "尾行", 60.0),
    ] {
        branch.write_string(row, 0, name)?;
        branch.write_number(row, 1, amount)?;
    }
    branch.write_formula(0, 7, "=SUM(B2:B7)")?;
    branch.insert_note(2, 2, &rust_xlsxwriter::Note::new("删除行批注"))?;
    branch.insert_note(3, 2, &rust_xlsxwriter::Note::new("保留行批注一"))?;
    branch.insert_note(6, 2, &rust_xlsxwriter::Note::new("保留行批注二"))?;
    branch.set_row_height(6, 24.0)?;
    branch.merge_range(8, 4, 9, 5, "合并区域", &rust_xlsxwriter::Format::new())?;
    let second = workbook.add_worksheet();
    second.set_name("乙分机构人民币")?;
    second.write_string(0, 0, "机构")?;
    second.write_string(
        1,
        0,
        if include_foreign {
            "外资乙"
        } else {
            "普通戊"
        },
    )?;
    second.write_string(2, 0, "普通丁")?;
    let non_target = workbook.add_worksheet();
    non_target.set_name("非目标工作表")?;
    non_target.write_string(0, 0, "机构")?;
    non_target.write_string(1, 0, "外资甲")?;
    let formula = workbook.add_worksheet();
    formula.set_name("公式引用")?;
    if include_foreign {
        formula.write_formula(0, 0, "='甲分机构本外币'!B7")?;
        workbook.define_name("BranchValues", "'甲分机构本外币'!$B$2:$B$7")?;
    } else {
        formula.write_string(0, 0, "无外部公式")?;
    }
    workbook.save(path)?;
    Ok(())
}

fn foreign_bank_menu_input(sources: &[&Path]) -> String {
    let paths = sources
        .iter()
        .map(|path| path.to_string_lossy())
        .collect::<Vec<_>>()
        .join(";");
    format!("5\n3\n1\n{paths}\n0\n0\n")
}

fn foreign_bank_capability(
    plan: &excel_editor::ExcelEditPlan,
    platform: &str,
) -> BackendCapability {
    BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Passed,
        format: plan.format,
        operations: plan.operations.iter().map(EditOperation::kind).collect(),
        supported_features: plan.workbook.features_present.clone(),
        preserved: plan.preserve.clone(),
        platform: platform.to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: "synthetic-pt-7-3-foreign-bank-menu".to_owned(),
            report_path: "pytools-core/tests/batch_tools_menu_cli.rs".to_owned(),
            adapter_version: "umya-spreadsheet-3.1.1".to_owned(),
            engine_version: "umya-spreadsheet-3.1.1".to_owned(),
            platform: platform.to_owned(),
        }),
    }
}

#[test]
fn interactive_institution_normalize_menu_creates_copies_in_user_developer_and_off_views(
) -> Result<()> {
    let root = TestDirectory::new()?;
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/institution_normalize.xlsx");
    let unsupported = root.path().join("unsupported.xlsm");
    fs::write(&unsupported, b"synthetic unverified XLSM placeholder")?;

    for view in ["user", "developer", "off"] {
        let view_root = root.path().join(view);
        let config = create_config(&view_root)?;
        let target = view_root.join("institution-target.xlsx");
        fs::copy(&fixture, &target)?;
        let target_before = fs::read(&target)?;
        let input_paths = if view == "user" {
            format!("{};{}", target.display(), unsupported.display())
        } else {
            target.display().to_string()
        };
        let input = format!("5\n2\n1\n{input_paths}\n0\n");
        let output =
            run_menu_with_arguments(&config, &input, &["--log-view".to_owned(), view.to_owned()])?;
        assert!(
            output.status.success(),
            "PT-7-2 menu ({view}) failed: {}",
            String::from_utf8_lossy(&output.stderr)
        );
        assert_eq!(fs::read(&target)?, target_before, "{view} modified input");
        let result_dir = view_root
            .join(pytools_core::institution_normalize::INSTITUTION_NORMALIZE_OUTPUT_DIRECTORY);
        let normalized_copy = result_dir.join("institution-target.xlsx");
        assert!(
            normalized_copy.is_file(),
            "{view} did not create output copy at {}\nstdout:\n{}\nstderr:\n{}",
            normalized_copy.display(),
            String::from_utf8_lossy(&output.stdout),
            String::from_utf8_lossy(&output.stderr)
        );

        let stdout = String::from_utf8_lossy(&output.stdout);
        assert!(stdout.contains("2  机构名称规范化"));
        assert!(stdout.contains(r#""feature_id": "PT-7-2""#));
        assert!(stdout.contains(r#""mapped_rows": 3"#));
        let stderr = String::from_utf8_lossy(&output.stderr);
        if view == "off" {
            assert!(stderr.trim().is_empty(), "off view emitted logs: {stderr}");
            assert!(!stdout.contains("history_report"));
            assert!(!stdout.contains("\"events\""));
        } else {
            assert!(find_file_recursively(&view_root.join("app/logs"), ".json")?.is_some());
            assert!(!view_root.join("app/output/plans/PT-7-2").exists());
            if view == "user" {
                assert!(stderr.contains("正在处理 1/1"));
                assert!(stdout.contains("[跳过]"));
                assert!(!stdout.contains("calamine"));
            } else {
                assert!(stderr.contains("\"phase\":\"engine_execution\""));
                assert!(stdout.contains("rust_xlsxwriter"));
            }
        }
        if view == "off" {
            let log_dir = view_root.join("app/logs");
            assert!(!log_dir.exists() || fs::read_dir(log_dir)?.next().is_none());
            assert!(!view_root.join("app/output/plans/PT-7-2").exists());
        }
    }
    Ok(())
}

#[test]
fn interactive_institution_normalize_menu_recursively_creates_copies_and_skips_result_folder(
) -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let source_root = root.path().join("sources");
    let first_folder = source_root.join("A");
    let second_folder = source_root.join("nested/B");
    fs::create_dir_all(&first_folder)?;
    fs::create_dir_all(&second_folder)?;
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/institution_normalize.xlsx");
    let first = first_folder.join("report.xlsx");
    let second = second_folder.join("report.xlsx");
    fs::copy(&fixture, &first)?;
    fs::copy(&fixture, &second)?;
    let previous_copy = first_folder
        .join(pytools_core::institution_normalize::INSTITUTION_NORMALIZE_OUTPUT_DIRECTORY)
        .join("report.xlsx");
    fs::create_dir_all(previous_copy.parent().context("result directory missing")?)?;
    fs::copy(&fixture, &previous_copy)?;
    let originals = [(&first, fs::read(&first)?), (&second, fs::read(&second)?)];
    fs::write(first_folder.join("unsupported.xlsm"), b"unverified format")?;
    fs::write(second_folder.join("~$locked.xlsx"), b"Excel lock file")?;

    let input = format!("5\n2\n3\n{}\n0\n", source_root.display());
    let output = run_menu(&config, &input)?;
    assert!(
        output.status.success(),
        "recursive PT-7-2 menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("选择文件夹（递归子目录）"));
    assert!(stdout.contains("已选择文件 3 个"));
    assert!(stdout.contains("这是机构名称规范化结果目录中的副本，不重复处理"));
    assert!(stdout.contains(r#""total_inputs": 2"#));
    assert!(stdout.contains(r#""completed_files": 2"#));
    let python_baseline: Value = serde_json::from_slice(&fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("probe/reports/institution_normalize_baseline.json"),
    )?)?;
    let python_rows_per_file = python_baseline["stats"]["mapped_rows"]
        .as_u64()
        .context("Python institution-normalize baseline mapped_rows missing")?;
    assert!(stdout.contains(&format!(r#""mapped_rows": {}"#, python_rows_per_file * 2)));
    assert!(!stdout.contains("unsupported.xlsm"));
    assert!(!stdout.contains("~$locked.xlsx"));

    for (path, original) in originals {
        assert_eq!(
            fs::read(path)?,
            original,
            "plan mode modified {}",
            path.display()
        );
    }
    let first_copy = first_folder
        .join(pytools_core::institution_normalize::INSTITUTION_NORMALIZE_OUTPUT_DIRECTORY)
        .join("report_1.xlsx");
    let second_copy = second_folder
        .join(pytools_core::institution_normalize::INSTITUTION_NORMALIZE_OUTPUT_DIRECTORY)
        .join("report.xlsx");
    assert!(
        first_copy.is_file(),
        "first normalized copy was not created"
    );
    assert!(
        second_copy.is_file(),
        "second normalized copy was not created"
    );
    assert!(!root.path().join("app/output/plans/PT-7-2").exists());
    Ok(())
}

#[test]
fn interactive_institution_normalize_menu_applies_only_with_temporary_capability() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/institution_normalize.xlsx");
    let target = root.path().join("institution-target.xlsx");
    let candidate = root.path().join("plan-candidate.xlsx");
    let capability_path = root.path().join("temporary-test-capability.json");
    fs::copy(&fixture, &target)?;
    let original = fs::read(&target)?;
    let result_dir = root
        .path()
        .join(pytools_core::institution_normalize::INSTITUTION_NORMALIZE_OUTPUT_DIRECTORY);
    fs::create_dir_all(&result_dir)?;
    let existing_copy = result_dir.join("institution-target.xlsx");
    fs::copy(&fixture, &existing_copy)?;
    let existing_copy_before = fs::read(&existing_copy)?;
    let mapping = pytools_core::config::load_institution_mapping_ordered(&config)?;
    let profile = excel_editor::inspect_xlsx_profile(&target)?;
    let plan = pytools_core::institution_normalize::build_institution_normalize_plan(
        &target,
        &candidate,
        profile,
        "synthetic-pt-7-2-menu-capability",
        &mapping,
    )?
    .plan
    .context("institution normalization plan missing")?;
    fs::write(
        &capability_path,
        serde_json::to_vec_pretty(&institution_normalize_capability(&plan))?,
    )?;

    let input = format!("5\n2\n1\n{}\n0\n", target.display());
    let output = run_menu_with_capability(
        &config,
        &input,
        &capability_path,
        "synthetic-pt-7-2-menu-test",
        "developer",
    )?;
    assert!(
        output.status.success(),
        "PT-7-2 menu apply failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stdout.contains("[执行模式]"));
    assert!(stdout.contains(r#""modified_files": 1"#));
    assert!(stderr.contains("\"actual_engine\":\"umya_spreadsheet\""));
    assert_eq!(fs::read(&target)?, original, "source workbook was modified");
    assert_eq!(
        fs::read(&existing_copy)?,
        existing_copy_before,
        "existing result copy was overwritten"
    );
    let normalized_copy = result_dir.join("institution-target_1.xlsx");
    assert!(
        normalized_copy.is_file(),
        "non-overwriting copy was not created"
    );
    assert!(stdout.contains("institution-target_1.xlsx"));
    let workbook = read_workbook(&normalized_copy)?;
    let branch = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "报表分机构本外币")
        .context("institution branch sheet missing after menu apply")?;
    assert_eq!(branch.values["A2"]["value"], "Exact Standard");
    let baseline: Value = serde_json::from_slice(&fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("probe/reports/institution_normalize_baseline.json"),
    )?)?;
    let expected_sheets = baseline["sheets"]
        .as_array()
        .context("institution baseline sheets missing")?;
    assert_eq!(workbook.sheets.len(), expected_sheets.len());
    for expected in expected_sheets {
        let name = expected["name"]
            .as_str()
            .context("baseline sheet name missing")?;
        let actual = workbook
            .sheets
            .iter()
            .find(|sheet| sheet.name == name)
            .with_context(|| format!("output sheet {name:?} missing"))?;
        assert_eq!(serde_json::to_value(&actual.values)?, expected["values"]);
        assert_eq!(
            serde_json::to_value(&actual.formulas)?,
            expected["formulas"]
        );
        assert_eq!(serde_json::to_value(&actual.merges)?, expected["merges"]);
        assert_eq!(
            serde_json::to_value(&actual.comments)?,
            expected["comments"]
        );
    }
    Ok(())
}

#[test]
fn interactive_rural_loan_menu_creates_city_adjustment_copy_and_writes_user_history() -> Result<()>
{
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let (current, previous) = make_rural_loan_periods(root.path(), false, false)?;
    let current_before = fs::read(&current)?;
    let previous_before = fs::read(&previous)?;

    let output = run_menu(&config, &rural_loan_menu_input(7, &current, &previous))?;
    assert!(
        output.status.success(),
        "PT-7-7 menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stdout.contains("7  涉农贷款比上月修正"));
    assert!(stdout.contains(r#""feature_id": "PT-7-7""#));
    assert!(stdout.contains("rust_xlsxwriter"));
    let output_copy = current
        .parent()
        .context("PT-7-7 source parent missing")?
        .join("PyTools_专用报表处理结果")
        .join(current.file_name().context("PT-7-7 source name missing")?);
    assert!(
        output_copy.is_file(),
        "PT-7-7 copy missing: {}\nstdout:\n{}\nstderr:\n{}",
        output_copy.display(),
        stdout,
        stderr
    );
    let adjusted = read_workbook(&output_copy)?;
    let adjusted_sheet = adjusted
        .sheets
        .iter()
        .find(|sheet| sheet.name == "惠州市涉农贷款分机构")
        .context("PT-7-7 result sheet missing")?;
    assert_eq!(adjusted_sheet.values["D3"]["type"], "number");
    let adjusted_value = adjusted_sheet.values["D3"]["value"]
        .as_f64()
        .context("PT-7-7 adjustment is not numeric")?;
    assert!((adjusted_value - 5.123456789).abs() <= 1e-15);
    assert!(stdout.contains(r#""matched": 1"#));
    assert!(stderr.contains("正在读取本期和上期工作簿"));
    assert!(stderr.contains("分地区总数一致"));
    assert_eq!(fs::read(&current)?, current_before);
    assert_eq!(fs::read(&previous)?, previous_before);

    assert!(!root.path().join("app/output/plans/PT-7-7").exists());

    let history_path = find_file_recursively(&root.path().join("app/logs"), ".run.json")?
        .context("PT-7-7 run history was not written")?;
    let history_bytes = fs::read(history_path)?;
    let history: Value = serde_json::from_slice(&history_bytes)?;
    assert_eq!(history["feature_id"], "PT-7-7");
    assert_eq!(history["status"], "succeeded");
    assert_eq!(history["completed_inputs"], 2);
    assert!(history["actual_engines"].as_array().is_some_and(|engines| {
        engines.iter().any(|engine| engine == "calamine")
            && engines.iter().any(|engine| engine == "rust_xlsxwriter")
    }));
    let history_text = String::from_utf8_lossy(&history_bytes);
    assert!(!history_text.contains(&current.display().to_string()));
    assert!(!history_text.contains(&previous.display().to_string()));
    assert!(!history_text.contains("甲行"));
    assert!(!history_text.contains("5.123456789"));
    Ok(())
}

#[test]
fn interactive_quarterly_rural_loan_menu_records_region_errors_in_developer_view() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let (current, previous) = make_rural_loan_periods(root.path(), true, false)?;
    let current_before = fs::read(&current)?;
    let previous_before = fs::read(&previous)?;
    let args = ["--log-view".to_owned(), "developer".to_owned()];

    let output = run_menu_with_arguments(
        &config,
        &rural_loan_menu_input(8, &current, &previous),
        &args,
    )?;
    assert!(
        output.status.success(),
        "PT-7-8 menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stdout.contains("8  季度涉农贷款修正及分地区核对"));
    assert!(stdout.contains(r#""feature_id": "PT-7-8""#));
    assert!(stdout.contains(r#""failed_regions": 1"#));
    assert!(stdout.contains("上期文件缺少江门市分机构表"));
    assert!(stdout.contains("rust_xlsxwriter"));
    assert!(stderr.contains(r#""actual_engine":"rust_xlsxwriter""#));
    assert!(stderr.contains(r#""phase":"planning""#));
    assert_eq!(fs::read(&current)?, current_before);
    assert_eq!(fs::read(&previous)?, previous_before);

    let output_copy = current
        .parent()
        .context("PT-7-8 source parent missing")?
        .join("PyTools_专用报表处理结果")
        .join(current.file_name().context("PT-7-8 source name missing")?);
    assert!(
        output_copy.is_file(),
        "PT-7-7 off-view copy missing: {}\nstdout:\n{}\nstderr:\n{}",
        output_copy.display(),
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    let output_book = read_workbook(&output_copy)?;
    let adjusted_sheet = output_book
        .sheets
        .iter()
        .find(|sheet| sheet.name == "惠州市涉农贷款分机构")
        .context("PT-7-8 result sheet missing")?;
    let adjusted_value = adjusted_sheet.values["D3"]["value"]
        .as_f64()
        .context("PT-7-8 adjustment is not numeric")?;
    assert!((adjusted_value - 5.123456789).abs() <= 1e-15);

    let history_path = find_file_recursively(&root.path().join("app/logs"), ".run.json")?
        .context("PT-7-8 run history was not written")?;
    let history_bytes = fs::read(history_path)?;
    let history: Value = serde_json::from_slice(&history_bytes)?;
    assert_eq!(history["feature_id"], "PT-7-8");
    assert_eq!(history["summary_metrics"]["failed_regions"], 1);
    assert_eq!(history["summary_metrics"]["consistent"], false);
    let history_text = String::from_utf8_lossy(&history_bytes);
    assert!(!history_text.contains(&current.display().to_string()));
    assert!(!history_text.contains(&previous.display().to_string()));
    assert!(!history_text.contains("江门市"));
    Ok(())
}

#[test]
fn interactive_rural_loan_menu_off_view_does_not_write_run_history() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let (current, previous) = make_rural_loan_periods(root.path(), false, false)?;
    let current_before = fs::read(&current)?;
    let previous_before = fs::read(&previous)?;
    let args = ["--log-view".to_owned(), "off".to_owned()];

    let output = run_menu_with_arguments(
        &config,
        &rural_loan_menu_input(7, &current, &previous),
        &args,
    )?;
    assert!(
        output.status.success(),
        "PT-7-7 off-view menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("rust_xlsxwriter"));
    let output_copy = current
        .parent()
        .context("PT-7-7 off-view source parent missing")?
        .join("PyTools_专用报表处理结果")
        .join(
            current
                .file_name()
                .context("PT-7-7 off-view source name missing")?,
        );
    assert!(output_copy.is_file());
    assert!(!String::from_utf8_lossy(&output.stderr).contains("phase"));
    assert_eq!(fs::read(&current)?, current_before);
    assert_eq!(fs::read(&previous)?, previous_before);
    let log_dir = root.path().join("app/logs");
    assert!(!log_dir.exists() || fs::read_dir(log_dir)?.next().is_none());
    Ok(())
}

#[test]
fn interactive_rural_loan_menu_executes_capability_and_safely_replaces_current_workbook(
) -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let (current, previous) = make_rural_loan_periods(root.path(), false, true)?;
    let capability_path = root.path().join("temporary-rural-loan-capability.json");
    let candidate_for_capability = root.path().join("capability-candidate.xlsx");
    let plan = build_rural_loan_plan(&current, &previous, &candidate_for_capability)?
        .0
        .plan
        .context("PT-7-7 capability plan missing")?;
    fs::write(
        &capability_path,
        serde_json::to_vec_pretty(&rural_loan_capability(&plan, "PT-7-7"))?,
    )?;
    let current_before = read_workbook(&current)?;
    let previous_before = fs::read(&previous)?;
    let current_bytes_before = fs::read(&current)?;
    let original_book = excel_editor::open_xlsx(&current)?;
    let original_bold = original_book
        .sheet_by_name("惠州市涉农贷款分机构")?
        .cell("D3")
        .and_then(|cell| cell.style().font())
        .is_some_and(|font| font.bold());
    assert!(original_bold);

    let output = run_menu_with_capability(
        &config,
        &rural_loan_menu_input(7, &current, &previous),
        &capability_path,
        "synthetic-pt-7-7-umya-menu-test",
        "developer",
    )?;
    assert!(
        output.status.success(),
        "PT-7-7 capability menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stdout.contains("[执行模式]"));
    assert!(stdout.contains(r#""operation": "replaced""#));
    assert!(stderr.contains(r#""actual_engine":"umya_spreadsheet""#));
    assert!(stderr.contains(r#""phase":"safe_commit""#));
    assert_ne!(fs::read(&current)?, current_bytes_before);
    assert_eq!(fs::read(&previous)?, previous_before);

    let updated = read_workbook(&current)?;
    assert_eq!(updated.sheets.len(), current_before.sheets.len());
    for (before, after) in current_before.sheets.iter().zip(&updated.sheets) {
        assert_eq!(before.name, after.name);
        assert_eq!(before.formulas, after.formulas);
        assert_eq!(before.merges, after.merges);
        assert_eq!(before.comments, after.comments);
        for (address, expected) in &before.values {
            if before.name == "惠州市涉农贷款分机构" && address == "D3" {
                continue;
            }
            assert_eq!(after.values.get(address), Some(expected), "{address}");
        }
    }
    let adjusted_sheet = updated
        .sheets
        .iter()
        .find(|sheet| sheet.name == "惠州市涉农贷款分机构")
        .context("adjusted institution sheet missing")?;
    let difference = adjusted_sheet.values["D3"]["value"]
        .as_f64()
        .context("adjusted difference is not numeric")?;
    assert!((difference - 5.123456789).abs() <= 1e-15);
    let updated_book = excel_editor::open_xlsx(&current)?;
    let adjusted_sheet = updated_book.sheet_by_name("惠州市涉农贷款分机构")?;
    let adjusted_bold = adjusted_sheet
        .cell("D3")
        .and_then(|cell| cell.style().font())
        .is_some_and(|font| font.bold());
    assert_eq!(adjusted_bold, original_bold);
    assert!(adjusted_sheet
        .merge_cells()
        .iter()
        .any(|range| range.range() == "G1:H1"));
    assert!(adjusted_sheet
        .comments()
        .iter()
        .any(|comment| comment.coordinate().to_string() == "E3"));

    let history_path = find_file_recursively(&root.path().join("app/logs"), ".run.json")?
        .context("PT-7-7 execution history was not written")?;
    let history: Value = serde_json::from_slice(&fs::read(history_path)?)?;
    assert!(history["actual_engines"].as_array().is_some_and(|engines| {
        engines.iter().any(|engine| engine == "calamine")
            && engines.iter().any(|engine| engine == "umya_spreadsheet")
    }));
    Ok(())
}

#[test]
fn interactive_quarterly_rural_loan_menu_executes_capability_and_safely_replaces_current_workbook(
) -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let (current, previous) = make_rural_loan_periods(root.path(), true, true)?;
    let capability_path = root
        .path()
        .join("temporary-quarterly-rural-loan-capability.json");
    let candidate_for_capability = root.path().join("quarterly-capability-candidate.xlsx");
    let plan = build_quarterly_rural_loan_plan(&current, &previous, &candidate_for_capability)?
        .0
        .plan
        .context("PT-7-8 capability plan missing")?;
    fs::write(
        &capability_path,
        serde_json::to_vec_pretty(&rural_loan_capability(&plan, "PT-7-8"))?,
    )?;
    let current_before = read_workbook(&current)?;
    let current_bytes_before = fs::read(&current)?;
    let previous_before = fs::read(&previous)?;

    let output = run_menu_with_capability(
        &config,
        &rural_loan_menu_input(8, &current, &previous),
        &capability_path,
        "synthetic-pt-7-8-umya-menu-test",
        "developer",
    )?;
    assert!(
        output.status.success(),
        "PT-7-8 capability menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stdout.contains("[执行模式]"));
    assert!(stdout.contains(r#""operation": "replaced""#));
    assert!(
        stdout.contains(r#""failed_regions": 1"#),
        "quarterly summary metrics missing from output: {stdout}"
    );
    assert!(stderr.contains(r#""actual_engine":"umya_spreadsheet""#));
    assert!(stderr.contains(r#""phase":"safe_commit""#));
    assert_ne!(fs::read(&current)?, current_bytes_before);
    assert_eq!(fs::read(&previous)?, previous_before);

    let updated = read_workbook(&current)?;
    assert_eq!(updated.sheets.len(), current_before.sheets.len());
    for (before, after) in current_before.sheets.iter().zip(&updated.sheets) {
        assert_eq!(before.name, after.name);
        assert_eq!(before.formulas, after.formulas);
        assert_eq!(before.merges, after.merges);
        assert_eq!(before.comments, after.comments);
        for (address, expected) in &before.values {
            if before.name == "惠州市涉农贷款分机构" && address == "D3" {
                continue;
            }
            assert_eq!(
                after.values.get(address),
                Some(expected),
                "{}!{address}",
                before.name
            );
        }
    }
    let hui_zhou_values = updated
        .sheets
        .iter()
        .find(|sheet| sheet.name == "惠州市涉农贷款分机构")
        .context("quarterly adjusted sheet missing")?;
    let difference = hui_zhou_values.values["D3"]["value"]
        .as_f64()
        .context("quarterly adjustment is not numeric")?;
    assert!((difference - 5.123456789).abs() <= 1e-15);
    let updated_book = excel_editor::open_xlsx(&current)?;
    let hui_zhou = updated_book.sheet_by_name("惠州市涉农贷款分机构")?;
    assert!(hui_zhou
        .cell("D3")
        .and_then(|cell| cell.style().font())
        .is_some_and(|font| font.bold()));
    assert!(hui_zhou
        .merge_cells()
        .iter()
        .any(|range| range.range() == "G1:H1"));
    assert!(hui_zhou
        .comments()
        .iter()
        .any(|comment| comment.coordinate().to_string() == "E3"));

    let history_path = find_file_recursively(&root.path().join("app/logs"), ".run.json")?
        .context("PT-7-8 execution history was not written")?;
    let history: Value = serde_json::from_slice(&fs::read(history_path)?)?;
    assert_eq!(history["feature_id"], "PT-7-8");
    assert_eq!(history["summary_metrics"]["failed_regions"], 1);
    assert!(history["actual_engines"].as_array().is_some_and(|engines| {
        engines.iter().any(|engine| engine == "calamine")
            && engines.iter().any(|engine| engine == "umya_spreadsheet")
    }));
    Ok(())
}

fn create_config(root: &Path) -> Result<PathBuf> {
    let app = root.join("app");
    let config = app.join("config/config.xlsx");
    fs::create_dir_all(config.parent().context("config parent missing")?)?;
    fs::create_dir_all(app.join("input"))?;
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let global = workbook.add_worksheet();
    global.set_name("全局配置")?;
    global.write_string(0, 0, "配置项")?;
    global.write_string(0, 1, "值")?;
    let settings = [
        ("输入目录", app.join("input").to_string_lossy().into_owned()),
        (
            "输出目录",
            app.join("output").to_string_lossy().into_owned(),
        ),
        ("日志目录", app.join("logs").to_string_lossy().into_owned()),
    ];
    for (index, (key, value)) in settings.iter().enumerate() {
        global.write_string((index + 1) as u32, 0, *key)?;
        global.write_string((index + 1) as u32, 1, value)?;
    }
    let mapping = workbook.add_worksheet();
    mapping.set_name("机构映射表")?;
    mapping.write_string(0, 0, "原始机构")?;
    mapping.write_string(0, 1, "标准机构")?;
    for (row, (source, target)) in [
        ("AB", "Short Standard"),
        ("ABC Bank", "Exact Standard"),
        ("branch", "Branch Standard"),
    ]
    .iter()
    .enumerate()
    {
        mapping.write_string((row + 1) as u32, 0, *source)?;
        mapping.write_string((row + 1) as u32, 1, *target)?;
    }
    workbook.save(&config)?;
    Ok(config)
}

fn make_rural_loan_periods(
    root: &Path,
    quarterly: bool,
    preserve_features: bool,
) -> Result<(PathBuf, PathBuf)> {
    let current = root.join("rural-current.xlsx");
    let previous = root.join("rural-previous.xlsx");
    make_rural_loan_workbook(&current, quarterly, true, preserve_features)?;
    make_rural_loan_workbook(&previous, quarterly, false, false)?;
    Ok((current, previous))
}

fn make_rural_loan_workbook(
    path: &Path,
    quarterly: bool,
    current_period: bool,
    preserve_features: bool,
) -> Result<()> {
    let mut workbook = Workbook::new();
    let regions = if quarterly && current_period {
        vec![
            ("惠州市", "甲行", 14.123456789, 9.0),
            ("江门市", "江门行", 8.0, 4.0),
        ]
    } else {
        vec![("惠州市", "甲行", 14.123456789, 9.0)]
    };
    for (region, institution, current_balance, previous_balance) in &regions {
        let sheet = workbook.add_worksheet();
        sheet.set_name(format!("{region}涉农贷款分机构"))?;
        for row in 0..2 {
            sheet.write_string(row, 0, "序号")?;
            sheet.write_string(row, 1, "机构名称")?;
            sheet.write_string(row, 2, "涉农贷款余额")?;
            sheet.write_string(row, 3, "比上月")?;
        }
        sheet.write_number(2, 0, 1.0)?;
        sheet.write_string(2, 1, *institution)?;
        let balance = if current_period {
            *current_balance
        } else {
            *previous_balance
        };
        sheet.write_number(2, 2, balance)?;
        if preserve_features && *region == "惠州市" {
            sheet.write_number_with_format(
                2,
                3,
                999.0,
                &rust_xlsxwriter::Format::new().set_bold(),
            )?;
            sheet.insert_note(2, 4, &rust_xlsxwriter::Note::new("保留批注"))?;
            sheet.write_formula(2, 5, "=C3*2")?;
            sheet.merge_range(
                0,
                6,
                0,
                7,
                "合并标题",
                &rust_xlsxwriter::Format::new().set_bold(),
            )?;
            sheet.set_row_height(8, 21.0)?;
        } else {
            sheet.write_number(2, 3, 999.0)?;
        }
    }

    let area = workbook.add_worksheet();
    area.set_name(if quarterly {
        "涉农贷款分地区统计表"
    } else {
        "本外币涉农贷款分地区"
    })?;
    area.write_string(0, 0, "地区")?;
    area.write_string(0, 1, "涉农贷款比上月")?;
    for (index, (region, _, current_balance, previous_balance)) in regions.iter().enumerate() {
        area.write_string((index + 1) as u32, 0, *region)?;
        let difference = if *region == "江门市" {
            4.0
        } else {
            current_balance - previous_balance
        };
        area.write_number((index + 1) as u32, 1, difference)?;
    }

    if preserve_features {
        let untouched = workbook.add_worksheet();
        untouched.set_name("非目标工作表")?;
        untouched.write_string(0, 0, "保持不变")?;
        untouched.write_formula(1, 0, "='惠州市涉农贷款分机构'!C3")?;
        untouched.merge_range(
            3,
            0,
            3,
            1,
            "非目标合并区域",
            &rust_xlsxwriter::Format::new(),
        )?;
        untouched.insert_note(0, 2, &rust_xlsxwriter::Note::new("非目标批注"))?;
    }
    workbook.save(path)?;
    Ok(())
}

fn make_village_period(path: &Path, current: bool) -> Result<()> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("村镇银行本外币")?;
    for (row, base) in [(4_u32, 100.0), (5_u32, 110.0)] {
        let factor = if current { 1.0 } else { 0.5 };
        sheet.write_number(row, 2, base * factor)?;
        sheet.write_string(row, 3, "存款备注")?;
        sheet.write_number(row, 4, row as f64 + 1.0)?;
        sheet.write_number(row, 86, (base + 100.0) * factor)?;
        sheet.write_string(row, 87, "贷款备注")?;
        sheet.write_number(row, 88, row as f64 + 2.0)?;
    }
    workbook.save(path)?;
    Ok(())
}

fn make_village_target(path: &Path) -> Result<()> {
    let mut workbook = Workbook::new();
    let branch = workbook.add_worksheet();
    branch.set_name("广东分机构本外币")?;
    for (row, name) in [(0_u32, "机构"), (1, "其他"), (2, "村镇银行"), (3, "尾行")] {
        branch.write_string(row, 0, name)?;
    }
    branch.write_number(2, 1, 9.0)?;
    branch.write_formula(3, 9, "=A4")?;
    branch.merge_range(
        3,
        10,
        3,
        11,
        "保留尾行合并区域",
        &rust_xlsxwriter::Format::new(),
    )?;
    branch.insert_note(3, 12, &rust_xlsxwriter::Note::new("保留尾行批注"))?;
    branch.set_row_height(3, 24.0)?;
    let untouched = workbook.add_worksheet();
    untouched.set_name("非目标工作表")?;
    untouched.write_string(0, 0, "标题")?;
    untouched.write_string(1, 0, "保持不变")?;
    untouched.write_formula(1, 1, "='广东分机构本外币'!A4")?;
    untouched.merge_range(
        3,
        3,
        3,
        4,
        "非目标合并区域",
        &rust_xlsxwriter::Format::new(),
    )?;
    untouched.insert_note(1, 5, &rust_xlsxwriter::Note::new("非目标批注"))?;
    untouched.set_row_height(4, 19.0)?;
    workbook.define_name("VillageTail", "'广东分机构本外币'!$J$4")?;
    workbook.save(path)?;
    Ok(())
}

fn village_split_capability(plan: &excel_editor::ExcelEditPlan) -> BackendCapability {
    BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Passed,
        format: plan.format,
        operations: plan.operations.iter().map(EditOperation::kind).collect(),
        supported_features: plan.workbook.features_present.clone(),
        preserved: plan.preserve.clone(),
        platform: "synthetic-pt-7-1-umya-menu-test".to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: "synthetic-pt-7-1-village-split".to_owned(),
            report_path: "pytools-core/tests/batch_tools_menu_cli.rs".to_owned(),
            adapter_version: "umya-spreadsheet-3.1.1".to_owned(),
            engine_version: "umya-spreadsheet-3.1.1".to_owned(),
            platform: "synthetic-pt-7-1-umya-menu-test".to_owned(),
        }),
    }
}

fn institution_normalize_capability(plan: &excel_editor::ExcelEditPlan) -> BackendCapability {
    BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Passed,
        format: plan.format,
        operations: plan.operations.iter().map(EditOperation::kind).collect(),
        supported_features: plan.workbook.features_present.clone(),
        preserved: plan.preserve.clone(),
        platform: "synthetic-pt-7-2-menu-test".to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: "synthetic-pt-7-2-institution-normalize-menu".to_owned(),
            report_path: "pytools-core/tests/batch_tools_menu_cli.rs".to_owned(),
            adapter_version: "umya-spreadsheet-3.1.1".to_owned(),
            engine_version: "umya-spreadsheet-3.1.1".to_owned(),
            platform: "synthetic-pt-7-2-menu-test".to_owned(),
        }),
    }
}

fn rural_loan_capability(
    plan: &excel_editor::ExcelEditPlan,
    feature_id: &str,
) -> BackendCapability {
    let evidence_id = feature_id.to_ascii_lowercase();
    let platform = format!("synthetic-{evidence_id}-umya-menu-test");
    BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Passed,
        format: plan.format,
        operations: plan.operations.iter().map(EditOperation::kind).collect(),
        supported_features: plan.workbook.features_present.clone(),
        preserved: plan.preserve.clone(),
        platform: platform.clone(),
        evidence: Some(CapabilityEvidence {
            sample_id: format!("synthetic-{evidence_id}-rural-loan-menu"),
            report_path: "pytools-core/tests/batch_tools_menu_cli.rs".to_owned(),
            adapter_version: "umya-spreadsheet-3.1.1".to_owned(),
            engine_version: "umya-spreadsheet-3.1.1".to_owned(),
            platform,
        }),
    }
}

fn rural_loan_menu_input(choice: u8, current: &Path, previous: &Path) -> String {
    format!(
        "5\n{choice}\n{}\n{}\n0\n0\n",
        current.display(),
        previous.display()
    )
}

fn village_menu_input(current: &Path, previous: &Path, target: &Path) -> String {
    format!(
        "5\n1\n{}\n{}\n1\n{}\n0\n0\n",
        current.display(),
        previous.display(),
        target.display()
    )
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

fn run_menu(config: &Path, input: &str) -> Result<Output> {
    run_menu_with_arguments(config, input, &[])
}

fn run_menu_with_capability(
    config: &Path,
    input: &str,
    capability: &Path,
    platform: &str,
    log_view: &str,
) -> Result<Output> {
    run_menu_with_arguments(
        config,
        input,
        &[
            "--capabilities".to_owned(),
            capability.to_string_lossy().into_owned(),
            "--platform".to_owned(),
            platform.to_owned(),
            "--log-view".to_owned(),
            log_view.to_owned(),
        ],
    )
}

fn run_menu_with_arguments(config: &Path, input: &str, extra: &[String]) -> Result<Output> {
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
