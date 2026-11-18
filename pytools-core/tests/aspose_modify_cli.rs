use anyhow::{Context, Result};
use excel_editor::{
    CellValue, EditOperation, EngineId, ExcelEditPlan, PreservationRequirement as P, WorkbookFormat,
};
use pytools_core::{
    aspose_edit::{ModifyExecutionScope, ModifyMode},
    excel_edit_router::{route_plan, RouteConfig, RouteMode},
};
use std::{
    collections::BTreeSet,
    fs,
    io::Write,
    path::{Path, PathBuf},
    process::{Command, Stdio},
    time::{SystemTime, UNIX_EPOCH},
};

fn root(label: &str) -> Result<PathBuf> {
    let p = std::env::temp_dir().join(format!(
        "aspose-modify-{label}-{}-{}",
        std::process::id(),
        SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos()
    ));
    fs::create_dir_all(&p)?;
    Ok(p)
}
fn fixture(p: &Path) -> Result<()> {
    let mut b = rust_xlsxwriter::Workbook::new();
    let s = b.add_worksheet().set_name("外汇数据")?;
    s.write_string(0, 0, "原值")?;
    s.write_number(11, 26, 123.0)?;
    s.write_formula(
        1,
        0,
        rust_xlsxwriter::Formula::new("=SUM(1,2)").set_result("3"),
    )?;
    s.set_header("&L单位：万元&R测试");
    s.set_footer("&C万元");
    b.save(p)?;
    Ok(())
}
fn config(root: &Path) -> Result<PathBuf> {
    let p = root.join("config.xlsx");
    let mut b = rust_xlsxwriter::Workbook::new();
    let s = b.add_worksheet().set_name("全局配置")?;
    for (i, (k, v)) in [("键", "值"), ("输出目录", "output"), ("日志目录", "logs")]
        .into_iter()
        .enumerate()
    {
        s.write_string(i as u32, 0, k)?;
        s.write_string(i as u32, 1, v)?;
    }
    b.save(&p)?;
    Ok(p)
}
fn plan(p: &Path) -> Result<ExcelEditPlan> {
    Ok(ExcelEditPlan {
        request_id: format!(
            "aspose-regression-{}",
            p.file_stem().unwrap().to_string_lossy()
        ),
        input: p.into(),
        output: p.with_extension("candidate.xlsx"),
        format: WorkbookFormat::Xlsx,
        workbook: excel_editor::inspect_xlsx_profile(p)?,
        operations: vec![EditOperation::SetCellValue {
            sheet: "外汇数据".into(),
            cell: "AA12".into(),
            value: CellValue::Text("新值".into()),
        }],
        preserve: BTreeSet::from([
            P::CellValuesAndTypes,
            P::FormulaText,
            P::CommentsAndPositions,
            P::Styles,
            P::MergedRanges,
            P::SheetIdentityAndOrder,
        ]),
        drop_allowed: Default::default(),
    })
}
#[test]
fn unsupported_and_precision_routes_fail_before_any_write() -> Result<()> {
    let r = root("unsupported")?;
    let p = r.join("input.xlsx");
    fixture(&p)?;
    let before = fs::read(&p)?;
    let mut plan = plan(&p)?;
    for value in [
        CellValue::Decimal("1234567890123456".into()),
        CellValue::Decimal("1.234567890123456789".into()),
        CellValue::Decimal("NaN".into()),
        CellValue::Error("#UNVERIFIED!".into()),
        CellValue::DateTime {
            serial: "not-a-serial".into(),
            number_format: "yyyy-mm-dd".into(),
        },
    ] {
        plan.operations = vec![EditOperation::SetCellValue {
            sheet: "外汇数据".into(),
            cell: "AA12".into(),
            value,
        }];
        assert!(route_plan(&plan, &RouteConfig::new(RouteMode::AsposeOnly)).is_err());
        assert!(!plan.output.exists());
        assert_eq!(before, fs::read(&p)?);
    }
    plan.operations = vec![EditOperation::DeleteRows {
        sheet: "外汇数据".into(),
        first_row: 2,
        count: 1,
    }];
    let routed = route_plan(&plan, &RouteConfig::new(RouteMode::AsposeOnly));
    if cfg!(any(windows, target_os = "linux")) {
        assert_eq!(routed?.engine, EngineId::AsposeWorker);
    } else {
        // Unavailable platforms must fail before a worker or candidate is created.
        let error = routed.expect_err("Aspose must remain unavailable on this platform");
        assert!(format!("{error:#}").contains("Windows"));
    }
    assert!(!plan.output.exists());
    assert_eq!(before, fs::read(p)?);
    Ok(())
}

#[cfg(any(windows, target_os = "linux"))]
#[test]
#[ignore = "requires explicitly configured resident Aspose Worker; synthetic isolated candidates"]
fn real_worker_structural_plan_primitives_validate_before_publish() -> Result<()> {
    let r = root("structural-primitives")?;
    let _scope = ModifyExecutionScope::enter(ModifyMode::Aspose)?;
    for (label, operations) in [
        (
            "delete-rows",
            vec![EditOperation::DeleteRows {
                sheet: "外汇数据".into(),
                first_row: 10,
                count: 2,
            }],
        ),
        (
            "rename",
            vec![EditOperation::RenameSheet {
                old_name: "外汇数据".into(),
                new_name: "新的名称".into(),
            }],
        ),
        (
            "dimensions",
            vec![
                EditOperation::SetRowHeight {
                    sheet: "外汇数据".into(),
                    row_index: 12,
                    height: 27.0,
                },
                EditOperation::SetColumnWidth {
                    sheet: "外汇数据".into(),
                    column_index: 27,
                    width: 24.0,
                },
            ],
        ),
        (
            "comment",
            vec![EditOperation::SetComment {
                sheet: "外汇数据".into(),
                cell: "AA12".into(),
                text: Some("合成批注".into()),
                author: Some("测试".into()),
            }],
        ),
        (
            "typed-values",
            vec![
                EditOperation::SetCellValue {
                    sheet: "外汇数据".into(),
                    cell: "Y10".into(),
                    value: CellValue::Decimal("5.1234567890000005".into()),
                },
                EditOperation::SetCellValue {
                    sheet: "外汇数据".into(),
                    cell: "AA12".into(),
                    value: CellValue::Decimal("1.2".into()),
                },
                EditOperation::SetCellValue {
                    sheet: "外汇数据".into(),
                    cell: "Z10".into(),
                    value: CellValue::Error("#N/A".into()),
                },
                EditOperation::SetCellValue {
                    sheet: "外汇数据".into(),
                    cell: "Z11".into(),
                    value: CellValue::DateTime {
                        serial: "45000".into(),
                        number_format: "yyyy-mm-dd".into(),
                    },
                },
            ],
        ),
        (
            "new-sheet",
            vec![
                EditOperation::CreateSheetAt {
                    name: "新表".into(),
                    index: 0,
                },
                EditOperation::MoveSheet {
                    sheet: "新表".into(),
                    new_index: 1,
                },
            ],
        ),
    ] {
        let p = r.join(format!("{label}.xlsx"));
        fixture(&p)?;
        let original = fs::read(&p)?;
        let mut edit = plan(&p)?;
        edit.operations = operations;
        edit.preserve
            .extend([P::RowAndColumnDimensions, P::DefinedNames, P::PageSettings]);
        pytools_core::inplace_edit::execute_candidate(
            &edit,
            pytools_core::inplace_edit::InPlaceEngine::Openpyxl,
        )
        .with_context(|| format!("structural case {label}; retained at {}", r.display()))?;
        assert_eq!(original, fs::read(&p)?);
        assert!(edit.output.exists());
    }
    Ok(())
}

#[cfg(any(windows, target_os = "linux"))]
#[test]
#[ignore = "requires real Aspose Worker; checks import loads across resident generations"]
fn real_worker_rotates_before_multi_workbook_request_budget() -> Result<()> {
    use excel_aspose::{
        AsposeWorkerBackend, ModifyEngine, ModifyRequest, WorkbookOperation, WorkbookRequest,
        WorkerConfig,
    };
    let r = root("import-budget")?;
    let source = r.join("source.xlsx");
    fixture(&source)?;
    let first = r.join("first.xlsx");
    let second = r.join("second.xlsx");
    fs::copy(&source, &first)?;
    fs::copy(&source, &second)?;
    let original = fs::read(&source)?;
    let mut backend = AsposeWorkerBackend::new(WorkerConfig {
        executable: PathBuf::from(
            std::env::var_os("RUSTTOOLS_ASPOSE_DOTNET").context("DOTNET required")?,
        ),
        arguments: vec![
            std::env::var("RUSTTOOLS_ASPOSE_WORKER_DLL")?,
            "--stdio".into(),
        ],
        diagnostics_dir: r.join("logs"),
        startup_timeout: std::time::Duration::from_secs(30),
        request_timeout: std::time::Duration::from_secs(180),
    });
    let mut pid = 0;
    for index in 0..97 {
        let result = backend.modify(&ModifyRequest {
            workbook: WorkbookRequest {
                run_id: "import-budget".into(),
                input: source.clone(),
                candidate: r.join(format!("candidate-{index}.xlsx")),
            },
            operations: vec![WorkbookOperation::SetValue {
                sheet: "外汇数据".into(),
                cell: "A1".into(),
                value: excel_aspose::CellValue::Text("合成修改".into()),
            }],
        })?;
        if index == 0 {
            pid = result.worker_pid;
        }
        assert_eq!(result.worker_pid, pid);
        assert_eq!(result.worker_attempts, index + 1);
    }
    let operations = [first, second].iter().enumerate().map(|(i, path)| WorkbookOperation::EditPlanOperation {
        operation: serde_json::json!({"operation":"import_range","source_workbook":path,"source_sheet":"外汇数据","source_range":"A1:A1","target_sheet":"外汇数据","target_start":format!("B{}",20+i),"copy_values":true,"copy_formulas":false,"style_mode":"none","copy_comments":false,"copy_merges":false}),
    }).collect();
    let result = backend.modify(&ModifyRequest {
        workbook: WorkbookRequest {
            run_id: "import-budget".into(),
            input: source.clone(),
            candidate: r.join("imported.xlsx"),
        },
        operations,
    })?;
    assert_ne!(result.worker_pid, pid);
    assert_eq!(backend.generations(), 2);
    assert_eq!(result.worker_attempts, 3);
    assert_eq!(original, fs::read(source)?);
    Ok(())
}

#[cfg(any(windows, target_os = "linux"))]
#[test]
#[ignore = "requires real Aspose Worker; validates large instruction sidecars and cleanup"]
fn real_worker_large_instruction_batch_uses_verified_temporary_sidecar() -> Result<()> {
    use excel_aspose::{
        AsposeWorkerBackend, ModifyEngine, ModifyRequest, WorkbookOperation, WorkbookRequest,
        WorkerConfig,
    };
    let r = root("large-instructions")?;
    let source = r.join("source.xlsx");
    fixture(&source)?;
    let original = fs::read(&source)?;
    let mut operations = (0..10000)
        .map(|i| WorkbookOperation::SetValue {
            sheet: "外汇数据".into(),
            cell: "A1".into(),
            value: excel_aspose::CellValue::Text(format!("合成大批量修改{i}")),
        })
        .collect::<Vec<_>>();
    operations.push(WorkbookOperation::SetValue {
        sheet: "外汇数据".into(),
        cell: "AA100".into(),
        value: excel_aspose::CellValue::Text("最后一条指令".into()),
    });
    assert!(serde_json::to_vec(&operations)?.len() > 1024 * 1024);
    let mut backend = AsposeWorkerBackend::new(WorkerConfig {
        executable: PathBuf::from(
            std::env::var_os("RUSTTOOLS_ASPOSE_DOTNET").context("DOTNET required")?,
        ),
        arguments: vec![
            std::env::var("RUSTTOOLS_ASPOSE_WORKER_DLL")?,
            "--stdio".into(),
        ],
        diagnostics_dir: r.join("logs"),
        startup_timeout: std::time::Duration::from_secs(30),
        request_timeout: std::time::Duration::from_secs(180),
    });
    let result = backend.modify(&ModifyRequest {
        workbook: WorkbookRequest {
            run_id: "large-instructions".into(),
            input: source.clone(),
            candidate: r.join("candidate.xlsx"),
        },
        operations,
    })?;
    let saved = excel_reader::read_workbook(&result.path)?;
    assert_eq!(saved.sheets[0].values["A1"]["value"], "合成大批量修改9999");
    assert_eq!(saved.sheets[0].values["AA100"]["value"], "最后一条指令");
    assert_eq!(original, fs::read(source)?);
    assert!(!fs::read_dir(r.join("logs"))?
        .filter_map(|e| e.ok())
        .any(|e| e.file_name().to_string_lossy().starts_with("operations-")));
    Ok(())
}
#[test]
fn modifier_menu_goes_to_file_selection_and_cancel_never_saves_config() -> Result<()> {
    for (sub, selection) in [
        ("2", "选择需要规范化机构名称的目标工作簿"),
        ("4", "选择要修改外汇页眉的 Excel 文件"),
    ] {
        let r = root("menu-cancel")?;
        let c = config(&r)?;
        let before = fs::read(&c)?;
        let mut child = Command::new(env!("CARGO_BIN_EXE_RustTools"))
            .current_dir(&r)
            .args(["menu", "--config"])
            .arg(&c)
            .env("PYTOOLS_CONFIG_PATH", &c)
            .env("RUSTTOOLS_ASPOSE_WORKER_EXE", r.join("missing.exe"))
            .stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .stderr(Stdio::piped())
            .spawn()?;
        child
            .stdin
            .take()
            .unwrap()
            .write_all(format!("5\n{sub}\n0\n0\n").as_bytes())?;
        let out = child.wait_with_output()?;
        let text = String::from_utf8_lossy(&out.stdout);
        assert!(
            out.status.success(),
            "{text} {}",
            String::from_utf8_lossy(&out.stderr)
        );
        assert!(!text.contains("修改执行模式"), "{text}");
        assert!(text.contains(selection), "{text}");
        assert!(text.contains("执行路径："), "{text}");
        assert_eq!(before, fs::read(c)?);
    }
    Ok(())
}
#[test]
fn retired_aspose_modify_does_not_fallback_or_touch_source() -> Result<()> {
    let r = root("missing")?;
    let p = r.join("input.xlsx");
    fixture(&p)?;
    let before = fs::read(&p)?;
    let o = Command::new(env!("CARGO_BIN_EXE_RustTools"))
        .current_dir(&r)
        .args([
            "fx-header",
            "--modify-engine",
            "aspose",
            "--in-place",
            "auto",
        ])
        .arg(&p)
        .env("RUSTTOOLS_ASPOSE_WORKER_EXE", r.join("missing.exe"))
        .output()?;
    assert!(!o.status.success());
    let error = String::from_utf8_lossy(&o.stderr);
    assert!(error.contains("仅保留用于 XLS→XLSX"), "{error}");
    assert!(!r.join("input.candidate.xlsx").exists());
    assert_eq!(before, fs::read(p)?);
    Ok(())
}

#[test]
fn explicit_openpyxl_rejects_xls_before_legacy_com() -> Result<()> {
    let r = root("no-com")?;
    let p = r.join("source.xls");
    let original = b"synthetic XLS must not reach COM";
    fs::write(&p, original)?;
    let o = Command::new(env!("CARGO_BIN_EXE_RustTools"))
        .current_dir(&r)
        .args([
            "fx-header",
            "--modify-engine",
            "openpyxl",
            "--in-place",
            "auto",
            "--view",
            "developer",
            "--output-dir",
        ])
        .arg(r.join("output"))
        .arg(&p)
        .env("PYTOOLS_OPENPYXL_PYTHON", r.join("missing.exe"))
        .output()?;
    ensure_success(&o)?;
    let v: serde_json::Value = serde_json::from_slice(&o.stdout)?;
    assert_eq!(v["failed_files"], 1, "{v}");
    assert!(
        v["items"][0]["error"]
            .as_str()
            .unwrap()
            .contains("所选修改模式仅支持 XLSX"),
        "{v}"
    );
    assert_eq!(fs::read(&p)?, original);
    let _scope = ModifyExecutionScope::enter(ModifyMode::Openpyxl)?;
    let task = pytools_core::config::DedupTask {
        config_sheet_name: "查重配置".into(),
        row_no: 2,
        enabled: true,
        task_name: "synthetic".into(),
        source_wb: Some(p.clone()),
        source_ws: "Sheet1".into(),
        key_cols: vec![1],
        target_wb: None,
        target_ws: String::new(),
        write_mode: String::new(),
    };
    let result = pytools_core::dedup_ops::mark_duplicates_in_place(&[task])?;
    assert_eq!(result.task_skipped, 1);
    assert_eq!(fs::read(p)?, original);
    Ok(())
}

#[test]
fn sanitized_validator_preserves_blank_and_reserved_business_names_without_mutation() -> Result<()> {
    let r = root("sanitized-business-names")?;
    let source = r.join("source.xlsx");
    let mut book = rust_xlsxwriter::Workbook::new();
    for name in ["Data", "Blank", "Evaluation Warning", "evaluation warning 2"] {
        let sheet = book.add_worksheet();
        sheet.set_name(name)?;
        if name != "Blank" { sheet.write_string(0, 0, "business value")?; }
    }
    book.save(&source)?;
    let candidate = r.join("candidate.xlsx");
    fs::copy(&source, &candidate)?;
    let original = fs::read(&source)?;
    let before = fs::read(&candidate)?;
    let mut p = plan(&source)?;
    p.operations = vec![EditOperation::SetCellValue {
        sheet: "Data".into(), cell: "B10".into(), value: CellValue::Blank,
    }];
    p.preserve = BTreeSet::from([P::CellValuesAndTypes, P::SheetIdentityAndOrder]);
    excel_editor::OpenpyxlXlsxValidator.validate_sanitized_aspose(&p, &candidate)?;
    assert_eq!(fs::read(&source)?, original);
    assert_eq!(fs::read(&candidate)?, before);
    Ok(())
}

#[test]
fn sanitized_aspose_validator_requires_business_roster_and_preserves_other_checks() -> Result<()> {
    use excel_editor::OutputValidator;
    let r = root("evaluation-validation")?;
    let source = r.join("source.xlsx");
    let write = |p: &Path, extra: &[&str], changed: bool| -> Result<()> {
        let mut b = rust_xlsxwriter::Workbook::new();
        b.add_worksheet().set_name("数据")?.write_string(
            0,
            0,
            if changed { "错误值" } else { "保留值" },
        )?;
        for name in extra {
            b.add_worksheet()
                .set_name(*name)?
                .write_string(0, 0, "extra")?;
        }
        b.save(p)?;
        Ok(())
    };
    write(&source, &[], false)?;
    let original = fs::read(&source)?;
    let mut p = plan(&source)?;
    p.operations = vec![EditOperation::SetCellValue {
        sheet: "数据".into(),
        cell: "B10".into(),
        value: CellValue::Blank,
    }];
    p.preserve = BTreeSet::from([P::CellValuesAndTypes, P::SheetIdentityAndOrder]);
    let candidate = r.join("candidate.xlsx");
    write(&candidate, &[], false)?;
    excel_editor::OpenpyxlXlsxValidator.validate_sanitized_aspose(&p, &candidate)?;
    excel_editor::OpenpyxlXlsxValidator.validate(&p, &candidate)?;
    p.preserve.insert(P::RowAndColumnDimensions);
    // An explicit unchanged default height is equivalent; a real height change
    // must still reject the candidate. No source or candidate is normalized on disk.
    for (height, expected_ok) in [("15", true), ("16", false)] {
        use std::io::Read;
        write(&candidate, &[], false)?;
        let mut zip = zip::ZipArchive::new(fs::File::open(&candidate)?)?;
        let mut entries = Vec::new();
        for i in 0..zip.len() {
            let mut entry = zip.by_index(i)?;
            let name = entry.name().to_owned();
            let mut bytes = Vec::new();
            entry.read_to_end(&mut bytes)?;
            if name == "xl/worksheets/sheet1.xml" {
                let xml = String::from_utf8(bytes)?;
                bytes = xml
                    .replacen("<row r=", &format!("<row ht=\"{height}\" r="), 1)
                    .into_bytes();
            }
            entries.push((name, bytes));
        }
        drop(zip);
        let mut zip = zip::ZipWriter::new(fs::File::create(&candidate)?);
        for (name, bytes) in entries {
            zip.start_file(name, zip::write::SimpleFileOptions::default())?;
            zip.write_all(&bytes)?;
        }
        zip.finish()?;
        assert_eq!(
            excel_editor::OpenpyxlXlsxValidator
                .validate_sanitized_aspose(&p, &candidate)
                .is_ok(),
            expected_ok
        );
    }
    for (extra, changed) in [
        (vec!["wrong"], false),
        (vec!["Evaluation Warning", "another"], false),
        (vec!["Evaluation Warning"], false),
        (vec!["Evaluation Warning"], true),
    ] {
        write(&candidate, &extra, changed)?;
        assert!(excel_editor::OpenpyxlXlsxValidator
            .validate_sanitized_aspose(&p, &candidate)
            .is_err());
        assert_eq!(original, fs::read(&source)?);
    }
    Ok(())
}

#[cfg(any(windows, target_os = "linux"))]
#[test]
#[ignore = "requires ASPOSE_TEST_PYTHON; protocol failure injection, not real Aspose acceptance"]
fn candidate_validation_failure_and_worker_loss_preserve_original_without_replay() -> Result<()> {
    let python = std::env::var_os("ASPOSE_TEST_PYTHON").context("set ASPOSE_TEST_PYTHON")?;
    for mode in ["invalid_edit", "crash"] {
        let r = root(mode)?;
        let p = r.join("输入.xlsx");
        fixture(&p)?;
        let original = fs::read(&p)?;
        let record = r.join("requests.txt");
        let o = Command::new(env!("CARGO_BIN_EXE_RustTools"))
            .current_dir(&r)
            .args([
                "fx-header",
                "--modify-engine",
                "aspose",
                "--in-place",
                "auto",
                "--view",
                "developer",
                "--output-dir",
            ])
            .arg(r.join("output"))
            .arg(&p)
            .env_remove("RUSTTOOLS_ASPOSE_WORKER_EXE")
            .env("RUSTTOOLS_ASPOSE_DOTNET", &python)
            .env(
                "RUSTTOOLS_ASPOSE_WORKER_DLL",
                Path::new(env!("CARGO_MANIFEST_DIR"))
                    .join("tests/fixtures/aspose_modify_fault_worker.py"),
            )
            .env("ASPOSE_FAULT_MODE", mode)
            .env("ASPOSE_FAULT_RECORD", &record)
            .output()?;
        ensure_success(&o)?;
        let v: serde_json::Value = serde_json::from_slice(&o.stdout)?;
        assert_eq!(v["failed_files"], 1, "{v}");
        assert_eq!(v["committed_files"], 0, "{v}");
        let error = v["items"][0]["error"]
            .as_str()
            .context("missing final error")?;
        if mode == "invalid_edit" {
            assert!(error.contains("OutputValidationFailed"), "{error}");
        } else {
            assert!(error.contains("NOT automatically retried"), "{error}");
        }
        assert_eq!(original, fs::read(&p)?);
        assert_eq!(fs::read_to_string(record)?, "Modify\n");
        assert!(!fs::read_dir(&r)?.filter_map(|e| e.ok()).any(|e| e
            .file_name()
            .to_string_lossy()
            .starts_with(".aspose-worker")));
    }
    Ok(())
}

#[cfg(any(windows, target_os = "linux"))]
#[test]
#[ignore = "requires explicitly configured Aspose .NET Worker; synthetic files only"]
fn real_worker_modifies_value_fill_formula_and_reuses_batch_process() -> Result<()> {
    let r = root("real")?;
    let _scope = ModifyExecutionScope::enter(ModifyMode::Aspose)?;
    let run_id = r.file_name().unwrap().to_string_lossy().into_owned();
    pytools_core::aspose_edit::bind_run_id(&run_id);
    for i in 0..2 {
        let p = r.join(format!("输入{i}.xlsx"));
        fixture(&p)?;
        let original = fs::read(&p)?;
        let mut plan = plan(&p)?;
        plan.operations.push(EditOperation::SetCellFillColor {
            sheet: "外汇数据".into(),
            cell: "AA12".into(),
            argb: "FFFF0000".into(),
        });
        plan.operations.push(EditOperation::SetFormulaText {
            sheet: "外汇数据".into(),
            cell: "B10".into(),
            formula: "=SUM(4,5)".into(),
        });
        let result = pytools_core::inplace_edit::execute_in_place(
            &plan,
            pytools_core::inplace_edit::InPlaceEngine::Openpyxl,
        )?;
        assert_eq!(result.execution.engine, EngineId::AsposeWorker);
        let dump = excel_reader::read_workbook(&p)?;
        assert_eq!(dump.sheets.iter().map(|s| s.name.as_str()).collect::<Vec<_>>(), ["外汇数据"]);
        let wb = umya_spreadsheet::reader::xlsx::read(&p)?;
        let s = wb.get_sheet_by_name("外汇数据").unwrap();
        assert_eq!(s.get_value("AA12"), "新值");
        assert_eq!(s.get_cell("B10").unwrap().get_formula(), "SUM(4,5)");
        assert_eq!(
            result.commit.sha256_before.as_deref(),
            Some(hex_hash(&original).as_str())
        );
        assert!(!plan.output.exists());
    }
    let diagnostics = pytools_core::run_logging::default_log_dir()
        .join("aspose-modify")
        .join(&run_id)
        .join("host.jsonl");
    let records = fs::read_to_string(diagnostics)?
        .lines()
        .map(serde_json::from_str::<serde_json::Value>)
        .collect::<std::result::Result<Vec<_>, _>>()?;
    assert_eq!(records.len(), 2);
    assert!(records[0]["result"]["worker_pid"].as_u64().is_some_and(|pid| pid > 0));
    assert_eq!(
        records[0]["result"]["worker_pid"],
        records[1]["result"]["worker_pid"]
    );
    assert_eq!(records[0]["result"]["worker_attempts"], 1);
    assert_eq!(records[1]["result"]["worker_attempts"], 2);
    assert_eq!(records[1]["result"]["timings"]["startup_ms"], 0.0);
    println!("REAL_ASPOSE_MODIFY={}", r.display());
    Ok(())
}

#[cfg(any(windows, target_os = "linux"))]
#[test]
#[ignore = "requires explicitly configured Aspose .NET Worker; synthetic institution batch"]
fn real_institution_batch_and_legacy_format_rejection() -> Result<()> {
    let r = root("institution")?;
    let c = r.join("mapping.xlsx");
    let mut config = rust_xlsxwriter::Workbook::new();
    let m = config.add_worksheet().set_name("机构映射表")?;
    for (row, a, b) in [(0, "原始机构", "标准机构"), (1, "ABC", "标准机构")] {
        m.write_string(row, 0, a)?;
        m.write_string(row, 1, b)?;
    }
    config.save(&c)?;
    let original_config = fs::read(&c)?;
    let mut sources = Vec::new();
    for i in 0..2 {
        let p = r.join(format!("机构{i}.xlsx"));
        let mut b = rust_xlsxwriter::Workbook::new();
        let s = b.add_worksheet().set_name("示例分机构")?;
        s.write_string(0, 0, "机构")?;
        s.write_string(1, 0, "ABC")?;
        s.write_number(1, 1, 3.0)?;
        b.save(&p)?;
        sources.push(p);
    }
    let legacy = r.join("legacy.xls");
    fs::write(
        &legacy,
        b"synthetic invalid BIFF, must refuse before COM/read",
    )?;
    let history = r.join("history.json");
    let o = Command::new(env!("CARGO_BIN_EXE_RustTools"))
        .current_dir(&r)
        .args([
            "institution-normalize",
            "--modify-engine",
            "aspose",
            "--in-place",
            "auto",
            "--apply",
            "--config",
        ])
        .arg(&c)
        .args(["--view", "developer", "--report-output"])
        .arg(&history)
        .args(&sources)
        .arg(&legacy)
        .env("PYTOOLS_OPENPYXL_PYTHON", r.join("no-python.exe"))
        .output()?;
    fs::write(r.join("stdout.txt"), &o.stdout)?;
    fs::write(r.join("stderr.txt"), &o.stderr)?;
    ensure_success(&o)?;
    let h: serde_json::Value = serde_json::from_slice(&fs::read(history)?)?;
    assert_eq!(h["completed_files"], 2, "{h}");
    assert_eq!(h["failed_files"], 1, "{h}");
    for p in sources {
        let b = excel_reader::read_workbook(p)?;
        assert_eq!(
            b.sheets[0].values["A2"]["value"],
            serde_json::json!("标准机构")
        );
        assert_eq!(b.sheets.iter().map(|s| s.name.as_str()).collect::<Vec<_>>(), ["示例分机构"]);
    }
    assert_eq!(
        fs::read(legacy)?,
        b"synthetic invalid BIFF, must refuse before COM/read"
    );
    assert_eq!(original_config, fs::read(c)?);
    println!("REAL_ASPOSE_INSTITUTION={}", r.display());
    Ok(())
}

#[cfg(any(windows, target_os = "linux"))]
#[test]
#[ignore = "requires explicitly configured Aspose .NET Worker; synthetic configured duplicate marking"]
fn real_duplicate_marking_uses_selected_aspose() -> Result<()> {
    let r = root("duplicates")?;
    let p = r.join("duplicates.xlsx");
    let mut b = rust_xlsxwriter::Workbook::new();
    let s = b.add_worksheet().set_name("数据")?;
    for (row, value) in [(0, "名称"), (1, "same"), (2, "same")] {
        s.write_string(row, 0, value)?;
    }
    b.save(&p)?;
    let task = pytools_core::config::DedupTask {
        config_sheet_name: "查重配置".into(),
        row_no: 2,
        enabled: true,
        task_name: "test".into(),
        source_wb: Some(p.clone()),
        source_ws: "数据".into(),
        key_cols: vec![1],
        target_wb: None,
        target_ws: String::new(),
        write_mode: String::new(),
    };
    let _scope = ModifyExecutionScope::enter(ModifyMode::Aspose)?;
    let legacy = r.join("unsafe.xls");
    fs::write(&legacy, b"synthetic XLS - never open COM")?;
    let mut legacy_task = task.clone();
    legacy_task.source_wb = Some(legacy.clone());
    let run = pytools_core::dedup_ops::mark_duplicates_in_place(&[task, legacy_task])?;
    assert_eq!(run.task_failed, 0, "{run:?}");
    assert!(run.marked_rows > 0, "{run:?}");
    assert_eq!(run.task_skipped, 1, "{run:?}");
    assert!(
        run.events
            .iter()
            .any(|event| event.engine == "aspose_worker"),
        "{run:?}"
    );
    assert_eq!(fs::read(legacy)?, b"synthetic XLS - never open COM");
    assert_eq!(excel_reader::read_workbook(&p)?.sheets.iter().map(|s| s.name.as_str()).collect::<Vec<_>>(), ["数据"]);
    println!("REAL_ASPOSE_DUPLICATES={}", r.display());
    Ok(())
}

#[test]
#[ignore = "requires bundled/explicit isolated Openpyxl runtime; selected route regression"]
fn forced_openpyxl_header_uses_semantic_validation_without_aspose() -> Result<()> {
    let r = root("openpyxl")?;
    let p = r.join("外汇.xlsx");
    fixture(&p)?;
    let out = Command::new(env!("CARGO_BIN_EXE_RustTools"))
        .current_dir(&r)
        .args([
            "fx-header",
            "--modify-engine",
            "openpyxl",
            "--in-place",
            "auto",
            "--view",
            "developer",
            "--output-dir",
        ])
        .arg(r.join("output"))
        .arg(&p)
        .env("RUSTTOOLS_ASPOSE_WORKER_EXE", r.join("missing.exe"))
        .output()?;
    ensure_success(&out)?;
    let result: serde_json::Value = serde_json::from_slice(&out.stdout)?;
    assert_eq!(result["failed_files"], 0, "{result}");
    assert_eq!(excel_reader::read_workbook(&p)?.sheets.len(), 1);
    assert!(excel_reader::read_header_inventory(&p)?[0]
        .headers
        .iter()
        .any(|h| h.text.contains("万美元")));
    Ok(())
}
#[cfg(any(windows, target_os = "linux"))]
#[test]
#[ignore = "requires explicitly configured Aspose .NET Worker; real CLI on synthetic copies"]
fn real_header_cli_commits_validated_clean_business_output() -> Result<()> {
    let r = root("header-cli")?;
    let p = r.join("外汇.xlsx");
    fixture(&p)?;
    let o = Command::new(env!("CARGO_BIN_EXE_RustTools"))
        .current_dir(&r)
        .args([
            "fx-header",
            "--modify-engine",
            "aspose",
            "--in-place",
            "auto",
            "--view",
            "developer",
            "--output-dir",
        ])
        .arg(r.join("output"))
        .arg(&p)
        .output()?;
    fs::write(r.join("stdout.txt"), &o.stdout)?;
    fs::write(r.join("stderr.txt"), &o.stderr)?;
    ensure_success(&o)?;
    let v: serde_json::Value = serde_json::from_slice(&o.stdout)?;
    assert_eq!(v["failed_files"], 0, "{v}");
    let inv = excel_reader::read_header_inventory(&p)?;
    assert!(
        inv[0].headers.iter().any(|h| h.text.contains("万美元")),
        "{inv:?}"
    );
    assert_eq!(excel_reader::read_workbook(&p)?.sheets.iter().map(|s| s.name.as_str()).collect::<Vec<_>>(), ["外汇数据"]);
    println!("REAL_ASPOSE_HEADER={}", r.display());
    Ok(())
}
fn ensure_success(o: &std::process::Output) -> Result<()> {
    anyhow::ensure!(
        o.status.success(),
        "{} {}",
        String::from_utf8_lossy(&o.stdout),
        String::from_utf8_lossy(&o.stderr)
    );
    Ok(())
}

fn hex_hash(bytes: &[u8]) -> String {
    use sha2::{Digest, Sha256};
    format!("{:x}", Sha256::digest(bytes))
}
