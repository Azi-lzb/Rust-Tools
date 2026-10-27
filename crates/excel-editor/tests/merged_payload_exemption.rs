//! End-to-end preflight coverage for the merged non-anchor payload
//! baseline-equivalence allowance, registered by the Audit S1-F01 flow: the
//! openpyxl preflight must keep rejecting merged non-anchor payloads by
//! default and let them through only for worksheets with an explicit
//! `MergedNonAnchorPayloads` allowance.
//!
//! These tests run on Windows without the packaged openpyxl runtime: the
//! rejection path fails before the worker spawns, and the allowance path is
//! expected to fail at worker startup (dummy executable), proving the plan
//! passed the payload preflight.

#![cfg(windows)]

use anyhow::Result;
use excel_editor::{
    inspect_xlsx_profile, BackendCapability, CapabilityStatus, EditBackend, EditOperation,
    EditOperationKind, EngineId, ExcelEditPlan, OpenpyxlWorkerConfig, OpenpyxlXlsxBackend,
    PreservationRequirement, WorkbookFormat,
};
use std::{
    fs,
    path::{Path, PathBuf},
    time::{SystemTime, UNIX_EPOCH},
};

struct TestDir(PathBuf);

impl TestDir {
    fn new(label: &str) -> Self {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("system clock after epoch")
            .as_nanos();
        let path = std::env::temp_dir().join(format!("excel-editor-payload-e2e-{label}-{nonce}"));
        fs::create_dir_all(&path).expect("create test directory");
        Self(path)
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

/// Workbook with a payload hidden on the non-anchor cell B1 of merged range
/// A1:C1, shaped like the real S1 formula-copy target sheets.
fn write_input(path: &Path) -> Result<()> {
    let mut workbook = umya_spreadsheet::new_file();
    let sheet = workbook.sheet_by_name_mut("Sheet1")?;
    sheet.set_name("汇总");
    sheet.cell_mut("A1").set_value_string("标题");
    sheet.cell_mut("B1").set_value_string("隐藏载荷");
    sheet.cell_mut("F5").set_value_string("unchanged");
    sheet.add_merge_cells("A1:C1");
    umya_spreadsheet::writer::xlsx::write(&workbook, path)?;
    Ok(())
}

fn plan_for(dir: &TestDir, allowance: Option<&str>) -> Result<ExcelEditPlan> {
    let input = dir.path("input.xlsx");
    write_input(&input)?;
    let mut plan = ExcelEditPlan {
        request_id: "merged-payload-allowance".to_owned(),
        input: input.clone(),
        output: dir.path("output.xlsx"),
        format: WorkbookFormat::Xlsx,
        workbook: inspect_xlsx_profile(&input)?,
        operations: vec![EditOperation::SetCellValue {
            sheet: "汇总".to_owned(),
            cell: "F10".to_owned(),
            value: excel_editor::CellValue::Text("plan write".to_owned()),
        }],
        preserve: [
            PreservationRequirement::CellValuesAndTypes,
            PreservationRequirement::MergedRanges,
        ]
        .into_iter()
        .collect(),
        drop_allowed: Default::default(),
    };
    if let Some(sheet) = allowance {
        plan.drop_allowed.insert(
            sheet.to_owned(),
            [PreservationRequirement::MergedNonAnchorPayloads]
                .into_iter()
                .collect(),
        );
    }
    plan.validate()?;
    Ok(plan)
}

fn backend(dir: &TestDir, plan: &ExcelEditPlan) -> Result<OpenpyxlXlsxBackend> {
    let dummy_python = dir.path("dummy-python.txt");
    fs::write(&dummy_python, b"not a real executable")?;
    let worker_script = Path::new(env!("CARGO_MANIFEST_DIR")).join("python/openpyxl_worker.py");
    OpenpyxlXlsxBackend::new(
        BackendCapability {
            engine: EngineId::OpenpyxlWorker,
            status: CapabilityStatus::Unverified,
            format: WorkbookFormat::Xlsx,
            operations: [EditOperationKind::SetCellValue].into_iter().collect(),
            supported_features: plan.workbook.features_present.clone(),
            preserved: plan.preserve.clone(),
            platform: "windows-preflight-test".to_owned(),
            evidence: None,
        },
        OpenpyxlWorkerConfig::new(&dummy_python, &worker_script)?,
    )
}

#[test]
fn payload_preflight_rejects_without_allowance_and_passes_with_it() -> Result<()> {
    // Without the allowance the payload on 汇总!B1 is rejected before any
    // write, exactly as before the exemption existed.
    let strict_dir = TestDir::new("strict");
    let strict_plan = plan_for(&strict_dir, None)?;
    let strict_failure = backend(&strict_dir, &strict_plan)?
        .apply(&strict_plan)
        .expect_err("merged payload without allowance must be rejected");
    assert_eq!(
        strict_failure.kind,
        excel_editor::BackendErrorKind::UnsupportedBeforeWrite
    );
    assert!(
        strict_failure
            .message
            .contains("merged non-anchor cell 汇总!B1"),
        "unexpected rejection: {}",
        strict_failure.message
    );
    assert!(!strict_plan.output.exists());

    // With the allowance and the payload outside every operation rectangle,
    // the payload preflight passes; the apply then fails at the dummy worker
    // startup instead of at the payload guard.
    let exempt_dir = TestDir::new("exempt");
    let exempt_plan = plan_for(&exempt_dir, Some("汇总"))?;
    let exempt_failure = backend(&exempt_dir, &exempt_plan)?
        .apply(&exempt_plan)
        .expect_err("dummy worker startup must fail");
    assert!(
        !exempt_failure.message.contains("merged non-anchor"),
        "payload preflight should have passed, got: {}",
        exempt_failure.message
    );
    assert!(!exempt_plan.output.exists());
    Ok(())
}
