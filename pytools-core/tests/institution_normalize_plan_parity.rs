use anyhow::{Context, Result};
use excel_editor::{CellValue, EditOperation, WorkbookFeature, WorkbookFormat, WorkbookProfile};
use excel_reader::{read_workbook, SheetDump};
use pytools_core::institution_normalize::build_institution_normalize_plan;
use serde_json::{json, Value};
use std::{
    collections::BTreeSet,
    fs,
    path::{Path, PathBuf},
    time::{SystemTime, UNIX_EPOCH},
};

#[test]
fn edit_plan_matches_python_baseline_without_modifying_source() -> Result<()> {
    let root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let fixture = root.join("tests/fixtures/institution_normalize.xlsx");
    let baseline_path = root.join("probe/reports/institution_normalize_baseline.json");
    let baseline: Value = serde_json::from_slice(&fs::read(&baseline_path)?)?;
    let mapping = baseline["mapping"]
        .as_array()
        .context("baseline mapping must be an array")?
        .iter()
        .map(|entry| {
            Ok((
                entry[0]
                    .as_str()
                    .context("mapping key must be a string")?
                    .to_owned(),
                entry[1]
                    .as_str()
                    .context("mapping value must be a string")?
                    .to_owned(),
            ))
        })
        .collect::<Result<Vec<_>>>()?;

    let unique = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
    let output = std::env::temp_dir().join(format!(
        "pytools-institution-normalize-plan-{}-{unique}.xlsx",
        std::process::id()
    ));
    let original_bytes = fs::read(&fixture)?;
    let profile = WorkbookProfile {
        format: WorkbookFormat::Xlsx,
        // The plan builder does not perform a complete package inventory.
        // Keep this false so the router cannot treat this test profile as safe.
        inventory_complete: false,
        features_present: BTreeSet::from([
            WorkbookFeature::Comments,
            WorkbookFeature::MergedCells,
            WorkbookFeature::Formulas,
        ]),
    };

    let result = build_institution_normalize_plan(
        &fixture,
        &output,
        profile,
        "institution-normalize-plan-parity",
        &mapping,
    )?;
    let source_dump = read_workbook(&fixture)?;
    let expected_operations = expected_operations(&baseline, &source_dump.sheets)?;
    let plan = result
        .plan
        .as_ref()
        .context("Python baseline predicts workbook changes, but plan is absent")?;

    assert_eq!(
        result.branch_sheets,
        baseline["stats"]["branch_sheets"].as_u64().unwrap() as usize
    );
    assert_eq!(
        result.mapped_sheets,
        baseline["stats"]["mapped_sheets"].as_u64().unwrap() as usize
    );
    assert_eq!(
        result.mapped_rows,
        baseline["stats"]["mapped_rows"].as_u64().unwrap() as usize
    );
    assert_eq!(plan.operations, expected_operations);
    assert!(!plan.workbook.inventory_complete);
    assert_eq!(fs::read(&fixture)?, original_bytes);
    assert!(!output.exists());

    if let Some(report_path) = std::env::var_os("PYTOOLS_INSTITUTION_PLAN_PARITY_REPORT") {
        let report_path = PathBuf::from(report_path);
        if let Some(parent) = report_path.parent() {
            fs::create_dir_all(parent)?;
        }
        let report = json!({
            "schema_version": 1,
            "suite": "pytools.institution_normalize_plan_parity",
            "status": "passed",
            "baseline": "probe/reports/institution_normalize_baseline.json",
            "fixture": "tests/fixtures/institution_normalize.xlsx (synthetic)",
            "stats": {
                "branch_sheets": result.branch_sheets,
                "mapped_sheets": result.mapped_sheets,
                "mapped_rows": result.mapped_rows,
                "plan_operations": plan.operations.len()
            },
            "plan": {
                "request_id": plan.request_id,
                "format": plan.format,
                "operations": plan.operations,
                "preserve": plan.preserve,
                "workbook": plan.workbook
            },
            "source_unchanged": true,
            "output_created_by_planning": false,
            "routing_status": "not_eligible_without_complete_workbook_inventory",
            "limitations": [
                "synthetic XLSX fixture only",
                "this test verifies planning parity, not engine execution or output preservation",
                "the existing direct umya candidate remains until a complete inventory, routed backend, and validator are accepted"
            ]
        });
        fs::write(report_path, serde_json::to_vec_pretty(&report)?)?;
    }

    Ok(())
}

fn expected_operations(
    baseline: &Value,
    source_sheets: &[SheetDump],
) -> Result<Vec<EditOperation>> {
    let mut operations = Vec::new();
    for expected_sheet in baseline["sheets"]
        .as_array()
        .context("baseline sheets must be an array")?
    {
        let sheet_name = expected_sheet["name"]
            .as_str()
            .context("baseline sheet name must be a string")?;
        if !sheet_name.contains("分机构") {
            continue;
        }
        let source_sheet = source_sheets
            .iter()
            .find(|sheet| sheet.name == sheet_name)
            .with_context(|| format!("source worksheet {sheet_name} is missing"))?;
        let mut cells = expected_sheet["values"]
            .as_object()
            .context("baseline values must be an object")?
            .iter()
            .filter_map(|(address, expected)| {
                target_cell_order(address).map(|order| (order, address, expected))
            })
            .collect::<Vec<_>>();
        cells.sort_by_key(|(order, _, _)| *order);

        for (_, address, expected) in cells {
            let previous = source_sheet.values.get(address);
            if previous == Some(expected) {
                continue;
            }
            let value = expected["value"].as_str().with_context(|| {
                format!("expected normalized cell {sheet_name}!{address} is not text")
            })?;
            operations.push(EditOperation::SetCellValue {
                sheet: sheet_name.to_owned(),
                cell: address.to_owned(),
                value: CellValue::Text(value.to_owned()),
            });
        }
    }
    Ok(operations)
}

fn target_cell_order(address: &str) -> Option<(usize, usize)> {
    let index = address.find(|character: char| character.is_ascii_digit())?;
    let column = match &address[..index] {
        "A" => 0,
        "B" => 1,
        _ => return None,
    };
    let row = address[index..].parse::<usize>().ok()?;
    (row > 0).then_some((row, column))
}
