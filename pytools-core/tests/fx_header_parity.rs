use excel_editor::{
    BackendCapability, CapabilityEvidence, CapabilityStatus, EngineId, ExcelModifier,
    OoxmlHeaderTextBackend, OoxmlHeaderTextValidator, WorkbookFeature, WorkbookFormat,
};
use excel_reader::{read_header_inventory, HeaderVariant};
use ooxml_patch::validate_header_text_patch;
use pytools_core::fx_header::build_fx_header_plan;
use serde_json::Value;
use std::{
    collections::BTreeSet,
    fs,
    path::{Path, PathBuf},
    sync::Arc,
};

const PLATFORM: &str = "ubuntu-wsl2-x86_64";

fn temp_dir() -> PathBuf {
    let nonce = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .expect("clock after epoch")
        .as_nanos();
    let path = std::env::temp_dir().join(format!(
        "pytools_fx_header_{}_{}",
        std::process::id(),
        nonce
    ));
    fs::create_dir_all(&path).expect("create temporary test directory");
    path
}

fn baseline() -> Value {
    serde_json::from_str(include_str!(
        "../probe/reports/fx_header_python_baseline.json"
    ))
    .expect("parse saved Python baseline")
}

fn header_capability(
    preserve: BTreeSet<excel_editor::PreservationRequirement>,
) -> BackendCapability {
    BackendCapability {
        engine: EngineId::OoxmlPatch,
        status: CapabilityStatus::Passed,
        format: WorkbookFormat::Xlsx,
        operations: BTreeSet::from([excel_editor::EditOperationKind::ReplaceHeaderText]),
        supported_features: BTreeSet::from([
            WorkbookFeature::HeadersFooters,
            WorkbookFeature::Comments,
            WorkbookFeature::ConditionalFormatting,
            WorkbookFeature::DataValidations,
            WorkbookFeature::MergedCells,
            WorkbookFeature::Formulas,
        ]),
        preserved: preserve,
        platform: PLATFORM.to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: "fx-header-openpyxl-synthetic-v1".to_owned(),
            report_path: "pytools-core/probe/reports/fx_header_parity.json".to_owned(),
            adapter_version: "excel-editor/0.1.0 + ooxml-patch/0.1.0".to_owned(),
            engine_version: "ooxml-patch/0.1.0 + zip/6".to_owned(),
            platform: PLATFORM.to_owned(),
        }),
    }
}

fn header_string(value: &Value, variant: &str) -> String {
    let mut output = String::new();
    for (section, marker) in [("left", "&L"), ("center", "&C"), ("right", "&R")] {
        if let Some(text) = value[variant][section].as_str() {
            output.push_str(marker);
            output.push_str(text);
        }
    }
    output
}

#[test]
fn xlsx_header_fix_matches_python_and_preserves_unmodified_workbook_content() {
    let baseline = baseline();
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/fx_header/source.xlsx");
    assert!(
        fixture.is_file(),
        "run the read-only Python baseline probe first"
    );
    let directory = temp_dir();
    let output = directory.join("fixed.xlsx");
    let profile = excel_editor::WorkbookProfile {
        format: WorkbookFormat::Xlsx,
        inventory_complete: true,
        features_present: BTreeSet::from([
            WorkbookFeature::Comments,
            WorkbookFeature::ConditionalFormatting,
            WorkbookFeature::DataValidations,
            WorkbookFeature::MergedCells,
            WorkbookFeature::Formulas,
        ]),
    };
    let inspection = build_fx_header_plan(&fixture, &output, profile, "fx-header-python-parity")
        .expect("build the engine-independent edit plan");

    assert_eq!(inspection.fx_sheets, 2);
    assert_eq!(inspection.modified_sheets, 1);
    assert_eq!(
        inspection.fx_sheets,
        baseline["python_baseline_stats"]["fx_sheets"]
    );
    assert_eq!(
        inspection.modified_sheets,
        baseline["python_baseline_stats"]["modified_sheets"]
    );
    let plan = inspection
        .plan
        .expect("matching header text produces a plan");
    assert_eq!(plan.operations.len(), 1);

    let capability = header_capability(plan.preserve.clone());
    let modifier = ExcelModifier::new(
        PLATFORM,
        vec![Arc::new(
            OoxmlHeaderTextBackend::new(capability).expect("construct OOXML header backend"),
        )],
        Arc::new(OoxmlHeaderTextValidator),
    )
    .expect("construct ExcelModifier");
    let report = modifier
        .execute(&plan)
        .expect("execute validated edit plan");
    assert_eq!(report.engine, EngineId::OoxmlPatch);
    assert!(report.route_decisions[0].eligible);
    assert!(report
        .attempts
        .iter()
        .any(|attempt| { attempt.outcome == excel_editor::BackendAttemptOutcome::Succeeded }));
    validate_header_text_patch(
        &fixture,
        &output,
        &[ooxml_patch::HeaderTextUpdate {
            sheet_name: "外汇报表".to_owned(),
            search: "万元".to_owned(),
            replacement: "万美元".to_owned(),
        }],
    )
    .expect("only requested header text changed");

    let after = read_header_inventory(&output).expect("read patched header inventory");
    let actual: BTreeSet<_> = after
        .iter()
        .flat_map(|sheet| {
            sheet.headers.iter().map(|header| {
                (
                    sheet.sheet_name.as_str(),
                    header.variant,
                    header.text.as_str(),
                )
            })
        })
        .collect();
    for variant in ["odd", "even", "first"] {
        let expected = header_string(&baseline["headers_after"]["外汇报表"], variant);
        if !expected.is_empty() {
            let raw = match variant {
                "odd" => HeaderVariant::Odd,
                "even" => HeaderVariant::Even,
                _ => HeaderVariant::First,
            };
            assert!(actual.contains(&("外汇报表", raw, expected.as_str())));
        }
    }
    assert!(actual.iter().any(|(sheet, _, text)| {
        *sheet == "外汇空白" && text.contains("单位：万美元")
    }));
    assert!(actual.iter().any(|(sheet, _, text)| {
        *sheet == "普通报表" && text.contains("普通报表：万元")
    }));

    let dump = excel_reader::read_workbook(&output).expect("Calamine can reopen output");
    let fx = dump
        .sheets
        .iter()
        .find(|sheet| sheet.name == "外汇报表")
        .expect("FX worksheet remains");
    assert_eq!(fx.merges, vec!["C3:D3"]);
    assert_eq!(
        fx.comments.get("A1").map(String::as_str),
        Some("原有批注保留")
    );
    assert_eq!(fx.conditional_formatting.len(), 1);
    assert_eq!(fx.data_validations.len(), 1);
    assert_eq!(
        fx.values["B2"]["value"],
        baseline["preserved_workbook_properties"]["cell_b2"]
    );
    assert_eq!(
        fx.formulas["C2"],
        baseline["preserved_workbook_properties"]["formula_c2"]
            .as_str()
            .unwrap()
            .trim_start_matches('=')
    );

    if let Some(report_path) = std::env::var_os("PYTOOLS_FX_HEADER_PARITY_REPORT") {
        let report_path = PathBuf::from(report_path);
        if let Some(parent) = report_path.parent() {
            fs::create_dir_all(parent).expect("create parity report directory");
        }
        let parity_report = serde_json::json!({
            "schema_version": 1,
            "status": "passed",
            "feature": "pytools.feature_1_4_fx_header.ordinary_unit_replacement",
            "scope": {
                "input_format": "xlsx",
                "sample_kind": "synthetic",
                "python_baseline_variant": "PyCode ordinary non-COM path",
                "platform": PLATFORM,
                "fx_sheets": inspection.fx_sheets,
                "modified_sheets": inspection.modified_sheets,
                "edit_operations": plan.operations.len()
            },
            "python_baseline": {
                "source_file": baseline["source_file"],
                "source_sha256": baseline["source_sha256"],
                "fixture": baseline["fixture"],
                "fixture_sha256": baseline["fixture_sha256"],
                "stats": baseline["python_baseline_stats"]
            },
            "rust_candidate": {
                "plan_builder": "pytools_core::fx_header::build_fx_header_plan",
                "metadata_reader": "excel_reader::read_header_inventory",
                "modifier": "excel_editor::ExcelModifier",
                "backend": "ooxml-patch/0.1.0",
                "backend_status": "passed for this synthetic XLSX operation and preservation contract"
            },
            "checks": {
                "fx_and_modified_sheet_counts_match_python": true,
                "odd_even_first_header_text_matches_python": true,
                "non_fx_header_remains_unchanged": true,
                "requested_patch_scope_only": true,
                "values_formulas_merges_comments_conditional_formatting_and_data_validations_preserved": true,
                "output_reopens_with_calamine": true
            },
            "not_verified": [
                "real business workbooks",
                "XLSM/VBA preservation",
                "Windows native Rust execution",
                "Kylin/UOS glibc 2.28 packaging",
                "menu or CLI integration",
                "in-place source replacement"
            ]
        });
        fs::write(
            &report_path,
            serde_json::to_vec_pretty(&parity_report).expect("serialize parity report"),
        )
        .expect("write parity report");
    }

    fs::remove_dir_all(directory).expect("remove temporary test directory");
}

#[test]
fn files_without_matching_headers_produce_no_edit_plan() {
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/fx_header/source.xlsx");
    let directory = temp_dir();
    let output = directory.join("unchanged.xlsx");
    let prepatched = directory.join("prepatched.xlsx");
    ooxml_patch::replace_header_text(
        &fixture,
        &prepatched,
        &[ooxml_patch::HeaderTextUpdate {
            sheet_name: "外汇报表".to_owned(),
            search: "万元".to_owned(),
            replacement: "万美元".to_owned(),
        }],
    )
    .expect("prepare a workbook with the target header already fixed");
    let profile = excel_editor::WorkbookProfile {
        format: WorkbookFormat::Xlsx,
        inventory_complete: true,
        features_present: BTreeSet::new(),
    };
    let inspection = build_fx_header_plan(&prepatched, &output, profile, "fx-header-noop")
        .expect("build no-op inspection");
    assert_eq!(inspection.fx_sheets, 2);
    assert_eq!(inspection.modified_sheets, 0);
    assert!(inspection.plan.is_none());
    fs::remove_dir_all(directory).expect("remove temporary test directory");
}
