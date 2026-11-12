//! Phase 2 whole-plan safety requirements, independent of workbook I/O.
use excel_editor::{
    EditOperation, EngineId, ExcelEditPlan, PageMargins, WorkbookFormat, WorkbookProfile,
};
use pytools_core::excel_edit_router::{route_plan, RouteConfig, RouteMode};
use std::collections::BTreeSet;

fn plan(operations: Vec<EditOperation>) -> ExcelEditPlan {
    ExcelEditPlan {
        request_id: "phase2-whole-plan".into(),
        input: "source.xlsx".into(),
        output: "candidate.xlsx".into(),
        format: WorkbookFormat::Xlsx,
        workbook: WorkbookProfile {
            format: WorkbookFormat::Xlsx,
            inventory_complete: true,
            features_present: BTreeSet::new(),
        },
        operations,
        preserve: BTreeSet::new(),
        drop_allowed: Default::default(),
    }
}

#[test]
fn auto_selects_one_engine_for_margins_and_row_insertion() {
    let p = plan(vec![
        EditOperation::SetMargins {
            sheet: "Data".into(),
            margins: PageMargins {
                left: Some(0.5),
                ..Default::default()
            },
        },
        EditOperation::InsertRows {
            sheet: "Data".into(),
            first_row: 2,
            count: 1,
        },
    ]);
    let route = route_plan(&p, &RouteConfig::new(RouteMode::Auto)).unwrap();
    assert_eq!(
        route.engine,
        EngineId::OpenpyxlWorker,
        "no single Rust adapter supports this complete business plan"
    );
}

#[test]
fn auto_checks_entire_style_and_margin_plan_before_choosing_umya() {
    let p = plan(vec![
        EditOperation::SetCellFillColor {
            sheet: "Data".into(),
            cell: "A1".into(),
            argb: "FFFF0000".into(),
        },
        EditOperation::SetMargins {
            sheet: "Data".into(),
            margins: PageMargins {
                left: Some(0.5),
                ..Default::default()
            },
        },
    ]);
    assert_eq!(
        route_plan(&p, &RouteConfig::new(RouteMode::Auto))
            .unwrap()
            .engine,
        EngineId::OpenpyxlWorker,
        "UMYA explicitly rejects SetMargins"
    );
}

#[test]
fn auto_routes_structural_and_sheet_rename_mix_without_preflight_error() {
    let p = plan(vec![
        EditOperation::DeleteRows {
            sheet: "Data".into(),
            first_row: 2,
            count: 1,
        },
        EditOperation::RenameSheet {
            old_name: "Data".into(),
            new_name: "Renamed".into(),
        },
    ]);
    assert_eq!(
        route_plan(&p, &RouteConfig::new(RouteMode::Auto))
            .unwrap()
            .engine,
        EngineId::OpenpyxlWorker
    );
}
