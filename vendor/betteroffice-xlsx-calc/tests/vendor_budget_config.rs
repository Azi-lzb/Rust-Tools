//! vendor tests (see the workspace `vendor/VENDOR_NOTES.md`).
//!
//! This file runs in its own process on purpose: it mutates the process-global
//! recalculation budget ceiling, and the lib tests must not observe that.

use xlsx_calc::engine::rebuild_and_recalc_all;
use xlsx_calc::eval::{max_recalculation_visits, set_max_recalculation_visits};
use xlsx_model::{Cell, CellProvider, CellRef, CellValue, Sheet, SheetId, Workbook};

fn put_num(wb: &mut Workbook, sheet: SheetId, cell: &str, v: f64) {
    wb.sheet_mut(sheet).unwrap().set_cell(
        CellRef::parse_a1(cell).unwrap(),
        Cell {
            value: CellValue::Number { value: v },
            ..Cell::default()
        },
    );
}

fn put_formula(wb: &mut Workbook, sheet: SheetId, cell: &str, f: &str) {
    wb.sheet_mut(sheet).unwrap().set_cell(
        CellRef::parse_a1(cell).unwrap(),
        Cell {
            value: CellValue::Empty,
            formula: Some(f.to_string()),
            style: None,
        },
    );
}

#[test]
fn tiny_budget_is_reported_and_restored_budget_calculates() {
    let previous = max_recalculation_visits();
    set_max_recalculation_visits(10);

    let mut wb = Workbook::default();
    wb.sheets.push(Sheet::new("Data"));
    wb.sheets.push(Sheet::new("Calc"));
    let data = SheetId(0);
    let calc = SheetId(1);
    for row in 1..=2_000u32 {
        put_num(&mut wb, data, &format!("A{row}"), row as f64);
    }
    put_formula(&mut wb, calc, "A1", "SUM(Data!A:A)");

    let (_, result) = rebuild_and_recalc_all(&mut wb, None);
    assert!(result.budget_exhausted, "tiny budget must be reported");

    set_max_recalculation_visits(previous);
    let (_, result2) = rebuild_and_recalc_all(&mut wb, None);
    assert!(!result2.budget_exhausted);
    assert!(
        matches!(
            wb.value(calc, CellRef::parse_a1("A1").unwrap()),
            CellValue::Number { .. }
        ),
        "restored ceiling must let the formula calculate"
    );
}
