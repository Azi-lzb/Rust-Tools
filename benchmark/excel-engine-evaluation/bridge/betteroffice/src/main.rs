use std::time::Instant;
use xlsx_calc::{rebuild_and_recalc_all, recalc_after};
use xlsx_model::{Cell, CellRef, CellValue, DefinedName, Sheet, SheetId, Workbook};

fn cell(value: CellValue, formula: Option<&str>) -> Cell {
    Cell { value, formula: formula.map(str::to_owned), style: None }
}
fn n(v: f64) -> CellValue { CellValue::Number { value: v } }
fn s(v: &str) -> CellValue { CellValue::Text { value: v.to_owned() } }
fn put(wb: &mut Workbook, sheet: SheetId, addr: &str, value: CellValue, formula: Option<&str>) {
    wb.sheet_mut(sheet).unwrap().set_cell(CellRef::parse_a1(addr).unwrap(), cell(value, formula));
}
fn shown(value: &CellValue) -> String {
    match value {
        CellValue::Empty => "<empty>".into(),
        CellValue::Number { value } => format!("{value:.15}"),
        CellValue::Text { value } => value.clone(),
        CellValue::Bool { value } => value.to_string(),
        CellValue::Error { value } => value.as_str().to_owned(),
    }
}
fn csv(value: &str) -> String { format!("\"{}\"", value.replace('"', "\"\"")) }
fn main() {
    let mut wb = Workbook::default();
    wb.sheets.push(Sheet::new("Data"));
    wb.sheets.push(Sheet::new("Calc"));
    wb.defined_names.push(DefinedName { name: "InputValues".into(), formula: "Data!$B$2:$B$4".into(), local_sheet: None, hidden: false });
    let data = SheetId(0); let calc = SheetId(1);
    put(&mut wb, data, "A1", s("Key"), None); put(&mut wb, data, "B1", s("Value"), None);
    put(&mut wb, data, "A2", s("a"), None); put(&mut wb, data, "B2", n(2.0), None);
    put(&mut wb, data, "A3", s("b"), None); put(&mut wb, data, "B3", n(3.0), None);
    put(&mut wb, data, "A4", s("b"), None); put(&mut wb, data, "B4", n(4.0), None);
    let formulas = [
        "SUM(InputValues)", "SUMIF(Data!A2:A4,\"b\",Data!B2:B4)", "COUNTIF(Data!B2:B4,\">2\")",
        "IF(AND(Data!B2>0,Data!B3>0),ROUND(Data!B2*Data!B3,2),0)",
        "IF(OR(Data!B2=0,Data!B3=3),1,0)", "VLOOKUP(\"b\",Data!A2:B4,2,FALSE)",
        "INDEX(Data!B2:B4,MATCH(\"b\",Data!A2:A4,0))", "DATE(2026,1,2)",
        "SUM(Data!B2:B4)+Data!B2", "1/0",
    ];
    for (i, formula) in formulas.iter().enumerate() {
        let addr = format!("A{}", i + 1);
        put(&mut wb, calc, &addr, CellValue::Empty, Some(formula));
    }

    let start = Instant::now();
    let (mut graph, first) = rebuild_and_recalc_all(&mut wb, None);
    let first_ms = start.elapsed().as_secs_f64() * 1000.0;
    println!("phase,engine,formula_count,changed,budget_exhausted,elapsed_ms,address,formula,result");
    for i in 0..formulas.len() {
        let addr = format!("A{}", i + 1); let at = CellRef::parse_a1(&addr).unwrap();
        let c = wb.sheet(calc).unwrap().cell(at).unwrap();
        println!("initial,betteroffice,{},{},{},{:.3},{},{},{}", formulas.len(), first.changed.len(), first.budget_exhausted, first_ms,
            addr, csv(c.formula.as_deref().unwrap_or("")), csv(&shown(&c.value)));
    }

    let input = CellRef::parse_a1("B2").unwrap();
    put(&mut wb, data, "B2", n(5.0), None);
    let start = Instant::now();
    let edited = recalc_after(&mut wb, &mut graph, &[(data, input)], None);
    let edit_ms = start.elapsed().as_secs_f64() * 1000.0;
    for i in 0..formulas.len() {
        let addr = format!("A{}", i + 1); let at = CellRef::parse_a1(&addr).unwrap();
        let c = wb.sheet(calc).unwrap().cell(at).unwrap();
        println!("edited_input,betteroffice,{},{},{},{:.3},{},{},{}", formulas.len(), edited.changed.len(), edited.budget_exhausted, edit_ms,
            addr, csv(c.formula.as_deref().unwrap_or("")), csv(&shown(&c.value)));
    }
    put(&mut wb, calc, "A1", CellValue::Empty, Some("SUM(InputValues)*2"));
    let start = Instant::now();
    let (_new_graph, formula_edit) = rebuild_and_recalc_all(&mut wb, None);
    let formula_edit_ms = start.elapsed().as_secs_f64() * 1000.0;
    let c = wb.sheet(calc).unwrap().cell(CellRef::parse_a1("A1").unwrap()).unwrap();
    println!("edited_formula,betteroffice,{},{},{},{:.3},A1,{},{}", formulas.len(), formula_edit.changed.len(), formula_edit.budget_exhausted,
        formula_edit_ms, csv(c.formula.as_deref().unwrap_or("")), csv(&shown(&c.value)));
}
