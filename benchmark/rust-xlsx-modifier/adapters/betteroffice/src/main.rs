use anyhow::{bail, Context, Result};
use betteroffice_xlsx::{
    CalculationOptions, CellRange, CellRef, CellState, CellValue, ColId, Op, RowId, SheetId,
    StylePatch, Workbook,
};
use rust_xlsxwriter::{
    Chart, ChartType, ConditionalFormatCell, ConditionalFormatCellRule, Format,
    Workbook as FixtureWorkbook,
};
use serde_json::json;
use std::{env, fs, path::Path, time::Instant};
use zip::{write::SimpleFileOptions, ZipArchive, ZipWriter};

fn main() -> Result<()> {
    let mut args = env::args().skip(1);
    match args.next().as_deref() {
        Some("fixture") => fixture(args.next().context("fixture path required")?.as_str()),
        Some("audit") => audit(args.next().context("output directory required")?.as_str()),
        Some("edit") => edit(args.next().context("input xlsx required")?.as_str(), args.next().context("output xlsx required")?.as_str()),
        Some("gen100k") => gen100k(args.next().context("output directory required")?.as_str()),
        Some("perf") => perf(args.next().context("output directory required")?.as_str()),
        _ => bail!("usage: betteroffice-modification-audit <fixture|audit|edit|gen100k|perf> <directory-or-path>"),
    }
}

fn edit(input: &str, output: &str) -> Result<()> {
    let source = fs::read(input)?;
    let mut wb = Workbook::open(&source)?;
    wb.apply_ops(
        vec![Op::SetCell {
            sheet: SheetId(0),
            at: CellRef::new(0, 4),
            cell: CellState {
                value: CellValue::Text {
                    value: "ordinary edit".into(),
                },
                formula: None,
                style: None,
            },
        }],
        CalculationOptions::default(),
    )?;
    fs::write(output, wb.save()?)?;
    Ok(())
}

fn fixture(path: &str) -> Result<()> {
    let mut wb = FixtureWorkbook::new();
    let ws = wb.add_worksheet();
    ws.set_name("Data")?;
    ws.write_number(0, 0, 10)?;
    ws.write_number(0, 1, 20)?;
    ws.write_formula(0, 2, "=A1+B1")?;
    ws.write_formula(0, 3, "=A1*2")?;
    ws.write_number(1, 0, 1)?;
    ws.write_number(1, 1, 2)?;
    ws.write_number(2, 0, 3)?;
    ws.write_number(2, 1, 4)?;
    ws.merge_range(4, 0, 4, 2, "merged marker", &Format::new().set_bold())?;
    let cf = ConditionalFormatCell::new()
        .set_rule(ConditionalFormatCellRule::GreaterThan(0))
        .set_format(Format::new().set_background_color("C6EFCE"));
    ws.add_conditional_format(0, 0, 2, 1, &cf)?;
    let mut chart = Chart::new(ChartType::Column);
    chart.add_series().set_values("Data!$A$1:$A$3");
    ws.insert_chart(6, 0, &chart)?;
    ws.set_landscape().set_paper_size(9);
    ws.set_print_area(0, 0, 4, 2)?;
    ws.set_margins(0.3, 0.3, 0.5, 0.5, 0.2, 0.2);
    ws.set_header("&LHeader fixture").set_footer("&CPage &P");
    wb.define_name("DataRange", "=Data!$A$1:$B$3")?;
    wb.save(path)?;

    // Add an opaque extension part. The content type is explicit so the fixture
    // contains a valid OPC part without relying on a vendor-specific feature.
    let bytes = fs::read(path)?;
    let with_custom = inject_custom_xml(
        &bytes,
        b"<extension marker=\"betteroffice-audit\">opaque-1</extension>",
    )?;
    fs::write(path, with_custom)?;
    Ok(())
}

fn inject_custom_xml(bytes: &[u8], payload: &[u8]) -> Result<Vec<u8>> {
    let mut input = ZipArchive::new(std::io::Cursor::new(bytes))?;
    let mut entries = Vec::new();
    for i in 0..input.len() {
        let mut file = input.by_index(i)?;
        let name = file.name().to_owned();
        let mut data = Vec::new();
        std::io::Read::read_to_end(&mut file, &mut data)?;
        if name == "[Content_Types].xml" {
            let s = String::from_utf8(data)?;
            let insert = "<Default Extension=\"xml\" ContentType=\"application/xml\"/>";
            let changed = if s.contains("Extension=\"xml\"") {
                s
            } else {
                s.replacen("</Types>", &format!("{insert}</Types>"), 1)
            };
            data = changed.into_bytes();
        }
        entries.push((name, data));
    }
    let mut out = ZipWriter::new(std::io::Cursor::new(Vec::new()));
    let opts = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    for (name, data) in entries {
        out.start_file(name, opts)?;
        std::io::Write::write_all(&mut out, &data)?;
    }
    out.start_file("customXml/item1.xml", opts)?;
    std::io::Write::write_all(&mut out, payload)?;
    Ok(out.finish()?.into_inner())
}

fn audit(dir: &str) -> Result<()> {
    fs::create_dir_all(dir)?;
    let source = format!("{dir}/audit_fixture.xlsx");
    fixture(&source)?;
    let original = fs::read(&source)?;
    let mut wb = Workbook::open(&original)?;
    let sheet = SheetId(0);
    let started = Instant::now();
    wb.apply_ops(
        vec![
            Op::DeleteRows {
                sheet,
                at: RowId::from(1u32),
                count: 1,
            },
            Op::InsertRows {
                sheet,
                at: RowId::from(1u32),
                count: 2,
            },
            Op::SetColWidth {
                sheet,
                col: 4,
                width: Some(21.0),
            },
            Op::SetRowHeight {
                sheet,
                row: 1,
                height: Some(27.0),
            },
            Op::PatchRangeStyle {
                sheet,
                range: CellRange::new(CellRef::new(0, 0), CellRef::new(0, 0)),
                patch: StylePatch {
                    bold: Some(true),
                    ..StylePatch::default()
                },
            },
        ],
        CalculationOptions::default(),
    )?;
    let structural_ms = started.elapsed().as_secs_f64() * 1000.0;
    let structural = wb.save()?;
    fs::write(format!("{dir}/structure.xlsx"), &structural)?;
    let reopened = Workbook::open(&structural)?;
    let model = reopened.model();
    let sheet_model = &model.sheets[0];
    let structure_json = json!({
        "chart_count": sheet_model.charts.len(),
        "conditional_format_count": sheet_model.conditional_formats.len(),
        "merge_count": sheet_model.merges.len(),
        "a1_has_style": sheet_model.cell(CellRef::new(0,0)).and_then(|c| c.style).is_some(),
        "column_width_e": sheet_model.col_widths.get(&4),
        "row_height_2": sheet_model.row_heights.get(&1),
        "defined_names": model.defined_names.iter().map(|n| json!({"name": n.name, "formula": n.formula})).collect::<Vec<_>>(),
        "structural_ms": structural_ms,
        "custom_xml_preserved": zip_part(&structural, "customXml/item1.xml")?.as_deref() == Some(b"<extension marker=\"betteroffice-audit\">opaque-1</extension>"),
    });

    let mut print_probe = Workbook::open(&original)?;
    print_probe.apply_ops(
        vec![Op::SetDefinedName {
            name: "_xlnm.Print_Area".into(),
            refers_to: "Data!$A$1:$C$4".into(),
            scope: Some(sheet),
            hidden: Some(true),
        }],
        CalculationOptions::default(),
    )?;
    let print_probe_bytes = print_probe.save()?;
    fs::write(format!("{dir}/print_area_probe.xlsx"), &print_probe_bytes)?;

    let formula_path = format!("{dir}/formula.xlsx");
    let mut formula_wb = Workbook::open(&original)?;
    formula_wb.apply_ops(
        vec![
            Op::SetCell {
                sheet,
                at: CellRef::new(0, 0),
                cell: CellState {
                    value: CellValue::Number { value: 15.0 },
                    formula: None,
                    style: None,
                },
            },
            Op::DeleteCols {
                sheet,
                at: ColId::from(1u32),
                count: 1,
            },
        ],
        CalculationOptions::default(),
    )?;
    let formula_bytes = formula_wb.save()?;
    fs::write(&formula_path, &formula_bytes)?;
    let after = Workbook::open(&formula_bytes)?;
    let s = &after.model().sheets[0];
    let formula = s
        .cell(CellRef::new(0, 1))
        .map(|c| json!({"formula": c.formula, "cached_value": c.value}));
    let formula_a_only = s
        .cell(CellRef::new(0, 2))
        .map(|c| json!({"formula": c.formula, "cached_value": c.value}));
    let formula_json = json!({"deleted_reference_formula": formula, "a1_dependent_formula_after_delete_b": formula_a_only, "calc_chain_part": zip_part(&formula_bytes, "xl/calcChain.xml")?.is_some()});
    fs::write(
        format!("{dir}/audit.json"),
        serde_json::to_vec_pretty(
            &json!({"structure": structure_json, "formula": formula_json, "print_area_probe": {"defined_names": print_probe.model().defined_names.iter().map(|n| json!({"name": n.name, "formula": n.formula, "scope": n.local_sheet.map(|s| s.0)})).collect::<Vec<_>>()}}),
        )?,
    )?;
    println!("{}", fs::read_to_string(format!("{dir}/audit.json"))?);
    Ok(())
}

fn gen100k(dir: &str) -> Result<()> {
    fs::create_dir_all(dir)?;
    let path = format!("{dir}/100k.xlsx");
    let mut x = FixtureWorkbook::new();
    let ws = x.add_worksheet();
    ws.set_name("Data")?;
    for r in 0..100_000u32 {
        ws.write_number(r, 0, r as f64)?;
        ws.write_number(r, 1, (r % 97) as f64)?;
        ws.write_string(r, 2, format!("row-{r}"))?;
    }
    x.save(&path)?;
    println!("generated {path} (100000 rows x 3 columns)");
    Ok(())
}

fn perf(dir: &str) -> Result<()> {
    fs::create_dir_all(dir)?;
    let path = format!("{dir}/100k.xlsx");
    if !Path::new(&path).exists() {
        bail!("run gen100k first to create {path}");
    }
    let bytes = fs::read(&path)?;
    let mut results = Vec::new();
    let workloads = [
        (
            "edit_1000",
            (0..1000)
                .map(|i| Op::SetCell {
                    sheet: SheetId(0),
                    at: CellRef::new(10 + i / 10, i % 10),
                    cell: CellState {
                        value: CellValue::Number { value: i as f64 },
                        formula: None,
                        style: None,
                    },
                })
                .collect::<Vec<_>>(),
        ),
        (
            "insert_delete_rows_cols",
            vec![
                Op::InsertRows {
                    sheet: SheetId(0),
                    at: 50_000,
                    count: 3,
                },
                Op::DeleteRows {
                    sheet: SheetId(0),
                    at: 100,
                    count: 2,
                },
                Op::InsertCols {
                    sheet: SheetId(0),
                    at: 1,
                    count: 1,
                },
                Op::DeleteCols {
                    sheet: SheetId(0),
                    at: 4,
                    count: 1,
                },
            ],
        ),
    ];
    // One unreported open/save warm-up; every measured repetition reopens the
    // same immutable source bytes and writes to a distinct output file.
    let warm = Workbook::open(&bytes)?;
    let _ = warm.save()?;
    for (name, ops) in workloads {
        for rep in 1..=3 {
            let t0 = Instant::now();
            let mut wb = Workbook::open(&bytes)?;
            let open_ms = t0.elapsed().as_secs_f64() * 1000.0;
            let t1 = Instant::now();
            wb.apply_ops(ops.clone(), CalculationOptions::default())?;
            let operation_ms = t1.elapsed().as_secs_f64() * 1000.0;
            let t2 = Instant::now();
            let output = wb.save()?;
            let save_ms = t2.elapsed().as_secs_f64() * 1000.0;
            fs::write(format!("{dir}/{name}-r{rep}.xlsx"), &output)?;
            results.push(json!({"scenario":name,"rep":rep,"open_ms":open_ms,"operation_ms":operation_ms,"save_ms":save_ms,"total_ms":open_ms+operation_ms+save_ms,"input_bytes":bytes.len(),"output_bytes":output.len()}));
        }
    }
    fs::write(
        format!("{dir}/perf.json"),
        serde_json::to_vec_pretty(&results)?,
    )?;
    println!("{}", serde_json::to_string_pretty(&results)?);
    Ok(())
}

fn zip_part(bytes: &[u8], part: &str) -> Result<Option<Vec<u8>>> {
    let mut zip = ZipArchive::new(std::io::Cursor::new(bytes))?;
    let result = match zip.by_name(part) {
        Ok(mut f) => {
            let mut b = Vec::new();
            std::io::Read::read_to_end(&mut f, &mut b)?;
            Ok(Some(b))
        }
        Err(zip::result::ZipError::FileNotFound) => Ok(None),
        Err(e) => Err(e.into()),
    };
    result
}
