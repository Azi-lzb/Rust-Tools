//! Fidelity fixture generation via Excel COM (Rust IDispatch, Windows only).
//!
//! Rationale: the PowerShell COM binder on this machine mis-resolves several
//! property-put shapes (array elements, 2D arrays via `Range.Value2`, plain
//! loop variables into `Cells.Item`), so fixture generation moved into this
//! deterministic Rust implementation. Every fixture gets its own Excel
//! instance; every cell write is an explicit `Range.Value2` / `Range.Formula`
//! property-put with a typed VARIANT.
//!
//! Fixture notes (recorded verbatim in the manifest):
//! - F09 dual BIFF8+BIFF5 (SaveAs format 43) and F13 BIFF4 (format 33) are
//!   refused by the installed Excel build.
//! - F11 keeps only the Excel4 macro sheet. A real VBA project was attempted
//!   first and got the workbook quarantined by Windows Defender, so VBA is not
//!   embedded on purpose.

use crate::combase::{sta_call, Apartment, AutoVariant, ComApp};
use crate::common::sha256_hex;
use anyhow::{Context, Result};
use std::path::{Path, PathBuf};
use windows::Win32::System::Com::IDispatch;

const XL_WORKBOOK_NORMAL: i32 = -4143;
const XL_EXCEL5: i32 = 39;
const XL_EXCEL9795: i32 = 43;
const XL_EXCEL4_WORKBOOK: i32 = 33;
const XL_CALC_AUTO: i32 = -4105;
const XL_CALC_MANUAL: i32 = -4135;
const XL_PORTRAIT: i32 = 1;
const XL_LANDSCAPE: i32 = 2;
const XL_PAPER_LETTER: i32 = 1;
const XL_PAPER_A3: i32 = 8;
const XL_PAPER_A4: i32 = 9;
const XL_GREATER: i32 = 5;
const XL_VALIDATE_LIST: i32 = 3;
const XL_ALERT_STOP: i32 = 1;
const XL_COLUMN_CLUSTERED: i32 = 51;
const XL_SHEET_HIDDEN: i32 = 0;
const XL_SHEET_VERY_HIDDEN: i32 = 2;
const XL_EXCEL4_MACRO_SHEET: i32 = 3;
const XL_CHART_SHEET: i32 = -4109;
const XL_OVER_THEN_DOWN: i32 = 2;
const MSO_FALSE: i32 = 0;
const MSO_TRUE: i32 = -1;

#[derive(Default, Clone)]
struct FixtureMeta {
    category: String,
    note: String,
}

/// `(file_name, meta)` produced by each fixture generator.
type FixtureResult = Result<(String, FixtureMeta)>;

pub fn run(out_dir: &Path) -> Result<()> {
    // Excel SaveAs resolves relative paths against its own process cwd, so
    // absolutize manually (std canonicalize yields `\\?\` verbatim paths that
    // Excel's SaveAs rejects).
    let out_dir = if out_dir.is_absolute() {
        out_dir.to_path_buf()
    } else {
        std::env::current_dir()?.join(out_dir)
    };
    std::fs::create_dir_all(&out_dir)?;
    let out_dir = &*out_dir;
    let mut metas: Vec<(String, FixtureMeta)> = Vec::new();
    metas.push(f01_plain(out_dir)?);
    metas.push(f02_formulas(out_dir)?);
    metas.push(f03_styles(out_dir)?);
    metas.push(f04_merges(out_dir)?);
    metas.push(f05_print(out_dir)?);
    metas.push(f06_multisheet(out_dir)?);
    metas.push(f07_unicode(out_dir)?);
    metas.push(f08_biff5(out_dir)?);
    metas.push(f09_dual(out_dir)?);
    metas.push(f10_objects(out_dir)?);
    metas.push(f11_macro(out_dir)?);
    metas.push(f12_password(out_dir)?);
    metas.push(f12b_protected(out_dir)?);
    metas.push(f13_biff4(out_dir)?);
    write_manifest(out_dir, metas)
}

/// One fixture = one owned Excel instance on one STA thread.
fn with_excel<T>(work: impl FnOnce(&ComApp) -> Result<T> + Send + 'static) -> Result<T>
where
    T: Send + 'static,
{
    sta_call(move || {
        let apartment = Apartment::initialize()?;
        let mut app = ComApp::create_and_configure("Excel.Application", false)
            .context("create Excel instance for fixture generation")?;
        let result = work(&app);
        let _ = app.shutdown();
        drop(app);
        drop(apartment);
        result
    })
}

fn new_book(app: &ComApp) -> Result<(IDispatch, IDispatch)> {
    use anyhow::Context as _;
    let app_object = app
        .object_dispatch()
        .context("new_book: app dispatch")?;
    let books = crate::combase::property(&app_object, "Workbooks")
        .and_then(|v| v.dispatch())
        .context("new_book: Workbooks")?;
    let book = crate::combase::method(&books, "Add", vec![])
        .and_then(|v| v.dispatch())
        .context("new_book: Workbooks.Add")?;
    loop {
        let sheets = crate::combase::property(&book, "Worksheets")
            .and_then(|v| v.dispatch())
            .context("new_book: Worksheets")?;
        let count = crate::combase::property(&sheets, "Count")
            .and_then(|v| v.integer())
            .context("new_book: Worksheets.Count")?;
        if count <= 1 {
            break;
        }
        let last = crate::combase::method(&sheets, "Item", vec![AutoVariant::int(count as i32)])
            .and_then(|v| v.dispatch())
            .with_context(|| format!("new_book: delete sheet {count}"))?;
        crate::combase::method(&last, "Delete", vec![]).context("new_book: sheet delete")?;
    }
    let sheets = crate::combase::property(&book, "Worksheets")
        .and_then(|v| v.dispatch())
        .context("new_book: Worksheets")?;
    let ws = crate::combase::method(&sheets, "Item", vec![AutoVariant::int(1)])
        .and_then(|v| v.dispatch())
        .context("new_book: first sheet")?;
    crate::combase::put(&ws, "Name", AutoVariant::text("Sheet1"))
        .context("new_book: rename")?;
    Ok((book, ws))
}

fn save_book(book: &IDispatch, path: &Path, format: i32) -> Result<()> {
    crate::combase::method(
        book,
        "SaveAs",
        vec![
            AutoVariant::text(&path.to_string_lossy().replace('/', "\\")),
            AutoVariant::int(format),
        ],
    )?;
    let bytes = std::fs::metadata(path).map(|m| m.len()).unwrap_or(0);
    eprintln!("saved {} ({bytes} bytes)", path.display());
    Ok(())
}

fn close_book(book: &IDispatch) {
    let _ = crate::combase::method(book, "Close", vec![AutoVariant::boolean(false)]);
}

fn put(object: &IDispatch, name: &str, value: AutoVariant) -> Result<()> {
    crate::combase::put(object, name, value)
}

fn get_dispatch(object: &IDispatch, name: &str) -> Result<IDispatch> {
    crate::combase::property(object, name)?.dispatch()
}

fn get_dispatch_arg(object: &IDispatch, name: &str, arg: AutoVariant) -> Result<IDispatch> {
    crate::combase::invoke_arg(object, name, arg)?.dispatch()
}

fn range_of(ws: &IDispatch, address: &str) -> Result<IDispatch> {
    get_dispatch_arg(ws, "Range", AutoVariant::text(address))
}

enum CellValue {
    Text(String),
    Num(f64),
    Bool(bool),
    Formula(String),
}

fn set_cell(ws: &IDispatch, address: &str, value: CellValue) -> Result<()> {
    let cell = range_of(ws, address)?;
    match value {
        CellValue::Text(t) => put(&cell, "Value2", AutoVariant::text(&t)),
        CellValue::Num(n) => put(&cell, "Value2", AutoVariant::float(n)),
        CellValue::Bool(b) => put(&cell, "Value2", AutoVariant::boolean(b)),
        CellValue::Formula(f) => put(&cell, "Formula", AutoVariant::text(&f)),
    }
}

fn set_cells(ws: &IDispatch, cells: &[(&str, CellValue)]) -> Result<()> {
    for (address, value) in cells {
        set_cell(ws, address, clone_value(value))
            .with_context(|| format!("set cell {address}"))?;
    }
    Ok(())
}

fn clone_value(value: &CellValue) -> CellValue {
    match value {
        CellValue::Text(t) => CellValue::Text(t.clone()),
        CellValue::Num(n) => CellValue::Num(*n),
        CellValue::Bool(b) => CellValue::Bool(*b),
        CellValue::Formula(f) => CellValue::Formula(f.clone()),
    }
}

fn set_number_format(ws: &IDispatch, address: &str, format: &str) -> Result<()> {
    let cell = range_of(ws, address)?;
    put(&cell, "NumberFormat", AutoVariant::text(format))
}


/// Best-effort number-format assignment for the style fixture; individual
/// format codes can be rejected by the Excel build, so log and continue.
fn try_number_format(ws: &IDispatch, address: &str, format: &str) {
    if let Err(e) = set_number_format(ws, address, format) {
        eprintln!("F03: NumberFormat {address} = {format} failed ({e:#}); skipped");
    }
}

fn put_font(ws: &IDispatch, address: &str, property: &str, value: AutoVariant) -> Result<()> {
    let cell = range_of(ws, address)?;
    let font = get_dispatch(&cell, "Font")?;
    put(&font, property, value)
}

fn put_interior(ws: &IDispatch, address: &str, color: i32) -> Result<()> {
    let cell = range_of(ws, address)?;
    let interior = get_dispatch(&cell, "Interior")?;
    put(&interior, "Color", AutoVariant::int(color))
}

/// Best-effort border styling: Excel rejects some argument shapes on this
/// build (medium/thick weights via BorderAround and per-edge Weight puts all
/// raised DISP_E_EXCEPTION), so failures are logged and skipped instead of
/// aborting the fixture. Whatever Excel stores is the source of truth.
fn try_border(ws: &IDispatch, address: &str, line_style: i32, weight: i32) {
    if border_around(ws, address, line_style, weight).is_ok() {
        return;
    }
    eprintln!(
        "F03: border {address} (style {line_style}, weight {weight}) could not be set; skipped"
    );
}

fn border_around(ws: &IDispatch, address: &str, line_style: i32, weight: i32) -> Result<()> {
    let cell = range_of(ws, address)?;
    if crate::combase::method(
        &cell,
        "BorderAround",
        vec![AutoVariant::int(line_style), AutoVariant::int(weight)],
    )
    .is_ok()
    {
        return Ok(());
    }
    // Fallback: per-edge Borders.Item put (some BorderAround argument shapes
    // raise DISP_E_EXCEPTION on this Excel build).
    let borders = get_dispatch(&cell, "Borders")?;
    for edge in [7i32, 8, 9, 10] {
        let edge_border =
            crate::combase::method(&borders, "Item", vec![AutoVariant::int(edge)])?.dispatch()?;
        put(&edge_border, "LineStyle", AutoVariant::int(line_style))?;
        put(&edge_border, "Weight", AutoVariant::int(weight))?;
    }
    Ok(())
}

fn add_sheet_after(
    book: &IDispatch,
    after: &IDispatch,
    name: &str,
    sheet_type: Option<i32>,
) -> Result<IDispatch> {
    let sheets = get_dispatch(book, "Worksheets")?;
    let mut args = vec![
        AutoVariant::missing(),
        AutoVariant::from_dispatch(after.clone())?,
        AutoVariant::int(1),
    ];
    match sheet_type {
        Some(t) => args.push(AutoVariant::int(t)),
        None => args.push(AutoVariant::missing()),
    }
    let sheet = crate::combase::method(&sheets, "Add", args)?.dispatch()?;
    put(&sheet, "Name", AutoVariant::text(name))?;
    Ok(sheet)
}

fn inches(app: &ComApp, value: f64) -> Result<f64> {
    let app_object = app.object_dispatch()?;
    let result =
        crate::combase::method(&app_object, "InchesToPoints", vec![AutoVariant::float(value)])?;
    Ok(result.as_f64()?)
}

fn safe_put(object: &IDispatch, property: &str, value: AutoVariant, tag: &str) -> bool {
    if let Err(e) = put(object, property, value) {
        eprintln!("{tag}: could not set {property} ({e:#})");
        false
    } else {
        true
    }
}

// --------------------------------------------------------------- F01

fn f01_plain(out_dir: &Path) -> FixtureResult {
    let out = out_dir.join("F01-plain.xls");
    let (name, note) = with_excel(move |app| -> Result<(String, String)> {
        let (book, ws) = new_book(app)?;
        put(&ws, "Name", AutoVariant::text("Data01"))?;
        set_cells(
            &ws,
            &[
                ("A1", CellValue::Text("文本text".into())),
                ("B1", CellValue::Num(42.0)),
                ("B2", CellValue::Num(3.125)),
                ("B3", CellValue::Num(-7.5)),
                ("B4", CellValue::Num(1e15)),
                ("B5", CellValue::Num(0.1 + 0.2)),
                ("B6", CellValue::Num(0.0001)),
                ("B7", CellValue::Num(1234567890.123456)),
                ("C1", CellValue::Num(0.125)),
                ("C2", CellValue::Num(0.005)),
                ("D1", CellValue::Num(45366.0)),
                ("D2", CellValue::Num(36537.9993055556)),
                ("E1", CellValue::Num(0.572916666666667)),
                ("F1", CellValue::Num(45139.2708333333)),
                ("G1", CellValue::Bool(true)),
                ("G2", CellValue::Bool(false)),
                ("I1", CellValue::Text("00123".into())),
                ("I2", CellValue::Text("1e5".into())),
                ("J1", CellValue::Text("0755-12345678".into())),
                ("K1", CellValue::Text("12345678901234567890".into())),
                ("K2", CellValue::Num(12345678901234567.0)),
                ("M1", CellValue::Num(0.0)),
            ],
        )?;
        set_cell(&ws, "H1", CellValue::Formula("=1/0".into()))?;
        set_cell(&ws, "H2", CellValue::Formula("=NA()".into()))?;
        set_cell(
            &ws,
            "H3",
            CellValue::Formula("=VLOOKUP(99,B1:B7,1,FALSE)".into()),
        )?;
        set_number_format(&ws, "C1", "0.00%")?;
        set_number_format(&ws, "C2", "0.0%")?;
        set_number_format(&ws, "D1", "yyyy-mm-dd")?;
        set_number_format(&ws, "D2", "yyyy-mm-dd hh:mm")?;
        set_number_format(&ws, "E1", "h:mm:ss")?;
        set_number_format(&ws, "F1", "yyyy/m/d h:mm:ss")?;
        set_number_format(&ws, "J1", "@")?;
        save_book(&book, &out, XL_WORKBOOK_NORMAL)?;
        close_book(&book);
        Ok((
            "F01-plain.xls".to_owned(),
            "text/number/percent/date/time/datetime/bool/error-cache/text-number/leading-zero/precision/zero"
                .to_owned(),
        ))
    })?;
    Ok((name, FixtureMeta { category: "data-types".into(), note }))
}

// --------------------------------------------------------------- F02

fn f02_formulas(out_dir: &Path) -> FixtureResult {
    let out = out_dir.join("F02-formulas.xls");
    let (name, note) = with_excel(move |app| -> Result<(String, String)> {
        let (book, ws) = new_book(app)?;
        put(&ws, "Name", AutoVariant::text("Formulas"))?;
        set_cells(
            &ws,
            &[
                ("B1", CellValue::Num(5.0)),
                ("C1", CellValue::Text("shared-block".into())),
                ("D1", CellValue::Text("array".into())),
                ("E1", CellValue::Text("legacy-fn".into())),
                ("F1", CellValue::Text("no-cache".into())),
            ],
        )?;
        for i in 1..=9 {
            set_cell(&ws, &format!("A{i}"), CellValue::Num((i * 10) as f64))?;
        }
        let formulas: Vec<(&str, &str)> = vec![
            ("B2", "=SUM(A2:A10)"),
            ("B3", "=IF(A2>50,\"BIG\",\"SMALL\")"),
            ("B4", "=VLOOKUP(30,A2:A10,1,FALSE)"),
            ("B5", "=COUNTIF(A2:A10,\">40\")"),
            ("B6", "=SUMIF(A2:A10,\">40\")"),
            ("B8", "=A2*$B$1"),
            ("B9", "=DATE(2024,3,15)+1"),
            ("B10", "=IF(SUM(A2:A3)>2,MAX(A2:A10),MIN(A2:A10))"),
            ("B11", "=ROUND(PI(),2)"),
            ("C2", "=A2*2"),
            ("E2", "=UNKNOWNFN1(1)"),
            ("E3", "=BESSELJ(2,1)"),
        ];
        for (address, formula) in formulas {
            set_cell(&ws, address, CellValue::Formula(formula.into()))
                .with_context(|| format!("formula {address}"))?;
        }
        let block = range_of(&ws, "C2:C10")?;
        crate::combase::method(&block, "FillDown", vec![])?;
        let array_cell = range_of(&ws, "D2")?;
        put(&array_cell, "FormulaArray", AutoVariant::text("=SUM(A2:A10*1)"))?;

        let ws2 = add_sheet_after(&book, &ws, "Data2", None)?;
        set_cells(
            &ws2,
            &[
                ("A1", CellValue::Num(100.0)),
                ("B1", CellValue::Num(200.0)),
            ],
        )?;
        set_cell(&ws, "B7", CellValue::Formula("=Data2!A1+Data2!B1".into()))?;

        let b12 = range_of(&ws, "B12")?;
        put(&b12, "Value2", AutoVariant::text("=IF(A2>50,\"大\",\"小\")"))?;

        let app_object = app.object_dispatch()?;
        put(&app_object, "Calculation", AutoVariant::int(XL_CALC_MANUAL))?;
        set_cell(&ws, "F2", CellValue::Formula("=SUM(A2:A10)".into()))?;
        set_cell(&ws, "F3", CellValue::Formula("=A2+A3".into()))?;
        put(&app_object, "Calculation", AutoVariant::int(XL_CALC_AUTO))?;

        save_book(&book, &out, XL_WORKBOOK_NORMAL)?;
        close_book(&book);
        Ok((
            "F02-formulas.xls".to_owned(),
            "SUM/IF/VLOOKUP/COUNTIF/SUMIF/cross-sheet/abs-ref/date/nested/shared/array/unknown-fn/ATP/no-cache/CJK-string-formula"
                .to_owned(),
        ))
    })?;
    Ok((name, FixtureMeta { category: "formulas".into(), note }))
}

// --------------------------------------------------------------- F03

fn f03_styles(out_dir: &Path) -> FixtureResult {
    let out = out_dir.join("F03-styles.xls");
    let (name, note) = with_excel(move |app| -> Result<(String, String)> {
        let (book, ws) = new_book(app)?;
        put(&ws, "Name", AutoVariant::text("Styles"))?;
        let fonts = ["宋体", "微软雅黑", "Arial", "Calibri"];
        let sizes = [8.0f64, 10.0, 12.0, 14.0];
        let columns = ["A", "B", "C", "D"];
        for (index, (font, size)) in fonts.iter().zip(sizes.iter()).enumerate() {
            let address = format!("{}1", columns[index]);
            set_cell(&ws, &address, CellValue::Text((*font).into()))?;
            put_font(&ws, &address, "Name", AutoVariant::text(font))?;
            put_font(&ws, &address, "Size", AutoVariant::float(*size))?;
        }
        set_cell(&ws, "A5", CellValue::Text("bold".into()))?;
        put_font(&ws, "A5", "Bold", AutoVariant::boolean(true))?;
        set_cell(&ws, "B5", CellValue::Text("italic".into()))?;
        put_font(&ws, "B5", "Italic", AutoVariant::boolean(true))?;
        set_cell(&ws, "C5", CellValue::Text("underline".into()))?;
        put_font(&ws, "C5", "Underline", AutoVariant::int(2))?;
        set_cell(&ws, "D5", CellValue::Text("strike".into()))?;
        put_font(&ws, "D5", "Strikethrough", AutoVariant::boolean(true))?;
        set_cell(&ws, "A6", CellValue::Text("红字".into()))?;
        put_font(&ws, "A6", "Color", AutoVariant::int(255))?;
        set_cell(&ws, "B6", CellValue::Text("蓝字".into()))?;
        put_font(&ws, "B6", "Color", AutoVariant::int(16711680))?;
        set_cell(&ws, "C6", CellValue::Text("黄底".into()))?;
        put_interior(&ws, "C6", 65535)?;
        set_cell(&ws, "D6", CellValue::Text("绿底".into()))?;
        put_interior(&ws, "D6", 65280)?;

        set_cell(&ws, "A8", CellValue::Text("细边框".into()))?;
        try_border(&ws, "A8", 1, 2);
        set_cell(&ws, "B8", CellValue::Text("双线".into()))?;
        try_border(&ws, "B8", 9, 4);
        set_cell(&ws, "C8", CellValue::Text("虚线".into()))?;
        try_border(&ws, "C8", -4115, 2);
        set_cell(&ws, "D8", CellValue::Text("中边框".into()))?;
        try_border(&ws, "D8", 1, -4136);
        set_cell(&ws, "E8", CellValue::Text("粗边框".into()))?;
        try_border(&ws, "E8", 1, 4);

        set_cell(&ws, "A12", CellValue::Text("居中".into()))?;
        put(&range_of(&ws, "A12")?, "HorizontalAlignment", AutoVariant::int(-4108))?;
        set_cell(&ws, "B12", CellValue::Text("右对齐".into()))?;
        put(&range_of(&ws, "B12")?, "HorizontalAlignment", AutoVariant::int(-4152))?;
        set_cell(
            &ws,
            "C12",
            CellValue::Text("换行文本测试这段文字足够长会触发自动换行".into()),
        )?;
        put(&range_of(&ws, "C12")?, "WrapText", AutoVariant::boolean(true))?;
        set_cell(&ws, "D12", CellValue::Text("缩进".into()))?;
        put(&range_of(&ws, "D12")?, "IndentLevel", AutoVariant::int(1))?;
        set_cell(&ws, "E12", CellValue::Text("旋转".into()))?;
        put(&range_of(&ws, "E12")?, "Orientation", AutoVariant::int(45))?;
        set_cell(&ws, "A13", CellValue::Text("顶端".into()))?;
        put(&range_of(&ws, "A13")?, "VerticalAlignment", AutoVariant::int(-4160))?;
        set_cell(&ws, "B13", CellValue::Text("底端".into()))?;
        put(&range_of(&ws, "B13")?, "VerticalAlignment", AutoVariant::int(-4107))?;

        let formats = [
            "General",
            "0.00%",
            "#,##0.00",
            "0.00E+00",
            "yyyy-mm-dd",
            "h:mm:ss",
            "@",
            "[Red]0.00;[Blue]-0.00;0.00",
            "¥#,##0.00",
            "0.000",
            "0.0\"kg\"",
            "00000",
        ];
        for (index, format) in formats.iter().enumerate() {
            let row = 15 + index as i32;
            set_cell(&ws, &format!("A{row}"), CellValue::Text(format!("fmt:{format}")))?;
            set_cell(&ws, &format!("B{row}"), CellValue::Num(1234.5678))?;
            try_number_format(&ws, &format!("B{row}"), format);
        }
        save_book(&book, &out, XL_WORKBOOK_NORMAL)?;
        close_book(&book);
        Ok((
            "F03-styles.xls".to_owned(),
            "fonts/sizes/decoration/colors/fills/borders/alignment/rotation/number-formats"
                .to_owned(),
        ))
    })?;
    Ok((name, FixtureMeta { category: "styles".into(), note }))
}

// --------------------------------------------------------------- F04

fn f04_merges(out_dir: &Path) -> FixtureResult {
    let out = out_dir.join("F04-merges-dims.xls");
    let (name, note) = with_excel(move |app| -> Result<(String, String)> {
        let (book, ws) = new_book(app)?;
        put(&ws, "Name", AutoVariant::text("Merges"))?;
        let merge1 = range_of(&ws, "A1:C1")?;
        crate::combase::method(&merge1, "Merge", vec![])?;
        set_cell(&ws, "A1", CellValue::Text("合并标题".into()))?;
        let merge2 = range_of(&ws, "B3:D5")?;
        crate::combase::method(&merge2, "Merge", vec![])?;
        set_cell(&ws, "B3", CellValue::Text("大块合并".into()))?;
        let merge3 = range_of(&ws, "E1:E3")?;
        crate::combase::method(&merge3, "Merge", vec![])?;
        set_cell(&ws, "E1", CellValue::Text("竖排合并长文本自动换行测试".into()))?;
        put(&range_of(&ws, "E1")?, "WrapText", AutoVariant::boolean(true))?;

        let col_b = get_dispatch_arg(&ws, "Columns", AutoVariant::text("B"))?;
        put(&col_b, "Hidden", AutoVariant::boolean(true))?;
        let row5 = get_dispatch_arg(&ws, "Rows", AutoVariant::text("5"))?;
        put(&row5, "Hidden", AutoVariant::boolean(true))?;
        let col_a = get_dispatch_arg(&ws, "Columns", AutoVariant::text("A"))?;
        put(&col_a, "ColumnWidth", AutoVariant::float(25.0))?;
        let col_c = get_dispatch_arg(&ws, "Columns", AutoVariant::text("C"))?;
        put(&col_c, "ColumnWidth", AutoVariant::float(15.0))?;
        let row1 = get_dispatch_arg(&ws, "Rows", AutoVariant::text("1"))?;
        put(&row1, "RowHeight", AutoVariant::float(30.0))?;
        let row2 = get_dispatch_arg(&ws, "Rows", AutoVariant::text("2"))?;
        put(&row2, "RowHeight", AutoVariant::float(18.0))?;
        let row3 = get_dispatch_arg(&ws, "Rows", AutoVariant::text("3"))?;
        put(&row3, "RowHeight", AutoVariant::float(45.0))?;

        let app_object = app.object_dispatch()?;
        crate::combase::method(&range_of(&ws, "B2")?, "Select", vec![])?;
        let window = get_dispatch(&app_object, "ActiveWindow")?;
        put(&window, "FreezePanes", AutoVariant::boolean(true))?;

        save_book(&book, &out, XL_WORKBOOK_NORMAL)?;
        close_book(&book);
        Ok((
            "F04-merges-dims.xls".to_owned(),
            "merges/hidden-row-col/column-widths/row-heights/freeze-panes".to_owned(),
        ))
    })?;
    Ok((name, FixtureMeta { category: "merges-dims".into(), note }))
}

// --------------------------------------------------------------- F05

fn f05_print(out_dir: &Path) -> FixtureResult {
    let out = out_dir.join("F05-print.xls");
    let (name, note) = with_excel(move |app| -> Result<(String, String)> {
        let (book, ws1) = new_book(app)?;
        put(&ws1, "Name", AutoVariant::text("PrintA"))?;
        for row in 1..=30 {
            for col in 1..=6 {
                set_cell(
                    &ws1,
                    &format!("{}{row}", ('A'..='Z').nth(col - 1).unwrap()),
                    CellValue::Text(format!("P{row}-{col}")),
                )?;
            }
        }
        let setup1 = get_dispatch(&ws1, "PageSetup")?;
        safe_put(&setup1, "PrintArea", AutoVariant::text("A1:F30"), "F05");
        safe_put(&setup1, "PrintTitleRows", AutoVariant::text("$1:$2"), "F05");
        safe_put(&setup1, "PaperSize", AutoVariant::int(XL_PAPER_A4), "F05");
        safe_put(&setup1, "Orientation", AutoVariant::int(XL_PORTRAIT), "F05");
        safe_put(&setup1, "LeftMargin", AutoVariant::float(inches(app, 0.7)?), "F05");
        safe_put(&setup1, "RightMargin", AutoVariant::float(inches(app, 0.7)?), "F05");
        safe_put(&setup1, "TopMargin", AutoVariant::float(inches(app, 0.9)?), "F05");
        safe_put(&setup1, "BottomMargin", AutoVariant::float(inches(app, 0.9)?), "F05");
        safe_put(&setup1, "CenterHeader", AutoVariant::text("&F"), "F05");
        safe_put(&setup1, "CenterFooter", AutoVariant::text("&P/&N"), "F05");
        safe_put(&setup1, "Zoom", AutoVariant::int(75), "F05");
        safe_put(&setup1, "PrintGridlines", AutoVariant::boolean(true), "F05");
        safe_put(&setup1, "CenterHorizontally", AutoVariant::boolean(true), "F05");

        let ws2 = add_sheet_after(&book, &ws1, "PrintB", None)?;
        for row in 1..=50 {
            for col in 1..=8 {
                set_cell(
                    &ws2,
                    &format!("{}{row}", ('A'..='Z').nth(col - 1).unwrap()),
                    CellValue::Text(format!("Q{row}-{col}")),
                )?;
            }
        }
        let setup2 = get_dispatch(&ws2, "PageSetup")?;
        safe_put(&setup2, "PaperSize", AutoVariant::int(XL_PAPER_A3), "F05");
        safe_put(&setup2, "Orientation", AutoVariant::int(XL_LANDSCAPE), "F05");
        if !safe_put(&setup2, "Zoom", AutoVariant::int(0), "F05") {
            safe_put(&setup2, "Zoom", AutoVariant::missing(), "F05");
        }
        safe_put(&setup2, "FitToPagesWide", AutoVariant::int(1), "F05");
        safe_put(&setup2, "FitToPagesTall", AutoVariant::int(2), "F05");
        safe_put(&setup2, "Order", AutoVariant::int(XL_OVER_THEN_DOWN), "F05");
        safe_put(&setup2, "PrintHeadings", AutoVariant::boolean(true), "F05");
        safe_put(&setup2, "LeftHeader", AutoVariant::text("左侧页眉"), "F05");

        let ws3 = add_sheet_after(&book, &ws2, "PrintC", None)?;
        for row in 1..=20 {
            for col in 1..=4 {
                set_cell(
                    &ws3,
                    &format!("{}{row}", ('A'..='Z').nth(col - 1).unwrap()),
                    CellValue::Text(format!("R{row}-{col}")),
                )?;
            }
        }
        let setup3 = get_dispatch(&ws3, "PageSetup")?;
        safe_put(&setup3, "PaperSize", AutoVariant::int(XL_PAPER_LETTER), "F05");
        safe_put(&setup3, "CenterVertically", AutoVariant::boolean(true), "F05");
        safe_put(&setup3, "FirstPageNumber", AutoVariant::int(5), "F05");
        safe_put(&setup3, "RightFooter", AutoVariant::text("右页脚"), "F05");

        save_book(&book, &out, XL_WORKBOOK_NORMAL)?;
        close_book(&book);
        Ok((
            "F05-print.xls".to_owned(),
            "print-area/titles/paper/orientation/margins/header-footer/zoom/fit/order/gridlines/centering (paper may be partial: printer context)"
                .to_owned(),
        ))
    })?;
    Ok((name, FixtureMeta { category: "print".into(), note }))
}

// --------------------------------------------------------------- F06

fn f06_multisheet(out_dir: &Path) -> FixtureResult {
    let out = out_dir.join("F06-multisheet.xls");
    let (name, note) = with_excel(move |app| -> Result<(String, String)> {
        let (book, ws) = new_book(app)?;
        put(&ws, "Name", AutoVariant::text("中文表"))?;
        set_cell(&ws, "A1", CellValue::Text("第一个工作表".into()))?;
        let ws2 = add_sheet_after(&book, &ws, "Data2", None)?;
        set_cell(&ws2, "A1", CellValue::Text("second".into()))?;
        let ws3 = add_sheet_after(&book, &ws2, "Hidden1", None)?;
        set_cell(&ws3, "A1", CellValue::Text("隐藏表内容".into()))?;
        put(&ws3, "Visible", AutoVariant::int(XL_SHEET_HIDDEN))?;
        let ws4 = add_sheet_after(&book, &ws3, "VeryHide", None)?;
        set_cell(&ws4, "A1", CellValue::Text("深度隐藏".into()))?;
        put(&ws4, "Visible", AutoVariant::int(XL_SHEET_VERY_HIDDEN))?;
        let last = add_sheet_after(&book, &ws4, "末页", None)?;
        set_cell(&last, "A1", CellValue::Text("最后".into()))?;
        let all_sheets = get_dispatch(&book, "Sheets")?;
        let chart_sheet = crate::combase::method(
            &all_sheets,
            "Add",
            vec![
                AutoVariant::missing(),
                AutoVariant::from_dispatch(last.clone())?,
                AutoVariant::int(1),
                AutoVariant::int(XL_CHART_SHEET),
            ],
        )?
        .dispatch()?;
        put(&chart_sheet, "Name", AutoVariant::text("图表页"))?;
        save_book(&book, &out, XL_WORKBOOK_NORMAL)?;
        close_book(&book);
        Ok((
            "F06-multisheet.xls".to_owned(),
            "sheet-order/chinese-names/hidden/very-hidden/chartsheet".to_owned(),
        ))
    })?;
    Ok((name, FixtureMeta { category: "workbook-structure".into(), note }))
}

// --------------------------------------------------------------- F07

fn f07_unicode(out_dir: &Path) -> FixtureResult {
    let out = out_dir.join("F07-chinese-legacy.xls");
    let (name, note) = with_excel(move |app| -> Result<(String, String)> {
        let (book, ws) = new_book(app)?;
        put(&ws, "Name", AutoVariant::text("Legacy"))?;
        let strings = vec![
            "简体中文测试".to_owned(),
            "繁體中文測試".to_owned(),
            "日本語のテキスト".to_owned(),
            "한국어 텍스트".to_owned(),
            "éàçüö Latin-1".to_owned(),
            "特殊字符 &<>\"'\\".to_owned(),
            "前导空格    ".to_owned(),
            "    尾随空格".to_owned(),
            "长文本".repeat(40),
            "=SUM(A1) 伪装文本".to_owned(),
            "00123".to_owned(),
            "3.14 是文本".to_owned(),
            "换行\n测试".to_owned(),
            "😀 emoji 行".to_owned(),
            "混排 English 中文 日本語".to_owned(),
        ];
        for (index, text) in strings.iter().enumerate() {
            let address = format!("A{}", index + 1);
            if index == 9 {
                set_number_format(&ws, &address, "@")?;
            }
            set_cell(&ws, &address, CellValue::Text(text.clone()))
                .with_context(|| format!("F07 item {index}"))?;
        }
        save_book(&book, &out, XL_WORKBOOK_NORMAL)?;
        close_book(&book);
        Ok((
            "F07-chinese-legacy.xls".to_owned(),
            "CJK/latin1/special-chars/long-text/spaces/emoji/multiline/formula-like-text"
                .to_owned(),
        ))
    })?;
    Ok((name, FixtureMeta { category: "unicode".into(), note }))
}

// --------------------------------------------------------------- F08

fn f08_biff5(out_dir: &Path) -> FixtureResult {
    let out = out_dir.join("F08-biff5.xls");
    let (name, note) = with_excel(move |app| -> Result<(String, String)> {
        let (book, ws) = new_book(app)?;
        put(&ws, "Name", AutoVariant::text("旧版"))?;
        set_cells(
            &ws,
            &[
                ("A1", CellValue::Text("旧格式内容".into())),
                ("A2", CellValue::Num(123.45)),
                ("A3", CellValue::Text("GBK 时代字符串".into())),
            ],
        )?;
        set_cell(&ws, "B1", CellValue::Formula("=A2*2".into()))?;
        save_book(&book, &out, XL_EXCEL5)?;
        close_book(&book);
        Ok((
            "F08-biff5.xls".to_owned(),
            "BIFF5 (Excel 5.0/95), codepage strings".to_owned(),
        ))
    })?;
    Ok((name, FixtureMeta { category: "biff5".into(), note }))
}

// --------------------------------------------------------------- F09

fn f09_dual(out_dir: &Path) -> FixtureResult {
    let out = out_dir.join("F09-biff5-dual.xls");
    let (name, note) = with_excel(move |app| -> Result<(String, String)> {
        let (book, ws) = new_book(app)?;
        put(&ws, "Name", AutoVariant::text("Dual"))?;
        set_cells(
            &ws,
            &[
                ("A1", CellValue::Text("双格式".into())),
                ("A2", CellValue::Num(7.25)),
            ],
        )?;
        let note = match save_book(&book, &out, XL_EXCEL9795) {
            Ok(()) => "dual BIFF8+BIFF5 container (Excel 97-2000 & 5.0/95)".to_owned(),
            Err(e) => {
                eprintln!("F09 dual save refused by this Excel build: {e:#}");
                "NOT GENERATED - this Excel build refuses xlExcel9795 (43)".to_owned()
            }
        };
        close_book(&book);
        Ok(("F09-biff5-dual.xls".to_owned(), note))
    })?;
    Ok((name, FixtureMeta { category: "biff5".into(), note }))
}

// --------------------------------------------------------------- F10

fn f10_objects(out_dir: &Path) -> FixtureResult {
    let out = out_dir.join("F10-objects.xls");
    let dir = out_dir.to_path_buf();
    let (name, note) = with_excel(move |app| -> Result<(String, String)> {
        let png_path = dir.join("_test.png");
        write_png(
            &png_path,
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
        )?;

        let (book, ws) = new_book(app)?;
        put(&ws, "Name", AutoVariant::text("Objects"))?;
        for row in 1..=5 {
            set_cell(&ws, &format!("A{row}"), CellValue::Num(row as f64))?;
            set_cell(&ws, &format!("B{row}"), CellValue::Num((row * 2) as f64))?;
        }
        let shapes = get_dispatch(&ws, "Shapes")?;
        crate::combase::method(
            &shapes,
            "AddPicture",
            vec![
                AutoVariant::text(&png_path.to_string_lossy().replace('/', "\\")),
                AutoVariant::int(MSO_FALSE),
                AutoVariant::int(MSO_TRUE),
                AutoVariant::float(260.0),
                AutoVariant::float(10.0),
                AutoVariant::float(60.0),
                AutoVariant::float(40.0),
            ],
        )?;
        let chart_objects = crate::combase::method(&ws, "ChartObjects", vec![AutoVariant::missing()])?
            .dispatch()?;
        let chart_object = crate::combase::method(
            &chart_objects,
            "Add",
            vec![
                AutoVariant::float(300.0),
                AutoVariant::float(10.0),
                AutoVariant::float(300.0),
                AutoVariant::float(200.0),
            ],
        )?
        .dispatch()?;
        let chart = get_dispatch(&chart_object, "Chart")?;
        let source = range_of(&ws, "A1:B5")?;
        crate::combase::method(&chart, "SetSourceData", vec![AutoVariant::from_dispatch(source)?])?;
        put(&chart, "ChartType", AutoVariant::int(XL_COLUMN_CLUSTERED))?;
        crate::combase::method(
            &shapes,
            "AddTextbox",
            vec![
                AutoVariant::int(1),
                AutoVariant::float(260.0),
                AutoVariant::float(60.0),
                AutoVariant::float(200.0),
                AutoVariant::float(40.0),
            ],
        )?;

        set_cell(&ws, "A8", CellValue::Text("带超链接".into()))?;
        let anchors = get_dispatch(&ws, "Hyperlinks")?;
        crate::combase::method(
            &anchors,
            "Add",
            vec![
                AutoVariant::from_dispatch(range_of(&ws, "A8")?)?,
                AutoVariant::text(""),
                AutoVariant::text("https://example.com/"),
                AutoVariant::missing(),
                AutoVariant::text("外部链接"),
            ],
        )?;
        set_cell(&ws, "A9", CellValue::Text("内部链接".into()))?;
        crate::combase::method(
            &anchors,
            "Add",
            vec![
                AutoVariant::from_dispatch(range_of(&ws, "A9")?)?,
                AutoVariant::text("Sheet1!A1"),
                AutoVariant::missing(),
                AutoVariant::missing(),
                AutoVariant::text("内部链接"),
            ],
        )?;
        set_cell(&ws, "A10", CellValue::Text("有批注".into()))?;
        let a10 = range_of(&ws, "A10")?;
        crate::combase::method(&a10, "AddComment", vec![AutoVariant::text("这是批注内容")])?;

        let names = get_dispatch(&book, "Names")?;
        crate::combase::method(
            &names,
            "Add",
            vec![
                AutoVariant::text("MyData"),
                AutoVariant::text("=Sheet1!$A$1:$B$5"),
            ],
        )?;
        let d1 = range_of(&ws, "D1")?;
        let validation = get_dispatch(&d1, "Validation")?;
        crate::combase::method(
            &validation,
            "Add",
            vec![
                AutoVariant::int(XL_VALIDATE_LIST),
                AutoVariant::int(XL_ALERT_STOP),
                AutoVariant::missing(),
                AutoVariant::text("甲,乙,丙"),
            ],
        )?;
        let e_range = range_of(&ws, "E2:E5")?;
        let conditions = get_dispatch(&e_range, "FormatConditions")?;
        crate::combase::method(
            &conditions,
            "Add",
            vec![
                AutoVariant::int(1),
                AutoVariant::int(XL_GREATER),
                AutoVariant::text("3"),
            ],
        )?;
        let filter_range = range_of(&ws, "A1:B5")?;
        crate::combase::method(&filter_range, "AutoFilter", vec![AutoVariant::int(1)])?;

        save_book(&book, &out, XL_WORKBOOK_NORMAL)?;
        close_book(&book);
        let _ = std::fs::remove_file(&png_path);
        Ok((
            "F10-objects.xls".to_owned(),
            "image/chart/textbox/hyperlinks/comment/defined-name/validation/conditional-format/autofilter"
                .to_owned(),
        ))
    })?;
    Ok((name, FixtureMeta { category: "objects".into(), note }))
}

fn write_png(path: &Path, base64: &str) -> Result<()> {
    const TABLE: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let cleaned: Vec<u8> = base64
        .bytes()
        .filter(|b| !b.is_ascii_whitespace() && *b != b'=')
        .collect();
    let mut out = Vec::new();
    for chunk in cleaned.chunks(4) {
        let mut values = [0u8; 4];
        for (slot, byte) in values.iter_mut().zip(chunk) {
            *slot = TABLE.iter().position(|t| t == byte).unwrap_or(0) as u8;
        }
        out.push((values[0] << 2) | (values[1] >> 4));
        if chunk.len() > 1 {
            out.push((values[1] << 4) | (values[2] >> 2));
        }
        if chunk.len() > 2 {
            out.push((values[2] << 6) | values[3]);
        }
    }
    std::fs::write(path, out)?;
    Ok(())
}

// --------------------------------------------------------------- F11

fn f11_macro(out_dir: &Path) -> FixtureResult {
    let out = out_dir.join("F11-legacy-macro.xls");
    let (name, note) = with_excel(move |app| -> Result<(String, String)> {
        let (book, ws) = new_book(app)?;
        put(&ws, "Name", AutoVariant::text("Normal"))?;
        set_cell(&ws, "A1", CellValue::Text("宏测试".into()))?;
        let all_sheets = get_dispatch(&book, "Sheets")?;
        let macro_sheet = crate::combase::method(
            &all_sheets,
            "Add",
            vec![
                AutoVariant::missing(),
                AutoVariant::from_dispatch(ws.clone())?,
                AutoVariant::int(1),
                AutoVariant::int(XL_EXCEL4_MACRO_SHEET),
            ],
        )?
        .dispatch()?;
        put(&macro_sheet, "Name", AutoVariant::text("MACRO"))?;
        set_cell(
            &macro_sheet,
            "A1",
            CellValue::Formula("=GET.CELL(1,Normal!A1)".into()),
        )?;
        set_cell(&macro_sheet, "A2", CellValue::Formula("=HALT()".into()))?;
        save_book(&book, &out, XL_WORKBOOK_NORMAL)?;
        close_book(&book);
        Ok((
            "F11-legacy-macro.xls".to_owned(),
            "Excel4 macro sheet only (a real VBA project was quarantined by Windows Defender during the first attempt; VBA deliberately not embedded, and this file may be AV-flagged on read)".to_owned(),
        ))
    })?;
    Ok((name, FixtureMeta { category: "macros".into(), note }))
}

// --------------------------------------------------------------- F12

fn f12_password(out_dir: &Path) -> FixtureResult {
    let out = out_dir.join("F12-password.xls");
    let (name, note) = with_excel(move |app| -> Result<(String, String)> {
        let (book, ws) = new_book(app)?;
        put(&ws, "Name", AutoVariant::text("Secret"))?;
        set_cells(
            &ws,
            &[
                ("A1", CellValue::Text("加密内容".into())),
                ("A2", CellValue::Num(31337.0)),
            ],
        )?;
        crate::combase::method(
            &book,
            "SaveAs",
            vec![
                AutoVariant::text(&out.to_string_lossy().replace('/', "\\")),
                AutoVariant::int(XL_WORKBOOK_NORMAL),
                AutoVariant::text("xls2xtest"),
            ],
        )?;
        eprintln!("saved {} (encrypted)", out.display());
        close_book(&book);
        Ok((
            "F12-password.xls".to_owned(),
            "open password 'xls2xtest' (encrypted BIFF8)".to_owned(),
        ))
    })?;
    Ok((name, FixtureMeta { category: "password".into(), note }))
}

fn f12b_protected(out_dir: &Path) -> FixtureResult {
    let out = out_dir.join("F12b-protect-structure.xls");
    let (name, note) = with_excel(move |app| -> Result<(String, String)> {
        let (book, ws) = new_book(app)?;
        put(&ws, "Name", AutoVariant::text("Protected"))?;
        set_cell(&ws, "A1", CellValue::Text("结构保护".into()))?;
        crate::combase::method(
            &book,
            "Protect",
            vec![
                AutoVariant::missing(),
                AutoVariant::boolean(true),
                AutoVariant::boolean(false),
            ],
        )?;
        save_book(&book, &out, XL_WORKBOOK_NORMAL)?;
        close_book(&book);
        Ok((
            "F12b-protect-structure.xls".to_owned(),
            "workbook structure protection, no open password".to_owned(),
        ))
    })?;
    Ok((name, FixtureMeta { category: "password".into(), note }))
}

// --------------------------------------------------------------- F13

fn f13_biff4(out_dir: &Path) -> FixtureResult {
    let out = out_dir.join("F13-biff4.xls");
    let (name, note) = with_excel(move |app| -> Result<(String, String)> {
        let (book, ws) = new_book(app)?;
        put(&ws, "Name", AutoVariant::text("Biff4"))?;
        set_cell(&ws, "A1", CellValue::Num(1.0))?;
        let note = match save_book(&book, &out, XL_EXCEL4_WORKBOOK) {
            Ok(()) => "BIFF4 workbook (Excel 4.0)".to_owned(),
            Err(e) => {
                eprintln!("F13 BIFF4 save refused by this Excel build: {e:#}");
                "NOT GENERATED - this Excel build refuses xlExcel4Workbook (33)".to_owned()
            }
        };
        close_book(&book);
        Ok(("F13-biff4.xls".to_owned(), note))
    })?;
    Ok((name, FixtureMeta { category: "biff4".into(), note }))
}

// --------------------------------------------------------------- manifest

fn write_manifest(out_dir: &Path, metas: Vec<(String, FixtureMeta)>) -> Result<()> {
    let mut files = std::fs::read_dir(out_dir)?
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .filter(|p| {
            p.extension()
                .map(|e| e.eq_ignore_ascii_case("xls"))
                .unwrap_or(false)
        })
        .collect::<Vec<_>>();
    files.sort();
    let mut fixtures = Vec::new();
    for path in files {
        let name = path
            .file_name()
            .map(|n| n.to_string_lossy().into_owned())
            .unwrap_or_default();
        let meta = metas
            .iter()
            .find(|(n, _)| *n == name)
            .map(|(_, m)| m.clone())
            .unwrap_or_default();
        let bytes = std::fs::metadata(&path).map(|m| m.len()).unwrap_or(0);
        let sha256 = match std::fs::read(&path) {
            Ok(data) => sha256_hex(&data),
            Err(_) => "hash-blocked-by-av".to_owned(),
        };
        fixtures.push(serde_json::json!({
            "name": name,
            "bytes": bytes,
            "sha256": sha256,
            "category": meta.category,
            "note": meta.note,
        }));
    }
    let manifest = serde_json::json!({
        "tool_version": crate::common::TOOL_VERSION,
        "fixtures": fixtures,
    });
    let manifest_path = out_dir.join("manifest.json");
    std::fs::write(&manifest_path, serde_json::to_string_pretty(&manifest)?)?;
    eprintln!(
        "manifest written: {} ({} fixture files)",
        manifest_path.display(),
        fixtures.len()
    );
    Ok(())
}
