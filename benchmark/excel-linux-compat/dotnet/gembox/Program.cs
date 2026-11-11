using GemBox.Spreadsheet;
using System.Drawing;
using System.Diagnostics;
using System.Globalization;

SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

static void Need(bool ok, string message)
{
    if (!ok) throw new InvalidOperationException(message);
}

static string Csv(string value) => "\"" + value.Replace("\"", "\"\"") + "\"";
static string Num(double value) => value.ToString("0.###", CultureInfo.InvariantCulture);

static int RunPerformance(string inputRoot, string outputRoot, string csvPath)
{
    Directory.CreateDirectory(outputRoot);
    using var writer = new StreamWriter(csvPath, false, System.Text.Encoding.UTF8);
    writer.WriteLine("engine,round,file,format,input_bytes,output_bytes,open_ms,append_ms,edit_ms,delete_rows_cols_ms,style_ms,print_setup_ms,calculate_ms,save_ms,convert_ms,total_ms,peak_memory_mb,cpu_ms,status,error");
    var process = Process.GetCurrentProcess();
    var files = Directory.GetFiles(inputRoot).Where(p => Path.GetFileName(p).StartsWith("batch-", StringComparison.Ordinal)).OrderBy(p => p).ToArray();
    if (files.Length > 0)
    {
        var warm = ExcelFile.Load(files[0]);
        warm.Calculate();
        warm.Save(Path.Combine(outputRoot, "warmup.xls"), SaveOptions.XlsDefault);
        Console.WriteLine("WARMUP_COMPLETE");
    }
    for (var round = 1; round <= 3; round++)
    foreach (var input in files)
    {
        var t = new double[9];
        var start = Stopwatch.GetTimestamp();
        var cpuStart = process.TotalProcessorTime.TotalMilliseconds;
        var output = Path.Combine(outputRoot, $"r{round}-{Path.GetFileName(input)}");
        try
        {
            var sw = Stopwatch.StartNew();
            var wb = ExcelFile.Load(input);
            var ws = wb.Worksheets[0];
            t[0] = sw.Elapsed.TotalMilliseconds;
            sw.Restart();
            var used = ws.GetUsedCellRange(false);
            var row = used == null ? 0 : used.LastRowIndex + 1;
            ws.Rows.InsertEmpty(row, 5);
            for (var i = 0; i < 5; i++) { ws.Cells[row + i, 0].Value = $"perf-{i}"; ws.Cells[row + i, 1].Value = i; }
            t[1] = sw.Elapsed.TotalMilliseconds;
            sw.Restart(); ws.Cells["A1"].Value = $"edit-r{round}"; t[2] = sw.Elapsed.TotalMilliseconds;
            sw.Restart(); ws.Rows.InsertEmpty(row + 5, 1); ws.Rows.Remove(row + 5, 1); ws.Columns.InsertEmpty(10, 1); ws.Columns.Remove(10, 1); t[3] = sw.Elapsed.TotalMilliseconds;
            sw.Restart(); var cell = ws.Cells["A1"]; cell.Style.Font.Weight = ExcelFont.BoldWeight; cell.Style.NumberFormat = "0.00"; t[4] = sw.Elapsed.TotalMilliseconds;
            sw.Restart(); ws.PrintOptions.PaperType = PaperType.A4; ws.PrintOptions.Portrait = false; ws.PrintOptions.FitToPage = true;
            ws.PrintOptions.FitWorksheetWidthToPages = 1; ws.HeadersFooters.DefaultPage.Header.CenterSection.Append("Phase 3");
            ws.HeadersFooters.DefaultPage.Footer.RightSection.Append(HeaderFooterFieldType.PageNumber); t[5] = sw.Elapsed.TotalMilliseconds;
            sw.Restart(); wb.Calculate(); t[6] = sw.Elapsed.TotalMilliseconds;
            sw.Restart(); var xls = Path.GetExtension(input).Equals(".xls", StringComparison.OrdinalIgnoreCase);
            wb.Save(output, xls ? SaveOptions.XlsDefault : SaveOptions.XlsxDefault); t[7] = sw.Elapsed.TotalMilliseconds;
            if (xls) { sw.Restart(); wb.Save(Path.ChangeExtension(output, ".converted.xlsx")); t[8] = sw.Elapsed.TotalMilliseconds; }
            var verify = ExcelFile.Load(output);
            if (!verify.Worksheets[0].Cells["A1"].StringValue.StartsWith("edit-r", StringComparison.Ordinal)) throw new InvalidOperationException("edit did not round-trip");
            process.Refresh();
            var total = Stopwatch.GetElapsedTime(start).TotalMilliseconds;
            writer.WriteLine(string.Join(',', Csv("gembox"), round, Csv(Path.GetFileName(input)), Csv(Path.GetExtension(input)),
                new FileInfo(input).Length, new FileInfo(output).Length, Num(t[0]), Num(t[1]), Num(t[2]), Num(t[3]), Num(t[4]), Num(t[5]),
                Num(t[6]), Num(t[7]), Num(t[8]), Num(total), Num(process.PeakWorkingSet64 / 1048576.0),
                Num(process.TotalProcessorTime.TotalMilliseconds - cpuStart), "PASS", ""));
        }
        catch (Exception ex)
        {
            process.Refresh();
            writer.WriteLine(string.Join(',', Csv("gembox"), round, Csv(Path.GetFileName(input)), Csv(Path.GetExtension(input)),
                new FileInfo(input).Length, 0, Num(t[0]), Num(t[1]), Num(t[2]), Num(t[3]), Num(t[4]), Num(t[5]), Num(t[6]), Num(t[7]), Num(t[8]),
                Num(Stopwatch.GetElapsedTime(start).TotalMilliseconds), Num(process.PeakWorkingSet64 / 1048576.0),
                Num(process.TotalProcessorTime.TotalMilliseconds - cpuStart), "FAIL", Csv(ex.Message)));
        }
    }

    foreach (var name in new[] { "large-150k.xlsx", "large-1m.xlsx" })
    {
        var input = Path.Combine(inputRoot, name); var output = Path.Combine(outputRoot, name);
        var start = Stopwatch.GetTimestamp(); var cpuStart = process.TotalProcessorTime.TotalMilliseconds; var sw = Stopwatch.StartNew();
        try
        {
            var wb = ExcelFile.Load(input); var openMs = sw.Elapsed.TotalMilliseconds;
            var ws = wb.Worksheets[0]; var used = ws.GetUsedCellRange(false);
            var lastRow = used.LastRowIndex; var value = ws.Cells[lastRow, 1].DoubleValue;
            var readMs = sw.Elapsed.TotalMilliseconds - openMs;
            ws.Cells["A1"].Value = "large-edited";
            sw.Restart(); wb.Save(output); var saveMs = sw.Elapsed.TotalMilliseconds;
            var verify = ExcelFile.Load(output);
            if (verify.Worksheets[0].Cells["A1"].StringValue != "large-edited") throw new InvalidOperationException("large edit did not round-trip");
            process.Refresh();
            writer.WriteLine(string.Join(',', Csv("gembox"), 1, Csv(name), Csv("xlsx-large"), new FileInfo(input).Length, new FileInfo(output).Length,
                Num(openMs), 0, Num(readMs), 0, 0, 0, 0, Num(saveMs), 0, Num(Stopwatch.GetElapsedTime(start).TotalMilliseconds),
                Num(process.PeakWorkingSet64 / 1048576.0), Num(process.TotalProcessorTime.TotalMilliseconds - cpuStart), "PASS", Csv($"lastRow={lastRow};lastB={value}")));
        }
        catch (Exception ex)
        {
            process.Refresh();
            writer.WriteLine(string.Join(',', Csv("gembox"), 1, Csv(name), Csv("xlsx-large"), File.Exists(input) ? new FileInfo(input).Length : 0, 0,
                Num(sw.Elapsed.TotalMilliseconds), 0, 0, 0, 0, 0, 0, 0, 0, Num(Stopwatch.GetElapsedTime(start).TotalMilliseconds),
                Num(process.PeakWorkingSet64 / 1048576.0), Num(process.TotalProcessorTime.TotalMilliseconds - cpuStart), "FAIL", Csv(ex.Message)));
        }
    }
    writer.Flush();
    return 0;
}

static int RunConversion(string inputRoot, string outputRoot, string csvPath)
{
    Directory.CreateDirectory(outputRoot);
    using var writer = new StreamWriter(csvPath, false, System.Text.Encoding.UTF8);
    writer.WriteLine("engine,round,file,input_bytes,output_bytes,open_ms,convert_save_ms,total_ms,status,error");
    var files = Directory.GetFiles(inputRoot, "batch-*.xls").OrderBy(p => p).ToArray();
    if (files.Length != 8) throw new InvalidOperationException($"expected 8 XLS fixtures, found {files.Length}");
    var warm = ExcelFile.Load(files[0]);
    warm.Save(Path.Combine(outputRoot, "warmup.xlsx"));
    Console.WriteLine("WARMUP_COMPLETE");
    for (var round = 1; round <= 3; round++)
    foreach (var input in files)
    {
        var output = Path.Combine(outputRoot, $"r{round}-{Path.GetFileNameWithoutExtension(input)}.xlsx");
        var total = Stopwatch.StartNew(); var sw = Stopwatch.StartNew();
        try
        {
            var workbook = ExcelFile.Load(input);
            var openMs = sw.Elapsed.TotalMilliseconds;
            sw.Restart(); workbook.Save(output); var convertMs = sw.Elapsed.TotalMilliseconds;
            var check = ExcelFile.Load(output);
            if (check.Worksheets[0].Cells["A2"].StringValue != "row-1") throw new InvalidOperationException("conversion readback mismatch at A2");
            writer.WriteLine(string.Join(',', Csv("gembox"), round, Csv(Path.GetFileName(input)), new FileInfo(input).Length,
                new FileInfo(output).Length, Num(openMs), Num(convertMs), Num(total.Elapsed.TotalMilliseconds), "PASS", ""));
        }
        catch (Exception ex)
        {
            writer.WriteLine(string.Join(',', Csv("gembox"), round, Csv(Path.GetFileName(input)), new FileInfo(input).Length,
                0, 0, 0, Num(total.Elapsed.TotalMilliseconds), "FAIL", Csv(ex.Message)));
        }
    }
    return 0;
}

static ExcelFile Load(string path) => ExcelFile.Load(path);

static void RoundTrip(string input, string noOp, string edited, bool xls)
{
    var wb = Load(input);
    var ws = wb.Worksheets[0];
    var before = ws.Cells["B3"].DoubleValue;
    wb.Save(noOp, xls ? SaveOptions.XlsDefault : SaveOptions.XlsxDefault);
    var saved = Load(noOp);
    Need(Math.Abs(saved.Worksheets[0].Cells["B3"].DoubleValue - before) < 1e-9, "no-op save changed B3");
    if (!xls)
    {
        var copy = saved.Worksheets[0];
        Need(copy.Cells["H1"].MergedRange != null, "no-op XLSX save lost merge");
        Need(copy.Charts.Count == ws.Charts.Count, "no-op XLSX save changed chart count");
        Need(copy.DataValidations.Count == ws.DataValidations.Count, "no-op XLSX save changed validations");
        Need(copy.Cells["A2"].HasComment, "no-op XLSX save lost comment");
        Need(copy.Cells["A3"].HasHyperlink, "no-op XLSX save lost hyperlink");
    }
    ws.Cells["B3"].Value = before + 1;
    wb.Save(edited, xls ? SaveOptions.XlsDefault : SaveOptions.XlsxDefault);
    Need(Math.Abs(Load(edited).Worksheets[0].Cells["B3"].DoubleValue - (before + 1)) < 1e-9,
        "edit did not round-trip");
    Console.WriteLine($"PASS {(xls ? "XLS" : "XLSX")} no-op and one-cell edit");
}

static void CapabilityRun(string input, string output)
{
    var wb = Load(input);
    var ws = wb.Worksheets["Data"];
    // The FREE-LIMITED license permits at most 150 rows. This tests 100 appended rows within that ceiling.
    var firstAppend = 11; // zero-based row 12, after the fixture's 11 used rows.
    ws.Rows.InsertEmpty(firstAppend, 100);
    for (var i = 0; i < 100; i++)
    {
        ws.Cells[firstAppend + i, 0].Value = $"append-{i}";
        ws.Cells[firstAppend + i, 1].Value = i;
    }
    ws.Rows.InsertEmpty(firstAppend + 100, 2);
    ws.Rows.Remove(firstAppend + 100, 2);
    ws.Columns.InsertEmpty(6, 1);
    ws.Columns.Remove(6, 1);

    var cell = ws.Cells["A12"];
    cell.Style.Font.Weight = ExcelFont.BoldWeight;
    cell.Style.Font.Size = 12;
    cell.Style.Font.Color = SpreadsheetColor.FromArgb(0, 0, 139);
    cell.Style.FillPattern.SetSolid(SpreadsheetColor.FromArgb(255, 255, 224));
    cell.Style.WrapText = true;
    cell.Style.NumberFormat = "0.00";
    cell.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
    cell.Style.Borders[IndividualBorder.Bottom].LineStyle = LineStyle.Thin;
    ws.Columns[0].Width = 22 * 256;
    ws.Columns[1].Width = 16 * 256;
    ws.Rows[11].Height = 24 * 20;
    ws.Rows[12].Height = 18 * 20;
    ws.Columns[8].Hidden = true;
    ws.Columns[8].Hidden = false;
    ws.Rows[3].Hidden = false;

    var page = ws.PrintOptions;
    page.PaperType = PaperType.A4;
    page.Portrait = false;
    page.TopMargin = 0.5;
    page.BottomMargin = 0.5;
    page.LeftMargin = 0.4;
    page.RightMargin = 0.4;
    page.HeaderMargin = 0.2;
    page.FooterMargin = 0.2;
    page.FitToPage = true;
    page.FitWorksheetWidthToPages = 1;
    page.FitWorksheetHeightToPages = 0;
    page.HorizontalCentered = true;
    page.VerticalCentered = false;
    ws.NamedRanges.SetPrintArea(ws.GetUsedCellRange(false).GetSubrangeAbsolute(0, 0, 110, 5));
    ws.NamedRanges.SetPrintTitles(ws.Rows[0], 1);
    ws.HeadersFooters.DefaultPage.Header.CenterSection.Append("Phase 2");
    ws.HeadersFooters.DefaultPage.Footer.RightSection.Append(HeaderFooterFieldType.PageNumber);
    ws.HeadersFooters.DefaultPage.Footer.RightSection.Append(" / ");
    ws.HeadersFooters.DefaultPage.Footer.RightSection.Append(HeaderFooterFieldType.NumberOfPages);

    // Independent formula workbook; calculate, edit inputs and formula text, save and inspect cache.
    var calc = new ExcelFile();
    var data = calc.Worksheets.Add("Data");
    data.Cells["A1"].Value = "Key"; data.Cells["B1"].Value = "Value";
    data.Cells["A2"].Value = "a"; data.Cells["B2"].Value = 2;
    data.Cells["A3"].Value = "b"; data.Cells["B3"].Value = 3;
    data.Cells["A4"].Value = "b"; data.Cells["B4"].Value = 4;
    data.NamedRanges.Add("InputValues", data.GetUsedCellRange(false).GetSubrangeAbsolute(1, 1, 3, 1));
    var formulas = calc.Worksheets.Add("Calc");
    var expr = new[] {
        "=SUM(Data!InputValues)", "=SUMIF(Data!A2:A4,\"b\",Data!B2:B4)", "=COUNTIF(Data!B2:B4,\">2\")",
        "=IF(AND(Data!B2>0,Data!B3>0),ROUND(Data!B2*Data!B3,2),0)",
        "=IF(OR(Data!B2=0,Data!B3=3),1,0)", "=VLOOKUP(\"b\",Data!A2:B4,2,FALSE)",
        "=INDEX(Data!B2:B4,MATCH(\"b\",Data!A2:A4,0))", "=DATE(2026,1,2)",
        "=SUM(Data!B2:B4)+Data!B2", "=1/0"
    };
    for (var i = 0; i < expr.Length; i++) formulas.Cells[i, 0].Formula = expr[i];
    calc.Calculate();
    Console.WriteLine($"FORMULA_DEBUG {formulas.Cells["A1"].Formula} value={formulas.Cells["A1"].Value} type={formulas.Cells["A1"].Value?.GetType().FullName}");
    Need(Math.Abs(formulas.Cells["A1"].DoubleValue - 9) < 1e-9, "SUM calculation incorrect");
    Need(Math.Abs(formulas.Cells["A2"].DoubleValue - 7) < 1e-9, "SUMIF calculation incorrect");
    Need(Math.Abs(formulas.Cells["A3"].DoubleValue - 2) < 1e-9, "COUNTIF calculation incorrect");
    Need(Math.Abs(formulas.Cells["A4"].DoubleValue - 6) < 1e-9, "nested IF/AND/ROUND calculation incorrect");
    Need(Math.Abs(formulas.Cells["A5"].DoubleValue - 1) < 1e-9, "OR calculation incorrect");
    Need(Math.Abs(formulas.Cells["A6"].DoubleValue - 3) < 1e-9, "VLOOKUP calculation incorrect");
    Need(Math.Abs(formulas.Cells["A7"].DoubleValue - 3) < 1e-9, "INDEX/MATCH calculation incorrect");
    Need(formulas.Cells["A8"].DateTimeValue.Date == new DateTime(2026, 1, 2), "DATE calculation incorrect");
    Need(Math.Abs(formulas.Cells["A9"].DoubleValue - 11) < 1e-9, "cross-sheet dependency calculation incorrect");
    Need(Convert.ToString(formulas.Cells["A10"].Value) == "#DIV/0!", "formula error type/result mismatch");
    calc.Calculate();
    Need(Math.Abs(formulas.Cells["A1"].DoubleValue - 9) < 1e-9, "repeat full recalculation mismatch");
    data.Rows.InsertEmpty(2, 1);
    data.Cells["A3"].Value = "c"; data.Cells["B3"].Value = 6;
    calc.Calculate();
    Need(Math.Abs(formulas.Cells["A1"].DoubleValue - 15) < 1e-9, "named-range dependency failed after row insertion");
    data.Rows.Remove(2, 1);
    data.Columns.InsertEmpty(4, 1);
    data.Columns.Remove(4, 1);
    calc.Calculate();
    Need(Math.Abs(formulas.Cells["A1"].DoubleValue - 9) < 1e-9, "formula dependency failed after row/column restoration");
    data.Cells["B2"].Value = 5;
    formulas.Cells["A1"].Formula = "=SUM(Data!InputValues)*2";
    calc.Calculate();
    Need(Math.Abs(formulas.Cells["A1"].DoubleValue - 24) < 1e-9, "value/formula recalculation incorrect");
    calc.Save("/output/gembox-formula-check.xlsx");
    var calcReopen = Load("/output/gembox-formula-check.xlsx");
    Need(calcReopen.Worksheets["Calc"].Cells["A1"].Formula.Contains("SUM"), "formula text lost after save");
    Need(Math.Abs(calcReopen.Worksheets["Calc"].Cells["A1"].DoubleValue - 24) < 1e-9, "formula cache lost after save");

    wb.Save(output);
    var check = Load(output);
    var outWs = check.Worksheets["Data"];
    Need(outWs.Cells["A111"].StringValue == "append-99", "append last-row mismatch");
    Need(Math.Abs(outWs.Cells["B111"].DoubleValue - 99) < 1e-9, "append numeric value mismatch");
    var savedStyle = outWs.Cells["A12"].Style;
    Need(savedStyle.Font.Weight == ExcelFont.BoldWeight && Math.Abs(savedStyle.Font.Size - 12) < 1e-9,
        "font did not round-trip");
    Need(savedStyle.Font.Color == SpreadsheetColor.FromArgb(0, 0, 139) &&
         savedStyle.FillPattern.PatternForegroundColor == SpreadsheetColor.FromArgb(255, 255, 224),
        "font/fill color did not round-trip");
    Need(savedStyle.WrapText && savedStyle.NumberFormat == "0.00" &&
         savedStyle.HorizontalAlignment == HorizontalAlignmentStyle.Center,
        "wrap/number/alignment style did not round-trip");
    Need(Math.Abs(outWs.Columns[0].Width / 256.0 - 22) < 0.1, "column width did not round-trip");
    Need(outWs.PrintOptions.PaperType == PaperType.A4 && !outWs.PrintOptions.Portrait, "paper/orientation mismatch");
    Need(outWs.PrintOptions.FitWorksheetWidthToPages == 1, "fit-to-width mismatch");
    Need(Math.Abs(outWs.PrintOptions.LeftMargin - 0.4) < 0.01 &&
         Math.Abs(outWs.PrintOptions.TopMargin - 0.5) < 0.01, "print margins mismatch");
    Need(outWs.PrintOptions.HorizontalCentered && !outWs.PrintOptions.VerticalCentered,
        "print centering mismatch");
    Need(outWs.Cells["H1"].MergedRange != null && outWs.Charts.Count == 1, "existing merge/chart lost during edit");
    Console.WriteLine("PASS append=100 rows (FREE-LIMITED cap) insert/delete rows+columns style dimensions print setup formula recompute");
}

try
{
    if (args.Length == 4 && args[0] == "--convert") return RunConversion(args[1], args[2], args[3]);
    if (args.Length == 4 && args[0] == "--perf") return RunPerformance(args[1], args[2], args[3]);
    RoundTrip("/input/capabilities.xlsx", "/output/gembox-noop.xlsx", "/output/gembox-edited.xlsx", false);
    RoundTrip("/input/capabilities.xls", "/output/gembox-noop.xls", "/output/gembox-edited.xls", true);
    var legacy = Load("/input/capabilities.xls");
    legacy.Save("/output/gembox-converted.xlsx");
    Need(Math.Abs(Load("/output/gembox-converted.xlsx").Worksheets[0].Cells["B3"].DoubleValue - 2) < 1e-9,
        "XLS to XLSX conversion value mismatch");
    Console.WriteLine("PASS XLS to XLSX conversion");
    CapabilityRun("/input/capabilities.xlsx", "/output/gembox-capabilities.xlsx");
    var pdfBook = Load("/input/capabilities.xlsx");
    pdfBook.Calculate();
    FontSettings.FontsBaseDirectory = "/opt/app/fonts";
    FontSettings.FontSelection += (_, e) =>
    {
        if (e.Font is null && e.FontSettings.FontFiles.Count > 0)
            e.Font = e.FontSettings.FontFiles[0];
    };
    pdfBook.Save("/output/gembox-smoke.pdf", SaveOptions.PdfDefault);
    var pdfBytes = File.ReadAllBytes("/output/gembox-smoke.pdf");
    Need(pdfBytes.Length > 8 && System.Text.Encoding.ASCII.GetString(pdfBytes, 0, 5) == "%PDF-",
        "PDF output signature missing");
    Console.WriteLine($"PASS PDF bytes={pdfBytes.Length}");
}
catch (Exception ex)
{
    Console.Error.WriteLine($"FAIL {ex}");
    return 1;
}
return 0;
