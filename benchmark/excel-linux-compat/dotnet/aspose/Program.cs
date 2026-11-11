using Aspose.Cells;
using System.Drawing;
using System.Diagnostics;
using System.Globalization;
using System.Text.Json;

static void Need(bool ok, string message)
{
    if (!ok) throw new InvalidOperationException(message);
}

static int RunWorker()
{
    string? line;
    while ((line = Console.ReadLine()) is not null)
    {
        string requestId = "unknown";
        try
        {
            using var request = JsonDocument.Parse(line);
            var root = request.RootElement;
            requestId = root.GetProperty("request_id").GetString() ?? "unknown";
            if (root.TryGetProperty("op", out var op) && op.GetString() == "ping")
            {
                Console.WriteLine(JsonSerializer.Serialize(new { request_id = requestId, ok = true, worker = "aspose-dotnet" }));
                continue;
            }
            var total = Stopwatch.StartNew();
            var input = Path.GetFullPath(root.GetProperty("input_path").GetString()!);
            var output = Path.GetFullPath(root.GetProperty("output_path").GetString()!);
            if (string.Equals(input, output, OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal))
                throw new InvalidOperationException("input_path and output_path must differ");
            var workbook = new Workbook(input);
            var sheet = workbook.Worksheets[0];
            var row = sheet.Cells.MaxDataRow + 1;
            var rows = root.GetProperty("append_rows");
            foreach (var item in rows.EnumerateArray())
            {
                sheet.Cells[row, 0].PutValue(item[0].GetString());
                sheet.Cells[row, 1].PutValue(item[1].GetDouble());
                row++;
            }
            var temp = output + ".candidate-" + Guid.NewGuid().ToString("N");
            try
            {
                workbook.Save(temp, SaveFormat.Xlsx);
                File.Move(temp, output, true);
            }
            finally { if (File.Exists(temp)) File.Delete(temp); }
            var verify = new Workbook(output);
            var last = verify.Worksheets[0].Cells.MaxDataRow;
            if (rows.GetArrayLength() > 0 && !verify.Worksheets[0].Cells[last, 0].StringValue.StartsWith("batch-", StringComparison.Ordinal))
                throw new InvalidOperationException("output readback did not match appended rows");
            Console.WriteLine(JsonSerializer.Serialize(new { request_id = requestId, ok = true,
                appended_rows = rows.GetArrayLength(), worker_total_ms = total.Elapsed.TotalMilliseconds }));
        }
        catch (Exception ex)
        {
            Console.WriteLine(JsonSerializer.Serialize(new { request_id = requestId, ok = false, error = ex.Message }));
        }
    }
    return 0;
}

static int RunPerformance(string inputRoot, string outputRoot, string csvPath)
{
    Directory.CreateDirectory(outputRoot);
    using var writer = new StreamWriter(csvPath, false, System.Text.Encoding.UTF8);
    writer.WriteLine("engine,round,file,format,input_bytes,output_bytes,open_ms,append_ms,edit_ms,delete_rows_cols_ms,style_ms,print_setup_ms,calculate_ms,save_ms,convert_ms,total_ms,peak_memory_mb,cpu_ms,status,error");
    var process = Process.GetCurrentProcess();
    var files = Directory.GetFiles(inputRoot).Where(p => Path.GetFileName(p).StartsWith("batch-", StringComparison.Ordinal)).OrderBy(p => p).ToArray();
    if (files.Length > 0)
    {
        var warm = new Workbook(files[0]);
        warm.CalculateFormula();
        warm.Save(Path.Combine(outputRoot, "warmup.xls"), SaveFormat.Excel97To2003);
        Console.WriteLine("WARMUP_COMPLETE");
    }
    for (var round = 1; round <= 3; round++)
    foreach (var input in files)
    {
        var phases = new double[10];
        var start = Stopwatch.GetTimestamp();
        var cpuStart = process.TotalProcessorTime.TotalMilliseconds;
        string output = Path.Combine(outputRoot, $"r{round}-{Path.GetFileName(input)}");
        try
        {
            var sw = Stopwatch.StartNew();
            var wb = new Workbook(input);
            var ws = wb.Worksheets[0];
            phases[0] = sw.Elapsed.TotalMilliseconds;
            sw.Restart();
            var row = ws.Cells.MaxDataRow + 1;
            for (var i = 0; i < 5; i++) { ws.Cells[row + i, 0].PutValue($"perf-{i}"); ws.Cells[row + i, 1].PutValue(i); }
            phases[1] = sw.Elapsed.TotalMilliseconds;
            sw.Restart(); ws.Cells["A1"].PutValue($"edit-r{round}"); phases[2] = sw.Elapsed.TotalMilliseconds;
            sw.Restart(); ws.Cells.InsertRows(row + 5, 1); ws.Cells.DeleteRows(row + 5, 1); ws.Cells.InsertColumns(10, 1); ws.Cells.DeleteColumns(10, 1, true); phases[3] = sw.Elapsed.TotalMilliseconds;
            sw.Restart(); var style = ws.Cells["A1"].GetStyle(); style.Font.IsBold = true; style.Number = 2; ws.Cells["A1"].SetStyle(style); phases[4] = sw.Elapsed.TotalMilliseconds;
            sw.Restart(); ws.PageSetup.PaperSize = PaperSizeType.PaperA4; ws.PageSetup.Orientation = PageOrientationType.Landscape;
            ws.PageSetup.FitToPagesWide = 1; ws.PageSetup.SetHeader(1, "Phase 3"); ws.PageSetup.SetFooter(2, "&P"); phases[5] = sw.Elapsed.TotalMilliseconds;
            sw.Restart(); wb.CalculateFormula(); phases[6] = sw.Elapsed.TotalMilliseconds;
            sw.Restart(); var xls = Path.GetExtension(input).Equals(".xls", StringComparison.OrdinalIgnoreCase);
            wb.Save(output, xls ? SaveFormat.Excel97To2003 : SaveFormat.Xlsx); phases[7] = sw.Elapsed.TotalMilliseconds;
            if (xls) { sw.Restart(); wb.Save(Path.ChangeExtension(output, ".converted.xlsx"), SaveFormat.Xlsx); phases[8] = sw.Elapsed.TotalMilliseconds; }
            if (Path.GetFileName(input) is "batch-01.xls" or "batch-09.xlsx")
            {
                var verify = new Workbook(output);
                if (!verify.Worksheets[0].Cells["A1"].StringValue.StartsWith("edit-r", StringComparison.Ordinal))
                    throw new InvalidOperationException("sample edit did not round-trip");
            }
            var total = Stopwatch.GetElapsedTime(start).TotalMilliseconds;
            process.Refresh();
            writer.WriteLine(string.Join(',', "aspose", round, Path.GetFileName(input), Path.GetExtension(input), new FileInfo(input).Length,
                new FileInfo(output).Length, phases[0], phases[1], phases[2], phases[3], phases[4], phases[5], phases[6], phases[7], phases[8],
                total, process.PeakWorkingSet64 / 1048576.0, process.TotalProcessorTime.TotalMilliseconds - cpuStart, "PASS", ""));
        }
        catch (Exception ex)
        {
            writer.WriteLine(string.Join(',', "aspose", round, Path.GetFileName(input), Path.GetExtension(input), new FileInfo(input).Length,
                0, phases[0], phases[1], phases[2], phases[3], phases[4], phases[5], phases[6], phases[7], phases[8],
                Stopwatch.GetElapsedTime(start).TotalMilliseconds, process.PeakWorkingSet64 / 1048576.0,
                process.TotalProcessorTime.TotalMilliseconds - cpuStart, "FAIL", '"' + ex.Message.Replace('"', '\'') + '"'));
        }
    }

    foreach (var name in new[] { "large-150k.xlsx", "large-1m.xlsx" })
    {
        var input = Path.Combine(inputRoot, name);
        var output = Path.Combine(outputRoot, name);
        var start = Stopwatch.GetTimestamp();
        var cpuStart = process.TotalProcessorTime.TotalMilliseconds;
        var sw = Stopwatch.StartNew();
        try
        {
            var wb = new Workbook(input);
            var openMs = sw.Elapsed.TotalMilliseconds;
            var ws = wb.Worksheets[0];
            var lastRow = ws.Cells.MaxDataRow;
            var value = ws.Cells[lastRow, 1].DoubleValue;
            var readMs = sw.Elapsed.TotalMilliseconds - openMs;
            ws.Cells["A1"].PutValue("large-edited");
            sw.Restart(); wb.Save(output, SaveFormat.Xlsx); var saveMs = sw.Elapsed.TotalMilliseconds;
            var verify = new Workbook(output);
            if (verify.Worksheets[0].Cells["A1"].StringValue != "large-edited") throw new InvalidOperationException("large read/write check failed");
            process.Refresh();
            writer.WriteLine(string.Join(',', "aspose", 1, name, "xlsx-large", new FileInfo(input).Length, new FileInfo(output).Length,
                openMs, 0, readMs, 0, 0, 0, 0, saveMs, 0, Stopwatch.GetElapsedTime(start).TotalMilliseconds,
                process.PeakWorkingSet64 / 1048576.0, process.TotalProcessorTime.TotalMilliseconds - cpuStart, "PASS", $"lastRow={lastRow};lastB={value}"));
        }
        catch (Exception ex)
        {
            process.Refresh();
            writer.WriteLine(string.Join(',', "aspose", 1, name, "xlsx-large", new FileInfo(input).Length, 0,
                sw.Elapsed.TotalMilliseconds, 0, 0, 0, 0, 0, 0, 0, 0, Stopwatch.GetElapsedTime(start).TotalMilliseconds,
                process.PeakWorkingSet64 / 1048576.0, process.TotalProcessorTime.TotalMilliseconds - cpuStart, "FAIL", '"' + ex.Message.Replace('"', '\'') + '"'));
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
    var warm = new Workbook(files[0]);
    warm.Save(Path.Combine(outputRoot, "warmup.xlsx"), SaveFormat.Xlsx);
    Console.WriteLine("WARMUP_COMPLETE");
    for (var round = 1; round <= 3; round++)
    foreach (var input in files)
    {
        var output = Path.Combine(outputRoot, $"r{round}-{Path.GetFileNameWithoutExtension(input)}.xlsx");
        var total = Stopwatch.StartNew(); var sw = Stopwatch.StartNew();
        try
        {
            var workbook = new Workbook(input);
            var openMs = sw.Elapsed.TotalMilliseconds;
            sw.Restart(); workbook.Save(output, SaveFormat.Xlsx); var convertMs = sw.Elapsed.TotalMilliseconds;
            var check = new Workbook(output);
            if (check.Worksheets[0].Cells["A2"].StringValue != "row-1") throw new InvalidOperationException("conversion readback mismatch at A2");
            writer.WriteLine(string.Join(',', "aspose", round, Path.GetFileName(input), new FileInfo(input).Length,
                new FileInfo(output).Length, openMs, convertMs, total.Elapsed.TotalMilliseconds, "PASS", ""));
        }
        catch (Exception ex)
        {
            writer.WriteLine(string.Join(',', "aspose", round, Path.GetFileName(input), new FileInfo(input).Length,
                0, 0, 0, total.Elapsed.TotalMilliseconds, "FAIL", '"' + ex.Message.Replace('"', '\'') + '"'));
        }
    }
    return 0;
}

static Workbook Open(string path) => new(path);

static void FormatRoundTrip(string source, string copy, string edited, bool xls)
{
    var original = Open(source);
    var sheet = original.Worksheets[0];
    var baselineValue = sheet.Cells["B3"].DoubleValue;
    original.Save(copy, xls ? SaveFormat.Excel97To2003 : SaveFormat.Xlsx);
    var reopenedCopy = Open(copy);
    var reopenedSheet = reopenedCopy.Worksheets[0];
    Need(Math.Abs(reopenedSheet.Cells["B3"].DoubleValue - baselineValue) < 1e-9, "no-op save changed B3");
    if (!xls)
    {
        Need(reopenedSheet.Cells["H1"].IsMerged, "no-op XLSX save lost H1:I1 merge");
        Need(reopenedSheet.Charts.Count == sheet.Charts.Count, "no-op XLSX save changed chart count");
        Need(reopenedSheet.Hyperlinks.Count == sheet.Hyperlinks.Count, "no-op XLSX save changed hyperlink count");
        Need(reopenedSheet.Validations.Count == sheet.Validations.Count, "no-op XLSX save changed validation count");
        Need(reopenedCopy.Worksheets.Names["InputValues"] != null, "no-op XLSX save lost named range");
        Need(reopenedSheet.Cells["A2"].Comment != null, "no-op XLSX save lost comment");
        Need(reopenedCopy.Worksheets["Calc"].Cells["A1"].Formula == original.Worksheets["Calc"].Cells["A1"].Formula,
            "no-op XLSX save changed formula text");
    }
    sheet.Cells["B3"].PutValue(baselineValue + 1.0);
    original.Save(edited, xls ? SaveFormat.Excel97To2003 : SaveFormat.Xlsx);
    var reopenedEdit = Open(edited);
    Need(Math.Abs(reopenedEdit.Worksheets[0].Cells["B3"].DoubleValue - (baselineValue + 1.0)) < 1e-9,
        "edit did not round-trip");
    Console.WriteLine($"PASS {(xls ? "XLS" : "XLSX")} no-op and one-cell edit");
}

static void ModifyXlsx(string input, string output)
{
    var wb = Open(input);
    var ws = wb.Worksheets[0];
    // Append a 1,000-row batch through the public cell API (not the trial-limited engines).
    for (var i = 0; i < 1000; ++i)
    {
        ws.Cells[11 + i, 0].PutValue($"append-{i}");
        ws.Cells[11 + i, 1].PutValue(i);
    }
    ws.Cells.InsertRows(1011, 2);
    ws.Cells.DeleteRows(1011, 2);
    ws.Cells.InsertColumns(6, 2);
    ws.Cells.DeleteColumns(6, 2, true);

    var styled = ws.Cells["A12"];
    var style = styled.GetStyle();
    style.Font.IsBold = true;
    style.Font.Size = 12;
    style.Font.Color = Color.DarkBlue;
    style.ForegroundColor = Color.LightYellow;
    style.Pattern = BackgroundType.Solid;
    style.IsTextWrapped = true;
    style.Number = 2;
    style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
    style.HorizontalAlignment = TextAlignmentType.Center;
    styled.SetStyle(style);
    ws.Cells.CreateRange(11, 1, 1, 1).CopyStyle(ws.Cells.CreateRange(11, 0, 1, 1));
    ws.Cells.SetColumnWidth(0, 22);
    ws.Cells.SetColumnWidth(1, 16);
    ws.Cells.SetRowHeight(11, 24);
    ws.Cells.SetRowHeight(12, 18);
    ws.Cells.HideColumn(8);
    ws.Cells.UnhideColumn(8, 13.0);
    ws.Cells.HideRow(3);
    ws.Cells.UnhideRow(3, 20.0);

    var page = ws.PageSetup;
    page.PaperSize = PaperSizeType.PaperA4;
    page.Orientation = PageOrientationType.Landscape;
    page.PrintArea = "A1:F1011";
    page.PrintTitleRows = "$1:$1";
    page.PrintTitleColumns = "$A:$A";
    page.LeftMarginInch = 0.4;
    page.RightMarginInch = 0.4;
    page.TopMarginInch = 0.5;
    page.BottomMarginInch = 0.5;
    page.HeaderMarginInch = 0.2;
    page.FooterMarginInch = 0.2;
    page.FitToPagesWide = 1;
    page.FitToPagesTall = 0;
    page.CenterHorizontally = true;
    page.CenterVertically = false;
    page.SetHeader(1, "&BPhase 2");
    page.SetFooter(2, "Page &P of &N");
    page.Order = PrintOrderType.OverThenDown;

    // Compute directly from formulas created through this SDK, not a stale fixture cache.
    var calc = new Workbook();
    var data = calc.Worksheets[0];
    data.Name = "Data";
    data.Cells["A1"].PutValue("Key"); data.Cells["B1"].PutValue("Value");
    data.Cells["A2"].PutValue("a"); data.Cells["B2"].PutValue(2);
    data.Cells["A3"].PutValue("b"); data.Cells["B3"].PutValue(3);
    data.Cells["A4"].PutValue("b"); data.Cells["B4"].PutValue(4);
    calc.Worksheets.Add("Calc");
    var cw = calc.Worksheets["Calc"];
    var inputName = calc.Worksheets.Names.Add("InputValues");
    calc.Worksheets.Names[inputName].RefersTo = "=Data!$B$2:$B$4";
    var formulas = new[] {
        "=SUM(InputValues)", "=SUMIF(Data!A2:A4,\"b\",Data!B2:B4)", "=COUNTIF(Data!B2:B4,\">2\")",
        "=IF(AND(Data!B2>0,Data!B3>0),ROUND(Data!B2*Data!B3,2),0)",
        "=IF(OR(Data!B2=0,Data!B3=3),1,0)", "=VLOOKUP(\"b\",Data!A2:B4,2,FALSE)",
        "=INDEX(Data!B2:B4,MATCH(\"b\",Data!A2:A4,0))", "=DATE(2026,1,2)",
        "=SUM(Data!B2:B4)+Data!B2", "=1/0"
    };
    for (var i = 0; i < formulas.Length; i++) cw.Cells[i, 0].Formula = formulas[i];
    calc.CalculateFormula();
    Need(Math.Abs(cw.Cells["A1"].DoubleValue - 9) < 1e-9, "SUM calculation incorrect");
    Need(Math.Abs(cw.Cells["A2"].DoubleValue - 7) < 1e-9, "SUMIF calculation incorrect");
    Need(Math.Abs(cw.Cells["A3"].DoubleValue - 2) < 1e-9, "COUNTIF calculation incorrect");
    Need(Math.Abs(cw.Cells["A4"].DoubleValue - 6) < 1e-9, "nested IF/AND/ROUND calculation incorrect");
    Need(Math.Abs(cw.Cells["A5"].DoubleValue - 1) < 1e-9, "OR calculation incorrect");
    Need(Math.Abs(cw.Cells["A6"].DoubleValue - 3) < 1e-9, "VLOOKUP calculation incorrect");
    Need(Math.Abs(cw.Cells["A7"].DoubleValue - 3) < 1e-9, "INDEX/MATCH calculation incorrect");
    Need(cw.Cells["A8"].DateTimeValue.Date == new DateTime(2026, 1, 2), "DATE calculation incorrect");
    Need(Math.Abs(cw.Cells["A9"].DoubleValue - 11) < 1e-9, "cross-sheet dependency calculation incorrect");
    Need(cw.Cells["A10"].StringValue == "#DIV/0!", "formula error type/result mismatch");
    calc.CalculateFormula();
    Need(Math.Abs(cw.Cells["A1"].DoubleValue - 9) < 1e-9, "repeat full recalculation mismatch");
    data.Cells.InsertRows(2, 1);
    data.Cells["A3"].PutValue("c"); data.Cells["B3"].PutValue(6);
    calc.CalculateFormula();
    Need(Math.Abs(cw.Cells["A1"].DoubleValue - 15) < 1e-9, "named-range dependency failed after row insertion");
    data.Cells.DeleteRows(2, 1);
    data.Cells.InsertColumns(4, 1);
    data.Cells.DeleteColumns(4, 1, true);
    calc.CalculateFormula();
    Need(Math.Abs(cw.Cells["A1"].DoubleValue - 9) < 1e-9, "formula dependency failed after row/column restoration");
    data.Cells["B2"].PutValue(5);
    cw.Cells["A1"].Formula = "=SUM(InputValues)*2";
    calc.CalculateFormula();
    Need(Math.Abs(cw.Cells["A1"].DoubleValue - 24) < 1e-9, "recalculation after value/formula edit incorrect");
    calc.Save("/output/formula-check.xlsx", SaveFormat.Xlsx);
    var calcReopen = Open("/output/formula-check.xlsx");
    Need(calcReopen.Worksheets["Calc"].Cells["A1"].Formula.Contains("SUM"), "formula text lost after save");
    Need(Math.Abs(calcReopen.Worksheets["Calc"].Cells["A1"].DoubleValue - 24) < 1e-9, "formula cache lost after save");

    wb.Save(output, SaveFormat.Xlsx);
    var check = Open(output);
    var outWs = check.Worksheets[0];
    Need(outWs.Cells["A1011"].StringValue == "append-999", "bulk append last row mismatch");
    Need(Math.Abs(outWs.Cells["B1011"].DoubleValue - 999) < 1e-9, "bulk append numeric value mismatch");
    var checkStyle = outWs.Cells["A12"].GetStyle();
    Need(checkStyle.Font.IsBold && Math.Abs(checkStyle.Font.Size - 12) < 1e-9, "font style did not round-trip");
    Need(checkStyle.Pattern == BackgroundType.Solid && checkStyle.ForegroundColor.ToArgb() == Color.LightYellow.ToArgb(),
        "fill style did not round-trip");
    Need(checkStyle.IsTextWrapped && checkStyle.Number == 2, "wrap/number style did not round-trip");
    Need(checkStyle.Borders[BorderType.BottomBorder].LineStyle == CellBorderType.Thin &&
         checkStyle.HorizontalAlignment == TextAlignmentType.Center, "border/alignment style did not round-trip");
    Need(Math.Abs(outWs.Cells.GetColumnWidth(0) - 22) < 1e-9, "column width did not round-trip");
    Need(Math.Abs(outWs.Cells.GetRowHeight(11) - 24) < 1e-9, "row height did not round-trip");
    Need(outWs.PageSetup.PaperSize == PaperSizeType.PaperA4, "paper size did not round-trip");
    Need(outWs.PageSetup.Orientation == PageOrientationType.Landscape, "orientation did not round-trip");
    Need(Math.Abs(outWs.PageSetup.LeftMarginInch - 0.4) < 0.01 &&
         Math.Abs(outWs.PageSetup.TopMarginInch - 0.5) < 0.01, "page margins did not round-trip");
    Need(outWs.PageSetup.PrintArea.Contains("F1011"), "print area did not round-trip");
    Need(outWs.PageSetup.PrintTitleRows.Contains("1"), "print title rows did not round-trip");
    Need(outWs.PageSetup.PrintTitleColumns.Contains("A"), "print title columns did not round-trip");
    Need(outWs.PageSetup.CenterHorizontally && !outWs.PageSetup.CenterVertically &&
         outWs.PageSetup.FitToPagesWide == 1 && outWs.PageSetup.FitToPagesTall == 0,
        "centering/scale setup did not round-trip");
    Need(outWs.PageSetup.GetFooter(2).Contains("&P"), "footer page marker did not round-trip");
    Need(outWs.Cells["H1"].IsMerged && outWs.Charts.Count == 1, "existing merge/chart lost during edit");
    Console.WriteLine("PASS append=1000 insert/delete rows+columns style dimensions print setup formula recompute");
}

try
{
    if (args.Length == 1 && args[0] == "--worker") return RunWorker();
    if (args.Length == 4 && args[0] == "--convert") return RunConversion(args[1], args[2], args[3]);
    if (args.Length == 4 && args[0] == "--perf") return RunPerformance(args[1], args[2], args[3]);
    FormatRoundTrip("/input/capabilities.xlsx", "/output/aspose-noop.xlsx", "/output/aspose-edited.xlsx", false);
    FormatRoundTrip("/input/capabilities.xls", "/output/aspose-noop.xls", "/output/aspose-edited.xls", true);
    var xls = Open("/input/capabilities.xls");
    xls.Save("/output/aspose-converted.xlsx", SaveFormat.Xlsx);
    Need(Math.Abs(Open("/output/aspose-converted.xlsx").Worksheets[0].Cells["B3"].DoubleValue - 2) < 1e-9,
        "XLS to XLSX conversion value mismatch");
    Console.WriteLine("PASS XLS to XLSX conversion");
    ModifyXlsx("/input/capabilities.xlsx", "/output/aspose-capabilities.xlsx");
    var pdf = Open("/input/capabilities.xlsx");
    pdf.CalculateFormula();
    pdf.Save("/output/aspose-smoke.pdf", SaveFormat.Pdf);
    var pdfBytes = File.ReadAllBytes("/output/aspose-smoke.pdf");
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
