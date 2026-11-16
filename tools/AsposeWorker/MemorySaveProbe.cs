using Aspose.Cells;
using System.Diagnostics;
using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;

internal static class MemorySaveProbe
{
    public static int Run(string[] args)
    {
// Capability gate before implementing the resident protocol/backend. This only
// operates on isolated synthetic inputs and never publishes a business output.
if (args.Length != 4 || args[0] != "--memory-save-probe")
{
    Console.Error.WriteLine("Usage: --memory-save-probe INPUT OUTPUT_DIRECTORY EXPECTED_ACTIVE_INDEX");
    return 64;
}
string input = Path.GetFullPath(args[1]);
string outputDirectory = Path.GetFullPath(args[2]);
int originalActiveIndex = int.Parse(args[3]);
Directory.CreateDirectory(outputDirectory);
bool allClean = true;
var originalSheets = new Dictionary<string, string[]>();
foreach (string operation in new[] { "Calculate", "Modify", "Calculate-Reloaded", "Modify-Reloaded" })
{
    string baseOperation = operation.Split('-')[0];
    bool reloaded = operation.EndsWith("-Reloaded");
    // Reload only a candidate created by this same probe invocation. This is
    // never permission to remove a real input's sheet based on its name.
    string actualInput = reloaded ? Path.Combine(outputDirectory, baseOperation.ToLowerInvariant() + ".xlsx") : input;
    string output = Path.Combine(outputDirectory, operation.ToLowerInvariant() + ".xlsx");
    if (File.Exists(output)) throw new IOException("Refuse existing output");
    using var workbook = new Workbook(actualInput);
    string[] inputSheets = workbook.Worksheets.Cast<Worksheet>().Select(s => s.Name).ToArray();
    if (!reloaded && inputSheets.Contains("Evaluation Warning"))
        throw new InvalidOperationException("Ambiguous input contains Evaluation Warning; refuse deletion");
    if (!reloaded) originalSheets[baseOperation] = inputSheets;
    else inputSheets = originalSheets[baseOperation];
    int loadedActive = workbook.Worksheets.ActiveSheetIndex;
    if (baseOperation == "Calculate") workbook.CalculateFormula();
    else
    {
        var cell = workbook.Worksheets["Data"].Cells["D12"];
        cell.PutValue(9999);
        var style = cell.GetStyle();
        style.Font.IsBold = true;
        cell.SetStyle(style);
    }
    string[] before = workbook.Worksheets.Cast<Worksheet>().Select(s => s.Name).ToArray();
    bool removed = false;
    if (workbook.Worksheets["Evaluation Warning"] is not null)
    {
        workbook.Worksheets.RemoveAt("Evaluation Warning");
        removed = true;
    }
    workbook.Worksheets.ActiveSheetIndex = originalActiveIndex;
    int activeBeforeSave = workbook.Worksheets.ActiveSheetIndex;
    string[] sheetsBeforeSave = workbook.Worksheets.Cast<Worksheet>().Select(s => s.Name).ToArray();
    var timer = Stopwatch.StartNew();
    workbook.Save(output, SaveFormat.Xlsx);
    timer.Stop();

    // Inspect the physical package; do not reopen it using the same library,
    // which could itself add evaluation content during load.
    using var zip = ZipFile.OpenRead(output);
    using var stream = zip.GetEntry("xl/workbook.xml")!.Open();
    var xml = XDocument.Load(stream);
    XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    string[] actualSheets = xml.Descendants(ns + "sheet").Select(s => (string)s.Attribute("name")!).ToArray();
    string? active = (string?)xml.Descendants(ns + "workbookView").FirstOrDefault()?.Attribute("activeTab");
    bool clean = actualSheets.SequenceEqual(inputSheets) && int.TryParse(active, out int index) && index == originalActiveIndex;
    allClean &= clean;
    Console.WriteLine(JsonSerializer.Serialize(new
    {
        sdk = CellsHelper.GetVersion(), licensed = workbook.IsLicensed, operation, input = actualInput, output,
        input_sheets = inputSheets, loaded_active = loadedActive,
        sheets_before_cleanup = before, removed_before_save = removed,
        sheets_before_save = sheetsBeforeSave, active_before_save = activeBeforeSave,
        saved_sheets = actualSheets, saved_active_tab = active,
        expected_active_tab = originalActiveIndex, save_ms = timer.Elapsed.TotalMilliseconds,
        clean_output = clean,
        status = clean ? "PASS" : "FAILED_EVALUATION_OUTPUT_CONTRACT"
    }));
}
return allClean ? 0 : 2;

    }
}
