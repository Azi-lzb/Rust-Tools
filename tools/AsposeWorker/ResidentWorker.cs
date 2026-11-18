using Aspose.Cells;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;

internal static class ResidentWorker
{
    private const int MaxAttempts = 99;
    private static void Emit(object value) => Console.WriteLine(JsonSerializer.Serialize(value));
    private static string Text(JsonElement e, string key) => e.GetProperty(key).GetString() ?? throw new ArgumentException(key);
    private static void Keys(JsonElement e, params string[] allowed)
    {
        foreach (var p in e.EnumerateObject())
            if (!allowed.Contains(p.Name)) throw new NotSupportedException("Unsupported field: " + p.Name);
    }
    private static void Address(string cell)
    {
        if (!Regex.IsMatch(cell, "^[A-Z]{1,3}[1-9][0-9]{0,6}$")) throw new ArgumentException("Invalid A1 coordinate");
        int column = 0, i = 0;
        while (i < cell.Length && char.IsLetter(cell[i])) column = checked(column * 26 + cell[i++] - 'A' + 1);
        if (column > 16384 || int.Parse(cell[i..], CultureInfo.InvariantCulture) > 1048576)
            throw new ArgumentException("Coordinate out of XLSX bounds");
    }
    private static JsonDocument? OperationFile(JsonElement request, string method)
    {
        if (!request.TryGetProperty("operations_file", out var property)) return null;
        if (method != "Modify") throw new ArgumentException("Operation sidecar is only valid for Modify");
        string path = property.GetString() ?? throw new ArgumentException("Missing operation sidecar path");
        if (!Path.IsPathFullyQualified(path) || new FileInfo(path).Length > 64L * 1024 * 1024) throw new ArgumentException("Invalid or excessive operation sidecar");
        if (request.GetProperty("operations").GetArrayLength() != 0) throw new ArgumentException("Ambiguous operation sources");
        var bytes = File.ReadAllBytes(path);
        if (!Convert.ToHexString(SHA256.HashData(bytes)).Equals(Text(request, "operations_sha256"), StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Operation sidecar hash mismatch");
        return JsonDocument.Parse(bytes);
    }
    private static void CheckOperation(JsonElement op, Workbook wb)
    {
        if (Text(op, "op") == "edit_plan_operation")
        {
            Keys(op, "op", "operation");
            _ = Text(op.GetProperty("operation"), "operation");
            return; // Plan primitives validate sequentially in memory before Save.
        }
        string kind = Text(op, "op"), sheet = Text(op, "sheet");
        if (wb.Worksheets[sheet] is null) throw new ArgumentException("Worksheet not found: " + sheet);
        if (kind == "replace_header_text")
        {
            Keys(op, "op", "sheet", "search", "replacement");
            if (Text(op, "search").Length == 0) throw new ArgumentException("Empty header search");
            _ = Text(op, "replacement");
            return;
        }
        string cell = Text(op, "cell");
        Address(cell);
        if (wb.Worksheets[sheet] is null) throw new ArgumentException("Worksheet not found: " + sheet);
        switch (kind)
        {
            case "set_value":
                Keys(op, "op", "sheet", "cell", "value");
                var value = op.GetProperty("value");
                string type = Text(value, "type");
                Keys(value, "type", "value");
                switch (type)
                {
                    case "text": _ = Text(value, "value"); break;
                    case "number": if (!double.IsFinite(value.GetProperty("value").GetDouble())) throw new ArgumentException("Nonfinite number"); break;
                    case "boolean": _ = value.GetProperty("value").GetBoolean(); break;
                    case "blank": break;
                    default: throw new NotSupportedException("Unsupported cell value type: " + type);
                }
                break;
            case "set_formula":
                Keys(op, "op", "sheet", "cell", "formula");
                if (!Text(op, "formula").StartsWith('=')) throw new ArgumentException("Formula must start with =");
                break;
            case "set_style":
                Keys(op, "op", "sheet", "cell", "style");
                var style = op.GetProperty("style");
                Keys(style, "bold", "italic", "font_name", "font_size", "fill_rgb", "number_format", "wrap_text");
                foreach (var p in style.EnumerateObject())
                {
                    if (p.Name is "bold" or "italic" or "wrap_text") _ = p.Value.GetBoolean();
                    else if (p.Name == "font_size")
                    {
                        int size = p.Value.GetInt32();
                        if (size < 1 || size > 409) throw new ArgumentException("Invalid font size");
                    }
                    else if (p.Name == "fill_rgb")
                    {
                        if (!Regex.IsMatch(p.Value.GetString() ?? "", "^[0-9A-Fa-f]{6}$")) throw new ArgumentException("RGB requires six hex digits");
                    }
                    else _ = p.Value.GetString() ?? throw new ArgumentException("Invalid style string");
                }
                break;
            default: throw new NotSupportedException("Unsupported operation: " + kind);
        }
    }
    private static void Apply(JsonElement op, Workbook wb, string runId, string requestId)
    {
        if (Text(op, "op") == "replace_header_text")
        {
            var setup = wb.Worksheets[Text(op, "sheet")].PageSetup;
            string search = Text(op, "search"), replacement = Text(op, "replacement");
            var sections = new (Func<int, string> Get, Action<int, string> Set)[] {
                (setup.GetHeader, setup.SetHeader),
                (setup.GetEvenHeader, setup.SetEvenHeader),
                (setup.GetFirstPageHeader, setup.SetFirstPageHeader)
            };
            foreach (var section in sections)
                for (int i = 0; i < 3; i++)
                {
                    string old = section.Get(i);
                    if (!string.IsNullOrEmpty(old) && old.Contains(search, StringComparison.Ordinal))
                        section.Set(i, old.Replace(search, replacement, StringComparison.Ordinal));
                }
            return;
        }
        var cell = wb.Worksheets[Text(op, "sheet")].Cells[Text(op, "cell")];
        switch (Text(op, "op"))
        {
            case "set_value":
                var v = op.GetProperty("value");
                switch (Text(v, "type"))
                {
                    case "text": cell.PutValue(Text(v, "value")); break;
                    case "number": cell.PutValue(v.GetProperty("value").GetDouble()); break;
                    case "boolean": cell.PutValue(v.GetProperty("value").GetBoolean()); break;
                    case "blank": cell.PutValue((string?)null); break;
                }
                break;
            case "set_formula": cell.Formula = Text(op, "formula"); break;
            case "set_style":
                var style = cell.GetStyle();
                foreach (var p in op.GetProperty("style").EnumerateObject())
                    switch (p.Name)
                    {
                        case "bold": style.Font.IsBold = p.Value.GetBoolean(); break;
                        case "italic": style.Font.IsItalic = p.Value.GetBoolean(); break;
                        case "font_name": style.Font.Name = p.Value.GetString(); break;
                        case "font_size": style.Font.Size = p.Value.GetInt32(); break;
                        case "number_format": style.Custom = p.Value.GetString(); break;
                        case "wrap_text": style.IsTextWrapped = p.Value.GetBoolean(); break;
                        case "fill_rgb":
                            style.Pattern = BackgroundType.Solid;
                            style.ForegroundColor = Color.FromArgb(int.Parse("FF" + p.Value.GetString(), NumberStyles.HexNumber));
                            break;
                    }
                // Apply only explicitly changed properties, retaining inherited
                // row/column formatting for the rest of the cell style.
                cell.SetStyle(style, true);
                var applied = cell.GetStyle();
                var readback = wb.Worksheets[Text(op, "sheet")].Cells[Text(op, "cell")].GetStyle();
                Console.Error.WriteLine(JsonSerializer.Serialize(new { run_id = runId, request_id = requestId, timestamp = DateTimeOffset.UtcNow.ToString("O"), stage = "style_applied", sheet = Text(op, "sheet"), address = Text(op, "cell"),
                    local_bold = applied.Font.IsBold, stored_bold = readback.Font.IsBold,
                    local_size = applied.Font.Size, stored_size = readback.Font.Size }));
                break;
        }
    }
    public static int Run()
    {
        Emit(new { ready = true, protocol_version = 1, engine = "Aspose.Cells", sdk = CellsHelper.GetVersion(),
                   pid = Environment.ProcessId, evaluation_mode = true, max_attempts = MaxAttempts });
        int attempts = 0;
        string? line;
        while ((line = Console.ReadLine()) is not null)
        {
            string? ownedCandidate = null;
            string? id = null, runId = null;
            string? sanitationStage = null;
            var elapsed = Stopwatch.StartNew();
            try
            {
                if (line.Length > 1024 * 1024) throw new ArgumentException("Request too large");
                using var document = JsonDocument.Parse(line);
                var q = document.RootElement;
                id = Text(q, "request_id"); runId = Text(q, "run_id");
                if (q.GetProperty("protocol_version").GetInt32() != 1) throw new NotSupportedException("Protocol version");
                string method = Text(q, "method");
                if (method == "Ping") { Emit(new { request_id = id, ok = true, pid = Environment.ProcessId, attempts }); continue; }
                if (method != "Calculate" && method != "Modify" && method != "Convert") throw new NotSupportedException("Unsupported method");
                Keys(q, "protocol_version", "request_id", "run_id", "method", "input", "candidate", "operations", "operations_file", "operations_sha256", "verify_content");
                string input = Text(q, "input"), output = Text(q, "candidate");
                if (!Path.IsPathFullyQualified(input) || !Path.IsPathFullyQualified(output)) throw new ArgumentException("Absolute paths required");
                input = Path.GetFullPath(input); output = Path.GetFullPath(output);
                if (string.Equals(input, output, OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal))
                    throw new ArgumentException("Input and candidate must differ");
                if (!new[] { ".xls", ".xlsx" }.Contains(Path.GetExtension(input).ToLowerInvariant()) || Path.GetExtension(output).ToLowerInvariant() != ".xlsx")
                    throw new NotSupportedException("Only XLS/XLSX input and XLSX candidate supported");
                if (File.Exists(output) || Directory.Exists(output)) throw new IOException("Candidate already exists");
                using var operationDocument = OperationFile(q, method);
                var requestedOperations = method == "Modify" ? operationDocument?.RootElement ?? q.GetProperty("operations") : default;
                var importPaths = method == "Modify" ? PlanOperations.ImportPaths(requestedOperations) : Array.Empty<string>();
                if (1 + importPaths.Length > MaxAttempts - attempts) throw new NotSupportedException("Request exceeds remaining workbook load budget; host must rotate before dispatch");
                attempts++; // Count failed loads too; never replay a partially executed request.
                using var wb = new Workbook(input);
                using var plans = new PlanOperations();
                if (method == "Modify" && PlanOperations.NeedsInputValues(requestedOperations)) plans.SetInput(input, wb);
                attempts += importPaths.Length;
                plans.LoadImports(importPaths);
                if (wb.HasMacro) throw new NotSupportedException("Macro preservation has not been validated");
                var sourceSheets = wb.Worksheets.Cast<Worksheet>().Select(s => s.Name).ToArray();
                int sourceActive = wb.Worksheets.ActiveSheetIndex;
                FormulaCachePolicy.PreserveStoredStyles(wb);
                var invalidated = method == "Modify" ? FormulaCachePolicy.Invalidated(wb, requestedOperations) : null;
                if (invalidated?.Count > 0)
                {
                    wb.Settings.FormulaSettings.CalculateOnOpen = true;
                    wb.Settings.FormulaSettings.ForceFullCalculation = true;
                }
                double loadMs = elapsed.Elapsed.TotalMilliseconds;
                if (method == "Modify")
                {
                    var operations = requestedOperations;
                    if (operations.ValueKind != JsonValueKind.Array || operations.GetArrayLength() == 0) throw new ArgumentException("Nonempty operations required");
                    foreach (var op in operations.EnumerateArray())
                    {
                        // A later write may target a worksheet created earlier
                        // in this same plan. Validate against the sequential
                        // in-memory state; Save still follows all operations.
                        CheckOperation(op, wb);
                        if (Text(op, "op") == "edit_plan_operation") plans.Apply(op.GetProperty("operation"), wb);
                        else Apply(op, wb, runId!, id!);
                    }
                }
                else wb.CalculateFormula();
                double transformMs = elapsed.Elapsed.TotalMilliseconds - loadMs;
                // Snapshot task-owned business sheets before SDK serialization.
                // Empty templates remain business sheets. Keep the original
                // active sheet's identity across rename/move/delete operations.
                sanitationStage = "output_sanitation_before_save";
                var identities = method == "Modify" ? PlanOperations.FinalIdentities(sourceSheets, requestedOperations) : sourceSheets.ToDictionary(s => s, s => s);
                string? preferredActive = sourceActive >= 0 && sourceActive < sourceSheets.Length && identities.TryGetValue(sourceSheets[sourceActive], out var activeName) ? activeName : null;
                string[] expectedBusiness = OutputSanitation.BusinessRoster(sourceSheets, requestedOperations);
                var removedBeforeSave = new List<string>();
                foreach (var sheet in wb.Worksheets.Cast<Worksheet>().Where(s => !expectedBusiness.Contains(s.Name)).ToArray())
                {
                    if (sheet.Name != OutputSanitation.SystemSheet || sheet.Shapes.Count != 0 || sheet.Comments.Count != 0 ||
                        sheet.ListObjects.Count != 0 || sheet.Hyperlinks.Count != 0 || sheet.PivotTables.Count != 0 || sheet.Validations.Count != 0)
                        throw new InvalidDataException("Unverified pre-save system worksheet");
                    int notices = 0;
                    foreach (Cell cell in sheet.Cells)
                        if (cell.IsFormula || (cell.Value is not null &&
                            (cell.Name != "A5" || cell.Value is not string text || !OutputSanitation.IsNotice(text) || ++notices > 1)))
                            throw new InvalidDataException("Pre-save system worksheet contains business data");
                    foreach (Worksheet business in wb.Worksheets) foreach (Cell cell in business.Cells)
                        if (cell.IsFormula && OutputSanitation.UnsafeReference(cell.Formula))
                            throw new InvalidDataException("Unsafe pre-save worksheet reference");
                    foreach (Name name in wb.Worksheets.Names)
                        if (OutputSanitation.UnsafeReference(name.RefersTo)) throw new InvalidDataException("Unsafe pre-save defined name");
                    removedBeforeSave.Add(sheet.Name);
                    wb.Worksheets.RemoveAt(sheet.Name);
                }
                var businessSheets = wb.Worksheets.Cast<Worksheet>().ToArray();
                if (!businessSheets.Select(s => s.Name).SequenceEqual(expectedBusiness)) throw new InvalidDataException("Pre-save business roster changed");
                var sanitationPlan = OutputSanitation.Prepare(businessSheets.Select(s => s.Name).ToArray(), businessSheets.Select(s => s.IsVisible).ToArray(), preferredActive);
                wb.Worksheets.ActiveSheetIndex = sanitationPlan.ActiveIndex;
                sanitationStage = null;
                using (var file = new FileStream(output, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                {
                    ownedCandidate = output;
                    wb.Save(file, SaveFormat.Xlsx);
                }
                double saveMs = elapsed.Elapsed.TotalMilliseconds - loadMs - transformMs;
                // Preserve exact SDK Save bytes before the existing preservation passes.
                var rawArchiveStart = elapsed.Elapsed.TotalMilliseconds;
                var sdkSaveArchive = RawOutputArchive.Capture(output, input, id!, runId!, method);
                var rawSdkArchiveMs = elapsed.Elapsed.TotalMilliseconds - rawArchiveStart;
                sanitationStage = "output_sanitation_after_save";
                var savedObjects = wb.Worksheets.Cast<Worksheet>().ToArray();
                if (savedObjects.Length < businessSheets.Length || savedObjects.Length > businessSheets.Length + 1 ||
                    businessSheets.Where((sheet, i) => !ReferenceEquals(sheet, savedObjects[i]) || sheet.Name != sanitationPlan.Sheets[i]).Any())
                    throw new InvalidDataException("SDK Save did not preserve business worksheet identities and order");
                string? appendedSheet = savedObjects.Length > businessSheets.Length ? savedObjects[^1].Name : null;
                if (appendedSheet is not null && (wb.IsLicensed || businessSheets.Any(sheet => ReferenceEquals(sheet, savedObjects[^1]))))
                    throw new InvalidDataException("SDK Save appended a worksheet without verifiable evaluation ownership");
                sanitationStage = null;
                var preservation = Stopwatch.StartNew();
                int preservedDefaults = method == "Modify" ? WorksheetDefaults.Preserve(input, output, PlanOperations.FinalIdentities(sourceSheets, requestedOperations)) : 0;
                if (method == "Modify") FormulaPreservation.Preserve(input, output, requestedOperations, invalidated: invalidated);
                if (method == "Modify") PlanSerialization.Apply(output, requestedOperations);
                if (method == "Modify") FontPreservation.Preserve(input, output, requestedOperations);
                if (method == "Convert")
                {
                    WorksheetDefaults.Preserve(input, output);
                    using var empty = JsonDocument.Parse("[]");
                    FormulaPreservation.Preserve(input, output, empty.RootElement, preserveCaches: false);
                    FontPreservation.Preserve(input, output, empty.RootElement);
                }
                if (method == "Calculate")
                {
                    WorksheetDefaults.Preserve(input, output);
                    using var empty = JsonDocument.Parse("[]");
                    FontPreservation.Preserve(input, output, empty.RootElement);
                }
                // Save can append an SDK notice and override the active view.
                // The exact SDK bytes above remain immutable evidence.
                sanitationStage = "output_sanitation";
                var sanitation = OutputSanitation.Clean(output, sanitationPlan, appendedSheet);
                Console.Error.WriteLine(JsonSerializer.Serialize(new { run_id = runId, request_id = id, timestamp = DateTimeOffset.UtcNow.ToString("O"),
                    stage = "output_sanitation", status = "validated", candidate = output, removed_before_save = removedBeforeSave, details = sanitation }));
                sanitationStage = null;
                double preserveDefaultsMs = preservation.Elapsed.TotalMilliseconds;
                using var zip = ZipFile.OpenRead(output);
                if (method == "Modify" && q.TryGetProperty("verify_content", out var verifyContent)
                    && verifyContent.ValueKind == JsonValueKind.True)
                    SavedStyleVerifier.Verify(zip, requestedOperations);
                using var xmlStream = zip.GetEntry("xl/workbook.xml")!.Open();
                var xml = XDocument.Load(xmlStream);
                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                string[] savedSheets = xml.Descendants(ns + "sheet").Select(s => (string)s.Attribute("name")!).ToArray();
                int active = int.Parse((string?)xml.Descendants(ns + "workbookView").FirstOrDefault()?.Attribute("activeTab") ?? "0");
                Emit(new { request_id = id, ok = true, candidate = output, evaluation_output = !wb.IsLicensed,
                           output_sanitation = sanitation,
                           sdk_save_archive = sdkSaveArchive,
                           raw_sdk_archive_ms = rawSdkArchiveMs,
                           source_sheets = sourceSheets, source_active_index = sourceActive, output_sheets = savedSheets,
                           output_active_index = active, load_ms = loadMs, transform_ms = transformMs, save_ms = saveMs,
                           preserve_defaults_ms = preserveDefaultsMs, preserved_default_sheets = preservedDefaults,
                           processing_ms = elapsed.Elapsed.TotalMilliseconds, attempts, pid = Environment.ProcessId, draining = attempts >= MaxAttempts });
                ownedCandidate = null; // The successful response has handed ownership to the host.
                Console.Error.WriteLine(JsonSerializer.Serialize(new { run_id = runId, request_id = id, timestamp = DateTimeOffset.UtcNow.ToString("O"), stage = method, input, candidate = output,
                    status = "candidate_saved", attempts, elapsed_ms = elapsed.Elapsed.TotalMilliseconds, evaluation_output = !wb.IsLicensed }));
            }
            catch (Exception ex)
            {
                if (sanitationStage is not null)
                    Console.Error.WriteLine(JsonSerializer.Serialize(new { run_id = runId, request_id = id,
                        timestamp = DateTimeOffset.UtcNow.ToString("O"), stage = sanitationStage,
                        status = "failed", error_code = "OUTPUT_SANITATION_FAILED", message = ex.Message }));
                // A post-save preservation failure must not leave a plausible
                // result behind. Never delete a pre-existing candidate: this
                // path is owned only after CreateNew successfully opened it.
                if (ownedCandidate is not null)
                {
                    try { File.Delete(ownedCandidate); }
                    catch (Exception cleanup) { Console.Error.WriteLine(JsonSerializer.Serialize(new { run_id = runId, request_id = id, stage = "failed_candidate_cleanup", exception = cleanup.ToString() })); }
                }
                Emit(new { request_id = id, ok = false, code = ex is NotSupportedException ? "UNSUPPORTED" : ex is InvalidDataException ? "OUTPUT_VALIDATION_FAILED" : "ENGINE_ERROR",
                           message = ex.Message, attempts, pid = Environment.ProcessId, draining = attempts >= MaxAttempts });
                Console.Error.WriteLine(JsonSerializer.Serialize(new { run_id = runId, request_id = id, timestamp = DateTimeOffset.UtcNow.ToString("O"), status = "failed", exception = ex.ToString() }));
            }
            if (attempts >= MaxAttempts) return 0;
        }
        return 0;
    }
}
