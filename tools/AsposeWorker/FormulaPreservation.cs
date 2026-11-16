using System.IO.Compression;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;

// Aspose normalizes otherwise unchanged formula text (notably sheet quotes).
// Restore original lexical text only for cells proven unaffected by the plan.
// Modify preserves proven-unaffected caches and clears invalidated ones;
// Convert retains the SDK's recalculated results. No formula evaluation here.
internal static class FormulaPreservation
{
    private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static string T(JsonElement p, string key) => p.GetProperty(key).GetString()!;
    private static XDocument Read(ZipArchive zip, string part)
    {
        using var stream = zip.GetEntry(part)!.Open(); return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }
    private static Dictionary<string, string> Parts(ZipArchive zip)
    {
        var rels = Read(zip, "xl/_rels/workbook.xml.rels").Root!.Elements().ToDictionary(e => (string)e.Attribute("Id")!, e => (string)e.Attribute("Target")!);
        return Read(zip, "xl/workbook.xml").Descendants(S + "sheet").ToDictionary(e => (string)e.Attribute("name")!,
            e => Uri.UnescapeDataString(new Uri(new Uri("http://ooxml/xl/workbook.xml"), rels[(string)e.Attribute(R + "id")!]).AbsolutePath.TrimStart('/')));
    }
    private static void Save(ZipArchive zip, string part, XDocument document)
    {
        zip.GetEntry(part)!.Delete(); using var stream = zip.CreateEntry(part, CompressionLevel.Optimal).Open(); document.Save(stream, SaveOptions.DisableFormatting);
    }
    private static bool Covered(string address, JsonElement p)
    {
        string kind = T(p, "operation");
        if (kind is "set_cell_value" or "set_formula_text") return T(p, "cell") == address;
        if (kind is not "copy_range" and not "import_range") return false;
        if (!p.GetProperty("copy_values").GetBoolean() && !p.GetProperty("copy_formulas").GetBoolean()) return false;
        var range = T(p, "source_range").Replace("$", "").Split(':');
        Aspose.Cells.CellsHelper.CellNameToIndex(range[0], out int r, out int c);
        Aspose.Cells.CellsHelper.CellNameToIndex(range[^1], out int z, out int x);
        Aspose.Cells.CellsHelper.CellNameToIndex(T(p, "target_start").Replace("$", ""), out int tr, out int tc);
        Aspose.Cells.CellsHelper.CellNameToIndex(address, out int ar, out int ac);
        return ar >= tr && ar <= tr + z - r && ac >= tc && ac <= tc + x - c;
    }
    private static string RenameReferences(string formula, string oldName, string newName)
    {
        // String literals are opaque; INDIRECT cannot be safely renamed.
        if (Regex.IsMatch(formula, "(?i)\\bINDIRECT\\s*\\(")) throw new NotSupportedException("Dynamic sheet references cannot be safely renamed");
        string replacement = "'" + newName.Replace("'", "''") + "'!";
        string quoted = "'" + oldName.Replace("'", "''") + "'!";
        return Regex.Replace(formula, "\"(?:[^\"]|\"\")*\"|" + Regex.Escape(quoted) + "|(?<![A-Za-z0-9_.'\\]])" + Regex.Escape(oldName) + "!",
            m => m.Value.StartsWith('"') ? m.Value : replacement);
    }
    public static void Preserve(string source, string candidate, JsonElement operations, bool preserveCaches = true, HashSet<(string Sheet, string Cell)>? invalidated = null)
    {
        if (!Path.GetExtension(source).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)) return;
        var ops = operations.EnumerateArray().ToArray();
        using var input = ZipFile.OpenRead(source); using var output = ZipFile.Open(candidate, ZipArchiveMode.Update);
        var outputParts = Parts(output);
        foreach (var (sourceSheet, part) in Parts(input))
        {
            string sheet = sourceSheet;
            var originals = Read(input, part).Descendants(S + "c").Where(c => c.Element(S + "f") is not null)
                .Select(c => (Address: (string)c.Attribute("r")!, Formula: new XElement(c.Element(S + "f")!), Value: c.Element(S + "v") is { } v ? new XElement(v) : null, Type: (string?)c.Attribute("t"))).ToArray();
            var expected = new Dictionary<string, XElement>();
            var caches = new Dictionary<string, (XElement? Value, string? Type)>();
            foreach (var original in originals)
            {
                string name = sheet, address = original.Address; var formula = new XElement(original.Formula); bool retained = true;
                if (formula.Attributes().Any(a => a.Name.LocalName is "t" or "si" or "ref"))
                    throw new NotSupportedException("Shared/array formula lexical preservation is not validated");
                foreach (var outer in ops)
                {
                    string outerKind = T(outer, "op");
                    if (outerKind != "edit_plan_operation")
                    {
                        if (outerKind is "set_value" or "set_formula" && T(outer, "sheet") == name && T(outer, "cell") == address) { retained = false; break; }
                        continue;
                    }
                    var p = outer.GetProperty("operation"); string kind = T(p, "operation");
                    if (kind == "rename_sheet") { formula.Value = RenameReferences(formula.Value, T(p, "old_name"), T(p, "new_name")); if (name == T(p, "old_name")) name = T(p, "new_name"); }
                    if (kind == "delete_sheet" && name == T(p, "name")) { retained = false; break; }
                    string target = p.TryGetProperty("sheet", out var targetSheet) ? targetSheet.GetString()! : p.TryGetProperty("target_sheet", out var t) ? t.GetString()! : "";
                    if (target != name) continue;
                    if (Covered(address, p)) { retained = false; break; }
                    if (kind is "delete_rows" or "insert_rows" or "delete_columns" or "insert_columns")
                    {
                        Aspose.Cells.CellsHelper.CellNameToIndex(address, out int row, out int col);
                        bool rows = kind.EndsWith("rows"); int start = p.GetProperty(rows ? "first_row" : "first_column").GetInt32() - 1, count = p.GetProperty("count").GetInt32();
                        int index = rows ? row : col;
                        if (kind.StartsWith("delete") && index >= start && index < start + count) { retained = false; break; }
                        if (index >= start) index += kind.StartsWith("delete") ? -count : count;
                        address = Aspose.Cells.CellsHelper.CellIndexToName(rows ? index : row, rows ? col : index);
                    }
                }
                if (retained)
                {
                    expected[address] = formula;
                    if (preserveCaches) caches[address] = invalidated?.Contains((sourceSheet, original.Address)) == true ? (null, null) : (original.Value, original.Type);
                }
            }
            var identities = PlanOperations.FinalIdentities(new[] { sourceSheet }, operations);
            if (!identities.TryGetValue(sourceSheet, out string? finalName) || expected.Count == 0) continue;
            var document = Read(output, outputParts[finalName]);
            // Index once: scanning the entire sheet for every formula is
            // quadratic on the complex-workbook probe. Duplicate addresses
            // remain an error rather than silently choosing one cell.
            var savedCells = document.Descendants(S + "c").ToDictionary(c => (string)c.Attribute("r")!, StringComparer.Ordinal);
            foreach (var (address, formula) in expected)
            {
                if (!savedCells.TryGetValue(address, out var cell))
                    throw new InvalidDataException("Saved formula cell missing: " + finalName + "!" + address);
                var actual = cell.Element(S + "f") ?? throw new InvalidDataException("Saved formula missing");
                actual.ReplaceWith(formula);
                if (caches.TryGetValue(address, out var cached))
                {
                    cell.Element(S + "v")?.Remove();
                    cell.SetAttributeValue("t", cached.Type);
                    if (cached.Value is not null) cell.Add(new XElement(cached.Value));
                }
            }
            Save(output, outputParts[finalName], document);
        }
        // Defined names also have lexical preservation requirements. Preserve
        // existing names except explicitly deleted/replaced ones; rename refs.
        var sourceBook = Read(input, "xl/workbook.xml"); var outputBook = Read(output, "xl/workbook.xml");
        var inputNames = sourceBook.Root!.Element(S + "definedNames"); var outputNames = outputBook.Root!.Element(S + "definedNames");
        if (inputNames is not null && outputNames is not null)
        {
            foreach (var original in inputNames.Elements())
            {
                if ((string?)original.Attribute("localSheetId") is not null) continue; // scope/order validation stays authoritative in Rust
                string name = (string)original.Attribute("name")!; string formula = original.Value; bool changed = false;
                foreach (var outer in ops)
                {
                    if (T(outer, "op") != "edit_plan_operation") continue;
                    var p = outer.GetProperty("operation"); string kind = T(p, "operation");
                    if (kind == "set_defined_name" && T(p, "name") == name) { changed = true; break; }
                    if (kind == "rename_sheet") formula = RenameReferences(formula, T(p, "old_name"), T(p, "new_name"));
                    if (kind.StartsWith("delete_") || kind.StartsWith("insert_")) { changed = true; break; }
                }
                if (!changed)
                {
                    var actual = outputNames.Elements().SingleOrDefault(n => (string?)n.Attribute("name") == name && n.Attribute("localSheetId") is null)
                        ?? throw new InvalidDataException("Saved defined name missing: " + name);
                    actual.Value = formula;
                }
            }
        }
        foreach (var outer in ops)
        {
            if (T(outer, "op") != "edit_plan_operation") continue;
            var p = outer.GetProperty("operation");
            if (T(p, "operation") != "set_defined_name" || !p.TryGetProperty("formula", out var f) || f.ValueKind == JsonValueKind.Null || p.GetProperty("local_sheet").ValueKind != JsonValueKind.Null) continue;
            var actual = outputBook.Descendants(S + "definedName").Single(n => (string?)n.Attribute("name") == T(p, "name") && n.Attribute("localSheetId") is null);
            actual.Value = f.GetString()!.TrimStart('=');
        }
        Save(output, "xl/workbook.xml", outputBook);
    }
}
