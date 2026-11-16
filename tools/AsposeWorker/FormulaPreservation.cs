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
    private static (int R1, int C1, int R2, int C2) Bounds(string range)
    {
        var cells = range.Replace("$", "").Split(':');
        if (cells.Length is < 1 or > 2) throw new NotSupportedException("Invalid grouped formula range");
        Aspose.Cells.CellsHelper.CellNameToIndex(cells[0], out int r, out int c);
        Aspose.Cells.CellsHelper.CellNameToIndex(cells[^1], out int z, out int x);
        if (r < 0 || c < 0 || z < r || x < c || z >= 1048576 || x >= 16384)
            throw new NotSupportedException("Invalid grouped formula bounds");
        return (r, c, z, x);
    }
    private static bool WritesInto(string sheet, string range, JsonElement outer)
    {
        string kind = T(outer, "op"); var p = outer;
        if (kind == "edit_plan_operation") { p = p.GetProperty("operation"); kind = T(p, "operation"); }
        string target = p.TryGetProperty("sheet", out var s) ? s.GetString()! : p.TryGetProperty("target_sheet", out s) ? s.GetString()! : "";
        if (target != sheet) return false;
        var group = Bounds(range);
        if (kind is "set_value" or "set_formula" or "set_cell_value" or "set_formula_text")
        {
            var cell = Bounds(T(p, "cell"));
            return cell.R1 >= group.R1 && cell.R1 <= group.R2 && cell.C1 >= group.C1 && cell.C1 <= group.C2;
        }
        if (kind is not "copy_range" and not "import_range") return false;
        if (!p.GetProperty("copy_values").GetBoolean() && !p.GetProperty("copy_formulas").GetBoolean()) return false;
        var source = Bounds(T(p, "source_range")); var at = Bounds(T(p, "target_start"));
        return at.R1 <= group.R2 && at.R1 + source.R2 - source.R1 >= group.R1 &&
               at.C1 <= group.C2 && at.C1 + source.C2 - source.C1 >= group.C1;
    }
    public static void Preserve(string source, string candidate, JsonElement operations, bool preserveCaches = true, HashSet<(string Sheet, string Cell)>? invalidated = null)
    {
        if (!Path.GetExtension(source).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)) return;
        var ops = operations.EnumerateArray().ToArray();
        var writes = new CellWriteIndex(operations);
        using var input = ZipFile.OpenRead(source); using var output = ZipFile.Open(candidate, ZipArchiveMode.Update);
        var outputParts = Parts(output);
        foreach (var (sourceSheet, part) in Parts(input))
        {
            var identities = PlanOperations.FinalIdentities(new[] { sourceSheet }, operations);
            if (!identities.TryGetValue(sourceSheet, out string? finalName)) continue;
            string sheet = sourceSheet;
            var originals = Read(input, part).Descendants(S + "c").Where(c => c.Element(S + "f") is not null)
                .Select(c => (Address: (string)c.Attribute("r")!, Formula: new XElement(c.Element(S + "f")!), Value: c.Element(S + "v") is { } v ? new XElement(v) : null, Type: (string?)c.Attribute("t"))).ToArray();
            var grouped = originals.Where(c => (string?)c.Formula.Attribute("t") is "shared" or "array").ToArray();
            var staleArrayRanges = new HashSet<string>(StringComparer.Ordinal);
            if (grouped.Length > 0)
            {
                // Only untouched, stationary groups are enabled. Appending values
                // elsewhere is safe; moving a grid or writing into a group is not.
                if (ops.Any(o => T(o, "op") == "edit_plan_operation" && T(o.GetProperty("operation"), "operation") is
                    "rename_sheet" or "insert_rows" or "delete_rows" or "insert_columns" or "delete_columns"))
                    throw new NotSupportedException("Moving/renaming shared or array formulas is not validated");
                var mastersById = grouped.Where(c => (string?)c.Formula.Attribute("t") == "shared" && c.Formula.Attribute("ref") is not null)
                    .GroupBy(c => (string?)c.Formula.Attribute("si") ?? "")
                    .ToDictionary(g => g.Key, g => g.ToArray());
                var checkedRanges = new HashSet<string>(StringComparer.Ordinal);
                foreach (var cell in grouped)
                {
                    string? range = (string?)cell.Formula.Attribute("ref");
                    if ((string?)cell.Formula.Attribute("t") == "shared")
                    {
                        string id = (string?)cell.Formula.Attribute("si") ?? throw new InvalidDataException("Shared formula id missing");
                        if (!mastersById.TryGetValue(id, out var masters) || masters.Length != 1 || string.IsNullOrEmpty(masters[0].Formula.Value)) throw new InvalidDataException("Shared formula master is ambiguous");
                        range = (string?)masters[0].Formula.Attribute("ref");
                    }
                    if (range is null) throw new NotSupportedException("Grouped formula range missing");
                    var bounds = Bounds(range); var address = Bounds(cell.Address);
                    if (address.R1 < bounds.R1 || address.R1 > bounds.R2 || address.C1 < bounds.C1 || address.C1 > bounds.C2)
                        throw new InvalidDataException("Formula member outside its group");
                    if (checkedRanges.Add(range) && (writes.StationaryScalarPlan
                        ? writes.Intersects(sourceSheet, bounds.R1, bounds.C1, bounds.R2, bounds.C2)
                        : ops.Any(o => WritesInto(sourceSheet, range, o))))
                        throw new NotSupportedException("Writing into a shared/array formula range is not validated: " + sourceSheet + "!" + range);
                    if (preserveCaches && (string?)cell.Formula.Attribute("t") == "array" && invalidated is not null &&
                        invalidated.Any(c => {
                            if (c.Sheet != sourceSheet) return false;
                            var at = Bounds(c.Cell);
                            return at.R1 >= bounds.R1 && at.R1 <= bounds.R2 && at.C1 >= bounds.C1 && at.C1 <= bounds.C2;
                        }))
                        staleArrayRanges.Add(range);
                }
            }
            var expected = new Dictionary<string, XElement>();
            var caches = new Dictionary<string, (XElement? Value, string? Type)>();
            foreach (var original in originals)
            {
                string name = sheet, address = original.Address; var formula = new XElement(original.Formula); bool retained = true;
                if ((string?)formula.Attribute("t") is { } type && type is not "normal" and not "shared" and not "array")
                    throw new NotSupportedException("Formula type is not validated: " + type);
                if ((string?)formula.Attribute("t") is not "shared" and not "array" && formula.Attributes().Any(a => a.Name.LocalName is "si" or "ref"))
                    throw new NotSupportedException("Ambiguous formula grouping metadata");
                if (writes.StationaryScalarPlan) retained = !writes.Contains(name, address);
                else foreach (var outer in ops)
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
            if (expected.Count == 0) continue;
            var document = Read(output, outputParts[finalName]);
            // Index once: scanning the entire sheet for every formula is
            // quadratic on the complex-workbook probe. Duplicate addresses
            // remain an error rather than silently choosing one cell.
            var savedCells = document.Descendants(S + "c").ToDictionary(c => (string)c.Attribute("r")!, StringComparer.Ordinal);
            // Never orphan SDK-generated shared followers by replacing only
            // part of one of its groups. Original groups get fresh per-sheet
            // ids, avoiding collisions with newly generated formula groups.
            foreach (var group in savedCells.Where(c => (string?)c.Value.Element(S + "f")?.Attribute("t") == "shared")
                .GroupBy(c => (string?)c.Value.Element(S + "f")?.Attribute("si")))
                if (group.Any(c => expected.ContainsKey(c.Key)) && group.Any(c => !expected.ContainsKey(c.Key)))
                    throw new NotSupportedException("Saved shared formula group crosses a modified range");
            int nextShared = savedCells.Values.Select(c => (int?)c.Element(S + "f")?.Attribute("si") ?? -1).DefaultIfEmpty(-1).Max();
            var sharedIds = new Dictionary<string, int>();
            foreach (var (address, formula) in expected)
            {
                if (!savedCells.TryGetValue(address, out var cell))
                    throw new InvalidDataException("Saved formula cell missing: " + finalName + "!" + address);
                var actual = cell.Element(S + "f") ?? throw new InvalidDataException("Saved formula missing");
                if ((string?)formula.Attribute("t") == "shared")
                {
                    string id = (string)formula.Attribute("si")!;
                    if (!sharedIds.TryGetValue(id, out int mapped)) sharedIds[id] = mapped = checked(++nextShared);
                    formula.SetAttributeValue("si", mapped);
                }
                actual.ReplaceWith(formula);
                if (caches.TryGetValue(address, out var cached))
                {
                    cell.Element(S + "v")?.Remove();
                    cell.SetAttributeValue("t", cached.Type);
                    if (cached.Value is not null) cell.Add(new XElement(cached.Value));
                }
            }
            // Array followers have no original <f> node, but their values are
            // caches too. Clear the whole affected result extent, never refuse
            // a content-safe edit merely because recalculation is needed.
            foreach (string range in staleArrayRanges)
            {
                var bounds = Bounds(range);
                foreach (var (address, cell) in savedCells)
                {
                    var at = Bounds(address);
                    if (at.R1 < bounds.R1 || at.R1 > bounds.R2 || at.C1 < bounds.C1 || at.C1 > bounds.C2) continue;
                    cell.Element(S + "v")?.Remove();
                    cell.Element(S + "is")?.Remove();
                    cell.SetAttributeValue("t", null);
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
