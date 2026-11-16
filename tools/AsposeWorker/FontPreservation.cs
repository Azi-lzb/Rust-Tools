using Aspose.Cells;
using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;
using LoadOptions = System.Xml.Linq.LoadOptions;
using SaveOptions = System.Xml.Linq.SaveOptions;

// Preserve the lexical font declaration of effective source styles. In
// particular <font><b/></font> must not become an explicitly different font.
// Font ids are remapped into the candidate, never copied across workbooks.
internal static class FontPreservation
{
    private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static string T(JsonElement p, string k) => p.GetProperty(k).GetString()!;
    private static XDocument Read(ZipArchive zip, string name) { using var f = zip.GetEntry(name)!.Open(); return XDocument.Load(f, LoadOptions.PreserveWhitespace); }
    private static Dictionary<string, string> Parts(ZipArchive zip)
    {
        var rels = Read(zip, "xl/_rels/workbook.xml.rels").Root!.Elements().ToDictionary(e => (string)e.Attribute("Id")!, e => (string)e.Attribute("Target")!);
        return Read(zip, "xl/workbook.xml").Descendants(S + "sheet").ToDictionary(e => (string)e.Attribute("name")!, e => Uri.UnescapeDataString(new Uri(new Uri("http://ooxml/xl/workbook.xml"), rels[(string)e.Attribute(R + "id")!]).AbsolutePath.TrimStart('/')));
    }
    private static (int Row, int Col) Address(string name) { CellsHelper.CellNameToIndex(name.Replace("$", ""), out int row, out int col); return (row, col); }
    private sealed class Fonts
    {
        public required XElement Default;
        public Dictionary<(int Row, int Col), XElement> Cells = new();
        public Dictionary<int, XElement> Rows = new(), Columns = new();
        public XElement At(int row, int col) => Cells.TryGetValue((row, col), out var f) ? f : Rows.TryGetValue(row, out f) ? f : Columns.TryGetValue(col, out f) ? f : Default;
        public void Shift(bool rows, int start, int count, bool delete)
        {
            int? Move(int index) => delete && index >= start && index < start + count ? null : index >= start ? index + (delete ? -count : count) : index;
            var cells = new Dictionary<(int, int), XElement>();
            foreach (var (at, font) in Cells) if (Move(rows ? at.Row : at.Col) is { } index) cells.Add(rows ? (index, at.Col) : (at.Row, index), font);
            Cells = cells;
            var dimensions = rows ? Rows : Columns; var shifted = new Dictionary<int, XElement>();
            foreach (var (index, font) in dimensions) if (Move(index) is { } to) shifted.Add(to, font);
            if (rows) Rows = shifted; else Columns = shifted;
        }
    }
    private static Dictionary<string, Fonts> Load(string input, string component)
    {
        using var zip = ZipFile.OpenRead(input); var styles = Read(zip, "xl/styles.xml");
        var fonts = styles.Root!.Element(S + component + "s")!.Elements().ToArray();
        var xfs = styles.Root.Element(S + "cellXfs")!.Elements().ToArray();
        // Source font nodes are immutable and shared by cells using the same
        // style. Do not allocate a complete font XML tree per data cell.
        XElement Font(int xf) => fonts[(int?)xfs[xf].Attribute(component + "Id") ?? 0];
        var result = new Dictionary<string, Fonts>(StringComparer.Ordinal);
        foreach (var (name, part) in Parts(zip))
        {
            var state = new Fonts { Default = Font(0) }; var xml = Read(zip, part);
            foreach (var row in xml.Descendants(S + "row").Where(r => r.Attribute("s") is not null)) state.Rows[(int)row.Attribute("r")! - 1] = Font((int)row.Attribute("s")!);
            foreach (var col in xml.Descendants(S + "col").Where(c => c.Attribute("style") is not null))
                for (int c = (int)col.Attribute("min")! - 1; c < (int)col.Attribute("max")!; c++) state.Columns[c] = Font((int)col.Attribute("style")!);
            foreach (var cell in xml.Descendants(S + "c"))
            {
                var at = Address((string)cell.Attribute("r")!);
                state.Cells[at] = cell.Attribute("s") is { } style ? Font((int)style) : state.At(at.Row, at.Col);
            }
            result[name] = state;
        }
        return result;
    }
    public static void Preserve(string input, string candidate, JsonElement operations)
    {
        PreserveComponent(input, candidate, operations, "font");
        // Aspose can materialize a neighbouring border after copying a range.
        // Keep exact planned full-copy borders; basic copies keep target borders.
        if (operations.EnumerateArray().Any(p => T(p, "op") == "edit_plan_operation" && T(p.GetProperty("operation"), "operation") is "copy_range" or "import_range"))
        {
            PreserveComponent(input, candidate, operations, "border");
            PreserveComponent(input, candidate, operations, "fill");
        }
    }
    private static void PreserveComponent(string input, string candidate, JsonElement operations, string component)
    {
        if (!Path.GetExtension(input).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)) return;
        var sheets = Load(input, component); var defaultFont = new XElement(sheets.Values.First().Default);
        var imports = new Dictionary<string, Dictionary<string, Fonts>>(OperatingSystem.IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal);
        foreach (var outer in operations.EnumerateArray())
        {
            var p = outer; string kind = T(p, "op");
            if (kind == "edit_plan_operation") { p = p.GetProperty("operation"); kind = T(p, "operation"); }
            if (kind is "create_sheet" or "create_sheet_at") { sheets.Add(T(p, "name"), new Fonts { Default = new XElement(defaultFont) }); continue; }
            if (kind == "delete_sheet") { sheets.Remove(T(p, "name")); continue; }
            if (kind == "rename_sheet") { var sheet = sheets[T(p, "old_name")]; sheets.Remove(T(p, "old_name")); sheets.Add(T(p, "new_name"), sheet); continue; }
            if (kind is "copy_range" or "import_range")
            {
                Dictionary<string, Fonts> sourceBook = sheets;
                if (kind == "import_range")
                {
                    string path = Path.GetFullPath(T(p, "source_workbook"));
                    if (!imports.TryGetValue(path, out var loaded)) imports[path] = loaded = Load(path, component);
                    sourceBook = loaded;
                }
                var source = sourceBook[T(p, "source_sheet")]; var target = sheets[T(p, "target_sheet")];
                string[] bounds = T(p, "source_range").Split(':'); var a = Address(bounds[0]); var z = Address(bounds[^1]); var at = Address(T(p, "target_start"));
                bool styles = kind == "copy_range" ? p.GetProperty("copy_styles").GetBoolean() : (component is "font" or "fill" ? T(p, "style_mode") != "none" : T(p, "style_mode") == "full");
                var copied = new Dictionary<(int, int), XElement>();
                for (int r = a.Row; r <= z.Row; r++) for (int c = a.Col; c <= z.Col; c++)
                {
                    int tr = at.Row + r - a.Row, tc = at.Col + c - a.Col;
                    copied[(tr, tc)] = styles ? source.At(r, c) : target.At(tr, tc);
                }
                foreach (var (key, font) in copied) target.Cells[key] = font;
                continue;
            }
            if (!p.TryGetProperty("sheet", out var sheetName) || !sheets.TryGetValue(sheetName.GetString()!, out var state)) continue;
            if (kind is "delete_rows" or "insert_rows" or "delete_columns" or "insert_columns")
            {
                bool rows = kind.EndsWith("rows"); state.Shift(rows, p.GetProperty(rows ? "first_row" : "first_column").GetInt32() - 1, p.GetProperty("count").GetInt32(), kind.StartsWith("delete")); continue;
            }
            if (!p.TryGetProperty("cell", out var address)) continue;
            var coordinate = Address(address.GetString()!); var desired = new XElement(state.At(coordinate.Row, coordinate.Col));
            if (kind == "set_style" && component == "font")
                foreach (var v in p.GetProperty("style").EnumerateObject())
                {
                    string? name = v.Name switch { "bold" => "b", "italic" => "i", "font_name" => "name", "font_size" => "sz", _ => null };
                    if (name is null) continue;
                    desired.Element(S + name)?.Remove();
                    if (v.Value.ValueKind == JsonValueKind.True) desired.Add(new XElement(S + name));
                    else if (v.Value.ValueKind != JsonValueKind.False) desired.Add(new XElement(S + name, new XAttribute("val", v.Value.ToString())));
                }
            if (kind == "set_style" && component == "fill" && p.GetProperty("style").TryGetProperty("fill_rgb", out var rgb))
                desired = new XElement(S + "fill", new XElement(S + "patternFill", new XAttribute("patternType", "solid"), new XElement(S + "fgColor", new XAttribute("rgb", "FF" + rgb.GetString()!))));
            state.Cells[coordinate] = desired;
        }
        using var output = ZipFile.Open(candidate, ZipArchiveMode.Update);
        var styleDoc = Read(output, "xl/styles.xml"); var fontRoot = styleDoc.Root!.Element(S + component + "s")!; var xfRoot = styleDoc.Root.Element(S + "cellXfs")!;
        var fontsOut = fontRoot.Elements().ToList(); var xfsOut = xfRoot.Elements().ToList(); var styleCache = new Dictionary<(int, int), int>();
        var fontCache = new Dictionary<XElement, int>(ReferenceEqualityComparer.Instance);
        var parts = Parts(output);
        foreach (var (name, state) in sheets)
        {
            if (name == "Evaluation Warning") continue;
            var document = Read(output, parts[name]);
            foreach (var cell in document.Descendants(S + "c"))
            {
                var coordinate = Address((string)cell.Attribute("r")!);
                if (!state.Cells.TryGetValue(coordinate, out var desired)) continue;
                if (!fontCache.TryGetValue(desired, out int fontId))
                {
                    fontId = fontsOut.FindIndex(f => XNode.DeepEquals(f, desired));
                    if (fontId < 0) { fontId = fontsOut.Count; var font = new XElement(desired); fontsOut.Add(font); fontRoot.Add(font); }
                    fontCache[desired] = fontId;
                }
                int index = (int?)cell.Attribute("s") ?? 0;
                if (((int?)xfsOut[index].Attribute(component + "Id") ?? 0) == fontId) continue;
                if (!styleCache.TryGetValue((index, fontId), out int newStyle))
                {
                    var xf = new XElement(xfsOut[index]); xf.SetAttributeValue(component + "Id", fontId); xf.SetAttributeValue("apply" + char.ToUpperInvariant(component[0]) + component[1..], "1");
                    newStyle = xfsOut.Count; xfsOut.Add(xf); xfRoot.Add(xf); styleCache[(index, fontId)] = newStyle;
                }
                cell.SetAttributeValue("s", newStyle);
            }
            Save(parts[name], document);
        }
        fontRoot.SetAttributeValue("count", fontsOut.Count); xfRoot.SetAttributeValue("count", xfsOut.Count); Save("xl/styles.xml", styleDoc);
        void Save(string part, XDocument document) { output.GetEntry(part)!.Delete(); using var f = output.CreateEntry(part, CompressionLevel.Optimal).Open(); document.Save(f, SaveOptions.DisableFormatting); }
    }
}
