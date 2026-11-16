using System.Globalization;
using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;

// Check actual serialized styles, not the in-memory model that is being tested.
internal static class SavedStyleVerifier
{
    private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static XElement Read(ZipArchive zip, string path)
    {
        using var stream = zip.GetEntry(path)?.Open() ?? throw new InvalidDataException("Missing part " + path);
        return XElement.Load(stream);
    }
    private static bool Enabled(XElement? node) => node is not null && (string?)node.Attribute("val") is not "0" and not "false";
    private static int Index(XElement? node, string attribute) => int.Parse((string?)node?.Attribute(attribute) ?? "0", CultureInfo.InvariantCulture);
    public static void Verify(ZipArchive zip, JsonElement operations)
    {
        var expected = new Dictionary<(string Sheet, string Cell), Dictionary<string, JsonElement>>();
        foreach (var op in operations.EnumerateArray())
        {
            if (op.GetProperty("op").GetString() != "set_style") continue;
            var key = (op.GetProperty("sheet").GetString()!, op.GetProperty("cell").GetString()!);
            if (!expected.TryGetValue(key, out var properties)) expected[key] = properties = new();
            foreach (var p in op.GetProperty("style").EnumerateObject()) properties[p.Name] = p.Value;
        }
        if (expected.Count == 0) return;
        var book = Read(zip, "xl/workbook.xml");
        var relationships = Read(zip, "xl/_rels/workbook.xml.rels");
        var styles = Read(zip, "xl/styles.xml");
        var xfs = styles.Element(S + "cellXfs")!.Elements().ToArray();
        var fonts = styles.Element(S + "fonts")!.Elements().ToArray();
        var fills = styles.Element(S + "fills")!.Elements().ToArray();
        var sheets = new Dictionary<string, XElement>();
        foreach (var (key, properties) in expected)
        {
            if (!sheets.TryGetValue(key.Sheet, out var xml))
            {
                var sheet = book.Element(S + "sheets")!.Elements().Single(s => (string?)s.Attribute("name") == key.Sheet);
                string id = (string)sheet.Attribute(R + "id")!;
                string target = (string)relationships.Elements().Single(r => (string?)r.Attribute("Id") == id).Attribute("Target")!;
                if (target.Contains("..") || target.Contains('\\')) throw new InvalidDataException("Unsafe sheet target");
                sheets[key.Sheet] = xml = Read(zip, target.StartsWith('/') ? target.TrimStart('/') : "xl/" + target);
            }
            var cell = xml.Descendants(S + "c").SingleOrDefault(c => (string?)c.Attribute("r") == key.Cell);
            int styleIndex = Index(cell, "s");
            // Explicit target styles normally have a cell xf; account for
            // inherited row/column styles when the serialized cell omits it.
            if (cell?.Attribute("s") is null)
            {
                string row = new(key.Cell.Where(char.IsDigit).ToArray());
                var rowNode = xml.Element(S + "sheetData")?.Elements(S + "row").SingleOrDefault(r => (string?)r.Attribute("r") == row);
                if (rowNode?.Attribute("s") is not null) styleIndex = Index(rowNode, "s");
                else
                {
                    int col = 0;
                    foreach (char c in key.Cell.TakeWhile(char.IsLetter)) col = col * 26 + c - 'A' + 1;
                    var column = xml.Element(S + "cols")?.Elements(S + "col").LastOrDefault(c => Index(c, "min") <= col && Index(c, "max") >= col);
                    styleIndex = Index(column, "style");
                }
            }
            var xf = xfs[styleIndex];
            var font = fonts[Index(xf, "fontId")];
            var fill = fills[Index(xf, "fillId")].Element(S + "patternFill");
            foreach (var (name, value) in properties)
            {
                bool valid = name switch
                {
                    "bold" => Enabled(font.Element(S + "b")) == value.GetBoolean(),
                    "italic" => Enabled(font.Element(S + "i")) == value.GetBoolean(),
                    "font_name" => (string?)font.Element(S + "name")?.Attribute("val") == value.GetString(),
                    "font_size" => double.Parse((string?)font.Element(S + "sz")?.Attribute("val") ?? "0", CultureInfo.InvariantCulture) == value.GetInt32(),
                    "fill_rgb" => (string?)fill?.Attribute("patternType") == "solid" && string.Equals((string?)fill.Element(S + "fgColor")?.Attribute("rgb"), "FF" + value.GetString(), StringComparison.OrdinalIgnoreCase),
                    "wrap_text" => ((string?)xf.Element(S + "alignment")?.Attribute("wrapText") is "1" or "true") == value.GetBoolean(),
                    "number_format" => NumberFormat(styles, Index(xf, "numFmtId")) == value.GetString(),
                    _ => false,
                };
                if (!valid) throw new InvalidDataException($"OUTPUT_VALIDATION_FAILED: saved style {name} differs at {key.Sheet}!{key.Cell}");
            }
        }
    }
    private static string? NumberFormat(XElement styles, int id) =>
        (string?)styles.Element(S + "numFmts")?.Elements().SingleOrDefault(n => Index(n, "numFmtId") == id)?.Attribute("formatCode")
        ?? id switch { 0 => "General", 1 => "0", 2 => "0.00", 9 => "0%", 10 => "0.00%", 49 => "@", _ => null };
}
