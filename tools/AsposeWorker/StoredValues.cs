using Aspose.Cells;
using System.IO.Compression;
using System.Xml.Linq;

// An absent formula cache is a blank cached value, never the SDK's implicit 0.
// Take the data-only snapshot before modifications, matching the existing
// openpyxl range-copy contract without calculating a formula.
internal sealed class StoredValues
{
    internal sealed record Value(object? Data, bool Error);
    private readonly Dictionary<(string Sheet, string Cell), Value> values = new();
    public StoredValues(string input, Workbook workbook)
    {
        var missing = new HashSet<(string, string)>();
        if (Path.GetExtension(input).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            using var zip = ZipFile.OpenRead(input);
            XDocument Read(string name) { using var f = zip.GetEntry(name)!.Open(); return XDocument.Load(f); }
            XNamespace s = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            var relationships = Read("xl/_rels/workbook.xml.rels").Root!.Elements().ToDictionary(e => (string)e.Attribute("Id")!, e => (string)e.Attribute("Target")!);
            foreach (var sheet in Read("xl/workbook.xml").Descendants(s + "sheet"))
            {
                string part = Uri.UnescapeDataString(new Uri(new Uri("http://ooxml/xl/workbook.xml"), relationships[(string)sheet.Attribute(r + "id")!]).AbsolutePath.TrimStart('/'));
                foreach (var cell in Read(part).Descendants(s + "c"))
                    if (cell.Element(s + "f") is not null && string.IsNullOrEmpty(cell.Element(s + "v")?.Value)) missing.Add(((string)sheet.Attribute("name")!, (string)cell.Attribute("r")!));
            }
        }
        foreach (Worksheet sheet in workbook.Worksheets) foreach (Cell cell in sheet.Cells)
            values[(sheet.Name, cell.Name)] = missing.Contains((sheet.Name, cell.Name)) ? new(null, false) : new(cell.Value, cell.Type == CellValueType.IsError);
    }
    public Value At(string sheet, string cell) => values.GetValueOrDefault((sheet, cell), new(null, false));
}
