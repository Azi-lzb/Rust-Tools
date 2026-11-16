using System.Globalization;
using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;

// Exact serialization of properties explicitly requested by the plan. This is
// not a broad round-trip normalizer and never edits the evaluation worksheet.
internal static class PlanSerialization
{
    private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static string T(JsonElement p, string key) => p.GetProperty(key).GetString()!;
    private static XDocument Read(ZipArchive zip, string name) { using var f = zip.GetEntry(name)!.Open(); return XDocument.Load(f, LoadOptions.PreserveWhitespace); }
    public static void Apply(string candidate, JsonElement operations)
    {
        using var zip = ZipFile.Open(candidate, ZipArchiveMode.Update);
        var book = Read(zip, "xl/workbook.xml"); var rels = Read(zip, "xl/_rels/workbook.xml.rels").Root!.Elements().ToDictionary(r => (string)r.Attribute("Id")!, r => (string)r.Attribute("Target")!);
        var parts = book.Descendants(S + "sheet").ToDictionary(s => (string)s.Attribute("name")!, s => Uri.UnescapeDataString(new Uri(new Uri("http://ooxml/xl/workbook.xml"), rels[(string)s.Attribute(R + "id")!]).AbsolutePath.TrimStart('/')));
        var documents = new Dictionary<string, XDocument>();
        var plan = operations.EnumerateArray().Where(o => T(o, "op") == "edit_plan_operation").Select(o => o.GetProperty("operation")).ToArray();
        var created = new HashSet<string>(StringComparer.Ordinal);
        foreach (var p in plan)
        {
            string kind = T(p, "operation");
            if (kind is "create_sheet" or "create_sheet_at") created.Add(T(p, "name"));
            if (kind == "delete_sheet") created.Remove(T(p, "name"));
            if (kind == "rename_sheet" && created.Remove(T(p, "old_name"))) created.Add(T(p, "new_name"));
        }
        XDocument Sheet(string name)
        {
            if (!documents.TryGetValue(name, out var d)) documents[name] = d = Read(zip, parts[name]);
            return d;
        }
        foreach (var name in created)
        {
            var d = Sheet(name); var defaults = d.Root!.Element(S + "sheetFormatPr");
            if (defaults is null) throw new InvalidDataException("Created sheet has no defaults");
            defaults.ReplaceWith(new XElement(S + "sheetFormatPr", new XAttribute("baseColWidth", "8"), new XAttribute("defaultRowHeight", "15")));
        }
        for (int i = 0; i < plan.Length; i++)
        {
            var p = plan[i]; string kind = T(p, "operation");
            if (kind == "set_workbook_calculation_properties")
            {
                var calculation = book.Root!.Element(S + "calcPr");
                if (calculation is null) { calculation = new XElement(S + "calcPr"); book.Root.Add(calculation); }
                foreach (var pair in new[] { ("full_calc_on_load", "fullCalcOnLoad"), ("force_full_calc", "forceFullCalc") })
                    if (p.TryGetProperty(pair.Item1, out var value) && value.ValueKind != JsonValueKind.Null) calculation.SetAttributeValue(pair.Item2, value.GetBoolean() ? "1" : "0");
                continue;
            }
            if (kind is not "set_column_width" and not "set_row_height" and not "set_page_settings" and not "import_range") continue;
            string name = T(p, kind == "import_range" ? "target_sheet" : "sheet"); bool deleted = false;
            for (int j = i + 1; j < plan.Length; j++)
            {
                var later = plan[j]; string other = T(later, "operation");
                if (other == "rename_sheet" && T(later, "old_name") == name) name = T(later, "new_name");
                if (other == "delete_sheet" && T(later, "name") == name) { deleted = true; break; }
            }
            if (deleted) continue;
            var d = Sheet(name);
            if (kind == "set_column_width")
            {
                int column = p.GetProperty("column_index").GetInt32(); var cols = d.Root!.Element(S + "cols");
                if (cols is null) { cols = new XElement(S + "cols"); d.Root.Element(S + "sheetData")!.AddBeforeSelf(cols); }
                XElement? old = cols.Elements().SingleOrDefault(c => (int)c.Attribute("min")! <= column && (int)c.Attribute("max")! >= column);
                var single = old is null ? new XElement(S + "col") : new XElement(old);
                if (old is not null)
                {
                    int first = (int)old.Attribute("min")!, last = (int)old.Attribute("max")!;
                    if (first < column) { var left = new XElement(old); left.SetAttributeValue("max", column - 1); old.AddBeforeSelf(left); }
                    if (last > column) { var right = new XElement(old); right.SetAttributeValue("min", column + 1); old.AddAfterSelf(right); }
                    old.Remove();
                }
                single.SetAttributeValue("min", column); single.SetAttributeValue("max", column); single.SetAttributeValue("width", p.GetProperty("width").GetDouble().ToString("R", CultureInfo.InvariantCulture)); single.SetAttributeValue("customWidth", "1");
                cols.Add(single); cols.ReplaceNodes(cols.Elements().OrderBy(c => (int)c.Attribute("min")!).ToArray());
            }
            else if (kind == "set_row_height")
            {
                int row = p.GetProperty("row_index").GetInt32(); var data = d.Root!.Element(S + "sheetData")!;
                var element = data.Elements(S + "row").SingleOrDefault(r => (int)r.Attribute("r")! == row);
                if (element is null) { element = new XElement(S + "row", new XAttribute("r", row)); data.Add(element); data.ReplaceNodes(data.Elements().OrderBy(r => (int)r.Attribute("r")!).ToArray()); }
                element.SetAttributeValue("ht", p.GetProperty("height").GetDouble().ToString("R", CultureInfo.InvariantCulture)); element.SetAttributeValue("customHeight", "1");
            }
            else if (kind == "set_page_settings")
            {
                var settings = p.GetProperty("settings"); var setup = d.Root!.Element(S + "pageSetup");
                if (setup is null) { setup = new XElement(S + "pageSetup"); var footer = d.Root.Element(S + "headerFooter"); if (footer is null) d.Root.Add(setup); else footer.AddBeforeSelf(setup); }
                foreach (var pair in new[] { ("orientation", "orientation"), ("fit_to_width", "fitToWidth"), ("fit_to_height", "fitToHeight"), ("paper_size", "paperSize") })
                    if (settings.TryGetProperty(pair.Item1, out var value) && value.ValueKind != JsonValueKind.Null) setup.SetAttributeValue(pair.Item2, value.ToString());
                if (settings.TryGetProperty("clear_scale", out var clear) && clear.ValueKind == JsonValueKind.True) setup.Attribute("scale")?.Remove();
                if (settings.TryGetProperty("fit_to_page", out var fit) && fit.ValueKind != JsonValueKind.Null)
                {
                    var properties = d.Root.Element(S + "sheetPr"); if (properties is null) { properties = new XElement(S + "sheetPr"); d.Root.AddFirst(properties); }
                    var page = properties.Element(S + "pageSetUpPr"); if (page is null) { page = new XElement(S + "pageSetUpPr"); properties.Add(page); }
                    page.SetAttributeValue("fitToPage", fit.GetBoolean() ? "1" : "0");
                }
            }
            // Blank-cell materialization is checked by the existing import
            // validator; never synthesize values or caches during this pass.
        }
        foreach (var (name, document) in documents) Save(parts[name], document);
        Save("xl/workbook.xml", book);
        void Save(string part, XDocument document) { zip.GetEntry(part)!.Delete(); using var f = zip.CreateEntry(part, CompressionLevel.Optimal).Open(); document.Save(f, SaveOptions.DisableFormatting); }
    }
}
