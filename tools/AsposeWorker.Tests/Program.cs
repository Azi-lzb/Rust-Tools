using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;

// Dependency-free unit runner: directly exercises the production sanitizer.
if (args.Length != 1) throw new ArgumentException("Supply a new isolated output directory");
string root = Path.GetFullPath(args[0]);
if (Directory.Exists(root)) throw new IOException("Refuse existing test directory");
Directory.CreateDirectory(root);
XNamespace s = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
XNamespace p = "http://schemas.openxmlformats.org/package/2006/relationships";
var results = new List<object>();
void Assert(bool value, string reason) { if (!value) throw new Exception(reason); }
string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));
string Read(string path, string part) { using var zip = ZipFile.OpenRead(path); using var reader = new StreamReader(zip.GetEntry(part)!.Open()); return reader.ReadToEnd(); }
void Write(ZipArchive zip, string part, string text) { using var writer = new StreamWriter(zip.CreateEntry(part).Open()); writer.Write(text); }
string Fixture(string label, string extra = "notice", string formula = "Data!A1+1", bool objectPart = false, bool extraInMiddle = false)
{
    string path = Path.Combine(root, label + ".xlsx");
    using var zip = ZipFile.Open(path, ZipArchiveMode.Create);
    string[] names = extraInMiddle ? ["Data", OutputSanitation.SystemSheet, "Blank"] : ["Data", "Blank", OutputSanitation.SystemSheet];
    var sheets = names.Select((name, index) => new XElement(s + "sheet", new XAttribute("name", name), new XAttribute("sheetId", index + 1), new XAttribute(r + "id", "rId" + (index + 1))));
    var book = new XElement(s + "workbook", new XAttribute(XNamespace.Xmlns + "r", r), new XElement(s + "bookViews", new XElement(s + "workbookView", new XAttribute("activeTab", 2))), new XElement(s + "sheets", sheets),
        new XElement(s + "definedNames", new XElement(s + "definedName", new XAttribute("name", "Local"), new XAttribute("localSheetId", Array.IndexOf(names, "Blank")), "Blank!$A$1")));
    Write(zip, "xl/workbook.xml", book.ToString());
    Write(zip, "xl/_rels/workbook.xml.rels", new XElement(p + "Relationships", names.Select((n, i) => new XElement(p + "Relationship", new XAttribute("Id", "rId" + (i + 1)), new XAttribute("Type", r.NamespaceName + "/worksheet"), new XAttribute("Target", "worksheets/sheet" + (i + 1) + ".xml")))).ToString());
    XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
    Write(zip, "[Content_Types].xml", new XElement(ct + "Types", names.Select((n, i) => new XElement(ct + "Override", new XAttribute("PartName", "/xl/worksheets/sheet" + (i + 1) + ".xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")))).ToString());
    for (int i = 0; i < names.Length; i++)
    {
        var data = new XElement(s + "sheetData");
        if (names[i] == "Data") data.Add(new XElement(s + "row", new XAttribute("r", 1), new XElement(s + "c", new XAttribute("r", "A1"), new XElement(s + "v", 7)), new XElement(s + "c", new XAttribute("r", "B1"), new XElement(s + "f", formula), new XElement(s + "v", 8))));
        if (names[i] == OutputSanitation.SystemSheet && extra != "blank")
            data.Add(new XElement(s + "row", new XAttribute("r", 5), new XElement(s + "c", new XAttribute("r", "A5"), new XAttribute("t", "inlineStr"), new XElement(s + "is", new XElement(s + "t", extra == "notice" ? "Evaluation Only. Created with Aspose.Cells for .NET. Copyright 2003 - 2026 Aspose Pty Ltd." : extra)))));
        var sheet = new XElement(s + "worksheet", new XElement(s + "sheetViews", new XElement(s + "sheetView", new XAttribute("workbookViewId", 0), new XAttribute("tabSelected", names[i] == OutputSanitation.SystemSheet ? "1" : "0"))), data);
        Write(zip, $"xl/worksheets/sheet{i + 1}.xml", sheet.ToString());
        if (names[i] == OutputSanitation.SystemSheet && objectPart) Write(zip, $"xl/worksheets/_rels/sheet{i + 1}.xml.rels", new XElement(p + "Relationships").ToString());
    }
    Write(zip, "customXml/item1.xml", "<business>keep exactly</business>");
    XNamespace vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
    XNamespace app = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
    Write(zip, "docProps/app.xml", new XElement(app + "Properties",
        new XElement(app + "HeadingPairs", new XElement(vt + "vector", new XAttribute("size", 4), new XAttribute("baseType", "variant"),
            new XElement(vt + "variant", new XElement(vt + "lpstr", "Worksheets")), new XElement(vt + "variant", new XElement(vt + "i4", 3)),
            new XElement(vt + "variant", new XElement(vt + "lpstr", "Named Ranges")), new XElement(vt + "variant", new XElement(vt + "i4", 1)))),
        new XElement(app + "TitlesOfParts", new XElement(vt + "vector", new XAttribute("size", 4), new XAttribute("baseType", "lpstr"),
            names.Append("Local").Select(n => new XElement(vt + "lpstr", n))))).ToString());
    return path;
}
void Case(string label, Action action) { action(); results.Add(new { name = label, status = "PASS" }); }
void Reject(string path, OutputSanitation.Plan plan)
{
    string before = Hash(path);
    try { OutputSanitation.Clean(path, plan); throw new Exception("Unsafe cleanup accepted"); }
    catch (InvalidDataException) { }
    Assert(Hash(path) == before, "Rejected candidate was modified");
}
var plan = OutputSanitation.Prepare(["Data", "Blank"], [true, true], "Blank");
foreach (string kind in new[] { "notice", "blank" }) Case(kind, () =>
{
    string path = Fixture(kind, kind);
    string custom = Read(path, "customXml/item1.xml");
    var result = OutputSanitation.Clean(path, plan);
    Assert(result.removed_count == 1 && result.active_index == 1 && result.validated, "Incorrect sanitation result");
    Assert(Read(path, "customXml/item1.xml") == custom, "Business part changed");
    Assert(Read(path, "xl/worksheets/sheet1.xml").Contains("Data!A1+1"), "Business formula changed");
    using var zip = ZipFile.OpenRead(path);
    Assert(zip.GetEntry("xl/worksheets/sheet3.xml") is null, "Orphan system part remains");
    Assert(!Read(path, "xl/_rels/workbook.xml.rels").Contains("sheet3.xml"), "Orphan relationship remains");
    Assert(!Read(path, "[Content_Types].xml").Contains("sheet3.xml"), "Orphan content type remains");
    Assert(!Read(path, "docProps/app.xml").Contains(OutputSanitation.SystemSheet) && Read(path, "docProps/app.xml").Contains("Local"), "Extended worksheet properties not cleaned safely");
} );
Case("idempotence", () => { string path = Fixture("idempotence"); OutputSanitation.Clean(path, plan); var result = OutputSanitation.Clean(path, plan); Assert(result.removed_count == 0 && !result.view_corrected, "Repeated cleanup changed logical state"); OutputSanitation.Verify(path, plan); });
Case("business-warning-name-preserved", () => { string path = Fixture("business-warning", "business data"); var business = OutputSanitation.Prepare(["Data", "Blank", OutputSanitation.SystemSheet], [true, true, true], OutputSanitation.SystemSheet); Assert(OutputSanitation.Clean(path, business).removed_count == 0, "Business namesake deleted"); Assert(Read(path, "xl/worksheets/sheet3.xml").Contains("business data"), "Business namesake content lost"); });
Case("unexpected-business-data-rejected", () => Reject(Fixture("unknown-data", "business data"), plan));
Case("objects-rejected", () => Reject(Fixture("objects", objectPart: true), plan));
Case("formula-reference-rejected", () => Reject(Fixture("reference", formula: "'Evaluation Warning'!A1"), plan));
Case("dynamic-reference-rejected", () => Reject(Fixture("dynamic", formula: "INDIRECT(A1)"), plan));
Case("middle-system-sheet-local-name-rebased", () => { string path = Fixture("middle", extraInMiddle: true); OutputSanitation.Clean(path, plan); Assert(Read(path, "xl/workbook.xml").Contains("localSheetId=\"1\""), "Local name index not rebased"); });
Case("hidden-active-fallback", () => Assert(OutputSanitation.Prepare(["Hidden", "Visible"], [false, true], "Hidden").ActiveIndex == 1, "Hidden active not corrected"));
Case("deleted-active-fallback", () => Assert(OutputSanitation.Prepare(["First", "Next"], [true, true], "Deleted").ActiveIndex == 0, "Deleted active not corrected"));
Case("no-visible-sheet-rejected", () => { try { OutputSanitation.Prepare(["Hidden"], [false], null); throw new Exception("No visible worksheet accepted"); } catch (InvalidDataException) { } });
Case("ordered-business-roster", () =>
{
    using var operations = JsonDocument.Parse("""
        [{"op":"edit_plan_operation","operation":{"operation":"rename_sheet","old_name":"Data","new_name":"Renamed"}},
         {"op":"edit_plan_operation","operation":{"operation":"create_sheet_at","name":"NewBlank","index":0}},
         {"op":"edit_plan_operation","operation":{"operation":"move_sheet","sheet":"Blank","new_index":0}},
         {"op":"edit_plan_operation","operation":{"operation":"delete_sheet","name":"Renamed"}}]
        """);
    Assert(OutputSanitation.BusinessRoster(["Data", "Blank"], operations.RootElement).SequenceEqual(new[] { "Blank", "NewBlank" }), "Ordered business roster lost template identity");
});
Case("unknown-extra-name-rejected", () =>
{
    string path = Fixture("unknown-name");
    var book = XDocument.Parse(Read(path, "xl/workbook.xml"));
    using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
    {
        book.Descendants(s + "sheet").Last().SetAttributeValue("name", "Sheet99");
        zip.GetEntry("xl/workbook.xml")!.Delete(); Write(zip, "xl/workbook.xml", book.ToString());
    }
    Reject(path, plan);
});
foreach (bool businessUsesNotice in new[] { false, true }) Case("shared-string-preservation-" + businessUsesNotice, () =>
{
    string path = Fixture("shared-string-" + businessUsesNotice);
    var business = XDocument.Parse(Read(path, "xl/worksheets/sheet1.xml"));
    var warning = XDocument.Parse(Read(path, "xl/worksheets/sheet3.xml"));
    string notice = warning.Descendants(s + "t").Single().Value;
    var a1 = business.Descendants(s + "c").First(); a1.SetAttributeValue("t", "s"); a1.Element(s + "v")!.Value = businessUsesNotice ? "1" : "0";
    var a5 = warning.Descendants(s + "c").Single(); a5.SetAttributeValue("t", "s"); a5.ReplaceNodes(new XElement(s + "v", 1));
    using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
    {
        foreach (var (part, doc) in new[] { ("xl/worksheets/sheet1.xml", business), ("xl/worksheets/sheet3.xml", warning) })
        { zip.GetEntry(part)!.Delete(); Write(zip, part, doc.ToString()); }
        Write(zip, "xl/sharedStrings.xml", new XElement(s + "sst", new XAttribute("count", 2), new XAttribute("uniqueCount", 2),
            new XElement(s + "si", new XElement(s + "t", "Business text")), new XElement(s + "si", new XElement(s + "t", notice))).ToString());
    }
    OutputSanitation.Clean(path, plan);
    var strings = XDocument.Parse(Read(path, "xl/sharedStrings.xml")).Descendants(s + "si").ToArray();
    Assert(strings.Length == 2 && strings[0].Value == "Business text" && strings[1].Value == (businessUsesNotice ? notice : ""), "Business shared strings changed or orphan notice retained");
});
File.WriteAllText(Path.Combine(root, "RESULT.json"), JsonSerializer.Serialize(new { status = "PASS", tests = results }, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine(JsonSerializer.Serialize(new { status = "PASS", tests = results.Count }));
