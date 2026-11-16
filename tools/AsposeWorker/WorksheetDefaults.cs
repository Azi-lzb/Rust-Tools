using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

// Modify currently supports only values/formulas/styles/header text: none of
// these operations is allowed to change worksheet-wide dimension defaults.
// Aspose can otherwise serialize an implicit baseColWidth as a different
// explicit defaultColWidth. Preserve the source declaration, never the watermark.
internal static class WorksheetDefaults
{
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static XDocument Read(ZipArchive zip, string part)
    {
        using var stream = (zip.GetEntry(part) ?? throw new InvalidDataException("Missing worksheet package part: " + part)).Open();
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }
    private static Dictionary<string, string> Parts(ZipArchive zip)
    {
        var relationships = Read(zip, "xl/_rels/workbook.xml.rels").Root!.Elements().ToDictionary(
            e => (string)e.Attribute("Id")!, e => e, StringComparer.Ordinal);
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var sheet in Read(zip, "xl/workbook.xml").Descendants(Main + "sheet"))
        {
            var relationship = relationships[(string)sheet.Attribute(Rel + "id")!];
            if ((string?)relationship.Attribute("TargetMode") == "External") throw new InvalidDataException("External worksheet target");
            var target = new Uri(new Uri("http://ooxml/xl/workbook.xml"), (string)relationship.Attribute("Target")!);
            if (target.Host != "ooxml" || target.Query.Length != 0 || target.Fragment.Length != 0) throw new InvalidDataException("Invalid worksheet target");
            result.Add((string)sheet.Attribute("name")!, Uri.UnescapeDataString(target.AbsolutePath.TrimStart('/')));
        }
        return result;
    }
    public static int Preserve(string source, string candidate, Dictionary<string, string>? identities = null)
    {
        if (!Path.GetExtension(source).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)) return 0;
        using var input = ZipFile.OpenRead(source);
        using var output = ZipFile.Open(candidate, ZipArchiveMode.Update);
        var sourceParts = Parts(input);
        var outputParts = Parts(output);
        int changed = 0;
        foreach (var (sheet, sourcePart) in sourceParts)
        {
            if (identities is not null && !identities.ContainsKey(sheet)) continue;
            string name = identities is null ? sheet : identities[sheet];
            string outputPart = outputParts[name];
            var original = Read(input, sourcePart).Root!.Element(Main + "sheetFormatPr");
            var document = Read(output, outputPart);
            var actual = document.Root!.Element(Main + "sheetFormatPr");
            if (XNode.DeepEquals(original, actual)) continue;
            if (original is null) actual?.Remove();
            else if (actual is not null) actual.ReplaceWith(new XElement(original));
            else throw new InvalidDataException("Saved worksheet lost its defaults element");
            output.GetEntry(outputPart)!.Delete();
            using (var stream = output.CreateEntry(outputPart, CompressionLevel.Optimal).Open())
            using (var writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new System.Text.UTF8Encoding(false), Indent = false })) document.Save(writer);
            changed++;
        }
        return changed;
    }
}
