using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Text.Json;
using System.Xml.Linq;

// Operates only on an owned, isolated candidate. The pre-Save business roster
// is authoritative; a blank business/template sheet is never a cleanup target.
internal static class OutputSanitation
{
    internal const string SystemSheet = "Evaluation Warning";
    private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace Vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
    private const string WorksheetType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

    internal sealed record Plan(string[] Sheets, int ActiveIndex);
    internal sealed record Result(string stage, string[] removed_sheets, int removed_count,
        int active_index, bool view_corrected, bool validated);

    internal static bool IsNotice(string text) => Regex.IsMatch(text,
        @"\AEvaluation Only\. Created with Aspose\.Cells for \.NET\. Copyright 2003 - [0-9]{4} Aspose Pty Ltd\.\z", RegexOptions.CultureInvariant);

    internal static bool UnsafeReference(string text) => text.Contains(SystemSheet, StringComparison.OrdinalIgnoreCase) ||
        Regex.IsMatch(text, @"\bINDIRECT\s*\(", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    internal static string[] BusinessRoster(string[] source, JsonElement operations)
    {
        var sheets = source.ToList();
        if (operations.ValueKind == JsonValueKind.Array)
        foreach (var outer in operations.EnumerateArray())
        {
            if (outer.GetProperty("op").GetString() != "edit_plan_operation") continue;
            var op = outer.GetProperty("operation");
            string Text(string key) => op.GetProperty(key).GetString()!;
            switch (Text("operation"))
            {
                case "rename_sheet": sheets[sheets.IndexOf(Text("old_name"))] = Text("new_name"); break;
                case "delete_sheet": sheets.Remove(Text("name")); break;
                case "create_sheet": sheets.Add(Text("name")); break;
                case "create_sheet_at": sheets.Insert(op.GetProperty("index").GetInt32(), Text("name")); break;
                case "move_sheet":
                    string name = Text("sheet"); sheets.Remove(name); sheets.Insert(op.GetProperty("new_index").GetInt32(), name); break;
            }
        }
        return sheets.ToArray();
    }

    internal static Plan Prepare(string[] sheets, bool[] visible, string? preferredActive)
    {
        if (sheets.Length == 0 || sheets.Length != visible.Length ||
            sheets.Distinct(StringComparer.OrdinalIgnoreCase).Count() != sheets.Length)
            throw new InvalidDataException("Invalid business worksheet roster");
        int active = Array.FindIndex(sheets, s => s == preferredActive);
        if (active < 0 || !visible[active]) active = Array.FindIndex(visible, v => v);
        if (active < 0) throw new InvalidDataException("No visible business worksheet");
        return new Plan(sheets.ToArray(), active);
    }

    private static XDocument Read(ZipArchive zip, string part)
    {
        using var stream = (zip.GetEntry(part) ?? throw new InvalidDataException("Missing candidate part: " + part)).Open();
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }

    private static string Resolve(string owner, XElement relationship)
    {
        if ((string?)relationship.Attribute("TargetMode") == "External")
            throw new InvalidDataException("External worksheet relationship");
        var uri = new Uri(new Uri("http://package/" + owner), (string?)relationship.Attribute("Target") ?? "");
        if (uri.Host != "package" || uri.Query.Length != 0 || uri.Fragment.Length != 0)
            throw new InvalidDataException("Invalid worksheet relationship");
        string part = Uri.UnescapeDataString(uri.AbsolutePath.TrimStart('/'));
        if (!part.StartsWith("xl/worksheets/", StringComparison.Ordinal) || !part.EndsWith(".xml", StringComparison.Ordinal))
            throw new InvalidDataException("Unsupported worksheet part");
        return part;
    }

    private static string RelsPart(string part) => part[..(part.LastIndexOf('/') + 1)] + "_rels/" + part[(part.LastIndexOf('/') + 1)..] + ".rels";

    private static void VerifySystemSheet(ZipArchive zip, string part)
    {
        var sheet = Read(zip, part);
        if (sheet.Root?.Name != S + "worksheet") throw new InvalidDataException("Invalid system worksheet");
        // Unknown structures (drawings, comments, tables, hyperlinks, OLE, etc.)
        // make ownership ambiguous. Do not recursively delete their parts.
        var allowed = new HashSet<string> { "dimension", "sheetViews", "sheetFormatPr", "sheetData", "pageMargins" };
        if (sheet.Root.Elements().Any(e => e.Name.Namespace != S || !allowed.Contains(e.Name.LocalName)) ||
            zip.GetEntry(RelsPart(part)) is not null)
            throw new InvalidDataException("System worksheet contains unverified objects");
        var strings = zip.GetEntry("xl/sharedStrings.xml") is null ? Array.Empty<XElement>() : Read(zip, "xl/sharedStrings.xml").Root!.Elements(S + "si").ToArray();
        int populated = 0;
        foreach (var cell in sheet.Descendants(S + "c"))
        {
            if (cell.Element(S + "f") is not null) throw new InvalidDataException("System worksheet contains a formula");
            string text;
            if ((string?)cell.Attribute("t") == "s")
            {
                if (!int.TryParse(cell.Element(S + "v")?.Value, out int index) || index < 0 || index >= strings.Length)
                    throw new InvalidDataException("Invalid system shared string");
                text = string.Concat(strings[index].Descendants(S + "t").Select(t => t.Value));
            }
            else if ((string?)cell.Attribute("t") == "inlineStr") text = string.Concat(cell.Descendants(S + "t").Select(t => t.Value));
            else
            {
                if (cell.Element(S + "v") is not null) throw new InvalidDataException("System worksheet contains business values");
                continue;
            }
            if (text.Length == 0) continue;
            // A known SDK-generated notice, never an arbitrary text-only sheet.
            if (++populated > 1 || (string?)cell.Attribute("r") != "A5" ||
                !IsNotice(text))
                throw new InvalidDataException("System worksheet contains unrecognized data");
        }
    }

    private static void CheckReferences(ZipArchive zip, HashSet<string> removedParts, HashSet<int> removedIndices,
        string removedName, string[] businessSheets)
    {
        // Recognize direct qualifiers of retained business sheets. Leave
        // ambiguous forms (including 3-D and dynamic references) to the
        // conservative check against the actual removed name.
        var retainedQualifiers = businessSheets
            .Where(name => name.Contains(removedName, StringComparison.OrdinalIgnoreCase))
            .Select(name => "'" + name.Replace("'", "''") + "'!")
            .ToArray();
        bool MentionsRemoved(string text)
        {
            foreach (string qualifier in retainedQualifiers)
                text = text.Replace(qualifier, "!", StringComparison.Ordinal);
            return text.Contains(removedName, StringComparison.OrdinalIgnoreCase);
        }
        foreach (var entry in zip.Entries.Where(e => e.FullName.EndsWith(".xml", StringComparison.Ordinal) && !removedParts.Contains(e.FullName)))
        {
            // Inspect reference-bearing XML elements without changing business
            // parts. Dynamic references cannot safely prove non-use of a sheet.
            var document = Read(zip, entry.FullName);
            foreach (var element in document.Descendants())
            {
                bool formula = element.Name.LocalName is "f" or "formula" or "formula1" or "formula2" or "definedName" or "calculatedColumnFormula" or "totalsRowFormula";
                if (formula && (MentionsRemoved(element.Value) ||
                    Regex.IsMatch(element.Value, @"\bINDIRECT\s*\(", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)))
                    throw new InvalidDataException("Unsafe reference to system worksheet");
                if (element.Name == S + "definedName" && int.TryParse((string?)element.Attribute("localSheetId"), out int local) && removedIndices.Contains(local))
                    throw new InvalidDataException("System worksheet has a scoped defined name");
                if (element.Attributes().Any(a => (a.Name.LocalName is "location" or "sheet") &&
                    !businessSheets.Contains(a.Value, StringComparer.Ordinal) && MentionsRemoved(a.Value)))
                    throw new InvalidDataException("Unsafe worksheet object reference");
            }
        }
        foreach (var entry in zip.Entries.Where(e => e.FullName.EndsWith(".rels", StringComparison.Ordinal) && e.FullName != "xl/_rels/workbook.xml.rels"))
        {
            var document = Read(zip, entry.FullName);
            int marker = entry.FullName.LastIndexOf("_rels/", StringComparison.Ordinal);
            string owner = marker < 0 ? "" : entry.FullName[..marker] + entry.FullName[(marker + 6)..^5];
            foreach (var relation in document.Root!.Elements())
            {
                if ((string?)relation.Attribute("TargetMode") == "External") continue;
                var uri = new Uri(new Uri("http://package/" + owner), (string?)relation.Attribute("Target") ?? "");
                if (removedParts.Contains(Uri.UnescapeDataString(uri.AbsolutePath.TrimStart('/'))))
                    throw new InvalidDataException("System worksheet part has another owner");
            }
        }
    }

    private static void Save(ZipArchive zip, string part, XDocument document)
    {
        zip.GetEntry(part)?.Delete();
        using var stream = zip.CreateEntry(part, CompressionLevel.Optimal).Open();
        document.Save(stream, SaveOptions.DisableFormatting);
    }

    private static void ClearUnusedNoticeStrings(ZipArchive zip, string[] keptParts, string[] removedParts)
    {
        if (zip.GetEntry("xl/sharedStrings.xml") is null) return;
        int[] Indices(string part) => Read(zip, part).Descendants(S + "c")
            .Where(c => (string?)c.Attribute("t") == "s").Select(c => int.Parse(c.Element(S + "v")!.Value)).ToArray();
        var used = keptParts.SelectMany(Indices).ToHashSet();
        var removed = removedParts.SelectMany(Indices).ToArray();
        if (removed.Length == 0) return;
        var document = Read(zip, "xl/sharedStrings.xml");
        var strings = document.Root!.Elements(S + "si").ToArray();
        foreach (var index in removed.Distinct())
            if (!used.Contains(index))
            {
                // Keep indices stable for every business cell. An empty slot
                // replaces only an unreferenced, verified SDK notice.
                string text = string.Concat(strings[index].Descendants(S + "t").Select(t => t.Value));
                if (text.Length != 0 && !IsNotice(text))
                    throw new InvalidDataException("Unverified system shared string");
                strings[index].ReplaceNodes(new XElement(S + "t", ""));
            }
        if (long.TryParse((string?)document.Root.Attribute("count"), out long count))
            document.Root.SetAttributeValue("count", Math.Max(0, count - removed.Length));
        Save(zip, "xl/sharedStrings.xml", document);
    }

    internal static Result Clean(string candidate, Plan plan, string? sdkAppendedSheet)
    {
        // ZipArchive.Update may rewrite headers even when validation throws.
        // Validate/mutate a private sibling, replacing only this owned candidate
        // after the independent package check succeeds.
        string working = candidate + ".sanitation-" + Guid.NewGuid().ToString("N") + ".xlsx";
        try
        {
            File.Copy(candidate, working, false);
            var result = CleanCopy(working, plan, sdkAppendedSheet);
            File.Move(working, candidate, true);
            return result;
        }
        finally { if (File.Exists(working)) File.Delete(working); }
    }

    private static Result CleanCopy(string candidate, Plan plan, string? sdkAppendedSheet)
    {
        string[] removedNames;
        bool viewCorrected = false;
        using (var zip = ZipFile.Open(candidate, ZipArchiveMode.Update))
        {
            if (zip.Entries.Select(e => e.FullName).Distinct(StringComparer.Ordinal).Count() != zip.Entries.Count)
                throw new InvalidDataException("Duplicate candidate ZIP entries");
            var book = Read(zip, "xl/workbook.xml");
            var rels = Read(zip, "xl/_rels/workbook.xml.rels");
            var types = Read(zip, "[Content_Types].xml");
            var sheets = book.Root!.Element(S + "sheets")!.Elements(S + "sheet").ToArray();
            var byId = rels.Root!.Elements(P + "Relationship").ToDictionary(e => (string)e.Attribute("Id")!);
            int expectedCount = plan.Sheets.Length + (sdkAppendedSheet is null ? 0 : 1);
            var kept = sheets.Take(plan.Sheets.Length).ToArray();
            if (sheets.Length != expectedCount ||
                !kept.Select(s => (string)s.Attribute("name")!).SequenceEqual(plan.Sheets) ||
                plan.ActiveIndex < 0 || plan.ActiveIndex >= kept.Length ||
                ((string?)kept[plan.ActiveIndex].Attribute("state") ?? "visible") != "visible")
                throw new InvalidDataException("Candidate business roster or active view changed");
            var removed = sheets.Skip(kept.Length).Select((sheet, index) => (sheet, index: kept.Length + index)).ToArray();
            var parts = sheets.Select(s => Resolve("xl/workbook.xml", byId[(string)s.Attribute(R + "id")!])).ToArray();
            if (parts.Distinct(StringComparer.Ordinal).Count() != parts.Length)
                throw new InvalidDataException("Shared worksheet part");
            var removedParts = removed.Select(p => parts[p.index]).ToHashSet(StringComparer.Ordinal);
            foreach (var (sheet, index) in removed)
            {
                string name = (string)sheet.Attribute("name")!;
                bool observedName = name == sdkAppendedSheet &&
                    (name == SystemSheet || Regex.IsMatch(name, @"\AEvaluation Warning \([1-9][0-9]*\)\z", RegexOptions.CultureInvariant));
                // A case-only duplicate can exist in raw SDK XML. The exact
                // business prefix and the appended SDK object distinguish it.
                if (!observedName || plan.Sheets.Contains(name, StringComparer.Ordinal) ||
                    (string?)byId[(string)sheet.Attribute(R + "id")!].Attribute("Type") != WorksheetType)
                    throw new InvalidDataException($"Unrecognized extra worksheet after SDK Save (in memory: '{sdkAppendedSheet}', in package: '{name}')");
                VerifySystemSheet(zip, parts[index]);
            }
            var removedIndices = removed.Select(p => p.index).ToHashSet();
            if (removed.Length > 0) CheckReferences(zip, removedParts, removedIndices, (string)removed[0].sheet.Attribute("name")!, plan.Sheets);
            if (removed.Length > 0) ClearUnusedNoticeStrings(zip, parts.Where(part => !removedParts.Contains(part)).ToArray(), removedParts.ToArray());
            removedNames = removed.Select(p => (string)p.sheet.Attribute("name")!).ToArray();
            foreach (var (sheet, index) in removed)
            {
                byId[(string)sheet.Attribute(R + "id")!].Remove();
                sheet.Remove();
                zip.GetEntry(parts[index])!.Delete();
                foreach (var type in types.Root!.Elements().Where(e => (string?)e.Attribute("PartName") == "/" + parts[index]).ToArray()) type.Remove();
            }
            foreach (var name in book.Descendants(S + "definedName"))
                if (int.TryParse((string?)name.Attribute("localSheetId"), out int local))
                    name.SetAttributeValue("localSheetId", local - removedIndices.Count(i => i < local));
            var views = book.Root.Element(S + "bookViews");
            if (views is null) { views = new XElement(S + "bookViews"); book.Root.Element(S + "sheets")!.AddBeforeSelf(views); }
            if (!views.Elements(S + "workbookView").Any()) views.Add(new XElement(S + "workbookView"));
            foreach (var view in views.Elements(S + "workbookView"))
            {
                viewCorrected |= (string?)view.Attribute("activeTab") != plan.ActiveIndex.ToString() ||
                    (string?)view.Attribute("firstSheet") != plan.ActiveIndex.ToString();
                view.SetAttributeValue("activeTab", plan.ActiveIndex);
                view.SetAttributeValue("firstSheet", plan.ActiveIndex);
            }
            for (int i = 0; i < kept.Length; i++)
            {
                var part = parts[Array.IndexOf(sheets, kept[i])];
                var sheet = Read(zip, part);
                bool changed = false;
                foreach (var view in sheet.Descendants(S + "sheetView"))
                {
                    string selected = i == plan.ActiveIndex ? "1" : "0";
                    if (((string?)view.Attribute("tabSelected") ?? "0") != selected)
                    { view.SetAttributeValue("tabSelected", selected); changed = true; }
                }
                if (changed) { Save(zip, part, sheet); viewCorrected = true; }
            }
            if (zip.GetEntry("docProps/app.xml") is not null && removed.Length > 0)
            {
                var app = Read(zip, "docProps/app.xml");
                var headings = app.Root!.Elements().FirstOrDefault(e => e.Name.LocalName == "HeadingPairs")?.Element(Vt + "vector")?.Elements(Vt + "variant").ToArray();
                var titles = app.Root.Elements().FirstOrDefault(e => e.Name.LocalName == "TitlesOfParts")?.Element(Vt + "vector");
                if (headings is not null && titles is not null)
                {
                    int offset = 0;
                    for (int i = 0; i + 1 < headings.Length; i += 2)
                    {
                        var count = headings[i + 1].Element(Vt + "i4");
                        if (count is null || !int.TryParse(count.Value, out int n)) throw new InvalidDataException("Invalid extended worksheet properties");
                        if (headings[i].Element(Vt + "lpstr")?.Value == "Worksheets")
                        {
                            var names = titles.Elements().Skip(offset).Take(n).ToArray();
                            if (!names.Select(e => e.Value).SequenceEqual(sheets.Select(s => (string)s.Attribute("name")!)))
                                throw new InvalidDataException("Extended worksheet roster differs");
                            foreach (var index in removedIndices) names[index].Remove();
                            count.Value = kept.Length.ToString();
                            titles.SetAttributeValue("size", titles.Elements().Count());
                            break;
                        }
                        offset += n;
                    }
                    Save(zip, "docProps/app.xml", app);
                }
            }
            Save(zip, "xl/workbook.xml", book);
            if (removed.Length > 0) { Save(zip, "xl/_rels/workbook.xml.rels", rels); Save(zip, "[Content_Types].xml", types); }
        }
        Verify(candidate, plan);
        return new Result("output_sanitation", removedNames, removedNames.Length, plan.ActiveIndex, viewCorrected, true);
    }

    // Reopen the physical package without Aspose (which can add objects again).
    internal static void Verify(string candidate, Plan plan)
    {
        using var zip = ZipFile.OpenRead(candidate);
        var book = Read(zip, "xl/workbook.xml");
        var sheets = book.Root!.Element(S + "sheets")!.Elements(S + "sheet").ToArray();
        if (!sheets.Select(s => (string)s.Attribute("name")!).SequenceEqual(plan.Sheets) ||
            !book.Descendants(S + "workbookView").Any() ||
            book.Descendants(S + "workbookView").Any(v => (int?)v.Attribute("activeTab") != plan.ActiveIndex || (int?)v.Attribute("firstSheet") != plan.ActiveIndex) ||
            ((string?)sheets[plan.ActiveIndex].Attribute("state") ?? "visible") != "visible")
            throw new InvalidDataException("Final output sanitation verification failed");
        var rels = Read(zip, "xl/_rels/workbook.xml.rels").Root!.Elements(P + "Relationship").ToDictionary(e => (string)e.Attribute("Id")!);
        var parts = sheets.Select(sh => Resolve("xl/workbook.xml", rels[(string)sh.Attribute(R + "id")!])).ToHashSet(StringComparer.Ordinal);
        if (rels.Values.Where(e => (string?)e.Attribute("Type") == WorksheetType).Any(e => !parts.Contains(Resolve("xl/workbook.xml", e))))
            throw new InvalidDataException("Orphan worksheet relationship");
        foreach (var (sheet, index) in sheets.Select((sh, i) => (sh, i)))
        {
            var relation = rels[(string)sheet.Attribute(R + "id")!];
            if ((string?)relation.Attribute("Type") != WorksheetType) throw new InvalidDataException("Unsupported business worksheet type");
            var document = Read(zip, Resolve("xl/workbook.xml", relation));
            if (document.Root?.Name != S + "worksheet" || document.Descendants(S + "sheetView").Any(v =>
                ((string?)v.Attribute("tabSelected") ?? "0") != (index == plan.ActiveIndex ? "1" : "0")))
                throw new InvalidDataException("Final worksheet view verification failed");
        }
        foreach (var name in book.Descendants(S + "definedName"))
            if (int.TryParse((string?)name.Attribute("localSheetId"), out int local) && (local < 0 || local >= sheets.Length))
                throw new InvalidDataException("Invalid final scoped name");
        if (!plan.Sheets.Contains(SystemSheet) && zip.GetEntry("docProps/app.xml") is not null &&
            Read(zip, "docProps/app.xml").Root!.Elements().Where(e => e.Name.LocalName == "TitlesOfParts")
                .Descendants(Vt + "lpstr").Any(e => e.Value == SystemSheet))
            throw new InvalidDataException("System worksheet remains in extended properties");
    }
}
