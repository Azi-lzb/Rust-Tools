using Aspose.Cells;
using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;

// Excel primitives only. Selection, deduplication and report rules remain in Rust.
// No calculation is performed by Modify (including range imports).
internal sealed class PlanOperations : IDisposable
{
    private readonly Dictionary<string, Workbook> imports = new(OperatingSystem.IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal);
    private readonly Dictionary<string, StoredValues> importValues = new(OperatingSystem.IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal);
    private StoredValues? inputValues;
    public void SetInput(string path, Workbook workbook) => inputValues = new StoredValues(path, workbook);
    public static bool NeedsInputValues(JsonElement operations) => operations.EnumerateArray().Any(o =>
        S(o, "op") == "edit_plan_operation" && S(o.GetProperty("operation"), "operation") == "copy_range");
    private static string S(JsonElement p, string key) => p.GetProperty(key).GetString() ?? throw new ArgumentException(key);
    private static int I(JsonElement p, string key) => p.GetProperty(key).GetInt32();
    private static bool B(JsonElement p, string key) => p.TryGetProperty(key, out var v) && v.ValueKind == JsonValueKind.True;
    private static bool Present(JsonElement p, string key) => p.TryGetProperty(key, out var v) && v.ValueKind != JsonValueKind.Null;
    private static bool NumericZero(object? value) => value is byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal
        && Convert.ToDouble(value, CultureInfo.InvariantCulture) == 0.0;
    private static string NumberKey(string text)
    {
        var m = Regex.Match(text, "^([+-]?)([0-9]+)(?:\\.([0-9]*))?(?:[eE]([+-]?[0-9]+))?$");
        if (!m.Success) throw new ArgumentException("Invalid numeric text");
        string digits = (m.Groups[2].Value + m.Groups[3].Value).TrimStart('0');
        if (digits.Length == 0) return "0";
        int trailing = digits.Length - digits.TrimEnd('0').Length;
        int exponent = (m.Groups[4].Success ? int.Parse(m.Groups[4].Value, CultureInfo.InvariantCulture) : 0) - m.Groups[3].Length + trailing;
        return (m.Groups[1].Value == "-" ? "-" : "") + digits.TrimEnd('0') + "e" + exponent;
    }
    private static Worksheet Sheet(Workbook wb, string name) => wb.Worksheets[name] ?? throw new ArgumentException("Worksheet not found: " + name);
    private static (int Row, int Column) Address(string address)
    {
        if (!Regex.IsMatch(address, "^\\$?[A-Z]{1,3}\\$?[1-9][0-9]{0,6}$")) throw new ArgumentException("Invalid A1 address");
        address = address.Replace("$", "");
        CellsHelper.CellNameToIndex(address, out int row, out int column);
        if (row < 0 || row >= 1048576 || column < 0 || column >= 16384) throw new ArgumentException("Address outside XLSX bounds");
        return (row, column);
    }
    private static (int Row, int Column, int Rows, int Columns) Bounds(string address)
    {
        var parts = address.Split(':');
        if (parts.Length is < 1 or > 2) throw new ArgumentException("Invalid finite range");
        var a = Address(parts[0]); var z = Address(parts[^1]);
        if (z.Row < a.Row || z.Column < a.Column) throw new ArgumentException("Reversed range");
        return (a.Row, a.Column, z.Row - a.Row + 1, z.Column - a.Column + 1);
    }
    public static string[] ImportPaths(JsonElement operations) => operations.EnumerateArray()
        .Where(o => S(o, "op") == "edit_plan_operation")
        .Select(o => o.GetProperty("operation"))
        .Where(o => S(o, "operation") == "import_range")
        .Select(o => { string path = S(o, "source_workbook"); if (!Path.IsPathFullyQualified(path)) throw new ArgumentException("Absolute import path required"); return Path.GetFullPath(path); })
        .Distinct(OperatingSystem.IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal).ToArray();
    public void LoadImports(string[] paths)
    {
        foreach (string path in paths)
        {
            if (!Path.IsPathFullyQualified(path) || !Path.GetExtension(path).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException("Import sources must be absolute XLSX paths");
            var wb = new Workbook(path);
            if (wb.HasMacro) { wb.Dispose(); throw new NotSupportedException("Macro import not validated"); }
            imports.Add(path, wb);
            importValues.Add(path, new StoredValues(path, wb));
        }
    }
    public static Dictionary<string, string> FinalIdentities(IEnumerable<string> original, JsonElement operations)
    {
        var identities = original.ToDictionary(s => s, s => s, StringComparer.Ordinal);
        foreach (var outer in operations.EnumerateArray())
        {
            if (S(outer, "op") != "edit_plan_operation") continue;
            var p = outer.GetProperty("operation");
            if (S(p, "operation") == "rename_sheet")
                foreach (var key in identities.Keys.ToArray()) if (identities[key] == S(p, "old_name")) identities[key] = S(p, "new_name");
            if (S(p, "operation") == "delete_sheet")
                foreach (var key in identities.Keys.ToArray()) if (identities[key] == S(p, "name")) identities.Remove(key);
        }
        return identities;
    }
    private static void PutValue(Cell cell, JsonElement value)
    {
        string type = S(value, "type");
        switch (type)
        {
            case "text": cell.PutValue(S(value, "value"), false); break;
            case "blank": cell.PutValue((string?)null); break;
            case "boolean": cell.PutValue(value.GetProperty("value").GetBoolean()); break;
            case "decimal":
                string raw = S(value, "value");
                // Excel numeric cells are doubles. Refuse rather than silently
                // coerce higher-precision decimal text into a different value.
                double number = double.Parse(raw, CultureInfo.InvariantCulture);
                if (!double.IsFinite(number)) throw new NotSupportedException("Nonfinite Excel number");
                string exact = NumberKey(raw);
                if (exact != NumberKey(number.ToString("G15", CultureInfo.InvariantCulture)) &&
                    (number % 1.0 == 0 || exact != NumberKey(number.ToString("R", CultureInfo.InvariantCulture))))
                    throw new NotSupportedException("Numeric text cannot be round-tripped, or integer exceeds 15 significant digits");
                cell.PutValue(number); break;
            case "date_time":
                var date = value.GetProperty("value");
                double serial = double.Parse(S(date, "serial"), CultureInfo.InvariantCulture);
                if (!double.IsFinite(serial)) throw new ArgumentException("Invalid date serial");
                cell.PutValue(serial);
                var style = cell.GetStyle(); style.Custom = S(date, "number_format"); cell.SetStyle(style, true); break;
            case "error":
                string error = S(value, "value");
                if (!new[] { "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A" }.Contains(error))
                    throw new NotSupportedException("Unvalidated Excel error type");
                cell.PutValue(error, true); // Serialized type is independently verified by Rust.
                if (cell.Type != CellValueType.IsError) throw new NotSupportedException("SDK did not preserve Excel error type");
                break;
            default: throw new NotSupportedException("Unsupported value type: " + type);
        }
    }
    public void Apply(JsonElement p, Workbook wb)
    {
        string kind = S(p, "operation");
        Worksheet Ws() => Sheet(wb, S(p, "sheet"));
        switch (kind)
        {
            case "set_cell_value": PutValue(Ws().Cells[S(p, "cell")], p.GetProperty("value")); break;
            case "set_comment":
                var ws = Ws(); string at = S(p, "cell"); ws.Comments.RemoveAt(at);
                if (Present(p, "text")) { int i = ws.Comments.Add(at); ws.Comments[i].Note = S(p, "text"); ws.Comments[i].Author = Present(p, "author") ? S(p, "author") : "pytools"; } break;
            case "set_column_width": Ws().Cells.SetColumnWidth(I(p, "column_index") - 1, p.GetProperty("width").GetDouble()); break;
            case "set_row_height": Ws().Cells.SetRowHeight(I(p, "row_index") - 1, p.GetProperty("height").GetDouble()); break;
            // Openpyxl's row/column primitives do not rewrite formula references.
            // Keep that contract here; the plan validator checks preserved refs.
            case "delete_rows": Ws().Cells.DeleteRows(I(p, "first_row") - 1, I(p, "count"), false); break;
            case "insert_rows": Ws().Cells.InsertRows(I(p, "first_row") - 1, I(p, "count"), false); break;
            case "delete_columns": Ws().Cells.DeleteColumns(I(p, "first_column") - 1, I(p, "count"), false); break;
            case "insert_columns": Ws().Cells.InsertColumns(I(p, "first_column") - 1, I(p, "count"), false); break;
            case "rename_sheet": Sheet(wb, S(p, "old_name")).Name = S(p, "new_name"); break;
            case "create_sheet": wb.Worksheets.Add(S(p, "name")); break;
            case "create_sheet_at": wb.Worksheets.Add(S(p, "name")); Sheet(wb, S(p, "name")).MoveTo(I(p, "index")); break;
            case "move_sheet": Ws().MoveTo(I(p, "new_index")); break;
            case "delete_sheet":
                if (wb.Worksheets.Count <= 1) throw new ArgumentException("Cannot delete the only worksheet");
                wb.Worksheets.RemoveAt(S(p, "name")); break;
            case "set_freeze_panes":
                if (!Present(p, "cell")) Ws().UnFreezePanes();
                else { var a = Address(S(p, "cell")); Ws().FreezePanes(a.Row, a.Column, a.Row, a.Column); } break;
            case "set_auto_filter":
                if (Present(p, "range")) Ws().AutoFilter.Range = S(p, "range"); else Ws().RemoveAutoFilter(); break;
            case "set_workbook_calculation_properties":
                if (Present(p, "full_calc_on_load")) wb.Settings.FormulaSettings.CalculateOnOpen = B(p, "full_calc_on_load");
                if (Present(p, "force_full_calc")) wb.Settings.FormulaSettings.ForceFullCalculation = B(p, "force_full_calc"); break;
            case "set_defined_name": SetName(p, wb); break;
            case "set_page_settings": SetPage(p.GetProperty("settings"), Ws().PageSetup); break;
            case "set_margins": SetMargins(p.GetProperty("margins"), Ws().PageSetup); break;
            case "copy_range": case "import_range": CopyRange(p, wb); break;
            default: throw new NotSupportedException("Unsupported ExcelEditPlan primitive: " + kind);
        }
    }
    private static void SetName(JsonElement p, Workbook wb)
    {
        string name = S(p, "name"); int scope = Present(p, "local_sheet") ? Sheet(wb, S(p, "local_sheet")).Index + 1 : 0;
        for (int i = wb.Worksheets.Names.Count - 1; i >= 0; i--)
            if (wb.Worksheets.Names[i].Text == name && wb.Worksheets.Names[i].SheetIndex == scope) wb.Worksheets.Names.RemoveAt(i);
        if (!Present(p, "formula")) return;
        int index = wb.Worksheets.Names.Add(name); var n = wb.Worksheets.Names[index]; n.SheetIndex = scope;
        string formula = S(p, "formula"); n.RefersTo = formula.StartsWith('=') ? formula : "=" + formula;
    }
    private static void SetPage(JsonElement p, PageSetup s)
    {
        if (Present(p, "orientation")) s.Orientation = S(p, "orientation") == "landscape" ? PageOrientationType.Landscape : PageOrientationType.Portrait;
        if (Present(p, "fit_to_page")) s.IsPercentScale = !B(p, "fit_to_page");
        if (Present(p, "fit_to_width")) s.FitToPagesWide = I(p, "fit_to_width");
        if (Present(p, "fit_to_height")) s.FitToPagesTall = I(p, "fit_to_height");
        if (Present(p, "paper_size")) s.PaperSize = (PaperSizeType)I(p, "paper_size");
        if (Present(p, "print_area")) s.PrintArea = S(p, "print_area");
        if (Present(p, "horizontal_centered")) s.CenterHorizontally = B(p, "horizontal_centered");
        if (Present(p, "vertical_centered")) s.CenterVertically = B(p, "vertical_centered");
        if (Present(p, "footer")) s.SetFooter(1, S(p, "footer"));
    }
    private static void SetMargins(JsonElement p, PageSetup s)
    {
        // Aspose's API margins are centimetres; plan / OOXML values are inches.
        foreach (var v in p.EnumerateObject())
        {
            if (v.Value.ValueKind == JsonValueKind.Null) continue;
            double cm = v.Value.GetDouble() * 2.54;
            switch (v.Name) { case "left": s.LeftMargin = cm; break; case "right": s.RightMargin = cm; break; case "top": s.TopMargin = cm; break;
                case "bottom": s.BottomMargin = cm; break; case "header": s.HeaderMargin = cm; break; case "footer": s.FooterMargin = cm; break; }
        }
    }
    private sealed record Snapshot(int Row, int Column, object? Value, bool Error, string? Formula, Style Style, string? Comment, string? Author);
    private void CopyRange(JsonElement p, Workbook wb)
    {
        bool import = S(p, "operation") == "import_range";
        var sourceWb = import ? imports[Path.GetFullPath(S(p, "source_workbook"))] : wb;
        var cachedValues = import ? importValues[Path.GetFullPath(S(p, "source_workbook"))] : inputValues ?? throw new InvalidOperationException("Input cache snapshot missing");
        var source = Sheet(sourceWb, S(p, "source_sheet")); var target = Sheet(wb, S(p, "target_sheet"));
        var b = Bounds(S(p, "source_range")); var at = Address(S(p, "target_start"));
        if (at.Row + b.Rows > 1048576 || at.Column + b.Columns > 16384) throw new ArgumentException("Copy exceeds worksheet bounds");
        string styleMode = import ? S(p, "style_mode") : B(p, "copy_styles") ? "full" : "none";
        if (!new[] { "none", "basic", "full" }.Contains(styleMode)) throw new ArgumentException("Invalid style mode");
        var snapshots = new List<Snapshot>();
        for (int r = b.Row; r < b.Row + b.Rows; r++) for (int c = b.Column; c < b.Column + b.Columns; c++)
        {
            var cell = source.Cells[r, c];
            if (cell.IsArrayFormula && B(p, "copy_formulas")) throw new NotSupportedException("Array formula copying is not validated");
            if (cell.Value is null && !cell.IsFormula && styleMode == "none" && !B(p, "copy_comments") && !B(p, "materialize_blank_cells")) continue;
            var comment = source.Comments[r, c];
            var ownedStyle = wb.CreateStyle(); ownedStyle.Copy(cell.GetStyle());
            var stored = cell.IsFormula || cell.IsArrayFormula ? cachedValues.At(source.Name, cell.Name) : new StoredValues.Value(cell.Value, cell.Type == CellValueType.IsError);
            snapshots.Add(new(r, c, stored.Data, stored.Error, B(p, "copy_formulas") && cell.IsFormula ? cell.R1C1Formula : null, ownedStyle, comment?.Note, comment?.Author));
        }
        foreach (var item in snapshots)
        {
            var dst = target.Cells[at.Row + item.Row - b.Row, at.Column + item.Column - b.Column];
            if (styleMode == "full") dst.SetStyle(item.Style, new StyleFlag { All = true });
            else if (styleMode == "basic")
            {
                // Copy the same font/fill/alignment/number-format set as the
                // baseline. Border/protection intentionally remain at target.
                dst.SetStyle(item.Style, new StyleFlag { Font = true, CellShading = true, Alignments = true, NumberFormat = true });
            }
            if (B(p, "copy_formulas") && item.Formula is not null) dst.R1C1Formula = item.Formula;
            else if (B(p, "copy_values"))
            {
                object? value = item.Value;
                if (B(p, "hide_zero") && NumericZero(value)) value = null;
                if (item.Error && value is string error) dst.PutValue(error, true);
                else dst.PutValue(value);
            }
            if (B(p, "copy_comments"))
            {
                target.Comments.RemoveAt(dst.Name);
                if (item.Comment is not null) { int i = target.Comments.Add(dst.Name); target.Comments[i].Note = item.Comment; target.Comments[i].Author = item.Author; }
            }
        }
        if (import && B(p, "copy_merges"))
            foreach (CellArea area in source.Cells.GetMergedAreas())
            {
                if (area.EndRow < b.Row || area.StartRow >= b.Row + b.Rows || area.EndColumn < b.Column || area.StartColumn >= b.Column + b.Columns) continue;
                if (area.StartRow < b.Row || area.EndRow >= b.Row + b.Rows || area.StartColumn < b.Column || area.EndColumn >= b.Column + b.Columns)
                    throw new NotSupportedException("Source range cuts through merged cells");
                target.Cells.Merge(at.Row + area.StartRow - b.Row, at.Column + area.StartColumn - b.Column, area.EndRow - area.StartRow + 1, area.EndColumn - area.StartColumn + 1);
            }
        if (!import) return;
        foreach (Row row in source.Cells.Rows)
        {
            if (!B(p, "copy_all_row_dimensions") && (row.Index < b.Row || row.Index >= b.Row + b.Rows)) continue;
            int to = at.Row + row.Index - b.Row;
            if (to < 0 || to >= 1048576) throw new ArgumentException("Row dimension out of bounds");
            if (B(p, "copy_row_heights")) target.Cells.SetRowHeight(to, row.Height);
            if (B(p, "copy_row_hidden")) target.Cells.Rows[to].IsHidden = row.IsHidden;
            if (B(p, "copy_outline_levels")) target.Cells.Rows[to].GroupLevel = row.GroupLevel;
        }
        foreach (Column col in source.Cells.Columns)
        {
            if (!B(p, "copy_all_column_dimensions") && (col.Index < b.Column || col.Index >= b.Column + b.Columns)) continue;
            int to = at.Column + col.Index - b.Column;
            if (to < 0 || to >= 16384) throw new ArgumentException("Column dimension out of bounds");
            if (B(p, "copy_column_widths")) target.Cells.SetColumnWidth(to, col.Width);
            if (B(p, "copy_column_hidden")) target.Cells.Columns[to].IsHidden = col.IsHidden;
            if (B(p, "copy_outline_levels")) target.Cells.Columns[to].GroupLevel = col.GroupLevel;
        }
    }
    public void Dispose() { foreach (var wb in imports.Values) wb.Dispose(); }
}
