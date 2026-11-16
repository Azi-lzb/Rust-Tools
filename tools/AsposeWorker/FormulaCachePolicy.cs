using Aspose.Cells;
using System.Text.Json;

internal static class FormulaCachePolicy
{
    // Use the SDK parser for dependencies; this does not calculate formulas.
    // This mirrors the existing Worker's explicit cell-write invalidation scope.
    public static HashSet<(string Sheet, string Cell)> Invalidated(Workbook workbook, JsonElement operations)
    {
        var writes = new List<(string Sheet, int Row, int Column)>();
        foreach (var outer in operations.EnumerateArray())
        {
            var op = outer; string kind = op.GetProperty("op").GetString()!;
            if (kind == "edit_plan_operation") { op = outer.GetProperty("operation"); kind = op.GetProperty("operation").GetString()!; }
            if (kind is not "set_value" and not "set_formula" and not "set_cell_value" and not "set_formula_text") continue;
            CellsHelper.CellNameToIndex(op.GetProperty("cell").GetString()!, out int row, out int column);
            writes.Add((op.GetProperty("sheet").GetString()!, row, column));
        }
        var invalid = new HashSet<(string, string)>();
        if (writes.Count == 0) return invalid;
        foreach (Worksheet sheet in workbook.Worksheets) foreach (Cell cell in sheet.Cells)
        {
            if (!cell.IsFormula) continue;
            bool stale = false;
            try
            {
                // Without a proven dynamic target, preserving its old cache
                // after a write could publish a stale business result.
                if (cell.Formula.Contains("INDIRECT(", StringComparison.OrdinalIgnoreCase) || cell.Formula.Contains("OFFSET(", StringComparison.OrdinalIgnoreCase)) stale = true;
                else
                {
                    var areas = cell.GetPrecedents();
                    if (areas is not null) foreach (ReferredArea area in areas)
                        if (area.IsExternalLink || writes.Any(w => (area.SheetNames ?? new[] { area.SheetName }).Contains(w.Sheet) &&
                            w.Row >= area.StartRow && w.Row <= area.EndRow && w.Column >= area.StartColumn && w.Column <= area.EndColumn)) stale = true;
                }
            }
            catch { stale = true; }
            if (stale) invalid.Add((sheet.Name, cell.Name));
        }
        return invalid;
    }
    public static void PreserveStoredStyles(Workbook workbook)
    {
        // Aspose can read an inherited/default xf correctly but omit its font
        // when saving an openpyxl-authored source. Force the effective style
        // to be explicitly applied; the Rust validator still verifies it.
        foreach (Worksheet sheet in workbook.Worksheets) foreach (Cell cell in sheet.Cells)
            cell.SetStyle(cell.GetStyle(), new StyleFlag { All = true });
    }
}
