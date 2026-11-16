using Aspose.Cells;
using System.Text.Json;

internal static class FormulaCachePolicy
{
    // Use the SDK parser for dependencies; this does not calculate formulas.
    // Both scalar and finite range writes participate. Results needing refresh
    // are cleared and marked for calculation on open; never compared as data.
    public static HashSet<(string Sheet, string Cell)> Invalidated(Workbook workbook, JsonElement operations)
    {
        var writes = new CellWriteIndex(operations);
        var invalid = new HashSet<(string, string)>();
        if (!writes.HasWrites) return invalid;
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
                        if (area.IsExternalLink || (area.SheetNames ?? new[] { area.SheetName }).Any(name =>
                            writes.Intersects(name, area.StartRow, area.StartColumn, area.EndRow, area.EndColumn))) { stale = true; break; }
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
