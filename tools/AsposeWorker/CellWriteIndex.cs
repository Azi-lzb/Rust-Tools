using Aspose.Cells;
using System.Text.Json;

// Write footprints shared by dependency checking and lexical preservation.
// Each range query visits written columns, then seeks rows in a sorted set;
// it never scans every planned append cell for every existing formula.
internal sealed class CellWriteIndex
{
    private readonly Dictionary<string, SortedDictionary<int, SortedSet<int>>> sheets = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, List<(int R1, int C1, int R2, int C2)>> rectangles = new(StringComparer.OrdinalIgnoreCase);
    public bool StationaryScalarPlan { get; private set; } = true;
    public bool HasWrites => sheets.Count != 0 || rectangles.Count != 0;

    public CellWriteIndex(JsonElement operations)
    {
        foreach (var outer in operations.EnumerateArray())
        {
            var op = outer; string kind = op.GetProperty("op").GetString()!;
            if (kind == "edit_plan_operation") { op = op.GetProperty("operation"); kind = op.GetProperty("operation").GetString()!; }
            if (kind is "rename_sheet" or "delete_sheet" or "insert_rows" or "delete_rows" or "insert_columns" or "delete_columns" or "copy_range" or "import_range")
                StationaryScalarPlan = false;
            if (kind is "copy_range" or "import_range")
            {
                if (!op.GetProperty("copy_values").GetBoolean() && !op.GetProperty("copy_formulas").GetBoolean()) continue;
                var range = op.GetProperty("source_range").GetString()!.Replace("$", "").Split(':');
                CellsHelper.CellNameToIndex(range[0], out int r1, out int c1);
                CellsHelper.CellNameToIndex(range[^1], out int r2, out int c2);
                CellsHelper.CellNameToIndex(op.GetProperty("target_start").GetString()!.Replace("$", ""), out int tr, out int tc);
                string target = op.GetProperty("target_sheet").GetString()!;
                if (!rectangles.TryGetValue(target, out var ranges)) rectangles[target] = ranges = new();
                ranges.Add((tr, tc, tr + r2 - r1, tc + c2 - c1));
                continue;
            }
            if (kind is not "set_value" and not "set_formula" and not "set_cell_value" and not "set_formula_text") continue;
            CellsHelper.CellNameToIndex(op.GetProperty("cell").GetString()!.Replace("$", ""), out int row, out int column);
            string sheet = op.GetProperty("sheet").GetString()!;
            if (!sheets.TryGetValue(sheet, out var columns)) sheets[sheet] = columns = new();
            if (!columns.TryGetValue(column, out var rows)) columns[column] = rows = new();
            rows.Add(row);
        }
    }

    public bool Contains(string sheet, string address)
    {
        CellsHelper.CellNameToIndex(address.Replace("$", ""), out int row, out int column);
        return sheets.TryGetValue(sheet, out var columns) && columns.TryGetValue(column, out var rows) && rows.Contains(row);
    }

    public bool Intersects(string sheet, int r1, int c1, int r2, int c2)
    {
        if (rectangles.TryGetValue(sheet, out var ranges) && ranges.Any(r => r.R1 <= r2 && r.R2 >= r1 && r.C1 <= c2 && r.C2 >= c1)) return true;
        if (!sheets.TryGetValue(sheet, out var columns)) return false;
        foreach (var (column, rows) in columns)
        {
            if (column < c1) continue;
            if (column > c2) break;
            using var hits = rows.GetViewBetween(r1, r2).GetEnumerator();
            if (hits.MoveNext()) return true;
        }
        return false;
    }
}
