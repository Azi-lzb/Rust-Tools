using Aspose.Cells;
using System.Diagnostics;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

static object?[][] LoadData(string path) => JsonSerializer.Deserialize<JsonElement[][]>(File.ReadAllText(path))!
    .Select(row => row.Select(v => v.ValueKind switch {
        JsonValueKind.String => (object?)v.GetString(), JsonValueKind.Number => v.GetInt64(),
        JsonValueKind.True => true, JsonValueKind.False => false, JsonValueKind.Null => null,
        _ => throw new Exception("fixture type") }).ToArray()).ToArray();

static string Digest(object?[][] data) {
    using var h = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
    foreach (var row in data) foreach (var v in row) {
        string s = v switch {
            null => "E\n", bool b => b ? "B1\n" : "B0\n", string t => $"S{t}\n",
            double n when n == Math.Truncate(n) => $"N{n.ToString("0", CultureInfo.InvariantCulture)}\n",
            long n => $"N{n}\n", int n => $"N{n}\n", _ => throw new Exception($"unexpected cell {v.GetType()}") };
        h.AppendData(Encoding.UTF8.GetBytes(s));
    } return Convert.ToHexString(h.GetHashAndReset()).ToLowerInvariant();
}
static object?[][] Extract(Workbook wb, object?[][] shape) {
    var cells = wb.Worksheets["Data"].Cells;
    return Enumerable.Range(0, shape.Length).Select(r => Enumerable.Range(0, shape[0].Length)
        .Select(c => cells[r,c].Value).ToArray()).ToArray();
}
var sets = new Dictionary<string,object?[][]>();
int workbookCount = 0;
Console.WriteLine(JsonSerializer.Serialize(new {ready=true, engine="aspose", version=CellsHelper.GetVersion(), pid=Environment.ProcessId, stopwatch_frequency=Stopwatch.Frequency}));
string? line;
while ((line = Console.ReadLine()) != null) {
    try {
        using var doc = JsonDocument.Parse(line); var q = doc.RootElement;
        var op = q.GetProperty("op").GetString();
        if (op == "init") {
            foreach (var f in q.GetProperty("fixtures").EnumerateArray()) sets.Add(f.GetProperty("id").GetString()!, LoadData(f.GetProperty("data").GetString()!));
            Console.WriteLine("{\"ok\":true}"); continue;
        }
        if (++workbookCount > 90) throw new Exception("BLOCKED_LICENSE conservative 90-workbook budget; no automatic restart");
        var data = sets[q.GetProperty("id").GetString()!];
        string? output = q.TryGetProperty("output",out var o) ? o.GetString() : null;
        if (output != null && File.Exists(output)) throw new Exception("refuse existing output");
        object?[][] actual = data;
        double ms;
        var start = Stopwatch.GetTimestamp();
        using (var wb = op == "write" ? new Workbook() : new Workbook(q.GetProperty("input").GetString())) {
            if (op == "write") {
                wb.Worksheets[0].Name = "Data";
                var cells = wb.Worksheets["Data"].Cells;
                for (int r=0;r<data.Length;r++) for(int c=0;c<data[r].Length;c++) {
                    switch(data[r][c]) {
                        case null: break;
                        case string s: cells[r,c].PutValue(s); break;
                        case long n: cells[r,c].PutValue((double)n); break;
                        case bool b: cells[r,c].PutValue(b); break;
                        default: throw new Exception("fixture type");
                    }
                }
            }
            if (op == "read") actual = Extract(wb,data);
            else if (op is "write" or "roundtrip") wb.Save(output,SaveFormat.Xlsx);
            else throw new Exception("unknown op");
            // Include disposal in the timed interval, as in native library calls.
        }
        ms = Stopwatch.GetElapsedTime(start).TotalMilliseconds;
        Console.WriteLine(JsonSerializer.Serialize(new {ok=true, processing_ms=ms, digest=Digest(actual), rows=data.Length, cols=data[0].Length, workbook_count=workbookCount}));
    } catch(Exception ex) { Console.WriteLine(JsonSerializer.Serialize(new {ok=false,error=ex.ToString(),workbook_count=workbookCount})); }
}
