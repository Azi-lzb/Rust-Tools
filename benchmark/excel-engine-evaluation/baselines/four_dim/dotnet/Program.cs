using Aspose.Cells;
using System.Diagnostics;
using System.Text.Json;
using System.Text;
using System.Security.Cryptography;
using System.Globalization;

static object? V(int r,int c)=>(c%5) switch {0=>(object)((long)r*20+c),1=>$"row-{r:D6}-col-{c:D2}",2=>(r+c)%2==0,3=>null,_=>-(long)(r+c)};
static void Put(Cell cell,object? v) {switch(v) {case null:break;case string s:cell.PutValue(s);break;case bool b:cell.PutValue(b);break;case long n:cell.PutValue((double)n);break;default:throw new Exception("fixture type");}}
static string Digest(List<object?[]> rows) {
    using var h=IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
    foreach(var row in rows) foreach(var v in row) {
        var s=v switch {null=>"E\n",bool b=>b?"B1\n":"B0\n",string t=>$"S{t}\n",double d when d==Math.Truncate(d)=>$"N{d.ToString("0",CultureInfo.InvariantCulture)}\n",int n=>$"N{n}\n",_=>throw new Exception($"Unexpected value type {v.GetType()}")};
        h.AppendData(Encoding.UTF8.GetBytes(s));
    }return Convert.ToHexString(h.GetHashAndReset()).ToLowerInvariant();
}
if(args.Length==2 && args[0]=="--fixture") {
    var root=args[1];using var spec=JsonDocument.Parse(File.ReadAllText(Path.Combine(root,"calc-spec.json")));
    using var wb=new Workbook();wb.Worksheets[0].Name="Data";var calc=wb.Worksheets.Add("Calc");
    int r=0;foreach(var row in spec.RootElement.GetProperty("data").EnumerateArray()) {int c=0;foreach(var v in row.EnumerateArray()) {if(v.ValueKind==JsonValueKind.String)wb.Worksheets["Data"].Cells[r,c].PutValue(v.GetString());else wb.Worksheets["Data"].Cells[r,c].PutValue(v.GetDouble());c++;}r++;}
    r=0;foreach(var row in spec.RootElement.GetProperty("formulas").EnumerateArray()) {int c=0;foreach(var f in row.EnumerateArray())calc.Cells[r,c++].SetFormula("="+f.GetString(),-9999.0);r++;}
    wb.Settings.FormulaSettings.CalculationMode=CalcModeType.Manual;
    foreach(var ext in new[]{"xls","xlsx"}) {var path=Path.Combine(root,"calc-stale."+ext);if(File.Exists(path))throw new Exception("refuse existing fixture");wb.Save(path,ext=="xls"?SaveFormat.Excel97To2003:SaveFormat.Xlsx);}
    Console.WriteLine("{\"ok\":true}");return;
}
var sdk=CellsHelper.GetVersion();
if(args.Length==2 && args[0]=="--inspect-xml") {using var wb=new Workbook(args[1]);Console.WriteLine(JsonSerializer.Serialize(new{custom_xml_parts=wb.CustomXmlParts.Count}));return;}
if(args.Length==1 && args[0]=="--empty-run") {Console.WriteLine(JsonSerializer.Serialize(new{ready=true,empty_run=true,version=sdk}));return;}
Console.WriteLine(JsonSerializer.Serialize(new{ready=true,engine="aspose",version=sdk,pid=Environment.ProcessId,clock_frequency=Stopwatch.Frequency}));
int count=0;string? line;
while((line=Console.ReadLine())!=null) {
    try {
        using var d=JsonDocument.Parse(line);var q=d.RootElement;var op=q.GetProperty("op").GetString();
        if(op=="ping") {Console.WriteLine("{\"ok\":true,\"processing_ms\":0}");continue;}
        // 7 conditions * (3 warmup + 10 measured) = 91 Workbook instances.
        if(++count>95)throw new Exception("BLOCKED_LICENSE: stop; no automatic process restart");
        string? output=q.TryGetProperty("output",out var o)?o.GetString():null;
        if(output!=null && File.Exists(output))throw new Exception("refuse existing output");
        List<object?[]> data=new();List<object?> values=new();double loadMs=0,calcMs=0;
        var start=Stopwatch.GetTimestamp();
        using(var wb=op=="write"?new Workbook():new Workbook(q.GetProperty("input").GetString())) {
            loadMs=Stopwatch.GetElapsedTime(start).TotalMilliseconds;
            if(op=="read") {
                var cells=wb.Worksheets["Data"].Cells;var rows=q.GetProperty("rows").GetInt32();var cols=q.GetProperty("cols").GetInt32();
                for(int r=0;r<rows;r++) {var row=new object?[cols];for(int c=0;c<cols;c++)row[c]=cells[r,c].Value;data.Add(row);}
            } else if(op=="write") {
                wb.Worksheets[0].Name="Data";var cells=wb.Worksheets["Data"].Cells;
                for(int r=0;r<50000;r++)for(int c=0;c<20;c++)Put(cells[r,c],V(r,c));wb.Save(output,SaveFormat.Xlsx);
            } else if(op=="calc") {
                var before=wb.Worksheets["Calc"].Cells;
                for(int r=0;r<256;r++)for(int c=0;c<8;c++)if(before[r,c].DoubleValue!=-9999.0)throw new Exception("Input cache was not stale before timed calculation");
                var t=Stopwatch.GetTimestamp();wb.CalculateFormula();calcMs=Stopwatch.GetElapsedTime(t).TotalMilliseconds;
                var cells=wb.Worksheets["Calc"].Cells;for(int r=0;r<256;r++)for(int c=0;c<8;c++)values.Add(cells[r,c].Value);
            } else if(op=="modify") {
                var cell=wb.Worksheets["Data"].Cells["J12"];cell.PutValue(9999.0);var s=cell.GetStyle();s.Font.IsBold=true;cell.SetStyle(s);wb.Save(output,SaveFormat.Xlsx);
            } else if(op=="header") {
                var p=wb.Worksheets["Data"].PageSetup;p.SetHeader(0,p.GetHeader(0).Replace("benchmark","edited"));wb.Save(output,SaveFormat.Xlsx);
            } else throw new Exception("unknown operation");
        }
        var total=Stopwatch.GetElapsedTime(start).TotalMilliseconds;
        Console.WriteLine(JsonSerializer.Serialize(new{ok=true,processing_ms=op=="calc"?calcMs:total,load_ms=loadMs,request_api_total_ms=total,digest=op=="read"?Digest(data):null,values,workbook_count=count}));
    } catch(Exception ex) {Console.WriteLine(JsonSerializer.Serialize(new{ok=false,error=ex.ToString(),workbook_count=count}));}
}
