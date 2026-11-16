using Aspose.Cells;
using System.Text.Json;
if(args.Length==2 && args[0]=="--root-cause") return RootCause.Run(args[1]);
if(args.Length!=1) throw new Exception("usage: Capabilities requests.json");
var requests=JsonDocument.Parse(File.ReadAllText(args[0])).RootElement;
foreach(var q in requests.EnumerateArray()) {
    try {
        var input=q.GetProperty("input").GetString()!;
        var output=q.GetProperty("output").GetString()!;
        if(File.Exists(output) || Path.GetFullPath(input)==Path.GetFullPath(output)) throw new Exception("refuse overwrite");
        using var wb=new Workbook(input);
        var op=q.GetProperty("op").GetString();
        if(op=="formula") {
            var ws=wb.Worksheets["Calc"];
            var before=Enumerable.Range(0,4).Select(r=> new {cell=$"A{r+1}",formula=ws.Cells[r,0].Formula,value=ws.Cells[r,0].Value}).ToArray();
            // Explicit SDK calculation only. Never inferred from opening/saving a file.
            wb.CalculateFormula();
            var after=Enumerable.Range(0,4).Select(r=> new {cell=$"A{r+1}",formula=ws.Cells[r,0].Formula,value=ws.Cells[r,0].Value}).ToArray();
            wb.Save(output,SaveFormat.Xlsx);
            Console.WriteLine(JsonSerializer.Serialize(new {ok=true,op,input,output,before,after,calculation="explicit Workbook.CalculateFormula"}));
        } else {
            var cell=wb.Worksheets["Data"].Cells["B12"];
            if(op=="value") cell.PutValue(999.0);
            else if(op=="style") {var s=cell.GetStyle();s.Font.IsBold=true;cell.SetStyle(s);}
            else if(op=="header") {var page=wb.Worksheets["Data"].PageSetup;page.SetHeader(0,page.GetHeader(0).Replace("synthetic","edited"));}
            else if(op!="noop") throw new Exception("unknown op");
            wb.Save(output,SaveFormat.Xlsx);
            Console.WriteLine(JsonSerializer.Serialize(new {ok=true,op,input,output}));
        }
    } catch(Exception ex) {Console.WriteLine(JsonSerializer.Serialize(new {ok=false,error=ex.ToString()})); return 1;}
}
return 0;
