using Aspose.Cells;
using System.Text.Json;

static class RootCause {
    public static int Run(string root) {
        using var specDoc=JsonDocument.Parse(File.ReadAllText(Path.Combine(root,"spec.json")));
        var spec=specDoc.RootElement;
        var count=spec.GetProperty("formulas").GetArrayLength();
        var generated=Path.Combine(root,"aspose-references.xls");
        if(File.Exists(generated)) throw new Exception("refuse existing fixture");
        using(var wb=new Workbook()) {
            wb.Worksheets.Clear();
            foreach(var name in spec.GetProperty("sheet_names").EnumerateArray()) wb.Worksheets.Add(name.GetString());
            foreach(var name in spec.GetProperty("sheet_names").EnumerateArray()) {
                var sheet=wb.Worksheets[name.GetString()];
                foreach(var v in spec.GetProperty("values").EnumerateArray()) {
                    if(name.GetString()=="Calc" && v[0].GetString()![0]=='A' && char.IsDigit(v[0].GetString()![1])) continue;
                    var cell=sheet.Cells[v[0].GetString()];
                    if(v[1].ValueKind==JsonValueKind.String) cell.PutValue(v[1].GetString());
                    else cell.PutValue(v[1].GetDouble());
                }
                if(name.GetString()=="Calc") {int r=0;foreach(var f in spec.GetProperty("formulas").EnumerateArray()) sheet.Cells[r++,0].Formula="="+f.GetString();}
            }
            // No CalculateFormula here: this fixture tests parsing, not calculation.
            wb.Save(generated,SaveFormat.Excel97To2003);
        }
        foreach(var file in new[]{"xlwt-references.xls","aspose-references.xls","openpyxl-control.xlsx"}) {
            using var wb=new Workbook(Path.Combine(root,file));
            var formulas=Enumerable.Range(0,count).Select(r=>new {row=r+1,column=1,formula=wb.Worksheets["Calc"].Cells[r,0].Formula}).ToArray();
            Console.WriteLine(JsonSerializer.Serialize(new {ok=true,input=file,formulas}));
        }
        using(var wb=new Workbook(Path.Combine(root,"validation-minimal.xlsx"))) {
            var validations=wb.Worksheets["Data"].Validations.Cast<Validation>().Select(v=>new {formula1=v.Formula1,formula2=v.Formula2,op=v.Operator.ToString()}).ToArray();
            Console.WriteLine(JsonSerializer.Serialize(new {ok=true,input="validation-minimal.xlsx",validations}));
        }
        return 0;
    }
}
