using Aspose.Cells;
using System.Drawing;
using System.Text.Json;

internal static class StyleSaveDiagnostic
{
    public static int Run(string input, string directory)
    {
        if (Directory.Exists(directory)) throw new IOException("Refuse existing diagnostic directory");
        Directory.CreateDirectory(directory);
        foreach (string coordinate in new[] { "D12", "AA100", "D100", "AA12", "A100" })
        foreach (bool stream in new[] { false, true })
        foreach (bool range in new[] { false, true })
        {
            using var wb = new Workbook(input);
            var cell = wb.Worksheets["Data"].Cells[coordinate];
            cell.PutValue(123);
            var style = cell.GetStyle();
            style.Font.IsBold = true; style.Font.Size = 14;
            style.Pattern = BackgroundType.Solid; style.ForegroundColor = Color.FromArgb(255, 230, 153);
            if (range) wb.Worksheets["Data"].Cells.CreateRange(cell.Row, cell.Column, 1, 1).ApplyStyle(style, new StyleFlag { All = true });
            else cell.SetStyle(style);
            var output = Path.Combine(directory, coordinate + (range ? "-range" : "-cell") + (stream ? "-stream.xlsx" : "-file.xlsx"));
            if (stream) { using var file = new FileStream(output, FileMode.CreateNew); wb.Save(file, SaveFormat.Xlsx); }
            else wb.Save(output, SaveFormat.Xlsx);
            Console.WriteLine(JsonSerializer.Serialize(new { coordinate, stream, range, output }));
        }
        return 0;
    }
}
