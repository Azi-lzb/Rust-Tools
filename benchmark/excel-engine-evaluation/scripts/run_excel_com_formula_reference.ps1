$ErrorActionPreference = 'Stop'
$experimentRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$outputRoot = Join-Path $experimentRoot 'reports\samples\formula-reference'
$evidenceRoot = Join-Path $experimentRoot 'reports\evidence\formula-reference'
New-Item -ItemType Directory -Force -Path $outputRoot,$evidenceRoot | Out-Null
$excel = $null
$rows = [System.Collections.Generic.List[object]]::new()
$book = $null
$data = $null
$calc = $null
$timer = [System.Diagnostics.Stopwatch]::StartNew()
try {
    $excel = New-Object -ComObject Excel.Application
    $startupMs = $timer.Elapsed.TotalMilliseconds
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AskToUpdateLinks = $false
    $excel.ScreenUpdating = $false; $excel.EnableEvents = $false
    $version = "Excel $($excel.Version) build $($excel.Build)"
    $book = $excel.Workbooks.Add()
    $data = $book.Worksheets.Item(1); $data.Name = 'Data'
    $calc = $book.Worksheets.Add(); $calc.Name = 'Calc'
    $data.Range('A1').Value2 = 'Key'; $data.Range('B1').Value2 = 'Value'
    $data.Range('A2').Value2 = 'a'; $data.Range('B2').Value2 = 2
    $data.Range('A3').Value2 = 'b'; $data.Range('B3').Value2 = 3
    $data.Range('A4').Value2 = 'b'; $data.Range('B4').Value2 = 4
    [void]$book.Names.Add('InputValues','=Data!$B$2:$B$4')
    $formulas = @(
        '=SUM(InputValues)', '=SUMIF(Data!A2:A4,"b",Data!B2:B4)', '=COUNTIF(Data!B2:B4,">2")',
        '=IF(AND(Data!B2>0,Data!B3>0),ROUND(Data!B2*Data!B3,2),0)',
        '=IF(OR(Data!B2=0,Data!B3=3),1,0)', '=VLOOKUP("b",Data!A2:B4,2,FALSE)',
        '=INDEX(Data!B2:B4,MATCH("b",Data!A2:A4,0))', '=DATE(2026,1,2)',
        '=SUM(Data!B2:B4)+Data!B2', '=1/0'
    )
    for ($i=0; $i -lt $formulas.Count; $i++) { $calc.Cells.Item($i+1,1).Formula = $formulas[$i] }
    $sw = [System.Diagnostics.Stopwatch]::StartNew(); $excel.CalculateFullRebuild(); $firstMs = $sw.Elapsed.TotalMilliseconds
    $initialPath = Join-Path $outputRoot 'excel-com-formulas-initial.xlsx'
    $book.SaveAs($initialPath,51)
    for ($i=1; $i -le $formulas.Count; $i++) {
        $cell = $calc.Cells.Item($i,1)
        $rows.Add([pscustomobject]@{phase='initial';engine='excel-com';address="A$i";formula=[string]$cell.Formula;display=[string]$cell.Text;value2=[string]$cell.Value2;elapsed_ms=[math]::Round($firstMs,3)})
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell)
    }
    $data.Range('B2').Value2 = 5
    $calc.Range('A1').Formula = '=SUM(InputValues)*2'
    $sw.Restart(); $excel.CalculateFullRebuild(); $editMs = $sw.Elapsed.TotalMilliseconds
    $editedPath = Join-Path $outputRoot 'excel-com-formulas-edited.xlsx'
    $book.SaveAs($editedPath,51)
    $cell = $calc.Range('A1')
    $rows.Add([pscustomobject]@{phase='edited';engine='excel-com';address='A1';formula=[string]$cell.Formula;display=[string]$cell.Text;value2=[string]$cell.Value2;elapsed_ms=[math]::Round($editMs,3)})
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell)
} finally {
    if ($calc) { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($calc) }
    if ($data) { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($data) }
    if ($book) { try { $book.Close($false) } catch {}; [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($excel) { try { $excel.Quit() } catch {}; [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [gc]::Collect(); [gc]::WaitForPendingFinalizers()
}
$rows | Export-Csv -LiteralPath (Join-Path $evidenceRoot 'excel-com-formula-results.csv') -NoTypeInformation -Encoding utf8
@($version,"startup_ms=$([math]::Round($startupMs,3))","full_rebuild_initial_ms=$([math]::Round($firstMs,3))","full_rebuild_after_edit_ms=$([math]::Round($editMs,3))","formula_fixture=10 formulas; input B2 changed 2→5; A1 formula changed") |
    Set-Content -LiteralPath (Join-Path $evidenceRoot 'excel-com-formula-environment.txt') -Encoding utf8
Write-Output "Excel COM formula reference rows=$($rows.Count) version=$version startup_ms=$([math]::Round($startupMs,3)) calc_ms=$([math]::Round($firstMs,3)) edited_calc_ms=$([math]::Round($editMs,3))"
