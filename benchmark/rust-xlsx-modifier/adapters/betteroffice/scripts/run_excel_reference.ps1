param([string]$Results = (Join-Path $PSScriptRoot '..\results'))
$ErrorActionPreference = 'Stop'
$Results = [System.IO.Path]::GetFullPath($Results)
$Source = Join-Path $Results 'audit_fixture.xlsx'
$Output = Join-Path $Results 'excel-com-structure-reference.xlsx'
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false
try {
    $Wb = $Excel.Workbooks.Open($Source, 0, $false)
    $Ws = $Wb.Worksheets.Item('Data')
    $Ws.Rows('2:2').Delete()
    $Ws.Rows('2:3').Insert()
    $Wb.SaveAs($Output, 51)
    $Report = [ordered]@{
        excel_version = $Excel.Version
        conditional_format_count = $Ws.Range('A1:B20').FormatConditions.Count
        conditional_format_applies_to = $Ws.Range('A1:B20').FormatConditions.Item(1).AppliesTo.Address()
        merged_a6_c6 = $Ws.Range('A6:C6').MergeCells
        chart_count = $Ws.ChartObjects().Count
        print_area = $Ws.PageSetup.PrintArea
        data_range_name = $Wb.Names.Item('DataRange').RefersTo
    }
    $Wb.Close($false)
    $Report | ConvertTo-Json -Depth 4 | Set-Content -Encoding utf8 (Join-Path $Results 'com-reference.json')
    $Report | ConvertTo-Json -Depth 4
} finally {
    if ($Wb) { try { $Wb.Close($false) } catch {} }
    $Excel.Quit()
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($Excel) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
