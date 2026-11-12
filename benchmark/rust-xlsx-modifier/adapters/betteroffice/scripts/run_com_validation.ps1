param([string]$Results = (Join-Path $PSScriptRoot '..\results'))
$ErrorActionPreference = 'Stop'
$Results = [System.IO.Path]::GetFullPath($Results)
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false
try {
    $StructurePath = Join-Path $Results 'structure.xlsx'
    $Wb = $Excel.Workbooks.Open($StructurePath, 0, $true)
    $Ws = $Wb.Worksheets.Item('Data')
    $Structure = [ordered]@{
        excel_version = $Excel.Version
        charts = $Ws.ChartObjects().Count
        conditional_format_count = $Ws.Range('A1:B3').FormatConditions.Count
        conditional_format_applies_to = $Ws.Range('A1:B3').FormatConditions.Item(1).AppliesTo.Address()
        merged_a6_c6 = $Ws.Range('A6:C6').MergeCells
        data_range_name = $Wb.Names.Item('DataRange').RefersTo
        print_area_after_structural_edit = $Ws.PageSetup.PrintArea
        orientation = $Ws.PageSetup.Orientation
        header_left = $Ws.PageSetup.LeftHeader
        footer_center = $Ws.PageSetup.CenterFooter
        left_margin_points = $Ws.PageSetup.LeftMargin
        e_column_width = $Ws.Columns.Item(5).ColumnWidth
        row_2_height = $Ws.Rows.Item(2).RowHeight
        a1_bold = $Ws.Range('A1').Font.Bold
    }
    $Wb.Close($false)
    $ProbePath = Join-Path $Results 'print_area_probe.xlsx'
    $Wb = $Excel.Workbooks.Open($ProbePath, 0, $true)
    $Ws = $Wb.Worksheets.Item('Data')
    $Probe = [ordered]@{ print_area_after_defined_name_attempt = $Ws.PageSetup.PrintArea; names = @($Wb.Names | ForEach-Object { $_.Name + '=' + $_.RefersTo }) }
    $Wb.Close($false)
    $FormulaPath = Join-Path $Results 'formula.xlsx'
    $Wb = $Excel.Workbooks.Open($FormulaPath, 0, $true)
    $Ws = $Wb.Worksheets.Item('Data')
    $Formula = [ordered]@{ b1_formula = $Ws.Range('B1').Formula; b1_value2 = $Ws.Range('B1').Value2; c1_formula = $Ws.Range('C1').Formula; c1_value2 = $Ws.Range('C1').Value2 }
    $Wb.Close($false)
    $Reference = Get-Content -Raw (Join-Path $Results 'com-reference.json') | ConvertFrom-Json
    $Report = [ordered]@{
        structure = $Structure
        excel_com_reference = $Reference
        conditional_format_range_matches_excel = ($Structure.conditional_format_applies_to -eq $Reference.conditional_format_applies_to)
        print_area_probe = $Probe
        formulas = $Formula
    }
    $Report | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 (Join-Path $Results 'com-validation.json')
    $Report | ConvertTo-Json -Depth 5
} finally {
    if ($Wb) { try { $Wb.Close($false) } catch {} }
    $Excel.Quit()
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($Excel) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
