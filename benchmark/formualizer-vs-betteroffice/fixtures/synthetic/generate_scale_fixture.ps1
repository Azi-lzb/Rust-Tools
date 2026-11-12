param(
    [ValidateRange(100, 100000)][int]$FormulaRows = 10000,
    [string]$Tag = "$FormulaRows",
    [switch]$IncludeIndirect
)
$ErrorActionPreference = 'Stop'
$fixtureDir = $PSScriptRoot
$xlsx = Join-Path $fixtureDir "formula-scale-$Tag.xlsx"
$specPath = Join-Path $fixtureDir "cases-scale-$Tag.json"
$baseOraclePath = Join-Path $fixtureDir "excel-com-scale-$Tag-base.json"
$input25OraclePath = Join-Path $fixtureDir "excel-com-scale-$Tag-input25.json"
$input35OraclePath = Join-Path $fixtureDir "excel-com-scale-$Tag-input35.json"
$excel = $null
$book = $null

function Get-Oracle([object]$CalcSheet, [int]$Rows, [int]$LastRow) {
    $matrix = $CalcSheet.Range("B2:B$LastRow").Value2
    $values = [ordered]@{}
    for ($i = 1; $i -le $Rows; $i++) {
        $excelRow = $i + 1
        $value = $matrix[$i, 1]
        $values["Calc!B$excelRow"] = "NUM:$([double]$value)"
    }
    return $values
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $book = $excel.Workbooks.Add()
    $data = $book.Worksheets.Item(1)
    $data.Name = 'Data'
    $calc = $book.Worksheets.Add()
    $calc.Name = 'Calc'
    $data.Range('A1').Value2 = 'Region'
    $data.Range('B1').Value2 = 'Amount'
    $data.Range('C1').Value2 = 'Type'
    for ($i = 1; $i -le 100; $i++) {
        $row = $i + 1
        $region = @('East', 'West', 'South', 'North')[($i - 1) % 4]
        $kind = if (($i % 2) -eq 0) { 'Retail' } else { 'Wholesale' }
        $data.Cells.Item($row, 1).Value2 = $region
        $data.Cells.Item($row, 2).Value2 = [double]$i
        $data.Cells.Item($row, 3).Value2 = $kind
    }
    $groups = @(
        '=SUM(Data!R2C2:R101C2)+ROW()',
        '=SUMIF(Data!R2C1:R101C1,"East",Data!R2C2:R101C2)+ROW()',
        '=SUMIFS(Data!R2C2:R101C2,Data!R2C1:R101C1,"East",Data!R2C3:R101C3,"Retail")+ROW()',
        '=COUNTIF(Data!R2C1:R101C1,"West")+ROW()',
        '=VLOOKUP("South",Data!R2C1:R101C3,2,FALSE)+ROW()',
        '=INDEX(Data!R2C2:R101C2,MATCH("North",Data!R2C1:R101C1,0))+ROW()',
        '=ROUND(SUM(Data!R2C2:R101C2)/ROW(),3)'
    )
    if ($IncludeIndirect) { $groups += '=INDIRECT("Data!B2")+ROW()' }
    $targets = [Collections.Generic.List[object]]::new()
    $cursor = 0
    $groupIndex = 0
    foreach ($formula in $groups) {
        if ($cursor -ge $FormulaRows) { break }
        $count = [math]::Ceiling(($FormulaRows - $cursor) / ($groups.Count - $groupIndex))
        $count = [math]::Min($count, $FormulaRows - $cursor)
        $firstRow = 2 + $cursor
        $lastRow = $firstRow + $count - 1
        $calc.Range("B${firstRow}:B${lastRow}").FormulaR1C1 = $formula
        for ($r = $firstRow; $r -le $lastRow; $r++) {
            $targets.Add([pscustomobject]@{ sheet='Calc'; row=[int]($r - 1); col=[int]1; address="Calc!B$r" })
        }
        $cursor += $count
        $groupIndex++
    }
    $calc.Range('A1').Value2 = 'Formula group'
    $calc.Range('B1').Value2 = 'Result'
    $calc.Range('A1:B1').Font.Bold = $true
    $calc.PageSetup.PrintArea = "`$A`$1:`$B`$$($FormulaRows + 1)"
    $excel.CalculateFullRebuild()
    $base = Get-Oracle $calc $FormulaRows ($FormulaRows + 1)
    $book.SaveAs($xlsx, 51)
    $base | ConvertTo-Json -Depth 4 -Compress | Set-Content -LiteralPath $baseOraclePath -Encoding utf8
    $data.Range('B2').Value2 = 25
    $excel.CalculateFullRebuild()
    $input25 = Get-Oracle $calc $FormulaRows ($FormulaRows + 1)
    $input25 | ConvertTo-Json -Depth 4 -Compress | Set-Content -LiteralPath $input25OraclePath -Encoding utf8
    $data.Range('B2').Value2 = 35
    $excel.CalculateFullRebuild()
    $input35 = Get-Oracle $calc $FormulaRows ($FormulaRows + 1)
    $input35 | ConvertTo-Json -Depth 4 -Compress | Set-Content -LiteralPath $input35OraclePath -Encoding utf8
    $spec = [pscustomobject]@{
        targets = $targets
        mutations = @([pscustomobject]@{ sheet='Data'; row=1; col=1; value=25 })
        notes = @('Seven repeated formula groups', 'Scale fixture generated with Excel COM', 'Coordinates are zero-based')
    }
    $spec | ConvertTo-Json -Depth 6 -Compress | Set-Content -LiteralPath $specPath -Encoding utf8
    $book.Worksheets.Item('Data').Range('B2').Value2 = 1
    $excel.CalculateFullRebuild()
    $book.Save()
    Write-Output "XLSX=$xlsx"
    Write-Output "FORMULAS=$FormulaRows"
    Write-Output "SPEC=$specPath"
    Write-Output "ORACLES=$baseOraclePath;$input25OraclePath;$input35OraclePath"
} finally {
    if ($book) { $book.Close($false) | Out-Null }
    if ($excel) { $excel.Quit(); [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
