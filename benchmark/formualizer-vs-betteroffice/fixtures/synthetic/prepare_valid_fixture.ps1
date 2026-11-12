param([string]$FixtureDir = $PSScriptRoot)
$ErrorActionPreference = 'Stop'
$source = Join-Path $FixtureDir 'formula-suite.xlsx'
$validPath = Join-Path $FixtureDir 'formula-suite-valid.xlsx'
$casePath = Join-Path $FixtureDir 'cases-valid.json'
$oraclePath = Join-Path $FixtureDir 'excel-com-valid-oracle.json'
$incrementalPath = Join-Path $FixtureDir 'excel-com-valid-input25-oracle.json'
$spec = Get-Content -LiteralPath (Join-Path $FixtureDir 'cases.json') -Raw | ConvertFrom-Json
$spec.targets | Where-Object address -eq 'Calc!B33' | ForEach-Object { $_.formula = '=SUM(Data!B2:B6)+1' }
$excel = $null
$book = $null
function Get-Oracle([object]$Workbook, [object[]]$Targets) {
    $values = @{}
    foreach ($target in $Targets) {
        $cell = $Workbook.Worksheets.Item([string]$target.sheet).Cells.Item([int]$target.row + 1, [int]$target.col + 1)
        $value = $cell.Value2
        if ($cell.Text -match '^#(DIV/0!|N/A|NAME\?|REF!|VALUE!|NUM!|NULL!)$') { $values[$target.address] = "ERROR:$($cell.Text)" }
        elseif ($null -eq $value) { $values[$target.address] = 'EMPTY' }
        elseif ($value -is [bool]) { $values[$target.address] = "BOOL:$($value.ToString().ToLowerInvariant())" }
        elseif ($value -is [string]) { $values[$target.address] = "TEXT:$value" }
        else { $values[$target.address] = "NUM:$([double]$value)" }
    }
    return $values
}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Open($source, 0, $false)
    $book.Worksheets.Item('Calc').Range('B33').Formula = '=SUM(Data!B2:B6)+1'
    $excel.CalculateFullRebuild()
    $base = Get-Oracle $book $spec.targets
    $book.SaveAs($validPath, 51)
    $base | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $oraclePath -Encoding utf8
    $book.Worksheets.Item('Data').Range('B2').Value2 = 25
    $excel.CalculateFullRebuild()
    $changed = Get-Oracle $book $spec.targets
    $changed | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $incrementalPath -Encoding utf8
    $book.Worksheets.Item('Data').Range('B2').Value2 = 10
    $excel.CalculateFullRebuild()
    $book.Save()
    $spec | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $casePath -Encoding utf8
    Write-Output "VALID_FIXTURE=$validPath"
    Write-Output "TARGETS=$($spec.targets.Count); B33=$($base['Calc!B33'])"
} finally {
    if ($book) { $book.Close($false) | Out-Null }
    if ($excel) { $excel.Quit(); [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
