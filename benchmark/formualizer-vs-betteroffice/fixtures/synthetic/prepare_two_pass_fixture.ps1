param([string]$FixtureDir = $PSScriptRoot)
$ErrorActionPreference = 'Stop'
$source = Join-Path $FixtureDir 'formula-suite-valid.xlsx'
$output = Join-Path $FixtureDir 'formula-two-pass.xlsx'
$specPath = Join-Path $FixtureDir 'cases-valid.json'
$targetsPath = Join-Path $FixtureDir 'cases-two-pass.json'
$oracle10Path = Join-Path $FixtureDir 'excel-com-two-pass-input10.json'
$oracle25Path = Join-Path $FixtureDir 'excel-com-two-pass-input25.json'
$oracle35Path = Join-Path $FixtureDir 'excel-com-two-pass-input35.json'
$targets = @((Get-Content -LiteralPath $specPath -Raw | ConvertFrom-Json).targets)
$additional = @(
    [pscustomobject]@{ sheet='Calc'; row=39; col=1; address='Calc!B40'; formula='=INDIRECT("Data!B2")+1' },
    [pscustomobject]@{ sheet='Calc'; row=40; col=1; address='Calc!B41'; formula='=B40+1' },
    [pscustomobject]@{ sheet='Calc'; row=41; col=1; address='Calc!B42'; formula='=INDIRECT("Data!B2")+B41' },
    [pscustomobject]@{ sheet='Calc'; row=42; col=1; address='Calc!B43'; formula='=SUM(B40:B42)' }
)
$targets += $additional
$spec = [pscustomobject]@{
    targets = $targets
    mutations = @([pscustomobject]@{ sheet='Data'; row=1; col=1; value=25 })
    notes = @('Excel COM generated', 'INDIRECT formulas with downstream dependencies', 'zero-based coordinates')
}
$excel = $null
$book = $null
function Get-Oracle([object]$Workbook, [object[]]$FormulaTargets) {
    $values = @{}
    foreach ($target in $FormulaTargets) {
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
    $excel.EnableEvents = $false
    $book = $excel.Workbooks.Open($source, 0, $false)
    foreach ($target in $additional) {
        $book.Worksheets.Item('Calc').Cells.Item([int]$target.row + 1, [int]$target.col + 1).Formula = $target.formula
    }
    $excel.CalculateFullRebuild()
    $oracle10 = Get-Oracle $book $targets
    $oracle10 | ConvertTo-Json -Depth 5 -Compress | Set-Content -LiteralPath $oracle10Path -Encoding utf8
    $book.Worksheets.Item('Data').Range('B2').Value2 = 25
    $excel.CalculateFullRebuild()
    $oracle25 = Get-Oracle $book $targets
    $oracle25 | ConvertTo-Json -Depth 5 -Compress | Set-Content -LiteralPath $oracle25Path -Encoding utf8
    $book.Worksheets.Item('Data').Range('B2').Value2 = 35
    $excel.CalculateFullRebuild()
    $oracle35 = Get-Oracle $book $targets
    $oracle35 | ConvertTo-Json -Depth 5 -Compress | Set-Content -LiteralPath $oracle35Path -Encoding utf8
    $targets | Out-Null
    $spec | ConvertTo-Json -Depth 6 -Compress | Set-Content -LiteralPath $targetsPath -Encoding utf8
    $book.Worksheets.Item('Data').Range('B2').Value2 = 10
    $excel.CalculateFullRebuild()
    $book.SaveAs($output, 51)
    Write-Output "TWO_PASS_FIXTURE=$output"
    Write-Output "FORMULA_TARGETS=$($targets.Count); INDIRECT_TARGETS=4"
    Write-Output "ORACLES=$oracle10Path;$oracle25Path;$oracle35Path"
} finally {
    if ($book) { $book.Close($false) | Out-Null }
    if ($excel) { $excel.Quit(); [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
