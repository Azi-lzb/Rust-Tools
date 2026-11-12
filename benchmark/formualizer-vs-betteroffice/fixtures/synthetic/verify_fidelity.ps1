param(
    [string]$FixtureDir = $PSScriptRoot,
    [string]$OutputDir = (Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'results\synthetic')
)
$ErrorActionPreference = 'Stop'
$spec = Get-Content -LiteralPath (Join-Path $FixtureDir 'cases.json') -Raw | ConvertFrom-Json
$excel = $null
function Read-BookDigest([string]$Path, $SheetName, $FormulaTargets) {
    $workbook = $null
    try {
        $workbook = $script:excel.Workbooks.Open($Path, 0, $true)
        $sheets = @()
        foreach ($ws in $workbook.Worksheets) {
            $sheets += [pscustomobject]@{
                name = [string]$ws.Name
                orientation = [int]$ws.PageSetup.Orientation
                paper_size = [int]$ws.PageSetup.PaperSize
                print_area = [string]$ws.PageSetup.PrintArea
                print_title_rows = [string]$ws.PageSetup.PrintTitleRows
                print_title_columns = [string]$ws.PageSetup.PrintTitleColumns
                center_horizontally = [bool]$ws.PageSetup.CenterHorizontally
                center_vertically = [bool]$ws.PageSetup.CenterVertically
                left_header = [string]$ws.PageSetup.LeftHeader
                center_footer = [string]$ws.PageSetup.CenterFooter
                left_margin = [double]$ws.PageSetup.LeftMargin
                right_margin = [double]$ws.PageSetup.RightMargin
                merged_d1_e1 = [bool]$ws.Range('D1:E1').MergeCells
                col_a_width = [double]$ws.Columns.Item(1).ColumnWidth
                row_1_height = [double]$ws.Rows.Item(1).RowHeight
                hyperlinks = [int]$ws.Hyperlinks.Count
                header_bold = [bool]$ws.Range('A1').Font.Bold
                fill_color_d1 = [int]$ws.Range('D1').Interior.Color
            }
        }
        $formulas = @{}
        foreach ($target in $FormulaTargets) {
            $ws = $workbook.Worksheets.Item([string]$target.sheet)
            $cell = $ws.Cells.Item([int]$target.row + 1, [int]$target.col + 1)
            $formulas[[string]$target.address] = [pscustomobject]@{
                formula = [string]$cell.Formula
                value2 = [string]$cell.Value2
                text = [string]$cell.Text
            }
        }
        return [pscustomobject]@{
            path = $Path
            status = 'OPENED'
            sheet_names = @($sheets | ForEach-Object name)
            sheets = $sheets
            defined_names = [int]$workbook.Names.Count
            formula_targets = $formulas
            formula_target_count = $formulas.Count
        }
    } catch {
        return [pscustomobject]@{ path = $Path; status = 'OPEN_FAILED'; error = $_.Exception.Message }
    } finally {
        if ($workbook) { $workbook.Close($false) | Out-Null; [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) }
    }
}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $files = [ordered]@{
        source = Join-Path $FixtureDir 'formula-suite.xlsx'
        betteroffice = Join-Path $OutputDir 'betteroffice.xlsx'
        formualizer = Join-Path $OutputDir 'formualizer.xlsx'
        formualizer_cache = Join-Path $OutputDir 'formualizer-cache.xlsx'
    }
    $result = [ordered]@{}
    foreach ($entry in $files.GetEnumerator()) {
        if (Test-Path -LiteralPath $entry.Value) {
            $result[$entry.Key] = Read-BookDigest $entry.Value $null $spec.targets
        } else {
            $result[$entry.Key] = [pscustomobject]@{ path = $entry.Value; status = 'MISSING' }
        }
    }
    $destination = Join-Path $OutputDir 'fidelity-comparison.json'
    $result | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $destination -Encoding utf8
    Write-Output "FIDELITY=$destination"
    $result.GetEnumerator() | ForEach-Object { Write-Output "$($_.Key): $($_.Value.status)" }
} finally {
    if ($excel) { $excel.Quit(); [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
