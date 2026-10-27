param(
    [string]$Workspace = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path,
    [string]$ReportPath = (Join-Path $PSScriptRoot 'reports/dedup_comment_delete_excel_com_20261023.json')
)

$ErrorActionPreference = 'Stop'

function Convert-ToWslPath([string]$Path) {
    $absolute = [IO.Path]::GetFullPath($Path)
    if ($absolute -notmatch '^([A-Za-z]):\\(.*)$') {
        throw "Path is not on a mounted Windows drive: $absolute"
    }
    return "/mnt/$($Matches[1].ToLowerInvariant())/$($Matches[2].Replace('\', '/'))"
}

function Invoke-Wsl([string]$Command) {
    $output = & wsl.exe -d Ubuntu -- bash -lc $Command 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "WSL command failed ($LASTEXITCODE): $($output -join "`n")"
    }
    return $output -join "`n"
}

function Write-Utf8NoBom([string]$Path, [string]$Content) {
    $encoding = [System.Text.UTF8Encoding]::new($false)
    [IO.File]::WriteAllText($Path, $Content, $encoding)
}

function Get-CommentText($Cell) {
    try {
        if ($null -ne $Cell.Comment) { return $Cell.Comment.Text() }
    } catch { }
    return $null
}

function Get-Snapshot($Workbook) {
    $data = $Workbook.Worksheets.Item('Commented')
    $refs = $Workbook.Worksheets.Item('Refs')
    $cells = [ordered]@{}
    $comments = [ordered]@{}
    for ($row = 1; $row -le 6; $row++) {
        for ($column = 1; $column -le 12; $column++) {
            $cell = $data.Cells.Item($row, $column)
            $address = $cell.Address($false, $false)
            $formula = $null
            if ($cell.HasFormula) { $formula = [string]$cell.Formula }
            $cells[$address] = [ordered]@{
                value = $cell.Value2
                formula = $formula
                number_format = [string]$cell.NumberFormat
                fill_color = [double]$cell.Interior.Color
                bold = [bool]$cell.Font.Bold
            }
            $comment = Get-CommentText $cell
            if ($null -ne $comment) { $comments[$address] = $comment }
        }
    }
    $sheets = @()
    for ($index = 1; $index -le $Workbook.Worksheets.Count; $index++) {
        $sheets += $Workbook.Worksheets.Item($index).Name
    }
    $rowHeights = [ordered]@{}
    for ($row = 1; $row -le 6; $row++) {
        $rowHeights[[string]$row] = [double]$data.Rows.Item($row).RowHeight
    }
    $columnWidths = [ordered]@{}
    for ($column = 1; $column -le 12; $column++) {
        $letter = [char](64 + $column)
        $columnWidths[[string]$letter] = [double]$data.Columns.Item($column).ColumnWidth
    }
    return [ordered]@{
        sheets = $sheets
        cells = $cells
        formulas = [ordered]@{
            'Refs!A1' = $refs.Range('A1').Formula
            'Refs!A2' = $refs.Range('A2').Formula
        }
        comments = $comments
        merged_range = $data.Range('E2').MergeArea.Address($false, $false)
        defined_name = $Workbook.Names.Item('DataValues').RefersTo
        row_heights = $rowHeights
        column_widths = $columnWidths
        autofilter_range = $data.AutoFilter.Range.Address($false, $false)
    }
}

$workspaceWsl = Convert-ToWslPath $Workspace
$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("pt-2-2-com-" + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $tempRoot
$source = Join-Path $tempRoot 'source.xlsx'
$expected = Join-Path $tempRoot 'excel_expected.xlsx'
$actual = Join-Path $tempRoot 'rust_candidate.xlsx'
$planPath = Join-Path $tempRoot 'plan.json'
$capabilityPath = Join-Path $tempRoot 'test-capability.json'
$executionPath = Join-Path $tempRoot 'execution.json'
$excel = $null
$sourceBook = $null
$oracleBook = $null
$expectedBook = $null
$actualBook = $null
$report = [ordered]@{
    schema_version = 1
    suite = 'pytools.dedup.comment_delete_excel_com_parity'
    date = '2026-10-23'
    feature_id = 'PT-2-2'
    status = 'failed'
    sample_kind = 'synthetic_workbook_generated_by_excel_com'
    oracle = $null
    rust_engine = 'umya_spreadsheet'
    production_capability_created = $false
    default_route_changed = $false
    source_unchanged = $false
    semantic_equal = $false
    differences = @()
    not_verified = @('real pytools workbook/configuration', 'Kylin V10', 'UOS V20', 'glibc 2.28', 'offline deb packaging')
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $report.oracle = "Microsoft Excel COM $($excel.Version)"

    $sourceBook = $excel.Workbooks.Add()
    while ($sourceBook.Worksheets.Count -gt 1) {
        $sourceBook.Worksheets.Item($sourceBook.Worksheets.Count).Delete()
    }
    $data = $sourceBook.Worksheets.Item(1)
    $data.Name = 'Commented'
    $data.Range('A1').Value2 = '源数据'
    $data.Range('B1').Value2 = 'Key'
    $data.Range('C1').Value2 = 'Value'
    $data.Range('A1').AddComment('本表为源数据') | Out-Null
    $data.Range('B1').AddComment('标识列') | Out-Null

    $data.Range('A2').Value2 = 'first'
    $data.Range('B2').Value2 = 'dup'
    $data.Range('C2').Value2 = 10
    $data.Range('A3').Value2 = 'duplicate'
    $data.Range('B3').Value2 = 'dup'
    $data.Range('C3').Value2 = 20
    $data.Range('A4').Value2 = 'keep'
    $data.Range('B4').Value2 = 'keep'
    $data.Range('C4').Value2 = 30
    $data.Range('A5').Value2 = 'tail'
    $data.Range('B5').Value2 = 'tail'
    $data.Range('C5').Value2 = 40
    $data.Range('E2:F5').Merge()
    $data.Range('E2').Value2 = 'merged-marker'
    $data.Range('G3').AddComment('deleted-row-comment') | Out-Null
    $data.Range('G4').AddComment('survivor-comment') | Out-Null
    $data.Range('G5').AddComment('tail-comment') | Out-Null
    $data.Range('C4').Interior.Color = 65280
    $data.Rows.Item(4).RowHeight = 24
    $data.Range('A1:C5').AutoFilter() | Out-Null
    $data.Range('H1').Formula = '=SUM(C2:C5)'
    $sourceBook.Names.Add('DataValues', "='Commented'!`$C`$2:`$C`$5") | Out-Null
    $refs = $sourceBook.Worksheets.Add()
    $refs.Name = 'Refs'
    $refs.Range('A1').Formula = "='Commented'!C4"
    $refs.Range('A2').Formula = '=SUM(Commented!C2:C5)'
    $sourceBook.SaveAs($source, 51)
    $sourceBook.Close($false)
    $sourceBook = $null

    $oracleBook = $excel.Workbooks.Open($source)
    $oracleBook.Worksheets.Item('Commented').Rows.Item(3).Delete()
    $oracleBook.SaveAs($expected, 51)
    $oracleBook.Close($false)
    $oracleBook = $null

    $sourceHashBefore = (Get-FileHash -Algorithm SHA256 -LiteralPath $source).Hash
    $sourceWsl = Convert-ToWslPath $source
    $actualWsl = Convert-ToWslPath $actual
    $planWsl = Convert-ToWslPath $planPath
    $binary = "$workspaceWsl/target/debug/pytools"
    $command = "cd '$workspaceWsl' && '$binary' dedup-comment-plan --input '$sourceWsl' --output '$actualWsl' --action delete --plan-output '$planWsl'"
    Invoke-Wsl $command | Out-Null
    $plan = Get-Content -Raw -LiteralPath $planPath | ConvertFrom-Json
    $capability = [ordered]@{
        engine = 'umya_spreadsheet'
        status = 'passed'
        format = 'xlsx'
        operations = @($plan.operations | ForEach-Object { $_.operation })
        supported_features = @($plan.workbook.features_present)
        preserved = @($plan.preserve)
        platform = 'windows11-x64-excel-com-probe'
        evidence = [ordered]@{
            sample_id = 'synthetic-pt-2-2-excel-com-row-delete'
            report_path = 'pytools-core/probe/compare_dedup_row_delete_excel_com.ps1'
            adapter_version = 'umya-spreadsheet-3.1.1'
            engine_version = 'umya-spreadsheet-3.1.1'
            platform = 'windows11-x64-excel-com-probe'
        }
    }
    Write-Utf8NoBom $capabilityPath ($capability | ConvertTo-Json -Depth 16)
    $capWsl = Convert-ToWslPath $capabilityPath
    $executionWsl = Convert-ToWslPath $executionPath
    $execute = "cd '$workspaceWsl' && '$binary' dedup-comment-plan --input '$sourceWsl' --output '$actualWsl' --action delete --capabilities '$capWsl' --platform windows11-x64-excel-com-probe > '$executionWsl'"
    Invoke-Wsl $execute | Out-Null
    $execution = Get-Content -Raw -LiteralPath $executionPath | ConvertFrom-Json

    $expectedBook = $excel.Workbooks.Open($expected, 0, $true)
    $actualBook = $excel.Workbooks.Open($actual, 0, $true)
    $expectedProjection = Get-Snapshot $expectedBook
    $actualProjection = Get-Snapshot $actualBook
    $expectedJson = ConvertTo-Json -InputObject $expectedProjection -Depth 16 -Compress
    $actualJson = ConvertTo-Json -InputObject $actualProjection -Depth 16 -Compress
    $report.execution = $execution.execution
    $report.plan_operations = @($plan.operations)
    $report.expected_projection = $expectedProjection
    $report.rust_projection = $actualProjection
    $report.source_unchanged = $sourceHashBefore -eq (Get-FileHash -Algorithm SHA256 -LiteralPath $source).Hash
    $report.semantic_equal = $expectedJson -ceq $actualJson
    if (-not $report.semantic_equal) {
        $report.differences = @('Excel COM and Rust/umya semantic projections differ')
    }
    if ($execution.execution.engine -ne 'umya_spreadsheet') {
        $report.differences += "Unexpected Rust engine: $($execution.execution.engine)"
    }
    if (-not $report.source_unchanged) {
        $report.differences += 'Source workbook changed during execution'
    }
    $report.status = if ($report.semantic_equal -and $report.source_unchanged -and $execution.execution.engine -eq 'umya_spreadsheet') { 'passed_synthetic_com_parity' } else { 'failed_synthetic_com_parity' }
}
catch {
    $report.error = $_.Exception.Message
    $report.error_line = $_.InvocationInfo.ScriptLineNumber
    $report.error_command = $_.InvocationInfo.Line.Trim()
    $report.status = 'failed'
}
finally {
    foreach ($workbook in @($actualBook, $expectedBook, $oracleBook, $sourceBook)) {
        if ($null -ne $workbook) {
            try { $workbook.Close($false) } catch { }
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } catch { }
        }
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { }
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Write-Utf8NoBom $ReportPath ($report | ConvertTo-Json -Depth 20)
    $resolvedTempRoot = [IO.Path]::GetFullPath($tempRoot)
    $separator = [IO.Path]::DirectorySeparatorChar
    $resolvedTempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd($separator) + $separator
    if (-not $resolvedTempRoot.StartsWith($resolvedTempBase, [StringComparison]::OrdinalIgnoreCase) -or -not (Split-Path -Leaf $resolvedTempRoot).StartsWith('pt-2-2-com-')) {
        throw "Refusing cleanup outside the test temporary directory: $resolvedTempRoot"
    }
    Remove-Item -LiteralPath $resolvedTempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Get-Content -Raw -LiteralPath $ReportPath
if ($report.status -notlike 'passed*') { exit 1 }
