param(
    [string]$Workspace = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path,
    [string]$ReportPath = (Join-Path $PSScriptRoot 'reports/PT-7-3_excel_com_parity_20261023_r2.json')
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
    [IO.File]::WriteAllText($Path, $Content, [System.Text.UTF8Encoding]::new($false))
}

function Get-CommentText($Cell) {
    try {
        if ($null -ne $Cell.Comment) { return $Cell.Comment.Text() }
    } catch { }
    return $null
}

function Get-Snapshot($Workbook) {
    $main = $Workbook.Worksheets.Item('甲分机构本外币')
    $second = $Workbook.Worksheets.Item('乙分机构人民币')
    $other = $Workbook.Worksheets.Item('非目标工作表')
    $cells = [ordered]@{}
    $comments = [ordered]@{}
    for ($row = 1; $row -le 8; $row++) {
        for ($column = 1; $column -le 8; $column++) {
            $cell = $main.Cells.Item($row, $column)
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
    for ($row = 1; $row -le 8; $row++) {
        $rowHeights[[string]$row] = [double]$main.Rows.Item($row).RowHeight
    }
    return [ordered]@{
        sheets = $sheets
        main_cells = $cells
        comments = $comments
        formula = [string]$main.Range('H1').Formula
        merged_range = $main.Range('E7').MergeArea.Address($false, $false)
        row_heights = $rowHeights
        second_sheet_values = @(
            [string]$second.Range('A1').Value2,
            [string]$second.Range('A2').Value2,
            [string]$second.Range('A3').Value2
        )
        non_target_values = @(
            [string]$other.Range('A1').Value2,
            [string]$other.Range('A2').Value2,
            [string]$other.Range('A3').Value2
        )
        cross_sheet_formula = [string]$second.Range('C1').Formula
        branch_values_name = [string]$Workbook.Names.Item('BranchValues').RefersTo
    }
}

$workspaceWsl = Convert-ToWslPath $Workspace
$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("pt-7-3-com-" + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $tempRoot
$source = Join-Path $tempRoot 'source.xlsx'
$config = Join-Path $tempRoot 'config.xlsx'
$expected = Join-Path $tempRoot 'excel_expected.xlsx'
$actual = Join-Path $tempRoot 'source(金融局).xlsx'
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
    suite = 'pytools.pt_7_3.row_delete_excel_com_parity'
    date = '2026-10-23'
    feature_id = 'PT-7-3'
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

    $configBook = $excel.Workbooks.Add()
    while ($configBook.Worksheets.Count -gt 1) {
        $configBook.Worksheets.Item($configBook.Worksheets.Count).Delete()
    }
    $configSheet = $configBook.Worksheets.Item(1)
    $configSheet.Name = '机构映射表'
    $configSheet.Range('A1').Value2 = '序号'
    $configSheet.Range('B1').Value2 = '机构名称'
    $configSheet.Range('C1').Value2 = '是否为外资行'
    $configSheet.Range('A2').Value2 = 1
    $configSheet.Range('B2').Value2 = '外资甲'
    $configSheet.Range('C2').Value2 = 1
    $configSheet.Range('A3').Value2 = 2
    $configSheet.Range('B3').Value2 = '外资乙'
    $configSheet.Range('C3').Value2 = 1
    $configSheet.Range('A4').Value2 = 3
    $configSheet.Range('B4').Value2 = '普通行'
    $configSheet.Range('C4').Value2 = 0
    $configBook.SaveAs($config, 51)
    $configBook.Close($false)

    $sourceBook = $excel.Workbooks.Add()
    while ($sourceBook.Worksheets.Count -gt 1) {
        $sourceBook.Worksheets.Item($sourceBook.Worksheets.Count).Delete()
    }
    $main = $sourceBook.Worksheets.Item(1)
    $main.Name = '甲分机构本外币'
    $main.Range('A1').Value2 = '机构'
    $main.Range('B1').Value2 = '余额'
    $rows = @(
        [pscustomobject]@{ institution = '普通甲'; amount = 10 },
        [pscustomobject]@{ institution = '外资甲'; amount = 20 },
        [pscustomobject]@{ institution = '普通乙'; amount = 30 },
        [pscustomobject]@{ institution = '外资乙'; amount = 40 },
        [pscustomobject]@{ institution = '普通丙'; amount = 50 },
        [pscustomobject]@{ institution = '尾行'; amount = 60 }
    )
    for ($index = 0; $index -lt $rows.Count; $index++) {
        $excelRow = $index + 2
        $main.Range("A$excelRow").Value2 = $rows[$index].institution
        $main.Range("B$excelRow").Value2 = [double]$rows[$index].amount
    }
    $main.Range('B2:B7').NumberFormat = '0.00'
    $main.Range('C3').AddComment('删除行批注') | Out-Null
    $main.Range('C4').AddComment('保留行批注一') | Out-Null
    $main.Range('C7').AddComment('保留行批注二') | Out-Null
    $main.Range('C4').Interior.Color = 65280
    $main.Range('H1').Formula = '=SUM(B2:B7)'
    $main.Rows.Item(7).RowHeight = 24
    $main.Range('E9:F10').Merge()
    $main.Range('E9').Value2 = '合并区域'

    $second = $sourceBook.Worksheets.Add()
    $second.Name = '乙分机构人民币'
    $second.Range('A1').Value2 = '机构'
    $second.Range('A2').Value2 = '外资乙'
    $second.Range('A3').Value2 = '普通丁'
    $second.Range('C1').Formula = "='甲分机构本外币'!B7"
    $other = $sourceBook.Worksheets.Add()
    $other.Name = '非目标工作表'
    $other.Range('A1').Value2 = '机构'
    $other.Range('A2').Value2 = '外资甲'
    $sourceBook.Names.Add('BranchValues', "='甲分机构本外币'!`$B`$2:`$B`$7") | Out-Null
    $sourceBook.SaveAs($source, 51)
    $sourceBook.Close($false)
    $sourceBook = $null

    $sourceHashBefore = (Get-FileHash -Algorithm SHA256 -LiteralPath $source).Hash
    $oracleBook = $excel.Workbooks.Open($source, 0, $false)
    $oracleMain = $oracleBook.Worksheets.Item('甲分机构本外币')
    $oracleMain.Rows.Item(5).Delete()
    $oracleMain.Rows.Item(3).Delete()
    $oracleBook.Worksheets.Item('乙分机构人民币').Rows.Item(2).Delete()
    $oracleBook.SaveAs($expected, 51)
    $oracleBook.Close($false)
    $oracleBook = $null

    $sourceWsl = Convert-ToWslPath $source
    $configWsl = Convert-ToWslPath $config
    $actualWsl = Convert-ToWslPath $actual
    $planWsl = Convert-ToWslPath $planPath
    $binary = "$workspaceWsl/target/debug/pytools"
    $planCommand = "cd '$workspaceWsl' && '$binary' remove-foreign-bank --config '$configWsl' '$sourceWsl' > '$planWsl'"
    Invoke-Wsl $planCommand | Out-Null
    $planReport = Get-Content -Raw -LiteralPath $planPath | ConvertFrom-Json
    $plan = $planReport.items[0].plan
    if ($null -eq $plan -or $plan.operations.Count -eq 0) {
        throw "Rust CLI produced no PT-7-3 delete-row plan: $($planReport | ConvertTo-Json -Depth 12 -Compress)"
    }

    $platform = 'windows11-x64-pt-7-3-excel-com-probe'
    $capability = [ordered]@{
        engine = 'umya_spreadsheet'
        status = 'passed'
        format = [string]$plan.format
        operations = @($plan.operations | ForEach-Object { $_.operation } | Sort-Object -Unique)
        supported_features = @($plan.workbook.features_present)
        preserved = @($plan.preserve)
        platform = $platform
        evidence = [ordered]@{
            sample_id = 'synthetic-pt-7-3-excel-com-row-delete'
            report_path = 'pytools-core/probe/compare_pt73_row_delete_excel_com.ps1'
            adapter_version = 'umya-spreadsheet-3.1.1'
            engine_version = 'umya-spreadsheet-3.1.1'
            platform = $platform
        }
    }
    Write-Utf8NoBom $capabilityPath ($capability | ConvertTo-Json -Depth 16)
    $capWsl = Convert-ToWslPath $capabilityPath
    $executionWsl = Convert-ToWslPath $executionPath
    $executeCommand = "cd '$workspaceWsl' && '$binary' remove-foreign-bank --config '$configWsl' --capabilities '$capWsl' --platform '$platform' '$sourceWsl' > '$executionWsl'"
    Invoke-Wsl $executeCommand | Out-Null
    $execution = Get-Content -Raw -LiteralPath $executionPath | ConvertFrom-Json
    $report.execution = $execution
    if ($execution.ok_files -ne 1 -or $execution.failed_files -ne 0) {
        throw "Rust PT-7-3 execution failed: $($execution | ConvertTo-Json -Depth 16 -Compress)"
    }
    $expectedBook = $excel.Workbooks.Open($expected, 0, $true)
    $actualBook = $excel.Workbooks.Open($actual, 0, $true)
    $expectedProjection = Get-Snapshot $expectedBook
    $actualProjection = Get-Snapshot $actualBook
    $expectedJson = ConvertTo-Json -InputObject $expectedProjection -Depth 18 -Compress
    $actualJson = ConvertTo-Json -InputObject $actualProjection -Depth 18 -Compress
    $report.plan_operations = @($plan.operations)
    $report.expected_projection = $expectedProjection
    $report.rust_projection = $actualProjection
    $report.source_unchanged = $sourceHashBefore -eq (Get-FileHash -Algorithm SHA256 -LiteralPath $source).Hash
    $report.semantic_equal = $expectedJson -ceq $actualJson
    if (-not $report.semantic_equal) {
        $report.differences = @('Excel COM and Rust/umya semantic projections differ')
    }
    if ($execution.ok_files -ne 1 -or $execution.failed_files -ne 0) {
        $report.differences += "Unexpected PT-7-3 execution summary: ok=$($execution.ok_files), failed=$($execution.failed_files)"
    }
    if (-not $report.source_unchanged) {
        $report.differences += 'Source workbook bytes changed'
    }
    if ($report.semantic_equal -and $report.source_unchanged -and $report.differences.Count -eq 0) {
        $report.status = 'passed_synthetic_com_parity'
    }
} catch {
    $report.error = $_.Exception.Message
    $report.error_stack = $_.ScriptStackTrace
    $report.differences += $_.Exception.Message
} finally {
    foreach ($book in @($actualBook, $expectedBook, $oracleBook, $sourceBook)) {
        if ($null -ne $book) {
            try { $book.Close($false) } catch { }
        }
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $reportPathAbsolute = [IO.Path]::GetFullPath($ReportPath)
    $null = New-Item -ItemType Directory -Force -Path (Split-Path -Parent $reportPathAbsolute)
    Write-Utf8NoBom $reportPathAbsolute ($report | ConvertTo-Json -Depth 24)
    $tempAbsolute = [IO.Path]::GetFullPath($tempRoot)
    $tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
    if ($tempAbsolute.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and
        (Split-Path -Leaf $tempAbsolute).StartsWith('pt-7-3-com-', [StringComparison]::Ordinal)) {
        Remove-Item -LiteralPath $tempAbsolute -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Output (Get-Content -Raw -LiteralPath $ReportPath)
if ($report.status -ne 'passed_synthetic_com_parity') { exit 1 }
