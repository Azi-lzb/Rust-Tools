$ErrorActionPreference = 'Stop'
$experimentRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$inputRoot = Join-Path $experimentRoot 'fixtures\performance'
$outputRoot = Join-Path $experimentRoot 'reports\samples\performance\excel-com-reference'
$evidenceRoot = Join-Path $experimentRoot 'reports\evidence\performance\excel-com-reference'
New-Item -ItemType Directory -Force -Path $outputRoot,$evidenceRoot | Out-Null
$files = @(Get-ChildItem $inputRoot -File -Filter 'batch-*.xls' | Sort-Object Name)
if ($files.Count -ne 8) { throw "Expected 8 XLS fixtures, found $($files.Count)" }

$excel = $null
$rows = [System.Collections.Generic.List[object]]::new()
$timer = [System.Diagnostics.Stopwatch]::StartNew()
try {
    $excel = New-Object -ComObject Excel.Application
    $startupMs = $timer.Elapsed.TotalMilliseconds
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    $version = "Excel $($excel.Version) build $($excel.Build)"

    $warmBook = $excel.Workbooks.Open($files[0].FullName, 0, $true)
    $warmBook.Close($false)
    $warmupMs = $timer.Elapsed.TotalMilliseconds - $startupMs

    for ($round = 1; $round -le 3; $round++) {
        foreach ($file in $files) {
            $destination = Join-Path $outputRoot ("r{0}-{1}.xlsx" -f $round,$file.BaseName)
            if (Test-Path -LiteralPath $destination) { Remove-Item -LiteralPath $destination -Force }
            $inputBytes = (Get-Item -LiteralPath $file.FullName).Length
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $book = $null
            $verify = $null
            $worksheets = $null
            $sheet = $null
            $cell = $null
            try {
                $book = $excel.Workbooks.Open($file.FullName, 0, $false)
                $openMs = $sw.Elapsed.TotalMilliseconds
                $book.SaveAs($destination, 51) # xlOpenXMLWorkbook
                $saveMs = $sw.Elapsed.TotalMilliseconds - $openMs
                $book.Close($false)
                [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($book); $book = $null

                $verify = $excel.Workbooks.Open($destination, 0, $true)
                $worksheets = $verify.Worksheets
                $sheet = $worksheets.Item(1)
                $cell = $sheet.Cells.Item(2,1)
                $cellValue = [string]$cell.Value2
                $sheetName = [string]$sheet.Name
                [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell); $cell = $null
                [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet); $sheet = $null
                [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($worksheets); $worksheets = $null
                $verify.Close($false)
                [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($verify); $verify = $null
                if ([string]::IsNullOrWhiteSpace($sheetName) -or -not (Test-Path -LiteralPath $destination)) {
                    throw 'converted workbook readback validation failed'
                }
                $rows.Add([pscustomobject]@{
                    engine='excel-com'; round=$round; file=$file.Name; input_bytes=$inputBytes
                    output_bytes=(Get-Item -LiteralPath $destination).Length
                    open_ms=[math]::Round($openMs,3); save_ms=[math]::Round($saveMs,3)
                    total_ms=[math]::Round($sw.Elapsed.TotalMilliseconds,3); startup_ms=[math]::Round($startupMs,3)
                    status='PASS'; sheet=$sheetName; sample_cell=$cellValue; error=''
                })
            } catch {
                if ($cell) { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell) }
                if ($sheet) { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet) }
                if ($worksheets) { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($worksheets) }
                if ($verify) { try { $verify.Close($false) } catch {}; [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($verify) }
                if ($book) { try { $book.Close($false) } catch {}; [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
                $rows.Add([pscustomobject]@{
                    engine='excel-com'; round=$round; file=$file.Name; input_bytes=$inputBytes; output_bytes=0
                    open_ms=0; save_ms=0; total_ms=[math]::Round($sw.Elapsed.TotalMilliseconds,3)
                    startup_ms=[math]::Round($startupMs,3); status='FAIL'; sheet=''; sample_cell=''; error=$_.Exception.Message
                })
            }
        }
    }
} finally {
    if ($excel) { try { $excel.Quit() } catch {}; [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [gc]::Collect(); [gc]::WaitForPendingFinalizers()
}

$csv = Join-Path $evidenceRoot 'excel-com-conversion.csv'
$rows | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding utf8
@("$version", "startup_ms=$([math]::Round($startupMs,3))", "warmup_open_ms=$([math]::Round($warmupMs,3))", "input=$inputRoot", "rounds=3", "files_per_round=8", "Excel process reused across all files and rounds") |
    Set-Content -LiteralPath (Join-Path $evidenceRoot 'environment.txt') -Encoding utf8
$failed = @($rows | Where-Object status -ne 'PASS')
Write-Output "Excel COM conversion rows=$($rows.Count) passed=$($rows.Count-$failed.Count) failed=$($failed.Count) startup_ms=$([math]::Round($startupMs,3))"
if ($failed.Count) { exit 1 }
