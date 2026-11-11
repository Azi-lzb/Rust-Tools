$ErrorActionPreference = 'Stop'
$experimentRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$inputRoot = Join-Path $experimentRoot 'fixtures\performance'
$samplesRoot = Join-Path $experimentRoot 'reports\samples\performance\windows-conversion'
$evidenceRoot = Join-Path $experimentRoot 'reports\evidence\performance\windows-conversion'
$excel = $null
$results = [System.Collections.Generic.List[object]]::new()
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AskToUpdateLinks = $false
    $excel.ScreenUpdating = $false; $excel.EnableEvents = $false
    foreach ($engine in @('aspose','gembox','excel-com')) {
        $base = if ($engine -eq 'excel-com') { Join-Path $experimentRoot 'reports\samples\performance\excel-com-reference' } else { Join-Path $samplesRoot $engine }
        foreach ($input in (Get-ChildItem $inputRoot -File -Filter 'batch-*.xls' | Sort-Object Name)) {
            $source = $null; $target = $null; $srcSheets=$null; $outSheets=$null; $srcSheet=$null; $outSheet=$null; $srcRange=$null; $outRange=$null
            try {
                $output = Join-Path $base ("r1-{0}.xlsx" -f $input.BaseName)
                $source = $excel.Workbooks.Open($input.FullName,0,$true)
                $target = $excel.Workbooks.Open($output,0,$true)
                $srcSheets=$source.Worksheets; $outSheets=$target.Worksheets
                $srcSheet=$srcSheets.Item(1); $outSheet=$outSheets.Item(1)
                $srcRange=$srcSheet.UsedRange; $outRange=$outSheet.UsedRange
                $srcValues=$srcRange.Value2; $outValues=$outRange.Value2
                $srcFormulas=$srcRange.Formula; $outFormulas=$outRange.Formula
                $rows=[int]$srcRange.Rows.Count; $cols=[int]$srcRange.Columns.Count
                $differences=0
                if ($rows -ne [int]$outRange.Rows.Count -or $cols -ne [int]$outRange.Columns.Count) { $differences++ }
                if ($srcSheets.Count -ne $outSheets.Count) { $differences++ }
                for ($r=1; $r -le $rows; $r++) { for ($c=1; $c -le $cols; $c++) {
                    $sv=$srcValues[$r,$c]; $ov=$outValues[$r,$c]
                    $sf=[string]$srcFormulas[$r,$c]; $of=[string]$outFormulas[$r,$c]
                    if ([string]$sv -cne [string]$ov -or $sf -cne $of) { $differences++ }
                } }
                $results.Add([pscustomobject]@{engine=$engine;file=$input.Name;source_rows=$rows;source_columns=$cols;source_sheets=$srcSheets.Count;output_sheets=$outSheets.Count;content_differences=$differences;status=$(if($differences -eq 0){'PASS'}else{'PARTIAL'})})
            } catch {
                $results.Add([pscustomobject]@{engine=$engine;file=$input.Name;source_rows=0;source_columns=0;source_sheets=0;output_sheets=0;content_differences=-1;status='FAIL';error=$_.Exception.Message})
            } finally {
                foreach ($com in @($srcRange,$outRange,$srcSheet,$outSheet,$srcSheets,$outSheets)) { if ($com) { try {[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($com)} catch {} } }
                if ($source) { try{$source.Close($false)}catch{}; [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($source) }
                if ($target) { try{$target.Close($false)}catch{}; [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($target) }
            }
        }
    }
} finally {
    if ($excel) { try{$excel.Quit()}catch{}; [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [gc]::Collect(); [gc]::WaitForPendingFinalizers()
}
$results | Export-Csv -LiteralPath (Join-Path $evidenceRoot 'conversion-semantic-check.csv') -NoTypeInformation -Encoding utf8
$results | Group-Object engine,status | ForEach-Object { "{0} {1}={2}" -f $_.Name.Split(',')[0],$_.Name.Split(',')[1],$_.Count } |
    Set-Content -LiteralPath (Join-Path $evidenceRoot 'conversion-semantic-summary.txt') -Encoding utf8
Write-Output "Windows conversion content checks=$($results.Count) PASS=$(@($results | Where-Object status -eq 'PASS').Count) PARTIAL=$(@($results | Where-Object status -eq 'PARTIAL').Count) FAIL=$(@($results | Where-Object status -eq 'FAIL').Count)"
