param(
    [string]$Results = (Join-Path $PSScriptRoot '..\results')
)
$ErrorActionPreference = 'Stop'
$Results = [System.IO.Path]::GetFullPath($Results)
$Exe = 'F:\RustBuild\pytools-test\release\betteroffice-modification-audit.exe'
$Stdout = Join-Path $Results 'perf.stdout.txt'
$Stderr = Join-Path $Results 'perf.stderr.txt'
$Started = Get-Date
$Proc = Start-Process -FilePath $Exe -ArgumentList @('perf', $Results) -NoNewWindow -PassThru -RedirectStandardOutput $Stdout -RedirectStandardError $Stderr
$PeakWorkingSet = 0L
$PeakPrivate = 0L
do {
    try {
        $Proc.Refresh()
        $PeakWorkingSet = [Math]::Max($PeakWorkingSet, [long]$Proc.WorkingSet64)
        $PeakPrivate = [Math]::Max($PeakPrivate, [long]$Proc.PrivateMemorySize64)
    } catch { }
    Start-Sleep -Milliseconds 25
    $Proc.Refresh()
} while (-not $Proc.HasExited)
$Proc.WaitForExit()
$Summary = [ordered]@{
    exit_code = $Proc.ExitCode
    wall_ms = [Math]::Round(((Get-Date) - $Started).TotalMilliseconds, 2)
    sampled_peak_working_set_mb = [Math]::Round($PeakWorkingSet / 1MB, 2)
    sampled_peak_private_mb = [Math]::Round($PeakPrivate / 1MB, 2)
    process_peak_working_set_mb = [Math]::Round($Proc.PeakWorkingSet64 / 1MB, 2)
    cpu_ms = if ($Proc.TotalProcessorTime) { [Math]::Round($Proc.TotalProcessorTime.TotalMilliseconds, 2) } else { $null }
    stdout = Get-Content -Raw $Stdout
    stderr = Get-Content -Raw $Stderr
}
$Summary | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 (Join-Path $Results 'perf-process.json')
$Summary | ConvertTo-Json -Depth 5
if ($Proc.ExitCode -ne 0) { exit $Proc.ExitCode }
