param(
    [Parameter(Mandatory=$true)][string]$Executable,
    [Parameter(Mandatory=$true)][string]$InputPath,
    [Parameter(Mandatory=$true)][string]$OutputRoot,
    [Parameter(Mandatory=$true)][string]$SpecPath,
    [Parameter(Mandatory=$true)][string]$Oracle25Path,
    [Parameter(Mandatory=$true)][string]$Oracle35Path,
    [int]$OfficialRounds = 3
)
$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$fixtureVerifier = Join-Path $PSScriptRoot 'fixtures\synthetic\verify_packages.ps1'
$fixtureDir = Split-Path -Parent $InputPath
$rows = [Collections.Generic.List[object]]::new()

for ($index = 0; $index -le $OfficialRounds; $index++) {
    $label = if ($index -eq 0) { 'warmup' } else { "round-$index" }
    $runDir = Join-Path $OutputRoot $label
    New-Item -ItemType Directory -Force -Path $runDir | Out-Null
    $arguments = @('run-two-pass', $InputPath, $runDir, $SpecPath, $Oracle25Path, $Oracle35Path)
    $proc = Start-Process -FilePath $Executable -ArgumentList $arguments -WorkingDirectory $repoRoot `
        -PassThru -RedirectStandardOutput (Join-Path $runDir 'stdout.json') `
        -RedirectStandardError (Join-Path $runDir 'stderr.log')
    $peakWorkingSet = 0L
    $cpuMs = 0.0
    do {
        try {
            $proc.Refresh()
            if (!$proc.HasExited) {
                $peakWorkingSet = [math]::Max($peakWorkingSet, $proc.WorkingSet64)
                $cpuMs = $proc.TotalProcessorTime.TotalMilliseconds
            }
        } catch { }
        if (!$proc.HasExited) { Start-Sleep -Milliseconds 20 }
    } while (!$proc.HasExited)
    try {
        $proc.Refresh()
        $cpuMs = $proc.TotalProcessorTime.TotalMilliseconds
        $peakWorkingSet = [math]::Max($peakWorkingSet, $proc.PeakWorkingSet64)
    } catch { }
    if ($proc.ExitCode -ne 0) {
        throw "Two-pass run $label failed: $(Get-Content (Join-Path $runDir 'stderr.log') -Raw)"
    }

    $run = Get-Content -LiteralPath (Join-Path $runDir 'two-pass-runs.json') -Raw | ConvertFrom-Json
    $statuses = @(
        $run.correctness.betteroffice_two_full_passes.status,
        $run.correctness.formualizer_first_full_pass.status,
        $run.correctness.formualizer_second_full_pass.status,
        $run.correctness.betteroffice_after_input_change.status,
        $run.correctness.formualizer_after_first_incremental_pass.status,
        $run.correctness.formualizer_after_second_incremental_pass.status
    )
    $correctness = ($statuses -notcontains 'MISMATCH')
    if (!$correctness) { throw "Two-pass run $label disagreed with Excel COM" }
    if ($run.formualizer_source_cache_pass1.status -ne 'EXECUTED' -or
        $run.formualizer_source_cache_pass2.status -ne 'EXECUTED') {
        throw "Strict source-cache pass did not complete in $label"
    }

    $rows.Add([pscustomobject]@{
        round = $label
        formula_targets = $run.formula_targets
        correctness_match = $correctness
        betteroffice_full_two_pass_ms = $run.betteroffice.calculate_ms
        betteroffice_open_calc_save_ms = $run.betteroffice.load_ms + $run.betteroffice.calculate_ms + $run.betteroffice.save_ms
        formualizer_full_pass1_ms = $run.formualizer.first_full_calculate_ms
        formualizer_full_pass2_ms = $run.formualizer.second_full_calculate_ms
        formualizer_full_two_pass_ms = $run.formualizer.first_full_calculate_ms + $run.formualizer.second_full_calculate_ms
        formualizer_load_calc2_save_ms = $run.formualizer.load_ms + $run.formualizer.first_full_calculate_ms + $run.formualizer.second_full_calculate_ms + $run.formualizer.save_after_second_pass_ms
        formualizer_incremental_pass1_ms = $run.formualizer.first_incremental_calculate_ms
        formualizer_incremental_pass2_ms = $run.formualizer.second_incremental_calculate_ms
        formualizer_incremental_changed_pass1 = $run.formualizer.first_incremental_changed_cells
        formualizer_incremental_changed_pass2 = $run.formualizer.second_incremental_changed_cells
        formualizer_cache_pass1_ms = $run.formualizer_source_cache_pass1.total_ms
        formualizer_cache_pass2_ms = $run.formualizer_source_cache_pass2.total_ms
        formualizer_cache_changed_pass1 = $run.formualizer_source_cache_pass1.cache_cells_changed
        formualizer_cache_changed_pass2 = $run.formualizer_source_cache_pass2.cache_cells_changed
        process_peak_memory_mb = [math]::Round($peakWorkingSet / 1MB, 2)
        process_cpu_ms = [math]::Round($cpuMs, 2)
    })
    $run = $null
    [GC]::Collect()
}

$csvPath = Join-Path $OutputRoot 'two-pass-runs.csv'
$rows | Export-Csv -NoTypeInformation -Encoding utf8 -LiteralPath $csvPath
$official = @($rows | Where-Object round -ne 'warmup')
$medians = [ordered]@{}
foreach ($property in @(
    'betteroffice_full_two_pass_ms', 'betteroffice_open_calc_save_ms',
    'formualizer_full_pass1_ms', 'formualizer_full_pass2_ms', 'formualizer_full_two_pass_ms',
    'formualizer_load_calc2_save_ms', 'formualizer_incremental_pass1_ms',
    'formualizer_incremental_pass2_ms', 'formualizer_cache_pass1_ms',
    'formualizer_cache_pass2_ms', 'process_peak_memory_mb', 'process_cpu_ms'
)) {
    $values = @($official | ForEach-Object { [double]$_.$property } | Sort-Object)
    $medians[$property] = $values[[math]::Floor($values.Count / 2)]
}
$summary = [ordered]@{
    rounds = $OfficialRounds
    warmup_included_in_csv = $true
    formulas = $rows[0].formula_targets
    all_correctness_match = (@($official | Where-Object { !$_.correctness_match }).Count -eq 0)
    medians = $medians
}
$summaryPath = Join-Path $OutputRoot 'two-pass-summary.json'
$summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $summaryPath -Encoding utf8
Write-Output "CSV=$csvPath"
Write-Output "SUMMARY=$summaryPath"
Write-Output "ALL_CORRECTNESS_MATCH=$($summary.all_correctness_match)"
