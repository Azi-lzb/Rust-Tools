param(
    [Parameter(Mandatory=$true)][string]$Executable,
    [Parameter(Mandatory=$true)][string]$InputPath,
    [Parameter(Mandatory=$true)][string]$OutputRoot,
    [Parameter(Mandatory=$true)][string]$SpecPath,
    [Parameter(Mandatory=$true)][string]$ReferenceOraclePath,
    [Parameter(Mandatory=$true)][string]$IncrementalOraclePath,
    [int]$OfficialRounds = 3
)
$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$fixtureVerifier = Join-Path $PSScriptRoot 'fixtures\synthetic\verify_packages.ps1'
$rows = [Collections.Generic.List[object]]::new()
$runCount = $OfficialRounds + 1 # first run warms the binary and filesystem cache.
for ($index = 0; $index -lt $runCount; $index++) {
    $label = if ($index -eq 0) { 'warmup' } else { "round-$index" }
    $runDir = Join-Path $OutputRoot $label
    New-Item -ItemType Directory -Force -Path $runDir | Out-Null
    $runJson = Join-Path $runDir 'engine-runs.json'
    $stdout = Join-Path $runDir 'stdout.json'
    $stderr = Join-Path $runDir 'stderr.log'
    $runArguments = @(
        'run', $InputPath, $runDir, $SpecPath,
        $ReferenceOraclePath, $IncrementalOraclePath
    )
    $proc = Start-Process -FilePath $Executable -ArgumentList $runArguments -WorkingDirectory $repoRoot `
        -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr
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
    if ($proc.ExitCode -ne 0) { throw "Benchmark run $label failed with exit code $($proc.ExitCode): $(Get-Content $stderr -Raw)" }
    $run = Get-Content -LiteralPath $runJson -Raw | ConvertFrom-Json
    $verifierOutput = Join-Path $runDir 'package-verification.json'
    & $fixtureVerifier -FixtureDir (Split-Path -Parent $InputPath) -OutputDir $runDir `
        -OraclePath $ReferenceOraclePath -InputPath $InputPath -SpecPath $SpecPath | Out-Null
    $verification = Get-Content -LiteralPath $verifierOutput -Raw | ConvertFrom-Json
    $correctness = @(
        $run.correctness.betteroffice_full.status,
        $run.correctness.formualizer_full.status,
        $run.correctness.betteroffice_after_input_change.status,
        $run.correctness.formualizer_after_input_change.status
    ) -notcontains 'MISMATCH'
    $betterFidelity = $verification.cache_correctness.betteroffice.status
    $cacheFidelity = $verification.cache_correctness.formualizer_cache.status
    $formulaCount = $run.betteroffice.formula_targets
    $better = $run.betteroffice
    $form = $run.formualizer
    $cache = $run.formualizer_cache_strict
    $rows.Add([pscustomobject]@{
        round = $label
        formulas = $formulaCount
        exit_code = $proc.ExitCode
        process_peak_memory_mb = [math]::Round($peakWorkingSet / 1MB, 2)
        process_cpu_ms = [math]::Round($cpuMs, 2)
        correctness_match = $correctness
        betteroffice_fidelity = $betterFidelity
        formualizer_cache_fidelity = $cacheFidelity
        betteroffice_load_ms = $better.load_ms
        betteroffice_calculate_ms = $better.calculate_ms
        betteroffice_extract_ms = $better.extract_ms
        betteroffice_save_ms = $better.save_ms
        betteroffice_total_ms = $better.load_ms + $better.calculate_ms + $better.extract_ms + $better.save_ms
        formualizer_load_ms = $form.load_ms
        formualizer_calculate_ms = $form.calculate_ms
        formualizer_extract_ms = $form.extract_ms
        formualizer_reserialize_ms = $form.save_ms
        formualizer_resident_total_ms = $form.load_ms + $form.calculate_ms + $form.extract_ms + $form.save_ms
        formualizer_cache_total_ms = $cache.total_ms
        formualizer_formula_cells = $cache.formula_cells
        formualizer_cache_cells_changed = $cache.cache_cells_changed
        formualizer_cache_changed_parts = $cache.worksheet_parts_changed
    })
}
$csvPath = Join-Path $OutputRoot 'performance-runs.csv'
$rows | Export-Csv -NoTypeInformation -Encoding utf8 -LiteralPath $csvPath
$official = @($rows | Where-Object round -ne 'warmup')
$summary = [ordered]@{ rounds=$OfficialRounds; warmup_included_in_csv=$true; input=$InputPath; formulas=$rows[0].formulas; correctness_valid=(@($official | Where-Object { !$_.correctness_match }).Count -eq 0); output_fidelity_valid=(@($official | Where-Object { $_.betteroffice_fidelity -ne 'MATCH' -or $_.formualizer_cache_fidelity -ne 'MATCH' }).Count -eq 0); medians=@{} }
foreach ($property in @('betteroffice_total_ms','betteroffice_calculate_ms','formualizer_resident_total_ms','formualizer_calculate_ms','formualizer_cache_total_ms','process_peak_memory_mb','process_cpu_ms')) {
    $sorted = @($official | ForEach-Object { [double]$_.$property } | Sort-Object)
    $middle = [math]::Floor($sorted.Count / 2)
    $median = if (($sorted.Count % 2) -eq 1) { $sorted[$middle] } else { ($sorted[$middle - 1] + $sorted[$middle]) / 2 }
    $summary.medians[$property] = $median
}
$summaryPath = Join-Path $OutputRoot 'performance-summary.json'
$summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $summaryPath -Encoding utf8
Write-Output "PERFORMANCE_CSV=$csvPath"
Write-Output "SUMMARY=$summaryPath"
Write-Output ("CORRECTNESS_VALID={0}; OUTPUT_FIDELITY_VALID={1}" -f $summary.correctness_valid, $summary.output_fidelity_valid)
