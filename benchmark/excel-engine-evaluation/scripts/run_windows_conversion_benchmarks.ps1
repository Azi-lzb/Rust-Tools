param([Parameter(Mandatory = $true)][string]$DotnetRoot)

$ErrorActionPreference = 'Stop'
$experimentRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$dotnet = Join-Path (Resolve-Path $DotnetRoot).Path 'dotnet.exe'
$inputRoot = Join-Path $experimentRoot 'fixtures\performance'
$evidenceRoot = Join-Path $experimentRoot 'reports\evidence\performance\windows-conversion'
$samplesRoot = Join-Path $experimentRoot 'reports\samples\performance\windows-conversion'
$runId = Get-Date -Format 'yyyyMMdd-HHmmss'
New-Item -ItemType Directory -Force -Path $evidenceRoot,$samplesRoot | Out-Null

foreach ($engine in @('aspose','gembox')) {
    $projectDir = Join-Path $experimentRoot "dotnet\$engine"
    $project = Get-ChildItem $projectDir -Filter '*.csproj' | Select-Object -First 1
    $publish = Join-Path $env:TEMP "excel-$engine-win-conversion-$runId"
    $output = Join-Path $samplesRoot $engine
    New-Item -ItemType Directory -Force -Path $publish,$output | Out-Null
    & $dotnet restore $project.FullName -r win-x64 --ignore-failed-sources -p:NuGetAudit=false -p:RestoreIgnoreFailedSources=true --verbosity minimal 2>&1 |
        Tee-Object -FilePath (Join-Path $evidenceRoot "$engine-restore.log")
    if ($LASTEXITCODE -ne 0) { throw "$engine Windows restore failed: $LASTEXITCODE" }
    & $dotnet publish $project.FullName -c Release -r win-x64 --self-contained true --no-restore `
        -p:PublishTrimmed=false -p:PublishSingleFile=false -p:NuGetAudit=false `
        -o $publish --verbosity minimal 2>&1 | Tee-Object -FilePath (Join-Path $evidenceRoot "$engine-publish.log")
    if ($LASTEXITCODE -ne 0) { throw "$engine Windows publish failed: $LASTEXITCODE" }
    $app = if ($engine -eq 'aspose') { 'Asposetest.exe' } else { 'GemboxTest.exe' }
    $csv = Join-Path $evidenceRoot "$engine-conversion.csv"
    & (Join-Path $publish $app) --convert $inputRoot $output $csv 2>&1 |
        Tee-Object -FilePath (Join-Path $evidenceRoot "$engine-runtime.log")
    if ($LASTEXITCODE -ne 0) { throw "$engine Windows conversion run failed: $LASTEXITCODE" }
}
$os = Get-CimInstance Win32_OperatingSystem
$cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
@("OS=$($os.Caption) $($os.Version)","CPU=$($cpu.Name)","logical_processors=$($cpu.NumberOfLogicalProcessors)","physical_memory_bytes=$($os.TotalVisibleMemorySize * 1024)","input=$inputRoot","workbooks=8 XLS x 3 rounds; one resident Release process per engine") |
    Set-Content -LiteralPath (Join-Path $evidenceRoot 'environment.txt') -Encoding utf8
