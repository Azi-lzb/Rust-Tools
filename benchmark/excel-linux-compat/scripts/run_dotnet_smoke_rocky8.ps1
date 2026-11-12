param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('aspose', 'gembox')]
    [string]$Engine,
    [ValidateSet('capabilities', 'performance')]
    [string]$Mode = 'capabilities',
    [Parameter(Mandatory = $true)]
    [string]$SdkRoot
)

$ErrorActionPreference = 'Stop'
$experimentRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$repoRoot = (Resolve-Path (Join-Path $experimentRoot '..\..')).Path
$fixtureRoot = Join-Path $experimentRoot 'fixtures'
$projectDir = Join-Path $experimentRoot "dotnet\$Engine"
$evidenceRoot = Join-Path $experimentRoot "reports\evidence\$Engine"
$image = 'rockylinux@sha256:9794037624aaa6212aeada1d28861ef5e0a935adaf93e4ef79837119f2a2d04c'
$runId = Get-Date -Format 'yyyyMMdd-HHmmss'
$tempRoot = Join-Path $env:TEMP "$Engine-rocky8-$runId"
$publishRoot = Join-Path $tempRoot 'publish'
$outputRoot = Join-Path $tempRoot 'output'
$inputRoot = Join-Path $tempRoot 'input'
$appName = if ($Engine -eq 'aspose') { 'Asposetest' } else { 'GemboxTest' }

New-Item -ItemType Directory -Force -Path $evidenceRoot, $publishRoot, $outputRoot, $inputRoot | Out-Null
$project = Get-ChildItem $projectDir -Filter '*.csproj' | Select-Object -First 1
if (-not $project) { throw "Missing project file in $projectDir" }
$dotnet = Join-Path $SdkRoot 'dotnet.exe'
if (-not (Test-Path $dotnet)) { throw ".NET SDK executable missing: $dotnet" }
& $dotnet --info | Out-File -Encoding utf8 (Join-Path $evidenceRoot 'dotnet-sdk-info.txt')
& $dotnet restore $project.FullName -r linux-x64 --ignore-failed-sources -p:NuGetAudit=false -p:RestoreIgnoreFailedSources=true --verbosity minimal 2>&1 |
    Tee-Object -FilePath (Join-Path $evidenceRoot 'dotnet-restore.log')
if ($LASTEXITCODE -ne 0) { throw "$Engine package restore failed with exit code $LASTEXITCODE" }
& $dotnet publish $project.FullName -c Release -r linux-x64 --self-contained true --no-restore `
    -p:PublishTrimmed=false -p:PublishSingleFile=false -p:NuGetAudit=false `
    -o $publishRoot --verbosity minimal 2>&1 |
    Tee-Object -FilePath (Join-Path $evidenceRoot 'dotnet-publish.log')
if ($LASTEXITCODE -ne 0) { throw "$Engine self-contained publish failed with exit code $LASTEXITCODE" }
$toolchainImage = 'rusttools-excel-eval-rocky8-toolchain:20261111'
$publishVolumeWritable = "$($publishRoot.Replace('\','/')):/app:rw"
& docker run --rm --network none -v $publishVolumeWritable $toolchainImage bash -lc 'cp -a /usr/lib64/libicu*.so* /app/' 2>&1 |
    Tee-Object -FilePath (Join-Path $evidenceRoot 'bundle-icu.log')
if ($LASTEXITCODE -ne 0) { throw "$Engine ICU dependency bundle failed with exit code $LASTEXITCODE" }
$fontBundleVolume = "$($publishRoot.Replace('\','/')):/app:rw"
& docker run --rm --network none -v $fontBundleVolume $toolchainImage bash -lc 'mkdir -p /app/fonts && cp -a /usr/share/fonts/dejavu/*.ttf /app/fonts/ && cp /usr/share/doc/dejavu-fonts-common/LICENSE /app/fonts/LICENSE.txt' 2>&1 |
    Tee-Object -FilePath (Join-Path $evidenceRoot 'bundle-fonts.log')
if ($LASTEXITCODE -ne 0) { throw "$Engine private font bundle failed with exit code $LASTEXITCODE" }

Copy-Item -LiteralPath (Join-Path $fixtureRoot 'capabilities.xlsx') -Destination (Join-Path $inputRoot 'capabilities.xlsx')
Copy-Item -LiteralPath (Join-Path $fixtureRoot 'capabilities.xls') -Destination (Join-Path $inputRoot 'capabilities.xls')
if ($Mode -eq 'performance') {
    Copy-Item -LiteralPath (Join-Path $fixtureRoot 'performance') -Destination (Join-Path $inputRoot 'performance') -Recurse
    New-Item -ItemType Directory -Force -Path (Join-Path $outputRoot 'performance') | Out-Null
}
$publishVolume = "$($publishRoot.Replace('\','/')):/opt/app:ro"
$outputVolume = "$($outputRoot.Replace('\','/')):/output:rw"
$inputVolume = "$($inputRoot.Replace('\','/')):/input:ro"
$experimentVolumeReadOnly = "$($experimentRoot.Replace('\','/')):/workspace:ro"
& docker run --rm --network none --read-only --tmpfs /tmp -v $publishVolume -v $outputVolume -v $inputVolume -v $experimentVolumeReadOnly $image `
    bash /workspace/scripts/run_dotnet_smoke_rocky8.sh $Engine $appName $Mode 2>&1 |
    Tee-Object -FilePath (Join-Path $evidenceRoot 'offline-runtime.log')
if ($LASTEXITCODE -ne 0) { throw "$Engine offline Rocky 8 smoke failed with exit code $LASTEXITCODE" }

if ($Mode -eq 'performance') {
    $perfEvidence = Join-Path $experimentRoot "reports\evidence\performance\$Engine"
    $perfSamples = Join-Path $experimentRoot "reports\samples\performance\$Engine"
    New-Item -ItemType Directory -Force -Path $perfEvidence, $perfSamples | Out-Null
    Copy-Item -LiteralPath (Join-Path $outputRoot "performance-$Engine.csv") -Destination $perfEvidence -Force
    Copy-Item -LiteralPath (Join-Path $evidenceRoot 'offline-runtime.log') -Destination $perfEvidence -Force
    foreach ($name in 'large-150k.xlsx', 'large-1m.xlsx', 'r1-batch-01.xls', 'r1-batch-09.xlsx') {
        $candidate = Join-Path (Join-Path $outputRoot 'performance') $name
        if (Test-Path $candidate) { Copy-Item -LiteralPath $candidate -Destination $perfSamples -Force }
    }
} else {
    New-Item -ItemType Directory -Force -Path (Join-Path $experimentRoot "reports\samples\$Engine") | Out-Null
    Get-ChildItem $outputRoot -File | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination (Join-Path (Join-Path $experimentRoot "reports\samples\$Engine") $_.Name) -Force
    }
}
Write-Output "$Engine evidence: $evidenceRoot"
