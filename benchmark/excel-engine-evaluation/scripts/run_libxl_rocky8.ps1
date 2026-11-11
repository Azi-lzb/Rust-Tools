param(
    [Parameter(Mandatory = $true)]
    [string]$SdkRoot,
    [ValidateSet('capabilities', 'performance')]
    [string]$Mode = 'capabilities'
)

$ErrorActionPreference = 'Stop'
$experimentRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$repoRoot = (Resolve-Path (Join-Path $experimentRoot '..\..')).Path
$SdkRoot = (Resolve-Path $SdkRoot).Path
$image = 'rockylinux@sha256:9794037624aaa6212aeada1d28861ef5e0a935adaf93e4ef79837119f2a2d04c'
$buildImage = 'rusttools-excel-eval-rocky8-toolchain:20261111'
$runId = Get-Date -Format 'yyyyMMdd-HHmmss'
$tempRoot = Join-Path $env:TEMP "libxl-rocky8-$runId"
$packageRoot = Join-Path $tempRoot 'package'
$outputRoot = Join-Path $tempRoot 'output'
$inputRoot = Join-Path $tempRoot 'input'
$evidenceRoot = Join-Path $experimentRoot 'reports\evidence\libxl'
$sampleRoot = Join-Path $experimentRoot 'reports\samples\libxl'

New-Item -ItemType Directory -Force -Path $packageRoot, $outputRoot, $inputRoot, $evidenceRoot, $sampleRoot | Out-Null
$sdkVolume = "$($SdkRoot.Replace('\','/')):/opt/libxl-sdk:ro"
$experimentVolume = "$($experimentRoot.Replace('\','/')):/workspace:rw"
$packageVolume = "$($packageRoot.Replace('\','/')):/package:rw"

if ($Mode -eq 'performance') {
    Copy-Item -LiteralPath (Join-Path $experimentRoot 'fixtures\performance') -Destination (Join-Path $inputRoot 'performance') -Recurse
}

& docker pull $image 2>&1 | Tee-Object -FilePath (Join-Path $evidenceRoot 'libxl-image-pull.log')
if ($LASTEXITCODE -ne 0) { throw "Docker image pull failed with exit code $LASTEXITCODE" }

& docker run --rm --network none -v $sdkVolume -v $experimentVolume -v $packageVolume $buildImage bash /workspace/scripts/build_libxl_rocky8.sh 2>&1 |
    Tee-Object -FilePath (Join-Path $evidenceRoot 'libxl-build-container.log')
if ($LASTEXITCODE -ne 0) { throw "LibXL build failed with exit code $LASTEXITCODE" }

$packageVolumeReadOnly = "$($packageRoot.Replace('\','/'))/libxl:/opt/libxl:ro"
$outputVolume = "$($outputRoot.Replace('\','/')):/output:rw"
$inputVolume = "$($inputRoot.Replace('\','/')):/input:ro"
$experimentVolumeReadOnly = "$($experimentRoot.Replace('\','/')):/workspace:ro"
if ($Mode -eq 'performance') { $inputVolumeOption = @('-v', $inputVolume) } else { $inputVolumeOption = @() }
$runtimeLog = if ($Mode -eq 'performance') { Join-Path $evidenceRoot 'libxl-performance-runtime.log' } else { Join-Path $evidenceRoot 'libxl-offline-runtime.log' }
& docker run --rm --network none --read-only --tmpfs /tmp -v $packageVolumeReadOnly -v $outputVolume @inputVolumeOption -v $experimentVolumeReadOnly $image bash /workspace/scripts/run_libxl_offline_rocky8.sh $Mode 2>&1 |
    Tee-Object -FilePath $runtimeLog
if ($LASTEXITCODE -ne 0) { throw "LibXL offline smoke failed with exit code $LASTEXITCODE" }

if ($Mode -eq 'performance') {
    $perfEvidence = Join-Path $experimentRoot 'reports\evidence\performance\libxl'
    $perfSamples = Join-Path $experimentRoot 'reports\samples\performance\libxl'
    New-Item -ItemType Directory -Force -Path $perfEvidence, $perfSamples | Out-Null
    Copy-Item -LiteralPath (Join-Path $outputRoot 'performance-libxl.csv') -Destination $perfEvidence -Force
    Copy-Item -LiteralPath $runtimeLog -Destination $perfEvidence -Force
    foreach ($name in 'large-150k.xlsx', 'large-1m.xlsx', 'r1-batch-01.xls', 'r1-batch-09.xlsx') {
        $candidate = Join-Path (Join-Path $outputRoot 'performance') $name
        if (Test-Path $candidate) { Copy-Item -LiteralPath $candidate -Destination $perfSamples -Force }
    }
} else {
    foreach ($name in 'libxl-input.xls', 'libxl-input.xlsx', 'libxl-modified.xls', 'libxl-modified.xlsx', 'libxl-converted.xlsx', 'libxl-capabilities.xls', 'libxl-capabilities.xlsx') {
        Copy-Item -LiteralPath (Join-Path $outputRoot $name) -Destination (Join-Path $sampleRoot $name) -Force
    }
}

Write-Output "LibXL evidence: $evidenceRoot"
Write-Output "LibXL samples: $sampleRoot"
