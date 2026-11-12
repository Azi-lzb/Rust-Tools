param(
    [Parameter(Mandatory = $true)]
    [string]$SdkRoot
)

$ErrorActionPreference = 'Stop'
$experimentRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$SdkRoot = (Resolve-Path $SdkRoot).Path
$sampleRoot = Join-Path $experimentRoot 'reports\samples\libxl'
$evidenceRoot = Join-Path $experimentRoot 'reports\evidence\aspose-cpp'
$image = 'rockylinux@sha256:9794037624aaa6212aeada1d28861ef5e0a935adaf93e4ef79837119f2a2d04c'
$toolchainImage = 'rusttools-excel-eval-rocky8-toolchain:20261111'
$runId = Get-Date -Format 'yyyyMMdd-HHmmss'
$tempRoot = Join-Path $env:TEMP "aspose-cpp-rocky8-$runId"
$packageRoot = Join-Path $tempRoot 'package'
$outputRoot = Join-Path $tempRoot 'output'
New-Item -ItemType Directory -Force -Path $evidenceRoot, $packageRoot, $outputRoot | Out-Null
Copy-Item -LiteralPath (Join-Path $sampleRoot 'libxl-input.xlsx') -Destination (Join-Path $outputRoot 'input.xlsx')
Copy-Item -LiteralPath (Join-Path $sampleRoot 'libxl-input.xls') -Destination (Join-Path $outputRoot 'input.xls')
$sdkVolume = "$($SdkRoot.Replace('\','/')):/opt/aspose-sdk:ro"
$experimentVolume = "$($experimentRoot.Replace('\','/')):/workspace:rw"
$packageVolume = "$($packageRoot.Replace('\','/')):/package:rw"
& docker run --rm --network none -v $sdkVolume -v $experimentVolume -v $packageVolume $toolchainImage bash /workspace/scripts/build_aspose_cpp_rocky8.sh 2>&1 |
    Tee-Object -FilePath (Join-Path $evidenceRoot 'aspose-cpp-build-container.log')
if ($LASTEXITCODE -ne 0) { throw "Aspose.Cells C++ build failed with exit code $LASTEXITCODE" }

$packageVolumeReadOnly = "$($packageRoot.Replace('\','/'))/aspose-cpp:/opt/aspose-cpp:ro"
$outputVolume = "$($outputRoot.Replace('\','/')):/output:rw"
$experimentVolumeReadOnly = "$($experimentRoot.Replace('\','/')):/workspace:ro"
& docker run --rm --network none --read-only --tmpfs /tmp -v $packageVolumeReadOnly -v $outputVolume -v $experimentVolumeReadOnly $image bash /workspace/scripts/run_aspose_cpp_offline_rocky8.sh 2>&1 |
    Tee-Object -FilePath (Join-Path $evidenceRoot 'aspose-cpp-offline-runtime.log')
if ($LASTEXITCODE -ne 0) { throw "Aspose.Cells C++ offline smoke failed with exit code $LASTEXITCODE" }

Copy-Item -LiteralPath (Join-Path $outputRoot 'modified.xlsx') -Destination (Join-Path $experimentRoot 'reports\samples\aspose-cpp-modified.xlsx') -Force
Copy-Item -LiteralPath (Join-Path $outputRoot 'modified.xls') -Destination (Join-Path $experimentRoot 'reports\samples\aspose-cpp-modified.xls') -Force
Write-Output "Aspose.Cells C++ evidence: $evidenceRoot"
