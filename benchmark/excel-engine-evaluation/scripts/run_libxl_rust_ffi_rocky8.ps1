param([Parameter(Mandatory = $true)][string]$SdkRoot)

$ErrorActionPreference = 'Stop'
$experimentRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$SdkRoot = (Resolve-Path $SdkRoot).Path
$image = 'rusttools-excel-eval-rocky8-bridge:20261112'
$runId = Get-Date -Format 'yyyyMMdd-HHmmss'
$tempRoot = Join-Path $env:TEMP "libxl-ffi-rocky8-$runId"
$packageRoot = Join-Path $tempRoot 'package'
$outputRoot = Join-Path $tempRoot 'output'
$targetRoot = 'F:\RustBuild\pytools-test'
$evidence = Join-Path $experimentRoot 'reports\evidence\bridge\libxl-ffi'
$samples = Join-Path $experimentRoot 'reports\samples\bridge\libxl-ffi'
New-Item -ItemType Directory -Force -Path $packageRoot, $outputRoot, $evidence, $samples, $targetRoot | Out-Null
Copy-Item (Join-Path $experimentRoot 'fixtures\capabilities.xlsx') (Join-Path $tempRoot 'input.xlsx')
$sdkVolume = "$($SdkRoot.Replace('\','/')):/opt/libxl-sdk:ro"
$experimentVolume = "$($experimentRoot.Replace('\','/')):/workspace:rw"
$packageVolume = "$($packageRoot.Replace('\','/')):/package:rw"
$targetVolume = "$($targetRoot.Replace('\','/')):/target:rw"
$outputVolume = "$($outputRoot.Replace('\','/')):/output:rw"
$inputVolume = "$($tempRoot.Replace('\','/')):/input:ro"

& docker run --rm --network none -e CARGO_TARGET_DIR=/target -e RUSTFLAGS='-L native=/package' `
    -v $sdkVolume -v $experimentVolume -v $packageVolume -v $targetVolume $image `
    bash /workspace/scripts/build_libxl_rust_ffi_rocky8.sh
if ($LASTEXITCODE -ne 0) { throw "Rust→LibXL FFI bridge build failed with exit code $LASTEXITCODE" }

$packageReadOnly = "$($packageRoot.Replace('\','/')):/opt/libxl:ro"
$binaryVolume = "$($targetRoot.Replace('\','/')):/target:ro"
$experimentReadOnly = "$($experimentRoot.Replace('\','/')):/workspace:ro"
$runtimeLog = Join-Path $evidence 'libxl-rust-ffi-runtime.log'
& docker run --rm --network none --read-only --tmpfs /tmp -e LD_LIBRARY_PATH=/opt/libxl `
    -v $packageReadOnly -v $binaryVolume -v $inputVolume -v $outputVolume -v $experimentReadOnly $image `
    bash /workspace/scripts/run_libxl_rust_ffi_offline_rocky8.sh 2>&1 | Tee-Object -FilePath $runtimeLog
if ($LASTEXITCODE -ne 0) { throw "Rust→LibXL FFI offline run failed with exit code $LASTEXITCODE" }
Copy-Item (Join-Path $outputRoot 'libxl-ffi-output.xlsx') (Join-Path $samples 'libxl-ffi-output.xlsx') -Force
Write-Output "Rust→LibXL FFI evidence: $evidence"
