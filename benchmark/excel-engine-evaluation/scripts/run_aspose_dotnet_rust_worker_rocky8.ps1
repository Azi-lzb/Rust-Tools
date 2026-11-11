param([Parameter(Mandatory = $true)][string]$SdkRoot)

$ErrorActionPreference = 'Stop'
$experimentRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$project = Join-Path $experimentRoot 'dotnet\aspose\Asposetest.csproj'
$SdkRoot = (Resolve-Path $SdkRoot).Path
$dotnet = Join-Path $SdkRoot 'dotnet.exe'
$image = 'rockylinux@sha256:9794037624aaa6212aeada1d28861ef5e0a935adaf93e4ef79837119f2a2d04c'
$bridgeImage = 'rusttools-excel-eval-rocky8-bridge:20261112'
$runId = Get-Date -Format 'yyyyMMdd-HHmmss'
$tempRoot = Join-Path $env:TEMP "aspose-dotnet-worker-$runId"
$publishRoot = Join-Path $tempRoot 'publish'
$inputRoot = Join-Path $tempRoot 'input'
$outputRoot = Join-Path $tempRoot 'output'
$targetRoot = 'F:\RustBuild\pytools-test'
$evidence = Join-Path $experimentRoot 'reports\evidence\bridge\aspose-dotnet-worker'
$samples = Join-Path $experimentRoot 'reports\samples\bridge\aspose-dotnet-worker'
New-Item -ItemType Directory -Force -Path $publishRoot, $inputRoot, $outputRoot, $evidence, $samples, $targetRoot | Out-Null

& $dotnet restore $project -r linux-x64 --ignore-failed-sources -p:NuGetAudit=false -p:RestoreIgnoreFailedSources=true --verbosity minimal 2>&1 |
    Tee-Object -FilePath (Join-Path $evidence 'dotnet-restore.log')
if ($LASTEXITCODE -ne 0) { throw "Aspose worker restore failed: $LASTEXITCODE" }
& $dotnet publish $project -c Release -r linux-x64 --self-contained true --no-restore `
    -p:PublishTrimmed=false -p:PublishSingleFile=false -p:NuGetAudit=false `
    -o $publishRoot --verbosity minimal 2>&1 | Tee-Object -FilePath (Join-Path $evidence 'dotnet-publish.log')
if ($LASTEXITCODE -ne 0) { throw "Aspose worker publish failed: $LASTEXITCODE" }

$publishVolume = "$($publishRoot.Replace('\','/')):/app:rw"
& docker run --rm --network none -v $publishVolume $bridgeImage bash -lc 'cp -a /usr/lib64/libicu*.so* /app/'
if ($LASTEXITCODE -ne 0) { throw "Failed to bundle ICU for the self-contained worker: $LASTEXITCODE" }

$targetVolume = "$($targetRoot.Replace('\','/')):/target:rw"
$experimentVolume = "$($experimentRoot.Replace('\','/')):/workspace:rw"
& docker run --rm --network none -e CARGO_TARGET_DIR=/target -v $targetVolume -v $experimentVolume $bridgeImage `
    cargo build --release --offline --manifest-path /workspace/bridge/dotnet/rust-client/Cargo.toml 2>&1 |
    Tee-Object -FilePath (Join-Path $evidence 'rust-client-build.log')
if ($LASTEXITCODE -ne 0) { throw "Rust .NET bridge client build failed: $LASTEXITCODE" }

Copy-Item (Join-Path $experimentRoot 'fixtures\capabilities.xlsx') (Join-Path $inputRoot 'capabilities.xlsx') -Force
$appVolume = "$($publishRoot.Replace('\','/')):/opt/app:ro"
$inputVolume = "$($inputRoot.Replace('\','/')):/input:ro"
$outputVolume = "$($outputRoot.Replace('\','/')):/output:rw"
$binaryVolume = "$($targetRoot.Replace('\','/')):/target:ro"
$runtimeLog = Join-Path $evidence 'aspose-dotnet-worker-runtime.log'
& docker run --rm --network none --read-only --tmpfs /tmp -e LD_LIBRARY_PATH=/opt/app `
    -v $appVolume -v $inputVolume -v $outputVolume -v $binaryVolume `
    $image /target/release/aspose-dotnet-worker-bridge /opt/app/Asposetest /input/capabilities.xlsx /output 2>&1 |
    Tee-Object -FilePath $runtimeLog
if ($LASTEXITCODE -ne 0) { throw "Aspose .NET worker bridge run failed: $LASTEXITCODE" }

Copy-Item (Join-Path $outputRoot 'bridge-aspose-dotnet.csv') (Join-Path $evidence 'bridge-aspose-dotnet.csv') -Force
foreach ($name in 'aspose-worker-10000.xlsx', 'aspose-worker-100000.xlsx') {
    Copy-Item (Join-Path $outputRoot $name) (Join-Path $samples $name) -Force
}
Write-Output "Rust→Aspose .NET Worker evidence: $evidence"
