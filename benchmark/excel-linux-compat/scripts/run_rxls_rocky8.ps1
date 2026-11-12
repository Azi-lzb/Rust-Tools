param(
    [string]$Image = 'rusttools-excel-linux-compat-rxls:20261112',
    [string]$InspectionImage = 'rusttools-excel-eval-rocky8-toolchain:20261111'
)

$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$repo = (Resolve-Path (Join-Path $root '..\..')).Path
$source = Join-Path $root 'rxls\source'
$target = 'F:\RustBuild\pytools-test\excel-linux-compat\rust-target'
$evidence = Join-Path $root 'reports\evidence\rxls'
$sample = Join-Path $root 'reports\samples\rxls'
$inputRoot = Join-Path $root 'fixtures'
$runId = Get-Date -Format 'yyyyMMdd-HHmmss'
$output = Join-Path $env:TEMP "excel-linux-compat-rxls-output-$runId"
New-Item -ItemType Directory -Force -Path $target, $evidence, $sample, $output | Out-Null

& docker build --pull=false -t $Image (Join-Path $root 'docker\rust-rocky8') 2>&1 |
    Tee-Object -FilePath (Join-Path $evidence 'toolchain-build.log')
if ($LASTEXITCODE -ne 0) { throw "Rust Rocky 8 image build failed: $LASTEXITCODE" }

$sourceVolume = "$($source.Replace('\','/')):/workspace:ro"
$targetVolume = "$($target.Replace('\','/')):/target:rw"
& docker run --rm -v $sourceVolume -v $targetVolume -e CARGO_TARGET_DIR=/target $Image `
    bash -lc 'cd /workspace && cargo build --release --locked' 2>&1 |
    Tee-Object -FilePath (Join-Path $evidence 'rxls-build.log')
if ($LASTEXITCODE -ne 0) { throw "rXLS release build failed: $LASTEXITCODE" }

$binary = Join-Path $target 'release\xls2x-bench'
if (-not (Test-Path $binary)) { throw "Missing Linux rXLS binary: $binary" }
$binaryVolume = "$((Join-Path $target 'release').Replace('\','/')):/opt/app:ro"
$inspectVolume = "$($target.Replace('\','/')):/opt/target:ro"
& docker run --rm --network none -v $inspectVolume $InspectionImage bash -lc 'set -e; file /opt/target/release/xls2x-bench; readelf -h /opt/target/release/xls2x-bench; readelf -d /opt/target/release/xls2x-bench; readelf --version-info /opt/target/release/xls2x-bench | grep -E "GLIBC(_|XX_)|CXXABI_" | sort -u; ldd /opt/target/release/xls2x-bench; ldd -r /opt/target/release/xls2x-bench' 2>&1 |
    Tee-Object -FilePath (Join-Path $evidence 'rxls-elf-dependencies.txt')
if ($LASTEXITCODE -ne 0) { throw "rXLS ELF inspection failed: $LASTEXITCODE" }
$inputVolume = "$($inputRoot.Replace('\','/')):/input:ro"
$outputVolume = "$($output.Replace('\','/')):/output:rw"
$imageRocky = 'rockylinux@sha256:9794037624aaa6212aeada1d28861ef5e0a935adaf93e4ef79837119f2a2d04c'
& docker run --rm --network none --read-only --tmpfs /tmp -v $binaryVolume -v $inputVolume -v $outputVolume $imageRocky `
    bash -lc 'set -e; uname -m; getconf GNU_LIBC_VERSION; ldd --version | sed -n "1p"; for t in dotnet java python3 libreoffice soffice; do if command -v "$t" >/dev/null 2>&1; then printf "PRESENT %s\n" "$t"; else printf "ABSENT %s\n" "$t"; fi; done; /opt/app/xls2x-bench convert /input/capabilities.xls --out-dir /output --backend rxls --validate full --report /output/rxls-convert-report.json --with-diag' 2>&1 |
    Tee-Object -FilePath (Join-Path $evidence 'rxls-offline-runtime.log')
if ($LASTEXITCODE -ne 0) { throw "rXLS Rocky 8 offline conversion failed: $LASTEXITCODE" }

Copy-Item -LiteralPath (Join-Path $output 'capabilities.xlsx') -Destination (Join-Path $sample 'capabilities-converted.xlsx') -Force
Copy-Item -LiteralPath (Join-Path $output 'rxls-convert-report.json') -Destination (Join-Path $evidence 'rxls-convert-report.json') -Force
Write-Output "rXLS Rocky 8 evidence: $evidence"
