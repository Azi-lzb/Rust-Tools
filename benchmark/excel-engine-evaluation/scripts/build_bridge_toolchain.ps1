$ErrorActionPreference = 'Stop'
$experimentRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$dockerRoot = Join-Path $experimentRoot 'docker\rocky8-toolchain'
$image = 'rusttools-excel-eval-rocky8-bridge:20261112'
& docker build --pull=false -t $image $dockerRoot
if ($LASTEXITCODE -ne 0) { throw "Rocky 8 bridge toolchain build failed with exit code $LASTEXITCODE" }
$environmentEvidence = Join-Path $experimentRoot 'reports\evidence\environment\rocky8-bridge-toolchain.txt'
& docker run --rm --network none $image bash -lc 'getconf GNU_LIBC_VERSION; rustc --version; cargo --version; g++ --version | head -1; rpm -qa | sort' |
    Tee-Object -FilePath $environmentEvidence
if ($LASTEXITCODE -ne 0) { throw "Bridge toolchain inventory failed with exit code $LASTEXITCODE" }
