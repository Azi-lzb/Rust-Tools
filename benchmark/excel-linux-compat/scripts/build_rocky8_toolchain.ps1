$ErrorActionPreference = 'Stop'
$experimentRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$dockerRoot = Join-Path $experimentRoot 'docker\rocky8-toolchain'
$image = 'rusttools-excel-eval-rocky8-toolchain:20261111'
& docker build --pull=false -t $image $dockerRoot
if ($LASTEXITCODE -ne 0) { throw "Rocky 8 toolchain image build failed with exit code $LASTEXITCODE" }
& docker run --rm --network none $image bash -lc 'getconf GNU_LIBC_VERSION; ldd --version | head -1; uname -m; cat /etc/os-release; gcc --version | head -1; cmake --version | head -1; rpm -qa | sort' |
    Tee-Object -FilePath (Join-Path $experimentRoot 'reports\evidence\environment\rocky8-toolchain.txt')
if ($LASTEXITCODE -ne 0) { throw "Toolchain inventory failed with exit code $LASTEXITCODE" }
