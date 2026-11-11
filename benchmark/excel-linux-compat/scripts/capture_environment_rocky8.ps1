param([string]$OutputDirectory = '')

$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $root 'reports\evidence' }
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$image = 'rockylinux@sha256:9794037624aaa6212aeada1d28861ef5e0a935adaf93e4ef79837119f2a2d04c'
$toolchain = 'rusttools-excel-eval-rocky8-toolchain:20261111'
docker run --rm --network none $image bash -lc 'cat /etc/os-release; uname -m; getconf GNU_LIBC_VERSION; ldd --version | head -1; for t in dotnet java python3 libreoffice soffice; do command -v "$t" || echo "ABSENT $t"; done' |
    Set-Content -Encoding utf8 (Join-Path $OutputDirectory 'rocky8-runtime-environment.txt')
docker image inspect $image --format '{{json .RepoDigests}} {{.Id}}' |
    Set-Content -Encoding utf8 (Join-Path $OutputDirectory 'rocky8-image-id.txt')
docker image inspect $toolchain --format '{{json .RepoDigests}} {{.Id}}' |
    Set-Content -Encoding utf8 (Join-Path $OutputDirectory 'toolchain-image-id.txt')
Write-Output "Captured Rocky 8 environment under $OutputDirectory"

