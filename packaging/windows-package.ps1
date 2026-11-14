# RustTools Windows release packaging.
# Layout (user requirement: double-click the exe in the package root):
#   RustTools-windows-x64-<stamp>.zip
#   ├── RustTools.exe            (double-click entry; auto-locates config)
#   ├── config/config.xlsx       (configuration; probed first)
#   ├── python/                  (bundled openpyxl runtime + worker)
#   └── doc/README.md            (documentation)
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File packaging/windows-package.ps1
# Prereqs: dist/RustTools.exe and dist/python/ exist (release built and
# runtime staged via packaging/openpyxl/stage-windows-runtime.py).
param(
    [string]$ProjectRoot = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = 'Stop'
$dist = Join-Path $ProjectRoot 'dist'
$exe = Join-Path $dist 'RustTools.exe'
$pythonDir = Join-Path $dist 'python'

if (-not (Test-Path -LiteralPath $exe)) {
    throw "Missing $exe. Build release first: cargo build --release --offline -p pytools-core --bin RustTools"
}
if (-not (Test-Path -LiteralPath $pythonDir)) {
    throw "Missing $pythonDir. Stage runtime first: python packaging/openpyxl/stage-windows-runtime.py --dist dist"
}

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$archive = Join-Path $dist ("RustTools-windows-x64-$stamp.zip")
if (Test-Path -LiteralPath $archive) { throw "Package already exists: $archive" }

# Stage inside the project (GetTempPath proved flaky under -File invocation:
# the freshly created directory vanished before the first copy).
$rootPath = [System.IO.Path]::GetFullPath($ProjectRoot).TrimEnd('\', '/')
$stage = [System.IO.Path]::GetFullPath((Join-Path $rootPath ".package-stage-$stamp"))
if (-not $stage.StartsWith($rootPath + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Package staging path escapes project root: $stage"
}
if (Test-Path -LiteralPath $stage) { throw "Package staging path already exists: $stage" }
New-Item -ItemType Directory -Path $stage | Out-Null
if (-not (Test-Path -LiteralPath $stage)) { throw "stage directory was not created: $stage" }

try {
    # Package root = the user's working root: exe + config/ + python/ + doc/.
    Copy-Item -LiteralPath $exe -Destination (Join-Path $stage 'RustTools.exe')

    $stageConfig = Join-Path $stage 'config'
    New-Item -ItemType Directory -Path $stageConfig -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $ProjectRoot 'config.xlsx') -Destination (Join-Path $stageConfig 'config.xlsx')

    Copy-Item -Path $pythonDir -Destination (Join-Path $stage 'python') -Recurse

    $stageDoc = Join-Path $stage 'doc'
    New-Item -ItemType Directory -Path $stageDoc -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $ProjectRoot 'README.md') -Destination (Join-Path $stageDoc 'README.md')

    Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $archive
    Write-Host ("Package created: " + $archive)
} finally {
    if (Test-Path -LiteralPath $stage) {
        $cleanupPath = [System.IO.Path]::GetFullPath((Resolve-Path -LiteralPath $stage).Path)
        if ($cleanupPath -ne $stage -or ((Get-Item -LiteralPath $stage -Force).Attributes -band [System.IO.FileAttributes]::ReparsePoint)) {
            throw "Refusing cleanup of changed package staging path: $stage"
        }
        Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue
    }
}
