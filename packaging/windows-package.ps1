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
    [string]$ProjectRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$StageRoot = '',
    [string]$ArchiveDirectory = '',
    [string]$WorkerDirectory = '',
    [ValidateSet('test', 'release')][string]$PackageKind = 'test',
    [switch]$ReleaseAccepted
)

$ErrorActionPreference = 'Stop'
$dist = Join-Path $ProjectRoot 'dist'
$exe = Join-Path $dist 'RustTools.exe'
$pythonDir = Join-Path $dist 'python'
$branch = (& git -C $ProjectRoot branch --show-current).Trim()
if ($LASTEXITCODE -ne 0 -or -not $branch) { throw 'Cannot identify packaging branch' }
if ($PackageKind -eq 'release' -and ($branch -ne 'main' -or -not $ReleaseAccepted)) {
    throw 'Release requires main and -ReleaseAccepted after acceptance'
}

if (-not (Test-Path -LiteralPath $exe)) {
    throw "Missing $exe. Build release first: cargo build --release --offline -p pytools-core --bin RustTools"
}
if (-not (Test-Path -LiteralPath $pythonDir)) {
    throw "Missing $pythonDir. Stage runtime first: python packaging/openpyxl/stage-windows-runtime.py --dist dist"
}

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
if (-not $StageRoot) { $StageRoot = $ProjectRoot }
if (-not $ArchiveDirectory) { $ArchiveDirectory = $dist }
if (-not (Test-Path -LiteralPath $StageRoot -PathType Container)) { throw 'StageRoot must exist' }
if (-not (Test-Path -LiteralPath $ArchiveDirectory -PathType Container)) { throw 'ArchiveDirectory must exist' }
$packageLabel = if ($PackageKind -eq 'test') { 'TEST-' + ($branch -replace '[/\\]', '-') + '-' } else { '' }
$archive = Join-Path $ArchiveDirectory ("RustTools-windows-x64-$packageLabel$stamp.zip")
if (Test-Path -LiteralPath $archive) { throw "Package already exists: $archive" }

# Stage inside the project (GetTempPath proved flaky under -File invocation:
# the freshly created directory vanished before the first copy).
$rootPath = [System.IO.Path]::GetFullPath($StageRoot).TrimEnd('\', '/')
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

    # Optional self-contained workers. No SDK, licenses, PDB or evaluation
    # samples are packaged; Auto never depends on these engines.
    $workers = if ($WorkerDirectory) { $WorkerDirectory } else { Join-Path $dist 'workers' }
    if (Test-Path -LiteralPath $workers) {
        foreach ($worker in @(@('aspose', 'AsposeWorker', 'Aspose.Cells'), @('aspose-words', 'AsposeWordsWorker', 'Aspose.Words'))) {
            $source = Join-Path $workers $worker[0]
            if (-not (Test-Path -LiteralPath $source)) { continue }
            foreach ($required in @(($worker[1] + '.exe'), ($worker[1] + '.dll'), ($worker[1] + '.runtimeconfig.json'), ($worker[2] + '.dll'), 'coreclr.dll')) {
                if (-not (Test-Path -LiteralPath (Join-Path $source $required) -PathType Leaf)) {
                    throw "Incomplete self-contained worker $source : $required"
                }
            }
            $destination = Join-Path (Join-Path $stage 'workers') $worker[0]
            New-Item -ItemType Directory -Path $destination -Force | Out-Null
            Get-ChildItem -LiteralPath $source -Force | Where-Object { $_.Extension -ne '.pdb' } | Copy-Item -Destination $destination -Recurse
        }
        $manifest = Join-Path $workers 'MANIFEST.json'
        if (Test-Path -LiteralPath $manifest -PathType Leaf) {
            Copy-Item -LiteralPath $manifest -Destination (Join-Path $stage 'workers/MANIFEST.json')
        }
    }

    $stageDoc = Join-Path $stage 'doc'
    New-Item -ItemType Directory -Path $stageDoc -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $ProjectRoot 'README.md') -Destination (Join-Path $stageDoc 'README.md')
    foreach ($guide in @('TEST_PACKAGE_20261116.md', 'ASPOSE_FULL_MENU_TRIAL_20261115.md', 'MENU_ENGINE_REFERENCE_50FILES_20261115.md')) {
        $guidePath = Join-Path (Join-Path $ProjectRoot 'docs') $guide
        if (Test-Path -LiteralPath $guidePath -PathType Leaf) { Copy-Item -LiteralPath $guidePath -Destination $stageDoc }
    }

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
