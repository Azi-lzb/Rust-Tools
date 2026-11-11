$ErrorActionPreference = 'Stop'
$experimentRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$repoRoot = (Resolve-Path (Join-Path $experimentRoot '..\..')).Path
$targetDir = 'F:\RustBuild\pytools-test\betteroffice-reference'
$env:CARGO_TARGET_DIR = $targetDir
$manifest = Join-Path $experimentRoot 'bridge\betteroffice\Cargo.toml'
$evidence = Join-Path $experimentRoot 'reports\evidence\formula-reference'
New-Item -ItemType Directory -Force -Path $evidence | Out-Null
Push-Location $repoRoot
try {
    cargo run --offline --release --manifest-path $manifest 2>&1 |
        Tee-Object -FilePath (Join-Path $evidence 'betteroffice-formula-runtime.log') |
        Set-Content -LiteralPath (Join-Path $evidence 'betteroffice-formula.csv') -Encoding utf8
    if ($LASTEXITCODE -ne 0) { throw "BetterOffice formula reference failed: $LASTEXITCODE" }
} finally { Pop-Location }
