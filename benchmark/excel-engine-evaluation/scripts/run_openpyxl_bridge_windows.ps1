$ErrorActionPreference = 'Stop'
$experimentRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$repoRoot = (Resolve-Path (Join-Path $experimentRoot '..\..')).Path
$targetDir = 'F:\RustBuild\pytools-test\bridge-python'
$env:CARGO_TARGET_DIR = $targetDir
$project = Join-Path $experimentRoot 'bridge\python\rust-client\Cargo.toml'
$worker = Join-Path $experimentRoot 'bridge\python\worker.py'
$input = Join-Path $experimentRoot 'fixtures\capabilities.xlsx'
$output = Join-Path $experimentRoot 'reports\samples\bridge\openpyxl'
$evidence = Join-Path $experimentRoot 'reports\evidence\bridge\openpyxl'
New-Item -ItemType Directory -Force -Path $output, $evidence | Out-Null
Push-Location $repoRoot
try {
    cargo build --offline --release --manifest-path $project
    if ($LASTEXITCODE -ne 0) { throw "Rust openpyxl bridge build failed: $LASTEXITCODE" }
} finally { Pop-Location }
$python = (Get-Command python).Source
$log = Join-Path $evidence 'rust-python-worker-runtime.log'
& (Join-Path $targetDir 'release\openpyxl-worker-bridge.exe') $python $worker $input $output 2>&1 |
    Tee-Object -FilePath $log
if ($LASTEXITCODE -ne 0) { throw "Rust→openpyxl bridge failed: $LASTEXITCODE" }
Copy-Item (Join-Path $output 'bridge-openpyxl.csv') $evidence -Force
