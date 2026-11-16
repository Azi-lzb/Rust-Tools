$ErrorActionPreference='Stop'
Set-Location -LiteralPath 'C:/Users/AZI/Desktop/VibeCode/Rust-Tools'
if ((git rev-parse --show-toplevel) -ne 'C:/Users/AZI/Desktop/VibeCode/Rust-Tools' -or (git branch --show-current) -ne 'main') { throw 'Unexpected worktree' }
$run='F:/RustBuild/worktree-retirement-20261116'
$env:CARGO_TARGET_DIR='F:/RustBuild/pytools'
$env:CARGO_BUILD_JOBS='2'
$env:CARGO_PROFILE_TEST_DEBUG='0'
$env:TEMP="$run/tmp"
$env:TMP=$env:TEMP
New-Item -ItemType Directory -Path $env:TEMP -Force | Out-Null
$env:PYTHONIOENCODING='utf-8'
$env:PYTHONDONTWRITEBYTECODE='1'
$env:RUST_WORKSPACE_OPENPYXL_PYTHON='F:/RustBuild/pytools/debug/python/python.exe'
$env:PYTOOLS_OPENPYXL_PYTHON=$env:RUST_WORKSPACE_OPENPYXL_PYTHON
$env:PYTOOLS_OPENPYXL_WORKER='C:/Users/AZI/Desktop/VibeCode/Rust-Tools/crates/excel-editor/python/openpyxl_worker.py'
cargo fmt --all *> "$run/fmt.log"
if ($LASTEXITCODE -ne 0) { throw 'Formatting failed' }
$ErrorActionPreference='Continue'
Write-Output 'START target_access_tests'
cargo test --offline --locked -p pytools-core --lib target_access_tests -- --test-threads=1 *> "$run/target-access.log"
$accessCode=$LASTEXITCODE
$accessCode | Set-Content "$run/target-access.exit"
Get-Content "$run/target-access.log" -Tail 7
Write-Output 'START summary and timeline integration tests'
cargo test --offline --locked -p pytools-core --test summary_menu_cli --test timeline_menu_cli --test summary_used_range_error --no-fail-fast -- --test-threads=1 *> "$run/summary-timeline.log"
$summaryCode=$LASTEXITCODE
$summaryCode | Set-Content "$run/summary-timeline.exit"
Get-Content "$run/summary-timeline.log" -Tail 12
if ($accessCode -ne 0 -or $summaryCode -ne 0) { exit 1 }
exit 0
