param([string]$Label='windows-workspace')
$ErrorActionPreference='Stop'
$repo=(Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
Set-Location -LiteralPath $repo
if ((git branch --show-current) -ne 'main') { throw 'Expected main' }
$run='F:/RustBuild/pytools/release-push-20261116'
$env:CARGO_TARGET_DIR='F:/RustBuild/pytools'
$env:CARGO_BUILD_JOBS='2'
$env:CARGO_PROFILE_TEST_DEBUG='0'
$env:TEMP='F:/RustBuild/pytools/aspose-task-temp'
$env:TMP=$env:TEMP
$env:PYTHONIOENCODING='utf-8'
$env:PYTHONDONTWRITEBYTECODE='1'
$env:PYTOOLS_OPENPYXL_PYTHON='F:/RustBuild/pytools/debug/python/python.exe'
$env:RUST_WORKSPACE_OPENPYXL_PYTHON=$env:PYTOOLS_OPENPYXL_PYTHON
$env:PYTOOLS_OPENPYXL_WORKER=Join-Path $repo 'crates/excel-editor/python/openpyxl_worker.py'
$env:PYTOOLS_COM_PYTHON=$env:PYTOOLS_OPENPYXL_PYTHON
$env:PYTOOLS_COM_WORKER=Join-Path $repo 'pytools-core/python/pywin32_excel_conversion_worker.py'
$env:DOTNET_ROOT='F:/RustBuild/pytools/aspose-worker-sdk-10.0.404'
$env:RUSTTOOLS_ASPOSE_DOTNET="$env:DOTNET_ROOT/dotnet.exe"
$env:RUSTTOOLS_ASPOSE_WORKER_DLL='F:/RustBuild/pytools/aspose-worker-release-print/bin/AsposeWorker.dll'
$env:RUSTTOOLS_ASPOSE_WORDS_DLL='F:/RustBuild/pytools/aspose-words-worker-release-menu50/bin/AsposeWordsWorker.dll'
$env:ASPOSE_TEST_PYTHON='C:/Users/AZI/AppData/Local/Programs/Python/Python312/python.exe'
Remove-Item Env:PYTOOLS_EXCEL_EDIT_ROUTE -ErrorAction SilentlyContinue
$ErrorActionPreference='Continue'
cargo test --offline --locked --workspace --no-fail-fast -- --test-threads=1 *> "$run/$Label.log"
$code=$LASTEXITCODE
$code | Set-Content "$run/$Label.exit"
Get-Content "$run/$Label.log" -Tail 15
exit $code
