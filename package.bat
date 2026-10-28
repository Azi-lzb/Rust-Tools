@echo off
setlocal

set "PROJECT_ROOT=%~dp0"
set "DIST_DIR=%PROJECT_ROOT%dist"
for %%R in ("%PROJECT_ROOT%.") do set "SCRIPT_ROOT=%%~fR"
where git >nul 2>nul
if errorlevel 1 (
    echo git was not found. Cannot identify this worktree safely.
    exit /b 9009
)
for /f "delims=" %%R in ('git -C "%PROJECT_ROOT%." rev-parse --show-toplevel 2^>nul') do set "GIT_ROOT=%%R"
for %%R in ("%GIT_ROOT%") do set "GIT_ROOT=%%~fR"
if /I not "%GIT_ROOT%"=="%SCRIPT_ROOT%" (
    echo This script is not at the root of the expected Rust-Tools Git worktree.
    echo Refusing to package another project.
    exit /b 3
)
for /f "delims=" %%B in ('git -C "%PROJECT_ROOT%." branch --show-current 2^>nul') do set "CURRENT_BRANCH=%%B"

if "%CURRENT_BRANCH%"=="main" (
    set "PACKAGE_SUFFIX=release"
    if /I not "%RUST_TOOLS_RELEASE_VALIDATED%"=="YES" (
        echo Official packaging requires completed main validation.
        echo After validation, set RUST_TOOLS_RELEASE_VALIDATED=YES in this command window and rerun.
        exit /b 6
    )
) else if "%CURRENT_BRANCH%"=="feature/business-development" (
    set "PACKAGE_SUFFIX=test-feature-business-development"
) else if "%CURRENT_BRANCH%"=="test/business-coverage" (
    set "PACKAGE_SUFFIX=test-business-coverage"
) else (
    echo Unsupported or detached Git branch: "%CURRENT_BRANCH%".
    echo Refusing to create a package with an ambiguous release label.
    exit /b 3
)

if not exist "%DIST_DIR%\pytools.exe" (
    echo Missing %DIST_DIR%\pytools.exe. Build this worktree explicitly first.
    exit /b 2
)
if not exist "%PROJECT_ROOT%config.xlsx" (
    echo Required local configuration is missing: %PROJECT_ROOT%config.xlsx
    exit /b 4
)
if not exist "%DIST_DIR%\python\openpyxl_worker.py" (
    echo openpyxl worker is missing: %DIST_DIR%\python\openpyxl_worker.py
    echo Run: python packaging\openpyxl\stage-windows-runtime.py --dist dist
    echo Refusing to build a package without the openpyxl runtime.
    exit /b 5
)
if not exist "%DIST_DIR%\python\python.exe" (
    echo Bundled Python is missing: %DIST_DIR%\python\python.exe
    echo Run: python packaging\openpyxl\stage-windows-runtime.py --dist dist
    echo Refusing to build a package without the openpyxl runtime.
    exit /b 5
)
echo Verifying bundled openpyxl worker protocol handshake...
"%DIST_DIR%\python\python.exe" -I "%DIST_DIR%\python\openpyxl_worker.py" --version > "%TEMP%\pytools-worker-version.txt" 2>&1
if errorlevel 1 (
    type "%TEMP%\pytools-worker-version.txt"
    echo Bundled openpyxl worker failed --version. Re-stage the runtime:
    echo   python packaging\openpyxl\stage-windows-runtime.py --dist dist
    exit /b 7
)
findstr /C:"protocol=1" "%TEMP%\pytools-worker-version.txt" >nul
if errorlevel 1 (
    type "%TEMP%\pytools-worker-version.txt"
    echo Bundled openpyxl worker reports an unexpected protocol version.
    echo Refusing to package a stale worker.
    exit /b 7
)
findstr /B /C:"openpyxl=" "%TEMP%\pytools-worker-version.txt" | findstr /V /C:"openpyxl=missing" >nul
if errorlevel 1 (
    type "%TEMP%\pytools-worker-version.txt"
    echo Bundled runtime cannot import openpyxl. Re-stage the runtime:
    echo   python packaging\openpyxl\stage-windows-runtime.py --dist dist
    exit /b 7
)
"%DIST_DIR%\python\python.exe" -I -c "import openpyxl, et_xmlfile" >nul 2>&1
if errorlevel 1 (
    echo Bundled runtime cannot import openpyxl/et_xmlfile. Re-stage the runtime:
    echo   python packaging\openpyxl\stage-windows-runtime.py --dist dist
    exit /b 7
)
del "%TEMP%\pytools-worker-version.txt" >nul 2>&1

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; $root=$env:PROJECT_ROOT; $dist=Join-Path $root 'dist'; $stamp=Get-Date -Format 'yyyyMMdd-HHmmss'; $archive=Join-Path $dist ('pytools-windows-x64-' + $env:PACKAGE_SUFFIX + '-' + $stamp + '.zip'); if (Test-Path -LiteralPath $archive) { throw 'Package already exists: ' + $archive }; $stage=Join-Path ([IO.Path]::GetTempPath()) ('pytools-package-' + [Guid]::NewGuid().ToString('N')); try { New-Item -ItemType Directory -Path $stage | Out-Null; $stageDist=Join-Path $stage 'dist'; New-Item -ItemType Directory -Path $stageDist | Out-Null; Copy-Item -LiteralPath (Join-Path $dist 'pytools.exe') -Destination $stageDist; Copy-Item -LiteralPath (Join-Path $root 'run.bat') -Destination $stage; Copy-Item -LiteralPath (Join-Path $root 'README.md') -Destination $stage; Copy-Item -LiteralPath (Join-Path $root 'config.xlsx') -Destination $stage; Copy-Item -Path (Join-Path $dist 'python') -Destination $stageDist -Recurse; Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $archive; Write-Host ('Package created: ' + $archive) } finally { if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Recurse -Force } }"
if errorlevel 1 exit /b %ERRORLEVEL%
exit /b 0
