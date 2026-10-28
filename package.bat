@echo off
setlocal

set "PROJECT_ROOT=%~dp0"
set "DIST_DIR=%PROJECT_ROOT%dist"
if not exist "%DIST_DIR%\pytools.exe" (
    echo Missing %DIST_DIR%\pytools.exe. Run build-release.bat explicitly first.
    exit /b 2
)
if not exist "%PROJECT_ROOT%config.xlsx" (
    echo Required local configuration is missing: %PROJECT_ROOT%config.xlsx
    exit /b 4
)
if not exist "%PROJECT_ROOT%dist\python\openpyxl_worker.py" (
    echo openpyxl worker is missing: %PROJECT_ROOT%dist\python\openpyxl_worker.py
    echo Run: python packaging\openpyxl\stage-windows-runtime.py --dist dist
    echo Refusing to build a package without the openpyxl runtime.
    exit /b 5
)
if not exist "%PROJECT_ROOT%dist\python\python.exe" (
    echo Bundled Python is missing: %PROJECT_ROOT%dist\python\python.exe
    echo Run: python packaging\openpyxl\stage-windows-runtime.py --dist dist
    echo Refusing to build a package without the openpyxl runtime.
    exit /b 5
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; $root=$env:PROJECT_ROOT; $dist=Join-Path $root 'dist'; $stamp=Get-Date -Format 'yyyyMMdd-HHmmss'; $archive=Join-Path $dist ('pytools-windows-x64-' + $stamp + '.zip'); if (Test-Path -LiteralPath $archive) { throw 'Package already exists: ' + $archive }; $stage=Join-Path ([IO.Path]::GetTempPath()) ('pytools-package-' + [Guid]::NewGuid().ToString('N')); try { New-Item -ItemType Directory -Path $stage | Out-Null; $stageDist=Join-Path $stage 'dist'; New-Item -ItemType Directory -Path $stageDist | Out-Null; Copy-Item -LiteralPath (Join-Path $dist 'pytools.exe') -Destination $stageDist; Copy-Item -LiteralPath (Join-Path $root 'run.bat') -Destination $stage; Copy-Item -LiteralPath (Join-Path $root 'README.md') -Destination $stage; Copy-Item -LiteralPath (Join-Path $root 'config.xlsx') -Destination $stage; Copy-Item -Path (Join-Path $dist 'python') -Destination $stageDist -Recurse; Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $archive; Write-Host ('Package created: ' + $archive) } finally { if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Recurse -Force } }"
if errorlevel 1 exit /b %ERRORLEVEL%
exit /b 0
