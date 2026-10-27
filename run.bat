@echo off
setlocal

set "PROJECT_ROOT=%~dp0"
set "DIST_DIR=%PROJECT_ROOT%dist"
set "PYTOOLS_EXE=%DIST_DIR%\pytools.exe"
if not defined PYTOOLS_CONFIG_PATH set "PYTOOLS_CONFIG_PATH=%PROJECT_ROOT%config.xlsx"

if not exist "%PYTOOLS_EXE%" (
    echo PyTools EXE is missing: %PYTOOLS_EXE%
    echo This script never invokes Cargo. Run build-release.bat explicitly to build from source.
    exit /b 2
)
pushd "%PROJECT_ROOT%"
if errorlevel 1 exit /b 4
"%PYTOOLS_EXE%" menu %*
set "RUN_EXIT_CODE=%ERRORLEVEL%"
popd
exit /b %RUN_EXIT_CODE%

