@echo off
setlocal

set "PROJECT_ROOT=%~dp0"
set "DIST_DIR=%PROJECT_ROOT%dist"
set "PYTOOLS_EXE=%DIST_DIR%\pytools.exe"
if not defined PYTOOLS_CONFIG_PATH set "PYTOOLS_CONFIG_PATH=%PROJECT_ROOT%config.xlsx"

if not exist "%PYTOOLS_EXE%" (
    echo PyTools EXE is missing: %PYTOOLS_EXE%
    echo This script never invokes Cargo. Run build-release.bat explicitly to build from source.
    pause
    exit /b 2
)
pushd "%PROJECT_ROOT%"
if errorlevel 1 (
    echo Cannot enter the project directory: %PROJECT_ROOT%
    pause
    exit /b 4
)
echo Starting the existing PyTools EXE: %PYTOOLS_EXE%
echo For development with incremental compilation, use run-dev.bat.
"%PYTOOLS_EXE%" menu %*
set "RUN_EXIT_CODE=%ERRORLEVEL%"

echo.
echo pytools.exe exited with code %RUN_EXIT_CODE%.
pause

popd
exit /b %RUN_EXIT_CODE%
