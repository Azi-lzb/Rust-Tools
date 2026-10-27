@echo off
setlocal

set "PROJECT_ROOT=%~dp0"
set "CARGO_TARGET_DIR=F:\RustBuild\pytools"
where cargo >nul 2>nul
if errorlevel 1 (
    echo cargo was not found. No build was started.
    exit /b 9009
)

pushd "%PROJECT_ROOT%"
if errorlevel 1 exit /b 3

echo Building PyTools Release binaries from this source into %CARGO_TARGET_DIR%.
cargo build --release --offline --locked -p pytools-core --bin pytools
if errorlevel 1 goto build_failed
cargo build --release --offline --locked -p lo-convert
if errorlevel 1 goto build_failed

if not exist "%PROJECT_ROOT%dist" mkdir "%PROJECT_ROOT%dist"
copy /y "%CARGO_TARGET_DIR%\release\pytools.exe" "%PROJECT_ROOT%dist\pytools.exe" >nul
if errorlevel 1 goto copy_failed
copy /y "%CARGO_TARGET_DIR%\release\lo-convert.exe" "%PROJECT_ROOT%dist\lo-convert.exe" >nul
if errorlevel 1 goto copy_failed

echo Release EXEs copied to %PROJECT_ROOT%dist.
popd
exit /b 0

:build_failed
set "BUILD_EXIT_CODE=%ERRORLEVEL%"
popd
echo Release build failed with exit code %BUILD_EXIT_CODE%.
exit /b %BUILD_EXIT_CODE%

:copy_failed
set "COPY_EXIT_CODE=%ERRORLEVEL%"
popd
echo Build succeeded, but copying the EXEs into dist failed with exit code %COPY_EXIT_CODE%.
exit /b %COPY_EXIT_CODE%
