@echo off
setlocal

set "PROJECT_ROOT=%~dp0"
for %%R in ("%PROJECT_ROOT%.") do set "SCRIPT_ROOT=%%~fR"
where git >nul 2>nul
if errorlevel 1 (
    echo git was not found. Cannot identify this worktree safely.
    pause
    exit /b 9009
)
for /f "delims=" %%R in ('git -C "%SCRIPT_ROOT%" rev-parse --show-toplevel 2^>nul') do set "GIT_ROOT=%%R"
for %%R in ("%GIT_ROOT%") do set "GIT_ROOT=%%~fR"
if /I not "%GIT_ROOT%"=="%SCRIPT_ROOT%" (
    echo This script is not at the root of the expected Rust-Tools Git worktree.
    echo Script root: %SCRIPT_ROOT%
    echo Git root: "%GIT_ROOT%"
    pause
    exit /b 3
)
for /f "delims=" %%B in ('git -C "%SCRIPT_ROOT%" branch --show-current 2^>nul') do set "CURRENT_BRANCH=%%B"

if "%CURRENT_BRANCH%"=="main" (
    set "CARGO_TARGET_DIR=F:\RustBuild\pytools"
) else if "%CURRENT_BRANCH%"=="feature/business-development" (
    set "CARGO_TARGET_DIR=F:\RustBuild\pytools-feature"
) else if "%CURRENT_BRANCH%"=="test/business-coverage" (
    set "CARGO_TARGET_DIR=F:\RustBuild\pytools-test"
) else (
    echo Unsupported or detached Git branch: "%CURRENT_BRANCH%".
    echo Refusing to build into an unknown or shared target directory.
    pause
    exit /b 3
)

where cargo >nul 2>nul
if errorlevel 1 (
    echo cargo was not found. No build was started.
    pause
    exit /b 9009
)

pushd "%PROJECT_ROOT%"
if errorlevel 1 (
    echo Cannot enter the project directory: %PROJECT_ROOT%
    pause
    exit /b 4
)

if not exist "%CARGO_TARGET_DIR%" mkdir "%CARGO_TARGET_DIR%"
if not exist "%CARGO_TARGET_DIR%" (
    popd
    echo Cannot create Cargo target directory: %CARGO_TARGET_DIR%
    pause
    exit /b 5
)

echo Building branch %CURRENT_BRANCH% into %CARGO_TARGET_DIR%.
if "%CURRENT_BRANCH%"=="main" echo This is the main release candidate. Publish only after validation.
if not "%CURRENT_BRANCH%"=="main" echo This binary belongs to a development/test worktree, not an official release.
cargo build --release --offline -p pytools-core --bin pytools
set "BUILD_EXIT_CODE=%ERRORLEVEL%"
if not "%BUILD_EXIT_CODE%"=="0" goto build_failed

set "BUILT_EXE=%CARGO_TARGET_DIR%\release\pytools.exe"
set "STAGED_EXE=%PROJECT_ROOT%dist\pytools.exe.new"
if not exist "%BUILT_EXE%" goto artifact_missing
if not exist "%PROJECT_ROOT%dist" mkdir "%PROJECT_ROOT%dist"
if not exist "%PROJECT_ROOT%dist" goto dist_create_failed
copy /y "%BUILT_EXE%" "%STAGED_EXE%" >nul
if errorlevel 1 goto stage_copy_failed
move /y "%STAGED_EXE%" "%PROJECT_ROOT%dist\pytools.exe" >nul
if errorlevel 1 goto install_copy_failed

echo This worktree's EXE was updated: %PROJECT_ROOT%dist\pytools.exe
popd
pause
exit /b 0

:build_failed
popd
echo Release build failed with exit code %BUILD_EXIT_CODE%. Existing dist\pytools.exe was not touched.
pause
exit /b %BUILD_EXIT_CODE%

:artifact_missing
popd
echo Build succeeded but the expected artifact was not found: %BUILT_EXE%
pause
exit /b 6

:dist_create_failed
popd
echo Could not create the current worktree's dist directory.
pause
exit /b 7

:stage_copy_failed
set "COPY_EXIT_CODE=%ERRORLEVEL%"
popd
echo Could not stage the new EXE. Existing dist\pytools.exe was not touched.
pause
exit /b %COPY_EXIT_CODE%

:install_copy_failed
set "COPY_EXIT_CODE=%ERRORLEVEL%"
popd
echo Could not replace the current worktree's EXE. The previous EXE should remain available.
echo The staged file, if present, is: %STAGED_EXE%
pause
exit /b %COPY_EXIT_CODE%
