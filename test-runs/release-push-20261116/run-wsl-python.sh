#!/usr/bin/env bash
set -eu
cd /mnt/c/Users/AZI/Desktop/VibeCode/Rust-Tools
run=/mnt/f/RustBuild/pytools/release-push-20261116
stage="$(cat "$run/wsl-stage.txt")"
while [ ! -f "$run/wsl-runtime.txt" ]; do sleep 1; done
test -x "$stage/runtime/bin/python3"
# Keep this distro invocation alive until its full Rust run has finished.
while [ ! -f "$run/wsl-workspace.exit" ]; do sleep 1; done
export PYTOOLS_OPENPYXL_PYTHON="$stage/runtime/bin/python3"
export RUST_WORKSPACE_OPENPYXL_PYTHON="$PYTOOLS_OPENPYXL_PYTHON"
export PYTOOLS_OPENPYXL_WORKER="$PWD/crates/excel-editor/python/openpyxl_worker.py"
export RUSTTOOLS_TEST_EXE=/mnt/f/RustBuild/pytools/x86_64-unknown-linux-gnu/debug/RustTools
export RUSTTOOLS_BOUNDARY_CLI="$RUSTTOOLS_TEST_EXE"
export RUSTTOOLS_LOGGING_CLI="$RUSTTOOLS_TEST_EXE"
export RUSTTOOLS_BOUNDARY_CAPTURE_DIR="$run/wsl-python-cli"
export RUSTTOOLS_OFFICE_FIXTURE_DIR="$run/office-fixtures"
unset PYTOOLS_EXCEL_EDIT_ROUTE RUSTTOOLS_TEST_COM RUSTTOOLS_OFFICE_PROGID RUSTTOOLS_WPS_READONLY
set +e
"$PYTOOLS_OPENPYXL_PYTHON" -X utf8 -I -m unittest discover -s pytools-core/tests/python -p 'test_*.py' -v > "$run/wsl-python.log" 2>&1
code=$?
printf '%s\n' "$code" > "$run/wsl-python.exit"
tail -10 "$run/wsl-python.log"
"$PYTOOLS_OPENPYXL_PYTHON" -I "$PWD/test-runs/release-push-20261116/linux-openpyxl-gates.py" wsl
gate_code=$?
if [ "$code" -ne 0 ]; then exit "$code"; fi
exit "$gate_code"
