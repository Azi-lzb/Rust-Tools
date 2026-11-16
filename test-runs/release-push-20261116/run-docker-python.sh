#!/usr/bin/env bash
set -eu
cd /src
mkdir -p /tmp/rusttools-state
cp -a /runtime-src /tmp/runtime
export XDG_STATE_HOME=/tmp/rusttools-state
export PYTOOLS_OPENPYXL_PYTHON=/tmp/runtime/bin/python3
export RUST_WORKSPACE_OPENPYXL_PYTHON="$PYTOOLS_OPENPYXL_PYTHON"
export PYTOOLS_OPENPYXL_WORKER=/src/crates/excel-editor/python/openpyxl_worker.py
export RUSTTOOLS_TEST_EXE=/build/docker-platform-validation/debug/RustTools
export RUSTTOOLS_BOUNDARY_CLI="$RUSTTOOLS_TEST_EXE"
export RUSTTOOLS_LOGGING_CLI="$RUSTTOOLS_TEST_EXE"
export RUSTTOOLS_BOUNDARY_CAPTURE_DIR=/evidence/docker-python-cli
export RUSTTOOLS_OFFICE_FIXTURE_DIR=/evidence/office-fixtures
unset PYTOOLS_EXCEL_EDIT_ROUTE RUSTTOOLS_TEST_COM RUSTTOOLS_OFFICE_PROGID RUSTTOOLS_WPS_READONLY
while [ ! -f /evidence/docker-workspace.exit ]; do sleep 1; done
set +e
"$PYTOOLS_OPENPYXL_PYTHON" -X utf8 -I -m unittest discover -s pytools-core/tests/python -p 'test_*.py' -v > /evidence/docker-python.log 2>&1
code=$?
printf '%s\n' "$code" > /evidence/docker-python.exit
tail -10 /evidence/docker-python.log
"$PYTOOLS_OPENPYXL_PYTHON" -I /src/test-runs/release-push-20261116/linux-openpyxl-gates.py docker
gate_code=$?
if [ "$code" -ne 0 ]; then exit "$code"; fi
exit "$gate_code"
