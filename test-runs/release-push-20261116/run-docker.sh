#!/usr/bin/env bash
set -eu
cd /src
label="${RUSTTOOLS_TEST_LABEL:-docker-workspace}"
mkdir -p /tmp/cargo /tmp/rusttools-state
ln -s /registry /tmp/cargo/registry
cp -a /runtime-src /tmp/runtime
# This overlay is private tmpfs; generated test reports cannot modify the host.
cp -a /report-source/. /src/pytools-core/probe/reports/
export CARGO_HOME=/tmp/cargo CARGO_TARGET_DIR=/build/docker-platform-validation
export CARGO_BUILD_JOBS=1 CARGO_PROFILE_TEST_DEBUG=0
export XDG_STATE_HOME=/tmp/rusttools-state
export PYTOOLS_OPENPYXL_PYTHON=/tmp/runtime/bin/python3
export RUST_WORKSPACE_OPENPYXL_PYTHON="$PYTOOLS_OPENPYXL_PYTHON"
export PYTOOLS_OPENPYXL_WORKER=/src/crates/excel-editor/python/openpyxl_worker.py
unset PYTOOLS_EXCEL_EDIT_ROUTE
"$PYTOOLS_OPENPYXL_PYTHON" -I -c 'import openpyxl; print(openpyxl.__version__)' > /evidence/docker-runtime.txt
{ cat /etc/os-release; cargo --version; ldd --version | head -1; } > /evidence/docker-environment.txt
date -u +%FT%TZ > /evidence/docker-start.txt
set +e
cargo test --offline --locked --workspace --no-fail-fast -- --test-threads=1 > "/evidence/$label.log" 2>&1
code=$?
printf '%s\n' "$code" > "/evidence/$label.exit"
date -u +%FT%TZ > /evidence/docker-end.txt
tail -20 "/evidence/$label.log"
exit "$code"
