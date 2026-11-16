#!/usr/bin/env bash
set -eu
export PATH=/root/.cargo/bin:$PATH
cd /mnt/c/Users/AZI/Desktop/VibeCode/Rust-Tools
run=/mnt/f/RustBuild/pytools/release-push-20261116
label="${RUSTTOOLS_TEST_LABEL:-wsl-workspace}"
stage="${RUSTTOOLS_TEST_STAGE:-}"
if [ -z "$stage" ]; then
    stage="$(mktemp -d /tmp/rusttools-regression-20261116-XXXXXX)"
    cp -a /mnt/f/RustBuild/pytools/openpyxl-runtime-src/runtime "$stage/runtime"
fi
printf '%s\n' "$stage" > "$run/wsl-stage.txt"
export CARGO_TARGET_DIR=/mnt/f/RustBuild/pytools
# Fixtures use native Linux filesystem semantics, not Windows DrvFS rename flags.
export TMPDIR=/tmp
export CARGO_BUILD_JOBS=1 CARGO_PROFILE_TEST_DEBUG=0
export PYTOOLS_OPENPYXL_PYTHON="$stage/runtime/bin/python3"
export RUST_WORKSPACE_OPENPYXL_PYTHON="$PYTOOLS_OPENPYXL_PYTHON"
export PYTOOLS_OPENPYXL_WORKER="$PWD/crates/excel-editor/python/openpyxl_worker.py"
export PYTOOLS_LO_PYTHON=/usr/bin/python3
export PYTOOLS_LO_CALC_WORKER="$PWD/pytools-core/python/libreoffice_calc_worker.py"
unset PYTOOLS_EXCEL_EDIT_ROUTE
"$PYTOOLS_OPENPYXL_PYTHON" -I -c 'import openpyxl; print(openpyxl.__version__)' > "$run/wsl-runtime.txt"
{ uname -sr; cargo --version; ldd --version | head -1; command -v soffice || true; } > "$run/wsl-environment.txt"
date -u +%FT%TZ > "$run/wsl-start.txt"
set +e
cargo test --offline --workspace --target x86_64-unknown-linux-gnu --no-fail-fast -- --test-threads=1 > "$run/$label.log" 2>&1
code=$?
printf '%s\n' "$code" > "$run/$label.exit"
date -u +%FT%TZ > "$run/wsl-end.txt"
tail -20 "$run/$label.log"
exit "$code"
