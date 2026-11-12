#!/usr/bin/env bash
set -euo pipefail

SDK_ROOT=/opt/libxl-sdk
SOURCE_ROOT=/workspace/libxl
PACKAGE_ROOT=/package/libxl
EVIDENCE_ROOT=/workspace/reports/evidence/libxl

command -v g++ >/dev/null
command -v cmake >/dev/null
command -v readelf >/dev/null
mkdir -p "$EVIDENCE_ROOT" "$PACKAGE_ROOT"

file "$SDK_ROOT/lib64/libxl.so" > "$EVIDENCE_ROOT/libxl-so-file.txt"
readelf -h "$SDK_ROOT/lib64/libxl.so" > "$EVIDENCE_ROOT/libxl-so-elf-header.txt"
readelf -d "$SDK_ROOT/lib64/libxl.so" > "$EVIDENCE_ROOT/libxl-so-dynamic.txt"
readelf --version-info "$SDK_ROOT/lib64/libxl.so" > "$EVIDENCE_ROOT/libxl-so-symbol-versions.txt"
ldd "$SDK_ROOT/lib64/libxl.so" > "$EVIDENCE_ROOT/libxl-so-ldd.txt"
ldd -r "$SDK_ROOT/lib64/libxl.so" > "$EVIDENCE_ROOT/libxl-so-ldd-r.txt"

cmake -S "$SOURCE_ROOT" -B /tmp/libxl-build -DLIBXL_ROOT="$SDK_ROOT"
cmake --build /tmp/libxl-build --verbose 2>&1 | tee "$EVIDENCE_ROOT/libxl-build.log"
cmake --install /tmp/libxl-build --prefix "$PACKAGE_ROOT"

file "$PACKAGE_ROOT/bin/libxl_smoke" > "$EVIDENCE_ROOT/libxl-smoke-file.txt"
readelf -h "$PACKAGE_ROOT/bin/libxl_smoke" > "$EVIDENCE_ROOT/libxl-smoke-elf-header.txt"
readelf -d "$PACKAGE_ROOT/bin/libxl_smoke" > "$EVIDENCE_ROOT/libxl-smoke-dynamic.txt"
readelf --version-info "$PACKAGE_ROOT/bin/libxl_smoke" > "$EVIDENCE_ROOT/libxl-smoke-symbol-versions.txt"
ldd "$PACKAGE_ROOT/bin/libxl_smoke" > "$EVIDENCE_ROOT/libxl-smoke-build-ldd.txt"
ldd -r "$PACKAGE_ROOT/bin/libxl_smoke" > "$EVIDENCE_ROOT/libxl-smoke-build-ldd-r.txt"

printf 'BUILD_COMPLETE\n'
