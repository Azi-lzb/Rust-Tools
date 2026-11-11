#!/usr/bin/env bash
set -euo pipefail

SDK_ROOT=/opt/aspose-sdk/Aspose.Cells
SOURCE_ROOT=/workspace/aspose-cpp
PACKAGE_ROOT=/package/aspose-cpp
EVIDENCE_ROOT=/workspace/reports/evidence/aspose-cpp

mkdir -p "$EVIDENCE_ROOT" "$PACKAGE_ROOT/lib"

for so in "$SDK_ROOT"/lib/linux_x86_64/*.so; do
    name=$(basename "$so")
    file "$so" > "$EVIDENCE_ROOT/$name.file.txt"
    readelf -h "$so" > "$EVIDENCE_ROOT/$name.elf-header.txt"
    readelf -d "$so" > "$EVIDENCE_ROOT/$name.dynamic.txt"
    readelf --version-info "$so" > "$EVIDENCE_ROOT/$name.symbol-versions.txt"
    ldd "$so" > "$EVIDENCE_ROOT/$name.ldd.txt" 2>&1 || true
done

cmake -S "$SOURCE_ROOT" -B /tmp/aspose-cpp-build -DASPOSE_SDK_ROOT="$SDK_ROOT"
cmake --build /tmp/aspose-cpp-build --verbose 2>&1 | tee "$EVIDENCE_ROOT/aspose-cpp-build.log"
cmake --install /tmp/aspose-cpp-build --prefix "$PACKAGE_ROOT"
cp -a "$SDK_ROOT"/lib/linux_x86_64/*.so "$PACKAGE_ROOT/lib/"

file "$PACKAGE_ROOT/bin/aspose_cpp_smoke" > "$EVIDENCE_ROOT/aspose-cpp-smoke.file.txt"
readelf -h "$PACKAGE_ROOT/bin/aspose_cpp_smoke" > "$EVIDENCE_ROOT/aspose-cpp-smoke.elf-header.txt"
readelf -d "$PACKAGE_ROOT/bin/aspose_cpp_smoke" > "$EVIDENCE_ROOT/aspose-cpp-smoke.dynamic.txt"
readelf --version-info "$PACKAGE_ROOT/bin/aspose_cpp_smoke" > "$EVIDENCE_ROOT/aspose-cpp-smoke.symbol-versions.txt"
ldd "$PACKAGE_ROOT/bin/aspose_cpp_smoke" > "$EVIDENCE_ROOT/aspose-cpp-smoke.ldd.txt"
ldd -r "$PACKAGE_ROOT/bin/aspose_cpp_smoke" > "$EVIDENCE_ROOT/aspose-cpp-smoke.ldd-reloc.txt" 2>&1
printf 'BUILD_COMPLETE\n'
