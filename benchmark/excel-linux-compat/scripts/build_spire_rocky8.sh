#!/usr/bin/env bash
set -euo pipefail

SPIRE_ROOT=/opt/spire-sdk
SOURCE_ROOT=/workspace/spire
PACKAGE_ROOT=/package/spire
EVIDENCE_ROOT=/workspace/reports/evidence/spire

mkdir -p "$EVIDENCE_ROOT" "$PACKAGE_ROOT/lib"

for so in "$SPIRE_ROOT"/lib/linux/*.so; do
    name=$(basename "$so")
    file "$so" > "$EVIDENCE_ROOT/$name.file.txt"
    readelf -h "$so" > "$EVIDENCE_ROOT/$name.elf-header.txt"
    readelf -d "$so" > "$EVIDENCE_ROOT/$name.dynamic.txt"
    readelf --version-info "$so" > "$EVIDENCE_ROOT/$name.symbol-versions.txt"
    ldd "$so" > "$EVIDENCE_ROOT/$name.ldd.txt" 2>&1 || true
done

cmake -S "$SOURCE_ROOT" -B /tmp/spire-build -DSPIRE_ROOT="$SPIRE_ROOT"
cmake --build /tmp/spire-build --verbose 2>&1 | tee "$EVIDENCE_ROOT/spire-build.log"
cmake --install /tmp/spire-build --prefix "$PACKAGE_ROOT"
cp -a "$SPIRE_ROOT"/lib/linux/*.so "$PACKAGE_ROOT/lib/"
cp -a /usr/lib64/libicu*.so* "$PACKAGE_ROOT/lib/"

file "$PACKAGE_ROOT/bin/spire_smoke" > "$EVIDENCE_ROOT/spire-smoke.file.txt"
readelf -h "$PACKAGE_ROOT/bin/spire_smoke" > "$EVIDENCE_ROOT/spire-smoke.elf-header.txt"
readelf -d "$PACKAGE_ROOT/bin/spire_smoke" > "$EVIDENCE_ROOT/spire-smoke.dynamic.txt"
readelf --version-info "$PACKAGE_ROOT/bin/spire_smoke" > "$EVIDENCE_ROOT/spire-smoke.symbol-versions.txt"
ldd "$PACKAGE_ROOT/bin/spire_smoke" > "$EVIDENCE_ROOT/spire-smoke.ldd.txt"
ldd -r "$PACKAGE_ROOT/bin/spire_smoke" > "$EVIDENCE_ROOT/spire-smoke.ldd-reloc.txt" 2>&1

printf 'BUILD_COMPLETE\n'
