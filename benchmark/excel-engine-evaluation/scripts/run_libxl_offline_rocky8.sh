#!/usr/bin/env bash
set -euo pipefail

PACKAGE_ROOT=/opt/libxl
OUTPUT_ROOT=/output
MODE=${1:-capabilities}
mkdir -p "$OUTPUT_ROOT"

printf '%s\n' '=== runtime environment ==='
cat /etc/os-release
uname -m
ldd --version | sed -n '1p'
getconf GNU_LIBC_VERSION
for tool in dotnet java python3 libreoffice soffice; do
    if command -v "$tool" >/dev/null 2>&1; then
        printf 'PRESENT %s %s\n' "$tool" "$(command -v "$tool")"
    else
        printf 'ABSENT %s\n' "$tool"
    fi
done

printf '%s\n' '=== app-private package dependencies ==='
ldd "$PACKAGE_ROOT/bin/libxl_smoke"
ldd -r "$PACKAGE_ROOT/bin/libxl_smoke"
if [[ "$MODE" == "performance" ]]; then
    ldd "$PACKAGE_ROOT/bin/libxl_perf"
    ldd -r "$PACKAGE_ROOT/bin/libxl_perf"
fi

if [[ "$MODE" == "performance" ]]; then
    printf '%s\n' '=== offline performance workload ==='
    "$PACKAGE_ROOT/bin/libxl_perf" /input/performance "$OUTPUT_ROOT/performance" "$OUTPUT_ROOT/performance-libxl.csv"
else
    printf '%s\n' '=== offline Excel API smoke ==='
    "$PACKAGE_ROOT/bin/libxl_smoke" "$OUTPUT_ROOT"
fi
