#!/usr/bin/env bash
set -euo pipefail

ENGINE=${1:?engine name required}
APP=${2:?apphost name required}
MODE=${3:-capabilities}
cd /output
for tool in dotnet java python3 libreoffice soffice; do
    if command -v "$tool" >/dev/null 2>&1; then
        printf 'FAIL unexpected runtime tool installed: %s\n' "$tool" >&2
        exit 1
    fi
done
printf 'TOOLS dotnet=absent java=absent python3=absent libreoffice=absent soffice=absent\n'
printf 'SELF_CONTAINED_APP %s\n' "$ENGINE"
export LD_LIBRARY_PATH="/opt/app:${LD_LIBRARY_PATH:-}"
if [[ "$MODE" == "performance" ]]; then
    "/opt/app/$APP" --perf /input/performance /output/performance "/output/performance-$ENGINE.csv"
    test -s "/output/performance-$ENGINE.csv"
else
    "/opt/app/$APP"
    test -s "/output/${ENGINE}-capabilities.xlsx"
    test -s "/output/${ENGINE}-converted.xlsx"
fi
printf 'OFFLINE_SELF_CONTAINED_SMOKE_COMPLETE %s\n' "$ENGINE"
