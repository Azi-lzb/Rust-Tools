#!/usr/bin/env bash
set -euo pipefail

cd /output
for tool in dotnet java python3 libreoffice soffice; do
    if command -v "$tool" >/dev/null 2>&1; then
        printf 'FAIL unexpected runtime tool installed: %s\n' "$tool" >&2
        exit 1
    fi
done
printf 'TOOLS dotnet=absent java=absent python3=absent libreoffice=absent soffice=absent\n'
ldd /opt/aspose-cpp/bin/aspose_cpp_smoke
/opt/aspose-cpp/bin/aspose_cpp_smoke
test -s /output/modified.xlsx
test -s /output/modified.xls
printf 'OFFLINE_SMOKE_COMPLETE Aspose.Cells.Cpp\n'
