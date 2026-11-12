#!/usr/bin/env bash
set -euo pipefail

cd /output
printf 'TOOLS dotnet=%s java=%s python3=%s libreoffice=%s soffice=%s\n' \
    "$(command -v dotnet || true)" "$(command -v java || true)" \
    "$(command -v python3 || true)" "$(command -v libreoffice || true)" \
    "$(command -v soffice || true)"
ldd /opt/spire/bin/spire_smoke
export LD_LIBRARY_PATH="/opt/spire/lib:${LD_LIBRARY_PATH:-}"
/opt/spire/bin/spire_smoke
test -s /output/modified.xlsx
test -s /output/modified.xls
printf 'OFFLINE_SMOKE_COMPLETE\n'
