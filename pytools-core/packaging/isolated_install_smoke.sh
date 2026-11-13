#!/usr/bin/env bash
# NEVER invoke from ordinary packaging. Disposable container only, no host mounts.
set -euo pipefail
[[ "${RUSTTOOLS_DISPOSABLE_INSTALL_TEST:-}" == yes && -f /.dockerenv ]] || {
  echo "Requires an explicitly disposable Docker container without host filesystem mounts" >&2; exit 1;
}
# Reject a non-overlay root or mounts that could expose host application/package data.
awk '
$5 == "/" { for(i=1;i<=NF;i++) if($i=="-") { if($(i+1)=="overlay") overlay=1; } }
$5 ~ /^\/(opt|usr|var)(\/|$)/ || $5 == "/etc" || $5 ~ /^\/etc\/dpkg(\/|$)/ { unsafe_mount=1 }
END { exit (!overlay || unsafe_mount) }
' /proc/self/mountinfo || { echo "Unsafe mount layout; disposable overlay container required" >&2; exit 1; }
# Refuse an existing package or application/config tree, even in the container.
if dpkg-query -W rusttools >/dev/null 2>&1 || [[ -e /opt/rusttools ]]; then
  echo "Existing installation/config detected; refusing install test" >&2; exit 1
fi
[[ $# == 1 && -f "$1" ]] || { echo "usage: $0 package.deb" >&2; exit 1; }
deb="$(readlink -f "$1")"
trap 'dpkg -P rusttools >/dev/null 2>&1 || true' EXIT
dpkg -i "$deb"
/opt/rusttools/python/runtime/bin/python3 -I - <<'PY'
import openpyxl
p='/opt/rusttools/config/config.xlsx'
w=openpyxl.load_workbook(p); w.active['Z99']='synthetic-upgrade-marker'; w.save(p)
PY
dpkg -i "$deb"
/opt/rusttools/python/runtime/bin/python3 -I - <<'PY'
import openpyxl
assert openpyxl.load_workbook('/opt/rusttools/config/config.xlsx').active['Z99'].value == 'synthetic-upgrade-marker'
PY
