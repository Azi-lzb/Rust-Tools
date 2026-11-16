#!/usr/bin/env bash
set -eu
cd /src
while [ ! -f /evidence/linux-deb-package.exit ]; do sleep 1; done
[ "$(tr -d '\r\n' < /evidence/linux-deb-package.exit)" = 0 ] || exit 1
while [ ! -f /evidence/linux-deb-metadata.txt ]; do sleep 1; done
mkdir -p /tmp/packaged-root /tmp/rusttools-state
tar -xf /evidence/linux-filesystem.tar -C /tmp/packaged-root
app=/tmp/packaged-root/opt/rusttools
export XDG_STATE_HOME=/tmp/rusttools-state
unset PYTOOLS_OPENPYXL_PYTHON PYTOOLS_OPENPYXL_WORKER PYTOOLS_COM_PYTHON PYTOOLS_COM_WORKER PYTOOLS_EXCEL_EDIT_ROUTE RUST_WORKSPACE_OPENPYXL_PYTHON
ldd "$app/RustTools" > /evidence/linux-packaged-ldd.txt
if grep -q 'not found' /evidence/linux-packaged-ldd.txt; then exit 1; fi
set +e
"$app/python/runtime/bin/python3" -X utf8 -I /src/test-runs/release-push-20261116/packaged_smoke.py "$app" /evidence/linux-docker-smoke /src > /evidence/linux-packaged-smoke.log 2>&1
code=$?
printf '%s\n' "$code" > /evidence/linux-packaged-smoke.exit
tail -12 /evidence/linux-packaged-smoke.log
exit "$code"
