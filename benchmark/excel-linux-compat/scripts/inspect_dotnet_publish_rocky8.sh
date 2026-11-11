#!/usr/bin/env bash
set -euo pipefail
engine=${1:?engine name required: aspose or gembox}
case "$engine" in
  aspose) app=Asposetest ;;
  gembox) app=GemboxTest ;;
  *) echo "unsupported engine: $engine" >&2; exit 2 ;;
esac
uname -m
getconf GNU_LIBC_VERSION
ldd --version | sed -n '1p'
echo '=== self-contained apphost ==='
file "/app/$app"
readelf -h "/app/$app"
ldd -r "/app/$app"
echo '=== self-contained native libraries ==='
for f in /app/*.so; do
  [ -e "$f" ] || continue
  echo "--- $f"
  file "$f"
  readelf --version-info "$f" | grep 'Name: GLIBC_' | sort -Vu | tail -2 || true
  ldd -r "$f"
done
echo '=== publish size ==='
du -sh /app
