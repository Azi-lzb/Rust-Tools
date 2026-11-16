#!/usr/bin/env bash
set -eu
cd /mnt/c/Users/AZI/Desktop/VibeCode/Rust-Tools
run=/mnt/f/RustBuild/pytools/release-push-20261116
while [ ! -f "$run/linux-release-build.exit" ]; do sleep 1; done
[ "$(tr -d '\r\n\357\273\277' < "$run/linux-release-build.exit")" = 0 ] || exit 1
export CARGO_TARGET_DIR=/mnt/f/RustBuild/pytools
export SKIP_BUILD=1 PACKAGE_KIND=test DPKG_DEB_THREADS_MAX=1
export BINARY=/mnt/f/RustBuild/pytools/docker-platform-validation/release/RustTools
export OPENPYXL_RUNTIME_SOURCE=/mnt/f/RustBuild/pytools/openpyxl-runtime-src/runtime
export XDG_STATE_HOME="$(mktemp -d /tmp/rusttools-package-state-XXXXXX)"
deb=/mnt/f/RustBuild/pytools/deb/rusttools_0.1.0-20261116-test-main_amd64.deb
[ ! -e "$deb" ] || { echo 'Refusing to overwrite an existing DEB'; exit 1; }
set +e
bash pytools-core/packaging/build_deb.sh > "$run/linux-deb-package.log" 2>&1
code=$?
printf '%s\n' "$code" > "$run/linux-deb-package.exit"
if [ "$code" -eq 0 ]; then
  dpkg-deb --fsys-tarfile "$deb" > "$run/linux-filesystem.tar"
  dpkg-deb -f "$deb" Package Version Architecture > "$run/linux-deb-metadata.txt"
  sha256sum "$deb" > "$run/linux-deb.sha256"
fi
tail -15 "$run/linux-deb-package.log"
exit "$code"
