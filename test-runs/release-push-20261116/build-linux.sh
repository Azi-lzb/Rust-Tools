#!/usr/bin/env bash
set -eu
cd /src
mkdir -p /tmp/cargo
ln -s /registry /tmp/cargo/registry
export CARGO_HOME=/tmp/cargo CARGO_TARGET_DIR=/build/docker-platform-validation CARGO_BUILD_JOBS=1
export XDG_STATE_HOME=/tmp/rusttools-build-state
cargo --version > /evidence/linux-build-environment.txt
{ cat /etc/os-release; ldd --version | head -1; } >> /evidence/linux-build-environment.txt
set +e
cargo build --release --offline --locked -p pytools-core --bin RustTools > /evidence/linux-release-build.log 2>&1
code=$?
printf '%s\n' "$code" > /evidence/linux-release-build.exit
if [ "$code" -eq 0 ]; then
  objdump -T /build/docker-platform-validation/release/RustTools > /evidence/linux-symbols.txt
  sha256sum /build/docker-platform-validation/release/RustTools > /evidence/linux-binary.sha256
fi
tail -12 /evidence/linux-release-build.log
exit "$code"
