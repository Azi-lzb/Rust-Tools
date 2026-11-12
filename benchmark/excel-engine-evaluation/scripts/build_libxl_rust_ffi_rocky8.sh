#!/usr/bin/env bash
set -euo pipefail
SDK_ROOT=/opt/libxl-sdk
WORKSPACE=/workspace
PACKAGE=/package
mkdir -p "$PACKAGE"
g++ -std=c++17 -O2 -fPIC -shared \
  -I"$WORKSPACE/bridge/cpp/include" -I"$SDK_ROOT/include_cpp" \
  "$WORKSPACE/bridge/cpp/src/libxl_bridge.cpp" \
  -L"$SDK_ROOT/lib64" -Wl,-rpath,'$ORIGIN' -lxl \
  -o "$PACKAGE/libxlbridge.so"
cp "$SDK_ROOT/lib64/libxl.so" "$PACKAGE/libxl.so"
cargo build --release --offline --manifest-path "$WORKSPACE/bridge/cpp/rust-client/Cargo.toml"
mkdir -p /workspace/reports/evidence/bridge/libxl-ffi
file /target/release/libxl-rust-ffi-prototype > /workspace/reports/evidence/bridge/libxl-ffi/rust-binary-file.txt
readelf --version-info /target/release/libxl-rust-ffi-prototype > /workspace/reports/evidence/bridge/libxl-ffi/rust-binary-symbol-versions.txt
