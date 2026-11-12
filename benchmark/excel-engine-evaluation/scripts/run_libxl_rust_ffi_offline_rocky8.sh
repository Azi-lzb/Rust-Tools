#!/usr/bin/env bash
set -euo pipefail
for tool in dotnet java python3 libreoffice soffice; do
  if command -v "$tool" >/dev/null 2>&1; then echo "FAIL unexpected runtime: $tool" >&2; exit 1; fi
done
/target/release/libxl-rust-ffi-prototype /input/input.xlsx /output/libxl-ffi-output.xlsx 5
test -s /output/libxl-ffi-output.xlsx
