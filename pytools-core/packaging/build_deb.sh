#!/usr/bin/env bash
# Assemble the pytools deb from the glibc 2.28 release binary.
# Runs inside WSL/Ubuntu (needs cargo-zigbuild, Zig, objdump, dpkg-deb).
# Stage-level packaging only: final packaging acceptance on Kylin V10/UOS V20
# belongs to the independent PyTools project acceptance stage.
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
WORKSPACE="$(cd "$SCRIPT_DIR/../.." && pwd)"
CRATE="pytools-core"
VERSION="0.1.0-$(date +%Y%m%d)"
TARGET="x86_64-unknown-linux-gnu.2.28"
STAGING="$(mktemp -d /tmp/pytools-deb-XXXXXX)"
CARGO_TARGET_DIR="${CARGO_TARGET_DIR:-/mnt/f/RustBuild/pytools}"
OUT_DIR="$CARGO_TARGET_DIR/deb"

cd "$WORKSPACE"
export CARGO_TARGET_DIR
export PATH="/root/.cargo/bin:/root/.local/bin:$PATH"

echo "== building $TARGET release =="
cargo zigbuild --release --offline --locked --target "$TARGET" -p "$CRATE" --bin pytools

binary="$CARGO_TARGET_DIR/x86_64-unknown-linux-gnu/release/pytools"
echo "== checking GLIBC symbol ceiling =="
max_glibc="$(objdump -T "$binary" \
    | grep -o 'GLIBC_[0-9.]*' | sed 's/GLIBC_//' \
    | sort -t. -k1,1n -k2,2n -k3,3n | tail -1)"
echo "max GLIBC symbol: $max_glibc"
if [[ "$max_glibc" != "2.28" ]]; then
    echo "FAIL: max GLIBC symbol $max_glibc exceeds the 2.28 delivery ceiling" >&2
    exit 1
fi

echo "== staging openpyxl runtime (required; failure aborts the deb) =="
# 复用统一暂存脚本：把 relocatable Python + openpyxl + worker 装进 deb。
# 运行时产物未提交到仓库时（README 所述），此处会明确失败，不产出缺 worker 的包。
"$WORKSPACE/packaging/openpyxl/stage-runtime.sh" pytools     "${OPENPYXL_RUNTIME_SOURCE:?set OPENPYXL_RUNTIME_SOURCE to the relocatable python runtime root}"     "$STAGING"

echo "== assembling package tree =="
install -D -m 0755 "$binary" "$STAGING/usr/bin/pytools"
install -d -m 0755 "$STAGING/usr/share/doc/pytools" "$STAGING/DEBIAN"
cat > "$STAGING/usr/share/doc/pytools/copyright" <<'EOF'
Format: https://www.debian.org/doc/packaging-manuals/copyright-format/1.0/
Upstream-Name: pytools
Source: workspace-internal Rust migration of py_statistics_tools
Disclaimer: workspace-internal build; licensing follows the workspace repository.
EOF
cp "$WORKSPACE/$CRATE/packaging/deb/README" "$STAGING/usr/share/doc/pytools/README"

installed_size="$(du -sk --exclude=DEBIAN "$STAGING/usr" | cut -f1)"
sed -e "s/@VERSION@/$VERSION/" -e "s/@INSTALLED_SIZE@/$installed_size/" \
    "$WORKSPACE/$CRATE/packaging/deb/control.in" > "$STAGING/DEBIAN/control"

echo "== building deb =="
mkdir -p "$OUT_DIR"
deb="$OUT_DIR/pytools_${VERSION}_amd64.deb"
dpkg-deb --build --root-owner-group "$STAGING" "$deb"

echo "== verifying package =="
dpkg-deb --info "$deb"
dpkg-deb --contents "$deb"

echo "== smoke test (extracted binary, no dpkg install) =="
smoke_dir="$(mktemp -d /tmp/pytools-deb-smoke-XXXXXX)"
dpkg -x "$deb" "$smoke_dir"
fixture="$WORKSPACE/$CRATE/tests/fixtures/summary_used_range/error_token_matrix.xlsx"
"$smoke_dir/usr/bin/pytools" summary-used-range \
    --output-dir "$smoke_dir/out" --log-dir "$smoke_dir/out/logs" --log-view off \
    "$fixture" | head -c 400
echo
echo "== smoke test: openpyxl worker availability =="
"$smoke_dir/usr/libexec/pytools/openpyxl/runtime/bin/python3" -I     "$smoke_dir/usr/libexec/pytools/openpyxl/worker/openpyxl_worker.py" --version
rm -rf "$smoke_dir" "$STAGING"
echo "== done: $deb =="
