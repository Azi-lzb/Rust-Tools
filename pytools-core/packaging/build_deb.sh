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
echo "== smoke test: openpyxl synthetic ExcelEditPlan via extracted binary =="
smoke_plan_dir="$(mktemp -d /tmp/pytools-deb-plan-XXXXXX)"
RUNTIME_PY="$smoke_dir/usr/libexec/pytools/openpyxl/runtime/bin/python3"
WORKER_PY="$smoke_dir/usr/libexec/pytools/openpyxl/worker/openpyxl_worker.py"
# 合成夹具 + 计划 + 能力记录，通过解压出的 pytools apply-edit-plan 执行。
"$RUNTIME_PY" -I - "$smoke_plan_dir" "$smoke_dir/usr/bin/pytools" <<'PYEOF'
import json, subprocess, sys
from pathlib import Path

import openpyxl

plan_dir = Path(sys.argv[1])
pytools = Path(sys.argv[2])
fixture = plan_dir / "fixture.xlsx"
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Data"
ws["A1"] = "original"
ws["B1"] = 7
wb.save(fixture)

# 与 Rust 包检查器对齐的部件清单特性（openpyxl 会写出空 <definedNames/>）。
import zipfile
with zipfile.ZipFile(fixture) as archive:
    workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")
features = ["relationships"]
if "<definedNames" in workbook_xml:
    features.append("defined_names")

output = plan_dir / "result.xlsx"
plan = {
    "request_id": "deb-smoke-1",
    "input": str(fixture),
    "output": str(output),
    "format": "xlsx",
    "workbook": {
        "format": "xlsx",
        "inventory_complete": True,
        "features_present": features,
    },
    "operations": [
        {"operation": "set_cell_value", "sheet": "Data", "cell": "A1",
         "value": {"type": "text", "value": "packaged-deb"}},
    ],
    "preserve": ["cell_values_and_types"],
    "drop_allowed": {},
}
capabilities = [{
    "engine": "openpyxl_worker",
    "status": "passed",
    "format": "xlsx",
    "operations": ["set_cell_value"],
    "supported_features": features,
    "preserved": ["cell_values_and_types"],
    "platform": "linux-x86_64",
    "evidence": {
        "sample_id": "deb-smoke",
        "report_path": "pytools-core/packaging/build_deb.sh",
        "adapter_version": "1",
        "engine_version": "openpyxl",
        "platform": "linux-x86_64",
    },
}]
(plan_dir / "plan.json").write_text(json.dumps(plan), encoding="utf-8")
(plan_dir / "capabilities.json").write_text(json.dumps(capabilities), encoding="utf-8")
result = subprocess.run(
    [str(pytools), "apply-edit-plan",
     "--plan", str(plan_dir / "plan.json"),
     "--capabilities", str(plan_dir / "capabilities.json"),
     "--platform", "linux-x86_64"],
    capture_output=True, text=True, timeout=300,
)
if result.returncode != 0:
    sys.stderr.write(result.stdout[-800:] + result.stderr[-800:])
    raise SystemExit("apply-edit-plan failed")
report = json.loads(result.stdout)
if report.get("engine") != "openpyxl_worker":
    raise SystemExit(f"unexpected engine: {report.get('engine')}")
verify = openpyxl.load_workbook(output)
value = verify["Data"]["A1"].value
if value != "packaged-deb":
    raise SystemExit(f"readback mismatch: {value!r}")
print("synthetic ExcelEditPlan via extracted binary: PASS")
PYEOF
if [ $? -ne 0 ]; then
    echo "FAIL: synthetic openpyxl edit smoke test" >&2
    exit 1
fi
echo "== install test: dpkg -i + runtime resolution from /usr/bin =="
if dpkg -i "$deb" >/dev/null 2>&1; then
    if command -v pytools >/dev/null 2>&1; then
        # worker 要求输出为全新文件；清掉解压路径冒烟生成的结果后重跑。
        rm -f "$smoke_plan_dir/result.xlsx"
        /usr/bin/pytools apply-edit-plan \
            --plan "$smoke_plan_dir/plan.json" \
            --capabilities "$smoke_plan_dir/capabilities.json" \
            --platform linux-x86_64 >/dev/null 2>&1 &&
            echo "installed-path openpyxl edit: PASS" ||
            echo "WARN: installed-path edit failed (see logs)" >&2
    else
        echo "WARN: pytools not on PATH after install" >&2
    fi
    dpkg -r pytools >/dev/null 2>&1 || true
else
    echo "WARN: dpkg -i failed in this WSL environment; extracted-tree smoke passed" >&2
fi
rm -rf "$smoke_plan_dir" "$smoke_dir" "$STAGING"
echo "== done: $deb =="
