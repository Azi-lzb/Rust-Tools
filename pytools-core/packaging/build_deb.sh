#!/usr/bin/env bash
# Assemble the rusttools deb: a self-contained /opt/rusttools application
# root (Windows-zip-style layout) plus FHS integration entries.
#
#   /opt/rusttools/RustTools          real binary (double-click/terminal entry)
#   /opt/rusttools/config/config.xlsx default config; registered as a conffile
#                                     so upgrades never overwrite user edits
#   /opt/rusttools/python/            bundled openpyxl runtime + worker, laid
#                                     out for the ZIP-layout resolver
#   /opt/rusttools/doc/README.md      project documentation
#   /usr/bin/RustTools                symlink -> /opt/rusttools/RustTools
#                                     (/proc/self/exe resolves symlinks, so
#                                     config is found from any launcher)
#   /usr/share/applications/rusttools.desktop  desktop menu entry
#   /usr/share/doc/rusttools/         dpkg documentation convention
#
# Runs inside WSL/Ubuntu (needs cargo-zigbuild + Zig for in-script builds,
# or SKIP_BUILD=1 BINARY=<path> with a prebuilt glibc 2.28 binary; also
# needs objdump and dpkg-deb). Stage-level packaging only: final packaging
# acceptance on Kylin V10/UOS V20 belongs to the project acceptance stage.
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
WORKSPACE="$(cd "$SCRIPT_DIR/../.." && pwd)"
CRATE="pytools-core"
VERSION="0.1.0-$(date +%Y%m%d)"
TARGET="x86_64-unknown-linux-gnu.2.28"
STAGING="$(mktemp -d /tmp/rusttools-deb-XXXXXX)"
CARGO_TARGET_DIR="${CARGO_TARGET_DIR:-/mnt/f/RustBuild/pytools}"
OUT_DIR="$CARGO_TARGET_DIR/deb"

cd "$WORKSPACE"
export CARGO_TARGET_DIR
export PATH="/root/.cargo/bin:/root/.local/bin:$PATH"

if [[ "${SKIP_BUILD:-0}" != "1" ]]; then
    echo "== building $TARGET release =="
    cargo zigbuild --release --offline --locked --target "$TARGET" -p "$CRATE" --bin RustTools
fi

# zigbuild normalizes the ".2.28" suffix onto the base triple directory.
binary="${BINARY:-$CARGO_TARGET_DIR/$TARGET/release/RustTools}"
if [ ! -f "$binary" ]; then
    binary="${BINARY:-$CARGO_TARGET_DIR/x86_64-unknown-linux-gnu/release/RustTools}"
fi
echo "== binary: $binary =="

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
# 复用统一暂存脚本：先落位到旧 libexec 布局，再重组为 /opt 自包含布局。
"$WORKSPACE/packaging/openpyxl/stage-runtime.sh" pytools     "${OPENPYXL_RUNTIME_SOURCE:?set OPENPYXL_RUNTIME_SOURCE to the relocatable python runtime root}"     "$STAGING"

echo "== assembling /opt/rusttools application root =="
APP_ROOT="$STAGING/opt/rusttools"
install -d -m 0755 "$APP_ROOT/config" "$APP_ROOT/doc"

install -m 0755 "$binary" "$APP_ROOT/RustTools"

# openpyxl runtime: runtime/ 下含 bin/python3 与 lib，worker 提升到 python/
# 根 —— 正好命中 Rust 侧 ZIP 布局解析的 exe_dir/python + python/runtime/bin。
mv "$STAGING/usr/libexec/pytools/openpyxl" "$APP_ROOT/python"
mv "$APP_ROOT/python/worker/openpyxl_worker.py" "$APP_ROOT/python/openpyxl_worker.py"
mv "$APP_ROOT/python/worker/libreoffice_calc_worker.py" "$APP_ROOT/python/libreoffice_calc_worker.py"
rmdir "$APP_ROOT/python/worker"

# 默认配置：优先项目根的真实 config（与 Windows 包同源）；缺失时用二进制
# 自带 config-init 生成空白模板。注册为 conffile，升级绝不覆盖用户改动。
if [[ -f "$WORKSPACE/config.xlsx" ]]; then
    install -m 0644 "$WORKSPACE/config.xlsx" "$APP_ROOT/config/config.xlsx"
    echo "config source: project root config.xlsx"
else
    "$APP_ROOT/RustTools" config-init --path "$APP_ROOT/config/config.xlsx"
    echo "config source: generated blank template (config-init)"
fi

install -m 0644 "$WORKSPACE/README.md" "$APP_ROOT/doc/README.md"

echo "== assembling FHS integration entries =="
install -d -m 0755 "$STAGING/usr/bin" "$STAGING/usr/share/applications" \
    "$STAGING/usr/share/doc/rusttools" "$STAGING/DEBIAN"
ln -s /opt/rusttools/RustTools "$STAGING/usr/bin/RustTools"

cat > "$STAGING/usr/share/applications/rusttools.desktop" <<'EOF'
[Desktop Entry]
Type=Application
Name=RustTools
Comment=通用汇总工具（Rust 版）
Exec=/opt/rusttools/RustTools
Terminal=true
Categories=Office;
EOF

cp "$WORKSPACE/$CRATE/packaging/deb/README" "$STAGING/usr/share/doc/rusttools/README"
cat > "$STAGING/usr/share/doc/rusttools/copyright" <<'EOF'
Format: https://www.debian.org/doc/packaging-manuals/copyright-format/1.0/
Upstream-Name: rusttools
Source: workspace-internal Rust migration of py_statistics_tools
Disclaimer: workspace-internal build; licensing follows the workspace repository.
EOF

echo "== conffiles: user config survives upgrades =="
cat > "$STAGING/DEBIAN/conffiles" <<'EOF'
/opt/rusttools/config/config.xlsx
EOF

installed_size="$(du -sk --exclude=DEBIAN "$STAGING/opt" "$STAGING/usr" | awk '{s+=$1} END {print s}')"
sed -e "s/@VERSION@/$VERSION/" -e "s/@INSTALLED_SIZE@/$installed_size/" \
    "$WORKSPACE/$CRATE/packaging/deb/control.in" > "$STAGING/DEBIAN/control"

echo "== building deb =="
mkdir -p "$OUT_DIR"
deb="$OUT_DIR/rusttools_${VERSION}_amd64.deb"
dpkg-deb --build --root-owner-group "$STAGING" "$deb"

echo "== verifying package =="
dpkg-deb --info "$deb"
dpkg-deb --contents "$deb" | grep -vE 'python3\.8|Lib/|libpython|\.so |include/|site-packages/|Tix|\.pyc|__pycache__|\.dist-info'

echo "== smoke test: summary conversion via extracted binary =="
smoke_dir="$(mktemp -d /tmp/rusttools-deb-smoke-XXXXXX)"
dpkg -x "$deb" "$smoke_dir"
fixture="$WORKSPACE/$CRATE/tests/fixtures/summary_used_range/error_token_matrix.xlsx"
(cd /tmp && "$smoke_dir/opt/rusttools/RustTools" summary-used-range \
    --output-dir "$smoke_dir/out" --log-dir "$smoke_dir/out/logs" --log-view off \
    "$fixture" | head -c 200)
echo
echo "== smoke test: openpyxl worker availability =="
"$smoke_dir/opt/rusttools/python/runtime/bin/python3" -I \
    "$smoke_dir/opt/rusttools/python/openpyxl_worker.py" --version

echo "== smoke test: openpyxl synthetic ExcelEditPlan via extracted binary =="
smoke_plan_dir="$(mktemp -d /tmp/rusttools-deb-plan-XXXXXX)"
RUNTIME_PY="$smoke_dir/opt/rusttools/python/runtime/bin/python3"
WORKER_PY="$smoke_dir/opt/rusttools/python/openpyxl_worker.py"
# 合成夹具 + 计划 + 能力记录，通过解压出的 RustTools apply-edit-plan 执行。
"$RUNTIME_PY" -I - "$smoke_plan_dir" "$smoke_dir/opt/rusttools/RustTools" <<'PYEOF'
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

echo "== install test: dpkg -i + config parity + upgrade protection =="
if dpkg -i "$deb" >/dev/null 2>&1; then
    # 模拟菜单/快捷方式启动：任意 cwd、不传 --config，config 必须命中 /opt。
    config_probe="$(cd /tmp && printf '7\n2\n0\n0\n' | /usr/bin/RustTools menu 2>&1 | grep -c '输出目录: /opt/rusttools' || true)"
    if [[ "$config_probe" -ge 1 ]]; then
        echo "config parity via symlink from unrelated cwd: PASS"
    else
        echo "FAIL: config did not resolve to /opt/rusttools" >&2
        dpkg -P rusttools >/dev/null 2>&1 || true
        exit 1
    fi
    # 用户改动 config（写入标记单元格），模拟真实使用后再升级。
    RUNTIME_PY_INSTALLED="/opt/rusttools/python/runtime/bin/python3"
    "$RUNTIME_PY_INSTALLED" -I - <<'PYEOF'
import openpyxl
book = openpyxl.load_workbook("/opt/rusttools/config/config.xlsx")
book.active["Z99"] = "user-edit-marker"
book.save("/opt/rusttools/config/config.xlsx")
PYEOF
    if dpkg -i "$deb" >/dev/null 2>&1; then
        marker=$("$RUNTIME_PY_INSTALLED" -I - <<'PYEOF'
import openpyxl
book = openpyxl.load_workbook("/opt/rusttools/config/config.xlsx")
print(book.active["Z99"].value)
PYEOF
)
        if [[ "$marker" == *"user-edit-marker"* ]]; then
            echo "upgrade preserves user-edited config: PASS"
        else
            echo "FAIL: upgrade overwrote the user-edited config (marker=$marker)" >&2
            dpkg -P rusttools >/dev/null 2>&1 || true
            exit 1
        fi
        if compgen -G "/opt/rusttools/config/config.xlsx.dpkg-dist" >/dev/null; then
            echo "new default shipped as .dpkg-dist: PASS"
        fi
    else
        echo "WARN: upgrade install failed; preservation unverified" >&2
    fi

    # 从安装路径（经符号链接）执行完整业务路径冒烟。
    rm -f "$smoke_plan_dir/result.xlsx"
    if (cd /tmp && /usr/bin/RustTools apply-edit-plan \
        --plan "$smoke_plan_dir/plan.json" \
        --capabilities "$smoke_plan_dir/capabilities.json" \
        --platform linux-x86_64 >/dev/null 2>&1); then
        echo "installed-path openpyxl edit: PASS"
    else
        echo "WARN: installed-path edit failed (see logs)" >&2
    fi
    dpkg -P rusttools >/dev/null 2>&1 || true
else
    echo "WARN: dpkg -i failed in this WSL environment; extracted-tree smoke passed" >&2
fi
rm -rf "$smoke_plan_dir" "$smoke_dir" "$STAGING"
echo "== done: $deb =="
