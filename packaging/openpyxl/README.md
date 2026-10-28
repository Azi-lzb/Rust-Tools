# Isolated openpyxl runtime staging

`stage-runtime.sh` prepares the openpyxl worker subtree inside an existing deb staging root. It does not build a Python runtime, install packages, create a deb, or modify product route capabilities.

The supplied relocatable runtime must already contain `bin/python3`, openpyxl, and et-xmlfile. The pinned versions are Python 3.8.18 (`python-version.txt`), openpyxl 3.1.5, and et-xmlfile 2.0.0 (`requirements.txt`). The staging script rejects a different Python version, then copies it to:

```text
<deb-root>/usr/libexec/<product>/openpyxl/runtime/
<deb-root>/<deb-root>/usr/libexec/<product>/openpyxl/worker/openpyxl_worker.py
```

This matches `OpenpyxlWorkerConfig::for_product_binary` for `/usr/bin/<product>`. The runtime must resolve its interpreter and internal symlinks within the staged runtime tree. Staging checks `python -I`, imports openpyxl, starts the worker help path, verifies every runtime symlink stays inside the staged tree, and rejects unresolved dynamic dependencies found by `ldd` (whitelisting only glibc base libraries) plus any GLIBC symbol requirement above the 2.28 delivery ceiling.

Shell scripts are forced to LF via `.gitattributes` (`*.sh text eol=lf`); checkouts that still carry CRLF break the WSL shebang invocation, so normalize before use (`sed -i 's/\r$//'`).

Example:

```sh
bash packaging/openpyxl/stage-runtime.sh pytools /opt/build/python-runtime deb-staging/pytools
```

## 可复现的 Linux 运行时构建（2026-10-29 实测）

仓库不提交运行时产物。构建/验证环境（WSL2 Ubuntu，无网络直连 GitHub 时可在
Windows 宿主下载后由 WSL 使用）：

1. 下载 python-build-standalone（astral-sh，可信来源）CPython 3.8.18
   install_only tarball（`x86_64-unknown-linux-gnu`，目标 glibc 2.17，低于
   2.28 交付上限），解压为 `<runtime>/bin/python3` 等布局；
2. 用宿主 Python `pip download --only-binary=:all: --no-deps` 取
   openpyxl==3.1.5 与 et-xmlfile==2.0.0 wheel（纯 Python，跨平台相同），
   解压进 `<runtime>/lib/python3.8/site-packages/`；
3. `python-version.txt` 已更新为 3.8.18（原 3.8.10 无可信可下载构建）；
4. 验证 `bin/python3 -I -c "import openpyxl, et_xmlfile"`；
5. `build_deb.sh`（WSL，需 cargo-zigbuild + zig 与 `dpkg-deb`）以
   `OPENPYXL_RUNTIME_SOURCE=<runtime>` 暂存并出包，包含 GLIBC 2.28 符号
   上限检查、ldd 白名单检查、worker `--version` 冒烟。

2026-10-29 实际执行：WSL2 Ubuntu 24.04 内完成 1–5（zig 0.13.0 +
cargo-zigbuild），deb 产出见该次构建输出；产物未提交仓库。

## Windows

`stage-windows-runtime.py --dist dist`：下载 CPython embeddable（版本见
`python-version-windows.txt`，默认 3.12.10）+ pip download 固定 wheel
（openpyxl 3.1.5 / et-xmlfile 2.0.0），组装到 `dist/python/` 并自检
（`python -I -c "import openpyxl"` + worker 启动）。每次暂存都会清理旧
`dist/python/`，避免误带过期 worker。`package.bat` 在运行时缺失或
`--version` 握手失败时拒绝打包。

ZIP 生成后运行解压验证：

```sh
python packaging/openpyxl/verify-windows-zip.py <zip>
# 省略路径时自动选择 dist 下最新的 pytools-windows-x64-*.zip
```

验证内容：解压布局完整性、bundled worker `--version` 握手（protocol=1）、
import openpyxl/et_xmlfile、用解压出的 pytools.exe `apply-edit-plan` 执行
合成 ExcelEditPlan（openpyxl 引擎路由 + PlanOutputValidator 回读校验）、
输出回读与输入快照。2026-10-29 实际执行：全部通过。

## 验收状态区分

- "暂存/打包脚本自检通过"指脚本内的 import 与 `--version` 冒烟；
- "解压/安装后验证"指 verify-windows-zip.py 或 build_deb.sh 的 smoke 测试；
- 目标系统（麒麟 V10 / 统信 UOS V20、干净 Windows）人工验收仍需单独执行，
  本仓库的自动化验证不能替代。
