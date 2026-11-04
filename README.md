# Rust-Tools

Rust-Tools 是独立维护的 PyTools Rust 项目。正式源码目录为 `C:\Users\AZI\Desktop\VibeCode\Rust-Tools`；业务代码位于 `pytools-core/`，本项目独立维护的 Excel-Core 位于 `crates/`。七类菜单、业务范围和已知限制见 [功能清单](pytools-core/MIGRATION_INVENTORY.md)。

## Windows 启动与构建

- 直接双击 `dist\RustTools.exe` 即可运行交互菜单；程序自动定位配置（依次探测 exe 目录、上级目录与当前目录的 `config.xlsx`），无需任何启动脚本。
- `run-dev.bat` 从当前项目运行 Debug 菜单：调用 `cargo run --offline -p pytools-core --bin pytools -- menu`，复用 `F:\RustBuild\pytools` 增量缓存，并默认读取项目根目录的 `config.xlsx`。首次编译可能较慢；后续按源码变化增量重编译，不打 Release 包。
- Release 构建入口：`cargo build --release --offline -p pytools-core --bin RustTools`（`CARGO_TARGET_DIR=F:\RustBuild\pytools`）；成功后将 `RustTools.exe` 复制到本项目 `dist\`。
- Windows 包：先 `python packaging/openpyxl/stage-windows-runtime.py --dist dist` 暂存运行时，再
  `powershell -File packaging/windows-package.ps1` 组装。包内布局（解压后在根目录双击 exe 即用）：
  `RustTools.exe` + `config/config.xlsx` + `python/`（openpyxl 运行时）+ `doc/README.md`。
- 格式转换（XLS/ET/ETT→XLSX、DOC/WPS/WPT→DOCX）由 pytools 内置转换引擎完成：Windows 全部 Office 格式统一使用随包 CPython + pywin32 调用本机 Excel/WPS（`.et/.ett` 走 WPS 表格、`.wps/.wpt` 走 WPS 文字，并核验 WPS 原生文件格式），在批次内复用同一 COM 会话；不再保留 Rust 直连 COM 通道。非 Windows 使用本机 LibreOffice `soffice`（可选，仅支持 `.xls/.xlsb/.xlsm→.xlsx`、`.doc→.docx`）。没有独立 sidecar EXE，也不依赖系统 Python。
- 项目根目录的 `config.xlsx` 是本地用户配置，已被 Git 忽略；不要将业务文件或本地配置提交到仓库。

日常开发和人工测试使用 `run-dev.bat`。修改 Rust 源码后，先退出正在运行的程序，再次双击即可按需增量编译并启动；修改配置无需重新编译，重新加载配置或重启程序即可。两个启动脚本在程序退出后都会显示退出码并等待按键，便于查看错误信息。

## Windows Python 运行时（COM 转换与 openpyxl 就地修改）

Windows ZIP 随包提供 CPython 3.12.10、pywin32 312、openpyxl 3.1.5 和
et-xmlfile 2.0.0。Windows Excel 格式转换 Worker 使用本机已安装的 Excel 或 WPS
建立批次 COM 会话，并在应用重新打开候选文件后对比工作表顺序、公式和非公式单元格值；
用户无需另行安装 Python。

部分就地修改功能依赖 openpyxl 子进程。三种分发场景的 openpyxl 运行时查找顺序一致（由
`OpenpyxlWorkerConfig::resolve_runtime` 统一实现，预检与执行共用同一入口，
预检还会实际启动运行时完成协议握手）：

1. 显式覆盖：`PYTOOLS_OPENPYXL_PYTHON` + `PYTOOLS_OPENPYXL_WORKER`（两者必须
   同时设置且路径有效；只设一个或路径无效会得到明确的配置错误，不会静默回退）；
2. Linux deb：`/usr/libexec/pytools/openpyxl/{runtime/bin/python3, worker/openpyxl_worker.py}`；
3. Windows ZIP：`<程序目录>\python\{python.exe, openpyxl_worker.py}`（兼容
   `python/runtime/bin/` 变体）；
4. 开发树（仅 debug 构建，或设置 `PYTOOLS_OPENPYXL_DEV_TREE=1` 显式允许）：
   worker 使用仓库内 `crates/excel-editor/python/openpyxl_worker.py`，解释器
   仍必须由 `PYTOOLS_OPENPYXL_PYTHON` 显式提供。

PATH 上的系统 Python 不构成受支持分发。运行时缺失时，功能会明确报错并列出
已查找位置，不会回退到系统 Python。

Worker 与 Rust 主程序通过协议版本握手（`protocol=1`）在派发前确认兼容；
握手成功结果按「解释器路径 + worker 路径 + 双方文件版本标识」缓存，worker
或运行时被替换后自动失效。worker 超时或失控时，Rust 侧通过 Windows Job
Object（子进程挂起建组后入 Job）或 Unix 进程组终止整棵受控进程树并回收，
不影响主程序或其他 Python 进程。worker 的 stdout/stderr 各有 64 MiB 单次
上限，超限即终止进程树；IPC 响应文件同样有上限，非 UTF-8 输出按 lossy
处理；失败信息包含类别、中文原因、修复建议与诊断行（run_id、协议版本、
退出码、耗时），不直接倾倒完整 traceback。

`.xlsm` 就地修改不在当前迁移范围：包检查器在计划构建阶段即拒绝 `.xlsm`
（仅接受 `.xlsx`），worker 预检同样只接受 `.xlsx`，相关请求在写入前被
明确拒绝；5-6 提取功能可读取 `.xlsm` 并输出新的 `.xlsx`。

Windows 打包：先执行 `python packaging/openpyxl/stage-windows-runtime.py
--dist dist` 暂存运行时，打包校验在 Python、pywin32 或 Worker 握手失败时拒绝打包；
生成后可用 `python packaging/openpyxl/verify-windows-zip.py <zip>` 做解压
验证（pywin32/openpyxl 握手与 import、合成 ExcelEditPlan 端到端修改与回读）。Linux 打包：
`build_deb.sh` 需要 `OPENPYXL_RUNTIME_SOURCE` 指向 relocatable Python
运行时（锁定版本见 `packaging/openpyxl/python-version.txt`），缺失时中止。

后续以用户逐项人工测试和反馈为主，按功能修复；实施、测试和各平台验收状态分别记录。未经用户明确授权，不编译、运行 Rust 程序、打包或推送。详细规则见 [AGENTS.md](AGENTS.md)、[当前开发目标](DEVELOPMENT_GOAL.md) 和 [业务范围](pytools-core/MIGRATION_PLAN.md)。

历史拆分来源及早期说明保留在 [归档](docs/archive/README.md)，不作为运行路径或当前任务。Linux 交付目标仍为 x86_64、glibc 2.28 兼容和离线 deb，外部引擎是按功能明确选择的可选项。
