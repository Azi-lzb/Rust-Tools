# Rust-Tools

Rust-Tools 是独立维护的 PyTools Rust 项目。正式源码目录为 `C:\Users\AZI\Desktop\VibeCode\Rust-Tools`；业务代码位于 `pytools-core/`，本项目独立维护的 Excel-Core 位于 `crates/`。七类菜单、业务范围和已知限制见 [功能清单](pytools-core/MIGRATION_INVENTORY.md)。

## Windows 启动与构建

- `run.bat` 只运行本项目 `dist\pytools.exe menu`；EXE 不存在时明确报错，不自动编译。
- `run-dev.bat` 从当前项目运行 Debug 菜单：调用 `cargo run --offline -p pytools-core --bin pytools -- menu`，复用 `F:\RustBuild\pytools` 增量缓存，并默认读取项目根目录的 `config.xlsx`。首次编译可能较慢；后续按源码变化增量重编译，不打 Release 包。
- `build-release.bat` 是用户明确要求构建时使用的唯一构建入口，Cargo 输出目录为 `F:\RustBuild\pytools`；成功后将 `pytools.exe` 复制到本项目 `dist\`。
- `package.bat` 使用已有的 `dist\` 产物、本地配置、启动脚本和 README 创建 Windows ZIP，不构建。
- 格式转换（XLS/ET→XLSX、DOC→DOCX）由 pytools 内置转换引擎直接完成：Windows 优先本机 Office/WPS COM（Excel→WPS 顺序），非 Windows 使用本机 LibreOffice `soffice`（可选）；没有独立 sidecar EXE。
- 项目根目录的 `config.xlsx` 是本地用户配置，已被 Git 忽略；不要将业务文件或本地配置提交到仓库。

日常开发和人工测试使用 `run-dev.bat`。修改 Rust 源码后，先退出正在运行的程序，再次双击即可按需增量编译并启动；修改配置无需重新编译，重新加载配置或重启程序即可。两个启动脚本在程序退出后都会显示退出码并等待按键，便于查看错误信息。

## openpyxl 运行时（就地修改类功能）

部分就地修改功能依赖 openpyxl 子进程。三种分发场景的运行时查找顺序一致（由
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
--dist dist` 暂存运行时，`package.bat` 在运行时缺失或握手失败时拒绝打包；
生成后可用 `python packaging/openpyxl/verify-windows-zip.py <zip>` 做解压
验证（握手、import、合成 ExcelEditPlan 端到端修改与回读）。Linux 打包：
`build_deb.sh` 需要 `OPENPYXL_RUNTIME_SOURCE` 指向 relocatable Python
运行时（锁定版本见 `packaging/openpyxl/python-version.txt`），缺失时中止。

后续以用户逐项人工测试和反馈为主，按功能修复；实施、测试和各平台验收状态分别记录。未经用户明确授权，不编译、运行 Rust 程序、打包或推送。详细规则见 [AGENTS.md](AGENTS.md)、[当前开发目标](DEVELOPMENT_GOAL.md) 和 [业务范围](pytools-core/MIGRATION_PLAN.md)。

历史拆分来源及早期说明保留在 [归档](docs/archive/README.md)，不作为运行路径或当前任务。Linux 交付目标仍为 x86_64、glibc 2.28 兼容和离线 deb，外部引擎是按功能明确选择的可选项。
