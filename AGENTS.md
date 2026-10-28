# Rust-Tools 项目开发规则

## 唯一开发目录与协作

- 唯一正式开发目录是 `C:\Users\AZI\Desktop\VibeCode\Rust-Tools`。日常开发优先使用 Codex Local；未经用户明确要求，不主动创建 Worktree 或开发分支。
- 本仓库独立拥有业务代码、Cargo workspace、配置和 `crates/` 内的 Excel-Core。不得依赖旧 Main、其他项目或旧共享 workspace 的源码、同步、合并或 cherry-pick；可按当前业务需要独立修改本仓库的 Excel-Core，不为统一其他项目而覆盖现有实现。
- Python 原项目 `C:\Users\AZI\Desktop\VibeCode\4.pytools-main` 只读，`probe/` 和测试辅助中的 Python 仅供按需比较。隔离 openpyxl worker 仅是既有工作簿修改的受限兼容路径，不得承载 PyTools 业务逻辑。
- 保留根目录本地 `config.xlsx` 和业务文件；它们不应提交到 Git。不得改写原始 Python 项目、RustHistory 原始归档或业务原件；不得删除业务源码、配置、测试夹具及人工验收记录。针对明确任务可按需维护合成测试夹具，保留既有验收证据。

## 工作方式和业务安全

- 后续以用户逐项人工测试与明确反馈为主。用户没有指定功能时，不主动扩大迁移、审计、对拍或优化；出现问题时仅修复该功能及必要依赖。既有已迁移功能优先保留，不自动重写，不得未经测试删除旧功能或以未测试的替代品移除现有业务行为。
- 保留七类菜单、既定子功能和明确排除项，见 `DEVELOPMENT_GOAL.md`、`pytools-core/MIGRATION_PLAN.md` 和 `pytools-core/MIGRATION_INVENTORY.md`。历史文档和修改记录是参考，不是新的执行任务。
- 保护原始 Excel 文件、业务数据、公式、必要结构和明确约定保留的对象；优先在副本上操作。涉及覆盖原文件时，必须有明确授权及备份或快照保护。不因普通字体、边框、颜色等外观差异阻塞功能交付；用户明确要求的业务格式和重要结构仍应保护。
- Excel 原语放在本仓库 `crates/`，PyTools 业务逻辑放在 `pytools-core/`。按需遵循 `rules/13_Rust_Excel读写修改与引擎路由规范.md` 的当前安全约束；遇到冲突，以本轮用户明确要求和本文件为准。
- 新增或修改业务功能时遵守 `docs/RUN_LOGGING_STANDARD.md`：关联 run_id、时间和阶段，按实际业务记录输入与必要 Sheet 的结果、可定位错误及运行汇总，并编写必要日志测试；不得凭读取成功推断业务成功。
- 未经用户批准，不执行 Git 强制操作、push 或远端修改。不得主动清除 Cargo target 缓存。

## 编译、运行和交付

- 实现或代码审查不等于构建授权。只有当前用户请求明确授权相应操作时，才可编译、构建、执行或安装 Rust 代码；禁止未授权的 `cargo build/check/test/run/clippy/bench/install`、`rustc` 及通过脚本间接触发这些操作。运行已有 EXE 也须有相应的明确运行授权。
- 允许静态检查、Git diff 和 `cargo fmt`（含 `cargo fmt --check`）。未执行编译或测试必须明确记录为未验证，不能把写了测试当作通过。
- `run.bat` 和兼容别名 `run-pytools-dev.bat` 只运行 `%~dp0dist\pytools.exe`，缺少 EXE 时明确失败，不调用 Cargo。用户主动要求构建时，`build-release.bat` 使用 `CARGO_TARGET_DIR=F:\RustBuild\pytools` 并将成功产物复制到本项目 `dist`；`package.bat` 只打包已有 `dist` 产物，不构建。BAT 均从 `%~dp0` 定位本项目。
- 分别记录实现状态、单元/冒烟测试、Windows 用户或真实样本验收、Linux/deb 验证；不得把菜单入口、编译或合成样本视为人工验收通过。
- Linux 交付要求保持为 x86_64、glibc 2.28 兼容和离线 deb；核心路径不要求用户预装 Python、Go、Office、WPS、LibreOffice 或 JVM。外部引擎仅为明确、按能力选择的可选项。
