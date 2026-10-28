# Rust-Tools 项目开发规则

## Worktree 角色与协作

- 多 Agent / Worktree 的完整职责、启动检查、构建路径和集成规则见 `docs/WORKTREE_DEVELOPMENT.md`。本文件中的业务安全和授权约束仍然适用；规则更新不会自动传播到其他分支。
- 根目录 `AGENTS.md` 和 Worktree 协作规范由 Main Agent 统一维护。Feature/Test Agent 不自行修改协作规则；调整建议提交 Main 处理并通过正常 Git 同步。

| 工作目录 | 分支 | 主要用途 | `CARGO_TARGET_DIR` |
| --- | --- | --- | --- |
| `C:\Users\AZI\Desktop\VibeCode\Rust-Tools` | `main` | 正式集成、回归和 Release 发布 | `F:\RustBuild\pytools` |
| `C:\Users\AZI\Desktop\VibeCode\Rust-Tools-feature` | `feature/business-development` | 业务功能与定点修复 | `F:\RustBuild\pytools-feature` |
| `C:\Users\AZI\Desktop\VibeCode\Rust-Tools-test` | `test/business-coverage` | 测试、夹具和测试发现的定点修复 | `F:\RustBuild\pytools-test` |

- 本仓库的正式主仓库、集成目录和唯一正式发布分支分别是上表中的 Rust-Tools 主目录与 `main`。日常工作在当前任务指定的工作树内进行。
- 开始修改或运行 Cargo 前，先确认 `git rev-parse --show-toplevel` 和 `git branch --show-current` 与当前工作目录/任务一致；不要在一个工作树里改写另一个工作树的源码、配置或 `dist`。
- Feature 与 Test 都可以修改 Rust 业务代码、Excel-Core 和测试。若要共同修改同一公共模块，先明确各自文件/接口边界；各自先提交到本分支，再通过普通 Git merge 集成到 `main`。冲突应针对性解决，禁止用强制 reset、覆盖或丢弃修改来处理。
- 只有 `main` 可进行正式集成和发布；Feature/Test 的构建与压缩包必须清楚标记为测试产物。正式发布前在 `main` 运行所需回归并完成验收。
- 任何直接 Cargo 命令都必须先按上表为当前工作树设置 `CARGO_TARGET_DIR`，避免默认在 C 盘生成 `target` 或多个工作树共用缓存。`build-release.bat` 会按当前 Git 分支选择对应目录。
- 所有 BAT 必须以脚本自身目录 `%~dp0` 定位项目，只操作当前工作树，不引用旧 Main、其他项目或旧 Codex 工作树。
- 本仓库独立拥有业务代码、Cargo workspace、配置和 `crates/` 内的 Excel-Core。不得依赖旧共享 Main、其他项目或旧共享 workspace 的源码、同步、合并或 cherry-pick；可按当前业务需要独立修改本仓库的 Excel-Core，不为统一其他项目而覆盖现有实现。
- Python 原项目 `C:\Users\AZI\Desktop\VibeCode\4.pytools-main` 只读，`probe/` 和测试辅助中的 Python 仅供按需比较。隔离 openpyxl worker 仅是既有工作簿修改的受限兼容路径，不得承载 PyTools 业务逻辑。
- 保留根目录本地 `config.xlsx` 和业务文件；它们不应提交到 Git。不得改写原始 Python 项目、RustHistory 原始归档或业务原件；不得删除业务源码、配置、测试夹具及人工验收记录。针对明确任务可按需维护合成测试夹具，保留既有验收证据。

## 工作方式和业务安全

- 后续以用户逐项人工测试与明确反馈为主。用户没有指定功能时，不主动扩大迁移、审计、对拍或优化；出现问题时仅修复该功能及必要依赖。既有已迁移功能优先保留，不自动重写，不得未经测试删除旧功能或以未测试的替代品移除现有业务行为。
- 保留七类菜单、既定子功能和明确排除项，见 `DEVELOPMENT_GOAL.md`、`pytools-core/MIGRATION_PLAN.md` 和 `pytools-core/MIGRATION_INVENTORY.md`。历史文档和修改记录是参考，不是新的执行任务。
- 保护原始 Excel 文件、业务数据、公式、必要结构和明确约定保留的对象；优先在副本上操作。涉及覆盖原文件时，必须有明确授权及备份或快照保护。不因普通字体、边框、颜色等外观差异阻塞功能交付；用户明确要求的业务格式和重要结构仍应保护。
- Excel 原语放在本仓库 `crates/`，PyTools 业务逻辑放在 `pytools-core/`。按需遵循 `rules/13_Rust_Excel读写修改与引擎路由规范.md` 的当前安全约束；遇到冲突，以本轮用户明确要求和本文件为准。
- 新增或修改业务功能必须遵守 `docs/RUN_LOGGING_STANDARD.md`；它是运行日志的唯一权威规范。
- 新增批处理入口须关联 run_id，记录开始/结束时间、阶段、文件最终状态、必要的 Sheet/单元格定位、输出结果、用户可读错误、结构化错误及统一运行汇总，并为统计和状态判定补充测试；不得凭读取成功推断业务成功，也不得删除旧运行 JSON 的兼容字段。
- 未经用户批准，不执行 Git 强制操作、push 或远端修改。不得主动清除 Cargo target 缓存。

## 编译、运行和交付

- 实现或代码审查不等于构建授权。只有当前用户请求明确授权相应操作时，才可编译、构建、执行或安装 Rust 代码；禁止未授权的 `cargo build/check/test/run/clippy/bench/install`、`rustc` 及通过脚本间接触发这些操作。运行已有 EXE 也须有相应的明确运行授权。
- 允许静态检查、Git diff 和 `cargo fmt`（含 `cargo fmt --check`）。未执行编译或测试必须明确记录为未验证，不能把写了测试当作通过。
- `run.bat` 只运行当前工作树 `%~dp0dist\pytools.exe`，缺少 EXE 时明确失败，不调用 Cargo。用户主动要求构建时，`build-release.bat` 按当前分支映射到上表中独立的 F 盘 target，成功后只更新当前工作树的 `dist`。`package.bat` 只打包当前工作树已有 `dist`；Feature/Test 包标记为测试包，正式包只允许从 `main` 在明确确认验证通过后生成。BAT 均从 `%~dp0` 定位本项目。
- 分别记录实现状态、单元/冒烟测试、Windows 用户或真实样本验收、Linux/deb 验证；不得把菜单入口、编译或合成样本视为人工验收通过。
- Linux 交付要求保持为 x86_64、glibc 2.28 兼容和离线 deb；核心路径不要求用户预装 Python、Go、Office、WPS、LibreOffice 或 JVM。外部引擎仅为明确、按能力选择的可选项。
