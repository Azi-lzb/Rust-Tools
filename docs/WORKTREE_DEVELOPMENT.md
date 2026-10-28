# Rust-Tools 多 Agent / Git Worktree 开发规范

本文是 Rust-Tools 多工作树协作的唯一详细规范。根目录 `AGENTS.md` 保留通用业务安全和授权规则，并引用本文；各分支不得维护互相冲突的 Worktree 规则。Git 分支和 Worktree 是独立工作副本，提交或修改不会自动传播到其他分支。

## 1. 工作树职责与路径

| 工作树 | 分支 | 主要职责 | Cargo target |
| --- | --- | --- | --- |
| `C:\Users\AZI\Desktop\VibeCode\Rust-Tools` | `main` | 集成 Feature/Test、解决集成问题、统一回归、Windows Release、正式打包交付 | `F:\RustBuild\pytools` |
| `C:\Users\AZI\Desktop\VibeCode\Rust-Tools-feature` | `feature/business-development` | Rust 业务开发、Umya/OOXML/Excel-Core、openpyxl Worker、业务 Bug 与对应测试 | `F:\RustBuild\pytools-feature` |
| `C:\Users\AZI\Desktop\VibeCode\Rust-Tools-test` | `test/business-coverage` | 测试数据和夹具、七类菜单自动化、集成/回归测试、测试发现的 Bug 与必要的 Excel-Core 修改 | `F:\RustBuild\pytools-test` |

Main 是正式集成和唯一正式发布分支，不承担日常业务开发。Feature 和 Test 的职责用于降低冲突，不限制代码写入范围：两者都可以修改 Rust 源码、业务模块、Excel-Core 和测试。

## 2. Agent 启动前检查

每项任务开始修改或运行命令前，在当前目录核实：

1. `git rev-parse --show-toplevel`：确认仓库根目录。
2. `git branch --show-current`：确认当前分支。
3. `git worktree list --porcelain`：确认当前路径实际登记在对应 Worktree 中。
4. `git status --short --branch`：检查已跟踪和未提交改动；涉及本地配置或样本时再检查 `git status --short --ignored`。
5. 对照上表确认分支、路径、任务职责和本次 Cargo target 一致。

若分支、路径或任务职责不匹配，先报告并停止写入；不得通过切换目录或直接访问另一个 Worktree 来绕过检查。不得假定缺失的 Worktree 或分支已经存在，也不得自行删除或重建它们。

## 3. 文件位置与 Cargo target

- 源码、Cargo 清单与锁文件、配置、BAT、文档、测试资源以及每个工作树自己的 `dist` 保留在 C 盘对应目录。
- Cargo 构建产物分别写入上表的 F 盘目录。直接执行 Cargo 命令前，必须按当前分支设置对应的 `CARGO_TARGET_DIR`；PowerShell 使用 `$env:CARGO_TARGET_DIR='对应的 F 盘路径'`，CMD 使用 `set "CARGO_TARGET_DIR=对应的 F 盘路径"`。不要让 Cargo 默认在 C 盘创建 `target`，也不要让并行 Worktree 共用 target。
- `build-release.bat` 根据当前仓库根目录和分支选择 target；成功后只复制到当前工作树的 `dist`。`run.bat` 只运行当前工作树的 `dist\pytools.exe`，缺少文件时失败退出，不自动编译。
- `package.bat` 只打包当前工作树已有的 `dist` 和资源，不隐式构建。Feature/Test 包必须标记为测试包；正式 Release 构建与正式交付包只允许来自 `main`。
- 本规范不授予编译、运行、测试或打包权限。仍须遵守当前用户对本次任务的明确授权；没有授权时不得运行 Cargo 或通过脚本间接编译。不得执行 `cargo clean` 或清理 F 盘缓存。

## 4. 修改权限与跨工作树边界

- Agent 只在任务指定的 Worktree 内读写和构建，不直接修改其他 Worktree 的源码、配置、`dist` 或测试输出。
- Feature 可按其职责修改业务实现及对应测试。Test 可修改任意必要的 Rust 业务源码、Excel-Core 和测试代码，不限于 `tests/` 目录。
- Main 的日常职责是集成和修复集成问题。需要在 Main 直接开发普通业务功能时，应先确认任务明确指定 Main。
- 各 Worktree 的忽略文件（如本地 `config.xlsx`、`dist`、业务样本、运行日志）不由 Git 提交或合并传播。不要覆盖、删除或假定另一个 Worktree 的本地资源与自己相同；确需复制时，先核对目标并保留原件。
- 不为统一其他项目而同步、覆盖或 cherry-pick 旧 Main、Rust-Audit、Python 原项目或旧共享 Workspace 的实现。不得操作 Rust-Audit 或 RustHistory 原始归档。

## 5. 提交、集成与同步

1. Feature/Test 在各自分支完成一批独立工作后，可提交本地 commit；提交前检查 `git diff`、状态和已获授权的相关验证结果。
2. 不自动 push；只有用户明确要求时才推送远端。
3. 集成前核实来源分支与 Main 的路径、分支和状态，并确认未提交改动已被保留和考虑。未经用户明确指定的集成任务，不跨分支合并。
4. 指定集成任务使用普通 Git merge 将已完成的 Feature/Test 修改合入 Main。发生冲突时逐项查看并有针对性地解决；禁止强制 reset、强制 push、覆盖文件或丢弃未提交修改。
5. 合并后在 Main 重新执行本次任务明确授权的必要编译、测试和回归。Feature/Test 单独通过不代表 Main 已通过；Main 验收通过后才可进行正式 Release。
6. Main 稳定后，Feature/Test 可在明确安排下通过普通 Git merge 同步 Main 的新提交；不得自动合并。
7. 本文或 `AGENTS.md` 更新后，应随普通 commit 记录，并在检查各工作树状态、安全确认无冲突后，通过正常 Git merge 传播。修改 Main 的规则不会自动更新 Feature/Test，也不得覆盖它们的未提交改动。

## 6. 冲突与业务数据保护

- 合并前后保留各工作树的未提交改动；若有无法安全纳入或存在冲突的本地修改，先停下并报告，不以 reset、checkout 覆盖或删除文件解决。
- 保护原始 Excel、业务数据、公式和必要结构；优先使用副本。未经明确授权和备份保护，不覆盖原始业务文件。
- 不得未经测试删除既有功能，不因普通字体、边框或外观差异阻塞交付。维持用户逐项人工测试和明确反馈的工作方式，不扩展到未指定的迁移、审计或优化。
- 正式构建和交付只能使用 Main 自己的源码、F 盘 target 与 C 盘 `dist`；Feature/Test 的 EXE 或压缩包不得复制覆盖 Main 的交付物。
