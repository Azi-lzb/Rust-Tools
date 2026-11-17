# Rust-Tools 项目开发规则

## Worktree 角色与协作

- 当前方式（2026-11-16）：用户已停止并行开发，后续业务开发、修复、测试和集成都在本主工作树的 `main` 进行。下表保留历史 Worktree/构建目录映射，不代表对应工作树仍存在，也不是自动重建或恢复并行开发的授权。副工作树退役与备份记录见 `docs/WORKTREE_RETIREMENT_20261116.md`。
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
- 任何直接 Cargo 命令都必须先按上表为当前工作树设置 `CARGO_TARGET_DIR`，避免默认在 C 盘生成 `target` 或多个工作树共用缓存。Release 构建使用 `cargo build --release --offline -p pytools-core --bin RustTools`（main 分支映射 `F:\RustBuild\pytools`）。
- 所有 BAT 必须以脚本自身目录 `%~dp0` 定位项目，只操作当前工作树，不引用旧 Main、其他项目或旧 Codex 工作树。
- 本仓库独立拥有业务代码、Cargo workspace、配置和 `crates/` 内的 Excel-Core。不得依赖旧共享 Main、其他项目或旧共享 workspace 的源码、同步、合并或 cherry-pick；可按当前业务需要独立修改本仓库的 Excel-Core，不为统一其他项目而覆盖现有实现。
- Python 原项目 `C:\Users\AZI\Desktop\VibeCode\4.pytools-main` 只读，`probe/` 和测试辅助中的 Python 仅供按需比较。所有交互菜单不要求外部 ExcelModifier capability 文件；需要修改既有 XLSX 时使用随包隔离 openpyxl worker 或该功能已确定的专用路径，并继续执行候选校验、源文件快照和安全提交。Openpyxl worker 不承载 PyTools 业务规则或公式计算。Windows 上经用户明确授权，可为 5-1 至 5-7 增加与 PyCode 对应的、按功能限定的 Excel/WPS COM 业务操作，优先用于需要处理 `.xls` 的既有功能；这是专用业务路由，不是通用 Excel 修改器，当前需逐项实现和验收。其他读取、新建、转换和公式重算仍分别按 Calamine、rust_xlsxwriter、pywin32 COM 或 Linux LibreOffice Calc 路由。4-5 对 `.xlsx` 使用 Openpyxl（改写跨表公式和命名区域引用），对 Windows `.xls` 使用专用 pywin32 Excel/WPS COM 改名路由（本轮 Excel.Application 已验证，WPS 尚需单独验收）。Linux `.xls` 改名不支持，可先通过 4-4 转为 `.xlsx`；4-5 的 Linux `.xlsx` 菜单路由使用随包通用 Openpyxl worker，改写跨表公式和命名区域、保留可验证的公式缓存，并在候选校验和源快照检查后安全提交；不能安全改写的动态引用拒绝提交。此接入不代表麒麟或 deb 安装验收通过。Linux 1-4 公式重算使用用户提供的 LibreOffice UNO worker。
- 用户确认的平台引擎路由见 `rules/13_Rust_Excel读写修改与引擎路由规范.md`：普通读取用 Calamine，OOXML 专属字段由窄范围 OOXML 补充；从零新建 XLSX 用 `rust_xlsxwriter`，OOXML 只补它不覆盖的明确部件；既有 `.xlsx` 修改用隔离 Openpyxl，必要时配合定点 OOXML。尽量不使用 Umya。1-4 的命名校验区域重算按平台路由：Windows 由 pywin32 调用 Microsoft Excel/WPS COM，Linux 由 LibreOffice Calc 处理已复制到隔离副本中的校验区域；其他功能没有 Linux LO 计算路由。`.xls → .xlsx` 格式转换只在 Windows 由 pywin32 调用 Excel/WPS COM 完成，Linux 不提供此转换。Linux 上需用 Openpyxl/OOXML 修改的 `.xls`，先提示用户在 Windows 转成 `.xlsx` 再处理。
- LibreOffice 当前只用于 Linux 的 1-4 校验区域重算：Rust/OOXML 先把校验区域复制到隔离临时工作簿副本，再由用户提供的 LO Calc 打开、计算并保存该副本，最后由 Calamine 读取结果。默认认为用户已安装/提供 LO；它不随 deb 打包。缺少 LO 时明确报告 1-4 计算依赖不可用，不把 LO 作为通用读取、新建、修改或格式转换引擎；其他功能未经明确任务不得调用 LO。Windows 的 1-4 计算走 pywin32 调用 Excel/WPS COM。`.xlsx` 扩展名不保证内容是 OOXML；4-4 对 OLE/BIFF `.xls` 内容仅改了 `.xlsx` 后缀的文件，会先暂存为 `.xls`、经 COM 确认为旧 Excel 工作簿后再 `SaveAs`；其他无法识别的非 ZIP 内容明确拒绝。
- 保留根目录本地 `config.xlsx` 和业务文件；它们不应提交到 Git。不得改写原始 Python 项目、RustHistory 原始归档或业务原件；不得删除业务源码、配置、测试夹具及人工验收记录。针对明确任务可按需维护合成测试夹具，保留既有验收证据。

## 工作方式和业务安全

- 后续以用户逐项人工测试与明确反馈为主。用户没有指定功能时，不主动扩大迁移、审计、对拍或优化；出现问题时仅修复该功能及必要依赖。既有已迁移功能优先保留，不自动重写，不得未经测试删除旧功能或以未测试的替代品移除现有业务行为。
- 保留七类菜单、既定子功能和明确排除项，见 `DEVELOPMENT_GOAL.md`、`pytools-core/MIGRATION_PLAN.md` 和 `pytools-core/MIGRATION_INVENTORY.md`。历史文档和修改记录是参考，不是新的执行任务。
- 保护原始 Excel 文件、业务数据、公式、必要结构和明确约定保留的对象；优先在副本上操作。涉及覆盖原文件时，必须有明确授权及备份或快照保护。不因普通字体、边框、颜色等外观差异阻塞功能交付；用户明确要求的业务格式和重要结构仍应保护。
- Excel 引擎统一选型与安全规范见 `rules/13_Rust_Excel读写修改与引擎路由规范.md`：读取、新建、既有工作簿修改、特殊属性处理、格式转换和公式计算按操作分别选择引擎。用户要求 PyCode 对齐或引擎调整时，先按该规则 §9.1 备份副本并对拍现状，再定点修复和复测；默认矩阵本身不要求存量功能立即迁移。Excel 原语放在本仓库 `crates/`，PyTools 业务逻辑放在 `pytools-core/`。规则冲突以本轮用户明确要求和根 AGENTS.md 为准。
- 新增或修改业务功能必须遵守 `docs/RUN_LOGGING_STANDARD.md`；它是运行日志的唯一权威规范。
- 新增批处理入口须关联 run_id，记录开始/结束时间、阶段、文件最终状态、必要的 Sheet/单元格定位、输出结果、用户可读错误、结构化错误及统一运行汇总，并为统计和状态判定补充测试；不得凭读取成功推断业务成功，也不得删除旧运行 JSON 的兼容字段。
- 未经用户批准，不执行 Git 强制操作、push 或远端修改。不得主动清除 Cargo target 缓存。

## 主工作区输出数据净化（2026-11-17）

- 本主工作区（项目 A）定位为生产环境；Aspose 作为计算 Worker 被合法调用，必须遵守产品输出规范。本节约束该 Worker 生成的最终 XLSX 候选，不据此变更其他功能的引擎路由。
- 测试/评估环境下，SDK 生成的候选文件可能包含系统级附加对象，例如默认占位的空白或提示工作表。这些对象不属于业务输出；不得仅根据工作表名称或空白状态认定其为系统对象。
- Worker 在保存输出前必须执行标准的“对象结构清理”：遍历工作表，依据本次任务的业务工作表清单及可验证的系统来源，剔除无业务数据的系统附加工作表，并修正文档的默认活动视图。保留业务工作表及其顺序、公式、必要结构和有意保留的空白模板；原活动业务表改名后继续按身份定位，被明确删除或隐藏时选用有效的可见业务表。
- 若 SDK 在序列化保存时重新附加系统对象，必须在隔离候选上完成同等清理及保存后独立结构校验，再交给上层验证和安全提交。无法确认对象来源、混有业务数据、存在不安全引用或没有有效可见业务表时拒绝交付，不得猜测删除。不得覆盖源文件；已有快照、候选验证和安全提交约束继续适用。
- 上述操作属于常规的数据清洗与结构规范，必须包含在单元测试中，断言最终文件仅含任务约定的业务工作表；覆盖系统附加空白/提示表、业务空白表、同名业务表、公式和对象保护、活动视图修正、重复处理及失败保护。实现、已执行测试、真实样本验收和平台验收分别记录。
- 开发诊断按 `docs/RUN_LOGGING_STANDARD.md` 关联 `run_id` 与请求编号，记录清理阶段、对象身份、实际删除数、视图修正及校验结果，不记录业务单元格原值。原始 SDK Save 字节及归档清单保持原样，净化后的最终候选与原始归档明确区分；既有评估模式/授权状态字段如实保留。本节优先于历史文档中“保留最终 XLSX 评估工作表及其活动视图”的约定，不修改 Word 评估内容或历史验收证据。

## 编译、运行和交付

- 实现或代码审查不等于构建授权。只有当前用户请求明确授权相应操作时，才可编译、构建、执行或安装 Rust 代码；禁止未授权的 `cargo build/check/test/run/clippy/bench/install`、`rustc` 及通过脚本间接触发这些操作。运行已有 EXE 也须有相应的明确运行授权。
- 允许静态检查、Git diff 和 `cargo fmt`（含 `cargo fmt --check`）。未执行编译或测试必须明确记录为未验证，不能把写了测试当作通过。
- 用户直接双击 `dist\RustTools.exe` 运行（程序自动向上探测 `config.xlsx`，无需启动脚本）；`dist` 缺少 EXE 时由构建流程补齐，不隐式调用 Cargo。Release 构建成功后只更新当前工作树的 `dist`。正式包只允许从 `main` 在明确确认验证通过后生成；Feature/Test 产物标记为测试包。打包/构建一律使用显式 Cargo 与 `packaging/` 脚本，不再使用 BAT。
- 分别记录实现状态、单元/冒烟测试、Windows 用户或真实样本验收、Linux/deb 验证；不得把菜单入口、编译或合成样本视为人工验收通过。
- Linux 交付要求保持为 x86_64、glibc 2.28 兼容和离线 deb；核心路径不要求用户预装 Python、Go、Office、WPS、LibreOffice 或 JVM。外部引擎仅为明确、按能力选择的可选项。
