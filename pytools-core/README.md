# pytools-core

此目录保存 Rust-Tools 的 PyTools 业务逻辑、配置处理、测试辅助和历史证据；Excel 原语由本仓库的 `crates/` 提供。当前七类菜单、具体子功能、保留的历史功能、已知限制及状态见 [功能清单](MIGRATION_INVENTORY.md)，业务范围与排除项见 [业务范围](MIGRATION_PLAN.md)。

既有工作簿修改通过 `ExcelEditPlan` / `ExcelModifier` 按具体能力处理。Python 基线及 probe 是按需比较辅助；隔离 openpyxl worker 仅用于受限的既有工作簿修改，不承载业务逻辑。原项目只读；不自动开展全量迁移或对拍。

历史测试和早期候选说明保留在 [归档](../docs/archive/README.md) 及本目录 `GPT修改记录/`、`probe/reports/`；旧菜单编号、旧路径和历史验收结果不自动成为当前指令。编译、运行、Git 和交付规则以 [根 AGENTS.md](../AGENTS.md) 为准。
