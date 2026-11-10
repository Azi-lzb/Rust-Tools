# rustTools 3-1 双引擎验收报告

数据：`results/e2e_compare.json`（8 场景双引擎逐格对拍）、`results/e2e_summary.txt`（执行状态）、
`results/office_compat.jsonl` + `results/office_pdf.jsonl`（Excel COM 实开 + PDF 导出 10 项）。
夹具：`benchmark/rusttools-3-1-dual-engine/fixtures/`（合成配置 8 份 + 双源簿）。
日期：2026-11-11。副本验证，原配置与原模板未动。

## 1. Openpyxl 入口验收（回归）

- print_menu_cli 7/7 通过（含双入口菜单断言、Openpyxl 入口端到端、XLSM/输出隔离）。
- 原 Openpyxl 入口的调用链、Python 脚本、追加既有目标、XLSM、mode 1/2/3 行为**未改动**。
- Openpyxl 输出文件（`out_m1.xlsx` 等）作为对拍参照物生成并验证。

## 2. Rust 入口验收

### 2.1 功能（8 场景 × 双引擎）

| 场景 | Openpyxl | Rust | 值差异 | 样式差异（136 格） |
| --- | --- | --- | ---: | ---: |
| mode 1（全样式） | 成功 | 成功 | 0 | 0 |
| mode 2（自适应列宽） | 成功 | 成功 | 0 | 0 |
| mode 3（快速） | 成功 | 成功 | 0 | 0 |
| multirange（多区域堆叠） | 成功 | 成功 | 0 | 0 |
| emptyrange（used_range 兜底） | 成功 | 成功 | 0 | 0 |
| multitarget-1（双源同目标） | 成功 | 成功 | 0 | 0 |
| multitarget-2 | 成功 | 成功 | 0 | 0 |
| repeatsheet（重复目标 Sheet） | 成功 | 成功 | 0 | 0 |

**结论：8 场景值/合并/行高列宽/页脚/页面设置全部一致（ALL MATCH）**。

### 2.2 样式对拍（136 个带样式单元格）

零差异。覆盖：字体 bold/size、填充 solid RGB、数字格式（含 `0.00%` 自定义百分比）、
对齐水平、合并锚点样式。源样式来自 openpyxl 写出的 styles.xml（含 cellXfs 无 apply*
属性的形态），经 OOXML 窄读取 → descriptor → rust_xlsxwriter Format 全链映射。

### 2.3 输出隔离

- Rust 入口输出 `printed_Rust.xlsx`，Openpyxl 目标 `printed.xlsx` 未被创建或覆盖。
- 源工作簿 SHA-256 前后一致。
- XLSM 目标：Rust 入口明确拒绝（`XLSM 输出，任务未执行，请使用 Openpyxl 入口`），不产出文件。

### 2.4 Excel COM 冒烟（10 项全 PASS）

每个输出在真机 Excel（COM DispatchEx，DisplayAlerts=False）中打开：
工作表数量与 openpyxl 预期一致 → PASS。COM 进程正确退出。

### 2.5 PDF 导出（10 项全 PASS）

每输出经 `ExportAsFixedFormat` 导出 PDF（17–21KB），PDF 文件生成成功，
打印区域经 COM PageSetup.PrintArea 回读确认。

## 3. 数据对拍明细

见 `results/e2e_compare.json`：每场景记录 openpyxl/Rust 双方的 max_row/max_column/
逐格值（repr）/合并列表。mode 1/2/3/多区域/多目标/重复 Sheet 的值差异均为 0。

## 4. 已知问题与限制

1. **XLSM**：Rust 入口明确拒绝（设计决定），Openpyxl 入口保留 XLSM 能力。
2. **追加既有目标**：Rust 入口走 `_Rust` 新路径，天然从新簿开始；如需追加语义，当前版本未实现（Openpyxl 入口保留该能力）。
3. **批注作者**：Rust 写入的批注文本不单独保留 author 字段（openpyxl 的 author 前缀在文本里时随文本带入）。
4. **theme tint**：标准近似公式，与 Excel 精确算法可能有 ±1 色阶差。
5. **页脚语义**：双引擎一致使用「工作表序号字面量」，非真页码（历史行为，本轮不改）。

## 5. 未测试项目

- WPS 抽查（环境无 WPS 自动化通道；标记 NOT_TESTED）。
- mode 2 在 186/210/222 列宽表上的 ### 表现（合成宽表夹具本轮未构建；
  真实配置中 55 条非空区域多为 mode 1，mode 2 需用户构造样本后补验）。
- 修复提示对话框路径（COM DisplayAlerts=False 抑制了对话框，需人工打开确认）。
