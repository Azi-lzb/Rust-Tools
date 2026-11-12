# Excel 真实业务修改能力矩阵

## 审计口径

- 仓库：`C:\Users\AZI\Desktop\VibeCode\Rust-Tools-feature`；分支 `feature/business-development`；代码基线 `497bee806cd9c150d751135a68115ff7f575488e`。本报告也记录随后在该基线上进行的未提交定点修复。
- 纳入菜单：用户指定的 13 项既有工作簿修改。转换、新建汇总、公式重算、打印生成和文件系统操作均不纳入。
- 业务操作以当前 Rust 菜单代码生成的 `ExcelEditPlan.operations` 为准；引擎“通过”只表示该菜单、该计划和该样本的验收，不外推到任意工作簿。
- `.xls` 专用 COM 路径不参与引擎迁移。专项 `.xls` 样本只复制后使用既有转换器派生隔离 `.xlsx`，测试不覆盖或修改原件。
- 根据用户补充的样本选择，5-1 使用 `测试用例\5\2026.09 商业银行共享一批.xls` 与 `2025.09 商业银行共享一批.xls`，分别对 `2026.09 本外币.xls`、`2026.09 外发用.xls` 的隔离 `.xlsx` 副本运行；5-7 使用 `测试用例\5\5-7\2026.09. 专项二批次.xls` 和 `2026.08. 专项二批次.xls` 派生副本。5-3 对 `2026.09 外发用.xls` 的派生副本先执行 5-2 标准化，再测试删除。

## 13 项计划与真实能力

| 菜单 | 计划实际需要的操作 | 当前 XLSX 主路径 | 本轮真实结果 |
|---|---|---|---|
| 1-1 既有目标导入 | 可选 `CreateSheet` + `ImportRange`（值、公式文本和样式按配置导入） | Openpyxl Worker | OpenpyxlOnly、UMYAOnly、Auto 均未提交：目标含 175 个公式缓存，候选守卫无法证明区域修改与缓存互不影响。不能算通过。 |
| 2-1 标记重复 | `SetCellFillColor`（整行标红） | Openpyxl Worker | OpenpyxlOnly、UMYAOnly、Auto 均标记预期 10 行；Auto 实际选择 UMYA。值、非目标样式和打印/合并属性未变。 |
| 2-2 删除重复 | `DeleteRows` | Openpyxl Worker | OpenpyxlOnly、UMYAOnly、BetterOfficeOnly、Auto 均删除 10 行，保留行值序列一致；Auto 实际选择 BetterOffice。COM 回读通过。 |
| 2-3 去重追加 | `ImportRange` / `SetCellValue` / `SetFormulaText`，依配置可选 `CreateSheet` | Openpyxl Worker | OpenpyxlOnly 与 Auto 候选验证拒绝：原有 `Sheet1!A10` 未保留。未提交，其他引擎未宣称通过。 |
| 4-5 改 Sheet 名 | `RenameSheet`；需同步直接公式和定义名称引用 | Openpyxl Worker | 测试文件含 ArrayFormula、跨表公式及 H21 驱动的 INDIRECT。OpenpyxlOnly/Auto 写前拒绝不完整动态引用；UMYA 拒绝带公式的改名；LegacyDefault 旧路径由 A21 校验拒绝。所有副本与原件未变。 |
| 5-1 村镇银行拆分 | `DeleteRows` + `InsertRows` + `SetCellValue`；当前计划没有新增行样式操作 | Openpyxl Worker | 两个目标副本的 OpenpyxlOnly/Auto 均写入正确拆分值并经 COM 回读；UMYA/BetterOffice 未提交。结构复核发现 2 行 × 9 列新单元格均为默认样式，而被替换模板行 9 列原有样式；因此格式验收失败。公式/缓存数量保留（530/530、785/785），但 Openpyxl 删除了 `calcChain.xml`，且大量 OOXML 部件被重写。 |
| 5-2 机构名称规范化 | `SetCellValue` | Openpyxl Worker | OpenpyxlOnly、RustNativeOnly、BetterOfficeOnly、Auto 的输出值一致；Auto 选择 RustNative。UMYA 未完成相同规范化（24 个机构名单元格与基准不同）。成功候选的 785 个公式缓存均保留；RustNative/Auto 保留 calcChain。 |
| 5-3 删除外资行并生成副本 | `DeleteRows` | Openpyxl Worker | 测试组合实际删除数为 0。虽产生可由 COM 打开的副本，但未证明删行行为，所有引擎均不得记为该业务通过。 |
| 5-4 修改外汇页眉单位 | `ReplaceHeaderText` | 定点 Rust OOXML，失败时有 Openpyxl 回退 | RustNativeOnly、Auto 成功，Auto 选择 RustNative；只改 `sheet7.xml`、`sheet8.xml`、`sheet9.xml` 页眉，公式/缓存及 calcChain 保持，COM 回读通过。OpenpyxlOnly 与 UMYAOnly 未提交。 |
| 5-6 工作表数据提取 | `CreateSheet` + `ImportRange` | Openpyxl Worker | OpenpyxlOnly、Auto 导入 462 行；原有公式与缓存相同，COM 回读通过。UMYA 因目标工作簿复杂结构拒绝；Native/BetterOffice 不覆盖完整导入计划。 |
| 5-7 涉农月度修正 | `SetCellValue` | Openpyxl Worker | OpenpyxlOnly、RustNativeOnly、Auto 的业务结果在 1e-8 数值容差内一致；Auto 因批注/复杂对象选择 Openpyxl。RustNative 仅改一个工作表部件且保留全部 236 个公式缓存和 calcChain；Openpyxl/Auto 丢失 5 个公式缓存（5 个表的 B12）及 calcChain。 |
| 5-8 季度修正和地区核对 | `SetCellValue`；地区核对是业务逻辑，不是编辑操作 | Openpyxl Worker | 用户确认样本已删除，本轮免除真实文件测试；只完成静态代码和计划盘点。 |
| 7-1 初始化/修复配置 | `SetCellValue`，部分路径可选 `CreateSheet`、`SetComment` | Openpyxl Worker | OpenpyxlOnly、Auto 均因 S1 表头“是否输出行序号”歧义拒绝；未修改配置副本。初始化新工作簿不在本轮范围。 |

完整命令行、输入/副本路径、选中引擎、COM 读回和包部件证据见 `artifacts/excel-edit-phase2/` 下对应 JSON 与日志；原件哈希复核见验收报告。

## ExcelEditPlan 表达能力

`crates/excel-editor/src/plan.rs::EditOperation` 当前具备：单元格值、公式文本、批注、填充色、行高/列宽、行列插删、工作表创建/删除/改名/移动、范围复制/导入、页面设置、页边距、定义名称、计算标志和页眉文本替换。计划可明确表达本轮 13 项所需的大部分动作，但“表达”不代表任何引擎都支持或能保真执行。

- **样式**：独立修改目前只有 `SetCellFillColor`；任意字体/边框/数字格式没有对应原子操作。范围导入/复制可按 `RangeStyleMode` 复制样式。
- **5-1 新行样式**：实际生成计划只有行插删和值写入，未表达从模板行复制样式。输出副本的新增行 30/31 均是默认样式（18 个样式单元格缺失）。不能把该菜单写入成功视为格式完整通过。
- **合并**：没有任意创建/取消合并操作；仅范围导入的 `copy_merges` 可复制合并区域。
- **打印**：`SetPageSettings`、`SetMargins` 和 `ReplaceHeaderText` 可表达限定页面属性；页眉文本当前是定点替换，不是任意页眉/页脚写入 API。
- **公式**：有 `SetFormulaText` 和计算属性标志，但没有公式重算操作。公式引用变换依赖具体后端；改名不是通用“重写所有公式引用”操作。
- **外链与关系**：外链、图表、图片、calcChain 等可在 `WorkbookFeature` 中声明并用于拒绝/保留判断；没有修改外链或图表数据的操作。普通 `.rels` 由包检查和关系校验覆盖。
- **业务流程**：地区校验、去重规则、银行映射等留在菜单业务代码中；Plan 只承载最终 Excel 修改，不能替代业务计算。

## 引擎可覆盖边界

- **RustNative OOXML**：本轮真实通过 5-2、5-4、5-7 的定点值/页眉计划；5-7 证明该样本中即使存在批注、公式和 calcChain，也能保留未修改部件。不能覆盖行列插删、导入、改名、任意样式。
- **UMYA**：本轮真实通过 2-1 标色、2-2 删除重复行。对带公式的 Sheet 改名、复杂批注/图表/外链对象和未验证的完整导入计划应拒绝；5-2 未得到正确业务值。
- **BetterOffice**：本轮真实通过 2-2 删除重复行、5-2 名称规范化。对 calcChain 行列结构修改、外部链接和复杂对象有拒绝守卫；5-7 因批注结构拒绝。
- **Openpyxl Worker**：覆盖最广，现有业务仍需保留。此轮也测出 Openpyxl 往返会删除 calcChain；5-7 的 Auto/Openpyxl 还清除了 5 个零值公式缓存。真实写入成功和 Excel 能打开不能掩盖这些差异。

## 特殊结构实测范围与空缺

本轮真实样本覆盖：公式及缓存、跨表公式、INDIRECT 动态引用拒绝、批注、合并范围、行列尺寸、页眉页脚/打印设置、calcChain、包关系部件及 COM 回读。没有代表性外部链接、图表系列引用、定义名称、数字签名、ActiveX 或宏文件；这些结构只能报告为未验收，不能由本轮结论外推。详见 `EXCEL_EDIT_ENGINE_GAPS.md`。
