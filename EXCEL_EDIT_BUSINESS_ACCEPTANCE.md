# ExcelEditPlan 第二阶段真实业务验收

## 范围和执行方式

- 仓库分支：`feature/business-development`；代码基线 `497bee806cd9c150d751135a68115ff7f575488e`。
- 只执行 13 项既有工作簿修改；10/50/100 文件性能测试未启动，因为正确性验收存在明确失败、未覆盖或未验收项。
- 用户测试集：`C:\Users\AZI\Desktop\VibeCode\Rust-Tools\测试用例`。每次路由均使用独立副本；7 个 `.xls` 专项样本经已有转换命令在临时目录派生 `.xlsx`，没有修改原始 `.xls`。
- 按用户补充的选择，5-1 以 `5\2026.09 商业银行共享一批.xls` / `5\2025.09 商业银行共享一批.xls` 为本期/去年同期，分别测试 `5\2026.09 本外币.xls`、`5\2026.09 外发用.xls` 派生副本；5-7 以 `5\5-7\2026.09. 专项二批次.xls` / `2026.08. 专项二批次.xls` 为本期/上期。5-3 对“外发用”派生副本先运行 5-2 标准化再检查删除数。所有这些业务路由均在临时 `.xlsx` 副本验收；不覆盖 `.xls` 输入原件。
- 读取配置、日志和验证副本不覆盖样本原件。所有执行结果需同时看应用级业务状态、实际路由行、候选提交状态和差异，不以进程 exit 0 单独判成功。
- 所有路径、命令、引擎行、输出 SHA256、ZIP 部件差异、COM 结果及完整日志均保存在 `artifacts/excel-edit-phase2/`。

## 菜单验收

| 菜单 | 路由及结果 | 修改前后差异与结构核验 | Excel COM / 原件 |
|---|---|---|---|
| 1-1 | OpenpyxlOnly、UMYAOnly、Auto 均未提交 | 目标已有 175 个公式缓存。ImportRange 预检无法证明导入范围与所有公式缓存相互独立，因此候选拒绝。现有公式/缓存未改。 | COM 只读打开未修改目标；测试副本与原件 SHA 未变。不能算写入业务通过。 |
| 2-1 | OpenpyxlOnly、UMYAOnly、Auto(UMYA) 通过 | 按第 8 列标记 10 行；所有值不变，非填充样式不变，打印属性和合并区域不变。COM 回读本菜单三份输出共核对 11,487 个常量值。 | 三份输出 COM 只读回读一致；原件 SHA 未变。 |
| 2-2 | OpenpyxlOnly、UMYAOnly、BetterOfficeOnly、Auto(BetterOffice) 通过 | 输入 463 行，输出 453 行；删除 10 行，剩余值序列一致；公式/缓存相同；打印设置和合并范围相同。按存活源行到目标行映射逐单元格比对，四种模式样式差异均为 0，行高差异为 0，列尺寸相同。差异包部件：Openpyxl 8 项、UMYA 10 项、BetterOffice/Auto 各 5 项。 | 四份输出 COM 回读通过，存活单元格值匹配；原件 SHA 未变。结构复核见 `structural-verify-results.json`。 |
| 2-3 | OpenpyxlOnly、Auto 均拒绝候选 | 候选验证报告既有 `Sheet1!A10` 未保留；没有提交。未用缩小数据或删掉格式让测试通过。 | 副本/原件未变；其他路由未宣称通过。 |
| 4-5 | LegacyDefault / OpenpyxlOnly / UMYAOnly / Auto 均未完成 | 最新 Debug 副本复验：OpenpyxlOnly 和 Auto 在候选写入前发现 `INDIRECT($H$21&"!A1")` 的 H21 仍是正在重命名的工作表名，明确拒绝不完整改写；UMYA 拒绝含公式引用的改名。旧 LegacyDefault 由其 A21 非公式单元格校验拒绝。直接公式及定义名称改写另有单元测试覆盖。 | 四份输出副本哈希未变，原件 SHA 未变，Excel COM 能只读打开未变副本。动态改名未验收通过。 |
| 5-1 | OpenpyxlOnly、Auto(Openpyxl) 写入分拆值；格式验收失败；UMYA/BetterOffice 未提交 | 共享银行本期/去年同期分拆：本外币目标改 20 个单元格，外发用目标改 60 个。两个计划均删除第 30 行后插入 2 行。新增行 30/31 各 9 个单元格全部丢失模板样式（共 18 个样式空洞），而被替换模板行具有样式；当前计划没有复制样式操作。公式数量/非空缓存数分别保持 530/530 与 785/785；`calcChain.xml` 被 Openpyxl 保存移除。38 页目标与 10 页目标重写了大量工作表/关系部件。 | 四份输出 COM 回读通过，核对 25,870 个常量值和 2,630 个公式文本；COM 成功不抵销新增行样式缺失。派生输入和 `.xls` 原件 SHA 未变。该测试不验证 `.xls` COM 路径。 |
| 5-2 | OpenpyxlOnly、RustNativeOnly、BetterOfficeOnly、Auto(RustNative) 通过；UMYA 未正确完成 | 规范化值与 Openpyxl 基准一致。UMYA 输出有 24 个机构名单元格与基准不同，故不通过。四种成功结果样式、批注、打印、行列尺寸、合并区域一致。公式/缓存数量 785/785；RustNative/Auto 保留 calcChain，Openpyxl/BetterOffice 删除 calcChain，BetterOffice 亦未保留 calcChain。 | 成功输出 COM 回读：各核对 1,947 常量值、785 个公式文本且无差异；原件 SHA 未变。 |
| 5-3 | 四种模式都生成副本，但业务未验收 | 本次规范化后删除数量为 0，故没有证明删除外资行；不把“副本可打开”写成通过，也不扩大现有格式支持范围。 | 四份副本 COM 回读通过；原件 SHA 未变。删行能力仍待有匹配行的脱敏 XLSX 样本。 |
| 5-4 | RustNativeOnly、Auto(RustNative) 通过；OpenpyxlOnly/UMYA 未提交 | 本外汇用输入三个目标表页眉替换单位文本。包仅变化 `xl/worksheets/sheet7.xml`、`sheet8.xml`、`sheet9.xml`，公式/缓存及 calcChain 数量均保持 785/785；页边距、打印属性无意外变化。 | 两份输出 COM 回读各核对 1,947 常量值、785 个公式文本，无差异；原件与派生输入 SHA 未变。 |
| 5-6 | OpenpyxlOnly、Auto(Openpyxl) 通过；UMYA 拒绝 | 导入 462 行；目标新建工作表。既有 175 个公式和缓存逐项未变；未执行公式重算。 | 两份输出 COM 回读各核对 3,841 常量值、175 个公式文本，无差异；原件 SHA 未变。 |
| 5-7 | OpenpyxlOnly、RustNativeOnly、Auto(Openpyxl) 业务值通过；UMYA/BetterOffice 因批注结构拒绝 | 修正 32 个 D 列值；Native 仅改 `sheet19.xml`，保留 236 个公式缓存、calcChain、样式、批注、打印设置、尺寸及合并。Native 与 Openpyxl 有 7 个数值最后位差异，最大绝对差 `1.1641532182693481e-10`，小于预设 `1e-8`；业务总计分别 `98274.899632` 与 `98274.89963199993`。Openpyxl/Auto 清除了 5 个缓存：`sheet10/11/12/18/20.xml!B12`（公式相同、缓存 0 变空），并删除 calcChain。 | 三份成功输出 COM 回读各核对 19,844 常量值、236 个公式文本；原件和派生输入 SHA 未变。 |
| 5-8 | 用户明确免除真实文件测试 | 用户确认本期/上期样本已删除；仅静态盘点，不能声称引擎验收通过。 | 无真实工作簿及 COM 验收。 |
| 7-1 | OpenpyxlOnly、Auto 均拒绝 | S1 表头“是否输出行序号”有歧义，当前计划期待列 20；候选未提交。未修改配置布局以制造成功。 | COM 只读打开未变副本；原配置 SHA 未变。 |

### COM 汇总

`com-readback-results.json` 对 26 个成功输出副本逐单元格读取 Excel COM `.Value2` / `.Formula`，共核对 **138,981 个常量单元格、11,538 个公式文本、0 差异**，所有打开均为只读且 `SaveChanges=False`。4-5 的四个失败副本另有 COM 只读打开记录；COM 能打开不代表业务改名成功。

### 原件完整性

- 最终扫描 `sample-inventory.json` 中 149 个用户样本：147 个哈希与初始盘点一致；2 个 `.xlsx` 不一致：`1（1-1，1-2，1-3，1-4）\2026.09 4010 交通银行 仲恺区存贷统计月报.xlsx`（初始 SHA-256 `464f06693993569a5275e75d869e0bc991647cbdcac236e52bb217723ee9bac6`，当前 `c76a655f72e76c5a4dc28fd8d6ebdb3fefd30b3d3fa6bef1d463d827451142ec`，mtime 2026-11-12 23:47:54）和 `1（1-1，1-2，1-3，1-4）\2026.09 4020 中信银行 仲恺区存贷统计月报.xlsx`（初始 `67fb93071b821578aabfb8fccf4c2963449bd646ba330a722864822f78d3ad72`，当前 `80bfb4cd69b8c1c69482bfc92a050b8729098ee0818f24bc2864ef8f49eaabfb`，mtime 2026-11-12 23:45:41）。可用备份的 SHA-256 与两份初始哈希一致；本轮没有覆盖或恢复它们。验收路由脚本均写隔离副本，仓库中也未找到全量测试直接引用这两个文件的路径，但修改发生在全量测试运行的时间段，无法排除外部进程或测试之外的操作，因此不能归因，也不能声称全部原件完整性通过。
- `special-staging.json` 中 7 个原始 `.xls` SHA256 变化数 0；7 个派生 `.xlsx` 输入哈希也保持不变。手动路由输出都写入 `AppData\Local\Temp` 隔离副本，没有把实验输出复制回样本目录。
- Feature 根目录 `config.xlsx` 与用户主工作树 `Rust-Tools\config.xlsx` 均只记录哈希，未被本轮脚本修改。

## Auto 与候选隔离

- `route_plan` 根据声明的完整操作集合和 WorkbookProfile 在内存中选路；路由判定自身不访问文件系统。
- Auto 目前优先选一个完整操作集合可接受的引擎；混合操作不能由某一引擎完整覆盖时，选单个 Openpyxl Worker。测试回归覆盖“页边距+插行”“填充+页边距”“删行+改名”等组合，不再自动分成 Composite 两段。
- Auto 在 RustNative/UMYA/BetterOffice 候选执行失败后，可按显式顺序尝试 Openpyxl；`ExcelModifier::execute` 为每次尝试重新复制原输入并分配独立候选，验证失败的候选不发布。只有全计划验证成功的候选进入就地安全提交。
- 4-5 真实失败证明计划能力分类并不能知道 H21 的单元格字符串依赖；最新 Worker 在写候选前拒绝该动态引用，所以没有先部分改名再交给其他引擎。此类拒绝是安全失败，不是业务成功。
- `OpenpyxlOnly` 路由仍调用现有 Worker 及候选校验/安全提交链；默认值未改。LegacyDefault 与显式 OpenpyxlOnly 在 4-5 上均未覆盖该真实计划的成功路径，其他成功菜单的原行为基准见上表。

## 测试和构建

| 验证 | 状态 | 证据 |
|---|---|---|
| Excel Worker Python 单测 | 已通过，26 项 | `openpyxl-python-regression-final.log` |
| Rust 公式改名回归 | 已通过，1 项 | `rename-formula-rust-regression-rerun.log` |
| Auto 路由混合计划回归 | 已通过，3 项 | `router-regression-final.log` |
| workspace `cargo fmt --all --check` | 未通过 | `rustfmt-check.log` 指向若干未改动既有文件的格式差异；本轮未为通过而改写无关业务文件 |
| 本轮修改 Rust 文件定点 rustfmt（edition 2021） | 已通过 | 5 个修改/新增 Rust 文件 |
| `cargo test --offline -p pytools-core --no-fail-fast` | 失败：422 passed、12 failed、12 ignored；4 个目标失败 | `cargo-test-final.log`；本次设置隔离测试 Worker 后运行 |
| Debug / Release 构建 | Debug、Release 均通过；Release 6m 52s | `debug-build-final.log`；`release-build-final.log` |
| 10/50/100 文件性能对比 | 未执行 | 正确性存在未通过/未验收项，按任务门槛不启动性能比较 |

### 全量测试失败项

- `conversion_menu_cli`：6 项失败，包含 4 项 Sheet 改名测试和 2 项 COM 转换测试；运行时选择的 pywin32 Python 缺少 `pythoncom`，另一个转换用例找不到打包的 Excel Worker。
- `sheet_rename_execution`：4 项失败，均因打包 COM Worker 的 Python 缺少 `pythoncom`；1 项需 Office/WPS 的测试按要求忽略。
- `print_comment_execution`：1 项失败，候选缺少 `__pt_print_seed_1` 工作表。
- `print_config_new_target`：1 项失败，候选 `Printed!C2` 值/类型与预期不符。

UMYA 删除行回归 `dedup_comment_delete_umya`（2 项）和去重菜单回归 `dedup_menu_cli`（17 项）在本轮代码修正后通过；外资行批处理回归也通过。打印失败不是本轮目标，未扩大修复范围。全量测试仍未全通过，性能对比因此跳过。

### 本轮回归修复

- UMYA 行删除会物化一个默认列宽记录；适配器现在只在计划未改列定义且工作表集合不变时，把原始 `<cols>` 节点恢复到候选包，避免把默认宽度误写成自定义宽度。
- 候选验证器现在按行删除映射预期公式、合并区域与定义名称；无删除行计划不再经过 UMYA 名称地址规范化。
- `dedup_comment_delete_umya` 2 项通过；`dedup_menu_cli` 17 项通过；`cargo test` 全量结果及原始终端日志见 `cargo-test-final.log`。

## 最终问题 A-G

A. 13 项指定 XLSX 修改业务都存在 Openpyxl Worker 主路径或回退路径；其中 12 项以 Worker 为主要路径，5-4 的主要路径是既有定点 OOXML，Openpyxl 仅在失败时回退。2-1/2-2 的 `.xls` COM 路径不计入本轮。

B. 本轮证明 RustNative 可完整处理 5-2、5-7 的受测计划；5-4 本来已有 Rust OOXML 定点实现且本轮再次通过。对任意模板或超出所测输入结构，不外推为全菜单替代完成。

C. UMYA 已通过 2-1、2-2；BetterOffice 已通过 2-2、5-2。UMYA 的 5-2 真实值不同，5-1/5-6/5-7 的复杂结构被拒绝；BetterOffice 的 5-1/5-7 未提交，5-3 没有命中删除行。

D. 现有 Worker 必须保留，特别是导入/追加（1-1、2-3、5-6）、改名（4-5）、复杂拆分/修复（5-1、7-1）及本轮缺失样本的 5-3/5-8。具体业务是否可以改用别的引擎，须逐项补齐结构和结果验收。

E. Auto 具备单引擎优先、候选隔离、失败后从原件副本回退的执行保护；但因 Worker 往返可删除 calcChain/公式缓存，且某些动态公式语义要到适配器预检才知道，当前只能评为“候选提交安全，工作簿语义覆盖尚不充分”。不建议把 Auto 视为所有 13 项都已安全迁移。

F. 暂不改变默认引擎。可考虑后续仅给 5-2、5-7 增加业务级 RustNative 灰度；5-7 相比 Openpyxl 保留了 5 个公式缓存和 calcChain。需要先复核实际用户模板、数值精度策略以及平台/运行日志要求，不能用全局 `RustNativeOnly` 替代默认。

G. 剩余工作量为中到高：补齐 1-1/2-3 的公式缓存安全验证，完成 4-5 直接/动态/图表/定义名称引用矩阵，验证 5-1 结构改动引用与 calcChain 策略，补 5-3 命中样本、5-8 脱敏样本，再覆盖外链、图表、名称、关系文件和多版本模板；然后才能逐业务灰度默认。
