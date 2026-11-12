# BetterOffice 作为 XLSX 修改引擎的专项能力盘点

日期：2026-11-12
测试分支：`test/business-coverage`
BetterOffice：`betteroffice-xlsx 0.5.1`、`betteroffice-xlsx-ops 0.5.1`、`betteroffice-xlsx-model 0.5.1`
测试构建：Windows x86_64 Release，离线 Cargo 构建；Windows Excel COM 用于独立打开验证。

## 结论摘要

BetterOffice 0.5.1 能完成常见 XLSX 单元格修改、行列结构操作、简单样式、行高列宽设置。行操作后图表、合并区域、普通命名区域和已有打印配置通过 COM 验证；条件格式规则仍在，但其应用范围未像 Excel 一样扩展。`apply_ops` 对测试中的公式依赖进行了公式文本重写和缓存更新：A1 改为 15 后 `=A1*2` 保存缓存为 30；删除被引用的 B 列后 `=A1+B1` 被改为 `=A1+#REF!`，缓存成为 Excel 错误值。

它目前不适合直接作为通用 `ExcelEditPlan` 修改后端：第一类页面设置没有公开修改 API；通过通用 `SetDefinedName` 试图设置打印区域，Excel COM 读回的 `PageSetup.PrintArea` 是空值。列宽设置的 model 值与 Excel COM 读回值也不一致。历史对拍还记录过 calcChain 丢失和公式相关区域出现意外缓存值。适合先按明确能力白名单做隔离试点，不应把“支持 apply_ops”理解为全功能、高保真修改。

## 测试范围与方法

- 独立工程：[`adapters/betteroffice`](../adapters/betteroffice/)，未修改生产代码。
- 依赖：BetterOffice 三个 crates 均固定 `=0.5.1`；完整依赖锁定在该工程 `Cargo.lock`。
- Release 验证：`cargo test --release --offline` 通过，1 个集成审计测试。
- 合成夹具包含图表、条件格式、合并单元格、定义名称、公式、打印区域、横向页面、页边距、页眉页脚，以及 `customXml/item1.xml` 不透明扩展部件。
- 结构改动：删除第 2 行、再插入 2 行；另测单元格样式、行高、列宽。公式场景修改 A1 并删除 B 列。
- Excel COM 对结构、打印设置和公式结果作只读交叉验证。另用 Excel COM 创建真实批注夹具，再由 BetterOffice 改写普通单元格并重新用 Excel COM 读取批注。
- XLS 未测试。本报告只论 XLSX。

## 能力矩阵

状态说明：PASS=本轮实际操作、保存并验证通过；PARTIAL=有受限实测；FAIL=已尝试但结果不符合预期；UNSUPPORTED=接口/产品明确无此能力；API_UNAVAILABLE=公开 API 中没有相应能力；NOT_TESTED=本轮没有足够实测。

| 操作/属性 | 状态 | 本轮证据与边界 |
|---|---|---|
| 打开、修改、保存 XLSX | PASS | Rust 适配器读取、修改、保存并重新打开；COM 能打开输出。 |
| 修改单元格值 | PASS | A1 修改为 15；另在含批注夹具中改 E1，COM 读到 `ordinary edit`。 |
| 删除行、插入行 | PASS | 删除 1 行并插入 2 行；保存后图表、条件格式、合并和名称仍有效。 |
| 删除列、插入列 | PARTIAL | 删除 B 列后公式引用按删除语义重写；性能场景中插入/删除行列操作成功，但该大型输出未逐项做结构语义对拍。 |
| 追加 1000 行、追加表格行 | NOT_TESTED | 本专项没有测追加行为。 |
| 单元格样式 | PARTIAL | `PatchRangeStyle` 对 A1 加粗，COM 确认 `Font.Bold=True`；字体、颜色、边框、数字格式、范围批量复制没有逐项实测。 |
| 行高 | PASS | 将第 2 行设为 27，COM 读回 27。 |
| 列宽 | PARTIAL | Rust model 保存请求值 21.0，Excel COM 读回 20.38；Excel 的列宽单位转换需应用层确认，当前不能视为精确值往返。 |
| 隐藏行列、AutoFit | NOT_TESTED | API enum 中存在可见性操作；本轮未实测，未发现 AutoFit 专用 `Op`。 |
| 图表保留 | PASS | 结构改动后 BetterOffice model 和 COM 均读到 1 个图表。 |
| 条件格式保留与范围重映射 | PARTIAL | 规则本体还在，但范围没有按 Excel 行操作更新。BetterOffice 输出由 COM 读到 `A1:B3`；Excel COM 对相同原文件执行删除第 2 行、插入第 2–3 行后，范围为 `A1:B4`。 |
| 合并单元格保留/行偏移 | PASS | 合并区域从 A5:C5 随净增行数移动至 A6:C6，COM 确认合并仍存在。 |
| 普通定义名称引用维护 | PASS | `DataRange` 从 `Data!$A$1:$B$3` 因删除/插入操作变为 `Data!$A$1:$B$4`；COM 读回相同引用。 |
| 已有打印区域在结构操作后维护 | PASS | 原打印区域 A1:C5 随行操作变为 A1:C6；COM `PageSetup.PrintArea` 读回 `$A$1:$C$6`。 |
| 主动设置打印区域 | FAIL | 尝试 `SetDefinedName("_xlnm.Print_Area", …)`；COM 的 `PageSetup.PrintArea` 为空，名称被输出为普通局部名称 `Data!Print_Area`。不能用这个通用名称操作替代打印区域 API。 |
| 页边距、方向、纸张大小、页眉页脚主动修改 | API_UNAVAILABLE | 公开 `Op`/工作簿/模型类型没有 PageSetup、margin、orientation、paper-size、header/footer 修改操作；本轮不能执行修改。已有设置在结构修改后可保留。 |
| 公式引用重写和缓存更新 | PASS（限测试公式） | 删除 B 后公式改为 `=A1+#REF!`，缓存为 Excel 错误；A1=15 后 `=A1*2` 缓存为 30。COM 的 Formula/Value2 与输出一致。不是对所有 Excel 公式兼容性的结论。 |
| 计算链保留 | PARTIAL | 合成夹具原本不含 `xl/calcChain.xml`，所以本轮没有验证其保留；旧基准 `real_target.xlsx` 编辑记录显示该部件在重写后丢失。 |
| Excel 批注保留 | PASS | COM 创建 A1 批注，BetterOffice 修改 E1 后再由 COM 读回完整批注文本。 |
| 未知 customXml 部件保留 | PASS | `customXml/item1.xml` 编辑前后 SHA-256 都为 `F96A9DE3C2DF2050562E496B8CC989F6A1E874C7B614C7C7E190F65849DA4A8B`，内容字节相同。仅证明该不透明部件，不代表任意关联扩展关系都保真。 |
| XLS 原格式修改 / XLS 转换 | NOT_TESTED | 本专项针对 BetterOffice XLSX 修改能力。 |

## 公式引用维护细节

测试输入有 `C1 = A1+B1` 及 `D1 = A1*2`。对同一批操作先将 A1 从 10 改为 15，再删除 B 列。Excel COM 重新打开结果后读到：

| 输出位置 | 公式文本 | Excel `Value2` | 判定 |
|---|---|---:|---|
| B1（原 C1） | `=A1+#REF!` | `-2146826265`（Excel `#REF!` 错误值） | 正确反映引用列被删除 |
| C1（原 D1） | `=A1*2` | `30` | 对修改后的 A1 重算成功 |

因此，本测试中的 `apply_ops` 操作伴随计算并更新了缓存；这不是“只改公式文本、不重算”。但公式覆盖仅限以上基础公式，不能据此推断对 `INDIRECT`、外部链接、动态数组或所有 Excel 函数都具备兼容性。旧历史报告中 `f_complex.xlsx` 编辑被语义对拍判为失败：输出出现原先缺失的 G1/B3:B6 数值；该历史结果提示复杂公式/缓存的行为需要用真实业务文件逐项验收。

## 打印设置 API 盘点

使用 `scripts/inventory_public_api.ps1` 搜索 BetterOffice 0.5.1 SDK 源码及仓库实际使用的 patched model，并列出完整公开 `Op` 变体：有单元格、结构、尺寸、冻结窗格、超链接、自动筛选、条件格式、图表、合并、样式、数字格式和定义名称操作；未找到页边距、方向、纸张、页眉/页脚、分页/缩放、居中、打印顺序等 PageSetup 修改操作。工作簿模型同样没有对应可写字段。命中 XML schema、测试数据或 Print_Area 内部重映射逻辑，不等于存在相应公开修改 API。

`SetDefinedName` 可以维护 Excel 文件里已经存在的 `_xlnm.Print_Area` 引用，结构操作后也实际观察到打印区域正确偏移；但用它新设打印区域的探测失败（见能力矩阵）。其余打印参数因此标为 `API_UNAVAILABLE`，不能宣称支持打印设置修改。

## 10 万行性能

测试数据为 100,000 行 × 3 列、XLSX 输入 1,917,270 bytes。Release 常驻进程执行一次不计入统计的 open/save 预热，随后每种场景三轮；每轮从相同原始输入字节重新打开。进程峰值内存由 PowerShell 每 25ms 采样工作集。三轮统计如下：

| 场景 | open 中位数 | 操作中位数 | save 中位数 | 总耗时中位数 | 各轮总耗时 |
|---|---:|---:|---:|---:|---|
| 修改 1000 个单元格 | 3.529 s | 1.621 s | 0.911 s | 6.061 s | 5.959 / 6.814 / 6.061 s |
| 插入 3 行、删除 2 行、插入 1 列、删除 1 列 | 3.537 s | 9.469 s | 1.386 s | 14.424 s | 14.424 / 14.677 / 14.177 s |

六轮合计 wall time 67.22 s、CPU time 66.61 s；整个常驻进程采样峰值工作集 1,049.55 MiB，峰值私有内存 1,131.19 MiB。峰值是整个六轮进程口径，不能分配到某一个场景。输入/输出大小分别约 1.92 MB。没有在本机重测 emend 和 sheetpatch。

用户提供的历史数值：emend 0.839 s / 203 MiB；sheetpatch 0.565 s / 171 MiB。若假定完全同一工作负载与内存口径，BetterOffice 本轮修改 1000 单元格的 6.061 s 和 1.05 GiB 明显更慢、更占内存；但对比双方数据形状、引擎版本、测量是否包含打开/保存、运行进程生命周期及内存口径未齐全，因此这里只作性能风险提示，不作为严格同机排名。结构修改中位数为 14.424 s，当前没有同场景竞争库基线。

## 历史对拍信号

此前仓库 `benchmark/excel-modify-engine/results/` 中 BetterOffice 0.5.1 的历史结果只作为参考，不并入本轮计时排名：

- `f_complex.xlsx` 的编辑被原对拍器标为 FAIL；意外差异集中在原始文件中为空/缺失的 G1、B3:B6 输出数值。需要结合真实文件意图判断这些值是有益重算还是破坏性改写。
- `real_target.xlsx` 编辑被标为 FAIL，失败项明确包括 `xl/calcChain.xml` 被删除。
- `f_comments.xlsx` 的简单编辑历史样本 PASS；本轮又用 Excel COM 批注夹具实测确认批注仍可读。
- 这些结果和本轮的 customXml 夹具并不矛盾：源包保留是按具体部件/操作路径核验的，不能泛化成“所有 Excel 扩展部件总是逐字节保留”。

## 强项与弱项

**强项**

- `apply_ops` 提供批量操作接口，行列插删、单元格编辑、样式、尺寸、合并、图表、条件格式、名称等能力集中在一个修改模型中。
- 本轮行插删后，图表、合并区、名称及既有打印配置经 Excel COM 验证；条件格式对象保留，但范围重映射存在缺陷。
- 测试公式发生结构变化后，公式文本和缓存均得到相应更新。
- 未触及的 `customXml/item1.xml` 部件保留字节内容；批注对象也通过 Excel COM 保留。

**弱项**

- 页面设置写 API 缺失；通用定义名称写入不能可靠设置打印区域。
- 列宽的用户传入值与 Excel COM 读回值不一致，需明确单位语义或在桥接层校准。
- 公式测试覆盖面很窄；历史真实复杂文件中有意外输出变化，且已出现 calcChain 部件丢失。
- 10 万行简单三列样本的 1000 格编辑约 6 秒；结构操作约 14.4 秒，六轮运行峰值约 1.05 GiB。对用户给出的 emend/sheetpatch 参考值，性能/内存需要进一步同条件复测才能证明达标。

## 路由建议

当前不建议把 BetterOffice 作为通用 XLSX 修改后端直接接入 `ExcelEditPlan`。可考虑将其限定为独立、可选的结构/公式修改试验路由，但需先明确列宽换算规则，并对每类目标工作簿检查输出中未授权变化。打印修改继续走现有已验证路径。对复杂公式、calcChain、外部链接或重要扩展部件的真实文件，应先做副本试验并逐部件核验；不能依赖笼统的“源包保留”承诺。

## 可复现证据

源码和脚本位于 [`adapters/betteroffice`](../adapters/betteroffice/)。本次结果文件：

- `adapters/betteroffice/results/audit.json`
- `adapters/betteroffice/results/com-validation.json`
- `adapters/betteroffice/results/comment-roundtrip.json`
- `adapters/betteroffice/results/perf.json`
- `adapters/betteroffice/results/perf-process.json`
- `adapters/betteroffice/results/api-inventory.txt`
- 合成输入/输出工作簿及三轮性能输出均保留在同目录 `results/`。

Release 验证命令及环境要求见 [`README.md`](../adapters/betteroffice/README.md)。
