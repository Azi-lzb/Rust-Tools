# 阶段二：Excel 业务能力验证

状态：第一轮 API 能力门禁已执行。只有阶段一 `GLIBC228_PASS` 的 Aspose.Cells .NET、GemBox.Spreadsheet .NET 和 LibXL C++ 进入测试。Aspose.Cells C++ 与 Spire.XLS C++ 因阶段一 GLIBC 不兼容，标为 `BLOCKED`。麒麟 V10/UOS V20 实机尚未测试。

测试运行于 Rocky Linux 8.9 x86_64、glibc 2.28。两个 .NET 程序使用 .NET 10.0.404 的目录式 self-contained 发布；所有 SDK 在断网容器内直接运行。输入是 [`../fixtures/capabilities.xlsx`](../fixtures/capabilities.xlsx) 和 [`../fixtures/capabilities.xls`](../fixtures/capabilities.xls)。调用和断言在 `dotnet/aspose/Program.cs`、`dotnet/gembox/Program.cs`、`libxl/src/libxl_smoke.cpp`。独立语义与 OOXML 部件检查结果为 [`evidence/phase2-validation.json`](evidence/phase2-validation.json)。

## 能力对照

| 引擎 | XLS 原格式修改 | XLSX 修改 | XLS→XLSX | 追加 | 插入/删除行列 | 样式 | 行高/列宽 | 打印配置 | 公式计算 | 保存保真度 |
|---|---|---|---|---|---|---|---|---|---|---|
| Aspose.Cells C++ | BLOCKED | BLOCKED | BLOCKED | BLOCKED | BLOCKED | BLOCKED | BLOCKED | BLOCKED | BLOCKED | BLOCKED |
| Aspose.Cells .NET 26.10.0 | PASS* | PASS* | PASS* | PASS*（1000 行） | PASS*（含公式/命名范围依赖插入行验证） | PARTIAL* | PARTIAL* | PARTIAL* | PASS*（目标公式集） | PARTIAL* |
| Spire.XLS C++ | BLOCKED | BLOCKED | BLOCKED | BLOCKED | BLOCKED | BLOCKED | BLOCKED | BLOCKED | BLOCKED | BLOCKED |
| GemBox.Spreadsheet .NET 2026.10.101 | PASS* | PASS* | PASS* | PARTIAL*（100 行；FREE-LIMITED 150 行上限） | PASS*（含公式/命名范围依赖插入行验证） | PARTIAL* | PARTIAL* | PARTIAL* | PASS*（目标公式集） | PARTIAL* |
| LibXL C++ 5.2.0 | PASS* | PASS* | PARTIAL*（样例单元格/公式重建） | PARTIAL*（100 行；evaluation 上限 300 个可读取单元格） | PARTIAL*（结构操作通过，公式/命名引用语义未验证） | PARTIAL* | PASS*（固定尺寸/隐藏） | PARTIAL* | UNSUPPORTED | NOT_TESTED |

`*` 表示本轮指定合成样例和评估授权下的实测结果，不代表商业授权、麒麟/UOS 实机验收、任意工作簿的完整支持或生产资格。Aspose 评估输出加入 `Evaluation Warning` 工作表；GemBox 使用 150 行免费评估许可；LibXL SDK 插入试用提示并限制 300 个读取单元格。三者都需合法生产授权后重跑许可相关验证。

## 已执行操作与可复现证据

### XLS / XLSX / 转换

- Aspose .NET 与 GemBox .NET 分别打开 XLSX 和 XLS，保存为相同格式，再重新打开确认 B3 数值；另外修改 B3 并保存/重开确认数值变化。PASS。
- 两个 .NET SDK 分别加载 XLS 并保存 XLSX，重开后验证 B3 及工作表。PASS。此结论是测试样本转换结果，不等于所有 BIFF 功能保留。
- LibXL 在 XLS、XLSX 上完成原格式读改存并重开验证。PASS。LibXL XLS→XLSX 直接加载被 SDK 拒绝；另有从 XLS 手动复制测试标量值和公式到新工作簿的 `RECONSTRUCTION` 样例，不能算原生转换。
- 输出文件与 SHA-256、大小、语义摘要见 `evidence/phase2-validation.json` 和 `reports/samples/`。

### 追加、行列、样式、尺寸

- Aspose .NET 逐行追加 1000 行、随后插入/删除 2 行和 2 列；保存并重开后验证最后追加行、数值、样式、尺寸和工作簿结构。PASS。该原型使用逐单元格 API，不能推断为最优批量写入性能。
- GemBox 在 `FREE-LIMITED-KEY` 下追加 100 行（工作表总行数仍小于 150），并插入/删除行列；保存重开后验证。1000 行 append `BLOCKED_LICENSE`，未伪装为通过。
- LibXL evaluation build 在两种格式各追加 100 行并保存重开；进一步 1000 行测试受 300 个可读取单元格上限阻断。100 行样本通过。
- Aspose、GemBox 实测字体粗体/大小/颜色、填充、边框、对齐、换行和数字格式中的代表项；LibXL 实测字体、填充、底边框、水平对齐及样式复制。自定义日期格式矩阵、范围批量样式、样式表增长及跨单元格意外影响仍 `NOT_TESTED`，因此样式能力标 `PARTIAL`。
- 行高列宽、隐藏/取消隐藏的固定值操作已执行并重开检查。原生 AutoFit 未执行；不能把固定尺寸操作解释为 AutoFit。
- Aspose、GemBox 在行插入/删除场景检查了命名范围依赖公式重算；插入位于被引用区域内时结果扩展，删除后结果恢复。LibXL 结构改动后公式/命名引用语义没有独立确认。

### 打印设置

- 三个兼容引擎都在输出工作簿中写入 A4、横向、页边距、打印区域、重复标题、页面缩放/适应宽高、居中、页眉页脚及页码字段；Aspose/GemBox 对输出重开后再次检查设置，LibXL 使用 SDK getter 校验主要字段。
- 尚未完成全字段矩阵（例如纵向设置、多工作表分别变更、分页符、打印顺序的独立验证及 XLS/XLSX 所有字段逐项断言），因此总表标 `PARTIAL`。不导出 PDF。
- 无修改保存的语义对比发现 GemBox 输出没有保留原工作表的中心页眉文本；其他被检查的 A4/横向、打印区域、标题行列、页脚字段可读取。页眉损失是保真度限制。

### 公式计算

- Aspose.Cells .NET 与 GemBox 使用 SDK 创建公式工作簿，调用真实 SDK `CalculateFormula()` / `Calculate()`，验证首次结果、重复全量计算、修改输入后重新计算、修改公式文本、命名范围和跨表依赖、插入/删除依赖行、保存后公式文本与缓存结果。
- 目标公式包括 `SUM`、`SUMIF`、`COUNTIF`、`IF`、`AND`、`OR`、`VLOOKUP`、`INDEX`、`MATCH`、`ROUND`、`DATE`、跨工作表引用、命名范围、多层嵌套和 `#DIV/0!` 错误结果。公式工作簿输出及缓存摘要在 JSON 证据中。
- LibXL 公式计算 `UNSUPPORTED`（按设计无计算引擎）。
- 同一组 10 个公式、相同单元格输入已在 Excel 16.0 build 20527 COM 与当前 BetterOffice 引擎上执行。`SUM/SUMIF/COUNTIF/IF/AND/OR/VLOOKUP/INDEX/MATCH/DATE`、跨表、命名区域、嵌套和 `#DIV/0!` 输出一致；输入 B2 从 2 改为 5、A1 公式改成 `SUM(InputValues)*2` 后均得 24。证据：`evidence/formula-reference/excel-com-formula-results.csv` 与 `evidence/formula-reference/betteroffice-formula.csv`。这只覆盖同一合成公式集，不涵盖全部 Excel 函数、真实文件或 10k–100k 公式负载。
- BetterOffice 内存模型上的单次初算为 0.571 ms、输入增量重算 0.036 ms、公式文本改变后的全量重建 0.140 ms；Excel COM 的单次 `CalculateFullRebuild` 为 4.648 ms、编辑后为 0.992 ms。只记录同一函数集的观察值，不构成正式性能排名（单次小样本、初始化/模型装载边界不同）。商用 SDK 的阶段二计算断言与 Excel COM 参照值一致；阶段三公式密集性能仍 `NOT_TESTED`。

### 无修改保存与保真度

- XLSX 检查数值、公式文本、命名范围、合并单元格、批注、超链接、数据验证、图表、行列尺寸、隐藏属性和打印区域等语义字段。保存包部件清单显示两种 .NET SDK 都重写了注释部件名称/关系，但 openpyxl 语义读取仍找到原批注。
- Aspose 评估版无修改保存新增 `Evaluation Warning` 工作表；单元格修改的评估版还新增重复警告表。此为评估许可影响，输出整体保真度只能 `PARTIAL`，且不能用于授权外分发。
- GemBox 无修改保存保留工作表、合并区、图表、数据验证、批注、超链接、公式、命名区域和大部分检查的打印设置；原页眉文本丢失，故 `PARTIAL`。
- XLS 通过实际 SDK 读改存并重开，但 BIFF 记录级结构对拍、外部链接、图片/图表完整性和原始样本损失矩阵尚未完成，XLS 保存保真度 `PARTIAL`。LibXL 没有全结构无修改保存矩阵，标 `NOT_TESTED`。

## 阶段二限制与未测项

本轮夹具是脱敏合成小样本，不包含实际 rustTools 脱敏工作簿、不含 Excel 表格对象、图片、外部链接、复杂命名区域组合和企业级公式样本。没有批量 1000 行以上的 GemBox/LibXL 结果，也没有生产商业许可下移除评估水印后的重测。阶段三需严格按这些边界解释性能结果；能力不完整或授权受限的项不得写成普遍 `PASS`。
