# Aspose 菜单扩展与对拍（2026-11-15）

本轮基于 main `34dfcd905b1157dca531908083b156b41f440a03` 的已有脏工作区；未提交、未推送、未更新 dist。原有改动已保存启动快照，本报告只描述本轮增量，不将整个 Git diff 归于本轮。

## 运行边界

- 手动选择 Aspose 才接入新后端；自动/默认和 Linux 路由保持原行为。现有读取仍由 Calamine/窄范围 OOXML，新建 XLSX 仍由 rust_xlsxwriter 完成，业务规则仍在 Rust。
- C# 为独立 .NET 10 工程。Cells/Words 都使用 26.10.0 无许可证评估版；输出警告、额外 Sheet、活动页变化及 Word 评估内容完整保留。比对时仅在内存视图中排除已验证的评估标记，原文件不去水印。
- 生产修改适配器目前只允许 XLSX，明确拒绝 XLS/XLSM 和未验证的宏、外链、透视表、签名、ActiveX、保护或未知扩展。2-3 既有 XLS 追加仍在候选创建和 COM 调用前拒绝，未恢复该能力。
- Cells 转换允许 XLS/XLSX → XLSX；XLS 保真校验范围小于 XLSX，不能由 XLSX 的通过推断 XLS。Words 仅 DOC/DOCX → DOCX，WPS/WPT 明确未启用。没有把可选 Worker 放入 dist。
- 带水印输出仅供本轮隔离验收。未知来源的同名评估 Sheet 不自动删除或认领，跨进程重新输入评估结果可能安全拒绝。

## 接口与生命周期

`crates/excel-aspose` 定义 `CalcEngine::calculate`、`ModifyEngine::modify`、`WorkbookConverterEngine::convert`、`DocumentConverterEngine::convert_document`，统一接收带 run_id 的路径请求，返回候选和计时。未来后端可实现同一 Trait；业务代码不直接启动 .NET。

`AsposeWorkerBackend` 管理 JSONL stdio、启动/健康检查、超时、PID 和请求标识校验。每次 Workbook 加载（含不同导入源和失败尝试）占用配额，99 次后退出；宿主按下一笔请求的加载预算提前轮换，单笔超过预算明确拒绝。Worker 失联不自动重放任务。大型指令使用有大小上限、SHA-256 验证的独占临时侧文件，宿主 RAII 清理。

Worker 的 Modify 只执行 Excel 原语；按计划顺序校验内存状态，整笔操作完成后才 Save。它不计算公式。复制值使用源缓存快照，未缓存公式不会被 SDK 默认的 0 冒充；公式文本、未影响的缓存、默认尺寸和必要字体/完整复制边框由窄范围 OOXML 保留模块处理。结构坐标和类型严格校验，不能安全往返的数字明确拒绝。

最终仍经过原 ExcelEditPlan 校验、业务校验、源快照和安全提交。评估 Sheet 的特殊比较仅在显式 Aspose 路径生效。Word 使用独立正文/页眉页脚/表格/图片校验和通用安全发布原语，不拿 Excel ZIP 校验代替 DOCX 校验。

开发日志保留 run_id、请求/PID/轮换次数、候选、阶段耗时和完整错误链；终端不逐单元格输出。配置修复与其他适用运行记录使用真实执行报告标注引擎。

## 当前可达菜单及引擎

- **1-1 时序宽表汇总**：Calamine + rust_xlsxwriter；既有目标修改原 Openpyxl，可显式选 Aspose。1-2 使用区域、1-3 模板命名区域、1-4 批注汇总继续原读取/新建路径，不伪造修改任务。**1-5 模板校验区域**：原 Windows COM/BetterOffice 或显式 Aspose 计算；模板区域复制保留原专用 OOXML/受保护回退，未宣称全部改用 Aspose。1-6 表头比对、1-7 配置预校验为 Rust 读取/校验；1-8 是当前进程的编辑路由设置。
- **2-1 重复检查/标记、2-2 删除重复、2-3 去重追加**：业务与去重规则不变，修改前增加 Auto/Openpyxl/Aspose 选择；从零新建目标仍用 rust_xlsxwriter。2-4 预校验不需要修改引擎。
- **3-1 按配置打印**：统一 Openpyxl/Rust/Aspose 选择；mode 1/2/3 仍是业务复制模式，不能与菜单编号混淆。**3-2** 保留原 Rust 兼容入口（含原有 UMYA 静态打印实现），**3-3** 仍是预校验。本轮探针的 `3-1/3-2/3-3` case 标签分别表示 **3-1 的 mode 1/2/3**，不是三个子菜单。
- **4-1 文件名机构标准化、4-2 排序前缀、4-3 内容前缀**：原 Rust 文件操作/Calamine 读取；4-6 提交检查是 Rust 读取/检查。**4-4** 在选择文件前选择原路由或 Aspose（Cells/Words 独立）；**4-5 Sheet 改名**增加 Aspose 修改，原 XLS Windows 专用 COM 不变；**4-7 打开修复**增加 Aspose 的隔离另存分支，原窄范围 OOXML 修复步骤保留。
- **5-1 村镇拆分、5-2 机构规范、5-3 外资删行、5-4 页眉单位、5-7 涉农比上月**：规则/计划仍在 Rust，XLSX 原 Openpyxl/定点 OOXML路径保留，增加显式 Aspose。原 XLS 专用 COM 不变，强制 Aspose 不支持时不回退。**5-6 工作表数据提取**新建输出仍 Rust writer，仅追加既有目标时执行所选修改引擎。5-5 地区总分校验只读和新建结果，不强行加修改器。5-8 不可达，未恢复。
- **6-1/2 归档预览、6-3 执行计划、6-4/5 直接复制**：Rust 文件系统，Excel 引擎不适用。
- **7-1 配置初始化/修复**：新建仍 Rust writer，既有配置修复可选 Aspose，保持候选预览与最终确认；7-2 读取配置，7-3/4/5 使用系统打开操作，不适用 Aspose。

## 本轮对拍发现与定点修复

1. 导入字体丢失、默认列宽量化、页面 fit 默认属性省略、公式引号/命名区域词法改变、未缓存公式被误当 0、数值 0 类型识别、同目标 Sheet 重复打印计划覆盖：以真实失败断言驱动修复。没有改变打印模式或去重规则。
2. 1-1 完整复制会在保存时带入相邻边框，现按计划追踪完整复制的源边框；basic 复制仍保留目标边框。配置修复的“创建 Sheet 后写入”按顺序验证，失败整笔不保存。
3. 原 Openpyxl 的 delete_rows 未同步 RowDimension，2-2 合成对拍触发原严格校验拒绝。仅修正删除行的维度索引移动，保留原安全断言；不是引擎切换。
4. 原 COM 的 4-4 将合成源 Calibri/行高15规范为宋体/13.5。用户确认 Aspose 保留源格式，这两类外观差异单列，不复制 COM 改动；其他值、类型、公式、结构仍对拍。
5. 首轮有测试夹具/菜单输入/环境失败：不匹配的 1-1 旧表头被保护拒绝；配置路径应在比对视图中归一化；两个旧菜单输入需补 Auto 选择；故障注入需专用 Python；侧文件哈希初版类型不符。所有失败日志保留，不能因最终通过而删除。
6. 补充异常打印对拍复现：同目标中一条配置缺失 Sheet 时，初版 Aspose 整组失败，原路径仍保留两条有效配置。现在单条规划失败回滚其计划增量、单独计失败，继续有效配置；候选生成/校验/提交失败仍整份目标拒绝发布。原失败输出和统计断言保留。
7. 计算对照暴露原公共 OOXML 复制缺陷：Openpyxl 工作簿的 `r` 前缀可能仅在旧 Sheet 局部声明，新依赖 Sheet 的 `r:id` 未绑定，Excel 无法打开。`append_sheet_entry` 现为新增元素声明独立命名空间；失败单测先复现未知前缀，再复测缺失绑定和祖先冲突绑定两种情况。不是计算规则或 Linux 路由调整。
8. Word 首轮仅因中英文被 COM 拆为多个同样式 run 而判差异。比对工具只合并相邻、格式完全一致的 run，仍检查文字、粗体、斜体、字号、表格和页眉页脚，不改实际 Word 输出。
9. 独立复杂 XLSX 探针核对 10,001 个公式缓存、图表绑定和自定义 XML 通过。公式保留模块初版逐公式扫描全表，已改为一次建立地址索引；同一探针的单次处理从 75,703 ms 降至 4,453 ms，语义检查仍通过。这是单次诊断，不能替代十轮正式基准，也不能推广为所有复杂结构的生产验收。

## 验证与性能口径

原始证据：`test-runs/aspose-full-integration-20261115/`。大型基准输出隔离在 `F:/RustBuild/pytools/aspose-benchmark-*-final/`；每轮命令、终端、响应、输出和摘要保留。最终自动生成的 `DELIVERY.json`、`PERFORMANCE.csv`、`CHANGES.json` 是结果清单，与本报告一起阅读。

续跑/修复后的原始目录另包括 `aspose-benchmark-menus-rest`、`aspose-benchmark-print-fixed`、`aspose-benchmark-fx-menu`、`aspose-benchmark-calc-final-fixed`、`aspose-benchmark-word-verified`，均在同一F盘父目录。合并汇总用有效完整分组，打印用修复后的十轮数据覆盖原组；未删除被中断或失败的原始记录。

- 编译仅 Debug 针对性目标；.NET 无系统安装，Cargo target 固定 F 盘。本轮没有全量测试、Release 或打包。实际编译/测试命令和结果见清单，不把代码存在当通过。
- 常驻处理采用 .NET Stopwatch / Python perf_counter_ns，主控使用 perf_counter_ns。加载解释器/CLR 到 READY 的时间单列；3 轮预热剔除首请求 JIT/延迟加载。纯处理包含打开、修改、Save 和保真后处理，另列 SDK load/transform/save/保真 I/O；不拿裸 Save 与受保护的完整流程直接比较。
- 纯 Worker 栏包含各自 Worker 内部的校验，Rust 主流程最终候选校验和提交在该栏之外；完整安全链的耗时应看真实 CLI 栏。Python 探针和随包运行时均为 CPython 3.12.10 / Openpyxl 3.1.5，产品没有回退到系统 Python。
- CLI 全链路从进程创建计到退出，包括实际 Worker 启动、Rust 读取/计划/校验/提交/日志；3 轮预热 +10 正式轮，交替先后顺序、中位数，输出校验在计时外。OS 磁盘缓存不清空。100 文件批次是真实一次命令，不用单文件乘100估算。
- 1 MiB/10 MiB 夹具实际压缩大小为 1,032,702 /10,295,217 字节，31,500/315,000 个非空业务单元格。每轮源哈希、类型/值及常见样式摘要验证，水印不计有效数据量。该夹具不代表复杂图表/宏文件。
- 计时包含 SDK 实际生成/写入评估标记的成本，没有估算扣除；只有有效数据量与差异视图排除标记。`preserve_defaults_ms` 是历史字段名，本轮表示全部保真 XML 后处理合计，不能解释为仅恢复默认行高。
- 不将一次合成通过泛化为全部格式和真实业务验收。Aspose 全簿另存会改写大量 OOXML，**不等同 OOXMLPatch 的字节级精准修补**；窄页眉操作保留原高效专用路径供用户选择。

## 明确未验证与使用方式

NOT_TESTED：用户真实业务、WPS 独立验收、Linux/WSL/Docker/麒麟接入、FAT/exFAT/共享盘、授权版、去水印、发布包、宏和其他受拒绝结构、共享/数组公式的修改。Word 复杂版式、动态域/跟踪修订/非图片 Shape 未作广泛验收，已知不能验证的结构应安全拒绝。

读取和从零新建路径本轮未重新做 Calamine/rust_xlsxwriter 四维基准，也未重跑 BetterOffice 的计算基准；不引用旧数字作为本轮通过或收益。复杂探针与主流程可接受的功能清点范围分别报告。

`dist/RustTools.exe` 未更新，不能用它验收新模式。本轮使用 F 盘 Debug EXE 与显式 Worker 环境变量；不要把 Debug 探针或评估输出误当正式发布包。应使用本轮带水印样本检查，不能拿业务原件试验。

手工最小步骤：在独立副本上逐个选择修改模式后再选文件；检查结果、公式和必要格式；打印逐一选择 mode 1/2/3，并让两条配置写同名目标 Sheet；验证重复追加/目标冲突失败时原件哈希不变。明确选 Aspose 后遇到拒绝应保留失败信息，不自动降级。2-3 XLS 仍应明确拒绝。

## 本轮最终实测结果

Windows Debug，除100文件批次仅1次实际运行外，以下均3次预热、10次正式计时的中位数。常驻修改是 Worker 内部处理；其余为真实 CLI/菜单全链路，不能跨栏直接相除。

- 约1 MiB 常驻修改：Openpyxl 2,496.16 ms，Aspose 553.04 ms（约4.51倍）。约10 MiB：27,731.55 /5,826.02 ms（约4.76倍）。启动至 READY 单次观测分别725.33 /143.31 ms；首请求延迟加载/JIT由预热剔除，没有再扣旧空跑估计。
- 实际100文件、5-2同一命令批次：原路径75,518.44 ms，Aspose 32,533.39 ms（约2.32倍）。启动计数为原Python修改100次+版本探测1次，Aspose为2个PID、100次任务。计数用透明转发入口，记录成本包含在原路径耗时中。
- 1-5公式校验：原COM 7,286.62 ms，Aspose 1,626.08 ms（约4.48倍）。D2字符串与E2数值缓存逐轮核对，另对全部13轮最终业务汇总文件和统计对拍通过。
- Word DOC→DOCX：原COM 5,349.99 ms，Aspose 2,174.34 ms（约2.46倍）。正文/同格式文本段、表格和页眉页脚对拍通过，原始评估内容保留。
- 小文件逐项完整数据见 `PERFORMANCE.csv`：多数普通修改差距不大。打印mode1原574.71 /Aspose1,430.77 ms；Sheet改名731.76 /1,347.43 ms；窄范围打开修复94.20 /1,369.65 ms。5-4真实菜单页眉259.85 /1,407.42 ms，原定点路径明显更快，CLI auto另列记录。不能用大文件修改收益替代这些具体入口的结论。
- 4-4 XLSX规范另存：原COM 6,923.86 /Aspose1,234.57 ms。仅按用户明确同意的默认字体/行高例外单列COM外观差异，并独立要求Aspose结果与源文件一致；该项不是XLS转换已通过的证据。

结论：保留可选Aspose很有价值，特别是大文件修改、实际批量和含COM冷启动的计算/转换；没有证据支持全局替换。自动路由维持原状。

## 最终验证记录

- `cargo test --offline -p pytools-core --test ... --no-run -j2`：本轮指定集成目标编译通过，具体目标和每次命令见 `rust-build-*.log`；两个独立 .NET 项目编译通过。
- `aspose_modify_cli --include-ignored --test-threads=1`：最终14/14通过（含真实SDK、99配额/导入前轮换、大指令、验证失败/进程丢失不重放、源保护）。
- `aspose_menu_cli --include-ignored --test-threads=1`：最终10/10通过，包含真实100文件计算轮换；`print_menu_cli`最终7/7、`fx_header_menu_cli`最终5/5通过。
- 去重17、外资4、村镇5、机构3、涉农3、Sheet改名6、转换菜单1、Sheet提取15、宽表23、配置9项既有回归通过；相应过滤范围外和门控未运行项见各原始日志，不计入通过。namespace定点单测1项通过，内部涵盖两种绑定环境；Openpyxl行维度脚本6条断言通过。
- 指定菜单和打印mode1/2/3合成对拍通过；缺Sheet混合打印3种mode均为2成功/1失败、有效输出保留。原语脚本11例通过，复杂工作簿探针选定语义通过。
- 定点 `rustfmt --check` 通过，`git diff --check` 通过。`cargo fmt --check` 仍失败于5个既有文件：excel-editor的lib.rs/plan.rs、pytools-core的edit_plan.rs/inplace_edit.rs/lib.rs；本轮新增模块声明顺序已修正，lib.rs仅保留原有engine模块排序差异。没有顺带格式化或清理此前改动。
- 最后仅对模块声明排序做 `cargo check --offline -p pytools-core --lib`，通过；没有为排序重跑全量测试。完整构建仍有此前警告，未把警告清理扩成新任务。

水印样本位于 `test-runs/aspose-full-integration-20261115/samples/`。Excel实际增加并激活`Evaluation Warning`，A5是SDK版权评估文字；Word含评估正文和页脚/图像。样本没有去水印。启动命令须使用F盘Debug EXE并显式设置 `RUSTTOOLS_ASPOSE_DOTNET`、`RUSTTOOLS_ASPOSE_WORKER_DLL`（Word另设`RUSTTOOLS_ASPOSE_WORDS_DLL`）；不要用未更新的dist EXE验证本轮。

`CHANGES.json`按启动/修改前快照记录本轮增量与明确新建文件，混合文件应按该快照审阅，不能将整个工作区diff视为本轮。主分支HEAD未改变，暂存区为空；本地config.xlsx、dist EXE及dist配置SHA-256与启动时完全相同。没有提交、推送、合并或操作两个其他工作树。
