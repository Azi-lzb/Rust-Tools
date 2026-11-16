# Windows 四维纯处理基准：固定 JIT 控制组（2026-11-14）

**没有全面胜出，不支持整体替换组合流水线。** Aspose 快于本样本的 BetterOffice、UMYA 和裸 openpyxl，但 Calamine、rust_xlsxwriter、Excel COM 及 OOXML 页眉窄接口各自更快。修改保留度上，Aspose 比本样本的裸 UMYA/openpyxl 完整，但不等于 OOXML 定点补丁。

本表固定 .NET DOTNET_TieredCompilation=0、DOTNET_TieredPGO=0，Rust release；完整重跑全部组件，没有拼选不同配置的最优值。默认 CLR 配置的独立数据在 ../four-dim-20261114/；同输入更正后的修改项也有默认配置记录。两组方向一致。

每组 3 轮预热＋10 轮正式，单位 ms，中位数。API 区间已剥离所有进程/CLR/Python/Excel 启动。

| 维度/相同负载 | 组合基线 ms | Aspose ms | Aspose/基线耗时比（小于 1 更快） |
| --- | --- | ---: | --- |
| Read：1 MiB 档 | calamine 0.37.0 + 本地修补 35.51 | 79.46 | 2.24× rust |
| Read：10 MiB 档 | calamine 0.37.0 + 本地修补 314.22 | 861.57 | 2.74× rust |
| Write：50k×20 | rust_xlsxwriter 0.101.0 1024.63 | 1805.14 | 1.76× rust |
| Calc：XLS，2048 公式 | Excel COM 16.0 8.79 | 26.62 | 3.03× com |
| Calc：XLSX，2048 公式 | BetterOffice 0.5.1 + 仓库修补 60.79；Excel COM 16.0 8.40 | 23.79 | 0.39× rust；2.83× com |
| Modify：值＋加粗 | UMYA 3.1.1 + 本地修补 566.81；openpyxl 3.1.5 裸 API 1995.19 | 190.38 | 0.34× rust；0.10× openpyxl |
| Modify：页眉窄操作 | ooxml-patch 页眉窄接口 127.32；openpyxl 3.1.5 裸 API 2063.67 | 193.06 | 1.52× rust；0.09× openpyxl |

## 完整轮次范围

| 条件 | 引擎 | 中位数 ms | 最小–最大 ms | 正式有效轮次 |
| --- | --- | ---: | --- | ---: |
| read-1MiB | Aspose.Cells 26.10.0 评估版 | 79.46 | 63.95–87.13 | 10/10 |
| read-1MiB | calamine 0.37.0 + 本地修补 | 35.51 | 31.59–48.93 | 10/10 |
| read-10MiB | Aspose.Cells 26.10.0 评估版 | 861.57 | 845.43–948.52 | 10/10 |
| read-10MiB | calamine 0.37.0 + 本地修补 | 314.22 | 310.89–320.98 | 10/10 |
| write-50kx20 | Aspose.Cells 26.10.0 评估版 | 1805.14 | 1737.77–1905.69 | 10/10 |
| write-50kx20 | rust_xlsxwriter 0.101.0 | 1024.63 | 1000.94–1213.18 | 10/10 |
| calc-xls | Excel COM 16.0 | 8.79 | 8.36–9.79 | 10/10 |
| calc-xls | Aspose.Cells 26.10.0 评估版 | 26.62 | 23.77–39.56 | 10/10 |
| calc-xlsx | Aspose.Cells 26.10.0 评估版 | 23.79 | 22.27–26.09 | 10/10 |
| calc-xlsx | BetterOffice 0.5.1 + 仓库修补 | 60.79 | 60.14–72.55 | 10/10 |
| calc-xlsx | Excel COM 16.0 | 8.40 | 7.97–9.15 | 10/10 |
| modify-value-style | UMYA 3.1.1 + 本地修补 | 566.81 | 538.21–669.82 | 10/10 |
| modify-value-style | openpyxl 3.1.5 裸 API | 1995.19 | 1880.84–2133.99 | 10/10 |
| modify-value-style | Aspose.Cells 26.10.0 评估版 | 190.38 | 179.57–237.90 | 10/10 |
| modify-header-narrow | ooxml-patch 页眉窄接口 | 127.32 | 112.92–162.45 | 10/10 |
| modify-header-narrow | openpyxl 3.1.5 裸 API | 2063.67 | 2001.96–2191.27 | 10/10 |
| modify-header-narrow | Aspose.Cells 26.10.0 评估版 | 193.06 | 180.33–210.80 | 10/10 |

## 启动和计时口径

- 同一 Aspose 二进制只加载 DLL、不处理工作簿并退出的空跑 3 次：中位数 154.87 ms，范围 152.38–891.23 ms。
- 进程 READY 单独计时：rust 126.53 ms；aspose 185.62 ms；openpyxl 664.56 ms；com 2862.52 ms。
- 主表直接采用驻留进程的内部 Stopwatch/Instant/perf_counter_ns 区间，已经排除了创建进程、加载 CLR/DLL/Python 和启动 Excel。没有从已经排除启动的 API 时长再减空跑常数，避免二次扣除。空跑也含进程退出，不拿它当准确的每文件成本。
- 各条件首次实际请求和 3 预热轮完整保留；没有删除较慢的有效轮。固定 JIT 控制组避免默认 CLR 仍在分层优化造成明显过渡，GC、系统后台负载、实际文件 I/O和保存均不估计扣除。
- Read：同一 1,032,702 / 10,295,217 字节 XLSX，完整物化原 Data Sheet，值/类型摘要在计时外。
- Write：同一逻辑 50,000×20 数据，1,000,000 个位置、800,000 非空值；API 内生成整数、文本、布尔、空值，填充并保存。采用逐单元格调用，不是各库批量 API 的最佳极限；未测 bulk import 或空单元格访问的进一步优化。两份实际保存结果均独立完整回读校验。
- Calc：8 类公式 ×256，共2048；每轮新载入同一错误缓存原件，确认全部缓存为 -9999 后计时重算。加载、结果提取、序列化在主计算时间外，load_ms 单列。三者均对独立预期逐值通过，不把任何引擎默认当真值。BetterOffice 通过 Calamine 装入模型的成本明确排除于“纯计算”，不声称它原生支持 XLS。
- Modify：同一 10000 数据行复杂 XLSX，J12=9999＋加粗，或另测页眉；计加载、修改、保存和释放。输入不变、候选唯一，目标不覆盖。

## 更正后的结构保留度

早版未知 OPC 哨兵仅在包级关联，不能证明合法 Workbook 自定义 XML 丢失。现在使用 WorkbookPart→customXml→datastoreProperties 的标准关系，独立确认 Aspose CustomXmlParts 从 0 变为 1，Sheet、styles、chart、table 字节完全未变，并重跑所有修改。原记录保留为诊断，不把旧错误归因继续用作产品结论。
每个修改输出逐文件核查全值/类型/公式摘要、10001 个未受目标影响的公式缓存、合并、批注、链接、表格、验证、命名区域、原显式格式/尺寸、打印、页眉页脚、冻结、自定义 XML 数据和 datastore 属性，同时列出每个 ZIP 部件未压缩 SHA256 与 XML 规范化结果。图表补查类型、标题、系列数、解析后的绑定目标；原始公式写法也保留。
结构语义与原包字节不变分别报告：
- Aspose 两组 26 个输出保留了本夹具所查关键语义，包括合法自定义 XML、图表和公式缓存。Data!B1 写为 Data!$B$1 的图表绑定目标相同，但原始 XML 确实改变；没有把一般工作表公式的 $ 当作可忽略。它仍改写了多个非目标部件，不是精准 OOXML patch。
- UMYA 13 个单元格修改输出丢失图表和合法自定义 XML；上下限修复仍有效。其时间可记录，但不能据此把有损输出当作成功保真修改。
- 裸 openpyxl 两组 26 个输出丢失 10001 个缓存和合法自定义 XML，图表保留。主项目隔离 Worker 另有缓存恢复，裸库结果不能推广成生产 Worker 当前会丢缓存；完整生产封装性能未测。
- ooxml-patch 页眉接口 13/13 保留所查语义，且仅目标 worksheet 内容变，其他全部未压缩部件字节相同；它更快且真正满足部件不变。通用单元格值/加粗并不是该库提供的接口，明确 N/A，未为了评测新增修改器。

## 证据与边界

- 本控制组 221 请求＝17 条件组×13，正式170、预热51；目标值/计算结果校验全部通过。结构有损另列，不将API返回成功当作整簿保真通过。
- Aspose 本组一个驻留进程共91个Workbook操作，95硬停止，没有在达到100后重启。默认组及夹具更正是独立实验配置，不是生产分批规避机制。全部原始评估输出保留，水印只在比较/有效数据量视图排除；保存成本包含实际评估标记生成，不估算授权版性能。
- COM 是独立创建并证明 PID 不在原进程集合中的 Excel 16.0 会话，只读处理合成文件；结束后只关闭自有会话。没有附加用户工作簿或清理用户 Excel。
- 单文件请求端到端时间、加载时间、冷启动、首次调用、预热、正式轮次分别留痕；不把不同负载中位数相加声称完整业务流水线总耗时。
- 新增3项Rust与11项Python针对性测试均通过；不跑主项目全量。正式基准运行后源路径、配置/dist及以前工作簿均进行哈希复核，保护文件见 protection.json。
- 原有 SDK 临时目录在测量后不可用，原因未确认，本轮没有删除它。随后一个可选写入循环优化未能构建，已回退到实际测量的源实现；没有安装/下载软件，也没有给未编译代码编造性能。自包含测量 EXE/DLL 仍在 F 盘。

NOT_TESTED：Linux/WSL/Docker/麒麟、WPS COM、授权版、真实业务样本、宏/XLSM、外部链接、UDF/动态数组/公式全集、更大计算规模、并发/峰值内存、冷磁盘、bulk API 最佳优化、完整生产安全提交流水线。没有改生产路由，不提交、不push、不更新dist。

建议保留 Calamine 读取、rust_xlsxwriter 新建与 OOXML 窄操作。Aspose 可继续作为显式可选的计算/复杂修改适配器评估，但本数据既不是全面性能胜出，也不是完整对象保真认证。

## 默认配置敏感性对照

以下保留默认 CLR 的对应中位数；不与上表挑最小值拼接。默认小文件读取的169.60 ms含分层JIT过渡，控制组79.46 ms更稳定；计算则默认14.94 ms比关闭分层/PGO的23.79 ms快。尽管幅度变化，各维度相对结论一致。

| 条件 | 引擎 | 默认 ms | 固定 JIT 控制 ms |
| --- | --- | ---: | ---: |
| read-1MiB | Aspose.Cells 26.10.0 评估版 | 169.60 | 79.46 |
| read-1MiB | calamine 0.37.0 + 本地修补 | 38.45 | 35.51 |
| read-10MiB | Aspose.Cells 26.10.0 评估版 | 874.65 | 861.57 |
| read-10MiB | calamine 0.37.0 + 本地修补 | 343.93 | 314.22 |
| write-50kx20 | Aspose.Cells 26.10.0 评估版 | 1876.63 | 1805.14 |
| write-50kx20 | rust_xlsxwriter 0.101.0 | 1147.70 | 1024.63 |
| calc-xls | Excel COM 16.0 | 9.16 | 8.79 |
| calc-xls | Aspose.Cells 26.10.0 评估版 | 15.30 | 26.62 |
| calc-xlsx | Aspose.Cells 26.10.0 评估版 | 14.94 | 23.79 |
| calc-xlsx | BetterOffice 0.5.1 + 仓库修补 | 67.96 | 60.79 |
| calc-xlsx | Excel COM 16.0 | 9.32 | 8.40 |
| modify-value-style | UMYA 3.1.1 + 本地修补 | 603.61 | 566.81 |
| modify-value-style | openpyxl 3.1.5 裸 API | 2085.03 | 1995.19 |
| modify-value-style | Aspose.Cells 26.10.0 评估版 | 175.00 | 190.38 |
| modify-header-narrow | ooxml-patch 页眉窄接口 | 119.74 | 127.32 |
| modify-header-narrow | openpyxl 3.1.5 裸 API | 2156.91 | 2063.67 |
| modify-header-narrow | Aspose.Cells 26.10.0 评估版 | 172.16 | 193.06 |

官方依据：[SetFormula 缓存接口](https://reference.aspose.com/cells/net/aspose.cells/cell/setformula/)、[显式重算接口](https://reference.aspose.com/cells/net/aspose.cells/workbook/calculateformula/)、[WorkbookPart 自定义 XML](https://learn.microsoft.com/en-us/visualstudio/vsto/custom-xml-parts-overview)。
