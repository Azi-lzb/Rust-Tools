# Windows 四维纯处理基准（2026-11-14）

结论：Aspose 没有全面胜出。它的计算快于本样本的 BetterOffice、复杂修改快于 UMYA/裸 openpyxl，所查对象保留也较完整；读取、新建写入、Excel COM 计算以及 OOXML 页眉定点修改均有现有组件更快。不能据此整体替换现有组合。

主表每组为 3 轮预热＋10 轮正式计时，单位毫秒，中位数。所有错误/慢轮保留，没有删去较慢的有效样本。数值取自同一 Windows 主机串行实验。

| 维度/负载 | 组合基线中位数 ms | Aspose 中位数 ms | Aspose/基线耗时比（<1 更快） |
| --- | --- | ---: | --- |
| 读取 1 MiB 档 | calamine 0.37.0 + 本地修补：38.45 | 169.60 | 4.41× calamine 0.37.0 + 本地修补 |
| 读取 10 MiB 档 | calamine 0.37.0 + 本地修补：343.93 | 874.65 | 2.54× calamine 0.37.0 + 本地修补 |
| 写入 50k×20 | rust_xlsxwriter 0.101.0：1147.70 | 1876.63 | 1.64× rust_xlsxwriter 0.101.0 |
| 计算 XLS / 2048 公式 | Excel COM 16.0：9.16 | 15.30 | 1.67× Excel COM 16.0 |
| 计算 XLSX / 2048 公式 | BetterOffice 0.5.1 + 仓库修补：67.96；Excel COM 16.0：9.32 | 14.94 | 0.22× BetterOffice 0.5.1 + 仓库修补；1.60× Excel COM 16.0 |
| 修改值＋加粗 / 复杂 XLSX | UMYA 3.1.1 + 本地修补：603.61；openpyxl 3.1.5 裸 API：2085.03 | 175.00 | 0.29× UMYA 3.1.1 + 本地修补；0.08× openpyxl 3.1.5 裸 API |
| 补充：页眉定点修改 | ooxml-patch 页眉窄接口：119.74；openpyxl 3.1.5 裸 API：2156.91 | 172.16 | 1.44× ooxml-patch 页眉窄接口；0.08× openpyxl 3.1.5 裸 API |

修改时间只表示 API 完成及目标值/样式检查通过，不等于结构保留通过。下文单列保留情况，不能把结构有损的输出当作合格胜出。BetterOffice 的 XLS 原格式计算标 UNSUPPORTED；不把先转格式的不同工作量混进本表。ooxml-patch 没有通用单元格值/加粗接口，该任务标 N/A，页眉项才是双方支持的对照。

## 启动分离与计时边界

- Aspose 实际 DLL 加载并退出的空跑共 3 次，中位数 135.60 ms，范围 130.73–137.60 ms。
- 原读写/计算 Worker 启动到 READY：Rust 30.38 ms，Aspose 151.67 ms，Python/openpyxl 797.30 ms，Python＋独立 Excel COM 2417.93 ms。
- 主表取驻留进程内部 Instant / Stopwatch / perf_counter_ns 的 API 区间，已从计时边界排除进程创建、CLR/Python/Excel 初始化；因此不能再把空跑常数减一次。空跑含退出成本，也不等于所有格式代码的 JIT 成本。首次实际调用和 3 轮预热另存原始记录。
- .NET 默认分层 JIT、GC及后台负载造成的运行期波动没有估计扣除；尤其小文件读的范围较宽。完整中位数/最小/最大及全部轮次在 CSV/JSONL 中，不用单轮最快值排名。
- Read：打开同一 XLSX 并物化全部原 Data Sheet 值，摘要/IPC在计时外。1 MiB 档实际 1,032,702 字节，10 MiB 档实际 10,295,217 字节。
- Write：50,000×20，共 1,000,000 个位置、800,000 个非空值；API 内按同一规则生成整数、带零填充文本、布尔、空值，填充工作簿并保存。不是从 JSON 预载矩阵的旧写入口径。
- Calc：每轮从同一错误缓存原件重新加载，先确认 2,048 个缓存都是 -9999，再只计重算 API；文件加载、结果提取、JSON及保存不计计算。BetterOffice 的 Calamine→模型导入在 load_ms 单列，未假称其有原生 XLS 导入。COM 的跨进程 API 调用本身是实际计算调用的一部分，Excel 服务启动已排除。
- Modify：加载同一 10,000 数据行复杂 XLSX，在无公式依赖的 J12 设置 9999 并加粗后保存；页眉另测。均包括加载、修改、保存及必要释放，排除 Python/.NET 启动和外部结构审查。

## 计算验证

夹具含跨表 SUM、SUMIF、COUNTIF、IF/AND/ROUND、VLOOKUP、INDEX/MATCH、IF/OR、SUMPRODUCT 八类公式，各 256 个，共 2,048 个。输入 XLS 与 XLSX 的错误缓存由 xlrd/openpyxl 独立确认；每轮引擎内部也检查计时前仍为错误缓存。
所有计时轮结果都对独立 Python 预期逐值比较，三者均通过；没有因某引擎身份或多数一致就判另一引擎正确。Excel COM 本轮算对且更快，没有排除它。Aspose 显式调用 Workbook.CalculateFormula，COM 调用 CalculateFullRebuild，BetterOffice 调用 rebuild_and_recalc_all。未实现 Rust 公式计算器。

## 修改保留度与夹具纠正

最初自定义 XML 哨兵挂在包级：它只能测试未知 OPC 部件保留，不能证明合法 Workbook 自定义 XML 被丢弃。该输入 Aspose CustomXmlParts.Count 为 0。已保留原始记录，并改用 WorkbookPart 关联、数据存储属性及关系齐全的新夹具，识别数为 1；仅更正元数据，所有 Sheet、styles、图表、表格字节保持原样。修改维度完整重跑，最终主表只使用这次更正后的计时。此前把包级哨兵缺失概括为 Aspose 丢失合法自定义 XML 的结论撤回。
更正后的 78 个输出均逐文件检查：全值/类型/公式摘要，10,001 个不受目标修改影响的公式缓存，合并、批注、超链接、表格、验证、命名区域、显式格式/尺寸、打印、冻结窗格、页眉页脚、Workbook 自定义 XML，以及全部 ZIP 部件未压缩哈希。相同内容可按哈希复用语义分析，仍逐文件校验所有部件；不忽略未知部件来制造通过。

- Aspose：值/样式与页眉两组各 13/13 通过所查语义，包括图表、缓存与合法自定义 XML。它会改写原包多个非目标部件，不能称为 OOXML 定点补丁。
- UMYA：13 个单元格修改输出均丢图表与合法自定义 XML；验证上下限未再丢失。不能因目标值正确就声称整簿保真。
- 裸 openpyxl：两组 26 个输出均丢 10,001 个既有公式缓存及合法自定义 XML；图表保留。注意主项目隔离 Worker 另有缓存恢复逻辑，本组不是该生产封装的验收，不把裸库现象推广成产品当前缺陷。
- ooxml-patch 页眉窄接口：13/13 所查语义通过，只有目标 worksheet 部件变化，其他所有未压缩部件字节哈希相同；更快且满足真正的部件不变要求。
- 图表另查类型、标题、系列数和解析后的绑定目标。Aspose 将 Data!B1 写为 Data!$B$1，目标相同但 XML 表示不同，原始写法保留供审查；没有忽略一般单元格公式的绝对/相对标志。

## 原始证据、测试与运行保护

- 原始实验 221 次请求、更正修改实验 78 次，共 299 次，目标数据/计算结果校验均通过。最终主表选用原读写/计算 143 次＋更正修改 78 次＝221 次，17 组 × 13 轮，其中正式 170 次。早版修改 78 次保留，不用于合法 CustomXml 的结论。
- 原 Aspose 进程 91 个工作簿操作、更正修改独立实验 26 个；各自均未达到 100，没有达到上限后自动重启或重试。全部原始评估输出保留，仅比较视图排除经确认的新增评估 Sheet/独占标记和强制活动页效果。保存开销包含实际评估标记生成，不能推断授权版表现。
- 具体命令、引擎版本/进程号、加载耗时、端到端请求耗时、empty-run、预热与正式轮次、失败/结构差异分别记录。较慢有效轮次未删除。
- 针对性 Rust/Python 计时协议、类型摘要、预期值、错误计数和夹具关联测试另存日志；没有运行主项目全量或发布构建。
- 所有原输入 SHA256 前后核对；COM 使用 DispatchEx 创建、证明 PID 不在此前进程集合中的独立实例，只读打开合成夹具，最后关闭自有实例。未附加用户 Excel 会话。

## 结论范围

这是一组可复现的组件 API 对照，不是把四个不同负载的中位数相加得到的业务总耗时。没有测现有完整候选校验、快照、安全提交及所有生产 Worker 包装，也没有运行“授权 Aspose 完整替代产品”的端到端验证。
可以得出：本样本中 Aspose 适合进一步评估为计算/复杂修改的可选适配器；不支持删除现有 Calamine、rust_xlsxwriter、OOXML 定点能力。替换与否还要按格式、公式覆盖、对象保留、运行依赖和实际业务分别验收。
NOT_TESTED：Linux/WSL/Docker/麒麟、WPS COM、授权版 Aspose、真实业务样本、宏/XLSM、外部链接、动态数组/UDF/公式全集、不同规模计算与修改、冷磁盘、并发/峰值内存、完整生产流水线。没有改生产路由、配置、dist，没有提交或 push。

文件：FINAL_SUMMARY.csv/json 为最终表；run/ 为原始整轮；modify-corrected/ 为最终修改轮次与全部结构差异；fixtures/ 保存原输入及更正夹具；customxml-recognition.json 保存识别验证。旧记录没有覆盖。

官方接口/关系依据：
- https://reference.aspose.com/cells/net/aspose.cells/cell/setformula/ （明确设置错误缓存）
- https://reference.aspose.com/cells/net/aspose.cells/workbook/calculateformula/ （显式重算）
- https://learn.microsoft.com/en-us/visualstudio/vsto/custom-xml-parts-overview （WorkbookPart 自定义 XML）
