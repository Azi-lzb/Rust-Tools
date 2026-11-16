# 独立引擎探针与纯处理基准计划

状态：`FIRST_WINDOWS_BATCH_COMPLETE`。计划基线：main `34dfcd905b1157dca531908083b156b41f440a03`。
用户已确认先执行 Windows 1 MiB/10 MiB，并随后将三项能力均纳入核心 Goal。执行证据位于 `results/`；本文件是方案，不能单独作为通过证据。

## 1. 隔离与现状

- 扩展当前仓库 `benchmark/excel-engine-evaluation`，不新建另一套平行实验室。
- 既有 `dotnet/aspose/Program.cs --worker` 有 JSON-lines 常驻协议，但当前只实现追加、保存、回读的合并计时，不能直接作为纯读写成绩。
- 保留既有 Worker/性能模式和历史证据，拟新增 `--bench-worker` 模式及分阶段计时。
- 当前 Aspose 探针固定 Aspose.Cells .NET 26.10.0、net10.0。Rust 对照跟随主项目已锁定的 calamine 0.37.0、rust_xlsxwriter 0.101.0；其他版本按实际安装记录。
- Rust 基准适配器使用独立 `[workspace]`，不加入主 Cargo workspace；主程序不引用 Aspose、.NET、hostfxr，也不改菜单、引擎选择和 Linux 路由。
- 旧 Aspose 桥接脚本硬编码 `F:/RustBuild/pytools-test`，本轮不能原样调用。新入口使用 main 的 `F:/RustBuild/pytools`，实验产物放其专用子目录，禁止碰 feature/test 工作目录或缓存。
- .NET SDK/依赖复用现有独立环境，构建输出置于 F 盘实验区；不安装系统软件、不改 dist、不提交、不 push。
- 只用固定种子的合成文件和独立输出目录；输入 SHA256 前后必须一致，不读写真实业务原件或本地 config.xlsx。

## 2. 拟建结构

原拟结构如下；实际实现采用 `adapters/dotnet/` 和 `adapters/dotnet-capabilities/` 独立工程，保持旧 Aspose 探针所有模式不变；评估视图在 `capabilities.py` 中实现。公式缺缓存明确拒绝，不修改旧实验转换器的默认行为：

```text
benchmark/excel-engine-evaluation/
  dotnet/aspose/Program.cs         # 扩展 bench-worker，保留原模式
  baselines/
    README.md
    PLAN.md
    protocol.schema.json          # 请求、计时、许可状态、有效数据量
    runner.py                     # 调度、冷启动、预热、重复、汇总
    adapters/rust/                # 独立 Cargo workspace：读、写、降级、编辑
    adapters/python/worker.py     # openpyxl 常驻对照
    fixtures/generate.py          # 固定种子、大小校准、manifest
    compare/semantic.py           # 类型、公式/缓存、样式及关键结构
    compare/package_diff.py       # 部件增删、未压缩内容哈希、XML差异
    compare/trial_view.py         # 可审计的评估产物比较视图
    tests/                        # 缺缓存拒绝、数据遗漏、评估误过滤等回归
    results/<run-id>/             # JSONL/CSV、环境、汇总及原始输出
```

二进制、缓存与发布输出位于 `F:/RustBuild/pytools` 的专用实验子目录，不引入主程序依赖。

## 3. 计时协议

主方法采用常驻进程内计时，不用空 ConsoleApp 的耗时粗减 Aspose 总耗时。

- 调度端 Python `time.perf_counter_ns()`：记录进程创建至 READY、首次请求、每请求往返。
- .NET `Stopwatch.GetTimestamp()` / Stopwatch：在 API 操作边界分别计时；记录计时频率、SDK版本、进程标识。
- Rust `std::time::Instant`、Python Worker `time.perf_counter_ns()`：采用相同操作边界。
- 冷启动独立采样，报告实际 Aspose Worker 的启动至 READY；区分 CLR/程序集准备和首次引擎请求。READY 不冒称所有 JIT/格式路径都已预热。
- 新进程冷启动不等于冷磁盘缓存，不清空系统文件缓存，也不混入“冷磁盘”结论。
- 每个引擎、输入规模和操作先 3 轮预热，再 10 轮正式测量；首请求和预热原始数据保留，不进入正式统计。
- 同机、同操作串行运行，使用可复现的交错顺序减少缓存/温度偏差，不并行压测引擎。
- 编译、样本生成在计时前结束；记录 CPU、内存、OS、磁盘路径和后台负载，不终止其他用户任务来制造成绩。
- 主结果为进程内 `processing_ms`；外部另报 `launch_to_ready_ms`、`first_request_ms`、`request_e2e_ms`、`verify_ms`，不得相加冒充同一口径。
- 读取必须打开并实际遍历/取出 manifest 中所有有效单元格，不能只计打开句柄。数据提取/物化计入读取，后续摘要计算、JSON传输、差异比较不计入 API 处理时间。
- 纯写入从同一预生成、已在内存的数据集开始，计新建工作簿、填入单元格与保存；数据生成/输入解析不计入纯写入。
- 读写往返单独测真实完整操作，不能把两个不同实验的中位数相加。分列打开/读取、写入/保存及实际往返总时间。
- 实际文件读写和保存结束计入处理时间；不声称已排除磁盘 I/O。记录默认压缩与页缓存条件，不做不同持久化承诺的混排。
- 正式结果保留全部 10 轮，汇总中位数、最小/最大值及有效轮数；错误、超时、异常慢轮不静默删除。

## 4. 首轮：同机 1 MiB / 10 MiB XLSX

默认先在当前 Windows 主机完成四个对象的共同格式对照；Linux 后续单列，不与 Windows 结果混排。

- 生成约 1 MiB、10 MiB 的真实 XLSX，目标是压缩后文件大小，允许约 5% 偏差；记录实际字节数、行列数、有效单元格数、种子和 SHA256。
- 仅包含普通标量、固定工作表，无公式；使用适当混合的数字、文本、布尔和空值，控制字符串熵，禁止 ZIP 填充凑大小。
- 纯读取：calamine、openpyxl read_only（完整消费流）、Aspose；rust_xlsxwriter 标 N/A，不伪造读取成绩。
- 纯写入：rust_xlsxwriter、openpyxl write_only、Aspose；calamine 标 N/A。三个写入器输入同一逻辑矩阵，输出格式均为 XLSX。
- 往返：Aspose、openpyxl 和 calamine+rust_xlsxwriter 的纯值重建管线分别列出；后者不冒称保真转换。
- 所有引擎必须通过有效数据量与类型/值摘要核对后才进入可比结果；失败列具体状态，不发布为成功样本速度。
- 原生库 API 与主项目封装分开标注。原生 openpyxl 的缓存/保真行为不能直接冒充当前生产 Worker（后者还有缓存恢复、候选校验和安全提交）。
- .xls 下一批单列。calamine/Aspose 可以进入 XLS 读取组，openpyxl、rust_xlsxwriter 的 XLS 原格式能力标 UNSUPPORTED/N/A。
- 配对 XLS/XLSX 使用同一逻辑数据，但物理字节数允许不同，不为凑相同大小偷偷改变数据量。BIFF 行列限制单独遵守。

首轮交付分三组计时：读取、新建写入、往返；列出规模、实际字节/有效单元格、引擎及模式、中位数、最小/最大、轮数和校验结果。另列冷启动，不把不同工作量排成统一速度榜。

## 5. 评估产物处理

- 原始 Aspose 输出永久保留，不删除/隐藏评估工作表，不改 SDK、许可证、内部开关或输出以消除限制。
- 只在比较/统计视图中排除已确认的新增评估工作表/标记。须同时对照输入 manifest 与供应商标记，不能只按名字含 Evaluation 就过滤，以免丢失同名用户数据。
- 通过原工作表名/身份选取有效数据，禁止依赖 ActiveSheet 或第一个工作表；评估版会改变活动工作表。
- 有效行、单元格、逻辑数据量不包括评估内容；无法可靠归属的差异记为未知/失败，不一概解释为水印。
- 保存时 SDK 自带水印生成开销已经发生，不能凭空扣除。保留真实 save_ms，注明评估版；物理输出字节含水印另报，不用它充当有效数据量或授权版性能。
- 官方说明单程序最多打开 100 个 Excel 文件。首轮计划控制在 80 次以内的保守 Workbook 操作预算，记录计数；外部读者完成回读校验。触及限制标 BLOCKED_LICENSE 并停止，不通过自动重启、拆分或修改 SDK 绕过。

## 6. 公式降级能力

- 首先核实当前实验 Rust 转换适配器的缓存语义，不能把旧探针的默认写入缓存当作已支持。
- 普通性能样本不调用 CalculateFormula。计算只在独立的能力参照案例中由 Aspose 显式调用，并记录是否成功；不假设任意复杂函数都受支持。
- 构造相同公式的缓存存在、缺失、错误以及复杂跨表/查找/数组等小型合成样例；先核实 XLS 夹具的实际 BIFF 缓存状态。
- Rust 不实现计算。首版对缺失缓存默认明确 UNSUPPORTED 并保留原输入；若后续开放“保留公式文本”，必须能明确表示未知缓存，不能填 0/空值冒充计算结果。
- Aspose 在已确认支持的公式上计算成功而 Rust 不计算，差异记 EXPECTED_CAPABILITY_GAP；丢公式、伪造缓存、原件被改写仍是缺陷。
- 不修改生产 4-4 的 Excel/WPS COM 转换或任何现行公式计算路由。本轮 Rust 降级是独立实验适配器的约定，不代表生产 COM 不支持计算。

## 7. 修改保真与精准修补

- 同一输入分别做 no-op 保存、只改已有数字单元格、只改一项样式；openpyxl、umya、Aspose 各保存独立输出。
- 对全部原工作表核对值及类型、公式文本/缓存、合并、批注、命名区域、行列尺寸、打印区域/页面/页眉页脚等夹具包含的关键结构；不能只看修改目标值。
- ZIP 比较部件增删、每部件未压缩内容 SHA256、预期修改白名单和 XML 语义变化，区分容器压缩差异、无意义 XML 重排与实际数据/结构损坏。
- 原始全量差异和评估归因差异同时报告。水印相关 workbook/rels/content-types 改动不能整部件无条件忽略，否则会掩盖真实损坏。
- Aspose no-op 为其重写行为 Baseline；目标编辑与 no-op 分开比较。只有非目标部件满足相应不变条件才可宣称精准修补，不能以语义检查通过冒充字节级通过。
- 现有 ooxml-patch 主要是页眉、列宽等窄接口，不冒称通用单元格/样式修改器。已有 RustNative 单元格操作可作为定点对照；ooxml-patch 另用双方都支持的页眉/列宽操作对拍，不为测试重构生产引擎。

## 8. 确认门槛及交付

确认后仅实现并编译独立探针/适配器，先执行首轮 Windows XLSX 1 MiB/10 MiB 对照，再交付实测表格和原始记录。必要的针对性测试用于证明计时边界、完整读取、水印误过滤和输入保护。
主 Cargo 清单、生产业务、菜单、现行路由、config.xlsx、业务文件和 dist 不在修改范围；不运行全量业务测试、不打包、不提交或推送。
后续公式降级与修改保真按上述计划推进，无法执行的能力标 NOT_TESTED/UNSUPPORTED/BLOCKED_LICENSE，绝不填估计成绩。

官方依据：
- https://docs.aspose.com/cells/net/evaluation-version-limitations/
- https://docs.aspose.com/cells/net/calculate-formulas/

## 9. 本轮实际完成与方案差异

- Windows XLSX 1 MiB/10 MiB：18 组、234 次请求（54 预热 + 180 正式），全部数据验证通过，52 个 Aspose 正式/预热保存原件均保留。
- 独立 .NET 工程取代扩展旧 Program.cs 的方式，保留历史探针所有模式；没有接入主 Rust 运行时。
- XLS 合成未缓存公式 4 项，显式 Aspose CalculateFormula 对拍通过；Rust 拒绝并保留原件。另发现 calamine XLS 公式文本解析差异，因此本探针拒绝全部 XLS 公式，不能把解析异常归为预期计算缺口。
- no-op、单值、单样式、窄页眉修改 12 个输出完成部件/语义分析；UMYA 数据验证上下限丢失、整簿保存丢失 customXml 等边界如实记录，没有修改生产引擎。
- 当前 10 个不同针对性用例通过（按受影响分组运行）；正式输入、配置及 dist EXE 哈希未变。完整命令、原始失败与修复记录、NOT_TESTED 见 VALIDATION.md。
- 尚未运行的更大公式全集、其他格式速度及跨平台测试不算本轮通过；没有全量业务测试、提交、push 或打包。
