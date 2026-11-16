# 纯处理性能与修改保真基线

状态：已获授权并执行独立 Windows 探针。实际结果以 `results/` 下的环境、逐轮 JSONL、摘要及校验记录为准。

2026-11-14 后续核实：XLS 公式解析与 UMYA 双上下限缺失均已用原始包、两种 XLS 生成器和直接库 API 复现，并通过仓库 vendor 定点修补。当前依赖与历史速度二进制不同；旧速度表、原始输出及失败证据保留，不把旧性能数字推广为修补后的成绩。详细结论见 `results/root-cause-20261114/REPORT.md` 与根 `vendor/EXCEL_LIBRARY_FIXES.md`。

执行方案见 [PLAN.md](PLAN.md)。本目录扩展既有实验室，不接入生产菜单或运行时。
Aspose .NET 保持外部独立评估探针；SDK、运行时、许可证及生成工作簿不提交。
现有历史数据不能作为新计时协议下的实测结果。

## 入口与依赖

- `fixtures/generate.py <output>`：固定种子生成普通标量 XLSX 与共享矩阵，按真实压缩大小校准，不填充 ZIP。
- `adapters/rust/`：独立 Cargo workspace，calamine 0.37.0、rust_xlsxwriter 0.101.0；独立 `capabilities` 二进制使用 umya 3.1.1 和本仓库窄接口 ooxml-patch。均不加入主程序依赖图。
- `adapters/python/worker.py`：原生 openpyxl 对照，读为 read_only、写为 write_only、往返为 normal。不是生产 Worker 的性能或保真结论。
- `adapters/dotnet/`：独立 Aspose.Cells 26.10.0 常驻进程，.NET 10，现有 `dotnet/aspose/Program.cs` 模式完全不变。
- `adapters/dotnet-capabilities/`：独立公式与修改探针，不包含在速度进程或生产程序中。
- `runner.py`：3 次启动采样，实际进程一次驻留，预加载后每个规模/操作 3 轮预热、10 轮正式记录，按轮次轮换引擎顺序，串行运行。
- `compare/semantic.py`：独立全表回读验证类型/值 SHA256、实际尺寸、未归因新增 Sheet；忽略尺寸声明，避免漏读。
- `compare/package_diff.py`：全部 ZIP 部件未压缩内容哈希与单独 XML 规范化结果，不把容器压缩差异当作结构变化。
- `capabilities.py`：合成 XLS 未缓存公式（检查真实 BIFF 字节）和 XLSX no-op/值/样式/页眉修改，保存原始输出及差异。

所有直接 Cargo 命令必须先设置 `$env:CARGO_TARGET_DIR='F:\RustBuild\pytools'`，然后仅构建本目录独立 manifest。依赖离线复用，.NET SDK 使用现有 `%TEMP%/dotnet-sdk-10.0.404-expanded/dotnet.exe`，输出在 F 盘专用目录。不安装 SDK、不运行主项目构建、不更新 dist。

## 计时口径

主表采用 Rust Instant、Python perf_counter_ns、.NET Stopwatch 的进程内 API 总耗时。包括实际文件 I/O、单元格物化/填充与保存，排除进程启动、矩阵预载、摘要、IPC和外部校验。每次请求的端到端时间另存。往返独立实测，不能把读写中位数相加。Rust 往返是值重建，不保留结构；Aspose/openpyxl 往返是加载后重新保存。

操作级 3 次预热不保证 .NET 分层 JIT 此后不再优化，正式慢轮、GC和系统噪声均保留。报告中位数和完整范围，不以单次最快轮排名。进程启动至 READY 是独立指标；首次真正 Workbook 请求还有延迟初始化，原始预热记录保留。不是冷磁盘测试，不冲刷 OS 页缓存，不含 fsync 持久化保证。

## 评估与能力边界

Aspose 原始输出一直保留。仅对照原输入 Sheet 身份与确切供应商标记，在比较/有效数据量视图中排除新增评估 Sheet；不删除水印、不改评估限制。保存耗时包含评估版生成水印的实际工作，不能推算扣除，也不能推广为授权版成绩。性能驻留进程总共 78 个 Workbook 请求，硬门控最多 90 个，不在达到限制后自动重启。

Rust 公式探针不计算。实际发现 calamine 对合成 XLS 的公式文本解码不符，因此本探针拒绝所有 XLS 公式；缺缓存时同样明确 UNSUPPORTED。该解析差异单列 UNEXPECTED_PARSER_DIFFERENCE，不能当作“预期计算缺口”掩盖。OOXML 空/歧义/非数值缓存保守拒绝；数值缓存保留分支本轮未单独验收。旧 rxls 实验命令未改，不能将新适配器的安全保护推广到旧命令，也不承诺完整 XLS 结构转换。绝不影响生产 COM 路由。Aspose 参照显式调用 CalculateFormula，不把打开/保存视为自动计算保证。

修改测试包含不透明自定义 OOXML 部件，用于暴露整簿序列化的丢失行为；未知部件差异不能被“水印过滤”消除。原始差异、评估视图、相对同引擎 no-op 的差异并列保留。页眉作为 ooxml-patch 真正支持的窄操作基准，不冒充通用单元格编辑能力。
