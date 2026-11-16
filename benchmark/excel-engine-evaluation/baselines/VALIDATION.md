# 2026-11-14 独立探针验证记录

主工作区 main，HEAD `34dfcd905b1157dca531908083b156b41f440a03`。既有未跟踪资料保留。本轮只新增 `benchmark/excel-engine-evaluation/baselines/`；没有修改生产源码、根 Cargo 清单/锁文件、路由、旧探针模式或其他工作树。

## 实际运行

- 独立 Rust I/O 适配器：`CARGO_TARGET_DIR=F:\RustBuild\pytools`，`cargo build --release --offline --manifest-path benchmark/excel-engine-evaluation/baselines/adapters/rust/Cargo.toml`，通过。
- 独立 Rust 能力适配器：同 target，`cargo build --release --offline --manifest-path benchmark/excel-engine-evaluation/baselines/adapters/rust/Cargo.toml --bin capabilities`，最终通过（`capabilities-build-policy.exit=0`）。
- 两个独立 .NET 工程：现有 `dotnet-sdk-10.0.404-expanded/dotnet.exe publish`，Release win-x64 self-contained；中间/输出全部在 `F:/RustBuild/pytools/io-*`，通过。没有安装 SDK 或更改系统软件。
- 固定种子 XLSX 生成：`py -3.12 fixtures/generate.py fixtures/generated`，两个文件实际压缩尺寸均在目标 ±5% 内。
- 性能入口：`py -3.12 -u runner.py --manifest fixtures/generated/manifest.json --rust F:/RustBuild/pytools/release/excel-io-baseline.exe --aspose F:/RustBuild/pytools/io-dotnet/IoBaseline.exe --output results/windows-20261114-first`，通过；234/234 请求校验通过，18 组 × 10 个正式观测。完整控制台在 `benchmark-console.log`。
- 能力入口：`py -3.12 -u capabilities.py --rust F:/RustBuild/pytools/release/capabilities.exe --aspose F:/RustBuild/pytools/io-capabilities/Capabilities.exe --output results/capabilities-windows-20261114-verified`，通过。12 个编辑输出的目标修改、其他单元格值/类型/公式及合并区域已另由回归断言核实。发现的保真差异作为能力结果保留，不伪报所有结构通过。
- 评估过滤视图更新：`py -3.12 refresh_views.py results/capabilities-windows-20261114-verified`，仅重读已有输出；不执行引擎、不改工作簿字节。过滤确切新增评估 Sheet、关联边与仅由其引用的共享字符串；用户 Sheet 同时引用的字符串保留。
- 针对性测试：全组在新共享字符串用例加入前为 9/9 通过（`all-targeted-tests-final.log`）；随后只重跑受影响 `test_analysis.py`，8/8 通过（`analysis-sst-final.log`）。合计当前 10 个不同用例已通过，含两个未变化的协议测试；没有运行主项目全量测试。
- 涉及的 Rust 适配器 `cargo fmt --check --manifest-path ...` 通过；Python 入口 `py_compile` 通过；`git diff --check` 通过。没有运行仓库级格式或全量业务回归。

上述简写 Python 相对参数以本基准目录为参照；实际 PowerShell 从仓库根运行时使用完整 `benchmark/excel-engine-evaluation/baselines/` 前缀。真实绝对命令路径及引擎版本也记录在性能 `environment.json`。

## 已解决的开发期失败（保留原始记录）

1. 第一轮协议测试 1 失败：校验器错误地要求 write_only 输出有尺寸声明。已改为完整扫描实际单元格；再次 2/2 通过，后续全组与实际 156 个保存结果也通过。原失败在 `protocol-tests.log` / `.exit`。
2. 第一次能力适配器编译失败：UMYA 3.1.1 的 Sheet 查询返回 Result，不能调用 Option 的 ok_or。已按当前 API 修正并换用非 deprecated API，最终编译和实际能力回归通过。原失败在 `capabilities-build.log` / `.exit`。
3. 复用的 ooxml-patch 库仍发出 7 个既有 unused 类警告；没有为探针清理生产代码。

## 当前发现与限制

- 速度样本不含公式；结果不能推广到公式计算、保真修改、XLS/XLSM 或授权版。
- 4 个 XLS 公式 BIFF 缓存确认为空。Aspose 显式计算得 8、68、6、5；Rust 不计算且拒绝，没有输出，输入保持不变。计算能力差异标 EXPECTED_CAPABILITY_GAP。
- calamine 0.37.0 的该 XLS 公式文本解析有独立异常（错误工作表/坐标），标 UNEXPECTED_PARSER_DIFFERENCE。新实验适配器拒绝全部 XLS 公式；OOXML 数值缓存保留分支尚未单独验收。旧 rxls 实验命令和生产代码均未改，不能推广新保护。
- UMYA 丢失数据验证上下限；Aspose 有多项默认编码/外观表示变化；三条整簿序列化路径均丢失本合成夹具的 customXml 部件。没有据此切换任何生产引擎。ooxml-patch 仅页眉窄操作满足非目标部件哈希不变断言。
- NOT_TESTED：Windows Excel/WPS COM、Linux/WSL/Docker/麒麟、XLS 大文件速度、XLSM、真实业务样本、授权版、复杂公式全集、冷磁盘、并发及峰值内存。

## 安全与交付

`protected-after.json` 显示根 config.xlsx、dist/config/config.xlsx、dist/RustTools.exe 的前后 SHA256 均相同。HEAD 未变，已跟踪文件 diff 为空；没有提交、push、合并、打包或更新 dist，没有操作 feature/test 工作目录。

性能原始 Aspose 输出位于 `results/windows-20261114-first/raw/aspose-*.xlsx`（52 个）。两轮能力输出及最终协议验证输出同样保留，旧能力结果没有覆盖。生成工作簿按本目录 `.gitignore` 排除，实际文件仍在本地。初始最小协议冒烟曾使用自清理临时夹具，后续协议测试已改为持久保留；全部正式性能及能力评估原件一直保留。

最终表格与完整范围见 `results/windows-20261114-first/REPORT.md`，原始计时见 `rounds.jsonl`、`cold-start.json`；结构分析见 `results/capabilities-windows-20261114-verified/report.json`。
