# 三引擎静态打印 Release 基准（2026-11-15）

## 结论

使用与上一轮完全相同的源文件字节，重新构建 Rust 与 Aspose Worker Release，再运行三类夹具 × mode 1/2/3。9 组/27 份真实菜单输出先通过三引擎对拍，并全部与之前 Debug 输出对齐；之后 432 份基准输出逐份通过复核，保留原始水印。

本次复杂样式 rich 夹具中，Rust 剥离启动后的处理链比 Openpyxl 快约 1.6–1.8 倍；普通数据与同 Sheet 覆盖场景两者接近，部分中位数差异仅几个百分点，不作稳定胜出判断。Rust 的真实菜单总耗时在 9 组中均最短，Aspose 在 9 组中均没有速度优势。该结论只覆盖当前小型 XLSX 静态打印；不推广到 Calc、其他业务、大文件、其他格式或平台，不调整默认路由。

## 版本、构建与本轮改动

- 工作树：`C:/Users/AZI/Desktop/VibeCode/Rust-Tools`，`main`，HEAD `34dfcd905b1157dca531908083b156b41f440a03`，起止一致，包含任务开始时所有未提交修改。
- Cargo target：`F:/RustBuild/pytools`。执行 `cargo build --release --offline -p pytools-core --bin RustTools --example print_benchmark`，PASS，日志确认 `release [optimized]`，3m45s。没有复制到 dist，也未打包。
- Rust：`rustc 1.99.1 (58cf8f9c3 2026-10-13)`。主 EXE：`F:/RustBuild/pytools/release/RustTools.exe`；API 驱动：`F:/RustBuild/pytools/release/examples/print_benchmark.exe`。不是旧 dist EXE。
- Aspose：独立 .NET 10 Worker，Aspose.Cells 26.10.0，`dotnet build ... -c Release --no-restore -p:BaseIntermediateOutputPath=F:/RustBuild/pytools/aspose-worker/obj/ -o F:/RustBuild/pytools/aspose-worker-release-print/bin`，PASS，0 warnings / 0 errors，46.28s。属性查询确认 Configuration=Release、Optimize=true；未覆盖旧 Debug Worker 或安装 SDK。
- Python/Openpyxl：继续使用上一轮隔离 CPython 3.12.10、Openpyxl 3.1.5；它们没有 Cargo 的 Debug/Release 构建档。路径 `F:/RustBuild/pytools/debug/python/python.exe` 中的 debug 是 Rust 构建目录名，不表示本轮切换了 Python 解释器；原始运行时信息见 python-runtime.json。
- 本轮唯一已有源码文件编辑是测试脚本 `tools/AsposeWorker/benchmark_print_three.py` 增加 `--canonical-from`，用于复用旧夹具的实际字节。没有改变比较器、业务逻辑、产品代码、C# Worker 源码、计时逻辑、菜单或自动/Linux 路由。
- 新增 Release 证据目录、本文和独立的 `measure_driver_startup.py`（仅额外测量 API 驱动 READY，不处理工作簿）。已记录的其他既有文件 hash 与 BEFORE 全部一致，旧 Debug 报告、失败断言和旧样本保留。

二进制/source hash、完整状态和属性见 `test-runs/print-three-release-20261115/BUILD_INFO.json`、ENVIRONMENT、BEFORE、DELIVERY。不是一次新的提交或发布。

## 同一套夹具与结果门控

| 夹具 | 大小 | 内容 |
| --- | ---: | --- |
| reuse | 5,426 B | 两条配置写同一目标 Sheet；旧合并/尺寸清理；数字 0 与文本“0” |
| rich | 7,833 B | 12×27，含 AA12、合并、RGB 边框/填充、批注和作者、精确行高列宽、日期/布尔/错误/空值/长 ID、独立设定的公式缓存 |
| plain | 9,872 B | 100×10 普通数据，显式列宽 |

从 `F:/RustBuild/pytools/print-three-parity-third-20261115/canonical` 复制原始源文件。Release canonical SHA、旧 canonical SHA、每个引擎/每轮输入源 SHA 全部相同，配置只有运行路径等必要差异。每次检查源和配置未改变，所有输出为新的隔离副本。

真实入口均为 `3 → 1 → 引擎 1/2/3`；配置 mode 1/2/3 与引擎编号分开。Rust 是 Calamine/OOXML 窄读/rust_xlsxwriter 路径，不是保留的旧 `run_umya_engine`。mode 1 full + 行高列宽，mode 2 full + 列宽/默认行高，mode 3 basic；按既定规则处理批注、缓存值、合并、覆盖/多区域堆叠与打印设置。详细既有修复见上一轮 `PRINT_THREE_ENGINE_BENCHMARK_20261115.md`，本轮没有重新修改这些规则。

校验数据值/类型/缓存、合并、被覆盖的样式字段、批注文本/作者、行高列宽、打印区域、方向、A4/fit/居中/页边距、页眉页脚及定义名称。采用上一轮相同的有效默认值/列分组规范化，没有新增忽略规则。不是 ZIP 字节保真或任意 Excel 对象完整性承诺。

Aspose 额外 Sheet 必须带可归属的 Evaluation 标记才从业务比对排除。水印 Sheet 与其活动页变化保留在原始文件中，是用户接受的过渡例外。未去水印、未绕过评估限制，也没有宣称输出干净。时间包含评估输出的实际保存成本；水印不计入有效业务数据量。

已知任务计数差异继续保留：reuse 的 Openpyxl/Aspose task_ok=2（配置行），Rust task_ok=1（目标簿），实际都输出一个文件。本轮输出对齐不等于统计字段已经统一。

## 计时边界与启动

- 先完成 27 输出的 Release 对拍与 27 输出的跨 Debug/Release 对拍，通过后才计时。
- 每个夹具/mode/引擎预热 3 次、正式 API 10 次；轮转执行顺序；另测真实菜单 E2E 3 次。合计 81 预热 + 270 正式 API + 81 菜单输出，432 份均复核通过。构建结束后才开始基准，没有与编译或其他测试并跑。
- Python 使用 perf_counter_ns，Rust 使用 Instant，Worker 使用 Stopwatch。外部 Worker 启动采用本次实际进程创建到 READY 的耗时，包含 Python/openpyxl 导入或 CLR/DLL GetVersion 加载；不减历史常量或无关空程序耗时。
- “剥离后处理链”= 同一次 API 整链减同一次外部 Worker READY 启动。包含配置、读文件、规划、复制、保存、校验、安全提交、协议/关闭等成本，不是 SDK 裸 API；首个文件 API 的 JIT/延迟初始化没有被虚构成已剥离。Worker 内部时间另外存于 SAMPLES。
- Rust API 驱动常驻；生产 Aspose 打印仍按原作用域管理 Worker，没有为了数字而跨菜单复用。正式轮中的子进程是新进程；操作系统文件缓存经过预热，不是整机冷缓存测试。
- 表中 Rust 的 Worker 启动为 0，仅表示无外部 Worker。Rust 驱动本次首次创建到 READY 实测 **1072.5144 ms**，完全保留，未混进每文件 API 处理时间。基准结束后，独立补测相同驱动启动 10 次，中位数 **7.8374 ms**，范围 **6.8365–56.2827 ms**。首次偏慢原因未定位，不能写成“启动近乎零”，也不能把首次值当作每文件固定开销。
- API 驱动启动不是生产主菜单启动的替身：真实菜单列由 RustTools.exe 实际启动到菜单退出直接测得，含全部真实开销，不是把驱动或 Worker 的中位数相加。脚本即时输入，不含人工等待。
- 所有表列分别取中位数，不保证它们相加相等；API 和菜单是独立样本。本机后台负载未完全隔离，min/max 保留；Debug/Release 分时运行，不能把全部跨轮差异归因于优化档。

## Release 全部结果

单位 ms；API n=10，菜单 n=3，预热 n=3 不入表。

| 夹具 | mode | 引擎 | 外部 Worker 启动 | 剥离后处理链 | API 整链 | 真实菜单总耗时 | 处理链 min–max |
| --- | --- | --- | ---: | ---: | ---: | ---: | ---: |
| reuse | 1 | openpyxl | 466.6 | 85.0 | 551.7 | 660.1 | 77.9–98.7 |
| reuse | 1 | rust | 0.0 | 88.0 | 88.0 | 198.3 | 80.9–225.5 |
| reuse | 1 | aspose | 159.3 | 1297.8 | 1463.9 | 1569.6 | 1173.0–1526.1 |
| reuse | 2 | openpyxl | 466.4 | 91.8 | 561.0 | 714.8 | 78.1–110.0 |
| reuse | 2 | rust | 0.0 | 96.1 | 96.1 | 166.7 | 70.9–419.0 |
| reuse | 2 | aspose | 157.6 | 1262.6 | 1416.3 | 1476.1 | 1147.0–1506.5 |
| reuse | 3 | openpyxl | 440.1 | 78.2 | 513.7 | 576.6 | 66.8–506.0 |
| reuse | 3 | rust | 0.0 | 82.1 | 82.1 | 160.9 | 52.7–358.5 |
| reuse | 3 | aspose | 147.9 | 1162.1 | 1311.5 | 1394.0 | 1084.2–1372.5 |
| rich | 1 | openpyxl | 474.5 | 193.5 | 674.7 | 751.8 | 173.7–219.0 |
| rich | 1 | rust | 0.0 | 105.1 | 105.1 | 205.5 | 99.2–123.1 |
| rich | 1 | aspose | 145.2 | 1409.5 | 1553.3 | 1762.3 | 1310.6–1545.6 |
| rich | 2 | openpyxl | 466.6 | 203.9 | 671.2 | 692.0 | 177.1–252.3 |
| rich | 2 | rust | 0.0 | 115.8 | 115.8 | 199.0 | 95.3–202.2 |
| rich | 2 | aspose | 150.9 | 1408.4 | 1560.8 | 1599.2 | 1294.2–1549.2 |
| rich | 3 | openpyxl | 443.8 | 145.8 | 606.6 | 640.3 | 141.4–188.5 |
| rich | 3 | rust | 0.0 | 90.9 | 90.9 | 167.9 | 77.1–181.7 |
| rich | 3 | aspose | 149.2 | 1308.0 | 1456.1 | 1558.2 | 1255.8–2136.9 |
| plain | 1 | openpyxl | 444.7 | 110.1 | 555.1 | 634.7 | 101.8–128.5 |
| plain | 1 | rust | 0.0 | 110.9 | 110.9 | 179.6 | 99.3–163.0 |
| plain | 1 | aspose | 148.2 | 1280.7 | 1427.6 | 1565.2 | 1220.7–1424.0 |
| plain | 2 | openpyxl | 438.4 | 104.5 | 542.7 | 628.2 | 99.0–110.8 |
| plain | 2 | rust | 0.0 | 109.1 | 109.1 | 176.2 | 92.8–360.8 |
| plain | 2 | aspose | 151.3 | 1302.7 | 1460.3 | 1513.7 | 1246.8–1470.6 |
| plain | 3 | openpyxl | 485.8 | 107.3 | 594.1 | 638.6 | 100.1–121.5 |
| plain | 3 | rust | 0.0 | 105.7 | 105.7 | 192.2 | 97.8–475.9 |
| plain | 3 | aspose | 158.5 | 1418.3 | 1573.3 | 1525.3 | 1223.8–1791.1 |

本次中位数的机械排序：剥离后处理链 Rust 4 组、Openpyxl 5 组最低；其中 reuse/plain 多组仅相差几个百分点，应视作接近。真实菜单 Rust 9 组均最低。Aspose 均未胜出；不能以本轮性能建议用它替换全部组合流水线。

## 实际执行与证据

| 检查 | 结果 |
| --- | --- |
| Rust Release 主程序/API 驱动定点构建 | PASS（编译有既有 warnings；未清理） |
| 独立 Aspose Worker Release 构建及 Optimize 查询 | PASS |
| Release 真菜单 9 组 / 27 输出对拍 | PASS |
| 与原 Debug 27 输出交叉对拍 | PASS；同字节源输入 |
| 正式基准及输出保护 | PASS；432 份输出 |
| 额外 Rust API 驱动启动 10 次 | PASS；无工作簿任务、无外部 Worker |
| 测试脚本 AST 与定点 diff 检查 | PASS |
| 产品源码、HEAD、index、config、dist 保护 | PASS，见 DELIVERY |

本轮未运行 cargo test 或全量测试，没有把之前 Debug 的 18 项单测记录冒充 Release 单测通过。本轮 Release 的运行验收证据是上述真实菜单、输出对拍与逐轮基准。没有做新的产品修复。

定点命令（环境均指向独立 Release DLL 与同一隔离 Python）：

```text
cargo build --release --offline -p pytools-core --bin RustTools --example print_benchmark
python tools/AsposeWorker/benchmark_print_three.py --exe F:/RustBuild/pytools/release/RustTools.exe --canonical-from F:/RustBuild/pytools/print-three-parity-third-20261115/canonical --output F:/RustBuild/pytools/print-three-release-parity-20261115
python tools/AsposeWorker/run_print_three_benchmark.py --exe F:/RustBuild/pytools/release/RustTools.exe --driver F:/RustBuild/pytools/release/examples/print_benchmark.exe --parity F:/RustBuild/pytools/print-three-release-parity-20261115 --output F:/RustBuild/pytools/print-three-release-benchmark-20261115
python test-runs/print-three-release-20261115/measure_driver_startup.py
```

Cargo 前设置 `CARGO_TARGET_DIR=F:/RustBuild/pytools`。运行环境：`DOTNET_ROOT=F:/RustBuild/pytools/aspose-worker-sdk-10.0.404`，`RUSTTOOLS_ASPOSE_DOTNET=<DOTNET_ROOT>/dotnet.exe`，`RUSTTOOLS_ASPOSE_WORKER_DLL=F:/RustBuild/pytools/aspose-worker-release-print/bin/AsposeWorker.dll`，`PYTOOLS_OPENPYXL_PYTHON=F:/RustBuild/pytools/debug/python/python.exe`，`PYTOOLS_OPENPYXL_WORKER=<当前仓库>/crates/excel-editor/python/openpyxl_worker.py`，`TEMP/TMP=F:/RustBuild/pytools/aspose-task-temp`。驱动启动补测在基准结束后执行。重跑须换新的隔离输出目录，不能覆盖现有记录。

仓库内 `test-runs/print-three-release-20261115/` 保存 BEFORE、构建日志/exit、BUILD_INFO、CROSS_BUILD_PARITY、PARITY、ENVIRONMENT、SUMMARY.csv/json、SAMPLES.jsonl、启动样本与 DELIVERY。原始逐轮 XLSX、快照、host.jsonl、菜单 stdout/stderr 在：

- `F:/RustBuild/pytools/print-three-release-parity-20261115/`
- `F:/RustBuild/pytools/print-three-release-benchmark-20261115/`

可检查带水印样本：`F:/RustBuild/pytools/print-three-release-benchmark-20261115/cli/rich-m1-aspose-cli-0/target_Aspose.xlsx`，并与相邻的 openpyxl/rust 合成输出对照。未用业务原件做测试。

## 未验证与交付保护

NOT_TESTED：真实业务簿、MiB/GiB 大文件、百万单元格、100 文件批量常驻吞吐、实际打印机/PDF/分页预览、复杂公式重算、全部高级格式/图表/宏/表格保真、XLS/XLSM、WPS/COM/LO、Linux/WSL/Docker/麒麟。本轮静态 XLSX 打印不能推广为这些路径已通过。

config.xlsx、dist/RustTools.exe、dist/config/config.xlsx 与开始 hash 一致；旧 dist EXE 未包含本轮 Release 测试构建。现有源码和修改保留，未更新生产路由，未暂存、提交、push、合并或操作 Feature/Test 工作目录。没有安装软件或发布构建。
