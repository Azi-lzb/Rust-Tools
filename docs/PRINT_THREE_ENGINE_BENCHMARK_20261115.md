# 三引擎静态打印 mode 1/2/3 对齐与基准（2026-11-15）

## 结论与范围

本轮按用户确认，以当前 Openpyxl 输出规则修复 Rust 路径，再比较 Openpyxl、Rust 和 Aspose 的配置打印。9 个夹具/mode 组合的 27 份真实菜单输出全部通过语义对拍；正式基准另外逐份验证 432 份输出，无新增差异。仅允许排除可归属的评估水印 Sheet；原始评估文件未删除或改写水印。

剥离 Worker 启动后，Openpyxl 在本轮 9 组中的处理链中位数均最短；真实菜单总耗时则是 Rust 在 9 组中均最短。Aspose 在本轮静态打印夹具上没有性能优势。不能据此断言所有业务、Calc、其他格式或 Release 的性能排序。

基线 HEAD：`34dfcd905b1157dca531908083b156b41f440a03`，分支 `main`，起止 HEAD 相同。既有未提交/未跟踪修改保留。没有暂存、提交、push、合并、更新 dist 或操作其他工作树。

本机 Windows 11；Rust 和 Worker 均为当前 Debug 构建；Python 3.12.10、Openpyxl 3.1.5、Aspose.Cells 26.10.0。不是 Release 评测，没有全量测试或打包。Windows 宿主后台负载没有完全隔离，波动较大的 plain 组应结合 min/max 阅读。

## 实际入口与技术路径

- 真实菜单：`3 → 1 → 执行引擎 1/2/3`，分别是 Openpyxl、Rust、Aspose。3-2 是 Rust 的兼容入口。配置表中的 mode 1/2/3 与执行引擎编号是两个维度。
- Openpyxl：Rust 加载任务配置 → 隔离 Python/openpyxl 脚本（每批一次进程；源簿公式/缓存两种视图）→ 静态值与样式复制 → 设置页面/页脚 → 新目标保存。多条配置复用目标 Sheet 时清理后覆盖。
- Rust：`run_rust_print_engine → run_rust_writer_engine → build_config_print_plans → write_static_print_plans`，Calamine 读值/缓存 + 窄范围 OOXML 读样式/批注/尺寸 + rust_xlsxwriter 新建 + 定点类型/尺寸补丁 + 无覆盖发布。
- Aspose：Rust 规划 → 候选与源快照保护 → 独立 Aspose.Cells Worker 的导入/样式操作 → 原始评估输出 → Rust 校验与安全提交。
- `run_umya_engine` 是保留的旧库函数，本轮真实菜单未调用。不能把它与当前 Rust 路径混为一谈。
- 均为静态打印 XLSX 文件生成，使用已有公式缓存，不执行复杂公式计算，也未访问打印机、PDF、COM 或 LO。自动路由、Linux 路由和主菜单均未修改。

## 经用户确认后的业务规则

| 配置 mode | 值/合并 | 样式 | 行高 | 列宽 | 批注 |
| --- | --- | --- | --- | --- | --- |
| 1 | 静态缓存值；完整落在范围内的合并 | full | 复制 | 复制 | 按“不输出批注”配置 |
| 2 | 同上 | full | 默认行高，不复制 | 复制源列宽，不自动适配 | 按配置 |
| 3 | 同上 | basic：字体/填充/对齐/数字格式 | 不复制 | 不复制 | 不复制 |

同一个目标 Sheet 的后续配置覆盖前一条；同一配置内部的多个区域继续堆叠并保留分隔空行。打印区域、A4、方向、fit、居中、页边距和生成的页脚参与对拍。未指定方向时 Rust 对齐默认纵向。没有改变默认引擎。

## 本轮修复与既有修改的界限

1. `pytools-core/src/print_ops.rs`：撤销同 Sheet 上被后续配置取代的旧导入及尺寸操作；按规范化后的名称判断覆盖；mode 2 去掉旧自适应尺寸；按配置保留 mode 1/2 批注和作者，修正 A1 到零基坐标偏移；保留边框 RGB；写入真正的 Excel 错误值；在安全发布之前恢复精确行高列宽，避免 writer 的像素量化；对齐默认方向与页边距。尺寸补丁仅处理本 writer 新建的临时簿，不读改真实业务原件。
2. `pytools-core/src/summary_ops.rs`：被复用的错误值补丁在转换单元格类型时保留现有 `s` 样式索引；新增回归。此文件此前已有其他修改，本轮增量由 prior 快照单独核对，不能把 HEAD 的全部 diff 都归到本轮。
3. `pytools-core/src/static_print_engines.rs`：仅在 `PYTOOLS_PRINT_BENCH_TIMING=1` 时启用真实 Python READY 与内部处理计时；未设置时仍走原调用方式。此文件包含此前的 Aspose 接入，不应整体视为本轮新增。
4. `tools/AsposeWorker/FontPreservation.cs`：复制/导入时额外精确保留源 fill 声明，并处理混合计划中的显式 fill 修改；修复 SDK 补出 inactive bgColor 导致严格校验拒绝的问题，未放松校验器。文件本来就处于未跟踪的 tools 目录；未取得本轮修改前的逐字快照，因此不声称其完整文件 diff 均由本轮产生。本轮实际编辑限于 fill 分量的调用、style_mode 判断及显式 fill 设置分支。
5. 新增 `pytools-core/examples/print_benchmark.rs`：测试专用常驻 Rust API 驱动，不添加生产入口。
6. 新增 `tools/AsposeWorker/benchmark_print_three.py` 与 `run_print_three_benchmark.py`：合成夹具、逐份语义校验、真实菜单与 API 独立计时、原始日志/水印输出留存。
7. 本报告与 `test-runs/print-three-engines-20261115/` 为本轮证据。`print_menu_cli.rs` 相对本轮 prior 快照没有改动；其 Git dirty 状态来自此前工作。

`prior/` 保存了 print_ops/static_print_engines/print_menu_cli/PlanOperations 的起点副本。summary_ops 在本轮首次编辑前另行保存。四份 `*.task.diff` 用于区分本轮与此前增量；不是可直接整体暂存当前 dirty 文件的建议。`.gitignore`、Cargo 文件、其他业务模块和已有未跟踪目录未作清理或提交。

## 对拍夹具与判定

| 夹具 | 原始 XLSX 大小 | 内容 |
| --- | --- | --- |
| reuse | 5,426 B | 两条配置复用同名目标 Sheet；第一条有合并/尺寸，第二条为独立单元格和数字 0/文本“0” |
| rich | 7,833 B | 12 行 × 27 列（AA12）；合并、边框 RGB、实色填充、32/28 行高、23.125 列宽、带作者批注、日期/布尔/错误/空值/长 ID/前导零、公式缓存 |
| plain | 9,872 B | 100 行 × 10 列，普通文本/数值，显式列宽 |

每组先生成一份 canonical source，复制同样的字节给三引擎；正式样本中的 source SHA 在同一夹具内完全一致。源文件和配置逐次校验 SHA 不变，原有输出不覆盖。公式缓存由夹具独立写入（SUM 结果 15.625），不是由被测引擎计算。

检查值和数据类型、缓存、合并、可见样式字段、批注文本及作者、精确行高列宽、打印区域、方向/fit/A4/居中/页边距、页眉页脚和定义名称。不是 ZIP 字节一致性，也未证明所有 Excel 对象或全部字体特性保真。

比较器只规范化明确的表示差异：标准测试主题下默认字体颜色，省略的默认对齐/fit，inactive 实色填充背景和无边框时的颜色，以及 `<col min/max>` 分组。原始快照和原始 XLSX 全保留。日期缓存用原始 XLSX 类型直接比较，避免 JSON 序列化误判。评估 Sheet 必须同时存在可归属的 Aspose Evaluation 标记才排除；水印导致的活动页变化是用户已接受的过渡例外，不声称输出无水印或 active tab 已恢复。

仍有既有统计口径差异：reuse 的 Openpyxl/Aspose `task_ok=2`，Rust `task_ok=1`，实际输出均为 1 个工作簿。前者按配置行、后者按目标簿计数。本轮未改统计接口；输出语义对齐结论不等于统计字段完全一致。

## 性能计时边界

- 每个夹具/mode/引擎：3 轮预热不计入表，10 轮 API 正式计时；轮转引擎执行顺序。另测 3 次真正生产菜单的启动到退出总耗时（脚本即时输入，不含人等待）。
- Rust `Instant` 记录从真实配置打印 API 进入到最终返回，包含配置、读取、规划、生成、保存、必要校验、提交、协议和清理成本。Python `perf_counter_ns` 测外层请求与菜单壁钟，C# 使用 Stopwatch。
- Python 启动 = 创建实际进程到 import/openpyxl 完成并发出 READY；Aspose 启动 = 实际宿主已记录的创建进程/CLR/DLL GetVersion 加载到 READY。没有减去历史固定的 154.87 ms 或其他空程序估值。
- 表中“剥离后处理链” = 同一次 API 的 `pipeline_ms - worker_startup_ms`。它不是 SDK 裸 API，也不是只计 save；含首个文件 API 的 JIT/延迟初始化与所有安全处理。Python/.NET 内部 processing 另存于 SAMPLES，不与 Rust 整链混作同边界。
- Rust API 驱动提前 READY，其本次独立启动实际测得 **9.4511 ms**；Rust 行 Worker 启动为 0，含义是没有外部 Worker，不代表 Rust 进程启动为零。驱动启动不计入每次 API；真实菜单列已包含自己的实际进程启动，不能再叠加驱动启动。
- Rust 测试驱动常驻不代表 Aspose 跨菜单常驻。本轮保留真实打印作用域，每次独立配置调用创建/关闭其 Worker；未擅改生产生命周期来改善数字。操作系统文件缓存经预热，不能称为机器冷缓存测试。
- 以下各列分别取中位数，不保证“启动中位数 + 处理中位数 = 总链中位数”。真实菜单是独立 3 次测量，也不能由其他列相加推得；部分 plain 组波动较大。

## 全部性能结果

单位 ms。API n=10；菜单 n=3；预热 n=3（不计入）。全部样本通过输出复核。

| 夹具 | mode | 引擎 | Worker 启动 | 剥离后处理链 | API 整链 | 真实菜单 E2E | 处理链 min–max |
| --- | --- | --- | ---: | ---: | ---: | ---: | ---: |
| reuse | 1 | openpyxl | 361.9 | 72.3 | 434.7 | 481.2 | 67.9–111.7 |
| reuse | 1 | rust | 0.0 | 111.7 | 111.7 | 192.6 | 107.3–189.2 |
| reuse | 1 | aspose | 127.0 | 1100.2 | 1227.3 | 1309.9 | 1058.7–1182.1 |
| reuse | 2 | openpyxl | 335.5 | 67.4 | 402.4 | 489.6 | 62.7–83.6 |
| reuse | 2 | rust | 0.0 | 107.7 | 107.7 | 191.5 | 88.1–186.1 |
| reuse | 2 | aspose | 123.3 | 1067.9 | 1189.4 | 1304.4 | 1048.0–1209.9 |
| reuse | 3 | openpyxl | 332.6 | 66.7 | 399.6 | 493.6 | 64.8–75.9 |
| reuse | 3 | rust | 0.0 | 107.1 | 107.1 | 201.5 | 87.3–132.6 |
| reuse | 3 | aspose | 122.3 | 1070.7 | 1193.5 | 1307.4 | 1036.8–1130.4 |
| rich | 1 | openpyxl | 331.6 | 162.2 | 492.8 | 572.4 | 158.5–185.1 |
| rich | 1 | rust | 0.0 | 232.1 | 232.1 | 329.3 | 226.9–239.1 |
| rich | 1 | aspose | 126.2 | 1518.9 | 1643.7 | 1729.3 | 1466.8–1720.8 |
| rich | 2 | openpyxl | 340.0 | 162.2 | 500.4 | 596.6 | 157.3–176.9 |
| rich | 2 | rust | 0.0 | 233.3 | 233.3 | 470.3 | 225.7–503.5 |
| rich | 2 | aspose | 122.5 | 1514.4 | 1637.6 | 1932.4 | 1480.3–1744.1 |
| rich | 3 | openpyxl | 340.8 | 126.1 | 466.4 | 568.1 | 123.9–128.2 |
| rich | 3 | rust | 0.0 | 201.7 | 201.7 | 391.8 | 193.6–217.4 |
| rich | 3 | aspose | 122.8 | 1470.9 | 1603.3 | 1815.4 | 1421.5–2476.0 |
| plain | 1 | openpyxl | 440.8 | 142.4 | 685.8 | 874.1 | 99.6–367.9 |
| plain | 1 | rust | 0.0 | 487.9 | 487.9 | 573.6 | 353.7–1104.4 |
| plain | 1 | aspose | 237.8 | 3302.4 | 3644.6 | 3437.7 | 2117.4–5306.0 |
| plain | 2 | openpyxl | 619.7 | 148.1 | 766.2 | 699.3 | 119.2–177.0 |
| plain | 2 | rust | 0.0 | 484.4 | 484.4 | 549.9 | 412.5–912.0 |
| plain | 2 | aspose | 324.1 | 2789.2 | 3253.1 | 2751.8 | 2530.8–3127.9 |
| plain | 3 | openpyxl | 520.5 | 125.5 | 651.6 | 655.9 | 102.0–161.8 |
| plain | 3 | rust | 0.0 | 409.7 | 409.7 | 458.0 | 348.3–610.6 |
| plain | 3 | aspose | 177.2 | 2400.3 | 2617.1 | 2385.6 | 2085.2–3033.5 |

静态打印中 Aspose 没有性能提升，不能由本轮数据建议替换全部原引擎。Rust 的小文件菜单延迟有优势，Openpyxl 的剥离后处理更短；保持现有显式选择和自动路由。大文件、Release、真实业务以及 Calc 应另行验证。

## 测试命令与实际结果

直接 Cargo 命令均设置 `CARGO_TARGET_DIR=F:/RustBuild/pytools`；临时目录位于 F 盘合成测试目录。

| 实际执行 | 结果 |
| --- | --- |
| `cargo build --offline -p pytools-core --bin RustTools --example print_benchmark` | PASS（Debug） |
| 独立 AsposeWorker .NET 10 Debug 定点构建，输出到 F 盘 | PASS，0 warnings / 0 errors；没有更新 dist |
| `cargo test --offline -p pytools-core --lib print_ops::tests -- --nocapture` | 4 passed；含新增精确尺寸及缺失尺寸拒绝断言 |
| `cargo test --offline -p pytools-core --lib patch_error_cells_tests -- --nocapture` | 2 passed；含新增样式保留与已有输出保护 |
| `cargo test --offline -p pytools-core --test print_menu_cli -- --nocapture` | 7 passed |
| `cargo test --offline -p pytools-core --test aspose_modify_cli real_worker_modifies_value_fill_formula_and_reuses_batch_process -- --ignored --nocapture` | 1 passed，真实 Worker |
| `cargo test --offline -p pytools-core --test aspose_modify_cli real_worker_structural_plan_primitives_validate_before_publish -- --ignored --nocapture` | 1 passed，真实 Worker |
| `cargo test --offline -p pytools-core --test print_config_new_target -- --nocapture` | 3 passed；旧计划/UMYA 候选专项，非本轮性能中的 Rust 引擎 |
| `benchmark_print_three.py --output F:/RustBuild/pytools/print-three-parity-third-20261115` | 9 组/27 输出 ALIGNED |
| `run_print_three_benchmark.py --output F:/RustBuild/pytools/print-three-benchmark-final-20261115 --parity F:/RustBuild/pytools/print-three-parity-third-20261115` | PASS；351 API 输出（81 预热+270 正式）+81 菜单输出 |
| `rustfmt --check --edition 2021`（print_ops、summary_ops、static_print_engines、print_benchmark 四文件） | PASS |
| 两个新增 Python 脚本 `ast.parse` | PASS |
| 涉及 Rust 文件 `git diff --check` | PASS；只有 Git CRLF 提示 |

本轮针对性 Rust 测试合计 **18 passed，0 failed**。编译日志存在已有 warnings；未做无关 warning 清理。没有运行全量或仓库级 cargo fmt，也不声称消除了其他历史格式差异。

早期失败保留：parity-first/second 的真实 Rust 行为差异及比对器边框空侧错误；rich Aspose inactive bgColor 严格校验失败；首次正式 benchmark 中 JSON 日期缓存字符串导致 Openpyxl 对自身误判。它们不计入最终性能汇总。parity-third 和 benchmark-final 是最终通过记录。未靠忽略实际尺寸、错误值、边框或批注差异使断言变绿。

## 证据与复现

仓库内：`test-runs/print-three-engines-20261115/` 包含起点、prior、本轮 diff、编译/回归/失败日志、PARITY、SUMMARY.csv/json、ENVIRONMENT 与 DELIVERY。原始逐轮 `SAMPLES.jsonl`、host.jsonl、菜单 stdout/stderr、每份带水印输出/快照在：

- `F:/RustBuild/pytools/print-three-parity-third-20261115/`
- `F:/RustBuild/pytools/print-three-benchmark-final-20261115/`

可人工打开：`F:/RustBuild/pytools/print-three-benchmark-final-20261115/cli/rich-m1-aspose-cli-0/target_Aspose.xlsx`，保留 `Evaluation Warning`。对照同一目录层级的 openpyxl/rust 样本，查看批注、AA12、G2 错误类型、边框颜色、32/28 行高与 23.125 列宽、打印预览。仅检查合成副本。

运行脚本前需显式设置：`DOTNET_ROOT=F:/RustBuild/pytools/aspose-worker-sdk-10.0.404`、`RUSTTOOLS_ASPOSE_DOTNET=<DOTNET_ROOT>/dotnet.exe`、`RUSTTOOLS_ASPOSE_WORKER_DLL=F:/RustBuild/pytools/aspose-worker/bin/AsposeWorker.dll`、`PYTOOLS_OPENPYXL_PYTHON=F:/RustBuild/pytools/debug/python/python.exe`、`PYTOOLS_OPENPYXL_WORKER=<当前仓库>/crates/excel-editor/python/openpyxl_worker.py`、`TEMP/TMP=F:/RustBuild/pytools/aspose-task-temp`。两个脚本使用支持 openpyxl 的 Python 运行，`--output` 必须为新的隔离目录。不得使用业务原件。

## 未验证与保护状态

NOT_TESTED：Windows Excel/WPS 图形界面实际分页/打印/预览、打印机与 PDF、真实业务簿、大型文件与吞吐、Release、XLS/XLSM、Linux/WSL/Docker/麒麟、COM/LO、复杂公式重算、任意图表/表格/宏/全部高级格式保真；没有由本轮推断这些路径通过。旧 UMya 打印候选的 3 项专项通过不等于其菜单启用，也不构成该路径的性能数据。

保护 hash：根 config.xlsx、dist/RustTools.exe、dist/config/config.xlsx 与任务开始一致，完整 hash 见 DELIVERY.json。dist 旧 EXE 未包含本轮修复，不能用于验收本轮新 Rust 打印保护。未恢复 XLS 追加，也未调整它或任何 Linux 引擎路由。没有提交、push 或修改 Feature/Test 工作目录。
