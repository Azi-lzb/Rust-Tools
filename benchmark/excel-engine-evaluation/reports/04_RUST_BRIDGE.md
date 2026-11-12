# 阶段四：Rust 桥接实测

状态：LibXL C ABI / Rust FFI、Aspose .NET 常驻 Worker / Rust IPC、Rust→openpyxl 路径式批量 IPC 原型均已运行。桥接额外时间的比较仅在对应原型内成立；C++、.NET、Python 的端到端耗时来自不同文件/环境/接口，不能横向评选。

## A. Rust → C++ C ABI → LibXL

在 Rocky Linux 8.9 / glibc 2.28 的离线容器，使用 LibXL C++ 5.2.0 evaluation SDK 编译 `libxlbridge.so` 和 Rust Release 客户端。Rust 只使用稳定的 `extern "C"` 接口，不接触 LibXL C++ 对象布局。C ABI 用不透明句柄承载工作簿，导出 open、批量 append、save、reopen verification、close、last_error；每个入口捕获 C++ 异常并返回错误码。UTF-8 路径/文本由 Rust `CString` 传入，LibXL 由 wrapper 释放。

实测：Rust→C ABI 打开 XLSX、批量追加 5 行、保存到独立输出、释放后重开并检查最后一行，`PASS`，约 293 ms（含 SDK 打开/保存/重开），输出 12,349 bytes。原始记录：`evidence/bridge/libxl-ffi/libxl-rust-ffi-runtime.log`；保留输出：`samples/bridge/libxl-ffi/libxl-ffi-output.xlsx`。C++/Rust ELF 最大需求为 GLIBC_2.28，运行在无网络 Rocky 容器成功。

调试时首轮回读失败，发现 LibXL 的 `lastRow()` 是已用行数（exclusive end），而写入 API 以 0 为首行。桥接层加入边界转换后回读通过。这是重要实际证据：坐标映射应集中在适配器，不应泄漏到业务层。此最小 wrapper 未实测线程并发、安全取消、样式/删除/打印/公式全部接口，也未对比原生 C++ 对应操作成本。评估版只测试 5 行，不能据此推断正式版大批量行为。

## B. Rust → 常驻 Aspose .NET Worker

`dotnet/aspose/Program.cs --worker` 是 .NET 10 Linux x64 目录式 self-contained 发布。Rust Release 客户端启动一个 Worker，等 `ping` 握手后在同一进程内发送两个 JSON-lines 批量请求（每个请求包含文件路径和整批行数据），Worker 打开输入副本、追加、写到兄弟临时候选、重开验证最后追加值，再提交到独立输出。输入与输出路径相同时明确拒绝。worker stdout 只传协议响应，错误回传 JSON。

| 数据量 | JSON 字节 | Rust JSON 构造 | 端到端 | .NET worker 内部（载入/追加/保存/回读） | 桥接及 JSON 解析估值 | 输出字节 |
|---:|---:|---:|---:|---:|---:|---:|
| 10,000 行 | 197,914 | 1.520 ms | 1,745.907 ms | 1,729.925 ms | 15.982 ms | 189,950 |
| 100,000 行 | 2,177,916 | 13.608 ms | 2,504.959 ms | 2,414.276 ms | 90.683 ms | 1,783,620 |

一次冷启动握手往返 1,187 ms（进程创建 12.6 ms；其余主要为 .NET 初始化/握手）。额外批次开销分别约 0.9% 和 3.6% 的端到端时间。数值只适用于本次 Rocky 容器、Aspose 评估版、输入夹具和单任务序列。原始 CSV/日志：`evidence/bridge/aspose-dotnet-worker/`；保存工作簿：`samples/bridge/aspose-dotnet-worker/`。

实测证明 self-contained .NET Worker 可以由 Rust 常驻调度并执行多次路径/批量数据任务，但不证明 Native AOT、hostfxr 嵌入、GemBox Worker 或生产 deb 完成。`.NET Self-contained` 作为独立进程运行，不等于可直接作为 Rust 动态库嵌入。

## C. Rust → Python/openpyxl Worker

仅只读检查当前生产路由：根 `pytools-core/MIGRATION_INVENTORY.md` 的 2026-11-11 记录说明当前 Openpyxl worker 是隔离进程 + JSON 文件交换（不是 FFI）；每个 ExcelEditPlan 为一个批次/工作簿执行一次 Python 修改和安全验证，工作簿只加载一次。代码入口为 `crates/excel-editor/src/openpyxl_backend.rs`，进程隔离/清理为 `crates/excel-editor/src/worker_process.rs`。未修改或运行生产菜单代码。

在本地 Windows 开发机用一个独立实验 Rust 客户端与 `bridge/python/worker.py` 复刻批量路径协议：Rust 生成一个完整 JSON 文件，启动隔离 Python 进程；openpyxl 在候选副本写入并 `os.replace`，随后回读末行。10k 行：约 1,315 ms 端到端 / 1,131 ms Worker 内部；100k 行：约 6,689 ms / 6,376 ms。JSON 构造加落盘 2.2/21.1 ms。Python 3.12.10、openpyxl 3.1.5；结果不代表 Linux 离线 Python 包体或目标机性能，也不能与 Rocky .NET 用来排名。CSV、日志和 XLSX 在 `evidence/bridge/openpyxl/`、`samples/bridge/openpyxl/`。

当前 Python 路线已有真实生产集成和更丰富的安全校验；此原型只验证文件式批处理可行性。常驻 Python Worker、生产 Rust→Python 同机端到端时间与官方随包 Python 运行时启动成本 `NOT_TESTED`。

## IPC 与协议

完整统一请求字段、坐标和值语义、操作顺序、能力拒绝、原子输出和错误对象见 `../bridge/protocol/README.md` 与 `../bridge/protocol/excel-task-v1.schema.json`。当前三个原型都只实现协议子集。批量数据均一次性传整个请求，未逐单元格发送 IPC。

## 结果与适用面

- LibXL 通过 C ABI 的句柄式边界运行；商业 C++ ABI 留在 C++ 内，Rust 只接触 C ABI。对可靠性要求高时，这是 C++ SDK 的推荐封装形态。
- 对既有 Excel 文件修改、公式或页面配置，优先传路径和一个批量操作计划，由 Worker 在内部加载/保存，避免搬运整份单元格矩阵。
- 只有 Rust 创建的新数据或大批量追加确有结构化传输需要时，才传矩形批次。该次 .NET 测量表明 2.18 MB JSON 请求的解析/IPC余量不足 0.1 秒；目前没有证据值得换更复杂的共享内存/Arrow。
- C++ SDK 原生调用对比、Rust→LibXL 的同机额外开销、Aspose/GemBox 原生应用与 Worker 对比、Python Resident Worker、Worker 超时/取消与并发稳定性仍为 `NOT_TESTED`。

## 授权与分发

SDK 和许可证密钥未提交。当前测试使用试用/评估组件，不可据此进行生产再分发。最终部署必须按购买的产品、开发者/部署地点/OEM 许可及允许再分发文件核对：Aspose 说明不同终端分发场景需匹配对应 license 类型；GemBox EULA 单列可再分发文件和限制；LibXL 页面说明其商业应用可免额外 royalty 分发，同时另有开发者/企业许可证选择；Spire C++ 需要试用或正式 license，且本轮 SDK 已被 GLIBC 2.29 阻断。

- [Aspose 许可类型与终端分发说明](https://purchase.aspose.com/policies/license-types/)、[Aspose 使用许可 FAQ](https://purchase.aspose.com/faqs/licensing)
- [GemBox EULA](https://www.gemboxsoftware.com/EULA.pdf)、[GemBox 授权 FAQ](https://www.gemboxsoftware.com/support/faq)
- [LibXL 商业应用分发说明](https://www.libxl.com/home.html)、[LibXL 许可证选项](https://www.libxl.com/purchase.html)
- [Spire.XLS for C++ SDK / 试用说明](https://www.e-iceblue.com/Download/xls-for-cpp.html)
