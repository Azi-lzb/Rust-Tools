# rustTools Excel 引擎综合选型实验

状态：本轮可运行的实验与报告已完成，结论保持暂行。阶段一完成；阶段二三个兼容引擎第一轮能力实测；阶段三 Linux 普通批次/容量测试和 Windows 同机转换对照完成；阶段四三种 Rust 桥接路径已运行；阶段五形成按操作选引擎的建议。证据来自 Rocky Linux 8.9/glibc 2.28、Windows 11 与合成样本，不代表麒麟 V10/UOS V20 实机或生产业务样本。

## 引擎总表

| 引擎 | Linux 2.28 | XLS修改 | XLS转换 | 追加 | 删除/插入行列 | 样式 | 列宽/行高 | 打印设置 | 计算 |
|---|---|---|---|---|---|---|---|---|---|
| Aspose.Cells C++ 26.10 | FAIL | BLOCKED | BLOCKED | BLOCKED | BLOCKED | BLOCKED | BLOCKED | BLOCKED | BLOCKED |
| Aspose.Cells .NET 26.10.0 | PASS | PASS | PASS | PARTIAL | PARTIAL | PARTIAL | PARTIAL | PARTIAL | PASS |
| Spire.XLS C++ 16.7.0 | FAIL | BLOCKED | BLOCKED | BLOCKED | BLOCKED | BLOCKED | BLOCKED | BLOCKED | BLOCKED |
| GemBox.Spreadsheet .NET 2026.10.101 | PASS | PASS | PASS | PARTIAL | PARTIAL | PARTIAL | PARTIAL | PARTIAL | PASS |
| LibXL C++ 5.2.0 | PASS | PASS | PARTIAL | PARTIAL | PARTIAL | PARTIAL | PARTIAL | PARTIAL | UNSUPPORTED |

表内 PASS 仅覆盖报告列出的合成样本/API 操作，不表示完整功能矩阵或生产保真度 PASS。Aspose trial 会加 `Evaluation Warning` 工作表；GemBox FREE-LIMITED 限 150 行；LibXL evaluation 只允许最多 300 个读取单元格。两个 C++ SDK 不兼容 glibc 2.28，因此阶段二被门禁阻断。所有引擎在 Kylin V10/UOS V20 上均 `NOT_TESTED`。

## 已证实的结论

- Rocky 8.9/glibc 2.28 x86_64 上，LibXL 5.2.0、Aspose.Cells .NET 26.10.0、GemBox.Spreadsheet 2026.10.101 完成离线目录式实际运行。Aspose.Cells C++ 26.10 与 Spire.XLS C++ 16.7.0 的官方 Linux SDK 需要 GLIBC 2.29（Aspose C++ 另有 GLIBCXX/FreeType 链接要求），不能用于目标 glibc 2.28 基线。
- Aspose .NET、GemBox .NET、LibXL 在样本上都能原格式保存 XLS 与 XLSX；Aspose/GemBox 能将 XLS 输出 XLSX。LibXL 不是通用 XLS 转换器，本轮重建样例只复制标量值和公式。
- Aspose/GemBox 在指定数学/查找/日期/跨表/命名区域公式集上实际执行计算并验证保存缓存；同一合成公式集在 Excel COM 与 BetterOffice 上得到一致结果。此结果不是全函数兼容性或公式密集性能结论。LibXL 公式计算 `UNSUPPORTED`。
- 输出保真有限制：Aspose trial 会新增警告工作表；GemBox 对原工作簿页眉的无修改保存语义检查失败。二者整体保真度均不是全量 `PASS`。

## 性能结果

普通合成批次 XLS 工作流总耗时中位数：Aspose .NET 162.0 ms、GemBox 133.4 ms、LibXL 406.1 ms；XLSX：Aspose .NET 52.8 ms、GemBox 69.4 ms、LibXL 81.4 ms。LibXL 不执行公式计算，不能作为完整公式业务工作流对手。Aspose .NET 的 150k/1m XLSX 编辑保存通过，分别 6.42/34.18 秒、峰值 RSS 466/2,136 MiB；LibXL 分别 13.85/76.55 秒、591/3,167 MiB。GemBox 大文件因免费版 150 行上限为 `BLOCKED_LICENSE`。Windows 同机 8 个 XLS×3 轮转换中位数：Aspose .NET 39.82 ms、GemBox 44.16 ms、常驻 Excel COM 328.20 ms；Aspose 输出因 trial 警告表为 `PARTIAL`，GemBox 在 100 行样本语义检查 `PASS`。详见 [`03_PERFORMANCE.md`](03_PERFORMANCE.md) 和原始 CSV。

未测试 10k–100k 公式密集负载、openpyxl 的同机性能基线和真实业务转换。Windows 小样本表明两个 .NET 转换器在该夹具上的转换用时低于 Excel COM；这不足以支持生产替换决定。

## 桥接与架构建议

Rust→LibXL C ABI、Rust→Aspose .NET 常驻 Worker 和 test-only Rust→openpyxl 批量文件 IPC 均已完成最小运行原型。工作簿处理和桥接耗时见 [`04_RUST_BRIDGE.md`](04_RUST_BRIDGE.md)。用户提供的架构图及暂行建议见 [`05_ARCHITECTURE_RECOMMENDATION.md`](05_ARCHITECTURE_RECOMMENDATION.md)。当前证据支持保持生产路由不变；LibXL/Aspose .NET 可作为取得正式许可后继续验证的实验候选。Aspose/GemBox 的同机 XLS 转换结果值得用正式许可及真实样本复测，但当前仍不替换 Excel COM。

## 报告索引

- [01_LINUX_COMPATIBILITY.md](01_LINUX_COMPATIBILITY.md)
- [02_EXCEL_CAPABILITIES.md](02_EXCEL_CAPABILITIES.md)
- [03_PERFORMANCE.md](03_PERFORMANCE.md)
- [04_RUST_BRIDGE.md](04_RUST_BRIDGE.md)
- [05_ARCHITECTURE_RECOMMENDATION.md](05_ARCHITECTURE_RECOMMENDATION.md)
- 能力与包部件核验：[phase2-validation.json](evidence/phase2-validation.json)
