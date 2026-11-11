# 阶段三：适量性能测试

状态：已完成本轮可运行的三引擎性能测试；Aspose.Cells C++ 与 Spire.XLS C++ 因阶段一 glibc 2.28 不兼容而未进入性能阶段。

## 条件与覆盖范围

- 环境：与阶段一相同的 Rocky Linux 8.9 x86_64 / glibc 2.28 离线容器；Aspose/GemBox 使用 .NET 10 self-contained 发布，LibXL 使用官方 C++ SDK。每个 Worker 常驻运行，执行预热后开始三轮。
- 普通批次：8 个 XLS + 16 个 XLSX，每份合成工作簿约 100 行；每轮从只读原始输入副本重新加载，共每引擎 72 个批次记录（24 × 3）。单批包括打开、追加 5 行、改值、插入/删除行列、改单元格样式、打印设置、公式重算调用、保存；XLS 另保存一份 XLSX 转换输出。LibXL 不能计算，且其转换能力仅为重建，耗时没有与前两者作为完整工作流比较。
- 容量：150,000 行 × 10 列 XLSX 与 1,000,000 行 × 8 列 XLSX 各执行一次打开、读取末行值、改单元格、保存、重新打开验证。容量测试不是三轮，且不执行复杂样式与公式计算。
- 明确未覆盖：10,000–100,000 公式单元格专用负载、openpyxl 同机基准、真实业务文件与麒麟/UOS 实机。Excel COM 转换及 BetterOffice 小型公式参照已在 Windows 专项实测，和 Linux 容量/批次表分开报告。
- 所有计时均是容器单次运行观测；CSV 同时保留各轮、每文件数据、错误、输入/输出字节、峰值 RSS 与 CPU 时间。中位数是描述性统计，不能外推到其他硬件/工作簿。

## 实测摘要

表中普通批次是 24 种文件的 3 轮汇总；XLS 与 XLSX 分开统计。三家均为每项支持自身所执行步骤后的 `PASS`；LibXL 不计算且其 XLS 转换并未计入此耗时。

| 引擎 | 普通 XLS 总耗时中位数 | 普通 XLSX 总耗时中位数 | 150k XLSX | 1m XLSX |
|---|---:|---:|---|---|
| Aspose.Cells .NET 26.10.0 | 162.0 ms（n=24） | 52.8 ms（n=48） | PASS，6.42 s，峰值 466 MiB | PASS，34.18 s，峰值 2,136 MiB |
| GemBox.Spreadsheet 2026.10.101 | 133.4 ms（n=24） | 69.4 ms（n=48） | BLOCKED_LICENSE（150 行上限） | BLOCKED_LICENSE（150 行上限） |
| LibXL C++ 5.2.0 eval | 406.1 ms（n=24） | 81.4 ms（n=48） | PASS，13.85 s，峰值 591 MiB | PASS，76.55 s，峰值 3,167 MiB |

LibXL 与 .NET 两者在大文件场景完成末行读取、编辑、保存和输出回开校验，因此本次观测里 Aspose .NET 更快且峰值 RSS 更低；每个大文件只运行一次，结论仅限于该生成样本和容器。Aspose 评估版会在保存件加入 Evaluation Warning 工作表，容量 PASS 不代表正式授权版保真结论。

批次里的 XLS 与 XLSX 操作不是完全等价的商业能力排名。GemBox 在本次测试的 XLS 合成批次中总耗时中位数低于 Aspose，但 XLSX 高于 Aspose；LibXL 的 XLS 普通批次总耗时也包括其 XLS 支持路径，CPU/RSS 较低，但完整业务操作集与公式计算不同，不能宣称整体胜出。各 CSV 的第一轮较慢，包含进程级预热未能消除的文件/页缓存与首次格式路径成本，故同时报告每个文件全部记录，不将首轮值单独用作结论。

性能字段记录：`open_ms`, `append_ms`, `edit_ms`, `delete_rows_cols_ms`, `style_ms`, `print_setup_ms`, `calculate_ms`, `save_ms`, `convert_ms`, `total_ms`, `peak_memory_mb`, `cpu_ms`。细项请按 engine 读取原始 CSV，不将单项最快解释为端到端最佳。Aspose/GemBox 的计算计时仅对应本合成样本的 `Calculate` 调用，不是公式密集负载，也未进行 BetterOffice 正确性比对。

## 同机 XLS→XLSX 转换基准（Windows）

补充同一台 Windows 11 Pro 主机的专门转换轮次：AMD Ryzen 5 2600X（12 logical processors）、16 GiB RAM，Excel 16.0 build 20527。Aspose .NET 与 GemBox .NET 使用 Release self-contained 进程；Excel COM 单一 Application 进程复用 24 个转换任务。8 个 XLS 输入各执行 3 轮。每项均打开 XLS、另存为 XLSX、回开检查 A2；另由 Excel COM 比较源与首轮输出的第一张表 used-range 值/公式、尺寸和工作表数。

| 实现 | 每文件总耗时中位数 | 打开中位数 | XLSX 保存中位数 | 运行/内容检查 |
|---|---:|---:|---:|---|
| Aspose.Cells .NET 26.10 evaluation | 39.82 ms | 2.75 ms | 13.03 ms | 24/24 执行成功；语义比较 `PARTIAL` 8/8（Evaluation Warning 多出一个工作表） |
| GemBox.Spreadsheet 2026.10.101 FREE-LIMITED | 44.16 ms | 9.25 ms | 11.54 ms | 24/24 执行成功；简单样本语义比较 `PASS` 8/8；100 行低于 150 行许可上限 |
| 常驻 Excel COM | 328.20 ms | 90.44 ms | 84.42 ms | 24/24 执行成功；简单样本语义比较 `PASS` 8/8 |

这是同硬件、同 8 个文件的有效小样本转换对比：两种 .NET 实现的本轮每文件另存时间低于 Excel COM。Aspose evaluation 输出增加警告表；GemBox 测试数据受 150 行免费许可限制；不能据此声称正式许可版对任意工作簿都更快或保真更高。每文件时间不含引擎/Excel 进程冷启动；Excel COM 记录的启动 768.7 ms、首次预热打开约 983.0 ms，.NET 客户端冷启动未单独计时。

原始 CSV：`evidence/performance/windows-conversion/{aspose-conversion,gembox-conversion}.csv`、`evidence/performance/excel-com-reference/excel-com-conversion.csv`。内容校验：`evidence/performance/windows-conversion/conversion-semantic-check.csv`。输出保存于相邻 `reports/samples/performance/`。这些 Windows 数据与 Rocky 的 Linux 批次/容量结果分开，不混入 Linux 排名。

## 引擎门禁和异常

- Aspose.Cells C++ 与 Spire.XLS C++：`NOT_TESTED` 性能。阶段一原生 SDK 已出现 glibc 2.29 符号要求，未在 glibc 2.28 绕过系统库继续测试。
- GemBox 大文件：两个文件的 CSV 保存载入阶段报错，错误为免费版本限制 150 行/5 工作表；这是 `BLOCKED_LICENSE`，不是技术上限或内存失败。商业/限时评估许可未提供，不能推断持证版本容量。
- LibXL：所有普通批次 `PASS`；日志标明 `calculate=UNSUPPORTED_BY_DESIGN`、`conversion=RECONSTRUCTION_ONLY`。100 万行峰值约 3.1 GiB，提示目标部署内存预算需按实际进程并发重新评估。
- 24 工作簿的生成输入为规整的 100 行数据；不能代表复杂图表、外部链接、宏、稀疏百万行等真实文件。

## 证据和复现

- 原始样本生成脚本：`../fixtures/generate_performance_fixtures.py`；输入行列、字节和 SHA-256：`../fixtures/performance/manifest.json`。
- Aspose .NET CSV 与日志：`evidence/performance/aspose/performance-aspose.csv`、`evidence/performance/aspose/offline-runtime.log`。
- GemBox CSV 与日志：`evidence/performance/gembox/performance-gembox.csv`、`evidence/performance/gembox/offline-runtime.log`。
- LibXL CSV 与日志：`evidence/performance/libxl/performance-libxl.csv`、`evidence/performance/libxl/libxl-performance-runtime.log`。
- 复现：分别运行 `scripts/run_dotnet_smoke_rocky8.ps1 -Engine aspose -Mode performance -SdkRoot <官方 .NET SDK>`、同脚本 `-Engine gembox`、`scripts/run_libxl_rocky8.ps1 -Mode performance -SdkRoot <官方 Linux LibXL SDK>`。脚本将发布物/SDK 外置并在网络禁用的固定 Rocky 镜像内运行。

## 限制

本实验没有完成 openpyxl 同机性能对照和 10k–100k 公式密集计算，所以不能回答商业库是否整体超过这些现有路径。Excel COM 转换小样本与 BetterOffice 公式正确性已另行测试。授权版可能消除评估限制或改变文件输出，必须在获得合法 SDK/许可证后重复能力和性能测试。Rocky 8 结果不替代麒麟 V10 / UOS V20 实机验证。
