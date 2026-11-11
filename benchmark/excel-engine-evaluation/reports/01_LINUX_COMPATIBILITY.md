# 阶段一：Linux glibc 2.28 与离线部署

状态：阶段一 Rocky 8.9/glibc 2.28 门禁已完成。只有完成实际 Linux 运行才记为 `GLIBC228_PASS`；厂商文档和静态分析不能替代实测。麒麟 V10、统信 UOS V20 必须单独实机验证，当前均为 `NOT_TESTED`。

## 环境

- Rocky Linux 8.9 x86_64 容器，镜像 digest：`sha256:9794037624aaa6212aeada1d28861ef5e0a935adaf93e4ef79837119f2a2d04c`。
- `uname -m`、`ldd --version`、`getconf GNU_LIBC_VERSION` 与 `/etc/os-release` 原始输出：[`rockylinux8-glibc.txt`](evidence/environment/rockylinux8-glibc.txt)。确认 glibc 2.28。
- 本容器证明 Rocky 8/glibc 2.28 结果，不代表麒麟 V10 或 UOS V20 实机验收。

## 五引擎状态

| 引擎 | Linux SDK/版本 | GLIBC228 | Kylin V10 | UOS V20 | 离线自包含冒烟 | 当前说明 |
|---|---|---|---|---|---|---|
| Aspose.Cells C++ 26.10 | Linux x86_64 SDK | INCOMPATIBLE | NOT_TESTED | NOT_TESTED | FAIL | 自定义 C++17 smoke 编译成功但链接失败：库要求 GLIBC_2.29、GLIBCXX_3.4.26，且需要 Rocky 8 缺少的 FreeType 符号；未运行工作簿读写 |
| Aspose.Cells .NET 26.10.0 | .NET 10 self-contained | `GLIBC228_PASS` | NOT_TESTED | NOT_TESTED | PASS | XLSX/XLS 读改存重开 PASS；发布目录附带 SkiaSharp Linux native assets 与 Rocky 8 ICU |
| Spire.XLS C++ 16.7.0 | 官方 Linux C++ NuGet SDK | INCOMPATIBLE | NOT_TESTED | NOT_TESTED | FAIL | 编译成功；离线运行实例化 Workbook 失败，SDK 动态加载的 `libSkiaSharp.so` 要求 GLIBC_2.29；见 [`evidence/spire/sdk-metadata.txt`](evidence/spire/sdk-metadata.txt) |
| GemBox.Spreadsheet .NET 2026.10.101 | .NET 10 self-contained | `GLIBC228_PASS` | NOT_TESTED | NOT_TESTED | PASS | XLSX/XLS 读改存重开 PASS；发布目录附带 Rocky 8 ICU |
| LibXL C++ 5.2.0 | Linux x64 evaluation SDK | `GLIBC228_PASS` | NOT_TESTED | NOT_TESTED | PASS | XLSX/XLS 读改存冒烟成功；见下文和原始证据 |

## LibXL 5.2.0 实测

- 官方 Linux x64 SDK 包 MD5 与厂商发布值一致，元数据见 [`sdk-metadata.txt`](evidence/libxl/sdk-metadata.txt)。评估许可限制：首行插入 banner，读取上限 300 个单元格；无 key。不得把评估版结果当作正式商业授权。
- 使用 Rocky 8 GCC 8.5.0、CMake 3.26.5 编译了官方 SDK 的原生 C++ 样例程序。最终程序在同 digest 的干净 Rocky 8 容器断网运行，rootfs 只读；容器无 .NET、Java、Python、Office 或 LibreOffice。输出见 [`libxl-offline-runtime.log`](evidence/libxl/libxl-offline-runtime.log)。
- XLSX 数值/文本/公式表达式读取、值修改、保存、重新打开校验：PASS。XLS 同样完成原格式修改和重开验证：PASS。
- `$ORIGIN/../lib` 私有库目录包大小 21.55 MiB，仅含 smoke 程序和 `libxl.so`。这是目录式离线运行验证，不是 deb 包验证。
- SDK `libxl.so` 最高符号要求：`GLIBC_2.14`、`GLIBCXX_3.4.15`、`CXXABI_1.3.5`；冒烟程序：`GLIBC_2.14`、`GLIBCXX_3.4.21`、`CXXABI_1.3.9`。直接依赖 `libstdc++`、`libm`、`libgcc_s`、`libc`。ELF/readelf/ldd 原始材料位于 [`evidence/libxl/`](evidence/libxl/)。
- XLS→XLSX 样例是 LibXL 从 XLS 读取后把测试单元格写入新 XLSX 的重建；不是保留完整工作簿结构的转换器。直接将 BIFF `.xls` 交给 XMLBook 打开被 SDK 拒绝。
- LibXL 没有公式计算引擎；不计算公式缓存。没有执行任何 PDF 导出测试。
- 临时安装包清单及输入/输出 SHA-256：[`artifact-hashes.csv`](evidence/libxl/artifact-hashes.csv)。

## SDK、依赖和授权核查

官方下载记录、许可限制、完整动态库/间接依赖、编译器版本和离线复制步骤应逐引擎保存到 `evidence/`。对 .NET 程序，发布目录须直接执行，不以 `dotnet run` 代替；对 C++ 程序，需以 `file`、`readelf`、`ldd` 等检查二进制并在目标 glibc 实机运行。

| 引擎 | SDK 与许可证据 | 静态依赖检查 | runtime/离线验证 |
|---|---|---|---|
| Aspose.Cells C++ 26.10 | 已取得官方 Linux SDK 评估包；无商业 key | 实测链接失败，超出 GLIBC/GLIBCXX 并缺 FreeType 符号 | FAIL |
| Aspose.Cells .NET 26.10.0 | self-contained publish 成功，评估版无商业 key | 直接 `.so`/runtime 依赖留档，使用 SkiaSharp 2.88.3 Linux native assets + ICU 60.3 | Rocky 8 断网目录发布，XLSX/XLS PASS |
| Spire.XLS C++ 16.7.0 | 已取得官方 Linux C++ NuGet 评估包；无商业 key | 已检查三份 SDK `.so`；SkiaSharp 要求 GLIBC_2.29 | FAIL |
| GemBox.Spreadsheet .NET 2026.10.101 | FREE-LIMITED-KEY，150 行限制内小样本 | 自包含目录 + ICU 60.3；默认内部渲染 | Rocky 8 断网目录发布，XLSX/XLS PASS |
| LibXL C++ | 完成评估版证据；商业再分发需有效授权 | 已检查 `libxl.so` 与 smoke 程序 | Rocky 8/glibc 2.28 断网目录包 PASS |

