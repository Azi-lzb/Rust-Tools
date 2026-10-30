# XLS → XLSX Excel COM 性能基准

本目录只用于研究 Windows 上四种 Excel COM 批量转换实现，不属于 PyTools 生产运行路径。所有输入均为本工具生成的合成工作簿；禁止把业务工作簿放入本目录。

## 目录

- `fixtures/generate_dataset.py`：生成 100 本真实 BIFF8 `.xls`，并校验文件格式、固定 SHA-256 与 B10/B25/B50/B100 子集。
- `fixtures/generated/`：生成后的 XLS 数据集和清单。
- `workers/`：Python、PowerShell COM worker；C/D 方案共用同一 PowerShell 转换函数。
- `runner/`：Rust 基准调度、Rust 原生 COM 实现和 JSON 结果汇总。
- `results/`：每轮结构化 JSON 结果及汇总。
- 根目录 `perf_xls_to_xlsx_com_comparison.md`：最终分析报告。

## 四种方案与生命周期

| 代号 | 调用路径 | 批次复用 | 逐文件对照 |
| --- | --- | --- | --- |
| A | Rust 原生 COM，专用 STA 线程 | 一个 STA 与一个 Excel 实例处理整批 | 每本重建 STA 与 Excel |
| B | Rust → Python/pywin32 → Excel COM | 一个 Python worker 与一个 Excel 实例处理整批 | 每本重建 worker、STA 与 Excel |
| C | PowerShell → Excel COM | 一个 `-STA` PowerShell 与一个 Excel 实例处理整批 | 每本重建 PowerShell 与 Excel |
| D | Rust → PowerShell worker → Excel COM | Rust 通过 JSON 与一个 PowerShell worker 通信；整批复用 Excel | 每本重建 worker 与 Excel |

C、D 使用相同的 `workers/excel_com.ps1` COM 操作；C 走命令行参数，D 走 JSON stdin/stdout。Python 仅承担 COM worker，不承载业务逻辑，不嵌入 CPython，也不使用 FFI。

所有方案统一先创建临时空白工作簿以设置手动计算，再关闭该空白簿；随后使用 `Workbooks.Open(path, UpdateLinks=0, ReadOnly=true)`、`SaveAs(path, 51)`、`Close(false)`，关闭提示、事件及屏幕更新，不调用任何 Calculate 方法。该机 Excel 只有在存在工作簿时才接受 `Application.Calculation` 设置，因此空白簿初始化是每种方案一致的 Excel 配置阶段。每个任务只操作独立输入副本及新建输出目录。worker 记录创建的 Excel PID，只有证明进程由本任务新建且工作簿计数为零时才会退出该 Excel；不会按进程名结束已有 Excel。

## 运行要求

- Windows、Microsoft Excel、PowerShell 5.1、Python 3、pywin32、openpyxl、xlrd、Rust stable。
- Python/PowerShell 仅为研究 worker；不加入 PyTools 正式运行时。
- Cargo 构建目录使用 `F:\RustBuild\pytools` 下的独立 benchmark target 子目录。
- 首次运行生成合成数据集后记录 `fixtures/generated/manifest.json` 与各批次清单哈希。重复生成不覆盖现有文件。
- 正式测试包含预热和 5 轮；批次规模、方案及生命周期交错执行。转换计时不含样本生成及正确性复核。
- 启动前保存 Excel PID 基线；最终只把本轮新建的 PID 用于资源清理判定。基线中出现的 Excel 实例保持不动。

详细实验步骤、机器环境、失败批次与适用限制记录在最终报告中。
