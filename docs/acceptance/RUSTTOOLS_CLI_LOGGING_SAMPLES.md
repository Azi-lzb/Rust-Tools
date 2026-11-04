# rustTools CLI 精简日志样例集

验收日期：2026-11-04。下列内容来自本机实际 RustTools CLI 子进程输出，不是理想化示例。地区总分 10/50/100 捕获为合并终端流；其他文件链接注明了来源和捕获方式。

**当前样例只覆盖 7 项业务功能，不能替代 37 项全量逐项验收。** 具体缺项见同目录的 [完整验收报告](RUSTTOOLS_CLI_LOGGING_ACCEPTANCE.md) 覆盖矩阵。

## 1-2 按使用区域快速汇总（独立 CLI）

输入为从含公式的临时 xls 转换出的 xlsx。Calamine 读取 1 个文件并生成 3 行汇总，结果工作簿成功保存。此处测的是核心 CLI 命令，不代表交互菜单完整验收。

```text
[2026-11-04 22:02:24] [INFO] 开始按使用区域汇总：1 个输入，1 个可用工作簿
[2026-11-04 22:02:24] [INFO] 工作簿「formula-output.xlsx」：工作簿汇总完成 [1/1]：formula-output.xlsx
[2026-11-04 22:02:24] [INFO] 正在保存汇总工作簿
[2026-11-04 22:02:24] [INFO] 按使用区域汇总完成：1 个工作簿，3 行
[2026-11-04 22:02:24] [INFO] 运行记录已保存：PT-6-1_pytools-pt-6-1-25408-1793800944472217000.json
{
  "execution": {
    "ooxml_error_cells_patched": 0,
    "read_elapsed_ms": 7,
    "reader_engine": "calamine",
    "total_elapsed_ms": 29,
    "write_elapsed_ms": 15,
    "writer_engine": "rust_xlsxwriter"
  },
  "files": 1,
  "path": "C:\\Users\\AZI\\AppData\\Local\\Temp\\RustTools-COM-formula-26bed446ead746a3b806c99f0ec7b821\\calamine-output\\20261104_220224_按使用区域汇总（全量）.xlsx",
  "rows": 3,
  "run_id": "pytools-pt-6-1-25408-1793800944472217000",
  "run_log_path": "C:\\Users\\AZI\\AppData\\Local\\rust-tools\\logs\\PT-6-1_pytools-pt-6-1-25408-1793800944472217000.json",
  "saved": true
}
```

完整原始捕获：[summary-calamine.stdout.txt](captures/summary-calamine.stdout.txt)

## 2-1 按配置检查重复

这是测试副本上的一次真实菜单运行：启用 1 项配置，在 `Commented` 工作表按第 2 列标记 1 行重复数据。来源文件路径在日志里被完整打印，存在可缩短空间。

```text
[2026-11-04 22:38:55] [INFO] 正在检查：配置工作表「去重追加数据配置」第 2 行「DedupMenu」；源文件：C:\Users\AZI\AppData\Local\Temp\pytools-dedup-menu-28180-1793803135796208700-1\source.xlsx；源工作表：Commented；标识列：2
[2026-11-04 22:38:56] [INFO] 配置结果：第 2 行「DedupMenu」；标记重复行 1 行，已更新并保存源工作簿
[2026-11-04 22:38:56] [WARN] 执行完成，发现业务异常：按配置检查重复（在源工作簿上标红标记重复行）；保存工作簿数 1，失败配置任务 0，成功配置任务 1，标记重复行数 1，跳过配置任务 0；耗时 0.080 秒；运行编号 pytools-2-1-18396-1793803135946965700-0
[2026-11-04 22:38:56] [WARN] 业务异常：1 项，明细 1 条
[2026-11-04 22:38:56] [WARN] [重复数据] 配置检查发现重复行并已标记（1 项）
[2026-11-04 22:38:56] [WARN] 输出：C:\Users\AZI\AppData\Local\Temp\pytools-dedup-menu-28180-1793803135796208700-1\source.xlsx
[2026-11-04 22:38:56] [WARN] 详细事件：run-events\2-1_pytools-2-1-18396-1793803135946965700-0.jsonl
```

完整捕获：[configured_duplicate_check_marks_formula_workbook_in_place.captured.txt](captures/configured_duplicate_check_marks_formula_workbook_in_place.captured.txt)

## 2-2 按配置删除重复

输入为测试副本；重复 1 行并保存源工作簿。下面是实际 CLI 输出中的最终结果片段。

```text
[2026-11-04 22:38:55] [INFO] 正在检查：配置工作表「去重追加数据配置」第 2 行「DedupMenu」；源文件：C:\Users\AZI\AppData\Local\Temp\pytools-dedup-menu-28180-1793803135796255100-2\source.xlsx；源工作表：Commented；标识列：2
[2026-11-04 22:38:56] [INFO] 配置结果：第 2 行「DedupMenu」；删除重复行 1 行，已更新并保存源工作簿
[2026-11-04 22:38:56] [INFO] 执行成功：按配置删除重复（在源工作簿上就地删除重复行）；保存工作簿数 1，删除重复行数 1，失败配置任务 0，成功配置任务 1，跳过配置任务 0；耗时 0.056 秒；运行编号 pytools-2-2-4260-1793803135971298600-0
[2026-11-04 22:38:56] [INFO] 输出：C:\Users\AZI\AppData\Local\Temp\pytools-dedup-menu-28180-1793803135796255100-2\source.xlsx
[2026-11-04 22:38:56] [INFO] 详细事件：run-events\2-2_pytools-2-2-4260-1793803135971298600-0.jsonl
```

完整捕获：[configured_duplicate_delete_removes_duplicates_in_place.captured.txt](captures/configured_duplicate_delete_removes_duplicates_in_place.captured.txt)

## 2-3 按配置去重追加到目标工作簿

输入为独立测试副本：过滤目标已有/批内重复 2 行，向目标追加 1 行。输出统计正确，但源和目标绝对路径多次重复打印。

```text
[2026-11-04 22:38:55] [INFO] 正在检查：配置工作表「去重追加数据配置」第 2 行「DedupMenu」；源文件：C:\Users\AZI\AppData\Local\Temp\pytools-dedup-menu-28180-1793803135796637700-8\source.xlsx；源工作表：Commented；标识列：2；目标文件：C:\Users\AZI\AppData\Local\Temp\pytools-dedup-menu-28180-1793803135796637700-8\target.xlsx；目标工作表：Append
[2026-11-04 22:38:55] [INFO] 正在写入目标：C:\Users\AZI\AppData\Local\Temp\pytools-dedup-menu-28180-1793803135796637700-8\target.xlsx（追加 1 行，过滤重复 2 行）
[2026-11-04 22:38:56] [INFO] 已写入：C:\Users\AZI\AppData\Local\Temp\pytools-dedup-menu-28180-1793803135796637700-8\target.xlsx
[2026-11-04 22:38:56] [INFO] 目标汇总结果：目标 C:\Users\AZI\AppData\Local\Temp\pytools-dedup-menu-28180-1793803135796637700-8\target.xlsx；关联配置 第 2 行「DedupMenu」；过滤重复 2 行，追加 1 行，目标工作簿已保存
[2026-11-04 22:38:56] [INFO] 执行成功：按配置去重追加；写入目标数 1，受阻目标数 0，失败目标数 0，未处理目标数 0，跳过配置数 0，过滤重复行数 2，追加行数 1；耗时 0.180 秒；运行编号 pytools-2-3-11156-1793803135940355000-0
```

完整捕获：[interactive_dedup_append_direct_copies_source_formula_and_summarizes.captured.txt](captures/interactive_dedup_append_direct_copies_source_formula_and_summarizes.captured.txt)

## 3-1 按打印配置生成（umya-spreadsheet）

测试副本生成 1 个目标工作簿。测试原始输出还显示了通用“处理中...”行。

```text
[2026-11-04 22:39:11] [INFO] 开始按打印配置生成静态打印文件（umya-spreadsheet）
[2026-11-04 22:39:11] [INFO] 正在执行打印配置：工作表「打印配置」第 2 行；源文件：C:\Users\AZI\AppData\Local\Temp\pytools-print-menu-28308-1793803151669707600-1\source.xlsx；源工作表：Marked；目标文件：C:\Users\AZI\AppData\Local\Temp\pytools-print-menu-28308-1793803151669707600-1\printed.xlsx；目标工作表：Printed；打印区域：B2:D4
  处理中...
[2026-11-04 22:39:11] [INFO] 打印配置执行结果：第 2 行；重复行：不适用；修改：是；保存：是
[2026-11-04 22:39:11] [INFO] 执行成功：按配置生成静态打印文件；保存工作簿 1 个；失败配置项 0，成功配置项 1，跳过配置项 0；耗时 0.064 秒；运行编号 pytools-3-1-12864-1793803151842259700-0
[2026-11-04 22:39:11] [INFO] 输出：printed.xlsx
```

完整捕获：[interactive_print_menu_routes_umya_openpyxl_and_precheck.captured.txt](captures/interactive_print_menu_routes_umya_openpyxl_and_precheck.captured.txt)

## 4-4 Office 格式转换（真实 Excel COM；独立 CLI）

本机 `Excel.Application` 可用；通过 pywin32 worker 转换临时 xls 副本。样本含 1 个公式单元格，`CalculateFullRebuild` 后保存并验证通过。源文件哈希没有变化。

实际运行 `RustTools.exe convert-file --input <测试副本.xls> --to xlsx --output <临时输出.xlsx> --engine pywin32`。进程退出码为 0，输出文件存在，输入副本 SHA-256 未改变。该命令当前把 worker 的结果 JSON 直接写到 stdout；它不是交互菜单摘要。

```json
{"com_application":"Excel.Application","com_bridge":"pywin32","com_process_id":27476,"com_worker_process_id":26040,"engine":"pywin32","input":"C:\\Users\\AZI\\AppData\\Local\\Temp\\RustTools-CLI-acceptance-convert-907ead9450884882aea651fc8444dbc4\\formula-copy.xls","output":"C:\\Users\\AZI\\AppData\\Local\\Temp\\RustTools-CLI-acceptance-convert-907ead9450884882aea651fc8444dbc4\\formula-copy.xlsx","status":"converted","validation":{"calculation":"CalculateFullRebuild before SaveAs","formula_cells":1,"level":"excel_com_sheet_order_formula_text_and_cell_values","numeric_tolerance":1e-15,"recalc_method":"CalculateFullRebuild","used_cells":3,"value_cells":2,"verified":true,"worksheet_count":1}}
```

完整实际 CLI 捕获：[conversion-pywin32com-cli.terminal.txt](captures/conversion-pywin32com-cli.terminal.txt)。另有较短的 worker 返回捕获：[conversion-pywin32com-formula.stdout.txt](captures/conversion-pywin32com-formula.stdout.txt)。交互菜单输出仍待验收。

## 5-5 地区总分校验：10 / 50 / 100 个工作簿

以下来自真实交互菜单的合并输出。测试副本包含一个地区差异、一个损坏 xlsx 和 `north/duplicate.xlsx`、`south/duplicate.xlsx` 两个同名文件。

```text
[2026-11-04 22:43:53] [INFO] 开始地区总分校验，共 100 个工作簿
[2026-11-04 22:43:53] [WARN] 进度 2/100（2%）：工作簿「south/duplicate.xlsx」，工作表「分地区汇总」：分地区总分核对完成：发现 1 处差异
[2026-11-04 22:43:53] [ERROR] 进度 3/100（3%）：地区总分校验阶段 workbook_read 失败：工作簿「broken.xlsx」：第 3 个工作簿读取失败：open broken.xlsx: Xlsx error: Zip error: invalid Zip archive: Could not find EOCD: Zip error: invalid Zip archive: Could not find EOCD: invalid Zip archive: Could not find EOCD
[2026-11-04 22:43:54] [INFO] 地区总分校验完成：成功 99 个，失败 1 个，核对 99 张工作表，发现 1 项错误
[2026-11-04 22:43:54] [WARN] 部分完成：地区总分校验；选择 100，尝试 100，成功 99，失败 1，跳过 0，取消 0，未完成 0；工作表发现 99，尝试 99，成功 99，失败 0，跳过 0；保存工作簿 1 个；发现差异数 1，失败工作簿数 1，核对工作表数 99；耗时 0.657 秒；运行编号 pytools-pt-7-5-4316-1793803433703812700
[2026-11-04 22:43:54] [WARN] 业务异常：1 项，明细 1 条
[2026-11-04 22:43:54] [WARN] [地区总分校验] 校验发现不一致项（1 项）
[2026-11-04 22:43:54] [WARN] 失败阶段 workbook_read [EXCEL_READ_FAILED]：broken.xlsx：工作簿「broken.xlsx」：第 3 个工作簿读取失败：open broken.xlsx: Xlsx error: Zip error: invalid Zip archive: Could not find EOCD: Zip error: invalid Zip archive: Could not find EOCD: invalid Zip archive: Could not find EOCD
```

完整原始输出分别保存为：[10 文件](captures/region-sum-10.terminal.txt)、[50 文件](captures/region-sum-50.terminal.txt)、[100 文件](captures/region-sum-100.terminal.txt)。全部真实行数分别为 74、72、72；全部成功创建核对结果文件。程序摘要耗时分别为 0.101、0.334、0.657 秒。错误原因重复了 ZIP 链，虽能定位但仍偏长。
