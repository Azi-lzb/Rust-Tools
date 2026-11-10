# Excel XLSX 修改引擎能力矩阵（真实执行结果）

来源：`results/capability_*.json`（探针程序逐项执行，非文档推断）。
图例：PASS = 探针真实执行成功；UNSUPPORTED = 引擎无原生 API（探针与源码双重核实）；FAIL = 有 API 但执行失败。
执行环境：Windows 11 x64，openpyxl 3.1.5（随包）、betteroffice-xlsx 0.5.1（vendored 补丁版）、rxls 0.1.4、umya-spreadsheet 3.1.1。

## M01–M14 修改能力

| 操作 | openpyxl | betteroffice 0.5.1 | rxls 0.1.4 | umya 3.1.1 |
| --- | --- | --- | --- | --- |
| M01 单格改值 | PASS | PASS | PASS | PASS |
| M02 批量 100 格 | PASS | PASS | PASS | PASS |
| M03 批量 1000 格 | PASS | PASS | PASS | PASS |
| M04 公式文本（含缓存值） | PASS | PASS（`SetCell{formula}`） | PASS（`set_cell_formula` 带 cached） | PASS |
| M05 数字格式 | PASS | PASS（`SetRangeNumberFormat`） | **UNSUPPORTED**（无 API） | PASS |
| M06 字体/填充/边框/对齐 | PASS | PASS（`PatchRangeStyle`） | **UNSUPPORTED**（无 API） | PASS |
| M07 行高列宽 | PASS | PASS | PASS | PASS |
| M08 合并增删 | PASS | PASS | PASS | PASS |
| M09 批注改/增 | PASS | **UNSUPPORTED**（模型无批注表示） | PASS | PASS |
| M10 命名区域 | PASS | PASS（含作用域） | PASS（仅全局） | PASS |
| M11 打印设置 | PASS | 部分（打印区域可走 `_xlnm.Print_Area` 定义名；页面方向无 API → UNSUPPORTED） | 部分（打印区域 PASS；方向无 API → UNSUPPORTED） | PASS（页面方向原生；打印区域走定义名） |
| M12 表增删重命名 | PASS | PASS | PASS | PASS |
| M13 复杂模板单点 | PASS | PASS | PASS | PASS |
| M14 同簿混合修改 | PASS | 部分（4/5 子操作；批注缺失） | PASS | PASS |
| （补充）行列插入删除 | PASS | PASS（`InsertRows/DeleteRows`） | **UNSUPPORTED**（官方 non-goal） | PASS |
| （补充）XLSM/VBA 保留 | PASS（keep_vba） | no-op PASS；edit 后 Excel 打不开（见保真报告） | no-op PASS；edit 后 Excel 打不开 | PASS |

**计分（M01–M14）**：openpyxl 14/14；umya 14/14；betteroffice 12/14（M09、M11 缺）；rxls 11/14 + M11 部分（M05、M06 缺）。

## 关键限制原文（探针记录）

- betteroffice M09：`UNSUPPORTED: betteroffice-xlsx 0.5.1 has no comment representation (model + ops verified)` —— model 与 ops 两层均无批注类型。
- betteroffice M11：`no page-orientation representation in betteroffice-xlsx 0.5.1`。
- rxls M05/M06：`rxls 0.1.4 exposes no style/number-format edit API` —— docs.rs 方法表与探针一致。
- rxls 行列操作：`rxls does not insert/delete rows or columns (documented non-goal)`。
- rxls 大文件打开：`invalid ZIP package: OOXML package entry is too large`（100k 行夹具 sheet XML 11MB 超限，见性能报告）。

## 结论要点

1. **能力覆盖冠军是 umya（14/14）**，其次是 openpyxl（14/14）。
2. betteroffice 缺批注与页面方向；rxls 缺样式/数字格式——两者各有明确能力空洞。
3. rxls 的**部件字节级保留**（未触碰部件逐字节不变，含 VBA）是其他引擎都没有的独有特性。
4. 三个 Rust 引擎的样式类能力：betteroffice 走 `PatchRangeStyle`（完整），umya 走 Style 对象（完整），rxls 完全没有。
