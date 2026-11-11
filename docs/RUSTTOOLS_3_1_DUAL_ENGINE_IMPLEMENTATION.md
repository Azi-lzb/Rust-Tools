# rustTools 3-1 双引擎并行实现报告

状态：**已实现并通过对拍与性能验收**。功能分支 `feature/3-1-dual-engine`（基线 b7085f4）。

## 1. 双入口设计

按用户要求实现「一个业务功能、两个独立入口、两套独立实现、同一份业务配置」：

```
3 报表打印
  ├── 1  按打印配置生成（Openpyxl）  ← 原 3-1 入口，完整保留
  ├── 2  按打印配置生成（Rust）      ← 新增，calamine + OOXML + rust_xlsxwriter
  └── 3  打印配置预校验              ← 原 3-3，编号由 2 顺延
```

- **无自动路由**：两个入口各自只调用自己的实现。Rust 入口不调用 Openpyxl/Python/UMYA/BetterOffice。
- 原 Openpyxl 入口的调用链、Python 脚本、日志、追加既有目标、XLSM 处理、mode 1/2/3 行为**全部保留**，未改动任何一行。
- 原 3-2 入口此前已并入 1（2026-11-09 合并），本次保持。

## 2. 菜单修改（RustTools.rs run_interactive_print_menu）

```
1  按打印配置生成（Openpyxl）  【打印配置：源工作簿/源工作表/打印区域/目标工作簿】
2  按打印配置生成（Rust）  【同一份打印配置；输出文件名追加 _Rust；暂不支持 XLSM】
3  打印配置预校验  【打印配置】
0  返回
```

- Openpyxl 入口 feature_id 仍为 `3-1`；Rust 入口 feature_id 为 `3-1-Rust`（日志独立可辨）。
- 预校验编号从 2 顺延到 3；测试已同步。

## 3. Rust 引擎调用链

```
run_configured_static_print_with_python
  ├── load_print_tasks（calamine 读「打印配置」，15 列字段全兼容）
  ├── RustWriter 分支：目标重映射为 _Rust 后缀（校验前完成）
  │     ├── XLSM 目标 → 明确拒绝「请使用 Openpyxl 入口」
  │     ├── build_config_print_plans（按目标分组；重复 Sheet 清空重写语义）
  │     └── write_static_print_plans
  │           ├── calamine 读源（值/公式缓存值/日期系统/合并/行高列宽）
  │           ├── ooxml_cell_styles::SourceStyles（OOXML 窄读取 styles.xml + cell s 索引）
  │           ├── build_writer_format（descriptor → rust_xlsxwriter::Format，去重）
  │           └── rust_xlsxwriter 写出（值/样式/合并/行高列宽/页面设置/页脚/批注）
  └── Openpyxl 分支：原实现原样（python_override 分层解析保留）
```

### OOXML 样式读取（ooxml_cell_styles.rs，新增模块）

- 只解析 `styles.xml`（numFmts/fonts/fills/borders/cellXfs）、`theme1.xml` 颜色槽位、目标工作表的 cell `s` 索引。
- **不使用 apply\* 属性门控**——openpyxl/Excel 写出的 cellXfs 经常省略它们，numFmtId/fontId/fillId/borderId 是权威来源（实测修正项）。
- roxmltree 遍历必须从 `doc.root_element()` 出发（`doc.root()` 是文档容器节点，直接子元素查找会全部落空——实测修正项）。
- 颜色：RGB / theme（槽位 + tint）/ indexed（64 色表）全支持。
- 去重：CellStyleInfo（Hash/Eq）→ Format 缓存，输出 styles.xml 里仅出现实际用到的格式。

## 4. 自动路由明确关闭

`run_configured_static_print_with_python` 的引擎分发是显式 match：RustWriter → Rust 路径，Openpyxl → Python 路径。两个菜单入口各自传死引擎枚举。任何引擎失败都直接报错，**不切换到另一个引擎**。

## 5. Rust 输出路径规则

- 原目标 `报表.xlsx` → Rust 输出 `报表_Rust.xlsx`（同目录，仅追加 `_Rust` 后缀）。
- 重映射在目标校验**之前**完成，使「目标已存在」校验和写入都指向独立路径。
- `rust_output_path` 幂等（已带 `_Rust` 不重复追加）。
- Rust 专用输出已存在时明确拒绝（沿用现有安全拒绝机制），不覆盖任何文件。
- XLSM 目标在校验阶段即拒绝：`Rust 引擎暂不支持 XLSM 输出，任务未执行，请使用 Openpyxl 入口`。

## 6. 配置兼容情况

15 列「打印配置」字段含义、默认值、Y/是/1 兼容、源工作表精确匹配（不 trim）、备注不参与逻辑——全部沿用 `load_print_tasks` 既有解析，未新增配置字段。

## 7. 三种模式实现情况（Rust 入口）

| 模式 | 状态 | 实现要点 |
| --- | --- | --- |
| mode 1 | ✅ | 全样式（font/fill/border/align/numfmt）+ 行高 + 列宽 + 合并 + 批注 |
| mode 2 | ✅ | mode 1 全部 + 紧凑列宽自适应（`static_print_display_length`，min 6.5/max 18.0）+ 15.0 行高 |
| mode 3 | ✅ | 值 + font/fill/align/numfmt（无 border/行高/列宽/批注）——对齐 PyCode print_310 快速语义 |

mode 差异通过 `ImportRange` 的既有标志传递：copy_row_heights (mode==1)、copy_column_widths (mode 1|2)、style_mode (Full/Basic)、copy_comments (!drop && mode≠3)。样式读取侧按 style_mode 过滤边框（Full 才含 border）。

## 8. 批注支持

rust_xlsxwriter 0.98.2 通过 `insert_note(Note::new(text))` 支持批注写入。Rust 路径：copy_comments 启用时，从 calamine/excel-reader 读取的 `SheetDump.comments` 按坐标平移写入 `insert_note`。作者信息暂不支持（excel-reader 的 comments 是 author 前缀文本），已记录为限制。

## 9. 公式缓存值红线

- 只取缓存值（calamine），不复制公式文本，不自动重算，不调 BetterOffice。
- 缓存缺失 → 空白（openpyxl 路径 data_only 返回 None 的行为一致）。
- 缓存缺失计入 issues 的机制沿用。
- **运行摘要固定说明**（两引擎共用）：
  `公式缓存说明：公式缓存取自源文件保存时状态，本次打印未重新计算。`

## 10. 本轮对拍中修正的 Rust 侧缺陷

以下问题在补跑宽表/日期夹具时被逐格对拍暴露，均已修复并补回归测试：

| # | 缺陷 | 根因 | 修正 |
| --- | --- | --- | --- |
| 1 | 默认字体丢失（宋体→Calibri） | `fontId=0` 被当作「无字体」跳过；Excel 烘焙源的 font 0 就是工作簿默认字体（中文环境为宋体） | 按 `fontId` 恒查表，不再对 0 特判 |
| 2 | 1904 簿时间单元格偏移 | 1904→1900 纪元平移对纯时间值不适用 | 新增 `numfmt_is_time_only`：纯时间格式且串行值 < 1 时不平移；含整天数（公式 +1）照常平移 |
| 3 | 打印区域少一行 | Rust 侧 `final_row = next_row - 2`，openpyxl 引擎是 `out_row - 1`（含块间空行） | 改为 `next_row - 1`，与 openpyxl 引擎一致 |
| 4 | 颜色被静默丢弃 | OOXML 是 8 位 ARGB，rust_xlsxwriter 只接受 6 位 RGB，越界后 `is_valid()` 判否并**丢色**（同时打 stderr 噪声） | 新增 `writer_rgb6`：写入前截掉 alpha |
| 5 | 列宽全部漂移 +0.71 | rust_xlsxwriter 的 `set_column_width` 走「字符宽→像素→再转回」往返，12.0 写成 12.7109375 | 新增 `set_column_width_exact`：反解像素走 `set_column_width_pixels`，写出与 openpyxl 完全相同的 width 值 |
| 6 | 未设列宽缺省不一致 | openpyxl 对未显式设置的列取 `DEFAULT_COLUMN_WIDTH=13.0` 并落盘；分组列（`min..max`）只有 min 字母命中 | 列宽表按 `<col min>` 单键收录；逐列回落到 13.0，与 openpyxl 引擎逐列语义一致 |
| 7 | 源行高/列宽未复制 | calamine 不暴露行高列宽，原实现只复制了值与样式 | 新增 `SheetLayout`（窄读 `<col>`/`<row>`），mode 1/2 按 openpyxl 语义复制 |

## 11. 已知限制（Rust 入口）

1. XLSM 输出明确拒绝（提示走 Openpyxl 入口）。
2. 追加到既有目标：Rust 输出走独立 `_Rust` 新路径，因此天然从新簿开始；对既有 Rust 输出拒绝覆盖。
3. 批注作者信息不单独保留（文本含 author 前缀时随文本带入）。
4. theme 颜色 tint 算法为标准近似公式。
5. mode 2 的列宽自适应沿用 `static_print_display_length`（日期按 19 字符估算），未按数字格式精确计算显示宽度（与 Openpyxl 入口现状一致）。
6. **Openpyxl 引擎的既有缺口（Rust 已按 print_310.py 语义实现，故两引擎 mode 2 输出不同）**：
   - 内嵌 Openpyxl 脚本没有 `print_310.py` 的 `_autofit_block_columns` / `_compact_block_row_heights`
     （实测：脚本中 `autofit`/`compact`/`0.95` 出现次数均为 0），mode 2 只复制源列宽、不做紧凑自适应；
   - 内嵌脚本的 `parse_orientation` 在配置留空时返回 `None`，Excel 落为纵向，不做「按区域宽高比自动判断」。
   本轮按目标「Rust 以 print_310.py 真实源码语义为准」「不得改变原有 Openpyxl 实现」处理：
   Rust 保留自适应与自动方向，Openpyxl 入口行为原样不动，差异在验收报告中如实列出。

## 12. 性能观测开关

设置 `PYTOOLS_PRINT_STAGE_TIMING=1` 后，Rust 引擎会把分阶段耗时以
`[3-1][stage-timing] {json}` 形式输出到 stderr（不影响正常输出与日志格式），
覆盖目标 §14.4 的 12 项阶段：配置读取、源工作簿打开/数据读取、OOXML 解析、
样式映射、区域提取、区域重组、目标工作簿创建、写入、打印设置、保存、总耗时。

