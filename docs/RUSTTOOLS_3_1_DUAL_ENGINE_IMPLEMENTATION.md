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

## 10. 已知限制（Rust 入口）

1. XLSM 输出明确拒绝（提示走 Openpyxl 入口）。
2. 追加到既有目标：Rust 输出走独立 `_Rust` 新路径，因此天然从新簿开始；对既有 Rust 输出拒绝覆盖。
3. 批注作者信息不单独保留（文本含 author 前缀时随文本带入）。
4. theme 颜色 tint 算法为标准近似公式。
5. mode 2 的列宽自适应沿用 `static_print_display_length`（日期按 19 字符估算），未按数字格式精确计算显示宽度（与 Openpyxl 入口现状一致）。
