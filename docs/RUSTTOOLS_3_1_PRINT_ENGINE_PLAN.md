# rustTools 3-1 报表打印引擎优化计划（里程碑 1 交付）

前置：`RUSTTOOLS_3_1_PRINT_ENGINE_AUDIT.md`（现状审查）。本计划为推荐路线，**待人工确认后进入里程碑 2**。

---

## 1. 推荐技术架构（分两步走）

### 第一步：方案 B 落地——把已存在的 rust_xlsxwriter 引擎接回 3-1

`run_rust_writer_engine` + `print_ops.rs` 已实现 3-1 核心业务的约七成，当前问题是**菜单从未路由到它**。第一步只做路由切换与差距补齐：

```
菜单 3-1 → run_configured_static_print
  ├─ 目标为新建 .xlsx 且 mode 1/2 → run_rust_writer_engine（Rust 原生，无 Python）
  └─ 追加到既有目标 / .xlsm 目标 / mode 3 → run_openpyxl_engine（保留现状兜底）
```

补齐项（对齐 openpyxl 脚本口径）：
1. rust writer 支持追加到既有 xlsx 目标（当前拒绝已存在目标）。
2. rust writer 补 mode 3 语义（无行高列宽、精简样式）。
3. 页脚语义双引擎统一（维持现状表序号字面量，不改行为）。

### 第二步：方案 C——OOXML 窄范围样式读取，补齐打印观感

仅当第一步对拍确认「无样式」影响验收时实施。范围严格限定：

```
源 xlsx → [calamine: 值/缓存值/合并/行高列宽]
        → [新模块 ooxml_styles_read: styles.xml 窄解析]
             cellXfs → numFmt/font/fill/border/alignment（仅目标区域涉及的 xf 索引）
        → rust_xlsxwriter Format 重建（去重复用）
        → 打印工作簿
```

- 复用既有模式：`read_date_formatted_cells`（summary_ops.rs）已示范 styles.xml 窄读取 + JobObject 无关。
- **不解析**无关 sheet 的样式、theme 完整解析、dxf、差异化格式。
- theme/indexed/tint 颜色解析按需最小化（Color::Rgb 直读 + theme 槽位查表）。

## 2. 各引擎职责

| 引擎 | 3-1 中的职责 |
| --- | --- |
| calamine | 配置读取 + 源值/公式缓存值/日期系统/合并/行高列宽读取 |
| excel-reader 既有补充 | 日期格式元数据（已有）；新增 styles.xml 窄解析（第二步） |
| rust_xlsxwriter | 新建目标 + 值写入 + Format + 合并 + 行高列宽 + 页面设置/页脚 |
| openpyxl | 兜底：追加既有目标、.xlsm 目标、mode 3 |
| umya | 不参与 3-1（历史实现保留） |
| BetterOffice | 不参与（缓存缺失重算是用户手动走 1-4 的独立行为） |

## 3. 关键技术方案

### 3.1 公式缓存值

- 与现状一致：只取缓存，不复制公式文本；缺失时留空白并计入 issues（不静默补 0、不用公式文本顶替、不自动重算）。
- 缓存过期提示：运行摘要固定附注「公式缓存取自源文件保存时状态」。

### 3.2 样式映射（第二步）

- 读取：解析 styles.xml 的 `numFmts/cellXfs/fonts/fills/borders`，按目标区域实际出现的 xf 索引过滤。
- 映射：cellXfs → `rust_xlsxwriter::Format`（num_format/font bold·italic·size·color/fill color/border 四边/align_h/align_v/wrap_text）。
- 复用：以等效 xf 属性组合为键去重，避免 Format 爆炸。
- 已知限制：theme 颜色需要 theme1.xml 槽位表（最小解析 6 槽）；dxf/条件格式样式不支持（3-1 不要求）。

### 3.3 合并区域

- 维持现有规则：完全在块内 → 平移复制；跨块边界 → 计入 issues 不静默截断；值取左上角。
- 转置模式（如未来配置支持）需坐标交换后重建——当前业务无转置，标注为不适用。

### 3.4 ### 问题

- mode 1：继承源列宽（现状）——日期列宽随源，通常已足够。
- mode 2：`static_print_display_length` 自适应（现状），日期固定按 19 字符估算（datetime_serial 特判已有）。
- 第二步增强：自适应时**按 number format 计算显示宽度**（如 `0.00%` 乘 100 加百分号），并对日期格式用格式化后文本长度而非 19 定值——均为纯函数改动，行为可对拍。
- 明确不做：全列自动调整、数字转文本、字体压缩。

### 3.5 打印配置

- 全部走 rust_xlsxwriter 原生 API（orientation/fit_to_pages/paper 9/print_area/centered/footer）——已在 print_ops 实现并验证，无 OOXML 需求。
- 页眉页脚编码：rust_xlsxwriter `set_footer` 写 UTF-8；页码如需真页码变量用 `&P`/`&N`（行为变更须用户确认，本轮默认维持现状）。

## 4. 迁移步骤

1. **M2-1**：搭独立对拍工程（复用 excel-modify-engine 的夹具生成模式），用真实配置 + 脱敏模板。
2. **M2-2**：方案 B 差距补齐（追加目标 / mode 3）在独立分支实现，菜单加内部开关对比双引擎输出。
3. **M2-3**：正确性对拍：值/类型/合并/行高列宽/打印设置逐项 diff（openpyxl 输出为参照物），Excel COM 开启 + PDF 页数/截断/### 检查。
4. **M2-4**：10/50/100 簿批量性能（openpyxl 现状 vs rust writer），记录阶段耗时与峰值内存。
5. **M2-5**：若观感验收不达标 → 实施第二步 OOXML 样式读取 → 重跑 M2-3/M2-4。
6. **M3 迁移**：菜单路由默认切换，openpyxl 引擎保留为兜底；保留回滚开关。

## 5. 回滚方案

- 菜单路由是单点改动（StaticPrintEngine 枚举参数），回滚 = 恢复传入 Openpyxl。
- openpyxl 引擎代码与脚本全程保留，不删除。
- rust writer 差距补齐均为新增分支逻辑，不改动现有 print_ops 主路径。

## 6. 验收标准

1. 数值/文本/日期/错误值/缓存值与现状 openpyxl 输出逐格一致（日期系统转换含在内）。
2. mode 1/2 打印观感：字体/填充/边框/对齐/数字格式与源模板一致（第二步后）。
3. 合并/行高/列宽/打印区域/纸张/方向/居中/页脚一致。
4. Excel COM 打开无修复提示，PDF 页数与分页与现状一致，无 ###。
5. 批量 100 簿耗时显著优于现状（预期 ≥3x，以实测为准）。
6. 现有测试（print_menu_cli 等）零回归；新实现有独立测试。
