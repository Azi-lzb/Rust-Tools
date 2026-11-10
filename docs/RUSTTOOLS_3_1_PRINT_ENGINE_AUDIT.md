# rustTools 3-1 报表打印引擎审查报告（里程碑 1）

状态：**源码调查与方案比较完成，待人工确认**。未经确认未修改任何 3-1 业务代码。
关联：`docs/EXCEL_MODIFY_ENGINE_PLAN.md` 系列（通用修改引擎 Benchmark，本轮不重复其结论，仅引用）。

---

## 1. 当前真实业务流程（源码追踪，非文档推断）

### 1.1 调用链

```
主菜单 3（RustTools.rs:545,574,625）
  → run_interactive_print_menu（RustTools.rs:5144）
    → 子菜单 1「按打印配置生成：openpyxl」（RustTools.rs:5180）
      → run_configured_static_print_menu（RustTools.rs:4704）——日志/动画/事件追踪
        → static_print_engines::run_configured_static_print（static_print_engines.rs:1057）
          → load_print_tasks（calamine 读配置）+ 目标存在性拒绝 + 目标扩展名校验
          → run_openpyxl_engine（static_print_engines.rs:687）
            → 内嵌 Python 脚本（-c 传整段源码 + 临时 plan.json）
```

### 1.2 配置结构（config.rs:94 PrintTask）

`打印配置` 工作表列：是否启用 / 打印模式(1|2|3) / 源工作簿 / 源工作表 / 源工作表打印区域(可多段 `;`) / 目标工作簿 / 目标工作表 / FitToPagesWide / FitToPagesTall / 打印方向(横向|纵向) / 水平居中 / 垂直居中 / 不输出批注 / 零值不输出 / 备注。

### 1.3 openpyxl 脚本逐环节（static_print_engines.rs:712-939）

| 环节 | 实现 | 库 |
| --- | --- | --- |
| 源工作簿打开 | **每簿加载两份**：样式版 + `data_only=True` 缓存值版 | openpyxl |
| 源区域解析 | `parse_ranges`：`;` 分段 → range_boundaries；空则用 used_range | openpyxl |
| 目标工作簿 | 已存在则加载（xlsx/xlsm），否则新建并删默认 Sheet | openpyxl |
| 区域复制 copy_block | 值取自 data_only 版（公式缓存值），样式取自样式版；逐格复制 | openpyxl |
| mode 语义 | 1=完整（行高+列宽+全样式+边框+批注）；2=完整+紧凑列宽/行高；3=快速（无行高列宽，样式缺边框/保护/批注） | 对齐 PyCode print_310.py |
| 零值不输出 | 数值 0 → None；**文本 "0" 保留** | 脚本内 is_zero_like |
| 合并区域 | **仅完全落在块内的合并复制**，值取左上角；跨块边界的静默跳过 | openpyxl |
| 页面设置 apply_setup | orientation / fitToPage / fitToWidth / fitToHeight / paper=A4(9) / print_area=A1:末 / 水平垂直居中 | openpyxl |
| 页脚 | 每工作表 `第 {工作表序号} 页 / 共 {总表数} 页`（是**表序号**而非页码） | openpyxl oddFooter |
| 保存 | 新文件 `xb` 独占创建；同目标多任务追加时二次 save | openpyxl |

### 1.4 引擎使用现状

| 引擎 | 3-1 中状态 | 说明 |
| --- | --- | --- |
| **openpyxl** | **现役**（菜单 1/2 两入口统一走它） | 唯一实际执行引擎 |
| calamine | 仅用于读打印配置表 | 不参与源工作簿读取 |
| rust_xlsxwriter | 有完整内部实现 `run_rust_writer_engine` + `print_ops.rs`（1539 行），**菜单未路由** | `StaticPrintEngine::RustWriter` 枚举存在但两入口均传 Openpyxl |
| umya | `run_umya_engine` 历史实现，**菜单未路由** | 保留供旧结果定位 |
| BetterOffice | 未参与 | — |
| 手写 OOXML | 未参与 3-1 | — |

### 1.5 关键发现：Rust 内部实现已存在且覆盖大半业务

`print_ops.rs` 的 rust_xlsxwriter 路径已实现：

- calamine 读源（值 + `datetime_serial` 公式缓存值 + 1900/1904 日期系统转换）
- 区域提取/多区域堆叠（行间空一行）、合并区域平移复制（块内）、行高列宽复制
- mode 1/2 差异（mode==1 复制行高；mode 1|2 复制列宽）、mode 2 紧凑列宽（`static_print_display_length` 自适应）
- 隐藏零值（数值 0 → 空白）、页脚、页面设置全套（orientation/fit/paper/print area/居中）
- 同目标多任务合并为单个 ExcelEditPlan 一次写入

**已确认的差距（与 openpyxl 脚本对齐口径）**：

1. **不复制源单元格 number_format**——datetime 写死 `yyyy-mm-dd h:mm:ss`（print_ops.rs:1034），百分比/千分位/前导零/自定义格式全部丢失。
2. **不复制字体/填充/边框/对齐**——纯值输出，打印观感与源模板差异大。
3. **仅支持新建目标**（目标已存在即拒绝），openpyxl 引擎支持追加到既有 xlsx。
4. **mode 3 语义未实现**（快速模式样式差异）。
5. 页脚语义是字面文本（openpyxl 也是表序号字面量，此处两引擎一致，但都不是真页码）。

## 2. 性能与瓶颈分析（基于代码结构 + 已有 Benchmark）

- openpyxl 脚本每簿**双份加载**（样式版+值版）是最大固定成本；此前通用 Benchmark 显示其打开+导入约 250–330ms 固定开销，修改阶段反而快。
- 脚本以 `-c` 整段传入 + 临时 plan.json：进程启动 + openpyxl 导入 + 双份解析，小文件场景总耗时被固定成本主导（上一轮 Benchmark：openpyxl 原生 w1 中位 781ms，其中 modify 仅 0.1ms）。
- 常驻批量无可行方案：生产 worker 是单请求协议（上一轮报告已确认）。
- 3-1 的写入目标是**新建工作簿**（追加是少数场景），恰好落在 rust_xlsxwriter 的最强区（新建+写入），避开 openpyxl 的保存慢区。

## 3. 五种候选方案能力比较

### 方案 A：当前实现（openpyxl 全样式复制）
- 数据正确（双工作簿口径）、格式最接近源模板、mode 1/2/3 语义完整。
- 依赖 Python 运行时；批量受进程启动成本主导；页脚是表序号不是页码。

### 方案 B：calamine + rust_xlsxwriter
- 即把**已存在但未路由的 `run_rust_writer_engine` 接回菜单**，再补齐差距。
- 已有：值/缓存值/日期系统/合并/行高列宽/页面设置/页脚/零值/mode1/2 列宽。
- 缺口：源 number_format、字体/填充/边框/对齐的复制（保真第二优先级的核心）、追加到既有目标、mode 3。
- calamine 读值已验证足够（公式缓存值、日期系统、错误值均有现成通道）；**但 calamine 不暴露 number_format/样式**。

### 方案 C：calamine + OOXML + rust_xlsxwriter
- 在 B 之上增加窄范围 OOXML 补充读取：只解析源区域涉及的 `styles.xml`（numFmt/font/fill/border/alignment）与 cellXfs 索引 → 映射为 rust_xlsxwriter Format。
- 本项目已有同款先例：`read_date_formatted_cells`（summary_ops.rs）就是 styles.xml 窄读取；excel-reader 已读 styles 部件。
- 合并区域/行高列宽 calamine+excel-reader 已覆盖，无需额外 OOXML。

### 方案 D：UMYA 独立完成
- umya 有完整读取+写入 API（合并/样式/行高列宽/页面设置全有），`run_umya_engine` 历史实现已验证「整簿复制」可行。
- 但 3-1 需要**跨工作簿区域重组**（多源区域→新目标），umya 的跨簿样式复制正是其 80.4% 保真失分的重灾区（跨簿样式索引引用问题，PyCode 注释也确认过这一点）。
- umya 无 Python 依赖、常驻友好，但输出保真需专门对拍验证。
- 页面设置：umya 有 PageSetup API（此前探针 M11 PASS）。

### 方案 E：rXLS + rust_xlsxwriter
- rxls 读取快（w1 167ms），但：无样式/数字格式读取 API（M05/M06 UNSUPPORTED 只证明写侧缺失，读侧也未见公开 API——需探针确认）；100k 打开超限。
- 3-1 源文件常为大报表；读取优势不抵样式读取缺失 + 新依赖成本。**初步标记为不推荐**，除非 rxls 探针证实有 styles 读取 API。

## 4. 公式缓存值处理现状与风险

- openpyxl 脚本：`data_only=True` 取缓存；缓存缺失时值为 None（静默空）——**现状已存在**「缓存缺失显示空白」的行为，且不重算。
- rust writer：calamine 返回 `datetime_serial` 等缓存类型，同样不重算；与现状一致。
- 两个引擎都不把公式文本当值（符合业务红线）。
- 缓存过期的既有兜底：1-4/BetterOffice 是独立菜单功能，3-1 不自动触发——保持现状，报告中明示。

## 5. ### 问题与列宽（业务验收重点）

- 现状 openpyxl：mode 1 直接继承源列宽（源列宽对目标列生效）；mode 2 走 `_autofit_block_columns`（紧凑自适应，min 6.5 / max 18.0）。
- rust writer：mode 2 已有 `static_print_display_length` 自适应等价物；mode 1 继承源列宽。
- 差距：两引擎都不检查「日期列宽不足→###」；日期显示宽度 openpyxl 依赖源列宽、rust writer 写死 19 字符估算。**这是可以双侧共同改进的点，但属行为变化，须走里程碑 2 验证。**

## 6. 风险与限制汇总

| 项 | 风险 |
| --- | --- |
| rust writer 无样式复制 | 直接路由会造打印观感回退——必须先补 OOXML 样式读取（方案 C） |
| 追加到既有目标 | rust writer 拒绝已存在目标；openpyxl 支持——方案 B/C 需保留 openpyxl 兜底该场景 |
| 页脚语义 | 两引擎都写表序号字面量，用户若期望真页码需行为变更（单列） |
| 手工外链/批注夹具脆弱 | 上一轮 COM 冒烟中三方 edit1 同败，复测时需换真实样本 |
| rxls 读取 | 无样式读取 API 迹象；大文件单部件上限；**初步不推荐** |

## 7. 五方案结论表

| 方案 | 可行性 | 工作量 | 预期收益 |
| --- | --- | --- | --- |
| A 现状 | 基线 | 0 | 全能力，慢，依赖 Python |
| B calamine+rxw | 高（大部分已实现） | 小：路由 + 修差距 | 去除 openpyxl 依赖（追加场景除外），速度大幅提升 |
| C B+OOXML 样式 | 高（有同款先例） | 中：styles.xml 窄解析 + Format 映射 | 打印观感对齐源模板，保真第二优先级达标 |
| D umya 独立 | 中（跨簿样式复制是历史弱区） | 中：新写区域重组 | 去 openpyxl；保真需专门对拍 |
| E rxls | 低 | — | 初步不推荐 |
