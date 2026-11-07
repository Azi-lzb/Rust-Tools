# 悬空 `[N]` 外部引用与 Excel COM 打开失败：机制结论

日期：2026-11-07
适用仓库：Rust-Tools（rustTools）、3.open-data-audit（旧 audit）、Rust-Audit（参考实现）
状态：**已实锤，含两处与初始假设相反的实测结论**

## 1. 缺陷机制（已确认）

报送核查表模板的校验区域公式可直接带外部工作簿前缀 `[1]表名!`。审核副本若：

1. 把外部工作簿导入为**本地工作表**（同名），
2. 却**丢弃** `xl/externalLinks/externalLink1.xml` 与 `workbook.xml` 的 `<externalReferences>`（或模板外链本已断链），
3. 且公式文本**原样保留** `[1]` 前缀，

则公式指向"不存在的外部工作簿 1"。Excel COM `Workbooks.Open` 直接拒绝，HRESULT `0x80020009`（"Open 方法无效"）。WPS 宽松忽略、LibreOffice 同样受影响。

### 1.1 两个仓库的复制来源不同（不要混淆）⚠️

| | 旧 open-data-audit / Rust-Audit | Rust-Tools 1-4 |
| --- | --- | --- |
| 工作表来源 | **外部文件**（独立的 external workbook，按配置导入为本地 sheet） | **模板文件自身**（`plan_missing_template_sheets(template, …)`，把模板里同名依赖表复制进工作副本） |
| 触发 `[N]` 悬空的路径 | 导入外部表后丢弃 externalLink 部件，公式仍带 `[N]` | 校验区域公式原样复制到工作副本，模板的 externalLink 部件与 `<externalReferences>` 不随复制；公式残留 `[N]` 即指向不存在的索引 |
| 本地表名判定基准 | 导入后副本的 `<sheets>`（含新导入的表） | 复制后工作副本的 `<sheets>`（含从模板复制来的表） |

共同点：**只要副本里公式带 `[N]` 而副本没有对应 externalLink 定义，Excel 就拒绝打开**。所以两边都需要处理，但处理位置与判定基准要按各自的复制来源取。

1-4 的处置（2026-11-07 用户决策）：**模板 fail-closed**——扫描到模板含 externalLink 部件或 `[N]` 前缀即拒绝执行，要求用户先清理模板，不做静默剥离（剥离会让规则指向非预期数据，改模板属业务决策）。

## 2. 两处与初始假设相反的实测结论 ⚠️

### 2.1 openpyxl 完整读写周期**不会**剥离 `[N]`

初始假设："openpyxl 周期天然剥离断链外链，不受影响。"

**实测推翻**（2026-11-07，openpyxl 3.1.5）：

```
输入：A1 = "=[1]本地表!B2"，本地表存在 B2=42
openpyxl load_workbook → save 后：
  XML 中 [1] 仍然存在（仅被实体编码为 &#26412; 等）
  该文件用 Excel COM 打开 → 仍报 0x80020009
```

**结论：任何经 openpyxl 保存的副本，只要公式原文带悬空 `[N]`，同样打不开。openpyxl 不是安全网。**

### 2.2 旧 audit 的 `DIRECT_OOXML` 与公式区域写入器**都不剥离** `[N]`

- `core/src/base_audit/formula_region_writer.py::_shift_refs_outside_strings`：测 `[1]存量单位贷款信息!C5` 与 `[1]'…'!C5`，**原样返回**（`!` 前的单元格引用被 `if "!" in segment[:match.start()]` 跳过，既不剥离也不平移）。
- `core/src/base_audit/external_sheet_writer.py`（`DIRECT_OOXML`）：`_render_sheet` 把 `cell.formula` 原样写进 `<f>`，同样无剥离。

旧 audit 122/129 副本可打开，**不是因为机制剥离了 `[N]`，而是因为这批样本本身不含 `[N]` 前缀**（Rust-Tools 侧 129 份全量扫描：`[N]` 命中 0、externalLinks 部件命中 0；1-4 模板同样为 0）。

**结论：旧 audit 的 `DIRECT_OOXML` 与 Rust-Tools 的自研 OOXML 直写路径同属风险面，只是尚未遇到带外链的模板。**

## 3. 修复口径（两边必须一致）

对工作表 XML 的四类公式载体——`<f>`、`<formula>`、`<c:f>`、`<definedName>`——剥离开头/内部的 `[N]` 前缀，**仅当 N 后跟的表名存在于本包 `<sheets>`**（含 `'表名'` 引号与 `''` 转义变体）。指向真正外部簿的引用保持原样。

实现要点（Rust-Tools `crates/ooxml-patch/src/excel_open_repair.rs`）：

- **必须走 XML 解码**（`roxmltree`），不能纯文本正则：生成侧可能写 `&#23384;` 数字字符引用，纯文本匹配会漏。
- 剥离后用公式内容最小转义（仅 `& < >`）写回。
- 预检与修复共用同一实现，无嫌疑不重写。
- 计数入报告：`external_refs_stripped` / `external_refs_total`。

**决策：默认断开所有外链，不保留/复制 externalLink 定义。** 保留外链定义虽可让 Excel 打开，但重算会回落到模板缓存的外部旧值，审核结论将基于错误数据快照。真无关外链剥离后重算变 `#REF!`，按审核异常口径捕获，优于整文件拒绝打开。

## 4. 风险面清单

| 仓库 | 复制来源 | 路径 | 处置 | 状态 |
| --- | --- | --- | --- | --- |
| Rust-Tools 1-4 | 模板文件自身 | `recalculate_named_validation_regions` 入口 | ✅ **模板 fail-closed**：含 externalLink 或 `[N]` 即拒绝 | 已验收 |
| Rust-Tools 1-4 | 模板文件自身 | `ooxml-patch::copy_cell_region` → 工作副本 | 兜底：随后调用 4-7 剥离（顺序已确认 copy → repair → COM） | 已实现 |
| Rust-Tools 4-7 | 调用方副本 | `repair_excel_com_open_issues` | ✅ 剥离 `[N]`（仅指向本包本地表） | 已验收 |
| Rust-Tools | — | 表头补丁 / calcChain 清理 | 不改写公式原文 | 无风险 |
| 旧 audit | **外部文件** | `formula_region_writer.py`（DirectOoxml） | ✅ 生成侧剥离 | 已修（`ba69b35`） |
| 旧 audit | **外部文件** | `external_sheet_writer.py`（DIRECT_OOXML） | ✅ 生成侧剥离 | 已修（`ba69b35`） |
| 旧 audit | — | `excel_open_repair.py`（事后修复） | ✅ 剥离（既有实现） | 已存在 |
| Rust-Audit | **外部文件** | `audit-core/src/excel_open_repair.rs::strip_external_refs` | ✅ 剥离 | 已验收（28/28） |

**关键区别**：旧 audit / Rust-Audit 的 `[N]` 来自"外部文件导入为本地表"；Rust-Tools 1-4 的 `[N]` 来自"模板公式原样复制 + 模板依赖表复制"——模板自己的 `<sheets>` 就是本地表名判定基准，不存在外部工作簿参与。

## 5. 验收

Rust-Tools 合成样本实测（2026-11-07）：

| 检查 | 结果 |
| --- | --- |
| 修复前 Excel COM 打开（含 `[1]` 本地表引用） | ❌ 0x80020009 |
| 4-7 修复后 | ✅ 打开成功，2 表 |
| `[1]` 前缀 | 已剥离，公式 `=SUM(存量单位贷款信息!C5:C10)` |
| `[99]` 真外部引用 | 保留不动 |
| 原件保护 | 只写 `-Excel修复` 副本 |
| 1-4 真实模板（无外链） | ✅ 放行，3 条命中 |
| 1-4 注入 `[1]外部簿!` 的模板 | ✅ 拒绝（exit 2，错误含定位） |
| 旧 audit `_strip_external_prefixes` | ✅ 本地表剥离 / `''` 转义 / 真外链保留 三例通过 |

## 6. 待办

1. ~~旧 audit 补齐 `strip_external_refs`~~ — 已完成（`ba69b35`，含 `formula_region_writer` 与 `external_sheet_writer` 两条生成侧路径）。
2. ~~Rust-Tools 确认 1-4 管线顺序~~ — 已确认：`copy_cell_region` → `repair_excel_com_open_issues` → COM 重算。
3. 旧 audit 仓库存在**既存未提交改动**（`excel_open_repair.py`、`test_excel_open_repair.py`，非本次工作），提交时需与本次改动分开处理。
