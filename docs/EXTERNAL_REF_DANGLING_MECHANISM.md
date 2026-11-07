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

| 仓库 | 路径 | 是否剥离 `[N]` | 状态 |
| --- | --- | --- | --- |
| Rust-Tools | `ooxml-patch::copy_cell_region`（1-4 工作副本） | 依赖 4-7 前置修复；副本生成本身不剥离 | 需确认调用顺序 |
| Rust-Tools | `ooxml-patch::repair_excel_com_open_issues`（4-7） | ✅ **已实现** | 已验收 |
| Rust-Tools | `ooxml-patch` 表头补丁 / calcChain 清理 | 不涉及公式原文改写 | 无风险 |
| 旧 audit | `formula_region_writer.py` | ❌ 原样保留 | **未修（建议同口径补齐）** |
| 旧 audit | `external_sheet_writer.py`（DIRECT_OOXML） | ❌ 原样保留 | **未修（建议同口径补齐）** |
| Rust-Audit | `audit-core/src/excel_open_repair.rs::strip_external_refs` | ✅ 已实现 | 已验收（28/28） |

## 5. 验收

Rust-Tools 合成样本实测（2026-11-07）：

| 检查 | 结果 |
| --- | --- |
| 修复前 Excel COM 打开（含 `[1]` 本地表引用） | ❌ 0x80020009 |
| 4-7 修复后 | ✅ 打开成功，2 表 |
| `[1]` 前缀 | 已剥离，公式 `=SUM(存量单位贷款信息!C5:C10)` |
| `[99]` 真外部引用 | 保留不动 |
| 原件保护 | 只写 `-Excel修复` 副本 |

## 6. 待办

1. 旧 audit 补齐 `strip_external_refs`（`formula_region_writer` + `external_sheet_writer` 两条路径）。
2. Rust-Tools 确认 1-4 管线在 `copy_cell_region` 之后确实经过修复（当前 1-4 已有 `repair_excel_com_open_issues` 调用，需验证顺序）。
