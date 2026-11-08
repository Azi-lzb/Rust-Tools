# S1F1 业务模板计算引擎实测补充（Rust-Audit business_templates）

日期：2026-11-08
承接 `CALC_ENGINE_BENCHMARK.md`（四引擎对比）。本轮按用户指示改用 **Rust-Audit 的 S1F1
真实业务模板**（`测试数据/S1F1_新旧对比_20261105_231620/business_templates`，20 个报送模板）
做打靶，并补充 visi-core、formulas(Py) 两个新候选。

## 1. S1F1 模板负载画像（20 个模板全扫描）

- 公式数：16 ~ 722 条/文件；清缓存后实测最多 1331 格（共享公式子格展开）。
- 嵌套深度 ≤4；函数 2~17 种（INDEX/MATCH/VLOOKUP/INDIRECT/IF/AND/OR/ISBLANK/TEXT/RANK/SUMIF 等）。
- 仅 ！50 票据业务含 116 个整列引用；仅 ！70 债券业务含 12 处 `[N]` 悬空外链（即 4-7 修复样本）。
- 结论：**这是小负载、公式密集、含中文长表名/中文条件文本的真实形态**——正是 1-4/S1F1
  校验区域复制后要重算的东西。

## 2. 参赛样本与清缓存口径

选 4 个代表模板，全部经 `strip_cache.py` 清空公式缓存值（删 `<v>` + `fullCalcOnLoad="1"`）：

| 代号 | 模板 | 公式格 | 特点 |
| --- | --- | ---: | --- |
| danwei100 | ！100.单位存款 | 1331 | 最大负载，含中文条件文本拼接 |
| zhaiquan70 | ！70.债券业务 | 1210 | 12 处 `[N]` 悬空外链 |
| piaojun50 | ！50.票据业务 | 1011 | 116 个整列引用 |
| daifa120 | ！120.代发工资 | 160 | 最小 |

## 3. 实测结果（基准 = Excel COM CalculateFullRebuild）

### 3.1 结果一致率

| 引擎 | danwei100 | zhaiquan70 | piaojun50 | daifa120 |
| --- | --- | --- | --- | --- |
| betteroffice | 99.70% (4) | **99.26% (9)** | 99.90% (1) | **100%** |
| IronCalc | 99.70% (4) | 98.84% (14) | 99.90% (1) | **100%** |
| LibreOffice | **100%** | 98.84% (14) | **100%** | **100%** |
| visi-core | **panic** | 98.84%(未对拍)* | **panic** | 100%* |
| formulas (Py) | 未对拍 | — | — | — |

*visi 只产出文件未逐格对拍（已因 panic 与速度出局）。

### 3.2 差异定性（重要：都不是"算错数"）

- **danwei100 的 4 处（betteroffice 与 IronCalc 完全相同）**：校验文本里拼浮点数时，
  Excel 按隐式转换截到 15 位有效数字（`34.6014278977972`），纯 Rust 引擎输出 17 位全精度
  （`34.60142789779723`）。是**数值→文本的显示语义差异**，判错逻辑（ABS 阈值比较）不受影响。
  LibreOffice 模拟了 15 位截断所以 100%。
- **zhaiquan70 的差异（三个引擎都有）**：全部源于 `[1]集中系统数据!...` 悬空外链公式。
  Excel 靠外链缓存值继续算出 4.112037；三个引擎都给 `#ERROR!`/`#REF!` 级联。
  **这正是 4-7 定义的 fail-closed 场景**——按业务规则该文件本就该先剥离外链，不是引擎缺陷。
- **piaojun50 的 1 处**：同类显示精度差异。

### 3.3 耗时（load + eval + save，秒）

| 引擎 | danwei100 (1331) | zhaiquan70 (1210) | piaojun50 (1011) | daifa120 (160) |
| --- | ---: | ---: | ---: | ---: |
| betteroffice | **0.73** | **0.71** | **0.53** | **0.19** |
| IronCalc | 6.64 | 0.18 | **70.85** | 0.02 |
| LibreOffice | 13.3* | 6.1* | 4.3* | 3.3* |
| Excel COM | 6.59 | 2.90 | 2.85 | 2.48 |
| visi-core | panic | 77.1 | panic | 6.36 |
| formulas (Py) | 28.91 | — | — | 2.51 |

*LibreOffice 含 soffice 进程启动 1.8~5.4s。
IronCalc 在整列引用样本（piaojun50）上 eval 独占 70.8s——它对整列引用逐格扫，
betteroffice 因 used-extent 裁剪而免于此难（budget_exhausted=false）。

## 4. 本轮新增结论

### 4.1 betteroffice 通过 S1F1 真实业务验证 ✅

- 4/4 模板预算充足（`budget_exhausted=false`），亚秒完成，全部输出可被 Excel 打开，
  批注/VML 全保留（此前已验证包保真）。
- 大表失败机制也查清了（见 §5）：**不是嵌套、不是逐格算**，是共享预算
  `公式数 × used_rows` 达到 10M 上限。S1F1/1-4 的负载（≤1331 格）距离上限远。

### 4.2 IronCalc 出局（新增第二个拒开根因）⛔

- 输出 Excel 兼容性不可靠：battery 因非法 `#ERROR!` 错误码拒开（已实锤）；
  **danwei100 在无任何 `#ERROR!` 的情况下仍被 Excel 拒开**（修复模式 CorruptLoad=1 也失败）。
  已排除：`_xlnm._FilterDatabase` definedName（修复无效）、xl/metadata.xml、样式/dxf 越界、
  单元格乱序/重复。根因未定位——这本身就是要害：**无法保证它写出的文件 Excel 能打开**。
- 慢：整列引用样本 eval 70.8s（betteroffice 同样本 0.05s）。
- 定义名求值全败、`RANK` 联合区间不支持（上一轮已实锤）。

### 4.3 visi-core 出局 ⛔

- **中文字符串字节截断 panic**：`xlsx.rs:618` 处 `end byte index 31 is not a char boundary;
  it is inside '限'`——对中文内容直接崩溃（danwei100、piaojun50 两个 panic）。
- 没崩的也慢得离谱：guangda（614 公式）eval **640 秒**；daifa120（160 公式）5.5 秒。
- 函数库虽全（519 个），但工程上不可用。

### 4.4 formulas(Py)：语义最准，工程上不适合 ⛔

- 唯一在 `SUMIF(B2:B6,">=C",...)` 文本比较上与 Excel 一致的引擎（IronCalc/betteroffice 算成 60，正确 120）。
- 但：7KB 文件 load 7.7s；danwei100（1331 公式）28.9s——比 betteroffice 慢 ~40 倍；
  依赖 CPython + numpy + scipy + schedula + openpyxl + xlrd + tqdm，写入用 openpyxl 重写
  （回写保真度=LO 档）。对 Rust-Tools 的离线 deb 是重依赖负担。
- 定位：**参考实现**（用于对拍仲裁），不是生产引擎。

## 5. betteroffice 大表失败机制（2026-11-08 二次修正：样式空格撑大 extent）

**重要修正**：本节初版把省信贷的失败归因于"used extent 本身就是 16,400 行"——表述错误。
经用户指出并实测：**真实内容只到 42~68 行**，16,415 行的 extent 全是被纯样式空单元格撑出来的。
betteroffice 对这类"格式刷满全列"的文件**没有合理截断**，用户判断正确。

### 5.1 三层事实

| 层 | 数字 | 来源 |
| --- | --- | --- |
| `<dimension>` 声明 | 16,384~16,415 行 | sheet XML |
| 纯样式空格末行（无值无公式） | 展示 16,384 / temp 16,415 | 逐行解析 |
| **有值/公式末行** | **42~68 行** | 逐行解析 |

### 5.2 源码实锤

`betteroffice-xlsx-model/src/workbook.rs`：

```rust
pub fn used_range(&self) -> Option<CellRange> {
    let mut bounds = self
        .cells                       // <-- 所有存储格，含纯样式空格
        .keys() ...
}
impl CellProvider for Workbook {
    fn used_rows(&self, sheet: SheetId) -> RowId {
        self.sheet(sheet).and_then(Sheet::used_range)...
    }
}
```

解析器把带样式的空格存进 `cells`，`used_range` 按 `cells.keys()` 取边界——
所以 used_rows = 样式行 16,415，不是内容行 68。

### 5.3 决定性对照实验（唯一变量 = 样式空格）

700 条 `SUMIFS(数据!C:C,...)` + 5,000 行真实数据：

| 实验 | 差异 | calc_ms | limited | budget_exhausted |
| --- | --- | ---: | ---: | --- |
| X2 | 无样式行 | 532 | 0 | false |
| X1 | +15,999 行纯样式空格 | 962 | **109** | **true** |

### 5.4 预算重估（按内容行）

省信贷 1,213 个整列公式 × 各表内容末行（42~68）≈ **80,494 次访问**——仅预算的 0.8%。
**如果按内容范围截断，这本 73,807 公式的册子轻松算完**。
此前"大簿需要 COM/LO"的结论据此推翻：大簿失败是截断策略缺陷，不是负载问题。

### 5.5 嵌套与逐格的两个否定结论仍然成立

- 不是嵌套：三层嵌套对照只影响耗时（435ms vs 164ms），`limited=0`。
- 不是逐格算：`engine/order.rs` SCC + Kahn 拓扑 + 段树区间索引；真重算已实证。

### 5.6 修复路径（三选）

| 方案 | 做法 | 代价 |
| --- | --- | --- |
| A. fork/vendor 修 `used_range` | 区分内容格与样式格，criteria 截断用内容范围 | 修完连省信贷都能算；需 vendor 整个系列或等上游 PR |
| B. 候选预处理剥样式空格 | ooxml-patch ZIP 直写，删除无值纯样式 `<c>`（或压掉纯样式行） | 我们自己代码可做；丢纯样式格的外观格式（工作副本场景可评估） |
| C. 预算拦截回退 COM/LO | `budget_exhausted`/`limited_cells` 非空即回退（原方案） | 保底安全；放弃此类文件的纯 Rust 加速 |

报送模板"格式刷全列"很常见，接入 betteroffice 前必须三选一（可组合：B/C 立即，A 中期）。

## 6. 最终选型建议

| 场景 | 引擎 | 依据 |
| --- | --- | --- |
| 1-4/S1F1 校验区域重算（Windows） | **betteroffice，COM 兜底** | S1F1 4/4 亚秒 + Excel 可开 + 批注保留；差异仅显示精度级 |
| 1-4 校验区域重算（Linux） | betteroffice，LO 兜底 | 同上；可摆脱 LibreOffice 硬依赖 |
| 大簿（含样式刷全列的脏 extent） | COM / LO（短期）；betteroffice（修复截断后） | 见 §5.6 三选 |
| 对拍仲裁 | formulas(Py) 或 LO | formulas 语义最准；LO 已在管线内 |
| ~~通用替换~~ | ~~IronCalc / visi~~ | 拒开根因未定位 / 中文 panic |

接入防护（与上一轮结论一致）：`budget_exhausted`/`limited_cells` 非空即回退；
`now_serial` 显式注入；重算结果出现 `#NAME?`/`#NUM!` 即回退；保留 COM/LO 兜底。

## 7. 未验证项

- ic-danwei100 的 Excel 拒开根因未定位（已排除 5 类假设，修复模式失败）。
- formulas(Py) 未跑 zhaiquan70/piaojun50 对拍（加载慢，仅验证了语义样本与速度量级）。
- visi-core 的两个"可开"输出未做逐格对拍（panic + 速度已足以出局）。
- S1F1 其余 16 个模板未逐个跑（负载画像同量级，抽样已覆盖最大/整列/外链/最小四种形态）。
- 方案 B（剥样式空格）未实施验证；方案 A 未评估 vendor 改动面。
