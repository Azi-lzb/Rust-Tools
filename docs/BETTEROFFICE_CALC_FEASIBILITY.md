# betteroffice-xlsx 作为计算引擎的可行性验证

日期：2026-11-08
探针位置：`C:\Users\AZI\AppData\Local\Temp\bo-probe`（临时，未进仓库）
库：`betteroffice-xlsx` / `betteroffice-xlsx-calc` 0.5.1，Apache-2.0，github.com/openooxml/betteroffice

## 1. 结论

**可以作为 1-4 校验区域重算的纯 Rust 引擎**，但不建议立刻用它整体替换 Excel/WPS COM 与 LibreOffice 路由。

判定依据：算得对、写得保真、失败可察觉、离线可构建；唯一硬限制是**重算预算写死 1000 万次单元格访问、不可调**，大簿会静默保留缓存值——这必须靠 `budget_exhausted` / `limited_cells` 显式拦截。

## 2. 验证结果

### 2.1 函数覆盖（46 条断言，全 PASS）

样本 `calc-battery.xlsx`：33 条公式 + 13 条跨函数链，覆盖用户点名的易失与查找函数。

| 类别 | 覆盖 | 结果 |
| --- | --- | --- |
| 易失函数 | `NOW()` `TODAY()` `RAND()` `RANDBETWEEN()` | ✅ 需注入 `now_serial` |
| 间接引用 | `INDIRECT("数据!C4")`、拼接+`MATCH`、跨表间接链、带空格表名 `'汇总 表'!A1` | ✅ 全对 |
| 查找 | `INDEX`（一维/二维/定义名）、`MATCH`、`VLOOKUP`、`OFFSET`（含 `COUNT` 动态高度） | ✅ 全对 |
| 聚合 | `SUM` `SUMIF` `SUMIFS` `SUMPRODUCT`（含数组条件）`AVERAGE` `MAX` `MIN` `LARGE` `SMALL` `COUNTA` `COUNTIF` | ✅ 全对 |
| 逻辑/文本 | `IF` `AND` `IFERROR` `CHOOSE` `TEXT`（含 `"yyyy-mm-dd"`）`CONCATENATE` `MOD` `ROUND` | ✅ 全对 |
| 定义名 | `SUM(单价区域)`、`INDEX(单价区域,2)`、区域名、公式名 `合计` | ✅ 全对 |

### 2.2 易失函数必须注入时间

`CalculationOptions { now_serial: None }` 时 `NOW()`/`TODAY()` 返回 `#VALUE!`；注入 `Some(46300.5)` 后精确等于 46300.5 / 46300.0，`TEXT(TODAY(),"yyyy-mm-dd")` 得 `2026-10-05`（serial 对应的日期）。

**这意味着引擎没有"当前时间"概念，调用方必须给。** 对业务是好事（可复现），但接入时必须显式传值，否则整片日期公式变 `#VALUE!`。

### 2.3 不支持的函数：显式错误，不静默错值

40 条边界样本（`XLOOKUP`/`TEXTJOIN`/`IFS`/`MAXIFS`/`SEQUENCE`/`LET`/`FILTER`/`UNIQUE`/`SORT`/`RANDARRAY`/`CELL`/`INFO`/`HLOOKUP` 等）：

| 形态 | 返回值 |
| --- | --- |
| 完全不存在的函数 `NO_SUCH_FUNCTION_XYZ` | `#NAME?` |
| 未实现的现代函数（`FILTER`/`SORT`/`SEQUENCE`/`RANDARRAY`） | `#VALUE!` |
| `CELL("address",…)` / `INFO("recalc")` | `#NAME?` |
| 数组溢出 `TRANSPOSE` / `MMULT` 单格上下文 | `#VALUE!` |

**没有一条返回 0 或旧值**——都是 Excel 语义内的显式错误码，能被候选校验捕获。

### 2.4 真重算（依赖链证明）

改 `数据!C2` 10→1000 后：

| 格 | 前 | 后 | 说明 |
| --- | --- | --- | --- |
| `SUM(数据!E2:E6)` | 625 | 2110 | 直接依赖 |
| `SUMPRODUCT(C,D)` | 625 | 2110 | 数组依赖 |
| `SUM(OFFSET(...,COUNT(...)))` | 150 | 1140 | **OFFSET 动态区** |
| `合计`（定义名公式） | 625 | 2110 | **定义名链** |
| `IF(A7=625,…)` | `"OK"` | `"BAD"` | 二阶依赖 |

改 `数据!C4`→777 后 `INDIRECT("数据!C4")` 精确变 777。

**不是信任缓存值**：`changed` 与依赖传播都真实发生。

### 2.5 回写保真（关键，超出预期）

`open` → `save()`（不重算）→ 与原件比对 ZIP 包：

| 工作簿 | 部件数 | 丢失 | 新增 | 大小变化 |
| --- | --- | --- | --- | --- |
| 光大银行月报（23 部件） | 23→23 | 0 | 0 | 0 |
| 目标用例（13 部件） | 13→13 | 0 | 0 | 0 |
| 省信贷收支（41 部件，7.1MB） | 41→41 | 0 | 0 | 0 |

重算后公式文本：**73807 条公式格全部逐字一致**（丢失 0 / 新增 0 / 文本变化 0），zip 层 `<f>` 标签计数逐 sheet 相同。

原因：库内部保留 `SourceContainer` 原始字节，只对改动格打补丁，不重建包。**这不是 rust_xlsxwriter 那样的"重写整个工作簿"。**

### 2.6 真机数据对拍

| 工作簿 | 公式格 | 完全一致 | 差异 | 预算 |
| --- | --- | --- | --- | --- |
| 光大银行月报（1-4 源） | 614 | 614（100%） | 0 | 未耗尽 |
| 目标用例（INDIRECT 密集） | 175 | 175（100%） | 0 | 未耗尽 |
| 省信贷收支（7.1MB） | 73807 | 73798（99.9878%） | 9 | **耗尽，71095 格受限** |

省信贷 9 处差异全部是错误码形态：Excel `#DIV/0!`/`#REF!` vs betteroffice `0`/`#VALUE!`——**引擎在错误传播上与 Excel 有语义差异**。

### 2.7 硬限制：重算预算不可调

- `MAX_RECALCULATION_CELL_VISITS = 10_000_000`，`const`，公开 API 无参数可改。
- 省信贷样本因 `SUMIFS(INDIRECT($M8&"!"&$N8&":"&$N8), INDIRECT(M8&"!B:B"), I$6)` 这类**整列引用**（`B:B` = 104 万格）而被吃光。
- 预算耗尽后：公式**保留缓存值**或显示 `#NUM!`，`budget_exhausted=true`、`limited_cells` 列出具体格。

**这是接入的唯一红线**：必须检查 `budget_exhausted`，非空即 fail-closed，不能把"保留缓存值"当"算对了"。

### 2.8 工程可行性

| 项 | 结果 |
| --- | --- |
| 许可 | Apache-2.0（兼容商用/闭源分发） |
| `--offline` 构建 | ✅ 干净 target 目录通过（依赖已在 `~/.cargo/registry`） |
| `default-features = false`（去 raster） | ✅ 计算功能不受影响，二进制 5.9MB → 4.8MB |
| build.rs / C 依赖 / SIMD 硬需求 | 无（纯 Rust，可交叉编译 Linux） |
| 依赖规模 | 196 个 crate（含 `yrs` CRDT、`tiny-skia`、`rustybuzz`） |

## 3. 建议接入方式

### 3.1 短期：作为 1-4 的第一优先引擎（Windows + Linux 统一）

1-4 的真实负载（校验区域，614 / 175 条公式）**远低于预算**，且这是唯一跨平台可用的纯 Rust 路线：

```
校验区域复制 → ooxml-patch（不变）
   ↓
betteroffice 重算（纯 Rust，无外部依赖）
   ├─ budget_exhausted == false 且 limited_cells 为空 → 采用
   └─ 否则 fail-closed → 回退 COM / LibreOffice
   ↓
Calamine 回读校验区域（不变）
```

好处：Linux 不再需要 LibreOffice（当前 deb 必须用户自备 LO + python3-uno）；Windows 不再需要 Excel/WPS 才能算。

### 3.2 必须的防护（缺一不可）

1. **`budget_exhausted` / `limited_cells` 非空即失败**，不得静默采用。
2. **`now_serial` 显式注入**，取系统时间换算成 Excel serial；不要传 `None`。
3. **落盘前后 SHA-256 快照**（沿用现有安全提交）。
4. **候选校验**：重算后校验区域不得出现 `#NAME?`/`#VALUE!`（否则说明模板用了引擎未实现的函数，必须回退 COM）。

### 3.3 保留 COM / LO 作为回退

不要删除现有路由。理由：
- 错误码语义差异（`#REF!` vs `#VALUE!`）可能影响业务判定；
- 未实现的现代函数（`XLOOKUP`/`FILTER`/`LET` 等）在真实模板里一旦出现就是全区域 `#NAME?`；
- Excel 特有的 `CELL`/`INFO` 函数缺失。

## 4. 未验证项（如实记录）

- 未在 Linux 上实际构建/运行（仅确认无平台特定构建脚本；交叉编译未实测）。
- 未对「重算后 Excel COM 打开」做完整覆盖（仅对 1-4 源做过一次，通过）。
- 未验证含宏 `.xlsm`、含图表/条件格式工作簿的重算保真（仅验证了包结构不丢部件）。
- 未验证 `calcChain.xml` 处理（重算后该部件是否需同步更新/删除未测）。
- 未做与 Excel COM 的**同输入同输出**逐格对拍（当前对拍的是"引擎重算 vs 文件缓存"，不是"引擎 vs Excel 实时重算"）。
- 省信贷 9 处差异未逐条定性（已确认是错误码形态差异，但未确认业务上是否等价）。

## 5. 下一步（需用户决定）

1. 是否把 1-4 的计算路由切到 betteroffice（保留 COM/LO 回退）？
2. 若切换，是否同时移除 Linux 的 LibreOffice 依赖（deb 可去掉 `python3-uno` + `soffice` 前置）？
3. 是否需要我补做「同输入 → betteroffice vs Excel COM 实时重算」的逐格对拍（需要 Excel 在环境里跑一遍，较慢）？
