# betteroffice 接入 1-4 校验区域重算：vendor 修复 + 可配置回退

日期：2026-11-08
状态：**已实现并端到端验证**（Windows）。Linux 交叉编译与实机未验证。
承接：`CALC_ENGINE_BENCHMARK.md`、`S1F1_ENGINE_TEST_SUPPLEMENT.md`

## 1. 做了什么

把 betteroffice（纯 Rust 计算引擎）接入 1-4 校验区域重算路由，作为首选引擎，
COM（Windows）/LibreOffice（Linux）降为可配置回退。

三处改动：

| 层 | 改动 |
| --- | --- |
| vendor | 修 `used_range` 把纯样式空格算进 extent 的缺陷；重算预算上限改为运行时可配 |
| ooxml-patch | 新增 `strip_style_only_cells`：按配置一次性剥离纯样式空格 |
| pytools-core | 新增 `bo_calc` 模块 + 路由接入 + 回退链 + 日志引擎名按实际取值 |

## 2. 根因回顾（为什么要 vendor）

betteroffice 0.5.1 的 `Sheet::used_range()` 按 `self.cells.keys()` 取边界，而解析器把
**带样式的空单元格**也存进 `cells`——所以"格式刷满全列"的报送模板会把整列引用
（`A:A`、`INDIRECT(...&"!C:C")`）的求值范围撑到样式末行：

| 省信贷样本 | 行数 |
| --- | ---: |
| `<dimension>` 声明 | 16,384~16,415 |
| 纯样式空格末行 | 16,384 / 16,415 |
| **真实有值/公式末行** | **42~68** |

`MAX_RECALCULATION_CELL_VISITS = 10_000_000` 是全簿共享预算，1,213 个整列公式 ×
16,400 样式行 ≈ 1,989 万次访问 → 预算耗尽、41,228 格保留旧值。

**决定性对照实验**（700 条 `SUMIFS(数据!C:C,...)` + 5,000 行真实数据）：

| 实验 | 唯一差异 | calc | limited | budget_exhausted |
| --- | --- | ---: | ---: | --- |
| X2 | 无样式行 | 532 ms | 0 | false |
| X1 | +15,999 行纯样式空格 | 962 ms | **109** | **true** |

## 3. vendor 修改清单

见 `vendor/VENDOR_NOTES.md`。三处：

1. **`xlsx-model`**：新增 `Sheet::content_range()`（只算有值/有公式的格），
   `impl CellProvider for Workbook` 的 `used_rows`/`used_cols` 改用它；原 `used_range` 行为不变。
2. **`xlsx-calc`**：`MAX_RECALCULATION_CELL_VISITS` const → `AtomicU64` 默认值 +
   `set_max_recalculation_visits(u64)` / `max_recalculation_visits()`。
3. 许可证 Apache-2.0，保留上游声明；`[patch.crates-io]` path 替换，版本号仍 0.5.1，
   `--offline` 构建通过（`Cargo.lock` 中两条已无 source/checksum，即指向本地 path）。

vendor 自带测试：model **100 passed**，calc **257 lib + 全部集成测试 passed**
（含新增 vendor 测试：样式不撑 extent、内容行截断求值正确、上限可配）。

## 4. 配置项

| 环境变量 | 默认 | 说明 |
| --- | --- | --- |
| `PYTOOLS_BO_CALC` | `auto`（启用） | `off` 完全回到 COM/LibreOffice 旧路由 |
| `PYTOOLS_BO_CALC_MAX_VISITS` | `0`=库默认 1000 万 | 重算预算上限 |
| `PYTOOLS_BO_CALC_ROUNDS` | `2` | 重算轮数，上限 5 |
| `PYTOOLS_BO_CALC_STRIP_STYLE` | `off` | `on` 时**一次性**剥离纯样式空格（不做事后重试） |
| `PYTOOLS_BO_CALC_FALLBACK` | `auto` | `auto`/`com`/`libreoffice`/`none`；`none` 时超限明确报错 |

### 为什么默认 2 轮

betteroffice 的依赖图**看不到 `INDIRECT` 的动态引用**，跨表动态链第一轮可能读到
尚未更新的值。省信贷实测收敛过程：

| 轮数 | 与 Excel COM 一致率 |
| --- | --- |
| 1 | 63.89% |
| 2 | 98.20% |
| 3 | 98.45% |

S1F1 模板（1-4 的真实负载）**第二轮即 `changed=0`**，2 轮足够。

## 5. 验收结果

### 5.1 端到端（S1F1 真实模板 ！100 单位存款，1331 公式格）

| 配置 | 计算引擎 | 命中结果 | 重算阶段耗时 |
| --- | --- | ---: | ---: |
| `PYTOOLS_BO_CALC=auto` | **betteroffice (pure Rust)** | **20 条** | **6.84 s** |
| `PYTOOLS_BO_CALC=off` | Excel.Application (COM) | **20 条** | 10.59 s |

**命中结果逐条一致，betteroffice 快 35%**（含整个 1-4 管线）。

### 5.2 配置矩阵

| 场景 | 结果 |
| --- | --- |
| `PYTOOLS_BO_CALC=off` | ✅ 走 COM，20 条命中，与旧行为一致 |
| `MAX_VISITS=10`（人为压爆预算） | ✅ 自动回退 COM，20 条命中，日志记录 `betteroffice 回退（重算预算耗尽（1329 个公式格未求值…））` |
| `FALLBACK=none` + 预算不足 | ✅ 明确失败，不静默采用受限结果 |
| `STRIP_STYLE=on` | ✅ 剥离 15,087 格纯样式空壳，内容零损失，仍 20 条命中 |
| `STRIP_STYLE=off` | ✅ 同等结果 |

### 5.3 时间统计（同文件、清空缓存后逐引擎）

单位秒，wall clock（含进程启动与落盘）：

| 样本 | 公式格 | betteroffice | IronCalc | Excel COM | LibreOffice |
| --- | ---: | ---: | ---: | ---: | ---: |
| ！120 代发工资 | 160 | **0.15** | 0.03 | 4.03 | 9.88 |
| ！50 票据业务（116 整列） | 1011 | **0.35** | 30.90 | 4.56 | 10.19 |
| ！70 债券业务（12 处外链） | 1210 | **0.48** | 0.11 | 4.49 | 4.28 |
| ！100 单位存款 | 1331 | **0.45** | 2.51 | 4.66 | 6.03 |
| 省信贷（大簿） | 73807 | 269.53 | — | **10.63** | — |

**结论**：1-4 真实负载（≤1331 格）上 betteroffice 比 COM **快 9~25 倍**；
大簿（73,807 格）反而慢 25 倍——因为逐公式解释执行，且需 50M 预算 + 2 轮迭代。

### 5.4 strip_style_only_cells 内容完整性（！100 模板）

7 个工作表，有值格与公式数**逐表完全一致**（剥离 15,087 个纯样式空格）：

| 工作表 | 有值格 | 公式数 |
| --- | ---: | ---: |
| 单位存款余额 | 880 | 551 |
| 单位存款发生额 | 795 | 351 |
| 存量单位存款分产品期限 | 584 | 420 |
| 单位存款发生额分产品期限 | 185 | 0 |
| 极端数据核查信息 | 59 | 9 |
| 集中系统数据 | 1818 | 0 |
| 参照表 | 612 | 0 |

剥离产物：所有 XML 部件解析通过、Excel COM 可正常打开（7 表）。

## 6. 修复的一个真 bug

`strip_style_only_cells` 首版在重建被改动的 `<row>` 时：
`open_tag` 本身已以 `>` 结尾，却又追加一个 `>`，且 `row_tail_start` 停在
`</row>` **之前**导致原始闭合标签被重放——产物变成 `<row ...>>...</row></row>`，
Excel 与 openpyxl 都读不了（实测 `expected 'sheetData' tag, not 'row'`）。

已修复并加回归测试 `xml_worker_does_not_duplicate_row_close_tags`：
断言无 `>>`、每行一个 `</row>`、行数不重复、整段可通过 XML 解析。
ooxml-patch 测试 **44/44 passed**。

## 7. 未验证项

- **Linux 未实机验证**：仅确认 vendor 无 C 依赖、`--offline` 构建通过；deb 交叉编译未跑。
  路由代码里 Linux 分支复用既有 LibreOffice 回退，逻辑与 Windows 对称但未实测。
- 循环引用（`result.cycle_cells` 非空）时拒绝并回退，但**未构造样本验证**该分支。
- `#NAME?` 检测（未实现函数）逻辑存在但未构造样本触发验证。
- 省信贷 1,145 处剩余差异（`IFERROR` 除零语义子链）未逐格定性。
- 未与 Excel COM 做"同输入实时重算"的逐格对拍（对拍基准是 COM 重算后的输出文件）。
- `PYTOOLS_BO_CALC_ROUNDS` 上限 5 的行为未实测（只用过 1/2/3）。

## 8. 建议的后续（非本次范围）

1. Linux 实机验证（deb 构建 + 1-4 端到端）。
2. 若上游 betteroffice 修复 `used_range` 语义或实现 `INDIRECT` 静态依赖分析，
   可摘除 vendor patch 并去掉多轮迭代。
3. 构造 `#NAME?` 与循环引用样本，补齐这两条拒绝分支的测试。
