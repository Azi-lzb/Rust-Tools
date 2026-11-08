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

### 4.1 config.xlsx 全局配置（业务配置入口，优先）

| 键 | 默认值 | C 列备注（写入表内，用户可见） |
| --- | --- | --- |
| `校验计算引擎` | betteroffice | 1-4 校验区域重算用哪个引擎：betteroffice=纯Rust优先（快，推荐）；com=直接用Excel/WPS；off=同com |
| `校验计算预算上限` | （空=默认1000万） | 重算时全簿共享的公式求值次数上限；留空=默认1000万；算不完会自动按回退策略处理 |
| `校验计算轮数` | 2 | 重算跑几轮：1~5；第2轮起跨表动态引用（INDIRECT）才收敛；保持2即可 |
| `校验计算剥离样式空格` | 否 | 是=重算前删掉工作副本里只有格式没有内容的空格子（大表更快）；否=不删 |
| `校验计算回退策略` | auto | betteroffice算不动时怎么办：auto=Windows回Excel/WPS、Linux回LibreOffice；com=只回Excel/WPS；libreoffice=只回LibreOffice；none=不回退直接报错 |

解析在 `config.rs::parse_global_config`；键清单在 `config_admin.rs::GLOBAL_DEFAULTS`
（三元组：键/默认值/备注，init 与 repair 都会把备注写进 C 列）。

### 4.2 环境变量（应急覆盖，config 未配置时生效）

| 环境变量 | 默认 | 说明 |
| --- | --- | --- |
| `PYTOOLS_BO_CALC` | `auto`（启用） | `off` 完全回到 COM/LibreOffice 旧路由 |
| `PYTOOLS_BO_CALC_MAX_VISITS` | `0`=库默认 1000 万 | 重算预算上限 |
| `PYTOOLS_BO_CALC_ROUNDS` | `2` | 重算轮数，上限 5 |
| `PYTOOLS_BO_CALC_STRIP_STYLE` | `off` | `on` 时**一次性**剥离纯样式空格（不做事后重试） |
| `PYTOOLS_BO_CALC_FALLBACK` | `auto` | `auto`/`com`/`libreoffice`/`none`；`none` 时超限明确报错 |

**优先级**：config.xlsx 全局配置的键值**优先**；该键缺失或为空时才读对应环境变量；
两者都没有时用内置默认。落点在 `bo_calc.rs::BoCalcConfig::load(config_path)`，
config 路径由 `config.rs::default_config_workbook_path()` 探测（与交互菜单同口径，
支持 `PYTOOLS_CONFIG_PATH` 显式指定）。

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

## 5.5 大簿性能优化：content_range 记忆化（2026-11-08 补充）

首版实测省信贷大簿 1 轮 112.5 秒，与 S1F1 小表（毫秒级）相差 4 个数量级。
归因发现真正的瓶颈**不是计算**：

- 各公式形态单独测速（1200 条/形态）：静态 SUMIFS 22.9ms、INDIRECT SUMIFS 28.6ms、
  SUMPRODUCT 17.3ms、RANK 8.0ms、裸引用 3.9ms——公式形态本身没有慢路径。
- 但省信贷 12,162 个公式 × 771,721 个存储格 ≈ **94 亿次** extent 扫描。
  `content_range()` 每次被调用都全表遍历 `cells`，而整列引用每算一格就问一次 extent。
  （原版 `used_range()` 同样每次全扫，只是首版预算先耗尽、掩盖了这个开销。）

修复：`Sheet::content_range()` 加 `OnceLock` 记忆化（沿用 `CellStyles` 同一范式），
并在 `set_cell` 做**精确失效**——只有"有无内容"发生转换时才清缓存，重算改写已有
公式格的值不会击穿缓存；`remap_cells` 无条件失效。

效果（省信贷，同预算同轮数）：

| 场景 | 优化前 | 优化后 | 提速 |
| --- | ---: | ---: | ---: |
| 1 轮 @10M 预算（预算耗尽路径） | 75.4 s | **2.62 s** | 29× |
| 1 轮 @50M 预算（全部算完） | 112.5 s | **3.49 s** | 32× |
| 2 轮 @50M + 落盘（对拍用） | 223.5 s | **6.97 s** | 32× |

正确性零变化：省信贷 2 轮对拍仍为 72,482/73,807（98.2048%，与优化前逐位相同）；
S1F1 四模板仍为 99.70% / 99.90% / 99.26% / 100%。vendor 测试 model 101、calc 257+集成全绿，
新增 `content_range_cache_invalidates_only_on_content_change` 回归测试。

## 5.6 四引擎最终耗时（含大簿，含 LibreOffice）

| 样本 | 公式格 | betteroffice | IronCalc | Excel COM | LibreOffice |
| --- | ---: | ---: | ---: | ---: | ---: |
| ！120 代发工资 | 160 | **0.15 s** | 0.03 s | 4.03 s | 9.88 s |
| ！50 票据业务 | 1011 | **0.35 s** | 30.90 s | 4.56 s | 10.19 s |
| ！70 债券业务 | 1210 | **0.48 s** | 0.11 s | 4.49 s | 4.28 s |
| ！100 单位存款 | 1331 | **0.45 s** | 2.65 s | 8.06 s | 15.52 s |
| 省信贷（大簿） | 73807 | 17.88 s | — | **10.71 s** | 19.84 s |

小表（1-4 真实负载）betteroffice 快 9~34 倍；大簿与 COM/LO 同量级（慢 1.7~1.8 倍），
不再是 25 倍的差距。大簿仍建议走回退（COM 最快），但 betteroffice 已可用于无 Office
环境的兜底。

## 5.7 Rust-Audit 全部 20 个 S1F1 模板对拍（2026-11-08 补充）

用真实 CLI（`summary-named-validation`）把 Rust-Audit 的 20 个 business_templates
逐个双路由跑一遍，比对**最终校验结果表**（命中条数、单元格、错误类型、描述文本）：

| 结果 | 数量 | 说明 |
| --- | ---: | --- |
| 完全对齐（命中数相同 + 实质差异 0） | **18 / 20** | — |
| 计算前拦截（与引擎无关） | 2 / 20 | COM 路由报出**一字不差**的同样错误 |

### 18 个对齐模板

| 模板 | 命中 | 实质差异 | 仅浮点显示 | betteroffice | COM |
| --- | ---: | ---: | ---: | ---: | ---: |
| ！10 单位贷款 | 78 | **0** | 0 | 13.77 s | 9.70 s |
| ！100 单位存款 | 20 | **0** | 34 | **7.93 s** | 10.62 s |
| ！110 存量个人存款 | 1 | **0** | 0 | **3.14 s** | 8.92 s |
| ！120 代发工资 | 3 | **0** | 0 | **2.13 s** | 8.29 s |
| ！130 互联网金融（季度） | 1 | **0** | 0 | **0.39 s** | 8.48 s |
| ！130 互联网金融（月度） | 6 | **0** | 0 | **0.29 s** | 8.27 s |
| ！140 养老产业 | 0 | **0** | 0 | **0.82 s** | 8.59 s |
| ！140 数字经济 | 0 | **0** | 0 | **0.53 s** | 8.43 s |
| ！140 普惠贷款 | 0 | **0** | 0 | **0.88 s** | 8.61 s |
| ！140 科技贷款 | 0 | **0** | 0 | **1.28 s** | 8.82 s |
| ！140 绿色贷款 | 1 | **0** | 0 | **0.61 s** | 8.81 s |
| ！30 存量专项一批 | 2 | **0** | 0 | **0.24 s** | 8.26 s |
| ！40 委托贷款 | 2 | **0** | 0 | **2.61 s** | 8.82 s |
| ！50 票据业务 | 11 | **0** | 1 | **5.74 s** | 10.15 s |
| ！60 单位贷款置换 | 0 | **0** | 0 | **0.42 s** | 8.32 s |
| ！61 存量专项二批 | 0 | **0** | 0 | **0.22 s** | 8.32 s |
| ！80 股权和 SPV | 2 | **0** | 0 | **5.87 s** | 9.36 s |
| ！90 同业业务 | 70 | **0** | 30 | 9.87 s | 9.60 s |

**所有 18 个模板的实质差异均为 0。** 浮点显示差异只出现在 3 个模板
（34 / 1 / 30 处），即 §5.4 说明的 15 位 vs 17 位文本拼接规则，不影响判错结论。

### 2 个计算前拦截的模板

| 模板 | 拦截原因 | COM 路由表现 |
| --- | --- | --- |
| ！20 个人贷款 | 公式引用工作表 `REF`，模板中无此表且无法补装 | **同样失败**（错误文本一致） |
| ！70 债券业务 | 模板含外部工作簿引用（1 个 externalLink 部件），复制后成悬空引用 | **同样失败**（4-7 fail-closed 口径） |

两者都在**计算开始之前**的模板校验阶段被拦下，与计算引擎无关。！70 正是 4-7
定义要 fail-closed 拦截的场景（模板含 `[N]` 外链）。

### 口径说明：为什么"逐格 98%"不等于"结果不一致"

`CALC_ENGINE_BENCHMARK.md` 报的 98.20%~99.90% 是**逐格比中间值**的口径，
把"同一个数的第 15 位小数不同"也算作差异。按**业务结论**口径（最终校验结果表的
判错结论、错误类型、定位单元格、描述文本）重新统计：

- ！100：命中 20/20 相同，排序后单元格集合完全相同，**实质差异 0**（34 处仅浮点显示）
- 省信贷大表：1,327 处差异中，**含判错关键词的结论格差异为 0**

## 5.8 组级批量编排：并行 betteroffice + 回退项统一处理（2026-11-09）

### 问题

首版是逐文件流：每个副本独立 `bo_calc::recalculate` → 拒绝时当场回退平台引擎。
多模板场景下平台引擎被逐个启动，浪费在进程启动与重复加载上。

### 实现

三阶段编排（`pytools-core/src/bo_calc_batch.rs` + `summary_ops`）：

1. **准备阶段**（循环内，逐文件）：每个源创建校验副本、复制模板校验公式，
   只登记 `PendingValidationCalculation`，不计算。
2. **批量计算**（循环后一次）：`bo_calc_batch::recalculate_parallel` 用 rayon
   对所有副本并行跑 betteroffice（每本独立 `Workbook`，无共享可变状态）；
   成功项直接采用，`BoCalcRefusal` 项进待处理集合。
3. **回退阶段**：只有拒绝项交给平台引擎（COM/Linux 用既有 worker，会话复用，
   不重复启动）；回退引擎整簿重算，不依赖 betteroffice 残留缓存。

`PYTOOLS_BO_CALC=off` 时跳过批量阶段，全部副本走平台引擎（旧行为）。
fail-closed 语义不变：`FALLBACK=none` 时拒绝项直接报错；`application` 字段格式
与单文件路径逐字一致。

### 验收

| 场景 | 结果 |
| --- | --- |
| 3 副本全成功 | betteroffice 成功 3 / 拒绝 0，**平台引擎零启动**；耗时 5.47 s |
| 4 副本全成功（并行） | 5.88 s；串行基线 ~21.2 s ⇒ **并行提速 3.6×** |
| 2 副本全拒（MAX_VISITS=10） | 成功 0 / 拒绝 2；`application` = `Excel.Application；betteroffice 回退（重算预算耗尽（1329 个公式格未求值，保留旧值不可信））`，命中仍 20 条 |
| 4 副本全拒 | 全部回退平台引擎，4 份结果各 20 条命中 |
| 预算阈值 | 1300（<1331 格）全拒；1400 仍拒（含预算外余量），行为与单文件一致 |
| 单文件路径 | 与批量前一致：！100 命中 20 条 / betteroffice；！140 命中 1 条 |
| 全量回归 | **559 passed / 0 failed / 15 ignored**（+2 为批量编排单测） |

单元测试 `bo_calc_batch::tests`：
- `all_copies_succeed_and_order_is_preserved`：4 副本全成功、索引完整、顺序保持
- `a_broken_copy_is_refused_without_failing_the_group`：2 成 1 拒，坏副本被拒绝
  且不中断整组（这是端到端无法构造的"部分成/部分拒"混合场景的唯一覆盖）

### 未验证

- 未构造出**端到端**的"同组部分成功、部分拒绝"（同一模板的副本行为一致，
  预算阈值只能全成或全拒）；该路径由上述单元测试覆盖。
- Linux 上批量并行 + LibreOffice 会话复用未实测。

## 5.9 .xls 另存为 .xlsx 后修复：COM 与 LibreOffice 对比（2026-11-09）

### 4-7 扩展

4-7 原本对非 ZIP 容器直接记 skipped。现在 `.xls` 可走"另存为 .xlsx → 修复"
路由（`run_excel_open_repair_batch_with_legacy_conversion`）：

- Windows：Excel/WPS COM（与 4-4 同一条 pywin32 路由）
- Linux：LibreOffice 无头转换（`soffice --convert-to xlsx`；`PYTOOLS_SOFFICE`
  可指定 soffice 路径）
- 另存产物直接落到 `-Excel修复.xlsx` 候选名；另存后无修复点时保留副本
  （状态 unchanged，转换本身即修复价值）；转换失败记 skipped 带原因
- `.et` 等其它格式仍 skipped；默认不传引擎时 .xls 也 skipped（带指引）

### LO 另存可靠性实测（决定路由的依据）

| 样本 | 公式数 | 公式文本 COM vs LO | 值差异 | Excel 可开（两产物） |
| --- | ---: | --- | ---: | --- |
| 单位贷款202607.xls（真实，纯数据） | 0 | —（两侧均无公式） | 0 | ✅ / ✅ |
| wei40.xls（！40 模板往返，787 公式） | 787 | **零差异**（787:787） | **0** | ✅ / ✅ |
| 合成共享公式样本（IF 布尔结果） | 21 | 零差异（21:21） | 20（见下） | ✅ / ✅ |

结构差异（不影响打开与计算）：

- COM 另存保留 shared formula 结构（wei40：552 个 shared）；
  **LO 另存把 shared formula 全部展开为逐格公式**（0 个 shared）——语义相同、
  文件略大，对后续 calamine 读取与公式比对反而更简单。
- **布尔结果的缓存表示不同**：`IF(...,TRUE(),FALSE())` COM 缓存 `True`、
  LO 缓存 `1`。公式语义一致（Excel 打开后重算即为布尔），但 calamine
  `data_only` 回读会读到 1 而非 TRUE。1-4 的校验公式返回文本（"错|…"），
  不受影响；业务模板若有裸布尔结果需注意此差异。

结论：LO 另存与 COM 另存在公式文本与值上**一致**（布尔缓存表示除外），
已加入路由（Linux 侧）。Windows 侧仍优先 COM。

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
- 省信贷剩余差异未逐格定性（2 轮时 1,325 处，3 轮时 1,145 处，集中在 `IFERROR` 除零语义子链）。
- 未与 Excel COM 做"同输入实时重算"的逐格对拍（对拍基准是 COM 重算后的输出文件）。
- `PYTOOLS_BO_CALC_ROUNDS` 上限 5 的行为未实测（只用过 1/2/3）。

## 8. 建议的后续（非本次范围）

1. Linux 实机验证（deb 构建 + 1-4 端到端）。
2. 若上游 betteroffice 修复 `used_range` 语义或实现 `INDIRECT` 静态依赖分析，
   可摘除 vendor patch 并去掉多轮迭代。
3. 构造 `#NAME?` 与循环引用样本，补齐这两条拒绝分支的测试。
