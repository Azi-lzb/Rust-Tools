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

## 5.10 4-4 转换的 LO 路由：11 文件大规模对拍（2026-11-09）

4-4 此前在 Linux 上直接拒绝（"LibreOffice conversion is paused"）。现按平台路由：

| 平台 | .xls/.xlsm → .xlsx | Word / WPS 原生 / .xlsb |
| --- | --- | --- |
| Windows | Excel/WPS COM（pywin32） | COM |
| **Linux** | **LibreOffice（本次放开）** | 仍要求 Windows COM |

`select_conversion_engine`（auto）在非 Windows + `xlsx` 目标 + `xls`/`xlsm` 源时
返回 `lo`；显式 `--engine lo` 也不再被拒（仅限电子表格 xlsx 目标）。

### 对拍语料（11 个 .xls）

5 个真实业务文件 + 1 个模板（含 108 公式）+ wei40（787 公式）
+ 5 个由 xlsx 模板经 COM 另存的 .xls（后 5 个用于压测，见下方注意）。

### 结果

| 文件 | 公式 | 公式文本精确差 | 值差 | 纯精度表示 | 共享公式 COM/LO |
| --- | ---: | ---: | ---: | ---: | --- |
| wei40.xls | 787 | **0** | **0** | — | 19 / 0 |
| 统计工具-报表工具 v1.0.4.9 | 108 | **0** | **0** | — | 7 / 0 |
| 个人贷款202607.xls | 0 | — | 0 | — | — |
| 单位贷款202607.xls | 0 | — | 0 | — | — |
| 报表工具-校验结果（个人/单位） | 0 | — | 0 | — | — |
| ！120 代发工资 | 160 | **0** | 48 | **48** | 10 / 0 |
| ！90 同业业务 | 671 | 18 | 86 | **86** | 12 / 0 |
| ！100 单位存款 | 1331 | 18 | 218 | **215** | 32 / 0 |
| ！10 单位贷款 | 1072 | 13 | 1044 | 1 | 38 / 0 |
| ！50 票据业务 | 1011 | 15 | 2798 | 3 | 38 / 0 |

### 差异分类（全部定性，无引擎缺陷）

1. **公式文本差异（13~18 处）**只在整列引用上：COM 写 `$1:$1048576`，
   LO 写 `$1:$65536`（LO 的 65536 行限制）。语义等价——目标表实际只有几十行数据，
   且 Excel 打开重算后即归一到 1048576。
2. **值差异的绝大多数是纯精度表示**（！120：48/48；！90：86/86；
   ！100：215/218）。COM/LO 另存时的浮点文本位数不同，数值本身相同。
3. **！50 的 2795 处与！10 的 1043 处**：公式逐字相同（已核对 S5/T5），
   差异在**缓存值**——`=NA()` 类公式 LO 写出 `#N/A`、COM 写空；
   `INDIRECT` 链 LO 重算得 `#REF!`、COM 保留 0。属转换期是否重算的差异，
   Excel 打开后按 `fullCalcOnLoad` 重算即一致。
4. **结构差异**：LO 把 shared formula 全部展开（COM 保留 19~38 个），
   语义相同、读取更简单。
5. **布尔结果缓存表示**：COM `True`/`FALSE`，LO `1`/`0`（公式语义一致）。

### 结论

LO 另存的**公式语义与数值**与 COM 一致，差异集中在缓存表示、整列范围写法、
shared formula 展开——都不影响 Excel 打开与后续重算。已放开 Linux 路由。

### 两个必须记录的注意点

- **语料陷阱**：xlsx→xls（COM SaveAs FileFormat=56）产生的样本中有 5 个
  COM 自己再读会失败（"读取源工作簿公式和值"阶段报 com_error）。这不是 LO
  或路由的问题，而是这类二次转换样本本身不完整。真实业务 .xls 5/5 全部成功。
- **LO 路由限定**：仅 `xls`/`xlsm` → `xlsx`；`.xlsb`、`.et`/`.ett`、Word 目标
  在 Linux 仍明确报错要求 Windows COM（未实测 LO 对它们的保真度）。

## 5.11 4-4 转换失败的两个真实原因（2026-11-09 追加实测）

§5.10 记录了 5 个 .xls 转换失败。深入排查后分成两类，**都不是 LO 路由的问题**：

### 5.11.1 整列范围写法差异（已修复）

失败信息：`工作表 '单位存款余额' 的公式在 M55 不一致`，
源 `.xls` 写 `集中系统数据!$1:$65536`、产物 `.xlsx` 写 `$1:$1048576`。

BIFF 单表上限 65536 行、OOXML 上限 1048576 行；Excel 的 SaveAs 会把整列引用
规范到目标格式的上限。语义完全相同。

修复（两处，口径一致）：
- Rust `excel_reader::canonicalize_formula_for_recalculation` 新增
  `canonicalize_sheet_row_ceilings`：把 `$1:$65536` / `$D$1:$D$65536` 两种
  写法的行上限视为等价；`=A1+65536` 之类字面量不受影响（有反向测试）。
- pywin32 worker 的 `_compare` 原本用 `a.strip() == b.strip()` 比对公式，
  现改为 `_canonicalize_formula`（同样处理行上限、`TRUE()`、`=` 前缀与空白）。
  注意：worker 是**随包**的，改源码后必须同步 `dist/python/` 与
  `F:/RustBuild/pytools/{debug,release}/python/`，否则跑的是旧副本
  （本轮实测踩过：改了源码但 exe 读的是 exe 同级 `python/` 下的旧 worker）。

修复后该批 5 个文件从 0/5 转为 **4/5 成功**。

### 5.11.2 源文件自身缺陷（不可修复，非管线问题）

剩余 1 个失败：`统计工具-报表工具(基础数据用)v1.0.4.9 (1).xls`
（24 个工作表、2.3 MB、`HasVBProject=True`）。

- 失败阶段：`重新打开候选工作簿` —— Excel 打不开自己刚 SaveAs 出的 XLSX。
- **手工验证**：脱离本工具，直接用 Excel COM `SaveAs(FileFormat=51)`，
  产物同样打不开（`类 Workbooks 的 Open 方法无效`）。
- 结论：该源文件（带 VBA 工程的老工具簿）经 Excel 自身 SaveAs 就产出坏文件，
  与本工具或 LO 路由无关。这类文件应改用其他方式（如先在 Excel 里手工另存
  并确认可打开）处理。

### 5.11.4 .xlsb / .et 实测与路由放开

**结论：.xlsb 已放开给 Linux LO；.et/.ett 无法用真实样本验证，维持现状。**

#### .xlsb（已实测，已放开）

造样本：2 个 S1F1 模板经 COM `SaveAs(FileFormat=50)` 生成。

| 文件 | COM | LO | 表名一致 | 公式 | 归一后公式差 | 纯精度值差 | 真值差 |
| --- | --- | --- | --- | --- | ---: | ---: | ---: |
| ！100 单位存款 | ok | ok | 是 | 1331 / 1341 | **10**（`=NA()` 写法） | 215 | 3 |
| ！120 代发工资 | ok | ok | 是 | 160 / 160 | **0** | 48 | 0 |

公式差异只有 `=NA()` vs 空（LO 写出 `=#N/A`），值差全是浮点表示位数。
与 .xls 的结论完全一致 → `select_conversion_engine` 已加入 `xlsb`。

顺带修了一个既有缺陷：worker 的容器校验对 `.xlsb` 要求 `xl/workbook.xml`，
但二进制工作簿的主部件是 `xl/workbook.bin`，导致 .xlsb 一律被误判为"包被裁剪"。
已按后缀区分必需部件。修复前 .xlsb 全部失败（0/2），修复后 2/2 转换成功。

#### .et / .ett（未能验证，未放开）

用 WPS 2019 专业版（`ket.Application`）把一个 xlsx 存为 `.et`，
但 `Workbook.FileFormat` 读出 **56（xlExcel8 = 标准 BIFF XLS）**，
不是 WPS 原生 ET（65521）。因此：

- 现有 `.et` 原生格式校验**正确地**拒绝了它（"未将该文件识别为原生 ET/ETT"）；
- 手上没有 WPS 真正原生保存的 `.et` 样本，**无法完成 LO 对 .et 的保真度实测**；
- 因此 `.et`/`.ett` 的 Linux LO 路由**未放开**（`select_conversion_engine`
  里虽已列出，但 `ensure_wps_source_format_verified` 仍会在原生校验处拒绝非原生
  内容；真正原生 .et 在 Linux 上的行为**未验证**）。

要把 .et 加入路由，需要一份 WPS 原生保存的 .et 真实样本，再重复 §5.10 的对拍流程。

### 5.11.5 真实 .et 样本对拍（2026-11-09，用户提供）

用户提供了 2 个真实 `.et`（`2026.09 5k01 广州银行`、`5k05 广东南粤银行`
仲恺区存贷统计月报，各 2 表、32KB/47KB）。特性：

- 容器：OLE2/CFB
- `Workbook.FileFormat` 经 WPS 2019（`ket.Application`）与 Excel 均读出 **56**
  （xlExcel8，标准 BIFF），**不是** WPS 原生 ET（65521）

即：这些 `.et` 是"扩展名为 .et 的 BIFF 工作簿"。因此它们在 Windows 上会被
现有 `ensure_wps_source_format_verified` 拒绝（"未识别为原生 ET"），
但在 Linux 上可以直接按 BIFF 处理。

#### LO vs COM 对拍结果

| 文件 | COM | LO | 表名 | 公式 COM/LO | 归一后公式差 | 纯精度值差 | 业务值差 |
| --- | --- | --- | --- | --- | ---: | ---: | ---: |
| 5k01 广州银行 | ok | ok | 一致（2 表） | 47 / 47 | **0** | 4 | **0** |
| 5k05 广东南粤银行 | ok | ok | 一致（2 表） | 67 / 67 | **0** | 14 | **0** |

公式精确差 3 / 1 处，全部是 `TRUE()`、表名引号之类，归一后归零。

#### 唯一的实质差异：日期格的数字格式

两各文件各有 2 处（`1人民币!A2`、`2外币!A2`）：

| | COM | LO |
| --- | --- | --- |
| 原始值 | `46295` | `2026-09-30 00:00:00` |
| 等效 Excel 序列号 | **46295** | **46295**（相同） |
| 数字格式 | `General` | `[$-409]m/d/yyyy` |

**业务值完全相同**（同一序列号），差异只在"显示为数字 vs 显示为日期"。
这是 LO 在 SaveAs 时按内容猜测日期格式导致的显示差异，不改变数据。

#### 结论与路由决定

- LO 对这类"BIFF 内容的 .et"转换**保真度与 COM 一致**（公式与会话值零差异，
  唯一差异是日期格的显示格式）。
- 但**仍未把 et/ett 加入 Linux LO 路由**：现有 Windows 侧校验要求"原生 WPS ET"，
  而这批样本恰恰不是原生的；若在 Linux 放开 et，会出现"同样文件在 Windows 被拒、
  在 Linux 通过"的不一致语义。需要先明确产品口径：
  1. 是否允许".et 扩展名的 BIFF 内容"按 BIFF 处理（那 Windows 侧也要放开），或
  2. 维持"只接受原生 ET"（那 Linux 也维持拒绝）。
- 若用户确认口径为 1，改动量很小：去掉 `ensure_wps_source_format_verified`
  的强校验 + 在 `select_conversion_engine` 加回 et/ett。

### 5.11.6 .et 口径落地（方案 A，2026-11-09）

用户又提供 3 个新建 `.et`（工作簿1、工作簿21、测试测是）＋此前 2 个，共 5 个，
`Workbook.FileFormat` 全部是 **56（xlExcel8）**。结论：**WPS 2019 专业版通过
COM 另存 `.et` 时不写 WPS 原生格式（65521），一律写 BIFF 内容。**

因此按"内容决定"落地（方案 A）：

| 源内容 | 处理 |
| --- | --- |
| `.et`/`.ett` + 原生 WPS（65521/65522） | 保持原校验（`source_format_verified=true`），Windows COM |
| `.et`/`.ett` + BIFF（56/8/16） | **放行**，按 `.xls` 同样处理；worker 记 `source_content_kind="biff"` |
| 其它（如 OOXML 改名成 .et） | 仍拒绝 |

改动四处：
1. worker 的 ET 校验：`FileFormat < 65521` 时不再直接拒绝，仅当不在
   `BIFF_FILE_FORMATS={56,8,16}` 时才报错；校验证据新增 `source_content_kind`。
2. worker 的 `source_format_verified` 改为"仅原生为 true"（BIFF 时 false + kind=biff）。
3. Rust `ensure_wps_source_format_verified`：原生或 BIFF（且 kind=biff）都通过。
4. Rust `validate_conversion` 的 et/ett 分支：同样接受 BIFF 证据。
5. Linux `select_conversion_engine` 加回 `et`/`ett`。

#### 验收

| 检查 | 结果 |
| --- | --- |
| 5 个 .et（工作簿1/工作簿21/测试测是/5k01/5k05）CLI 转换 | **5/5 converted** |
| 产物 Excel COM 可打开 | **5/5**（2/2/1/2/2 表，公式 3/3/11/47/67 全保留） |
| CLI 产物 vs COM 直接转换（Excel COM 读值逐格） | 公式差 **0**、值差 **0** |
| 此前"B5=#REF!"差异复核 | **是 openpyxl `data_only` 误判**；Excel COM 读值两侧均为 `'fdaf '`，零差异 |

### 5.11.3 顺带修的一处健壮性缺陷

`_snapshot()` 原先一次读取 `UsedRange.Formula`，在 154x40 的「基本配置」表上
抛 com_error（同表 `Value2` 可正常读）。现加回退：区域批量读失败时改用
`SpecialCells(xlCellTypeFormulas)` 按 Area 逐块读公式文本。

## 5.12 用 calamine 复核：修正前几轮对拍里的假阳性（2026-11-09）

### 方法错误与更正

前几轮 §5.10~§5.11 的跨引擎对拍**用 openpyxl 读转换产物**，这是错的：

- openpyxl 只支持 OOXML，本身不解析 BIFF；
- 它的 `data_only=True` 直接读 `<v>` 缓存，对错误格 / 日期格 / 布尔格的解释与
  Excel 不一致，导致**把读取器差异误报成引擎差异**。

改用项目自带的 calamine（`excel_reader::read_workbook`）复核后，多个"差异"消失：

| 文件 | openpyxl 口径 | **calamine 口径** |
| --- | --- | --- |
| wei40（787 公式） | 0 / 0 | 0 / 0 |
| 统计工具模板（108） | 0 / 0 | 0 / 0 |
| ！120 代发工资 | 0 / **48** | **0 / 0** |
| ！90 同业业务 | 18 / **86** | 0 / **0** |
| ！100 单位存款 | 18 / **218** | 0 / **3** |
| ！50 票据业务 | 15 / **2798** | 0 / **754** |

**全部公式差归零。** 此前 5 个 `.et` 报告的 `B5=#REF!`、日期格差异同样是
openpyxl 误判——calamine 与 Excel COM 读出来都是 `'fdaf '`，零差异。

### 比较器还需要两条等价规则

1. **空串格 vs 缺格**：COM 会写 `<c><v></v></c>`（空串），LO 干脆不写该格。
   两者都是"空"，不算差异。
2. **`=NA()` vs 无公式**：LO 给 `#N/A` 缓存格补一个 `NA()` 公式，COM 留空。
   错误格场景下等价。

### 剩下的真差异：只有 LO 的 INDIRECT 求值

| 文件 | 剩余值差 | 根因 |
| --- | ---: | --- |
| ！120 / ！90 / wei40 / 统计工具 / 5 个 .et | **0** | — |
| ！100 单位存款 | 3 | 文本拼接的浮点第 15 位（`…748192` vs `…748193`） |
| ！50 票据业务 | **754** | **全部**源于 `INDIRECT` 求值差异 |

！50 的 754 处经逐条归因：536 处公式本身含 `INDIRECT`，其余 218 处是
`SUM(INDIRECT(...))` 的**下游级联**（LO 把 INDIRECT 算成 `#REF!`，Excel 的
`CalculateFullRebuild` 算成 `0`）。**公式文本逐字相同**，差异只在缓存值。

即：LO 转换 .xls 时对 `INDIRECT` 动态引用求值不可靠（会得 `#REF!`）。
影响面评估：只有"公式用了 INDIRECT 且依赖其缓存值"的文件会受影响；
本次 11 个样本里 1 个（！50）。而 1-4 的校验区域重算会重新计算，不受影响；
4-4 纯转换场景下该缓存差异会写进产物，但 Excel 打开后 `fullCalcOnLoad`
重算即恢复正确值。

### 后续对拍的正确姿势（写进规范）

| 读取对象 | 用哪个读取器 |
| --- | --- |
| `.xlsx` 产物 | **calamine**（`excel_reader::read_workbook`），或 Excel COM |
| `.xls` / `.et` / `.xlsb` 源 | 必须先转换；转换后同样用 calamine 读 |
| 判"缓存值对不对" | **不要用 openpyxl `data_only`** |

calamine 的优势：纯 Rust、快、无需 COM/LO、`.xlsx`/`.xls`/`.xlsb` 全支持。

## 5.13 求值交给 Rust：LO 只做格式转换（2026-11-09）

### 决策

按"主要用 Rust 求值"的原则，把 LO 的职责限定为**格式转换**，公式求值交给
betteroffice：

1. LO 转换 `.xls`/`.et`/`.xlsb` → `.xlsx`（格式保真，包含样式与结构）；
2. 转换产物立即交给 betteroffice 重算 2 轮，把 `<v>` 缓存刷成 Excel 口径；
3. 重算失败则**回滚**到 LO 原始产物（LO 的公式文本正确，不能留半新半旧缓存）。

落点：`conversion_engine::run_fidelity_conversion_with_runtime` 的 `lo` 分支，
新增 `refresh_converted_caches_with_betteroffice`。

**设计意图（2026-11-09 用户确认保留）**：LO 转换后的 betteroffice 重算**不是
多余步骤**——它让转换产物的 `<v>` 缓存回到 Excel 口径，这样下游用 calamine
`data_only` 读缓存值（1-1/2-1/4-4 等功能都这么做）时拿到的就是 Excel 算过的
结果，而不是 LO 的 `#REF!`/异形缓存。实测 ！50：LO 转换后 754 格缓存错误
（`#REF!`），betteroffice 重算后与 COM 直接转换**只差 1 格**（浮点文本末位）。

原因：LO 对 `INDIRECT` 的求值不可靠。实测 ！50 模板有 754 个格在 LO 转换后
得到 `#REF!`（原 `.xls` 缓存是 `0.0`），其中 536 处公式直接含 `INDIRECT`、
218 处是下游级联。公式文本逐字正确，只是缓存错。

### 连带修掉的第二个差异：15 位有效数字

用 betteroffice 重算后，！100/！90 反而各出现 20/21 处差异——全部是
**数字转文本的位数**：Excel 的 `&` 运算符用 15 位有效数字，betteroffice 原来用
Rust 的最短往返表示（最多 17 位）。

实测 8 个样本全部符合"按 15 位有效数字格式化即与 Excel 一致"：

| Excel 文本 | betteroffice 原值 | 15 位格式化 |
| --- | --- | --- |
| `901777.42` | `901777.4199999999` | `901777.42` ✓ |
| `1111.11` | `1111.1100000000001` | `1111.11` ✓ |
| `-1219.99999999997` | `-1219.999999999971` | `-1219.99999999997` ✓ |
| `74.5061355864011` | `74.50613558640113` | `74.5061355864011` ✓ |
| `-645.360000000001` | `-645.3600000000006` | `-645.360000000001` ✓ |

已改 vendor `betteroffice-xlsx-calc/src/eval.rs::format_number`（第 4 处 vendor
修改，见 `vendor/VENDOR_NOTES.md`）。

### 验收（calamine 读取，6 个真实 .xls）

| 文件 | 公式差 | 值差 |
| --- | ---: | ---: |
| wei40（787 公式） | 0 | **0** |
| 统计工具模板（108） | 0 | **0** |
| ！120 代发工资 | 0 | **0** |
| ！100 单位存款 | 0 | **0** |
| ！50 票据业务 | 0 | **0** |
| ！90 同业业务 | 0 | **0** |

**LO 转换 + betteroffice 重算与 COM 直接转换完全一致（零差异）。**

1-4 回归：！100 / ！50 / ！90 命中 20 / 11 / 70 条，与基准一致。

### 已验证（2026-11-09 补充）：betteroffice 在 Linux 可编译、可运行、求值正确

三项此前"未验证"中的一项已完成：

1. **Linux 编译**：`cargo zigbuild --target x86_64-unknown-linux-gnu.2.28`
   （zig 来自 pip `ziglang` 0.16）在 Windows 侧交叉编译通过，产物 5.3 MB
   ELF x86-64（glibc 2.28 目标）。
2. **Linux 运行**：WSL Ubuntu（x86_64，glibc 2.39）直接执行成功——
   `INDIRECT` 专项测试簿 15 格全部重算，`exhausted=false`，calc 1.7 ms。
3. **求值正确**：全部 15 格（静态跨表/字面 INDIRECT/拼接构造/单元格间接/
   整列 INDEX/中文表名/SUM 聚合）与 Excel 基准逐格一致。

结论：**betteroffice 在 Linux 可用**。1-4 Linux 的求值链路可以从 LO 切换到
betteroffice（与 Windows 同源同行为），LO 保留为格式转换与兜底。

### 附带澄清（2026-11-09 复核 INDIRECT 构造）

`xlsx → LO calculateAll → INDIRECT` 构造专项测试（4 链路同源对拍）：

| 构造 | COM | LO | betteroffice(Win) | betteroffice(Linux/WSL) |
| --- | --- | --- | --- | --- |
| 静态跨表 `Data!A2` | ✅ | ✅ | ✅ | ✅ |
| 字面 `INDIRECT("Data!A2")` | ✅ | ✅ | ✅ | ✅ |
| 拼接 `INDIRECT("Data"&"!A2")` | ✅ | ✅ | ✅ | ✅ |
| 单元格间接 `INDIRECT($E$1&"!A2")` | ✅ | ✅ | ✅ | ✅ |
| 整列 `INDEX(Data!$1:$1048576,…)` | ✅ | ✅ | ✅ | ✅ |
| 中文表名 `INDIRECT("参照表!B2")` | ✅ | ✅ | ✅ | ✅ |
| `SUM` 聚合（直引/间接） | ✅ | ✅ | ✅ | ✅ |

**LO 在 xlsx（OOXML 导入）路径下对这些 INDIRECT 构造求值全部正确**——此前的
`#REF!` 缓存问题**只发生在 BIFF 导入路径**（.xls/.et 源）。含义：
- 1-4 Linux（工作副本恒为 xlsx）用 LO 计算是安全的；
- 4-4/4-7 从 BIFF 源转换时才需要 betteroffice 刷缓存（当前实现已如此）。

### 未验证项的处理决定（2026-11-09 用户确认）

- ~~betteroffice Linux 编译运行~~ → 已验证（上节）。
- **原生 WPS ET（65521）样本**：不再追。COM 另存无法产生原生 ET，样本不可得；
  BIFF 内容的 .et（实际遇到的全部形态）已按内容放行并实测通过。**不用管。**
- **betteroffice 重算失败的回滚分支**：不再构造样本。betteroffice 能正确处理
  绝大多数文件；失败时回退链已有代码兜底。**不用管。**
- Windows 上未装 LibreOffice，LO 链路的端到端（转换+重算）只在组件层面验证
  （WSL 转换 + Windows 重算），未在单一进程内跑通 `--engine lo`。维持记录。

## 5.14 转换校验的数值容差：1e-15 → 1e-14（2026-11-09）

### 问题

`.xlsx` 转换（COM/LO 两条路由）在保存后会把源与产物逐格比对，数值用相对
容差 `1e-15`。用户现场反馈报 `conversion failed: pywin32 Excel 转换失败`，
差异样本：

```
0.0012083234879418526  vs  0.00120832348794185
```

这是纯粹的**浮点文本尾差**（相对差 `2.15e-15`），但超过了 `1e-15`。

### 根因

Excel 以 **15 位有效数字**保存/显示数值；BIFF→OOXML 重存时两侧可能分别落在
17 位与 15 位形式。实测该截断造成的相对差最大约 **2.7e-15**：

| 原值 | 15 位形式 | 相对差 |
| --- | --- | --- |
| `1234.5678901234567` | `1234.56789012346` | 2.76e-15 |
| `0.12345678901234566` | `0.123456789012346` | 2.70e-15 |
| `12345678.901234567` | `12345678.9012346` | 2.72e-15 |
| `1.2345678901234568e-05` | `1.23456789012346e-05` | 2.61e-15 |
| `901777.4199999999` | `901777.42` | 1.29e-16 |

`1e-15` 容差因此过紧，正常文件会被判成"转换改变了非公式数据"而整次失败。

### 修复

统一为 `VALUE_TOLERANCE = 1e-14`（比最坏实测大 3.7 倍，仍比 Excel 15 位精度
的固有误差紧一个数量级），并补两条保护：

1. **整数必须精确相等**：相对容差在 `1e15` 量级会放过 ±1 的改动，故两侧
   都是整数值时要求逐位相同。
2. **零附近用绝对容差** `VALUE_ZERO_TOLERANCE = 1e-12`：原实现要求两侧都
   非零，`0` 与 `1e-300` 会被误判。

落点（三处，口径一致）：
- `pywin32_excel_conversion_worker.py::_values_equal`（COM 路径）
- `conversion_engine::cell_values_equal`（Rust 路径）
- `conversion::numeric_values_equal`（旧路径，改为复用共享常量）

上报的 `numeric_tolerance` 字段同步为 `1e-14`。

### 验证

`conversion_value_tolerance_absorbs_float_tails_only` 单测覆盖：6 组 15 位尾差
（含用户现场样本）必须相等且对称；3 组零/极小值必须相等；6 组真实改动必须
拒绝；整数 ±1 与大数必须拒绝；数字与文本不等。

11 个真实 `.xls` 转换：**10 converted / 1 failed**，唯一失败仍是
「统计工具模板」（带 VBA，Excel 自身 SaveAs 产物也打不开，与本容差无关）。

全量回归 571 passed / 0 failed / 15 ignored。

## 5.15 非公式数据比对改为软失败（2026-11-09 用户决定）

### 决定

转换校验里**非公式数据**的差异不再阻断转换，只作提示；**公式文本**仍是硬失败。

| 比对项 | 行为 |
| --- | --- |
| 工作表数量/顺序、公式增删、公式文本 | **硬失败**（真实结构改动，终止转换） |
| 非公式数据的单元格增删、值不一致 | **软失败**：转换照常发布，差异写入校验报告 |

理由：Excel/WPS 重存本来就可能改写浮点尾数、空串表示与日期格式；把这类差异
当硬失败会让正常文件无法转换（§5.14 的 `2.7e-15` 尾差就是实例）。而公式文本
变化代表真实的结构改动，必须拦。

### 实现

- `pywin32_excel_conversion_worker.py::_compare`：值差异收集到
  `value_differences`，不再 `raise ValueError`；返回报告带上
  `value_difference_count` / `value_differences`（最多 50 条）/
  `value_differences_truncated`。
- `conversion_engine::compare_workbooks`：签名改为接收 `&mut
  ConversionValueNotes`，值差异只记录；`validate_conversion` 把提示并入上报
  JSON，并标注 `value_comparison = "soft (reported, not blocking)"`。
- `MAX_VALUE_DIFFERENCES = 50`：避免报告被单文件差异淹没。

### 验收

| 检查 | 结果 |
| --- | --- |
| 值差 1%（真实改动） | ✅ 转换成功，报告 `value_difference_count=1` 并附明细 |
| 值差 `2.15e-15`（尾差） | ✅ 不阻断，且不产生提示（在 `1e-14` 容差内） |
| 公式文本改动 | ✅ 仍硬失败（`formula text ... 不一致`） |
| 新增单测 `non_formula_value_differences_are_reported_but_not_blocking` | ✅ 值差返回 Ok + 提示；公式差异仍 Err；完全一致无提示 |
| 11 个真实 `.xls` 转换 | 10 converted / 1 failed（失败仍是带 VBA 的统计工具模板，与比对无关） |
| 全量回归 | 572 passed / 0 failed / 15 ignored |

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
