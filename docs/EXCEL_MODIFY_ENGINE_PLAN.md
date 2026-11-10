# Excel XLSX 修改引擎技术评估与 Benchmark —— 测试计划（里程碑 1 交付物）

状态：**待人工审阅**。本文件只是计划；在获得确认前不会实施任何 Benchmark。
任务来源：用户 2026-11-10 指令（BetterOffice / rXLS / openpyxl 三引擎修改能力、保真度、性能对比）。
核心原则：**保真度优先，性能其次，功能覆盖第三，维护成本与部署复杂度纳入决策。**

---

## 1. 仓库现状调查（今天做了什么）

### 1.1 当前引擎分工（已用源码核实）

| 引擎 | 版本 | 当前承担 | 证据 |
| --- | --- | --- | --- |
| **openpyxl**（Python 子进程 worker，协议 v3） | 3.1.5（随包 `dist/python`） | 绝大多数就地修改：2-3 追加候选、配置修复/初始化、4-5 `.xlsx` 改名、5-x 系列行级操作、1-1/1-3 目标写入回退、图表目标回退；操作面覆盖 M01–M14 | `crates/excel-editor/python/openpyxl_worker.py`（约 3200 行）；`crates/excel-editor/src/openpyxl_backend.rs` |
| **umya-spreadsheet** | `=3.1.1`（workspace 固定） | 2-3 追加的首选候选引擎（`.xlsx` 专用，含宏 `.xlsm` 明确拒绝并回退 openpyxl）；3-1 的历史 umya 打印引擎（已从菜单撤下，仅留代码）；测试夹具构造 | `crates/excel-editor/src/umya_backend.rs`、`pytools-core/src/inplace_edit.rs:69`、`Cargo.toml:37` |
| **手写 OOXML（窄范围补丁）** | — | 4-7 打开修复（`excel_open_repair`）、共享公式 ref 归一、样式空格剥离（`strip_style_only_cells`）、外部页眉替换（`ooxml_header_backend`）、读取补充（批注/命名区域/日期格式元数据，`excel-reader`）、1-4 缓存回写 | `crates/ooxml-patch/src/*`、`crates/excel-reader/src/lib.rs:2030+`、`pytools-core/src/excel_open_repair_batch.rs` |
| **betteroffice**（计算，非修改） | 0.5.1（registry + vendor 补丁） | 仅 1-4 公式重算（`bo_calc.rs`：`Workbook::open` → 计算 → `.save()`）。**当前不承担任何修改业务** | `pytools-core/src/bo_calc.rs:233-254`、`vendor/VENDOR_NOTES.md` |
| **calamine** | 0.37.0 | 只读：所有功能的源读取 | `crates/excel-reader` |
| **rust_xlsxwriter** | workspace | 只写新建：3-1 rust writer、测试夹具 | `pytools-core/src/static_print_engines.rs` |

### 1.2 真实业务修改操作 → 测试项映射

| 测试项 | 现有真实调用 |
| --- | --- |
| M01/M02/M03 单格/批量改值 | 2-3 追加（462 行实测）、配置写入 |
| M04 公式文本 | worker `set_formula_text`；5-1/5-7 行级公式守卫 |
| M05 数字格式 / M06 字体填充边框 | worker 样式操作；3-1 打印输出 |
| M07 行高列宽 | 3-1 打印输出、2-3 追加保留列宽 |
| M08 合并区域 | 1-1/1-3 输出、3-1 保留合并 |
| M09 批注 | 2-1/2-2 按批注去重（读+改）、3-1 批注丢弃/保留 |
| M10 命名区域 | 1-3/1-4 模板定义区域（读）；4-4 重算区域复制（写） |
| M11 打印设置 | 3-1 页面设置（方向/居中/FitToPages） |
| M12 工作表增删重命名 | 4-5 改名、1-6 提取新建 |
| M13 复杂模板少量修改 | 配置修复、5-2 机构规范化 |
| M14 混合修改 | 2-3 追加计划（值+批注+格式+合并多操作合一） |

### 1.3 需要严格保真的业务点（决定保真测试权重）

- **未修改部件字节不变**：2-3 追加、4-7 修复（图表、VBA、未知部件不能丢）。
- **公式文本与缓存值**：5-1/5-7 守卫、1-1/1-3 输出（缓存清空 vs 保留的口径已定型）。
- **合并/样式/列宽**：3-1 打印、1-1 宽表输出。
- **命名区域**：1-3/1-4 模板路由依赖定义名作用域（`localSheetId`）。
- **外部链接**：4-7 修复场景（悬空外链）。

---

## 2. 三个引擎的核实结果（真实源码 / 官方文档，非宣传）

### 2.1 openpyxl 3.1.5（基线）

- 修改 API 完整：值/公式/样式/合并/批注/命名区域/打印/表操作全部原生支持（现有 worker 已在用）。
- 已知保真特性（本项目实测沉淀）：保存时**丢弃全部公式缓存值**（需外部恢复）、重排 XML 部件、`xml:space` 需补写、未知扩展会丢弃并告警。
- 部署：需随包 Python 运行时（约 25MB）+ 子进程桥接（协议 v3，握手/超时/有界输出）。

### 2.2 betteroffice-xlsx 0.5.1（registry + vendor）

核实来源：`Cargo.lock:138-152`（betteroffice-xlsx + ops/parse/render/drawingml/opc 子 crate）、vendored `betteroffice-xlsx-model/src/workbook.rs`（本次通读）、`bo_calc.rs` 实际调用、上游 `github.com/openooxml/betteroffice`。

- 加载/保存：`Workbook::open(&bytes)` → 修改 model → `.save()`。**保存是全量重渲染**（render crate 重新序列化整个包），不是部件保留式。
- 修改 API（vendored model 源码核实）：`Sheet::cell_mut/set_cell/set_calculated_value/adopt_cells`、`merges: Vec<CellRange>`（有类型）、`DefinedName`、列样式 runs、数组公式、`numfmt`/`styles.rs`/`rich_text.rs`/`chart.rs` 类型齐备。
- **待 M2 探针核实**（render 覆盖面决定成败，文档未逐项承诺）：批注、数据验证、条件格式、打印设置、图片、外部链接、列宽行高、命名区域作用域（`localSheetId`）在保存后的往返保真。
- 已知历史证据：S1F1 20 模板计算往返 9822 个公式文本零改动；但那是「计算+保存」，修改+保存的保真没有测过。
- 部署：纯 Rust，已在 workspace 内（vendor 补丁），无运行时依赖。

### 2.3 rxls 0.1.3/0.1.4（= 用户所称 "rXLS"，2026-09-28 发布）

核实来源：docs.rs/rxls（0.1.4，2026-09-28，MIT，100% documented）、github.com/HyunjoJung/rxls README、crates.io 依赖清单。

- 修改 API（README/文档核实）：`Spreadsheet::open(&bytes)`（保留原始包字节用于**无损保存**）→ `set_cell_value("Data", 0, 0, Cell::Text(..))` → `save()`；`Spreadsheet::transaction` 原子批编辑；每次变更前检查 `EditCapability`，不可编辑时返回 `EditReadOnlyReason`。
- 保存语义：**未触碰的已声明部件逐字节保留（含 VBA）**——与 openpyxl（全量重写）、betteroffice（全量重渲染）根本不同。
- 官方声明覆盖：cell/formula/range 编辑、表生命周期、布局/冻结/打印区域、合并、批注（legacy notes）、超链接、精确范围数据验证、表格底部扩行。
- **官方明确不支持**：插入/删除行列（README non-goals）→ M 系列中的行操作类不可用，但 M01–M14 里没有行列插入删除，不影响本轮。
- 类型化损失报告：`StyleLoss`/`PrintLoss`/`ChartUnsupportedReason`/`FormulaUnsupportedReason`——保存时会明确告知哪类样式无法精确表示，这对保真评估非常有利。
- 成熟度：0.1.x 早期（12 stars、24 open PRs、0.1.4 尚未正式发布且 README 自述"preparing"）；但测试量可观：1092 项 all-target 测试、916 文件公开语料 0 意外失败、与 openpyxl 可见值对拍 387/387。
- 依赖：cfb、encoding_rs、thiserror + 可选 chrono/quick-xml/serde/zip；MSRV 1.85（M2 需确认与 glibc 2.28 交叉编译链兼容）。

### 2.4 能力矩阵（M2 探针将逐项验证；本表为计划基线）

| 操作 | openpyxl | betteroffice | rxls |
| --- | --- | --- | --- |
| M01–M03 改值（单/100/1000） | ✅ 现役 | ✅ API 在（`set_cell`） | ✅ API 在（`set_cell_value`） |
| M04 公式文本 | ✅ | ✅ 待探针 | ✅ 待探针 |
| M05 数字格式 | ✅ | ✅ 类型在，渲染待探针 | ✅ 待探针 |
| M06 字体/填充/边框/对齐 | ✅ | ⚠️ 类型在，往返待探针 | ⚠️ `CellStyle` 在，损失类型待探针 |
| M07 行高列宽 | ✅ | ⚠️ 待探针 | ⚠️ 未列入 README，待探针 |
| M08 合并增删 | ✅ | ✅ 类型在，待探针 | ✅ README 明确 |
| M09 批注 | ✅ | ⚠️ 待探针 | ✅ legacy notes |
| M10 命名区域 | ✅ | ✅ 类型在，作用域待探针 | ⚠️ 未见 API，待探针 |
| M11 打印设置 | ✅ | ⚠️ 待探针 | ✅ README 明确（打印区域）|
| M12 表增删重命名 | ✅ | ⚠️ 待探针 | ✅ sheet lifecycle |
| M13 复杂模板小改 | ✅ | ⚠️ 全量重渲染风险 | ✅ 部件保留式（卖点） |
| M14 混合修改 | ✅ | ⚠️ 待探针 | ✅ transaction |
| 行/列插入删除 | ✅ | ⚠️ 待探针 | ❌ 官方 non-goal |
| XLSM（宏） | ✅ 保留宏 | ❓ 待探针 | ✅ 含 VBA 字节保留 |

图例：✅ 已核实支持；⚠️ 类型/API 存在但往返保真未核实（M2 探针项）；❌ 官方明确不支持；❓ 完全未知。

---

## 3. 独立测试环境设计（里程碑 2，确认后实施）

- 位置：`benchmark/excel-modify-engine/`（新目录，加入 workspace `exclude`，不影响现有构建）。
- 结构：`Cargo.toml`（独立 crate）+ `src/`（驱动 + 计时）+ `fixtures/`（生成脚本产物）+ `results/`（CSV/JSON）+ `probe/`（每个引擎的能力探针小程序）。
- **不修改任何业务代码、不替换引擎、不动现有测试**；betteroffice 通过 path 依赖指向现有 `vendor/`（与生产同版本同补丁）。
- openpyxl 走两种模式测：A) 独立 Python 进程；B) 复用现有 worker 桥接（`openpyxl_backend` 同款子进程协议）——两者分开报告。
- 版本固定并写入报告：openpyxl 3.1.5 / betteroffice-xlsx 0.5.1(+vendor) / rxls 0.1.3（若 0.1.4 已发布则记录两者）/ umya 3.1.1（仅作参照）/ Rust 与 Python 版本 / OS。

## 4. 测试文件设计

**修改能力夹具**（M01–M14 共用，合成生成脚本入仓可复现）：

普通值表、公式表（含跨表引用+缓存值）、多表（5 表）、复杂格式（字体/填充/边框/数字格式/行高列宽）、合并区域、批注、命名区域（全局+局部作用域）、数据验证、条件格式、打印设置、图片+图表、外部链接（真实悬空外链形态）、XLSM（独立标记）。

**优先复用真实模板脱敏副本**：`测试用例/2/目标用例.xlsx`（含跨表镜像公式，M13 的现成样本）；若用户提供更复杂业务模板则以脱敏副本为准。

**性能规模样本**（合成）：1k / 10k / 50k / 100k 行 × 20 列 + 一个复杂模板（≥10 表、含公式/合并/样式/批注）。

## 5. 修改能力测试（M01–M14）

按任务表格执行；每项记录：原生 API 有无、成功/失败、是否需绕过 API 直改 OOXML、结构异常、额外依赖、限制与原始错误信息。引擎不支持 → 标 `Unsupported`，**不虚构耗时、不计为性能失败**。

## 6. 保真度验证（三层）

**A. 结构层**：ZIP 完整性、部件清单 diff、Relationships/Content Types 校验、XML 可解析、未修改部件 SHA-256 对照。区分「字节变化」与「语义变化」（合法重序列化不算 FAIL，记录为 RESERIALIZED）。

**B. 语义层**：以 calamine + `excel-reader`（现有中立读取器）对拍：值/类型、公式文本、缓存值、合并、数字格式、批注、命名区域（含作用域）、数据验证、条件格式、打印设置、图片/图表/外链。允许修改区域变化，其余出现非预期变化即 FAIL。

**C. 兼容层**：Windows 上用 Excel COM（pywin32，环境已有）+ WPS 手动双开抽查：修复提示、内容丢失、公式错误、再保存再打开。无法自动化时标 `Not Tested`，不冒充通过。

结果四态：`PASS / FAIL / UNSUPPORTED / NOT_TESTED`，绝不合并。

## 7. 性能测量方法

- 负载：1 / 100 / 1000 格、跨表分散、公式+格式、混合；完整链路 **打开→解析→修改→保存→关闭** 分段计时。
- 批量：10/50/100 簿 × 模式 A（每簿冷启动进程）/ 模式 B（常驻进程连续处理）。openpyxl 桥接模式单独一列，**不与 Rust 常驻直接对比下结论**。
- 规范：Release 编译；同文件同修改；串行避免竞争；预热后 ≥5 次取中位数/P95/min/max；记录峰值 RSS、CPU 时间；轮换引擎测试顺序抑制缓存偏差；记录硬件/OS/版本/构建配置。
- 试运行（如 100k 规模成本过高）允许，但报告中单列，不冒充完整结果。

## 8. 风险与预案

| 风险 | 影响 | 预案 |
| --- | --- | --- |
| betteroffice 全量重渲染丢部件（图表/图片/未知部件） | 保真 FAIL | 逐部件 diff 定位；结论里明确「只适合值/公式编辑场景」 |
| rxls 0.1.x 太早期（API 变动、MSRV 1.85 与交叉编译链冲突） | 选型风险 | 锁定 0.1.3 测试；记录 MSRV 冲突；结论给「观察期」选项 |
| rxls `EditCapability` 拒绝某些修改（如外部链接工作簿） | Unsupported | 如实记录拒绝原因（`EditReadOnlyReason` 有类型化输出，恰好便于归档） |
| 100k 行样本 openpyxl 内存爆炸 | 测试不可完成 | 先试运行；openpyxl 只保证跑到其能力上限，如实标注 |
| Excel COM 自动化不稳定 | 兼容层缺证 | COM 失败改 WPS 手动抽查 + 标 Not Tested |
| vendor betteroffice 与上游 0.5.1 行为分叉 | 结论失真 | 报告同时注明「vendored 补丁版」测试结果与上游差异点（VENDOR_NOTES 已列 4 处） |

## 9. 实施顺序（里程碑 2 内部）

1. **探针周**：三个引擎各写最小 probe（M01–M14 逐项跑通/跑不通），产出能力矩阵终稿 → `EXCEL_MODIFY_ENGINE_CAPABILITY.md`。
2. **夹具生成**：合成脚本 + 真实模板脱敏副本入库（`benchmark/excel-modify-engine/fixtures/`）。
3. **保真测试**：先能力允许的引擎 × 12 类夹具，三层验证 → `EXCEL_MODIFY_ENGINE_FIDELITY.md`。
4. **性能测试**：单文件 6 负载 × 5 规模 → 批量 10/50/100 × 2 模式 → `EXCEL_MODIFY_ENGINE_BENCHMARK.md` + CSV/JSON。
5. **汇总建议** → `EXCEL_MODIFY_ENGINE_RECOMMENDATION.md`（回答任务第七节 12 问，给 A/B/C/D 结论，允许分操作结论）。

## 10. 验收对照

- [ ] 计划经人工确认（本文件，当前状态）
- [ ] 能力矩阵全部基于真实执行（非 README 推断）
- [ ] 保真报告含 OOXML diff 与 Office 兼容证据
- [ ] 性能报告含阶段耗时、批量、冷/常驻、中位数/P95/min/max、内存/CPU
- [ ] 机器可读原始结果入库（CSV/JSON）
- [ ] 业务代码零改动、现有测试结果零变化
- [ ] 结论回答第七节 12 问并给明确建议

---

## 附：本轮核实来源清单

- 仓库：`crates/excel-editor/src/{openpyxl_backend,umya_backend,plan}.rs`、`crates/excel-editor/python/openpyxl_worker.py`、`crates/ooxml-patch/src/*`、`vendor/betteroffice-xlsx-model/src/workbook.rs`、`vendor/VENDOR_NOTES.md`、`pytools-core/src/bo_calc.rs`、`Cargo.toml`/`Cargo.lock`、`pytools-core/MIGRATION_INVENTORY.md`
- rxls：docs.rs/rxls（0.1.4）、github.com/HyunjoJung/rxls（README：修改 API、non-goals、保真合同、916 文件语料、成熟度）
- betteroffice：github.com/openooxml/betteroffice（上游）、vendored 源码（0.5.1 + 本仓库 4 处补丁，见 `vendor/VENDOR_NOTES.md`）
- 既有可复用方法论：`docs/CALC_ENGINE_BENCHMARK.md`（1-4 计算引擎对比的计时/对拍框架）
