# 现有 Excel 引擎业务审计

## 审计基线

- 仓库：`C:\Users\AZI\Desktop\VibeCode\Rust-Tools-feature`
- 分支与提交：`feature/business-development`，`09a314d085de50174d6fb0e6c46fc49c5f24b384`。该提交是用户授权将 `main` 快进合并到 Feature 后的共同基线。
- 方法：核对当前 Rust 菜单分发、计划执行入口、引擎适配器及具体业务调用点。历史清单只用于寻找线索；与当前代码不一致时，以当前代码为准。
- 下表记录路由层实施前、未设置路由覆盖时的生产调用基线；文末另列本轮新增路由，避免把原有默认调用与新灰度路径混为一谈。
- 下表中 `NOT_FOUND` 表示当前生产菜单代码中没有找到 Openpyxl Worker、UMYA 或 BetterOffice 的调用；它不表示该业务菜单不存在。其他实际调用的引擎仍单独列出。

## 菜单盘点

| 菜单编号及功能 | 当前实际引擎调用 | 实际 Excel 修改 | 公式计算 / 图表 |
| --- | --- | --- | --- |
| 1-1 时序宽表汇总 | 新建输出使用 Calamine + Rust writer；写入既有目标使用 Openpyxl Worker。 | 新建汇总工作簿；既有目标计划创建工作表或导入单元格区域，可按计划复制值、公式文本和样式。 | 不重算公式；目标导入可能写入公式文本或其缓存值。没有图表数据修改调用。 |
| 1-2 按使用区域快速汇总 | `NOT_FOUND`（目标引擎）；当前读取/输出路径使用 Calamine 和 Rust writer。 | 读取并生成汇总输出，不修改源工作簿。 | 没有公式重算或图表修改调用。 |
| 1-3 按模板命名区域汇总 | `NOT_FOUND`（目标引擎）；当前为 Rust 读取与汇总输出路径。 | 读取模板命名区域、比较结构并生成汇总结果，不修改源工作簿。 | 不重算公式；没有图表修改调用。 |
| 1-4 按批注汇总（模板+源文件） | `NOT_FOUND`（Openpyxl Worker、UMYA、BetterOffice）；当前实现为 Calamine 读值 + `rust_xlsxwriter` 写新结果。 | 读取模板批注定义和源数据，生成汇总工作簿；不修改源文件。 | 不执行公式重算；没有图表修改调用。 |
| 1-5 按模板校验区域提取校验（`PT-6-3`） | BetterOffice 在隔离副本上尝试公式重算，拒绝或不适用时转 Windows Excel/WPS COM 或 Linux LibreOffice Calc；Calamine 回读结果。复制校验区域使用定点 OOXML，存在 Openpyxl 复制回退。 | 在临时副本中复制校验区域、重算并读取结果；源工作簿不修改。 | **有公式重算**。没有图表数据修改调用。 |
| 1-6 表头路径比对与映射 | `NOT_FOUND`（目标引擎）；当前为 Rust/Calamine 读取和比对。 | 读取和比较配置路径，不修改工作簿。 | 没有公式重算或图表修改调用。 |
| 1-7 时序配置预校验 | `NOT_FOUND`（目标引擎）；当前为 Rust 配置校验及 Calamine 读取。 | 校验规则和配置，不修改工作簿。 | 没有公式重算或图表修改调用。 |
| 2-1 按配置检查重复并标记 | `.xlsx` 使用 Openpyxl Worker 候选路径；`.xls` 使用专用 Windows Excel/WPS COM 路径；Rust/Calamine 负责读取和规划。 | 将重复行标红，并按安全候选/快照流程提交。 | 不重算公式；没有图表修改调用。 |
| 2-2 按配置删除重复 | `.xlsx` 使用 Openpyxl Worker 候选路径；`.xls` 使用专用 Windows Excel/WPS COM 路径；Rust/Calamine 负责读取和规划。 | 删除重复行，保留首次出现；`.xlsx` 路径对含公式工作表保留写前守卫。 | 不运行公式计算引擎；没有图表修改调用。 |
| 2-3 按配置去重追加到目标工作簿 | 既有 `.xlsx` 目标使用 Openpyxl Worker；既有 `.xls` 目标走专用 COM 追加；新目标使用 Rust writer。 | 追加去重后的行/工作表，并按计划写值或公式文本。 | 不重算公式；既有公式文本可能被保留，依赖缓存按既有计划处理。没有图表数据修改调用。 |
| 2-4 去重/追加配置预校验 | `NOT_FOUND`（目标引擎）；当前为 Rust 配置校验和 Calamine 读取。 | 只校验配置，不修改工作簿。 | 没有公式重算或图表修改调用。 |
| 3-1 按打印配置生成（子菜单选项 1） | Openpyxl 专用打印脚本。 | 新建静态打印工作簿，复制配置区域的值/缓存值、格式、可选批注、合并、行列尺寸和打印设置。 | 使用公式缓存值，不重算公式；当前脚本没有复制或修改图表的调用。 |
| 3-1-Rust 按打印配置生成（子菜单选项 2） | Calamine + OOXML 辅助读取 + `rust_xlsxwriter`。 | 新建 `_Rust` 后缀静态打印工作簿；写数据、格式/合并和页面属性。 | 使用缓存值，不重算公式；没有图表修改调用。 |
| 历史 UMYA 打印 helper（当前 3-1 菜单调用 `NOT_FOUND`） | `NOT_FOUND`。`run_umya_engine` 函数仍存在，但没有生产调用者。 | 无当前菜单输出；历史函数保留在源码中。 | 当前菜单无公式计算或图表处理调用。 |
| 3-3 打印配置预校验（当前可见子菜单选项 3，内部 `PT-3-4`） | `NOT_FOUND`（目标引擎）；读取并校验打印配置。 | 不修改工作簿。 | 没有公式重算或图表修改调用。 |
| 4-1 文件名机构标准化 | `NOT_FOUND`（目标引擎）；文件系统重命名。 | 不修改 Excel 内容。 | 不涉及公式计算或图表。 |
| 4-2 文件排序前缀管理（添加/清理） | `NOT_FOUND`（目标引擎）；Rust 文件系统操作。 | 不修改 Excel 内容。 | 不涉及公式计算或图表。 |
| 4-3 工作簿内容命名前缀（添加/清理） | `NOT_FOUND`（目标引擎）；读取工作簿内容后对文件名执行 Rust 文件系统重命名。 | 不修改工作簿内容。 | 不重算公式、不修改图表。 |
| 4-4 Office 格式转换 | 转换使用平台专用引擎；特定转换分支在暂存 `.xlsx` 上调用 BetterOffice 刷新公式缓存。 | 转换输出格式；可选地改写暂存输出中的公式缓存，不改源文件。 | 仅缓存刷新分支执行 BetterOffice 重算；没有图表数据修改调用。 |
| 4-5 按配置修改 Sheet 名 | `.xlsx` 候选使用 Openpyxl；Windows `.xls` 候选使用 Excel/WPS COM。 | 重命名工作表；Openpyxl 路径更新支持范围内的跨表公式与命名区域引用。 | 改写公式引用，不执行公式计算；没有图表数据修改调用。 |
| 4-6 机构提交完整性检查 | `NOT_FOUND`（目标引擎）；当前是读取/检查路径。 | 不修改工作簿。 | 没有公式重算或图表修改调用。 |
| 4-7 修复 Excel COM 无法打开的工作簿 | `NOT_FOUND`（Openpyxl Worker、UMYA、BetterOffice）；使用专用修复/转换路径，不是通用编辑器路由。 | 为修复生成候选副本；原件保持不变。 | 当前业务目标是修复可打开性；没有公式计算或图表数据修改调用。 |
| 5-1 拆分村镇银行数据（`PT-7-1`） | `.xlsx` 使用 Openpyxl；`.xls` 使用专用 Windows Excel/WPS COM。 | 插入/删除目标行并写入 Rust 计算出的拆分值。 | 业务计算在 Rust 中完成并写值；不调用公式重算。没有图表修改调用。 |
| 5-2 机构名称规范化（`PT-7-2`） | `.xlsx` 使用 Openpyxl；`.xls` 使用专用 Windows Excel/WPS COM。 | 按映射表改写机构名称单元格。 | 不重算公式；没有图表修改调用。 |
| 5-3 删除外资行并生成副本（`PT-7-3`） | 当前菜单仅接受 `.xlsx`，使用 Openpyxl；`.xls` 和其他格式会跳过。 | 删除匹配行并生成“金融局”副本。 | 不重算公式；没有图表修改调用。 |
| 5-4 修改外汇工作表页眉单位（`PT-7-4`） | `.xlsx` 先走定点 Rust OOXML patch，失败后使用受保护的 Openpyxl 回退；当前跳过 `.xls/.xlsm`。 | 替换页眉单位文本，验证候选后安全替换源文件。 | 不涉及公式计算或图表数据修改。 |
| 5-5 地区总分校验（`PT-7-5`） | `NOT_FOUND`（目标引擎）；读取并生成核对结果。 | 不修改源工作簿。 | 不执行公式重算；没有图表修改调用。 |
| 5-6 工作表数据提取（`PT-7-6`） | 使用 Openpyxl Worker。区域提取当前接受 `.xlsx/.xlsm`；完整工作表导入仅接受 `.xlsx`。 | 创建/导入工作表并复制配置区域；普通菜单关闭公式、批注复制，不请求完整样式/合并/尺寸导入。 | 不重算公式；没有图表数据修改调用。 |
| 5-7 涉农贷款比上月修正（`PT-7-7`） | `.xlsx` 使用 Openpyxl；`.xls` 使用专用 Windows Excel/WPS COM。 | Rust 计算差值并写入本期工作簿；`.xls` 走候选和快照保护。 | 写计算结果值，不执行公式重算；没有图表修改调用。 |
| 5-8 季度涉农贷款修正及分地区核对（`PT-7-8`） | 当前菜单仅接受 `.xlsx`，使用 Openpyxl 候选路径。 | 写入 Rust 计算出的季度调整值并产出地区核对结果。 | 不调用公式重算；没有图表修改调用。 |
| 6-1 按日期归档（预览） | `NOT_FOUND`（目标引擎）；Rust 文件/日期检查。 | 只生成预览。 | 不涉及公式计算或图表。 |
| 6-2 按类型归档（预览） | `NOT_FOUND`（目标引擎）；Rust 配置与文件检查。 | 只生成预览。 | 不涉及公式计算或图表。 |
| 6-3 执行归档计划 | `NOT_FOUND`（目标引擎）；Rust 文件系统复制/归档。 | 归档文件，不改 Excel 内容。 | 不涉及公式计算或图表。 |
| 6-4 按日期归档（直接复制） | `NOT_FOUND`（目标引擎）；Rust 文件系统复制。 | 复制归档文件，不改 Excel 内容。 | 不涉及公式计算或图表。 |
| 6-5 按类型归档（直接复制） | `NOT_FOUND`（目标引擎）；Rust 文件系统复制。 | 复制归档文件，不改 Excel 内容。 | 不涉及公式计算或图表。 |
| 7-1 初始化/修复配置 | 修改既有 `.xlsx` 配置候选时使用 Openpyxl Worker。 | 写配置单元格、创建缺失工作表，并按计划添加/清除批注；按现有候选校验和确认流程提交。 | 不执行公式重算；没有图表修改调用。 |
| 7-2 重新加载并查看配置状态 | `NOT_FOUND`（目标引擎）；Calamine 读取。 | 不修改工作簿。 | 不涉及公式计算或图表。 |
| 7-3 打开配置文件 | `NOT_FOUND`（目标引擎）；启动系统文件打开操作。 | 不修改工作簿。 | 不涉及公式计算或图表。 |
| 7-4 打开最近输出 | `NOT_FOUND`（目标引擎）；查找并打开文件。 | 不修改工作簿。 | 不涉及公式计算或图表。 |
| 7-5 打开输出目录 | `NOT_FOUND`（目标引擎）；启动系统文件管理器。 | 不修改工作簿。 | 不涉及公式计算或图表。 |

## 引擎调用点核对

### 菜单到源码函数的定位

| 菜单 | 当前调用链定位 |
| --- | --- |
| 1-1 | `RustTools.rs::run_wide_summary_cli_report` → `execute_wide_summary_target_writes` → `commit_wide_summary_existing_target_openpyxl`；Openpyxl 候选调用位于最后一个函数。 |
| 1-5 | `RustTools.rs` 菜单分发 → `summary_ops::extract_named_validation_regions_with_progress` → `ConversionRuntime::prepare_named_validation_copy` / `compute_prepared_copy` → `bo_calc::recalculate`；平台回退位于 `conversion_engine.rs`。 |
| 2-1 / 2-2 | `RustTools.rs::run_interactive_dedup_menu` → `dedup_ops::mark_duplicates_in_place_with_progress` / `delete_duplicates_in_place_with_progress`；`.xlsx` 分支调用 `inplace_edit::execute_candidate`，`.xls` 分支调用专用 COM 候选。 |
| 2-3 | `RustTools.rs::run_dedup_append_execute_report` → `run_dedup_append_direct_existing`；`.xlsx` Worker 路径由 `dedup_ops` 计划执行，`.xls` 有单独 COM 追加分支。 |
| 3-1 | `RustTools.rs::run_interactive_print_menu` → `run_configured_static_print_menu` → `static_print_engines::run_openpyxl_engine` 或 `run_rust_print_engine`。 |
| 4-4 | `RustTools.rs::run_convert_batch_menu` → `conversion_engine` 转换；`refresh_converted_caches_with_betteroffice` 是转换后特定 XLSX 缓存刷新函数。 |
| 4-5 | 菜单分发 → `sheet_rename` 批处理；`.xlsx` 调用 `ConversionRuntime::rename_sheets_in_candidate`（Openpyxl），Windows `.xls` 调用 `rename_sheets_com_in_candidate`。 |
| 5-1 | `RustTools.rs::run_village_split_batch` → `batch_ops::build_village_split_plan`；`.xlsx` 走 Openpyxl 计划，Windows `.xls` 走 COM。 |
| 5-2 | `RustTools.rs::run_institution_normalize_batch` → `institution_normalize::build_institution_normalize_plan` → Openpyxl 就地候选；Windows `.xls` 走 COM。 |
| 5-3 | `RustTools.rs::run_remove_foreign_bank_batch` → `batch_ops::build_remove_foreign_plan` → Openpyxl 删除行候选；菜单只接收 `.xlsx`。 |
| 5-4 | `RustTools.rs::run_fx_header_batch` → `execute_fx_header_in_place`；先调用 `ooxml_patch`，失败后通过 `execute_optional_plan` 走 Openpyxl。 |
| 5-6 | `RustTools.rs` 菜单分发 → `sheet_extract` 计划执行 → `inplace_edit::execute_candidate`（Openpyxl Worker）。 |
| 5-7 / 5-8 | `RustTools.rs::run_rural_loan_menu` → `batch_ops::build_rural_loan_plan` / `build_quarterly_rural_loan_plan`；`.xlsx` 菜单路径使用 Openpyxl，5-7 `.xls` 使用 COM。 |
| 7-1 | `RustTools.rs::run_interactive_config_initialize_or_repair` → Openpyxl 候选执行和配置回读校验。 |

### Openpyxl Worker 与专用 Openpyxl 业务脚本

- 可复用适配器：`crates/excel-editor/src/openpyxl_backend.rs::OpenpyxlXlsxBackend`；通过 `pytools-core/src/edit_plan.rs` 注册。正常就地计划由 `pytools-core/src/inplace_edit.rs` 生成候选、调用 ExcelModifier 校验并安全提交。
- 菜单调用链：1-1 既有目标、2-1/2-2 `.xlsx`、2-3 既有 `.xlsx` 目标、4-5 `.xlsx`、5-1/5-2/5-3/5-6/5-7/5-8 对应 `.xlsx` 路径、7-1 既有配置工作簿。
- 3-1 Openpyxl 打印在 `pytools-core/src/static_print_engines.rs::run_openpyxl_engine` 使用专用打印脚本，不是通用 Worker 计划协议。
- `execute_optional_plan` 的普通 `--in-place` 路径选 Openpyxl；兼容参数 `--in-place umya` 当前也映射到 Openpyxl。另有显式 `--capabilities` + `--platform` 计划执行入口，可按传入能力记录选择注册适配器。

### UMYA

- 可复用适配器：`crates/excel-editor/src/umya_backend.rs::UmyaXlsxBackend`，在能力文件计划执行中注册。
- 当前可见 3-1 菜单不调用 UMYA；`pytools-core/src/static_print_engines.rs::run_umya_engine` 仍在源码中，但 `rg` 未找到生产调用者（`NOT_FOUND`）。
- UMYA 适配器代码本身支持 `SetComment`，但本任务的新路由必须在选择 UMYA 时拦截批注修改。

### BetterOffice

- 生产计算调用：1-5 校验区域重算路径，在 `pytools-core/src/bo_calc.rs` 打开隔离副本并调用 `Workbook::recalculate_all`；由 `summary_ops.rs` / `conversion_engine.rs` 调用并在拒绝时交平台引擎。
- 另一处生产调用：4-4 转换路径中，特定 `.xlsx` 暂存输出会调用 `refresh_converted_caches_with_betteroffice` 刷新缓存。
- 当前 `pytools-core/src/edit_plan.rs` 没有 BetterOffice `EditBackend` 注册；也未发现菜单用 BetterOffice 执行行列结构修改（`NOT_FOUND`）。依赖库本身有编辑 API 不等于当前生产菜单已经调用它。

### Rust Native / OOXML

- 通用 RustNative `ExcelEditPlan` 适配器当前 `NOT_FOUND`。
- 5-4 使用的是 `ooxml-patch` 定点页眉文本操作；现有 `OoxmlHeaderTextBackend` 只注册 `ReplaceHeaderText`，不是通用单元格/打印/命名区域后端。
- 1-5 校验区域复制及 4-4 转换属性也有窄范围 OOXML 操作；3-1-Rust 是新建工作簿 writer，不是既有文件的 RustNative OOXML 修改器。

## 对后续路由实现有影响的现状

1. `ExcelEditPlan` 已存在，当前包含单元格值/公式/批注、行插删、行高列宽、工作表操作、页面设置等操作；当前 `PageSettings` 尚无页边距字段，也没有独立命名区域修改和列插删操作。
2. `crates/excel-editor/src/workbook_profile.rs` 将 `xl/calcchain.xml` 作为已知包部件接受，但尚未把 CalcChain 是否存在放入 `WorkbookFeature`；Auto 路由不能在判定时重新扫描文件，应在既有计划/profile 检查阶段提供该元数据。
3. 当前默认就地修改调用 Openpyxl，并使用候选校验、源文件快照与安全提交。新开关的默认 `OpenpyxlOnly` 必须保持该 Worker 路径和既有业务输出不变。
4. 当前计划执行的通用输出校验由 `PlanOutputValidator` 负责。新适配器必须接入现有校验与候选提交边界，不能因手动选择引擎跳过保护。

## 主要源码位置

- 菜单分发：`pytools-core/src/bin/RustTools.rs` 的 `run_interactive_data_menu`、`run_interactive_dedup_menu`、`run_interactive_print_menu`、`run_interactive_convert_menu`、`run_interactive_batch_menu`、`run_interactive_config_menu`。
- 统一计划与适配器注册：`pytools-core/src/edit_plan.rs`、`pytools-core/src/inplace_edit.rs`、`crates/excel-editor/src/plan.rs`。
- Openpyxl/UMYA 适配器：`crates/excel-editor/src/openpyxl_backend.rs`、`crates/excel-editor/src/umya_backend.rs`。
- UMYA/打印业务代码：`pytools-core/src/static_print_engines.rs`。
- BetterOffice 重算与转换：`pytools-core/src/bo_calc.rs`、`pytools-core/src/summary_ops.rs`、`pytools-core/src/conversion_engine.rs`。
- OOXML 能力盘点与 5-4 页眉 patch：`crates/excel-editor/src/workbook_profile.rs`、`crates/excel-editor/src/ooxml_header_backend.rs`、`pytools-core/src/bin/RustTools.rs`。

## 本轮路由实现后的更新（2026-11-12）

- 菜单 1-8 增加当前进程路由开关：`OpenpyxlOnly`、`UMYAOnly`、`BetterOfficeOnly`、`RustNativeOnly`、`Auto`。没有设置 `PYTOOLS_EXCEL_EDIT_ROUTE` 时，`inplace_edit::execute_candidate` 和 `execute_in_place` 保持原有引擎调用路径；显式设置后，已有的计划式业务入口进入 `excel_edit_router`。
- 通用 `execute_plan`、候选执行和安全就地执行共用现有 ExcelModifier 计划执行、输出验证、源文件快照及安全提交链。2-3 既有目标追加、4-5 XLSX 改名和 5-4 外汇页眉的专用计划路径也读取同一进程路由配置。
- 计划新增列插删、页边距和命名区域操作，以及 CalcChain 特征。`edit_plan.rs` 现在注册真实的 Openpyxl、UMYA、BetterOffice、RustNative OOXML 和内存 Composite 后端；新实现分别位于 `betteroffice_edit.rs`、`native_ooxml_backend.rs` 和 `excel_edit_router.rs`。
- 路由只读取 `ExcelEditPlan` 与已生成的工作簿 Profile。普通 OOXML `.rels` 部件不会单独触发复杂对象拦截；完整清单、外链、图表/图片/批注、宏、未知扩展、CalcChain 结构修改等仍按适配器能力拒绝或转 Openpyxl。UMYA 批注修改会被拦截。
- 真实样本验证：测试集 2-1 使用 UMYA 标记 10 行；2-2 使用 BetterOffice 删除 10 行。完整终端日志、原件哈希与副本回读结果见 `EXCEL_EDIT_ROUTE_VALIDATION.md`。所有修改均在测试集副本上完成。
- 专用转换、公式计算和新建工作簿路径仍遵循审计表所列的平台业务路由；本轮通用开关作用于既有 ExcelEditPlan 修改入口，不将 1-5 的平台重算或 3-1 的新建打印业务改成通用编辑器。
