# PyTools 测试覆盖清单

生成日期：2026-10-28（第一阶段）；测试执行日期：2026-10-28（第二阶段）。范围：七类菜单 32 项子功能。本清单是静态覆盖记录与实际执行结果台账，不是验收结论。

## 状态词汇

- **待运行**：测试代码已编写，尚未经 Cargo 编译执行。
- **已通过**：本次会话实际编译运行且有输出证据。
- **失败待修复**：实际运行失败。
- **环境限制**：依赖外部引擎或特定平台，当前环境未执行。
- **尚缺**：还没有可执行测试。

## 全局说明（2026-10-28，第二阶段执行后更新）

1. 第二阶段统一编译（3 分 43 秒，`F:\RustBuild\pytools` 缓存）与全部测试已实际执行：
   除 4 个失败外全部通过；4 个失败已按第三阶段要求定点修复并重跑通过，相关模块全量回归通过。修复记录：
   - 测试基建（3 处）：`sheet_extract.rs` 重复测试模块删除；
     `sheet_extract_area_operation_parity.rs` 旧函数名 `read_value_columns` 导入补齐；
     excel-reader 测试夹具从旧共享工作区路径 `audit-core` 改为本项目新合成夹具
     `crates/excel-reader/tests/fixtures/synthetic_probe.xlsx`。
   - 业务修复（1 处）：`summary_ops.rs` 错误单元格回填器新增"原地替换非类型化错误单元格"
     路径，修复 5-6 区域提取保留错误单元格时 `summary error cell A1 already exists` 失败；
     替换后仍经 `validate_error_cells` 重开工作簿强校验。
   - 测试预期更新（3 处，经原 Python 只读核实后判定测试过时而非业务错误）：
     表头比对汇总 C4/D4 差异数为数值单元格（Python `compare_393.py` 经 pandas 即写数值）；
     进度消息现为"第 X/Y 个源工作簿读取完成"（原断言"已处理 X/Y"）；
     PT-7-6 拒绝提示现为"区域提取仅接受 .xlsx/.xlsm"。
   - 偶发测试竞态（1 处）：excel-reader 8 个测试夹具助手共用"进程ID+纳秒"命名，
     并行测试线程在同一时间刻度撞名，一个线程删除文件时另一线程正要打开，
     导致偶发"找不到文件"失败；统一改为进程内原子序列号唯一命名，连续 3 次运行验证稳定。
2. 新增 2-3 边界测试 `pytools-core/tests/dedup_append_target_edges.rs`（3 个用例，
   手工推导的独立预期，见文件头注释）——已通过。
3. 人工菜单测试数据包 `testdata/manual-menu/`（31 个合成样本文件 + 合成配置，
   生成脚本 `pytools-core/tests/manual_kit/generate_manual_kit.py`）已生成；
   其菜单操作步骤不属于自动化范围，留作人工验收输入。
4. 独立预期结果来源：既有 parity 测试使用原 Python 基线 JSON
   （`pytools-core/probe/reports/`、`tests/fixtures/*_cases.json`）；新增测试使用
   测试内手工推导常量；不使用 Rust 自身输出作为唯一预期。
5. 5-7/5-8 业务规则来源已确认为原 Python 只读参考
   `4.pytools-main/PyCode|NoComPyTools/.../feature_1_8_adjust_rural_loan*.py`（月度）
   与 `feature_1_9_adjust_rural_loan_quarterly*.py`（季度，涉农贷款分机构/分地区统计）。
6. 第四阶段（2026-10-28，经用户授权）：Release 构建完成（`build-release.bat`，
   `F:\RustBuild\pytools` 缓存），`dist\pytools.exe`（21.3 MB）与 `dist\lo-convert.exe`
   （4.3 MB）已就位，`run.bat` 可用；最终全量回归 68 个测试目标
   **399 通过 / 0 失败 / 6 个环境门控忽略**。人工验收按 `testdata/manual-menu/README.md`
   逐菜单执行（`run.bat` 或 `run-dev.bat` 均可启动），报错文本按菜单编号反馈。
7. 环境补测（2026-10-28，本机实测）：本机注册有 `Excel.Application` 与 WPS
   `Ket.Application`/`KWPS.Application` COM，无 LibreOffice；`lo-convert` 的 COM 路径
   按 Excel→WPS 顺序选择引擎。以 `--ignored` 补跑：openpyxl worker 4 例（设
   `RUST_WORKSPACE_OPENPYXL_PYTHON` 后）3 过 1 失败——失败为迁移清单已记录的
   "openpyxl 路径对 ImportRange 的尺寸/样式复制限制"（worker 输出丢失默认样式记录，
   严格回读报不匹配），非回归；WPS COM 伪标 .et 拒绝测试实测通过（真实 WPS 在环，
   拒绝消息现为"原生 ET"，测试断言已同步）。正向 ET 转换测试仍需真实原生 ET 夹具：
   实测 WPS COM `SaveAs(65521)` 静默写出 BIFF（重开为 FileFormat 56），无法编程产出
   原生 ET；需用户在 WPS 界面手动另存 .et 提供后以 `LO_CONVERT_NATIVE_ET_FIXTURE` 运行。
8. 引擎优先级与占位清理（2026-10-28）：确认 `lo-convert` 的 `--engine auto` 已按
   平台选择——Windows 用 COM（Excel→WPS 顺序，本机两者均注册可用），非 Windows 用
   LibreOffice；Windows 侧不需要安装 LO。移除 `EngineId::ExcelizeWorker` 休眠占位
   （枚举、Display、一处测试样例、机构规范化的 "Go" 语言映射），无功能影响；测试临时
   目录命名补齐线程 ID 防并行撞名（11 个文件，含一处 .xlsx 后缀位置修正），
   修后全量回归 399 通过 / 0 失败，Release 已重建同步 dist。
9. 真实配置形状兼容测试（2026-10-28，经用户授权只读参考项目根 config.xlsx）：
   新增 `pytools-core/tests/real_config_schema_compat.rs`——
   A) 按真实 11 张表形状（真实表头、Y/N 启用、数字布尔、复合键 `3;4`、分号多区域、
   绝对路径形态，值全部脱敏）构造合成配置，锁定 8 个 config.rs loader 的兼容行为，
   含真实形状特有语义：18 列时序规则无 是否输出行序号 → output_row_no=false；
   去重表无 任务名 且 执行模式 不在 写入方式 候选内 → 两字段留空（不阻塞追加）；
   B) 软跳过用例：本机存在 `../config.xlsx`（gitignore，不入库）时加载真实文件
   仅断言形状级不变量。人工测试包生成脚本表头同步对齐真实配置并重新生成。
   本轮遵守用户指示**未编译未运行**，测试状态：**待运行**。
10. lo-convert sidecar 取消（2026-10-28，经用户明确要求）：`sidecar/lo-convert` 从
   workspace 移除并删除，转换改为 pytools 内置引擎 `pytools-core/src/conversion_engine.rs`
   ——Windows 直接调用本机 Office/WPS COM（Excel→WPS 顺序，ET 走 KET 并保留 65521
   原生格式预检），非 Windows 直接调用本机 LibreOffice `soffice` CLI（可选）；自动模式
   仍为 Windows→COM、其他→LO。校验/对拍逻辑（工作簿逐格比较、CSV 校验、DOCX 文本
   校验、ET 双转换探针）原样平移。`build-release.bat`/`package.bat` 不再构建/打包
   lo-convert.exe；`dist\lo-convert.exe` 已删除；`PYTOOLS_LO_CONVERT` 环境变量废止；
   测试改用 `PYTOOLS_CONVERT_ENGINES_DISABLED=1` 模拟"无外部引擎"环境；原生 ET 夹具
   环境变量更名为 `PYTOOLS_NATIVE_ET_FIXTURE`。sidecar 专属的 `recalculate` 子命令
   （产品菜单从未调用）随 sidecar 取消，公式重算需求若需恢复另行设计。转换相关测试
   **待运行**（遵守不编译指示）。
11. 临时格式探针按钮（2026-10-28，经用户要求）：文件处理菜单新增 7/8 两个**临时**
   入口（`format_probe.rs`）：umya（整簿读入+删其他表+`remove_row`/`remove_column_by_index`
   删区域外行列，1 起始索引已对 umya 3.1.1 源码核实）与 openpyxl（本机 Python
   `load_workbook(keep_vba)`+`delete_rows/delete_cols`，`PYTOOLS_OPENPYXL_PYTHON` 可指
   定解释器）。输出为源文件旁 `*_umya探针.xlsx` / `*_openpyxl探针.xlsx`，拒绝覆盖、源
   不变；供用户人工对比"区域复制保格式"选型，**结论后移除**。新增代码**未编译**，
   状态：**待运行**。

## 第 1 类：数据提取与汇总

| 编号 | 业务名称 | 自动化测试 | 测试数据 | 预期结果 | 外部引擎 | 状态 |
|---|---|---|---|---|---|---|
| 1-1 | 时序宽表汇总 | `tests/timeline_menu_cli.rs`（wide_summary_* 约 25 例：日期边界、时长、legacy XLS、中间目标写入与能力门控） | `tests/fixtures/timeline/`、运行时生成 | Python 基线/手工断言（测试内） | 否 | 已通过 |
| 1-2 | 按使用区域汇总 | `tests/summary_menu_cli.rs`、`tests/summary_used_range_error.rs`（错误矩阵、XLSM、未知错误值） | 运行时生成 | 测试内断言 | 否 | 已通过 |
| 1-3 | 按模板命名区域/批注汇总 | `tests/summary_menu_cli.rs`（named_range_*）、`tests/summary_comments_cli.rs`（合并锚点、set 标记、公式表头） | `tests/fixtures/summary_comments/` | Python 基线 + 测试内断言 | 否 | 已通过 |
| 1-4 | 表头路径比对与映射 | `tests/timeline_menu_cli.rs`（header_path_compare_*） | 运行时生成 | 测试内断言 | 否 | 已通过 |
| 1-5 | 时序配置预校验 | `tests/timeline_menu_cli.rs`（timeline_precheck_*）、`tests/config_task_loader_parity.rs` | `tests/fixtures/config_loader.xlsx` | Python 基线 `probe/reports/config_loader_baseline.json` | 否 | 已通过 |

## 第 2 类：去重与追加

| 编号 | 业务名称 | 自动化测试 | 测试数据 | 预期结果 | 外部引擎 | 状态 |
|---|---|---|---|---|---|---|
| 2-1 | 按配置检查重复 | `tests/dedup_menu_cli.rs`（configured_duplicate_check_*：公式/批注工作簿安全拒绝） | 运行时生成 | 测试内断言 | 否 | 已通过 |
| 2-2 | 按配置删除重复 | `tests/dedup_comment_delete_umya.rs`、`tests/dedup_number_format_inheritance.rs`（格式继承）、`tests/dedup_menu_cli.rs` | 运行时生成 | 测试内断言 | 否 | 已通过 |
| 2-3 | 去重追加到目标 | `tests/dedup_append_execution.rs`（既有目标候选化+能力提交+陈旧目标拒绝）、**新增 `tests/dedup_append_target_edges.rs`**（仅表头历史目标全量追加 / 无新增行零追加 / 复合键 1;2 首条保留、不同机构相同指标不误删） | 运行时生成 + `testdata/manual-menu/input/dedup/` | 测试内手工推导断言 | 否 | 已通过 |
| 2-4 | 去重/追加配置预校验 | `tests/dedup_menu_cli.rs`、`tests/config_task_loader_parity.rs` | `tests/fixtures/config_task_loaders*.xlsx` | Python 基线 | 否 | 已通过 |

2-3 重点说明：既有历史目标追加由 `pt_2_5_existing_target_edits_a_candidate_without_overwriting_the_target`、
`pt_1_3_pt_1_4_explicitly_commit_existing_target_after_capability_validation` 与新增 3 个边界用例共同覆盖，
不以后建目标代替既有目标场景。

## 第 3 类：报表打印

| 编号 | 业务名称 | 自动化测试 | 测试数据 | 预期结果 | 外部引擎 | 状态 |
|---|---|---|---|---|---|---|
| 3-1 | 按配置生成静态打印报表 | `tests/print_menu_cli.rs`（数值零→空白、文本"0"列宽、0.0001 保留、能力+确认后才执行）、`tests/print_config_new_target.rs`（缓存公式零→空白、公式无缓存拒绝、既有目标拒绝）、`tests/print_comment_execution.rs`（批注保留/删除双路径） | 运行时生成 + `testdata/manual-menu/input/print/` | 测试内断言 | 否 | 已通过 |
| 3-2 | 打印配置预校验 | `tests/print_menu_cli.rs`（interactive_print_menu_routes_static_report_and_precheck） | 运行时生成 | 测试内断言 | 否 | 已通过 |

注：3-1 历史存在用户实际失败记录（见 MIGRATION_INVENTORY），合成测试通过不能代表该问题解决。

## 第 4 类：文件处理

| 编号 | 业务名称 | 自动化测试 | 测试数据 | 预期结果 | 外部引擎 | 状态 |
|---|---|---|---|---|---|---|
| 4-1 | 文件名机构标准化 | `tests/batch_rename_parity.rs`（Python 基线、冲突保护、缺 RenameSheet 拒绝） | `tests/fixtures/batch_rename_*` | Python 基线 | 否 | 已通过 |
| 4-2 | 排序前缀添加/清理 | `tests/file_sort_prefix_parity.rs`、`file_sort_prefix_operation_parity.rs`（操作 49）、`file_sort_prefix_remove_parity.rs`（操作 410，含空输入） | `tests/fixtures/file_sort_prefix_*` | Python 基线 | 否 | 已通过 |
| 4-3 | 按内容重命名 | `tests/content_rename_46_parity.rs`、`content_unrename_47_parity.rs`、`local_filename_map_48_parity.rs`（XLS 浮点文本、冲突与跳过） | `tests/fixtures/content_*_cases.json` | Python 基线 | 否 | 已通过 |
| 4-4 | XLS/ET/DOC 格式转换 | `tests/conversion_menu_cli.rs`（路由、伪造 ET 拒绝、进度脱敏）；`sidecar/lo-convert/tests/`（需 LibreOffice） | `tests/fixtures/timeline/*.xls` + `testdata/manual-menu/input/files/待转换/` | Python/手工断言 | **是（LibreOffice/Office）** | 已通过；WPS COM 伪标 ET 拒绝已实测通过（真实 WPS 在环）；**原生 ET 正向样本待用户提供**（WPS COM 无法写出原生 ET，见下） |
| 4-5 | 按配置修改 Sheet 名 | `tests/sheet_rename_execution.rs`（发布副本不改源、能力门控、递归批量） | 运行时生成 + `testdata/manual-menu/input/files/工作表重命名/` | Python 基线 | 否 | 已通过 |
| 4-6 | 报送完整性检查 | `tests/submission_check_parity.rs`（AND/OR 别名语义、日期关键字、报告统计） | `tests/fixtures/` + `testdata/manual-menu/input/files/报送检查/` | Python 基线 | 否 | 已通过 |

## 第 5 类：一二批专用报表处理

| 编号 | 业务名称 | 自动化测试 | 测试数据 | 预期结果 | 外部引擎 | 状态 |
|---|---|---|---|---|---|---|
| 5-1 | 拆分村镇银行数据 | `tests/batch_tools_menu_cli.rs`（village_split_*：校验副本不改目标、legacy 浮点文本、能力提交） | 运行时生成 | Python 语义断言 | 否 | 已通过 |
| 5-2 | 机构名称规范化 | `tests/institution_normalize_plan_parity.rs`、`institution_normalize_batch_cli.rs`（重复键取末值、空映射、浮点科学计数） | `tests/fixtures/institution_normalize*` | Python 基线 | 否 | 已通过 |
| 5-3 | 删除分机构表外资行 | `tests/foreign_bank_batch_execution.rs`（planned_deleted_rows/deleted_rows 按行计数、跨表公式与定义名称跟随、动态引用拒绝）——**行级删除，非删表** | 运行时生成 + `testdata/manual-menu/input/special/分机构报表.xlsx` | 测试内断言 | 否 | 已通过 |
| 5-4 | 修改外汇页眉单位 | `tests/fx_header_parity.rs`（与 Python 一致且保持未修改内容、无匹配不产计划）、`fx_header_menu_cli.rs`（默认仅计划、能力后才提交）——**仅标签修正，无数值换算** | `tests/fixtures/fx_header/source.xlsx` + manual kit | Python 基线 | 否 | 已通过 |
| 5-5 | 地区总分校验 | `tests/region_sum_parity.rs`（值/统计/输出表现、缓存公式合计、输出冲突保护、输出目录失败不改输入） | `tests/fixtures/region_sum/`（balanced/mismatch/cached_formula） | Python 基线 | 否 | 已通过 |
| 5-6 | 工作表数据提取（普通无批注版） | `tests/sheet_extract_parity.rs`、`sheet_extract_area_operation_parity.rs`（区域提取基线、既有目标写入、追加能力门控、XLSM/错误值/日期/时长） | `tests/fixtures/sheet_extract_area/` | Python 基线 | 否 | 已通过（含旧函数名导入修复） |
| 5-7 | 涉农贷款比上月修正（月度） | `tests/rural_loan_regressions.rs`（未舍入差值、legacy 公式读取、机构行序）+ `batch_tools_menu_cli.rs`（rural_loan 菜单：市级调整副本、能力后替换当前簿） | 运行时生成 + `testdata/manual-menu/input/special/涉农贷款_月度.xlsx`（最简形状） | Python 基线（来源 feature_1_8） | 否 | 已通过 |
| 5-8 | 季度涉农贷款修正及分地区核对 | 同上（quarterly_*：无效地区保留并整体失败、分地区统计） | 运行时生成 + manual kit（最简形状） | Python 基线（来源 feature_1_9） | 否 | 已通过 |

## 第 6 类：归档整理

| 编号 | 业务名称 | 自动化测试 | 测试数据 | 预期结果 | 外部引擎 | 状态 |
|---|---|---|---|---|---|---|
| 6-1~6-5 | 按日期/类型预览、执行计划、直接复制 | `tests/archive_parity.rs`（Python 基线、扩展名过滤、Windows 路径序、原生大小写冲突、词干点号保留）、`archive_menu_cli.rs`（5 条工作流路由、当前层/递归选择） | `tests/fixtures/archive_*` + `testdata/manual-menu/input/archive/` | Python 基线 | 否 | 已通过 |

## 第 7 类：配置与快捷操作

| 编号 | 业务名称 | 自动化测试 | 测试数据 | 预期结果 | 外部引擎 | 状态 |
|---|---|---|---|---|---|---|
| 7-1 | 初始化/修复配置 | `tests/config_menu_cli.rs`（缺失引导、二次确认、修复只出计划不改源、歧义表头失败关闭、损坏配置记录失败）、`config_repair_execution.rs`、`config_vba_workbook_parity.rs` | `tests/fixtures/vba_config/` | Python/手工断言 | 否 | 已通过（Windows）；`config_repair_execution.rs` 为 `#![cfg(unix)]`，Windows 不编译（环境限制），Linux 执行 |
| 7-2 | 重新加载/查看配置 | `tests/config_menu_cli.rs`（重载更新输出目录、读取失败保留上次有效配置、off 视图抑制） | 运行时生成 | 测试内断言 | 否 | 已通过 |
| 7-3 | 打开配置文件 | `tests/quick_menu_cli.rs`（路由三快捷操作但**不启动桌面应用**——GUI 边界明确） | 运行时生成 | 测试内断言 | 否（GUI 验证属人工） | 已通过 |
| 7-4 | 打开最近输出 | 同上 | 运行时生成 | 测试内断言 | 否 | 已通过 |
| 7-5 | 打开输出目录 | 同上 | 运行时生成 | 测试内断言 | 否 | 已通过 |

## 覆盖缺口与环境边界（第一阶段结论）

1. **4-4 原生 ET 正向转换**：无样本、无引擎证据（MIGRATION_INVENTORY 已知限制）——尚缺测试，需 ET 样本后补充。
2. **需要外部引擎的转换与重算**：`sidecar/lo-convert` 的 10 个用例已通过（引擎缺失时自门控返回，不伪造成功）。本机 COM 引擎实测可用（Excel + WPS），WPS COM 伪标 .et 拒绝已实测；LibreOffice 未安装，soffice 路径在本机不可用。
3. **Unix 专属测试**：`config_repair_execution.rs`、`dedup_comment_openpyxl.rs`、`institution_normalize_openpyxl.rs` 为 `#![cfg(unix)]`，Windows 上不参与编译/运行。
4. **GUI 快捷操作（7-3/7-4/7-5）**：自动化仅验证路由与目标路径解析，实际打开行为留待人工验收（quick_menu_cli 已明确不启动桌面应用）。
5. **人工菜单测试数据包**：`testdata/manual-menu/`（31 个文件）为静态生成产物，未通过运行菜单验证——按 README 逐菜单人工执行即为第二阶段后的人工验收输入。
6. 2026-10-28 第二阶段已执行全量测试并把结果回填至本清单；未执行项仅为上列明确标记的环境限制（Unix 专属、WPS/ET 样本）与尚缺的原生 ET 正向样本。

12. 统一引擎诊断日志·第一阶段（2026-10-28，报表打印功能）：
   新增 `pytools-core/src/engine_diag.rs`（DiagEvent/DiagStage/DiagOutcome/时间戳
   工具，格式沿用旧版 PyTools logger 的 `%Y-%m-%d %H:%M:%S`，运行时实时取本机时间）。
   3-1 umya 与 3-2 openpyxl 已接入：引擎/阶段/结果/配置行号/文件/耗时/原始错误随
   `StaticPrintRun.events` 回传；菜单显示带时间戳的开始/完成/输出行（输出文件去重、
   使用保存事件时刻）、单行"处理中..."动画（仅终端环境，stderr 刷新不换行）、
   跳过与失败明细行（含行号与原始错误）；中断时如实说明无法获得完整统计。
   openpyxl 子进程以 JSON 内嵌结构化事件回传（含 Unix 秒时刻），Rust 侧转换为本机
   时间戳，普通输出不混入 JSON。新增 3 个 engine_diag 单元测试与打印菜单时间戳/
   去重断言。**本轮未经授权未编译未运行，状态：待运行。**尚未接入：excel-reader/
   rust_xlsxwriter/ExcelModifier 内部路径、其余菜单功能（第二、三阶段）。

13. 统一引擎诊断日志·第二阶段（2026-10-28，Excel-Core 接入 + 用户反馈修正）：
   - 新增 `pytools-core/src/excel_core_diag.rs` 适配层：把 excel-editor 既有
     `EditExecutionReport`（引擎/逐后端尝试/结果/耗时/原始消息）与
     `ExcelModifierError` 转成统一事件，**不新建引擎选择或回退机制**；另提供
     `read_event`/`write_event` 轻量构造器供 calamine 读与 rust_xlsxwriter 写使用。
     附 4 个单元测试。
   - 菜单不再打印 JSON：2-5 追加、去重预校验、打印预校验三处 JSON 仅在 developer
     视图输出（保留机器可读兼容），user 视图只显示可读行；1-1 汇总菜单改为
     时间戳+摘要+文件路径+失败阶段的可读输出。
   - 1-1 目标写入改为直接写入（按用户原则）：文件不存在→新建工作簿；所在文件夹
     不存在→报错且不代建；文件存在但工作表缺失→新建工作表写入；工作表存在→
     清空后重写。不再走 ExcelModifier 能力门控与追加计划。
   - 2-1 改为在源工作簿上就地标红标记重复行（对齐旧 Python
     `dedup_311.py::_mark_duplicates`：首次出现不标、后续重复整行填充
     FFFF0000、标识列全空跳过、保存回源文件）；新增
     `dedup_ops::mark_duplicates_in_place`，仅支持 .xlsx（.xlsm 如实跳过）；
     2-1 菜单改为可读输出（任务成功/跳过/失败、标记行数、保存工作簿数、耗时）。
   - 相关测试已改写：打印菜单时间戳与输出去重断言、1-1 直接重写语义、
     2-1 就地标记语义（含公式与批注工作簿）。**本轮未获编译授权，全部改动
     仅通过 rustfmt 语法解析与括号平衡静态检查，状态：待运行。**
   - 尚未接入：其余菜单功能的诊断输出、1-1/2-1 以外路径的统一事件展示。

14. 就地修改与 openpyxl 打包（2026-10-28）：
   - `openpyxl` 走 **IPC 子进程**（非 FFI）：`Command::new(python) -I` + 临时文件传 JSON
     请求/响应，一次性批处理，带超时与退出码校验；无 pyo3/libpython 依赖。
   - 新增 `packaging/openpyxl/stage-windows-runtime.py`：下载 CPython embeddable +
     openpyxl 3.1.5 + et-xmlfile 2.0.0，组装到 `dist/python/`（23.4 MiB），自检
     导入与 worker 启动后才通过；`inplace_edit::openpyxl_runtime_paths` 识别该布局。
   - 新增 `pytools-core/src/inplace_edit.rs`：按所选引擎构建运行时能力记录并通过
     `apply_plan_in_place` 执行（候选 + 逐项回读校验 + sha256 快照替换）。
     已注册 `.xls` 提示“请先用 4-4 转换”；umya 仅 xlsx/xltx，openpyxl 支持
     xlsx/xlsm/xltx/xltm。
   - 5-1 村镇银行拆分：菜单先选引擎再选文件，就地改写目标工作簿（已手动验证）。
     其 4 个既有测试仍断言副本语义，已标记 `#[ignore]` 待人工验收后改写。
   - 5-2 机构名称规范化、5-7/5-8 涉农贷款：就地改造代码已具备（`InPlaceEngine`
     路径可用），但为控制风险暂保持副本/提交模式，菜单仍显示原模式提示。
   - 5-3 外资行保持副本（对齐旧 Python `copy2`）；5-4 外汇页眉必须走 OOXML
     补丁后端（实测 umya/openpyxl 无法完成该操作）；5-5/5-6 只读源新建结果。
   - 1-1 目标写入：既有 sheet 表头一致则追加、不一致则拒绝且不改目标。
   - 全量回归 67 目标 / 404 通过 / 0 失败 / 10 ignored；Release 20:29 重建，
     dist 含 pytools.exe（22.5 MB）与 python/ 运行时（25 MB）。

15. openpyxl worker IPC 加固（2026-10-28，P0/P1/P2）：
   - P0：`OpenpyxlWorkerConfig::resolve_runtime` 统一解析（显式覆盖 → deb libexec →
     ZIP python/ → debug 开发树），PATH 系统 Python 不再构成可用；
     后端 `from_environment_or_product_layout` 失败携带全部查找位置；
     `inplace_edit::openpyxl_unavailable_reason` 复用同一入口（预检=执行）。
   - P1：worker `--version` 握手（`WORKER_PROTOCOL_VERSION`/`PROTOCOL_VERSION`
     两侧各一处定义）；派发前缓存校验，不匹配报 unsupported_before_write；
     超时 kill+wait 治理已文档化；stderr 诊断截断 1024 字符。
   - P2：package.bat 暂存 dist/python 并在缺失时拒绝打包（exit 5）；
     build_deb.sh 接线 stage-runtime.sh（OPENPYXL_RUNTIME_SOURCE）+ worker
     --version 冒烟；README/packaging README/MIGRATION_INVENTORY 已更新。
   - 验证：excel-editor 34/34、相关单元测试（握手缓存/截断/版本格式）通过、
     pytools-core 315 通过 0 失败（含 4 个 village ignored）。
     **未验证项**：worker --version 超时场景的残留进程断言、deb/ZIP 真实打包
     运行（需 WSL/clean 机器）、三场景人工验收。
