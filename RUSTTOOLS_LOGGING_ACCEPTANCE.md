# RustTools 日志验收

代码与自动化交付：PASS。人工业务样本和 Linux/deb 不在本轮自动化结论内。基线 main / af7059e100d5a0311f361b5a7ea1ec2bb3991f81；未 push、未合并，未更新 dist。

## 最终测试门禁

证据根目录：docs/evidence/rusttools-logging-20261112。

- cargo fmt --check：PASS，fmt-regression4.exit=0。
- cargo test --offline -p pytools-core --no-fail-fast：PASS，full-regression4.exit=0；合计 434 passed，0 failed，12 ignored。各测试目标原始结果完整保留，ignored 不视为通过。
- Openpyxl Rust Worker：8 项实际执行通过，包含 ignored 专项，worker-complete-final.log。
- Python Worker：25 项通过，worker-python-full-trace.log；包含 21 层深栈及原状态/退出码不变断言。
- 真实 CLI Python 专项：4 项通过，formula-cli-final.log；部分缺表、缺字段、全部缺表、受控公式计算失败。
- 受保护 Openpyxl 回退：1 项显式 ignored 测试通过，fallback-final.log。
- 宽表 Openpyxl 专项：2 项显式 ignored 测试通过，wide-ignored-final.log。首次缺少隔离运行时环境失败后，按测试要求设置运行时再通过；未修改业务代码。

Cargo 使用 F:/RustBuild/pytools、CARGO_BUILD_JOBS=1、CARGO_PROFILE_TEST_DEBUG=0，仅关闭测试调试符号以减少链接开销，业务优化级别和源配置不变。Python 使用 dist/python 隔离解释器；Worker 专项命令包含 --include-ignored --test-threads=1。全量和专项均在同一最终 Rust/Worker 源快照执行。

## 14 类场景证据与边界

1. 单文件成功：batch-1，CLI 捕获、实际输出和源字节对拍 PASS。
2. 10 文件成功：batch-10，同上 PASS。
3. 50 文件成功：batch-50，同上 PASS。
4. 100 文件成功：batch-100，同上 PASS。100 成功、提取/写入 200 行，含源表表头 100 行，实际业务数据 100 行；所有源字节不变。
5. 部分跳过：missing-worksheet、missing-required-header 及汇总 CLI 结构跳过用例 PASS；1 成功/1 跳过，实际输出仅有效来源 1 行。
6. 部分失败：user_corrupt_input_logging_keeps_file_identity_without_internal_errors 混合正常/损坏输入 PASS；整次输出中止时不把未输出文件计成功。
7. 文件损坏：上述用例及 interactive_menu_returns_failure_when_summary_input_is_corrupt PASS；用户简短错误，诊断留原链。
8. 缺工作表：missing-worksheet PASS；全部缺表 all-missing-worksheet 为 NoOp/跳过，无虚假失败或输出。
9. 缺必要字段：missing-required-header PASS，实际结构表头缺失，原件不变。
10. 全部失败：损坏唯一输入的真实 CLI 用例 PASS；失败 1，无生成结果。
11. 公式重算成功：conversion_writes_recalculated_formula_cache_for_calamine PASS，实际 Excel COM CalculateFullRebuild，Calamine 重读缓存值 27。
12. 公式重算失败：formula-recalc-failure PASS，真实 CLI + 受控 COM Worker 故障注入，成功 0、失败 1、生成工作簿/Sheet 均 0。不是实际 Office 崩溃验收。
13. 修改引擎回退：pt_2_5_chart_target_falls_back_to_openpyxl_and_keeps_charts 与显式命名依赖回退 PASS；真实文件、图表/依赖保存断言。通用 emitter 测试保证正常回退成功事件不展示技术过程。
14. 候选验证/提交失败：file_commit 和 dedup_append_execution 的无效候选、源变更/目标冲突保护 PASS；统计测试 candidate_and_failed_workbooks_do_not_contribute_completed_changes 只计 4 个实际已提交单元格。端到端菜单强制提交失败捕获 NOT_TESTED，此项证据为真实文件提交层与摘要统计层组合。

## 诊断、对拍与未测项

cli-evidence 保存原始 stdout/stderr、历史 JSON/JSONL、输出工作簿、statistics/verification；developer-sample.json 和 developer-verification.json 从最终 batch-100 刷新。Worker 原始堆栈响应及跨 Rust 错误链样例分别在 worker-trace 与 worker-rust-error.json。

文件清单 modified-files.txt；最终代码 SHA256 在 source-sha256.json。测试结果只覆盖自动化合成样本/临时夹具副本，不改根 config.xlsx、业务原件、Python 原项目或其他工作树。源文件字节、输出行数、页眉、公式缓存、图表保存等断言仍在原测试中。

NOT_TESTED：Windows 用户真实业务样本逐菜单人工验收、真实 Excel/WPS 计算故障、WPS 单独验收、原生 ET/Word/WPS 专项、人工 VBA 夹具、旧已移除样式导入路径、快捷操作真实桌面打开、Linux/deb。未为这些项目宣称通过，未生成正式包。

原始终端捕获保留空白行及行尾空格，不作内容清洗；源码/报告 diff 空白检查通过，原始捕获不纳入空白格式门禁。目录还保留早期测试捕获，最终门禁以 full-regression4 与 formula-cli-final 及其关联 run_id 为准。
