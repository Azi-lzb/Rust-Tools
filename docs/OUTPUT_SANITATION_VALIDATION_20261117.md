# 主工作区 XLSX 输出数据净化验证记录

日期：2026-11-17。工作目录为 `C:\Users\AZI\Desktop\VibeCode\Rust-Tools`，分支为 `main`。

## 规范与实现

先单独提交根 `AGENTS.md` 的生产输出规范，提交号 `708afc2b`（`docs: define production XLSX output sanitation contract`），再实现代码。本轮代码保留在当前工作区，未另行提交；开始任务时已有其他未提交修改，未将它们纳入规范提交。

`tools/AsposeWorker/ResidentWorker.cs` 在保存前根据源工作表及有序结构操作推导业务清单，检查并清理可确认的系统附加表，按原活动表身份处理改名、移动、删除和可见性。序列化后调用 `OutputSanitation`，处理 SDK 保存时新附加的对象；最终候选才交给现有上层验证和安全提交流程。

当前系统对象识别限定为本次 SDK 保存新附加的 `Evaluation Warning`：其内容必须为空，或仅含固定结构的 SDK 提示文字。源文件同名表继续保护性拒绝，任务明确保留的业务表、空白模板和同名业务表不会仅因名称或空白状态被清理。其他名称、未知内容、公式、附属对象、危险引用及缺少可见业务表均拒绝交付。含 `INDIRECT` 等无法证明安全的动态引用时保守拒绝，本轮不扩展动态引用改写能力。

清理在私有候选副本内完成，更新工作表关系、内容类型、局部命名索引、扩展属性和视图选择；不重编号业务共享字符串。提示文字仅在已无业务单元格引用时清空其共享字符串槽。独立重开 ZIP/XML 检查业务清单和活动视图后，才替换内部候选；失败不发布候选，源文件不覆盖。

原始 SDK Save 归档和哈希保持原样。`evaluation_output` 等既有字段保持真实语义，新增 `output_sanitation` 记录实际删除及视图修正；成功和失败诊断关联 `run_id`、请求编号及清理阶段，不记录业务单元格原值。Word、其他引擎路由及历史证据不在本轮改动范围。

SDK 的评估附加表和活动视图行为另见 [Aspose 官方说明](https://docs.aspose.com/cells/net/evaluation-version-limitations/)。本轮结果只证明已测合成样本的输出契约，不据此声明授权状态或解除其他 SDK 限制。

## 已执行验证

使用现有 .NET SDK `F:/RustBuild/pytools/aspose-worker-sdk-10.0.404` 和 Aspose.Cells `26.10.0`；Worker 和单元测试均隔离构建到 `F:/RustBuild/pytools/output-sanitation-20261117/`，最终构建均为 0 警告、0 错误。没有运行 Cargo、Rust EXE 或更新 `dist`。

- 单元测试：`tools/AsposeWorker.Tests` 的 16 项通过，结果在 `F:/RustBuild/pytools/output-sanitation-20261117/unit-06/RESULT.json`。覆盖提示表、空系统表、业务空白表、同名业务内容、未知附加表拒绝、公式/动态引用/对象保护、局部命名索引、扩展属性、隐藏/删除活动表回退、有序操作、共享字符串保护和重复清理。
- 真实 SDK 定点测试：`test_output_sanitation.py` 的 6 项成功场景及 3 项预期拒绝通过，结果在 `F:/RustBuild/pytools/output-sanitation-20261117/sdk-04/RESULT.json`。成功场景为 Calculate、Convert、Modify、活动表改名并移动、删除活动表及连续处理；独立 openpyxl 读取检查业务表、活动表、公式、计算缓存、批注和命名区域。拒绝场景为已有目标、输入同名业务表和不安全动态引用。校验源哈希、原始归档哈希、失败候选清理及诊断事件。
- 既有结构操作回归：`test_plan_operations.py` 的 11 项通过，证据在 `F:/RustBuild/pytools/output-sanitation-20261117/plan-01/`。
- 共享/数组公式回归：`test_grouped_formula_preservation.py` 的 19 项通过，证据在 `F:/RustBuild/pytools/output-sanitation-20261117/grouped-01/`。
- 大型共享公式回归：`test_large_shared_append.py` 的 3 项通过，包含 20000 条共享公式和 1000 个追加单元格，证据在 `F:/RustBuild/pytools/output-sanitation-20261117/large-01/`；未提供 `--exe`，未运行可选 Rust 菜单链。
- 原始归档回归：`test_raw_sdk_archive.py` 的 7 组场景通过，证据在 `F:/RustBuild/pytools/output-sanitation-20261117/archive-01/RESULT.json`。原始归档保留 SDK 内容，最终候选执行净化；重复归档无覆盖，归档 IO 失败不发布候选。
- 6 个涉及的 Python 测试文件 AST 解析通过；相关已跟踪文件的 `git diff --check` 通过。`test_resident.py` 的预期已同步，但该脚本需要运行 Rust host，本轮未执行。

单元测试初轮发现 ZIP 更新模式在拒绝时仍可能改写文件头，已改用私有副本并复测哈希保护；补充扩展属性用例发现格式化 XML 的空白影响分组识别，已改为读取具体文本子元素并复测。失败目录 `unit-01`、`unit-04` 及中间成功目录保留，不删除既有证据。最后两处扩展属性修正后执行了最终单元和 SDK 定点验证；先前结构/公式/归档回归没有全量重复执行。

## 验证边界

全部自动验证使用隔离合成文件。Rust 编译、Rust host 端到端、完整菜单回归、真实业务样本、Windows 用户人工验收、Linux/麒麟/deb 和正式打包均未执行；不将 .NET 构建或合成样本成功视为生产发布验收。当前 SDK 测试主要覆盖保存时附加提示表；不能据此声称所有其他 SDK 版本或未知系统对象都已支持。
