# main 审计核实与修复（2026-11-13）

基线：main，fa4ba947a64300bbcac1134529848e34565b5925。审计参考 b1e6b068，不将旧结论当作当前事实。已读根 AGENTS.md、pytools-core/AGENTS.md 与工作树规范；未发现更深层 AGENTS.md。开始时没有已跟踪的未提交修改，原有未跟踪目录全部保留。只在本 main 工作目录写入；不提交、不 push、不合并、不发布、不更新 dist，不操作 feature/test 工作目录。没有全量测试、cargo clean、系统安装/卸载或真实业务原件测试。

## P0：2-3 既有 XLS 追加

**仍存在，已按用户明确答复改为安全拒绝，不宣称恢复 XLS 追加能力。** 菜单配置去重追加通过 run_dedup_append_direct_existing 进入此路径；旧代码 rfind 数字使 A12 拆为 A1/2，类型化对象直接放入 JSON 矩阵；SetFormulaText 没有 value 会变空。按最大坐标补空矩阵并从第 1 行写入，可以覆盖旧区域。Worker 仅抽查前三行且跳过空值，返回 verified 后 Rust 会 replace_target_from_candidate；快照仅防并发变更，不证明旧内容保留。

用户已选择“接受先安全拒绝 XLS 追加”。RustTools.rs 移除危险矩阵执行分支，改为扩展名不区分大小写的显式拒绝，发生在候选创建、COM 与提交之前；Python Worker append_rows_com 同样在任何 IO/打开之前拒绝，封住其他直接调用。原目标不被替换。去重、源数据提取与 XLSX 追加规则不改。

本轮支持口径：既有 XLS 的文本、数字、布尔、日期、空值、错误值、公式和缓存追加全部拒绝，不静默转空；没有声称 A12、多列和公式 COM 写入已修复。未来恢复须实现严格 A1 解析、逐项类型协议、只写新增区域、旧区域与结构全量验证和持久化后回读。公式及缓存仍须遵循现有 AppendCell/去重规则，不擅自改成缓存值追加。另一个特殊报表代码中的相似 rfind 不在此追加调用链，未扩大修改。

## P1：3-1 同目标 Sheet

**仍存在，已修复。** run_openpyxl_engine 内嵌脚本每条配置 out_row=1，明确为覆盖内容，不是条目累加；原 delete_rows 不清合并和尺寸。改为先解除旧合并、删除全部旧行，再清 row_dimensions/column_dimensions；保留 Sheet 对象、页面属性、打印区域/设置及页脚处理，仍按现有 apply_setup 和最终页码规则生成。不改 mode 1/2/3 的 copy_block。

新增 Python 测试直接提取并运行当前 Rust 内嵌完整脚本，使用临时合成源：第一条 A1:B1 合并且行高70、列宽50；第二条 left/right 两个独立单元格。检查最终值、无残留合并、无旧尺寸、打印区域与页脚。测试不只是复制清理逻辑。

旧 run_umya_engine 的锚点字符串化仍在，但全仓调用搜索只有函数定义；注释明确历史保留。当前菜单 3-2 使用 RustWriter -> run_rust_writer_engine -> print_ops，并非该 UMYA 函数。用户未授权启用旧路径，本轮仅报告，不修或切换默认引擎。

## P1：Linux DEB

**仍存在，已分离。** build_deb.sh 原本普通打包后直接 dpkg -i、写 /opt 配置 Z99 再 dpkg -P。该段从普通打包移出。新增 isolated_install_smoke.sh：要求显式 RUSTTOOLS_DISPOSABLE_INSTALL_TEST=yes、Docker 标识、overlay 根且无受保护目录外部挂载；拒绝已有软件/配置。仅用于用户另行授权的一次性容器；本轮从未执行其安装路径。宿主安装不可用该脚本测试。

glibc 改为有效版本解析并比较上限，2.17/2.28/2.28.0 允许，2.28.1/2.29 拒绝，无法识别明确失败。不再要求恰好2.28。测试只提取检查片段，以 mock objdump 输出驱动 Bash，不执行构建/打包。

默认 PACKAGE_KIND=test，版本含 test 和分支名；release 需 main 且 RELEASE_ACCEPTED=yes（只能在实际验收后显式提供）。feature/test target 分别映射独立 F 盘构建目录；未知分支必须显式独立 target。未实际构建 DEB 或发布，防误发布不是正式验收。

## P2：内容前缀与无覆盖改名

**仍存在，已修复。** 生成、识别、撤销使用同一 canonical_content_prefix：最多180字符且UTF-8不超180字节。兼容旧版180字符截断前缀的识别与撤销；保留原有短前缀拆解兼容。完整新文件名需同时满足UTF-8字节和UTF-16单元不超240，为冲突后缀留余量；超长明确跳过，绝不截原文件名。路径层面的平台限制仍由文件系统拒绝并保留源文件。

Windows 改用 MoveFileExW，不设 REPLACE_EXISTING，不依赖硬链接，并编码绝对长路径/UNC；现有冲突递增与回滚复用该原语。Linux 使用 renameat2(RENAME_NOREPLACE)；若内核/文件系统不支持则明确失败保留源，不退回可能覆盖的 rename。其他未验证平台明确 Unsupported，不做复制删除的高风险回退。FAT32/exFAT/SMB 与麒麟实际行为尚需人工验收。

新增长中文、旧截断前缀、重复识别、真实临时文件改名/撤销与冲突原件保护回归；原有冲突/回滚测试一并运行。当前产品为 4-3 下添加/清理内容前缀，历史 4-6/4-7 名称仍有兼容包装；未修改菜单编号。

## 5-8

**已不可达，不恢复。** 菜单与命令分发无 5-8/quarterly-rural-loan-plan 入口。run_rural_loan_menu 的产品调用传 quarterly=false；季度分支仅残留实现。DEVELOPMENT_GOAL.md 与 MIGRATION_INVENTORY.md 最后追加的下线说明明确取代前文历史状态。未对季度校验算法做修复或大范围删源码。

## 验证记录

Python 专项：dist/python/python.exe -X utf8 -I pytools-core/tests/python/test_audit_safety.py，5 项通过。包含真实 Openpyxl 打印脚本、Worker 无IO拒绝、普通打包无宿主副作用静态断言、mock objdump 的安全 Bash 版本检查，以及 main/feature/test 正式包门禁执行检查。

Git Bash -n：build_deb.sh 与 isolated_install_smoke.sh 语法检查通过；未运行打包主体或安装主体。Python Worker 使用 AST 隔离执行拒绝函数，不导入/调用 COM。

Rust 前缀专项最终通过：cargo test --offline -p pytools-core --lib file_sort_prefix --no-fail-fast，7 passed / 0 failed，audit-safety-prefix-verified.exit=0；长路径修正后 audit-safety-prefix-final.log 也已通过。RustTools bin 内 XLS 拒绝单元专项通过：cargo test --offline -p pytools-core --bin RustTools audit_xls_append_guard --no-fail-fast，1 passed / 0 failed，audit-safety-xls-guard.exit=0（不是实际菜单端到端）。未执行全量测试。编译使用 CARGO_TARGET_DIR=F:/RustBuild/pytools、CARGO_BUILD_JOBS=1、CARGO_PROFILE_TEST_DEBUG=0。仓库新合入部分源码原本未格式化，本轮不保留 cargo fmt 导致的无关文件改动，修改文件单独 rustfmt。

## 最小人工步骤（待验收）

1. 在隔离目录放合成 XLS（表头、至少12行、AA列、各类型和公式），记录哈希；进入2-3选择既有XLS，确认提示停用、原哈希及目录文件列表不变。不要使用真实业务原件。当前预期是拒绝，不是追加成功。
2. 两条3-1配置同一新目标Sheet，第一条合并A1:B1/大尺寸，第二条独立值；核对无旧合并与尺寸、页面设置和页码。分别按现有 mode1/2/3 做实际菜单验收。
3. 在可丢弃的FAT32/exFAT/共享盘测试目录，用合成XLSX的超长内容添加、再次执行、撤销，并预置同名冲突文件；核对内容哈希及冲突文件未变。
4. 麒麟/统信一次性环境仅先审阅构建产物，不在宿主安装。DEB安装升级测试必须另行授权在无宿主挂载的全新一次性容器执行；确认配置标记升级保留。Linux renameat2、glibc实际产物、桌面和安装路径均 NOT_TESTED。

未验证：恢复后的XLS类型/公式/缓存保真（未实现）、Windows Excel/WPS COM端到端、FAT/exFAT/SMB、麒麟/UOS、实际DEB安装升级与正式包。自动化合成样本通过不等于真实业务人工验收。

## 修改文件清单

- pytools-core/src/bin/RustTools.rs：既有 XLS 追加 IO 前拒绝及针对性单元测试。
- pytools-core/python/pywin32_excel_conversion_worker.py：直接 Worker 调用同样拒绝，不保留危险写入逻辑。
- pytools-core/src/static_print_engines.rs：清理复用 Sheet 的旧合并和尺寸。
- pytools-core/src/file_sort_prefix.rs：统一前缀、旧前缀兼容、完整名称长度防护、无覆盖原生改名与临时文件回归。
- pytools-core/packaging/build_deb.sh：宿主安装分离、glibc 上限及测试包/正式包门禁。
- pytools-core/packaging/isolated_install_smoke.sh：独立、显式、一次性隔离容器安装测试（本轮不执行）。
- pytools-core/tests/python/test_audit_safety.py：安全专项自动测试。
- 本报告与 test-runs/audit-safety-* 本地运行记录。

注册表静态检查：Excel.Application 与 ket.Application 均存在。注册项不等于可用或验收通过，本轮未用它们执行 XLS 写入。XLS 拒绝测试使用临时占位字节文件，不将其冒充有效 BIFF 夹具。

## 首轮回归发现

首轮前缀专项 6 通过、1 失败：长前缀叠加临时目录超 Windows 传统路径上限，MoveFileExW 返回路径不存在。已补绝对长路径及 UNC 编码，未缩短测试数据规避；以 audit-safety-prefix-verified.log 的最终复测为准，早期失败记录保留。

注意：本轮不更新 dist。旧 dist/RustTools.exe 不包含本轮安全拒绝，不能用旧 EXE 验证修复，也不应继续对既有 XLS 执行追加。本轮 cargo test 生成测试程序，不等于交付可交互 EXE；人工菜单验证须先从本轮源码显式构建可交互程序。

格式与静态检查：修改的三个 Rust 文件 rustfmt --check --edition 2021 通过（audit-safety-fmt-changed.exit=0）；全仓 cargo fmt --check 返回1，差异仅为基线已存在的 excel-editor 模块顺序/plan 格式及 edit_plan/inplace_edit/lib 格式，已避免把这些无关格式变动纳入修复。Git diff --check 通过。

最终自动化结论：前缀7项、Rust拒绝1项、Python/脚本5项通过；无全量测试。真实菜单COM追加/打印人工验收、麒麟与实际文件系统/DEB安装均保持 NOT_TESTED。证据 test-runs/audit-safety-verification.json 记录最终修改源码哈希和测试退出码；HEAD仍为fa4ba947，无暂存、提交、push或merge。
