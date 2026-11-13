# 通用功能首批 P0/P1 边界测试

main / fa4ba947a64300bbcac1134529848e34565b5925。仅补测试与本报告，无产品代码修改。开始时5个已跟踪产品修改全部以SHA256核对保持不变；原有未跟踪目录保留。不操作feature/test，不调整Linux路由，不做功能5通用化，不提交、不push、不更新dist。

## 新增文件与覆盖

- pytools-core/tests/general_boundary_p0_p1.rs：公共提交层同路径/规范化别名拒绝、损坏新候选零发布、Windows目标锁定和只读保护；归档预览零复制、计划后目标碰撞；归档源变更、自目标和嵌套目标保护断言；显式平台门控的Python/CLI入口。
- pytools-core/tests/python/test_general_boundaries.py：3-1 mode1/2/3复用同Sheet，逐项核对值/合并/行高/列宽/边框/零值/AA10/A100/打印区域/方向/缩放/页码及源哈希；2-3真实XLS拒绝与XLSX追加/重复执行对照；2-2跨表公式/命名区域/合并与文件锁导致失败时原件保护；4-3真实菜单的中间文本、重复执行、撤销重名、配置排序前缀。

3-1既运行当前Rust内嵌完整Python脚本，也经当前Cargo构建的真实菜单3-1读取合成打印配置。2-3真实XLS用Excel COM从零生成BIFF FileFormat56，检查OLE签名；只操作临时合成文件。CLI被测程序来自CARGO_BIN_EXE_RustTools，不是dist旧EXE。COM本轮仅用于合成夹具创建；未宣称XLS追加恢复。所有源/配置均只读，目标及失败文件的字节保护有断言。

## 执行命令和结果

Cargo统一环境：CARGO_TARGET_DIR=F:/RustBuild/pytools，CARGO_BUILD_JOBS=1，CARGO_PROFILE_TEST_DEBUG=0；既有专项另外显式设置PYTOOLS_OPENPYXL_PYTHON与PYTOOLS_OPENPYXL_WORKER、RUST_WORKSPACE_OPENPYXL_PYTHON。仅关闭测试调试符号，没有修改Cargo配置。

1. cargo test --offline -p pytools-core --test file_commit --test dedup_menu_cli --test print_menu_cli --test print_config_new_target --test archive_parity --no-fail-fast
   - file_commit：5通过。
   - dedup_menu_cli：17通过。
   - print_menu_cli：7通过。
   - archive_parity：5通过。
   - print_config_new_target：2通过、1失败。
   合计36通过、1失败，退出101。证据general-boundary-existing-20261113.log/exit。
2. cargo test --offline -p pytools-core --test general_boundary_p0_p1 --no-fail-fast -- --include-ignored --nocapture --test-threads=1
   新增Rust目标共9项：6通过、3失败、0忽略，退出101。其中一个通过项实际运行Python专项8项，8项全部通过（不能再把这8项加进Rust测试总数）。平台项默认ignore，显式执行要求Windows隔离Python/openpyxl/pywin32与Excel，不允许依赖缺失时静默通过。
   证据general-boundary-final-20261113.log/exit。
3. 将三种mode精确检查补到真实菜单后，仅重跑受影响门控项：cargo test --offline -p pytools-core --test general_boundary_p0_p1 synthetic_python_cli_and_print_boundaries -- --include-ignored --nocapture --test-threads=1
   Rust入口1通过；内部Python8项全通过、0跳过，退出0。证据general-boundary-cli-final-20261113.log/exit。其他Rust测试逻辑未变，无需无理由重复运行。

rustfmt --check --edition 2021 pytools-core/tests/general_boundary_p0_p1.rs及git diff --check通过。未跑全量测试。

## 失败项：保持红灯，不改产品或降低断言

1. 归档计划后源文件改变：期望拒绝复制，实际copied=1。当前计划未绑定源文件快照，执行复制了执行时版本；源文件本身未损坏。该项是计划一致性安全缺口，后续应确认并实现快照/失效策略。
2. 源目标为同一文件：期望拒绝，实际自动后缀后copied=1。原文件仍在，但产生了额外副本并报告成功；不是源文件丢失，属于源目标关系防护缺口。
3. 目标嵌套于源目录：预览把目标目录已有归档文件再次纳入源集合；当前未排除目标子树，有重复归档/增长风险。测试检查预览，无无限递归或真实业务写入。
4. 既有print_config_new_target中的pt_3_3_new_target_creates_a_plan_without_touching_source_or_target：执行旧print-config-plan + UMYA capability路径时，报OutputValidationFailed / copied cell value/type mismatch at Printed!C2。发生于旧机器接口执行分支，不等同当前菜单3-1/3-2失败；本轮未改该测试或引擎，根因及发布后状态须另行定点核实。当前打印菜单7项及新增3种mode菜单精确验证通过。

这些失败未标ignore、未改为“期望现状成功”、未删除；总体结论是测试补齐后发现缺口，不能称首批P0/P1全部通过。修复产品需下一步单独安排，本轮不擅自扩大测试任务为实现任务。

## 未验证及限制

- Windows使用本机Excel COM生成合成XLS并验证拒绝；没有WPS独立验收，没有既有XLS成功追加能力，也没有真实业务样本验收。
- 2-2含跨表公式/命名区域/合并的样本验证的是拒绝与原件保护，不代表这些结构可以成功删行并正确移动。锁定、只读测试不能替代完整Windows ACL无权限测试。
- 4-3真实菜单已覆盖XLSX；带真实VBA的XLSM、XLS添加/清理格式不对称仍待单项验收。
- FAT/exFAT、共享盘/UNC断连、麒麟V10、Linux环境和实际安装打包均NOT_TESTED；没有调整Linux路由。
- 本轮未覆盖所有类型/格式、0/10/50/100文件组合；旧日志不能替代当前工作区的新验收。
- dist/RustTools.exe保持2026-11-11 21:12:29，未复制或更新；不要用旧EXE验证当前保护或继续既有XLS追加。Cargo针对性测试按需生成的调试程序不等于正式发布产物。

## 证据位置

全部本轮记录位于test-runs/general-boundary-*，终端捕获位于test-runs/general-boundary-20261113/（各菜单及mode stdout/stderr/exit）。general-boundary-verification-20261113.json保存基线HEAD、已有修改未变核对、当前测试源码哈希和各命令真实退出码。失败测试临时目录仅含合成文件，由测试自身清理；没有清理原有未跟踪目录或构建缓存。
