# 4-3 内容前缀位置定点修复

基线 main / fa4ba947a64300bbcac1134529848e34565b5925。开始时已有5个已跟踪修改文件及审计报告、脚本和测试记录；本轮只增量修改 pytools-core/src/file_sort_prefix.rs，新增本说明与 prefix-position-* 本地测试记录，不将此前修改归为本轮成果。旧 AUDIT_FIX_REPORT_20261113.md 和 audit-safety-verification.json 保留为上一轮证据；其中 file_sort_prefix.rs 哈希不再代表本轮源码。

## 规则

添加前的重复识别与撤销共用同一完整前缀匹配规则：文件名开头，或已配置排序前缀之后。添加现在读取与撤销相同的可选排序配置，不改变菜单或前缀内容取值。无匹配时添加仍在原名开头拼接完整前缀，不改写原名文本。

内容 A、原名 memo_A_file.xlsx、无 memo 排序配置时，不再误判重复；添加得到 A_memo_A_file.xlsx，撤销得到 memo_A_file.xlsx。配置排序前缀001时，001_A_file.xlsx可识别且撤销保留001。新180字节/字符上限前缀及旧180字符截断前缀保持兼容。

只移除一份完整前缀，不反复剥离，不按各个内容词拆删。例如内容A/B，A_B_A_file.xlsx只恢复A_file.xlsx，不再删成file.xlsx；仅有A_file.xlsx的旧部分词名称不猜测删除。重复执行添加会跳过，避免自动制造多层前缀。

若开头匹配和排序前缀之后匹配产生不同还原结果（例如内容A、排序前缀A_B、名称A_B_A_file.xlsx），添加明确提示位置有歧义并保留原名，撤销也不改名。未配置位置的中间文本不识别为旧前缀；需要用户自行核实旧命名来源。完整开头前缀按本功能命名约定识别，程序没有历史清单证明文件名来源，不推断残缺词或自动清除多层旧名称。

无覆盖 MoveFileExW/renameat2、Windows长路径及UNC编码、240字节/UTF16文件名限制、冲突后缀和回滚实现均未改。

## 测试

最终命令：CARGO_TARGET_DIR=F:/RustBuild/pytools，CARGO_BUILD_JOBS=1，CARGO_PROFILE_TEST_DEBUG=0；cargo test --offline -p pytools-core --lib file_sort_prefix --no-fail-fast。

prefix-position-verified-20261113.log/exit：11 passed，0 failed，160 filtered out，退出0；不是全量测试。既有7项通过（其中原“删除到file.xlsx”的断言按保护普通原名词的新契约改为A_file.xlsx）；新增4项通过：中间文本/单次撤销、配置排序及长/旧截断位置、歧义名称保护、实际合成XLSX添加/重复/撤销/重名及移动失败保护。

实际文件用例断言：memo_A_file.xlsx -> A_memo_A_file.xlsx，重复执行不改名；原名位置预置冲突文件后撤销得到memo_A_file_1.xlsx；两个文件字节各自保留。失败目标父目录不存在时源字节不变。歧义名称添加和撤销都保持字节与原名。既有长路径、重名和报告写入失败回滚测试继续通过。

涉及文件 rustfmt --check --edition 2021 通过；Git diff 空白检查通过。未运行全量测试、发布构建、打包、安装、提交或push；不修改feature/test目录，不清理原有未跟踪目录。

## 尚需人工验收

使用另行构建的新可交互程序和合成/备份数据：2-3 XLSX实际追加及旧区域保留；3-1 mode1/2/3精确值、合并、行列尺寸、打印设置及页脚；4-3真实菜单新增中间文本、已配置排序前缀及歧义名称；FAT/exFAT的冲突与失败保护；共享盘/UNC及断连后的实际文件状态；麒麟V10的renameat2、glibc和隔离DEB行为。上述本轮均NOT_TESTED，不把源码或NTFS自动化通过当作平台验收。

dist/RustTools.exe未更新（检查时间戳仍2026-11-11 21:12:29）；测试程序不是可交互产品EXE。旧EXE不能验证新保护，不应继续用于既有XLS追加。XLS追加仍按上一轮用户选择安全拒绝，未恢复能力。

## 续修：新旧截断候选歧义与真实配置流程

本次仍在main/fa4ba947保留全部既有修改，仅增量修改file_sort_prefix.rs与本说明，新增prefix-ambiguity-20261113本地测试记录。

leading_content_prefix_candidates返回全部新/旧截断还原候选；content_prefix_restore_candidates汇总开头及各个已配置排序位置的全部候选，不再将内层多候选提前折叠为None。添加和撤销均在最终候选数>1时明确提示“跳过：前缀位置有歧义，原文件未改名”；只有唯一候选可识别或撤销。无匹配仍按原业务处理，不推断中间文本。

新增2项测试实际通过：一项构造新旧截断都能匹配且结果不同的名称，验证开头/排序前缀之后的两候选均传递到外层，并通过公共业务入口读取合成配置，断言添加/撤销均明确歧义、原件与配置字节不变且无额外输出；另一项从合成config.xlsx读取“重命名配置”和“文件排序配置”，验证启用001、禁用memo，执行添加、重复、清理的完整库业务流程，核对排序前缀和中间普通文本保留，工作簿/配置内容不变。该流程不是交互菜单EXE端到端验收。

命令：设置CARGO_TARGET_DIR=F:/RustBuild/pytools、CARGO_BUILD_JOBS=1、CARGO_PROFILE_TEST_DEBUG=0后，cargo test --offline -p pytools-core --lib file_sort_prefix --no-fail-fast。结果13 passed / 0 failed / 160 filtered out，包含此前11项与本次新增2项；日志test-runs/prefix-ambiguity-20261113.log，exit=0。rustfmt --check --edition 2021 pytools-core/src/file_sort_prefix.rs及git diff --check通过。未跑全量测试。

此前未测平台和真实菜单场景继续保持NOT_TESTED；dist未更新、XLS追加未恢复、不提交不推送、不操作feature/test目录。

本次测试源码SHA256：87f3ecc8e8c5201d6a273dad4496000de78a7d708eed3dc11664afd61576f405
