# Windows 人工验收测试包（2026-11-13）

本次按用户授权，在 Rust-Tools 主工作区 main 上重新构建 Windows Release 模式程序、更新 dist、生成本地人工验收 ZIP，并提交和推送本轮累积修复。测试包不代表所有菜单或真实业务样本已经验收；不创建 GitHub Release，不生成新 DEB，不改变 Linux 引擎路由。

构建前基线为 fa4ba947a64300bbcac1134529848e34565b5925，本地与远程 main 一致。已有九个已跟踪修改、相关测试及报告属于此前各轮累积成果，本次打包没有重新归为新增业务修复。原始业务样本、config.xlsx、output、备份及无关未跟踪目录不提交、不清理。

## 本包包含的累积修复

- 2-3 既有 XLS 追加在候选创建和 COM 调用前明确拒绝；Worker 直接调用也拒绝。没有恢复 XLS 追加，XLS 读取和其他既有专用 COM 路径不关闭。
- 3-1 同目标 Sheet 复用时清除旧合并、行高、列宽；保持原 mode 1/2/3 和页面处理规则。
- 4-3 内容前缀统一生成、重复识别及撤销位置；保留中间普通文本、排序前缀和新旧截断兼容；多种还原结果时提示歧义并跳过。无覆盖改名和长路径保护随包交付。
- 归档计划保存源 SHA256，执行时拒绝规划后变化、同源目标和旧无快照计划，递归时排除目标及结果目录；旧计划需要重新预览。
- 显式旧打印路径校验正确处理数值零隐藏和删除的种子 Sheet，保留其他 Sheet、样式、元数据和关系保护。
- DEB 普通打包不再安装、改写或卸载宿主程序；隔离安装测试单独显式执行，glibc 上限和测试包/正式包门禁已有专项覆盖。本次没有构建或安装 DEB。

## 已有验证与限制

此前平台记录见 PLATFORM_BOUNDARY_VALIDATION_20261113.md。当前 15 个源码/测试文件的 SHA256 与该记录一致：Windows 75 项、WSL 64 项、Docker 22 项 Rust 定点检查有通过记录；不把本次打包称为重新执行这些 Rust 测试。

WPS 严格 BIFF 字节不变检查仍有一项失败：只读打开合成 XLS 后出现 16 字节变化。不能将 Excel COM 的通过推广为 WPS 全部通过。WSL 挂载路径曾发生 Worker 握手超时；字节一致的 Linux 本地临时副本通过，但原挂载问题没有产品修复。

NOT_TESTED：真实业务文件及用户人工验收、FAT/exFAT、共享盘/UNC 断连、麒麟 V10 实机、实际打印机、DEB 安装升级卸载、其余 COM 专用业务及完整菜单组合。Spire 没有接入，Linux XLS 通用修改没有新增实现。

## 最小人工验收

请完整解压 ZIP，在解压目录双击 RustTools.exe；保留 python 和 config 目录。使用合成或备份数据，先核对配置中的输入、输出和日志目录，避免默认路径指向业务原件。

1. 2-3：已有 XLS 目标应提示安全停用，原文件不变；另用 XLSX 检查旧表头/旧数据保留、只追加新记录，重复执行不重复写入。
2. 3-1：两条配置写同一目标 Sheet，先写 A1:B1 合并及大行列尺寸，再写两个独立单元格；分别验 mode 1/2/3 的值、合并、尺寸、页面设置和页脚。
3. 4-3：内容 A、memo_A_file.xlsx 且 memo 未配置排序时，添加再清理应恢复原名；验证配置的排序前缀、长内容、旧截断名称和重复执行。歧义应明确跳过；预置重名文件不得被覆盖，失败时源文件应保留。
4. 归档：预览后修改合成源文件，执行应失败且不发布副本；旧计划重新生成；目标子目录的旧输出不应再次入计划。

此前报告中的“dist 未更新”“不提交、不推送”是对应历史任务的状态，保留原文作为证据；本次构建、打包和 Git 交付结果以本文件末尾及当前提交为准。

## 本次实际构建和包验证

- 设置 CARGO_TARGET_DIR=F:/RustBuild/pytools、CARGO_BUILD_JOBS=1，运行 cargo build --release --offline -p pytools-core --bin RustTools：退出 0，release optimized，耗时 12m55s，有既有编译警告。
- dist/RustTools.exe 已更新为本次构建，31,186,944 字节，SHA256：236192967c5f2b92a799be1acc5fd7b54a301d9295909b235a654827ce0ddb95。
- 同步两个 Worker 源码到既有 dist/python，不下载安装依赖；运行时自检确认 openpyxl 3.1.5、et_xmlfile 2.0.0、pywin32 312 及协议握手通过。
- 使用 dist/python/python.exe -X utf8 -I 执行 test_audit_safety.py：5 通过；test_general_boundaries.py 指向本次 Release EXE：7 通过、1 项创建 COM 夹具的检查跳过。覆盖真实菜单 XLSX 追加/重复、打印 mode 1/2/3、前缀冲突、锁定与结构保护。
- test_platform_xls_refusal.py 复用此前 Excel 创建的合成 BIFF 副本：1 通过，确认 XLS 追加明确拒绝，源、目标、配置不变，无候选遗留。本次没有启动 Excel/WPS COM。
- 修改的 8 个 Rust 文件 rustfmt --check --edition 2021、Git diff --check 通过；cargo fmt --check 仍退出 1，差异仅在此前未修改的 excel-editor/src/lib.rs、plan.rs，以及 pytools-core/src/edit_plan.rs、inplace_edit.rs、lib.rs，没有清理这些既有差异。
- 使用 packaging/windows-package.ps1 生成 ZIP，再由本次 finish-package.py 加入测试说明和 BUILD_INFO.json。最终 ZIP 完整性、EXE/Worker/config 字节比对、独立工作目录下解压程序自动发现配置和启动菜单均通过。
- 使用 packaging/openpyxl/verify-windows-zip.py 验证最终 ZIP：ALL PASS。实际执行解压后的 Worker 握手、依赖导入、RustTools apply-edit-plan、输出值/列宽回读及源文件快照校验；没有仅凭 ZIP 存在判定成功。

本地测试包：dist/RustTools-windows-x64-TEST-main-20261113-173519.zip，32,968,480 字节。

ZIP SHA256：ada353e68f8c0ed1bdc87c6c9dfd66be49044b232c87bc61e1f86d79e5d437c3。

首个 ZIP 的追加说明步骤实际失败：PowerShell ZIP 原文件头使用反斜杠，Python 追加重写目录后改用斜杠，导致头部路径不一致。失败 ZIP 和日志保留在本地证据目录；重新打包并将全部条目重写为一致的斜杠路径后，才通过上述完整验证。没有修改产品代码或默认打包脚本来掩盖失败。

证据目录：test-runs/windows-package-20261113/。verification.json 记录最终检查、产物哈希和源码/配置不变验证；previous-dist/ 保留更新前的 EXE 与两个 Worker，仅作本地备份。两个本地 config.xlsx 均未修改。构建、测试和打包期间，全部起点源码/测试文件哈希保持一致。

本轮提交包含此前累积修复、专项测试、历史报告、平台证据摘要和本次交付记录；dist、ZIP、本地配置、业务样本、原始捕获及备份保持本地，不推送二进制或创建 GitHub Release。没有操作 feature/test 或其他工作树。
