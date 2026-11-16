# 2026-11-16 全量验证、测试包和提交交付

起点：主工作树 `C:/Users/AZI/Desktop/VibeCode/Rust-Tools`，分支 `main`，HEAD `34dfcd905b1157dca531908083b156b41f440a03`。开始时本地 HEAD 与远端 main 一致，有 35 个已跟踪修改文件以及多批历史未跟踪文件。没有 pull、切换分支、reset、merge，也没有操作 feature/test 工作目录。

## 最终实际结果

三个平台均完成默认 workspace 全量测试，每个平台 80 个结果组（包含空结果组和文档测试）：

- Windows：671 passed，0 failed，33 ignored，退出码 0。
- WSL Ubuntu，原生 `/tmp` 测试目录：663 passed，0 failed，22 ignored，退出码 0。
- Docker Rocky 8 / glibc 2.28：663 passed，0 failed，22 ignored，退出码 0。

命令：`cargo test --offline --locked --workspace --no-fail-fast -- --test-threads=1`。WSL 使用 `--target x86_64-unknown-linux-gnu`；所有构建缓存位于 main 对应的 F 盘目录及其平台子目录。没有执行 cargo clean。没有把 ignored 算成通过，也没有运行 `--all-features` 或上游 vendor 全套测试。

额外执行，单列以免重复计数：

- 独立 Rust resident host：11 passed / 0 failed。
- Aspose 修改、故障注入、明确 Openpyxl 路由：9 passed / 0 failed。
- Aspose 计算：随包 apphost 和 SDK/DLL 两种调用各 3 passed / 0 failed。包括 2 文件复用、100 文件在第 99 次后换进程、未知函数不误报通过。
- Openpyxl 门控：Windows、WSL、Docker 各 7 passed / 0 failed。
- Windows VBA 读取对照 1 项、COM 格式转换 2 项、转换菜单 1 项、XLS Sheet 改名 1 项、ET 内容识别/候选拒绝 2 项通过。
- Python CLI discovery：Windows 28 个方法、21 passed / 7 skipped；WSL、Docker 各 28 个方法、19 passed / 9 skipped。
- Excel COM 专项：3 个方法通过，其中打印方法分别检查 mode 1/2/3；另检查真实合成 BIFF 的追加拒绝、XLSX 追加与重复执行。读取、关闭后的文件哈希不变。
- 先前基准框架的分析/协议测试 16 项、四维框架测试 11 项，本轮实际通过。它们读取历史实验缓存，部分断言保护旧实验时点的 dist；这组证据是在更新 dist 前检查的，不是新版 EXE 的哈希验收。原始评估工作簿和二进制不提交，复现历史结果断言需先生成相应本地实验数据。
- Windows 打包脚本的 3 项隔离回归通过：测试包命名、自包含 Worker 完整性及符号排除、缺失 Worker 拒绝、正式包 main/验收门控。
- `cargo fmt --all --check` 通过。原有五个文件的排版差异只做格式化。业务源码的 staged diff 检查通过；完整检查仍报告 vendor 上游已有空白差异和原始终端捕获的结尾空行，均保留，未清洗原始记录来伪造零告警。与原始 crate 归档按规范化行尾对拍，仅 calamine `src/xls.rs` 和 UMYA `src/structs/data_validation.rs` 有预期补丁差异；归档 SHA 和检查结果随证据保留。

具体命令与原始结果见 `test-runs/release-push-20261116/` 的脚本、结果索引及捕获。完整本机原始数据保留于 `F:/RustBuild/pytools/release-push-20261116/`。

## 首轮失败及处理，未覆盖旧记录

- Windows 首轮：669 passed / 2 failed / 33 ignored。一项 XLS 删除重复用例漏选新执行模式，实际取消了任务；只补模式输入，原保留行断言不变。另一项“填充已有空 Sheet”单独两次及最终全量复跑通过；首轮具体原因未能从当时断言确认，仅增加失败诊断，没有放宽成功、数据和其他 Sheet 的断言，也没有猜测性修改业务逻辑或超时。
- 多个 Python 验收脚本仍使用新增执行模式之前的输入，首轮 10 failures / 3 errors；补上明确模式选择后通过。Excel 打印专项中同类遗漏也已修正，mode 1/2/3 的值、合并、尺寸和打印属性断言保留。
- WSL 首轮：636 passed / 27 failed / 22 ignored。其中 25 项把临时文件放在 F 盘 DrvFS 上，触发原子无覆盖改名限制；另 2 项与 Docker 首轮的 2 项相同：测试错误地要求 Linux 接受 Windows 专用 Aspose 路由。现在 Windows 保留原肯定断言，Linux 明确断言拒绝且不写文件。没有启用或调整 Linux Aspose/LO 路由。
- WSL 改回原生 `/tmp` 后重新全量验证。DrvFS 未被宣称兼容：独立合成探针确认目标冲突时 errno 17，新目标执行 `renameat2(RENAME_NOREPLACE)` 时 errno 22；源、已有目标和配置字节均不变，没有生成改名后的文件。业务日志只给出 `FILE_RENAME_FAILED`，原始 errno 另由底层探针记录。当前策略是安全拒绝，不回退到覆盖改名或先删除源。
- SDK/DLL 计算门控在并发编译时曾有 2 项 30 秒期限超时；停止编译后用相同默认期限、相同 SDK/DLL 调用方式复跑 3/3 通过。随包 apphost 的相同 3 项也通过。保留两轮日志，不用当前空载成功抹去负载敏感性记录；本轮诊断时间不作为新的性能基准。
- 状态记录脚本在运行期间被修改，导致两份首轮退出标记漏写。原 Cargo 日志完整且明确失败，外层执行实际返回 1；该脚本失误及补记的外层状态单独存于 `INITIAL_WRAPPER_ERRORS.json`。最终轮使用固定脚本并正常记录退出码，未把首轮算作通过。
- 初次 Windows 打包回归遇到测试采集器 GBK/UTF-8 解码失配，已显式使用 UTF-8，原错误捕获保留；Words 发布首次缺少缓存源，增加已有 F 盘缓存后构建通过。一次复制 Worker 因 C 盘空间不足中止，残余只属于本轮生成文件，已完整移存 F 盘；最终 dist 为实体 Worker 目录，不依赖此前临时目录引用。

## 与此前修改的边界

本次提交纳入此前已经完成但未提交的 Aspose Cells/Words 独立 Worker、Rust trait/适配器、明确选择入口、Rust 打印对齐、50 文件参考耗时及两项 vendor 修补。本轮没有重做这些业务规则，也没有改自动路由或 Linux 产品路由。

本轮新增或修正的是 Windows 打包/门控检查、验收脚本、遗漏的模式输入、平台契约断言、空 Sheet 失败诊断、格式化和本交付记录。`CHANGE_LIST.txt` 列出实际提交范围。历史 output、备份、其他未跟踪验收目录仍保留，不将来源不明的目录一并提交。

另外补齐了默认 reader 测试依赖的 `synthetic_probe.xlsx` 的 Git 跟踪：它在本轮前已存在，生成来源未确认；已审阅全部单元格及包部件，为占位文本/数字/布尔值，没有宏、外链、真实机构名或评估水印。内容未改，三平台全量均实际读取同一字节。SHA256：`2a31ad9ad545594fb6f85076ae42e98071ecffee1be3641f886ac65525001da0`。没有把本地业务文件作为这个夹具提交。

## 制品验证

- `dist/RustTools.exe` 已更新，SHA256 `67ca06c64d9c00eada0e8e45ecf00c7eaa2e2fdd721e4dd8e433d66a039e0712`。
- Windows ZIP：`RustTools-windows-x64-TEST-main-20261116-093030.zip`，139,000,105 字节，SHA256 `109537b64b7ed2aa9618ece87767d8197a68b37e389eb74cf2b5336e7a5c1bbf`。
- Linux DEB：`rusttools_0.1.0-20261116-test-main_amd64.deb`，32,428,976 字节，SHA256 `9deb2279c45cc6f936bae47b647172f74dd0f4733ead743a9cdb3c4d0fcb3119`。
- Linux Release 二进制 SHA256 `e73d48701ba4ae534f813c1f7abecfda45cdfa388b4df75bf040e32e46355334`，最大 glibc 符号 2.28。
- ZIP 解压后：Python/pywin32/Worker 握手、真实 apply-edit-plan、源快照保护、Sheet 改名及引用/缓存、重复执行、三条追加路径、既有 XLS 追加拒绝通过。391 个 Worker 文件逐项核对。清除 SDK/Worker 环境覆盖后，Rust 通过随包 apphost 完成页眉修改及 Word 转换，水印保留。最终 ZIP 与功能实测包仅测试说明文字不同，所有功能载荷逐文件一致；本地实体 dist 的相同专项也通过。
- DEB：构建门控及解包后的 Openpyxl 修改通过；在只读宿主/只读仓库的 Rocky 8 容器中运行解包程序，动态依赖、Sheet 引用/缓存、追加及 XLS 拒绝通过。没有执行宿主 dpkg 安装、卸载，也没有改宿主 `/opt/rusttools/config`。
- 包含自包含的 Windows Cells/Words 评估 Worker；无 SDK 安装要求。Linux 不包含这些 Worker。原始水印输出、DLL/PDB/SDK 中间产物不进入 Git；分发二进制只放被忽略的本地测试包和 dist 中。
- 根 `config.xlsx` SHA256 始终为 `b8211c33ffbc35f12b762d6eb9a10c9e3c95377129c4c7a1177eb8c510be256f`；`dist/config/config.xlsx` 始终为 `04b6535db132c7f0f41cf9bee6db354e126c46e5afc403934595a554e35d0cec`。两份均保留、不提交。ZIP/DEB 按原打包规则使用根配置副本，本地 dist 保留自己的既有配置。

旧报告中的“未打包/未更新 dist/未提交”是历史时点。本轮不重新计算 50 文件基准，也不把全量通过等同于 50 文件审计中的 24 个门控路径已实现或已对齐；例如 Aspose XLS 含批注的转换安全拒绝仍保留。现有 XLS 追加没有恢复。

## NOT_TESTED 与最小人工验收

- Windows：在隔离目录和配置副本中，用合成或备份文件验证 2-3 XLS 明确拒绝且原件不变；XLSX 追加保留旧数据、重复执行不重复写入。3-1 各引擎分别用 mode 1/2/3，让两条规则复用同 Sheet，检查独立单元格、合并、行高列宽、打印区域、页面设置和页脚。Aspose 额外水印及活动页变化属于当前评估输出。
- 4-3：中间普通文本、开头内容前缀、配置排序前缀后、长内容、新旧截断名、重复添加、撤销及重名；歧义明确提示，不猜测性删除普通文本。已有目标不得覆盖，必要时使用可用后缀；失败时源文件不得丢失。
- FAT/exFAT、共享盘和真实文件锁/权限组合：仍需目标文件系统实测；不能把 NTFS、WSL 原生目录或容器结果推广过去。DrvFS 改名限制已实测确认，仍未提供成功改名支持。
- WPS 独立 COM、麒麟 V10 实体机、DEB 安装/升级/卸载及既有安装配置保护：本轮 NOT_TESTED。先用专用虚拟机/隔离环境，不在宿主已有安装上直接做冒烟。
- 用户真实业务文件、带复杂对象/外链的任意工作簿、宏在 Aspose 路径中的修改、授权版和无水印输出、未提供的真实 native ET 夹具，以及默认 ignored 中未单列执行的旧专用路径：NOT_TESTED。没有将编译或合成样本当成人工验收。

最终提交号和远端同步结果见 Git 及本地 dist 的 BUILD_INFO；不创建 GitHub Release，不合并其他分支。
