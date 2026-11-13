本轮在 main 工作区验证通用边界修复。HEAD 为 fa4ba947a64300bbcac1134529848e34565b5925，包含此前未提交修改。记录目录为 test-runs/platform-validation-20261113/。指定检查已完成；WPS 的严格 BIFF 字节不变检查仍失败，仓库级格式检查有既有差异，未记为全部通过。

**修改及归属**

- 本轮产品代码仅修改 crates/excel-editor/src/openpyxl_backend.rs：隐藏数值零的目标必须同时没有值、没有公式。新增五行回归断言，模拟“公式存在但缓存为空”；修改前该断言实际失败，修复后 Windows、WSL、Docker 均通过。
- 原已存在但未跟踪的 test_general_boundaries.py：新增独立捕获目录，原业务断言及超时行为保留。test_audit_safety.py：抽取 Worker 方法时恢复原模块的延迟注解语义，兼容项目现有 Python 3.8；产品 Worker 未修改。
- 新增 test_platform_office.py、test_platform_xls_refusal.py、test_wps_readonly.py，分别验证独立 Excel COM、跨平台 BIFF 拒绝及 WPS 只读行为。COM 默认门控关闭，Linux 不导入或启动 COM。WPS 脚本不创建、配置或退出应用，只关闭自己打开的临时工作簿。
- 新增本报告及记录目录内的运行脚本、快照验证辅助、原始输出、退出码和合成副本。
- 起点的其他八个已跟踪修改、general_boundary_p0_p1.rs、根 config.xlsx 和 dist/RustTools.exe，结束前核对 SHA256 均与起点一致。三个混有此前修改的文件，在内存中仅逆向本轮补丁后，SHA256 均与起点完全一致，见 scope-proof.json；没有实际逆向或覆盖工作区。
- 其他既有未跟踪目录不重判归属，也未清理；不能把当前相对 HEAD 的全部 diff 当成本轮修改。

**实际通过的检查**

1. Windows：75 项 Rust 定点检查通过。7 个集成目标共 48 项；前缀 13、归档 5、零值/样式/元数据/关系 4、尺寸 5。嵌套执行的 8 项 Python/真实 CLI 检查通过，未重复加入 Rust 数量；当前脚本又独立复核 8 项通过。安全脚本 5 项通过。Excel COM 16.0 的 3 项检查通过，其中打印方法覆盖 mode 1/2/3。
2. WSL：64 项 Rust 定点检查通过：7 个集成目标 45 项，前缀 13、归档 5、零值 1。安全脚本 5 项通过。Linux 本地临时副本上的 CLI 实际执行 6 项通过，2 项 Windows 专属跳过；真实 BIFF 重放 1 项通过。临时程序逐字节比较、运行时全目录 diff 通过，见 wsl-isolated-copy.sha256。
3. Docker：Rocky Linux 8.9、glibc 2.28、Rust 1.99.1；21 项 Rust 集成检查及零值单测 1 项通过。实际调试 CLI 的最高 glibc 符号版本为 GLIBC_2.28。项目现有 Python 3.8.18/openpyxl 3.1.5 已在该容器运行。隔离副本逐文件 SHA256 核验通过；实际执行 6 项 CLI、1 项 BIFF 重放及 5 项安全脚本检查，全部通过。2 项 Windows 专属 CLI 检查跳过，不计通过。
4. 另在 Debian/Python 3.12 容器运行安全脚本 5 项通过。未授权的平台测试收集为 7 项 SKIPPED，证明门控生效；不计为业务覆盖。

Windows 主要命令：cargo test --offline -p pytools-core --test general_boundary_p0_p1 --test archive_parity --test archive_menu_cli --test print_config_new_target --test print_menu_cli --test dedup_menu_cli --test file_commit --no-fail-fast -- --include-ignored --nocapture --test-threads=1。CARGO_TARGET_DIR=F:/RustBuild/pytools，CARGO_BUILD_JOBS=1，CARGO_PROFILE_TEST_DEBUG=0。库级分别过滤 file_sort_prefix::、archive_ops::tests、range_copy_hide_zero、style_validation_tracks、metadata_follows_planned_sheet_identity、deleted_sheet_relationship_exemption、dimensions。

WSL 的精确命令保存在 run-wsl.sh、run-wsl-cli-isolated.sh；使用 --target x86_64-unknown-linux-gnu，构建仍在项目 F 盘缓存。Docker 的精确目标保存在 run-docker-rust.sh、run-docker-cli.ps1、run-docker-python.sh；源码及依赖只读挂载，容器无网络，单独使用主工作区 F 盘缓存子目录。没有执行全量测试、正式 Release 或打包。

**保留的失败和环境限制**

- WPS COM 12.0：最终只读补充检查为 2 通过、1 失败。打印三种模式及 XLSX 追加/重复执行通过；Excel 生成的合成 BIFF 在 ReadOnly=True、UpdateLinks=0 打开后就变化 16 字节，Close(False) 后仍不同。差异偏移为 552、554–560、25708–25715；两份文件均为 26,112 字节。值和 AA12 公式抽查通过，严格字节不变断言失败，未作语义豁免或恢复原字节来伪装通过。原始夹具只被复制；WPS 打开的是临时副本。前后文件、阶段哈希和差异保存在 wps-biff-diagnostic/，最终日志为 wps-readonly-final.log，退出码 1。
- WPS 首轮的新测试有默认行高写死 15、源/目标表头宽度不匹配的问题；按实际 StandardHeight 和一致的 27 列合成表头修正测试，未修改业务规则。之后重启 COM 曾遇到服务器失败/忙碌，改为只读活动会话验证；不把连接失败算通过。
- WSL 挂载路径运行时出现一次 30 秒 Worker --version 握手超时，导致补充 CLI 为 5 通过、1 失败、2 跳过。日志显示写入前拒绝发布；该次用例未完成成功断言。记录保留在 wsl-cli-final.log。字节一致的程序/运行时临时副本重跑为 6 通过、2 跳过，见 wsl-cli-isolated.log；原挂载路径问题没有产品修复，未延长超时或改变 Linux 路由。
- 首次并行编译耗时过长，曾结束本任务未完成的 WSL/Docker 编译并保留缓存，之后串行完成。Windows 首轮 COM 打印有 90 秒 CLI 超时，独立调试副本重跑通过。过程记录保留，不用失败/取消记录代替最终结果。
- Docker 补充测试准备阶段先发现没有 cmp，后发现默认 /tmp 为 noexec；前两次业务用例尚未开始，分别记录退出 1、126。已改用 Python 标准库逐文件核验，并仅对一次性容器的私有 tmpfs 开启执行。没有安装 diffutils，也没有修改宿主权限或配置。
- 自动审批拒绝了按旧 PID 关闭所谓空 WPS 测试会话：旧 PID 不足以证明当前归属，可能误关用户会话。该关闭操作未执行；后续 WPS 验证没有退出或重配应用。

**格式检查及未验证项**

涉及的 Rust 文件 rustfmt --check、git diff --check、5 个 Python 文件语法检查通过。cargo fmt --check 退出 1，差异在 crates/excel-editor/src/lib.rs、plan.rs，以及 pytools-core/src/edit_plan.rs、inplace_edit.rs、lib.rs。它们均不在本轮或起点的已修改列表中，与此前格式失败记录一致；没有为通过格式检查而改动这些文件。

NOT_TESTED：真实业务文件、FAT/exFAT、SMB/UNC 断连、硬链接文件身份和完整 ACL 矩阵、麒麟 V10 实机/桌面、DEB 安装升级卸载、正式包自动发现/双击布局、实际打印机、更大批量全组合、其余 COM 专用业务/公式重算及其余业务菜单。Docker 的 glibc 2.28 调试验证不等于正式 DEB 或麒麟实机验收。所有测试工作簿均为合成数据，不要求用户再用真实业务原件做破坏性验收。

2-3 既有 XLS 追加继续安全拒绝，未恢复 COM 追加；XLSX 追加及重复执行已验证。3-1/3-2 的默认引擎和 mode 语义未切换；旧 print-config-plan/UMYA 仅作显式回归。没有功能 5 通用化、Linux 路由调整、宿主安装卸载或 feature/test 工作目录操作。

HEAD 和暂存区未变，没有提交、推送、合并或发布。dist/RustTools.exe 未更新，仍是 2026-11-11 的旧 EXE；不能用它验证本轮保护，也不要用它继续向既有 XLS 追加。本轮实际执行的是当前编译的调试程序及校验一致的测试副本。
