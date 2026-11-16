# 2026-11-16 Feature/Test 集成与副工作树退役

## 集成结果

起点为主工作树 main `01a34ce90e518eeb103abed137867c8d62218059`。用户要求以 Feature/Test 为主完成必要集成，停止并行开发，并清理其他 Worktree。

- Feature：普通 merge 提交 `4f1ad27618b4ae8faa271f0c64985e358222fcbd`，完整合入分支唯一未集成提交 `2788e78a`，只涉及样本恢复说明和已有日志；没有重做样本恢复，也没有操作业务原件。
- Test：普通 merge 提交 `5c1c111aa9dbe5a4b68be6e1276d821a66f65e1c`，合入产品修复 `8b629f086aa410ec4964de7855d1711413866d5c`。包含目标可写性预检、使用区域汇总尾部空表头处理、可选输入目录预校验，以及请求内 Openpyxl 跨簿导入视图复用。
- 两处冲突均保留 main 当前业务：保留“先选择计算模式”的菜单说明；保留仅异常状态显示用户明细的日志规则，避免旧分支把正常阶段重新刷到用户终端。自动路由、Linux 引擎路由、Aspose 水印政策未调整。
- 新增回归 `used_range_summary_trims_style_only_headers_and_retains_numeric_zero`：合成输入包含空字符串、右侧只有格式的 Z 列和数值 0，确认不生成冗余表头、0 保留且源文件字节不变。
- Test 后续 8 个商业基准提交未合入产品树，原分支完整保留。其树中存在 EXE/PDB/NuGet 等生成物，因此没有整支导入。需要的历史源码、报告、已提交夹具均可从保留分支恢复。

商业引擎、rXLS 和 Rust XLSX modifier 实验改动均位于基准或报告范围，没有主产品调用链；Spire REA 分支约 6.7 万个差异文件主要是研究与证据，外加研究技能及忽略规则，不是当前生产依赖。这些实验没有合入 main。

## 实际验证

收到用户明确授权后，设置 main 的 `CARGO_TARGET_DIR=F:/RustBuild/pytools`、测试临时目录位于 F 盘，使用合成或测试夹具执行：

- `cargo test --offline --locked -p pytools-core --lib target_access_tests -- --test-threads=1`：1 passed / 0 failed。目标文件独占时拒绝，释放后成功，内容未改变。
- `cargo test --offline --locked -p pytools-core --test summary_menu_cli --test timeline_menu_cli --test summary_used_range_error --no-fail-fast -- --test-threads=1`：分别 25、33、6 passed，共 64 passed / 0 failed；时序目标默认忽略 2 个 Openpyxl 用例。
- 显式提供隔离 Openpyxl 运行时后执行 `cargo test --offline --locked -p pytools-core --test timeline_menu_cli -- --ignored --test-threads=1`：上述 2 个门控均通过，包括单次 Worker 多目标 Sheet 写入和宽松表头映射。
- `python -m unittest discover -s crates/excel-editor/tests/python -p test_openpyxl_worker.py -v`：30 passed / 0 failed，包含新增导入视图复用回归。
- Rust 针对性合计 67 passed / 0 failed；Python 30 passed / 0 failed。原始日志保留在 `test-runs/worktree-retirement-20261116/`；不同 PowerShell 捕获含 UTF-16 和 UTF-8，保留原字节，其中编译警告的 NativeCommandError 包装不代表测试退出失败，实际退出码均为 0。
- 合并片段只做必要格式化，最终 `cargo fmt --all --check` 及源码、文档的 diff 空白检查通过。完整 staged 检查仍提示门控原始日志结尾有空行，该日志未清洗，未将此警告报告为零。既有编译警告未借清理任务扩大修复。

本轮未运行新合并代码的全量、WSL/Docker 回归、WPS COM、FAT/exFAT、共享盘或麒麟安装验收。此前 `01a34ce9` 的全量通过记录不能推广为本次合并后的全量通过。

## 清理与保留

五个副工作树均已注销，只保留主工作树 main。四个目录已完全移除：Feature、商业基准、Spire REA 和 XLS 基准。

`Rust-Tools-test` 的 Git 注销和文件清理已完成，但根目录被其他进程占用。首次 `git worktree remove` 返回 Permission denied；随后确认目录为空、已不在 `git worktree list` 中，PowerShell 删除空目录也报告正在被其他进程使用。保留这个空壳，未终止其他用户进程。关闭占用它的旧终端或编辑器后可删除空目录；执行前仍应确认其中没有新文件。

清理前均核验了绝对路径、分支/HEAD、Git 状态、所有本地文件集合、大小和修改时间，以及修改文件和配置的哈希。REA 唯一已跟踪未提交修改 `AGENTS.md`、未跟踪 `.codex/` 和全部忽略文件已归档。该目录使用 `git worktree remove --force`，仅用于用户要求退役且已完成备份校验的工作副本，不删除分支，不改写提交历史。

- 五个工作树逻辑文件总量约 25.76 GiB。C 盘可用空间由清理前约 1.39 GiB 增至约 26.81 GiB，实测增加约 25.42 GiB；后台系统活动和硬链接等会使其与逻辑大小不同。
- `F:/RustBuild/worktree-retirement-20261116/` 保留 5 份本地文件 ZIP，共约 12.79 GiB、45,790 个文件；每个文件均从 ZIP 重新读取并通过 SHA-256 校验。所有归档哈希见同目录 `backup-receipts.json`，文件清单见每个 ZIP 的 `_BACKUP_MANIFEST.json`。
- 原分支、提交历史和已提交源码/夹具/记录保留在主仓库的 Git 对象中。归档只补充 Git 未保存的文件及未提交修改。没有删除任何本地分支，没有 reset、历史改写、prune 或 gc。
- 恢复说明见备份目录 `RESTORE.md`：以后按保留分支重建新的 worktree，再核验和解压相应本地备份；不要直接覆盖 main。
- 不运行 cargo clean，不清理 F 盘 Cargo target 或历史构建缓存。只删除了本轮生成、被完整备份替代的一份中断 ZIP 临时文件。
- 主目录 `config.xlsx` 与 `dist/config/config.xlsx` 各自保持原哈希，主目录业务文件和历史未跟踪内容保留。备份可能包含私人配置和业务副本，留在本地，不加入 Git、不上传。

## 后续方式及交付边界

根 AGENTS.md 和 WORKTREE_DEVELOPMENT.md 已注明只在 main 开发；历史 Worktree 表只是参考，不授权自动恢复并行。

本轮没有 push、更新 dist、重建 Release 或打包。dist 仍对应此前已推送的 `01a34ce9`，不包含本次 `5c1c111a` 的产品修复。若继续交付新版，应另行执行用户授权的验证与构建、更新版本凭据后再推送。
