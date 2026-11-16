# Aspose 显式计算接入（2026-11-15）

后续同日的完整菜单扩展和性能结果见 `ASPOSE_FULL_MENU_TRIAL_20261115.md`。
下文保留此前阶段的实现与验收记录；后续能力范围以新报告和 DELIVERY.json 为准。

实际入口是 **1-5 按模板校验区域提取校验**，1-4 是批注汇总。选择本次计算模式后，
才进入原有模板与文件选择流程：1 自动、2 Excel/WPS COM、3 BetterOffice、4 Aspose。
自动保持既有配置与路由；强制模式禁止跨引擎回退，选择不写入 config.xlsx。

命令行：`RustTools summary-named-validation --calc-engine aspose --template TEMPLATE.xlsx --output-dir OUTPUT --log-dir LOG SOURCE.xlsx ...`。
省略 calc-engine 等价于 auto。BetterOffice 强制模式沿用原预算和轮数，但禁止回退。
本轮强制 COM/Aspose 仅接入 Windows，Linux 的既有路由未改。

## 架构

共享进程适配器归入 `crates/excel-aspose`，包含 CalcEngine、ModifyEngine、
AsposeWorkerBackend。原 tools/aspose-worker-host 保留兼容转发，独立探针复用同一
实现。业务调用依赖 Trait，不直接启动 .NET；C# 工程仍独立位于 tools/AsposeWorker。

ConversionRuntime 在整批内持有一个 Aspose 计算会话，复用 Worker；第 99 个响应
返回后回收，下一请求重新启动。并未把所有既有业务修改迁移到 Aspose。
修改引擎菜单、4-4 纯转换接口及自动路由升级不属于本轮。

## 本地启动

优先使用 RUSTTOOLS_ASPOSE_WORKER_EXE；或同时配置 RUSTTOOLS_ASPOSE_DOTNET 与
RUSTTOOLS_ASPOSE_WORKER_DLL。均未配置时仅查找程序目录下
`workers/aspose/AsposeWorker.exe`。显式配置无效就报错，不扫描旧探针、不自动安装。

```powershell
$env:DOTNET_ROOT='F:\RustBuild\pytools\aspose-worker-sdk-10.0.404'
$env:RUSTTOOLS_ASPOSE_DOTNET="$env:DOTNET_ROOT\dotnet.exe"
$env:RUSTTOOLS_ASPOSE_WORKER_DLL='F:\RustBuild\pytools\aspose-worker\bin\AsposeWorker.dll'
& 'F:\RustBuild\pytools\debug\RustTools.exe' menu --config 'C:\Users\AZI\Desktop\VibeCode\Rust-Tools\config.xlsx'
```

再选择 `1 → 5 → 4`，使用合成或备份数据。dist 未更新，旧 EXE 不包含此入口。
SDK、DLL、PDB、bin/obj、测试水印文件均不提交。

## 文件和日志

只计算现有流程生成的 `副本/`。Worker 另存新候选；校验源快照、原 Sheet 顺序、
公式文本与位置、非公式数据、合并和批注后，复用 file_commit 带快照/备份的安全替换。
替换目标仅为副本，业务原件和模板不就地保存。公式比较不混淆字符串常量或绝对引用。

当前用户接受评估输出：副本与原始候选保留 Evaluation Warning 及活动页变化。
业务按 Sheet 名和校验区域提取，评估页不计入结果。暂仅计算 XLSX 副本，XLSM 不降级。
旧 XLS 输入仍先走既有副本转换规则；选择 Aspose 只影响计算，不替换转换引擎。

每次 CLI 报告保留原字段并增加 requested_calc_engine。实际引擎进入原诊断事件；
Worker 和每次运行独立的 host-calculations-*.jsonl 使用同一 run_id，记录 PID、阶段耗时、候选及提交。
计算和提交失败不算成功；业务校验异常仍提取为结果，不误报为进程执行失败。
校验区域出现未识别函数/名称错误时，候选明确拒绝提交，避免误判为“没有校验异常”。
新增 error_summary 兼容字段保留失败文件定位，旧 failure 字段不删除。
Worker 标准输入/输出显式使用 UTF-8，避免 Windows 本地代码页破坏“副本”等中文路径。

## 验证

仅运行模式/公式保护单元测试、aspose_menu_cli、真实 Worker 的两文件与 100 文件
专项及独立 host 协议回归。通过与否以本轮日志为准，不以编写了测试替代实测。
新增代码的定点格式检查通过；全仓 cargo fmt --check 的既有差异未清理，
归属核实见本轮 format-audit.json，不将它报告为全仓格式通过。

NOT_TESTED：新 Linux/WSL/Docker/麒麟/WPS 接入、真实业务人工验收、正式打包安装、
授权版、去水印/恢复活动页、修改引擎菜单和 4-4 Aspose 转换。原 xlwt BIFF 部分坐标
样式兼容问题仍由 Worker 的保存后校验拒绝，本轮不声称修复该库问题。


## 后续：显式修改引擎选择（2026-11-15）

上文的修改引擎 NOT_TESTED 是前一轮计算接入的历史状态。本轮增加 Windows
**2-1 重复标记、5-2 机构名称规范化、5-4 页眉单位修改**的修改模式：
1 自动（保留既有路由，包括专用 XLS COM）、2 Openpyxl、3 Aspose .NET。
5-2/5-4 在文件选择之前询问；2-1 使用配置中的输入，因此在读取任务配置之前询问。
选择只在当前任务线程/批次有效，退出、取消和报错后恢复；不写环境变量或 config.xlsx。
旧全局编辑菜单仍是原有五种模式，未把 Aspose 纳入默认或自动候选。

命令行 `institution-normalize`、`fx-header` 增加可选
`--modify-engine auto|openpyxl|aspose`。Aspose 当前要求原有 `--in-place auto`
执行模式；5-2 另需 `--apply --config CONFIG.xlsx`。省略新参数保持原行为。
显式 Openpyxl/Aspose 在这些入口只接受 XLSX，不转用旧 XLS COM；自动模式不受影响。

沿用上一节本地 SDK/Worker 环境变量，再运行 F 盘 debug EXE。未更新 dist，
旧 EXE 不含本轮修改模式。应使用未含评估页的合成或备份数据测试；原有
`Evaluation Warning` 的所有权无法确认时拒绝修改，不按名称删除普通业务 Sheet。
本轮保留原始评估页以及 SDK 可能改变的活动页，不声称输出无水印或恢复活动页。

Rust 业务继续构建 ExcelEditPlan，通过统一修改路由调用 AsposeEditBackend，
由共享 ModifyEngine Trait 调用外部常驻 Worker。每批复用同一进程，复用既有
99 次加载后排空并重启机制；丢失响应不自动重放。没有把 DLL 链接进 Rust。
当前支持文本、布尔、空值、可精确保留的至多 15 位整数、公式文本、RGB/不透明填色、
普通/首页/偶数页页眉文字替换；页脚不变。小数/指数数字文本、日期/错误值的显式写入、
结构删行/复制等未接通操作明确拒绝。未知或未验证的敏感结构拒绝修改。
4-5 Sheet 改名、5-1/5-3/5-7 结构操作等尚未增加 Aspose 选择，原路径保留。

Worker 保存独立候选，ExcelModifier 继续校验数据、公式、样式和计划声明的结构，
再由 file_commit 检查源快照并安全提交。评估校验只在内存比对中排除唯一追加的评估 Sheet，
不改写/删除水印；对已验证的默认行高显式化与 paperSize=1 作窄范围等价比较。
实际行高、隐藏、样式、默认尺寸及页面属性变化仍拒绝。Aspose 保存可能改变
Openpyxl 来源文件的默认列宽，Worker 因此仅在 Modify 的非结构操作后，从源 ZIP
恢复各原有 Sheet 的 sheetFormatPr 声明，再接受独立 Rust 校验。该步骤不触碰评估页、
活动页或业务数据；额外 I/O 耗时单列为 preserve_defaults_ms，未混入 API Save 耗时。
页眉校验仅把真正空的页眉/页脚 XML 节点与省略节点视为等价；空格、控制码和实际
页脚内容仍保留并检查。其他引擎校验不放宽。
重写引擎按语义校验页眉，不能要求完整 ZIP 部件与 OOXML 定点修补一致。

原始编译/测试、真实菜单捕获、失败注入记录与本轮相对先前工作区的 diff 保存在
`test-runs/aspose-modify-integration-20261115/`。最终结果以 DELIVERY.json 为准。
NOT_TESTED：本轮没有运行全量测试、新 Linux/WSL/Docker/麒麟/WPS 接入、真实业务
人工验收、授权版、去水印或发布包。原 xlwt BIFF 保存后样式拒绝断言保持，未恢复 XLS 追加。

本轮最终：50 项针对性 Rust/协议检查通过，另有 2 项 Windows 真实菜单合成验收通过。
全仓格式检查仍为上述 5 个文件的既有差异；新修改的定点格式检查通过。
