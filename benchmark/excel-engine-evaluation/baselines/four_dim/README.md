# Windows 四维引擎基准

仅独立探针；不加入主运行时、不改变产品路由、不更新 dist。主项目已有库修补和未提交工作全部保留。

## 计时

每个条件 3 轮预热、10 轮正式，所有引擎串行、按轮次轮换顺序。进程内 API 单调时钟是主成绩；另测 Aspose 加载 DLL 的空进程和各 Worker READY。驻留 API 区间已经排除启动，不能再把空跑常数扣一次。

默认 CLR 的早版小文件读取出现明显分层 JIT 过渡，因此另有固定 `DOTNET_TieredCompilation=0`、`DOTNET_TieredPGO=0` 控制组，所有组件均重跑。两组均保留，不能混用各组最快值。控制组也不排除 GC、磁盘 I/O和后台负载。

Read 使用相同 1/10 MiB 档 XLSX 并完整物化 Data；Write 在 API 内生成 50k×20、填充并保存；Calc 先重新载入 2048 个 -9999 错误缓存再计重算；Modify 为同一复杂 XLSX 的 J12=9999＋加粗，以及补充页眉窄操作。计算结果由独立预期验证，不默认任何引擎为真值。

BetterOffice 当前只测 XLSX 的计算，Calamine→模型装载在计时外单列。COM 用全新且 PID 可证明独立的 Excel 实例，只读打开合成输入并最后 Quit；不附加用户会话。ooxml-patch 仅测真实支持的页眉 API，通用值/样式标 N/A。

## 命令

所有 Cargo 先设置 `$env:CARGO_TARGET_DIR='F:\RustBuild\pytools'`。

1. `cargo build --release --offline --manifest-path benchmark/excel-engine-evaluation/baselines/adapters/rust/Cargo.toml --bin four_dim`
2. 用现有 .NET 10 SDK publish `four_dim/dotnet/FourDim.csproj` 到 `F:/RustBuild/pytools/fourdim-dotnet`，中间产物也置 F 盘，不安装 SDK。
3. `fixtures.py <新结果根>/fixtures`，使用已有初版 1/10 MiB 合成样本；不存在时先运行父级 fixtures/generate.py。
4. `FourDim.exe --fixture <新结果根>/fixtures`，生成 XLS/XLSX 错误缓存；independent openpyxl/xlrd 验证 2048 个缓存均为 -9999。
5. `run.py --root <新结果根> --rust .../four_dim.exe --aspose .../FourDim.exe --no-tiered`。读写检查点输出后等待结果根的 `phase2.ready`，用于先给用户第一版表格；写入该标记后继续，不重启 Worker。
6. `structure.py <新结果根> --source <实际修改夹具>`；`chart_audit.py` 另查图表绑定；最后生成总表。

`run.py --fixtures <现有夹具根> --modify-input <实际文件>` 可复用只读合成输入。`results/four-dim-20261114-controlled/experiment.json` 记录本次控制组确切输入和参数。

## 评估与修正

原始评估输出保留；比较视图仅排除确切供应商标记对应的新增评估 Sheet、独占字符串和强制活动页影响。SDK 保存实际包含标记生成成本，不估计授权版。每个完整实验 91 个 Workbook 操作，95 处硬停止，不在达到限制后自动重启。

首版未知 OPC 哨兵挂在包级，只能证明未知包部件是否复制，不代表合法 Workbook 自定义 XML。经 Microsoft WorkbookPart 关系与 Aspose CustomXmlParts 识别核实，已更正关联/属性并重跑修改项。旧输入、错误归因和所有原始输出保留并注明；新生成器默认生成更正版本。

结构审查分开记录“目标修改正确”“所查语义保留”“未修改部件字节不变”。合法 CustomXml 以 Workbook 关系、XML 数据语义和 datastore 属性核对；XML 表示变化另留原始哈希差异。图表同时保存原始公式字符串和解析后的目标范围，不能把绝对写法变化叫作数据丢失。每份输出都逐部件哈希，相同内容可复用已完成的语义解析；docProps/core 不属于所选业务语义，但其原始差异仍逐文件报告。

裸 openpyxl 不等于产品隔离 Worker，后者另有缓存恢复；本轮不借裸库差异误报产品缺陷。所有结果仅对实际格式/数据/对象/平台成立，不据此删除多引擎接口。
