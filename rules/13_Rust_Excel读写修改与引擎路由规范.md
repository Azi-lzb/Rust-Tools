# Rust-Tools Excel 读写、修改与引擎路由规范

> 当前项目范围：仅适用于独立 Rust-Tools 仓库。用户本轮明确要求和根 AGENTS.md 优先；本文件保留具体功能所需的 Excel 安全与能力边界。历史基准和性能数据仅供按需参考，不是全量迁移、全面样式对拍或跨项目同步任务。
> 旧联合工作区规范原文已保存在 docs/archive/Excel规则_旧共享工作区_20261027.md；旧分支、旧路径和 Agent 分工不再生效。

## 1. 项目目标与已确认边界

Rust-Tools 独立维护现有 PyTools 业务和仓库内 Excel-Core。Rust 负责业务核心、读取、规则计算、新建工作簿及 ExcelEditPlan 调度。所有交互菜单不要求外部 ExcelModifier capability 文件；菜单需要修改既有 XLSX 时直接调用随包 openpyxl worker，按功能已有语义生成候选或就地提交，并执行候选校验、源快照及安全提交。读取、Rust 新建工作簿、Office/WPS 转换和公式重算仍由各自功能使用 Calamine、rust_xlsxwriter 或指定引擎。openpyxl worker 不承载业务逻辑、读取统计或公式计算。Windows Excel/WPS COM 与 Linux LibreOffice 仅作为本规范明确限定的公式计算后端；当前 Linux LO 只服务于 1-4 的隔离校验副本重算，COM 不作为通用工作簿修改器。任何 Python/COM 组件均不得承载 PyTools 业务逻辑。原 Python 项目只读，仅在用户需要时作为行为参考。公式计算与工作簿修改分开。面向未来新项目/新功能的通用默认选型见 §4.6；它不触发自动改写既有业务语义。

目标交付为麒麟 V10 x86_64 和统信 UOS V20 x86_64 离线 deb，兼容 glibc 2.28；核心路径不得要求用户预装 Python、Go、Office、WPS、LibreOffice 或 JVM。1-4 的 Linux 校验计算默认使用用户已安装或自行提供的 LibreOffice；该可选依赖不作为核心运行依赖，也不随 deb 打包。程序须检测 LO 是否可用，并在缺失时明确提示如何提供，不得假报计算成功。启用 openpyxl 引擎时可在 deb 内置隔离 Python runtime 与依赖，不得使用系统 Python、用户 PATH 或联网 pip。发布说明须列明 1-4 的 LibreOffice 依赖和检测方式；Windows COM 不进入 Linux deb 或其必要依赖。ABI 和打包约束见根 AGENTS.md。

### 1.1 主路径和旁路

- 核心主路径不依赖 Office、WPS、LibreOffice、COM 或 UNO。
- 公式计算、同类别格式转换及未来 Windows COM 修改只能作为独立、显式选择的可选能力；其不可用不影响读取、新建和其他业务路径。
- 公式计算按平台路由：Windows 由随包 pywin32 worker 调用 Microsoft Excel 或 WPS COM；Linux 的 1-4 使用 LibreOffice Calc。1-4 先由 Rust/OOXML 将校验区域复制到隔离临时工作簿副本，再由 LO Calc 打开副本、计算并保存缓存，最后由 Calamine 读取结果。LO 在此链路中允许保存重算后的临时副本；这属于公式计算，不等于将 LO 用作通用读取、新建、修改或格式转换引擎。当前只有 1-4 可调用 Linux LO；其他功能未经明确任务不得接入。各平台须检查计算完成状态、保存结果并核验关键公式。
- `.xls → .xlsx` 格式转换只由 Windows 随包 pywin32 worker 调用 Microsoft Excel/WPS COM 完成；Linux 不提供该转换，不得改用 LibreOffice 或值重建冒充格式转换。需要 Openpyxl/OOXML 修改 `.xls` 的功能，应先让用户在 Windows 转为 `.xlsx`。其他格式转换仍按各自已验证路由单独登记。
- PDF、纸张渲染、打印机输出和跨类别转换不在当前迁移范围。打印只写页面设置。
- 外部 Worker 是独立进程，主程序启动时不检查外部引擎。用户显式触发时才检查依赖；缺失/失败不能阻塞纯 Rust 主路径，未验证输出不得报告成功。
- 未来 ExcelModifier 可选接入 Microsoft Excel COM、WPS 表格 COM 两个独立 Windows 后端：优先使用 pywin32 / `win32com.client` 与工作簿级 JSON IPC，Excel 使用 `Excel.Application`，WPS 优先验证 `KET.Application`。二者使用独立 EngineId 和 capability 证据，按真实 COM 激活及具体操作验收；不能仅凭软件已安装判定可用。显式选择不可用后端须报错，不静默切换。COM 不提供 Reader/Writer，也不进入 Linux 依赖；不使用 Rust COM、PyO3 或 FFI。只有端到端测量证实 IPC 为瓶颈后才评估替代通信方式。
- LibreOffice Calc 当前仅为 Linux 1-4 的公式计算后端，不是通用读取、新建、修改或格式转换后端。1-4 可在 OOXML 已复制好的隔离副本上调用 LO 计算并保存缓存；复制、业务逻辑和结果提取仍分别由 OOXML/Rust 与 Calamine 完成。Windows 1-4 公式计算由 pywin32 调用 Microsoft Excel/WPS COM。其他功能不得因通用平台矩阵而自动调用 LO。
- COM 批处理应优先评估在一个受控批次内复用应用实例，按工作簿串行处理；COM 初始化、对象创建和释放须在所属 STA 线程内完成，禁止跨线程共享 COM 对象。复用前后应处理工作簿关闭、计算状态、事件/警告设置、文件占用、超时、异常退出和宿主进程归属，避免关闭用户自行打开的 Office。不得未经测量默认每个文件启动一个 Office 实例，也不得仅凭复用推断性能或稳定性改善。
- Windows 计算应使用真正的 Excel 计算调用，等待计算完成并检查计算状态，再保存和验证结果；仅设置自动计算/全量计算标志或执行格式转换不等于已经生成公式缓存。需要重建依赖时才使用完整重建，普通计算不默认承担该成本。Linux 1-4 使用用户已安装或提供的 LibreOffice，以隔离用户配置目录及无界面/UNO 调用，等待进程完成、检查输出并核验关键公式；LO 缺失时应清楚报告依赖不可用。格式转换不等同于强制计算，LO 不作为格式转换引擎。
- PDF、纸张渲染、打印机输出和跨类别转换不在当前迁移范围。打印只写页面设置。旧 COM 业务工作流不得整体搬入 sidecar/worker 以规避 Rust 迁移。
- Linux 1-4 默认由用户提供 LibreOffice，不要求将 LO 打进 deb。发布说明须明确用户侧依赖和运行时检测方式；核心功能仍可在没有 LO 时运行，只有 1-4 校验计算会报告不可用。发布前仍须完成目标 Linux 平台验证。

### 1.2 已确认模板与打印范围

历史已检查模板不含图表、透视表、切片器、图片、工作表保护或 VBA；这不代表其他输入没有这些对象。修改前按当前操作确认实际文件中相关的必要对象。

工作表提取保留普通无批注版，带批注的 COM 分支已明确排除。PyCode 与 NoComPyTools 的其他差异保留为历史记录，仅按用户具体需求核实。

打印只写 orientation、fitToPage、fitToWidth、fitToHeight、清除 scale、paperSize、printArea、horizontalCentered、verticalCentered 和页脚页码。旧代码没有写 margins、printTitleRows/Columns、gridLines、headings、blackAndWhite、draft、firstPageNumber；新增属性须另立需求，不得顺手加入。

## 2. 项目结构与依赖边界

本仓库的业务逻辑位于 `pytools-core/`，Excel 原语位于本仓库 `crates/`。Windows Office/WPS COM Worker 位于 `pytools-core/python/pywin32_excel_conversion_worker.py`，由 Rust 的 `pytools-core/src/conversion_engine.rs` 管理；Linux LibreOffice 仅作为公式计算程序，不执行 Office 格式转换。它们是 Rust-Tools 自己维护的实现，不依赖旧 Main、其他项目或旧共享 workspace，不通过合并、cherry-pick 或自动同步取得更新。

- 通用第三方依赖由本项目根 Cargo.toml 管理；Excel 原语不得承载 PyTools 业务规则。
- 既有工作簿修改以引擎无关的 ExcelEditPlan/ExcelModifier 表达，按具体操作及保留项选择已验证路径，避免反复打开和保存。
- sidecar 是独立的可选进程；Linux 核心路径不得要求预装外部运行时。业务 fixtures、probe、报告和本地配置由本项目独立保留。

## 3. 当前功能的操作契约

一个端到端功能可能是 R01 → 规则计算 → W01。仅针对本次指定功能及其受影响阶段，按需记录下列字段；不要求为其他功能建立完整盘点：

| 字段 | 要求 |
|---|---|
| feature_id / operation | 业务功能 ID 与第 5 节操作编号；多阶段分别列出。 |
| platform | OS/版本、架构、glibc、库/引擎版本；开发和目标环境分开记录。 |
| input_format / output_format | 按文件内容识别，扩展名只初筛；不靠改后缀转换。新建默认 xlsx，编辑默认原格式，格式变化须登记为转换。 |
| output_kind | new_report、edit_existing、format_convert、filesystem；不得把重建报表称为原位修改。render 当前不支持。 |
| write_pattern | append_only、random_access、edit_existing；“读后重建 sheet”不能标成追加。 |
| data_owner | rust、python_legacy、none；迁移期用来记录数据与业务计算所在侧，并测量跨语言传输成本。最终生产业务计算不依赖 Python；仅 ExcelModifier 的可选 openpyxl worker 可在隔离进程运行。 |
| value_view / formula_mode | stored 默认；displayed 需静态显示格式能力验收。公式模式为 cached、text、text_and_cache、recalculate；旧缓存不等于本轮重算。 |
| preserve / drop_allowed | 明确要保留和明确允许丢弃的对象；默认不丢弃未授权对象。 |
| allow_conversion / external_apps | 默认 false；格式转换、公式计算、Windows COM 修改分别显式授权，不因某项开启而自动启用其他外部能力。 |
| reference_app / engine_mode | none、Microsoft Excel、WPS、LibreOffice；Excel/WPS 作为独立参考结果，不假设一致。Windows COM 脚本可为读取/新建测试生成参考数据或文件，但不进入 Reader/Writer 生产路由。外部 COM 功能仍按其运行时要求明确报告不可用；交互菜单的 Openpyxl 修改不要求外部 capability 文件。 |
| workload_limits | 记录覆盖的文件/sheet/行列/对象数量、耗时和内存；没有用户阈值先测，不编造硬限制。 |

修改已有工作簿时，按本次实际操作和文件特征保护值/类型、公式、合并、工作表、批注、名称、VBA 及其他约定必须保留的对象和结构。对可能受影响但未检查的对象注明未验证，不假设其不存在；输入特征变化时应重新判断能力是否匹配。普通字体、边框等外观差异不自动阻塞交付。

## 4. 技术分工

### 4.1 读取

- Rust-Tools 当前生产普通值读取使用 Rust reader；未来 Rust 功能默认选择 calamine，按数据归属、文件格式、输出契约和对象验收。禁止将 python-calamine、PyO3 或本仓库现有 openpyxl 编辑 worker 用作普通值读取路由。若未来功能确需复杂属性读取，可按 §4.6 单独评估 openpyxl 读取能力；它必须保持 Excel 专项职责，业务计算仍由 Rust 承担，并单独满足隔离运行、部署和能力验收要求。
- 当 `data_owner=rust`、输入为 XLSX、操作只读取普通存储值，且坐标、类型、空值和值语义与 §6.2、§6.3 登记的基准契约一致时，Calamine 是优先评估的 reader 候选；这不是该业务功能已经验收的证明。其他格式或不同契约按各自能力登记和对拍。
- Direct OOXML 是直接操作 ZIP/XML 的实现方式，不是固定库，也不是普通值读取在 Calamine 之后的通用 fallback。Calamine 读取失败时必须按实际错误报告，不得再用 Direct OOXML 掩盖或改写错误结果。
- 功能明确需要 Calamine 未提供的 OOXML 专属元数据、模板结构或特定 OOXML 部件时，单独选择已针对该对象和契约验收的专用 OOXML reader/适配器；适配器能力不满足时报告具体缺项，不把普通值读取失败转换成元数据补读。
- Python 拥有数据的迁移期流程保留该项目现有、同契约验收通过的路由。不能由本次 Rust/Python 实现比较推导 Python `python-calamine` 与 Python Direct OOXML 的 reader 排序；候选调整须对该流程的完整调用链同机测量，并计入进程、导入、序列化、传输、数据转换和校验。
- 对 xls、xlsx、xlsm、xlsb 的值、sheet 名、坐标、公式/缓存、合并等字段按格式和锁定版本逐项验收。不能从扩展名或 ZIP 容器推断完整支持。
- Calamine 不提供且业务确需的批注、条件格式、数据验证等字段，用针对这些对象并已验收的 namespace-aware OOXML reader 定点补读；umya 可用于特定字段对拍/补充，但不默认引入第二套完整 reader。
- “API 未提供”“没有此字段”“空值”“没有合并/批注”必须区分。
- CSV 用经对拍的 Rust CSV 实现，固定编码、分隔符、引号、空字段和类型解释。
- 多文件处理用 rayon 做文件级并行，不并行写同一 workbook。

### 4.2 新建工作簿

- Rust 功能从零生成 XLSX 默认使用 rust_xlsxwriter；按业务输出契约检查值、类型、公式文本、必要结构及资源需求。它不是既有工作簿编辑器，存量生产路由不因默认选型自动变更。不要为未指定功能自动开展对拍或性能测试。
- 按本次业务需要核验值/类型、公式、必要结构和文件可打开性；格式或性能比较仅在用户要求或当前问题确有必要时针对相关范围进行。
- 基准报告中的 Direct OOXML 写入是 Python S5/S2 专用 writer；报告没有比较 rust_xlsxwriter 与 Rust Direct OOXML writer。不得把该实现称为 Rust writer，也不得仅凭名称、压缩设置或 ZIP 行写方式推断不同 writer 等价。
- Direct OOXML 仅在存在专用输出契约、明确 OOXML 对象要求，或同契约实测证明资源收益时进入候选；constant-memory 模式、既有工作簿编辑和通用 OOXML writer 均须单独实现与验收。

### 4.3 修改已有工作簿

Rust-Tools 当前修改业务由 ExcelEditPlan 表达。既有 OOXML `.xlsx` 修改统一优先走内置 Openpyxl worker，并用经过验证的 OOXML 定点补丁补齐 Openpyxl 不适合处理的明确部件。不要为方便保留 Umya 修改路径；只有 Openpyxl 与窄范围 OOXML 均无法完成、且同类业务对拍和公式/合并/必要样式保留检查通过时，才可保留 Umya，并须记录具体原因。非 OOXML `.xls` 不能直接交给 Openpyxl/OOXML 修改；除已实现的 Windows 5-1、5-7 功能专用 COM 路由外，需提示用户先在 Windows 通过 4-4 使用 pywin32 驱动 Excel/WPS COM 另存为 `.xlsx`。以上不豁免候选验证和安全提交。

- **Rust OOXML Patch**：适合明确部件、明确位置、依赖关系可证明的窄修改；仅对已验收白名单执行，无法证明安全时拒绝。
- **Rust umya-spreadsheet**：非默认修改引擎。仅当 Openpyxl 与有边界的 OOXML 均不能满足具体保留契约，且 PyCode 业务对拍、公式/合并/必要样式验证通过后，才允许在该单项功能中保留；能由 Openpyxl + OOXML 完成时不使用 Umya。
- **Go Excelize worker**：条件性候选。仅当 umya/定点补丁对真实模板有大面积失败或维护成本过高，且 Excelize 对同一整批操作验证适用时才评估。Rust 以一个批量 IPC 请求传文件路径和完整计划，不逐单元格跨进程。若采用，随包交付，CGO_ENABLED=0。
- **隔离 openpyxl worker**：既有 OOXML XLSX 修改的默认便捷实现。Rust 提交结构化 ExcelEditPlan，再启动随包 Python/openpyxl；不要求外部 capability 文件。worker 仅执行已接入的工作簿操作，不做普通读取、新建报表或公式重算。openpyxl 会重写 XLSX package，因此候选仍须验证数据、结构、未修改区域和计划要求保留的对象，再根据操作模式安全提交。运行时缺失、格式不支持、操作不支持、崩溃、超时或验证不一致时应定位并报告，不得静默切换到另一个修改引擎。
- **功能专用 Microsoft Excel/WPS COM worker**：用户已明确授权 Windows 上为 5-1 至 5-7 的既有业务功能增加与 PyCode 对应的 COM 操作，尤其用于 PyCode 按 `.xls` 路由到 COM 的场景。由 Rust 负责配置解析、业务规则和结果统计；worker 可执行对应功能所需的 Excel/WPS 操作并返回结构化结果。5-1、5-7 的 `.xls` 路径已接入并在限定样本上以 Excel `Excel.Application`、WPS `ket.Application` 完成 PyCode 对拍；5-2 至 5-6 仍需逐项实现和验收。每个菜单功能、Excel/WPS 应用、输入格式、公式/样式/合并等保留契约分别实现和验收；尽量批次复用 COM 会话，源文件修改遵守该功能的授权、备份和候选提交规则。这是明确限定的业务兼容例外，不构成通用 Excel 修改器或任意公式计算 API。Linux 不提供 COM，相关 `.xls` 业务操作不支持，需在 Windows 处理或转换。
- **其他 COM 能力**：仍只按已接入功能使用，例如 1-4 公式重算、4-4 Office 格式转换及 4-5 的专用 `.xls` Sheet 改名。新功能不得将 COM 默认为通用回退；须有明确范围授权并独立验收。Worker 通过隔离 Python runtime 中的 pywin32 / `win32com.client` 工作，不依赖系统 Python/PATH。
- 生产引擎按“具体功能操作 + 输入格式/对象 + 保留契约”选择。Openpyxl 是优先考虑的便捷实现；Umya 只有在相同或代表性输入上完成 PyCode 业务对拍，且所需公式、合并区域和样式通过保留检查后才可用于该操作。两者都不满足时，评估有边界的 OOXML 补丁或明确报告尚不支持，不以静默回退掩盖差异。一个操作通过验收不代表该引擎对其他操作、格式或工作簿对象也通过。
- umya 已覆盖且未证明 Go 有保真、性能或维护收益时，不引入 Go。

### 4.4 业务计算、公式计算与禁止项

规则、期间、校验、统计、去重、映射和审核结果由 Rust 业务逻辑计算；财务数值、Decimal、舍入、日期、空值、错误和文本前导零策略按当前功能契约明确并验证。业务计算、公式文本、旧缓存读取、Excel 公式重算是不同能力，不能互相冒充。

Rust 业务规则、统计和结果判定始终由 Rust 承担。普通读取用 Calamine，OOXML 字段由局部补读；新建写入用 rust_xlsxwriter + 必要 OOXML；既有 XLSX 修改用 Openpyxl + 必要 OOXML。Umya 仅按 §4.3 的例外条件保留。这不把业务逻辑迁入 Python，也不把 worker 作为通用公式计算 API。pywin32 / `win32com.client` 只通过独立 Windows worker 提供指定的 COM 转换或真实公式计算；Linux LO 当前只接入 1-4 的隔离副本重算与保存，不作为通用修改后端。worker 不依赖系统 Python/PATH；修改按工作簿批量处理，不逐格启动。对拍/测试 Python 仍可在 tests/compare 使用。禁止 python-calamine、PyO3、pandas、xlrd、pyxlsb 进入 Rust-Tools 生产依赖。不得无证据重写已验收模块；不得以完整 OOXML writer 为默认目标；不得无测量串联 Rust 读取→中间文件→Go 修改→Go 写回。

`ExcelCalculator` 与 ExcelModifier 分离。当前需要公式重算的 1-4 在 Windows 通过 pywin32 调用 Microsoft Excel/WPS COM，在 Linux 通过 LibreOffice Calc；计算仅在用户操作确实需要时运行，并在隔离副本上计算、保存和验证。1-4 允许 LO 打开并保存已由 OOXML 准备好的临时校验工作簿，以生成公式缓存；这不是通用读取、写入、修改或格式转换路由。其他功能未经明确授权不得接入 LO。公式计算尚未接入的功能必须明确报不可用，不能退回普通读取或修改路径伪称已计算。普通重算和完整重算须按需求区分，并验证公式结果及保存影响；具体公式及平台语义仍须按功能验收。

### 4.5 本项目的 Excel-Core 维护

- 按用户明确提出的业务问题维护本地 Excel-Core，保留已有实现和兼容路径，不为其他项目统一而覆盖或重写。
- 普通值读取优先使用 Calamine；Calamine 不提供的必要工作簿字段用窄范围 OOXML 补读。新建 XLSX 使用 rust_xlsxwriter + 必要 OOXML；既有 XLSX 修改优先使用 Openpyxl + 必要 OOXML；Umya 能避免时不使用。
- 保留现有 openpyxl worker、接口、测试和兼容依赖；未经测试，不得移除旧功能。Go/Excelize、COM 或其他扩展不是默认任务。
- 既有区域复制默认保留公式原文；引用平移和公式重算须单独明确。用户未提出相关功能时，不主动开展引擎扩展或全面验收。

### 4.6 未来 Excel 工具的统一默认选型

本节定义 Rust 与 Python 项目今后新增 Excel 功能时的默认候选；“默认”表示优先采用并按目标契约验证，不代表任意文件/格式的能力已通过验收。Rust-Tools 交互菜单的现有功能按 §4.1–§4.5、§6 约定执行；不得为满足本节而改变用户已确认的业务行为或删除既有实现。

| 项目语言 | 操作 | 默认候选 | 补充或边界 |
|---|---|---|---|
| Rust | 常规读取 | `calamine` | 复杂属性按需使用经过验证的 OOXML reader；`openpyxl` 可作为复杂读取辅助，不默认整簿重复读取。 |
| Rust | 新建 XLSX | `rust_xlsxwriter` + 必要的窄范围 OOXML | 业务数据、样式和公式由 Rust 写入；OOXML 仅补充 writer 不覆盖且有明确契约的部件；不用于修改既有工作簿。 |
| Rust | 修改既有 XLSX | `openpyxl` + 必要的窄范围 OOXML | 通过隔离、按工作簿批处理的 worker 调用；验证重存对无关内容及约定保留对象的影响。Umya 仅按 §4.3 的例外条件保留。 |
| Rust | 严格保真定点修改 | 经验证的 OOXML 定点修改 | 仅限关系及依赖可证明、范围明确并有结果校验的操作；OOXML 不是通用整簿引擎。 |
| Rust | 公式计算 | Windows：pywin32 + Microsoft Excel/WPS COM；Linux：LibreOffice Calc | 与读写/业务计算分离；隔离副本计算，检查重算完成状态、公式结果和保存影响。每项业务需单独接入并验证。 |
| Python | 常规读取 | `python-calamine` | 复杂属性由 OOXML 或 `openpyxl` 补充；不默认用 `openpyxl` 读取全部数据。 |
| Python | 从零新建 XLSX | `XlsxWriter` | 适用于新建和批量写入，不用于修改已有工作簿。 |
| Python | 基于模板生成新文件 | 复制模板后用 `openpyxl` 填充 | 先评估模板内容及高级对象的保真风险。 |
| Python | 修改既有 XLSX | `openpyxl` | 按目标对象和格式验证；高保真局部修改可评估 OOXML。 |
| Python | 严格保真定点修改 | 经验证的 OOXML 定点修改 | 同样需要处理 ZIP/XML 关系并验证结果，不能仅凭使用 OOXML 宣称无损。 |
| Python | 公式计算 | Windows：pywin32 + Microsoft Excel/WPS COM；Linux：LibreOffice Calc | 与公式读写、普通新建/修改区分；只在实际需要计算时检查依赖。 |

### 4.7 Rust-Tools 用户确认的平台路由

以下项目路由优先于本规则中的通用候选表，并按用户确认的目标及旧 PyCode 行为对齐：

| 操作与输入 | Windows | Linux |
|---|---|---|
| 普通值读取（Calamine 支持的格式） | Calamine | Calamine |
| OOXML 专属结构/元数据补读 | 窄范围 OOXML reader | 窄范围 OOXML reader |
| 从零新建 XLSX | `rust_xlsxwriter`，必要部件用 OOXML 定点补充 | `rust_xlsxwriter`，必要部件用 OOXML 定点补充 |
| 修改既有 OOXML `.xlsx` | 隔离 Openpyxl worker + 必要的 OOXML 定点补充 | 隔离 Openpyxl worker + 必要的 OOXML 定点补充 |
| 5-1 至 5-7 的专用业务修改 | 按 PyCode 实际入口和输入格式路由：PyCode 对 `.xls` 使用 COM 的功能可实现对应的专用 pywin32 Excel/WPS COM 操作；`.xlsx` 默认走 Openpyxl/OOXML，只有 PyCode 对该功能确有 COM 分支时才对齐该专用分支。每项均须独立实现和验收，不是通用修改器 | 不支持 COM；按具体功能的 Openpyxl/OOXML 能力处理 `.xlsx`，`.xls` 先在 Windows 处理或通过 4-4 转成 `.xlsx` |
| 其他既有 BIFF `.xls` 修改，且功能需要 Openpyxl/OOXML | 先由用户通过 4-4 转成 `.xlsx`，再修改 | 不支持；先在 Windows 转成 `.xlsx` 再使用 |
| 4-5 按配置改 Sheet 名，输入为 BIFF `.xls` | 专用 pywin32 worker 调用 Excel/WPS COM 改名；不是通用 COM 编辑能力 | 不支持；可先在 Windows 通过 4-4 转为 `.xlsx`，再使用 4-5 |
| `.xls → .xlsx` 格式转换 | pywin32 调用 Microsoft Excel/WPS COM，批次复用会话 | 不支持；提示用户在 Windows 转换 |
| `.xlsx` 扩展名但内容不是 OOXML ZIP | OLE/BIFF `.xls` 内容仅改了 `.xlsx` 后缀时，复制到隔离目录并暂存为 `.xls`，由 pywin32 调用 Excel/WPS COM `SaveAs`；COM FileFormat 必须确认是旧 Excel BIFF。其他无法识别的非 ZIP 内容明确拒绝 | 不支持；先在 Windows 正确转换为 OOXML `.xlsx` |
| 1-4 命名校验区域重算 | pywin32 调用 Microsoft Excel/WPS COM，在隔离副本上计算 | LibreOffice Calc，在 OOXML 复制好的隔离校验副本上计算并保存缓存；之后由 Calamine 提取结果 |

约束：

- Calamine 负责普通值读取；OOXML 专属字段由定点补读补充，不重复整簿读取。新建 XLSX 走 `rust_xlsxwriter` + 必要 OOXML；既有 OOXML XLSX 修改走 Openpyxl + 必要 OOXML，均不承载 PyTools 业务规则。
- Windows COM 可用于已实现的 1-4 公式计算、4-4 格式转换、4-5 专用 `.xls` Sheet 改名，以及经本轮明确授权后逐项开发的 5-1 至 5-7 功能专用业务操作；5-x 业务 COM 路径必须对应 PyCode 功能，不能扩成通用修改器，并逐项注明实现和验收状态。LO 当前只按 Linux 1-4 公式计算任务调用；其保存仅作用于已准备好的临时重算副本。两种引擎都不是未授权功能的隐式回退。Windows COM 应优先复用批次会话，并兼容 WPS `KET.Application`。
- Linux 不支持一般 `.xls` 修改、4-5 `.xls` 改名或 `.xls → .xlsx` 转换；需要 Openpyxl/OOXML 修改 `.xls` 时，先让用户在 Windows 通过 4-4 转换。Linux 有公式计算需求时使用用户提供的隔离 LibreOffice 环境；不能把 LO 当成格式转换器。
- 4-4 的输入须检查容器而不能只看后缀；用户要求的格式另存由 Windows COM `SaveAs` 执行，真实类型及输出可读性须验证。当前针对 `.xlsx` 后缀的非 ZIP 输入，只自动识别并处理 OLE/BIFF `.xls` 内容；未知或其他容器不做猜测。
- 当前差距必须如实登记：4-4 的容器识别及 `.xlsx` `SaveAs` 行为须按实际代码核验；Linux 公式计算只在已接入并验证的功能上可用；4-5 `.xls` COM 路径本轮以 Excel.Application 完成样本验收，WPS 仍待单独验收；4-5 的 `.xlsx` 改名目前尚未接入 Linux 通用 Openpyxl worker；其他存量功能需依 §9.1 隔离对拍逐项迁移，不因本表推断为已验收。

统一决策要求：

1. 先确认项目主语言、操作类型、文件格式、模板/保真要求和是否需要公式计算，再选对应引擎；读取、新建、修改、特殊属性处理、公式计算可以分别使用不同引擎。
2. 公式写入、公式文本读取、缓存结果读取、触发重算和保存重算结果是不同操作。`calamine`、`python-calamine`、`rust_xlsxwriter`、`XlsxWriter`、`openpyxl` 均不作为通用公式计算引擎。
3. `.xls`、`.xlsx`、`.xlsm`、`.xlsb`、`.xltx`、`.xltm` 的结构能力不同；数据读取后重写只代表数据迁移，不代表高保真格式转换。不得从扩展名或库的语言/实现方式推断兼容性或性能。
4. 对已有文件明确输出副本/覆盖策略，优先临时文件写出并按业务契约验证后再替换；不得在失败时损坏源文件或静默丢弃未声明对象。
5. 变更引擎或声称性能提升时，以代表性文件对比业务结果、必要结构、耗时和资源占用；未实测不得宣称更快或更保真。替代默认引擎时记录原因、依赖变化、兼容风险及回退条件，并遵循用户授权。
6. 以上默认是未来开发的选型指导。存量功能不要求立即符合；发现差异先记录，只有用户明确要求迁移/修复时才修改实现。

## 5. 操作编号和目标方向

编号是能力分类，不代表当前已经实现。每项须按操作子类、格式、对象和平台登记。本表中的“目标方向/候选顺序”仅指导评估，不等于已通过当前功能、格式、对象和平台验收的生产路由。

| 编号 | 操作 | 目标方向 |
|---|---|---|
| F01/F02 | 文件操作；按单元格内容命名 | 文件系统 API；读取另走 R01/R04，不重存 workbook。 |
| R01 | 读值、sheet 名 | 仅对 Rust 持有数据、普通存储值 XLSX 且坐标/类型/空值/值语义与专项契约相同的功能，优先评估 Calamine；其他数据归属、格式或契约保留各自已验收路由。Calamine 失败不以 Direct OOXML 通用回退。 |
| R02/R03 | 坐标、合并；公式/缓存、批注、样式、名称、打印元数据 | Calamine 已验收字段；OOXML 专属元数据使用对象/契约已验收的专用 reader，不作为普通值读取 fallback。 |
| R04/R05 | 显示文本或高级格式元数据 | 仅使用对应已验收静态解释器/reader；否则报告未验证或不支持。 |
| W01 | 新建 XLSX | Rust 持有结果、从零生成纯数据 XLSX 时优先评估 rust_xlsxwriter；对应输出契约通过且峰值内存满足预先定义的功能预算后，才可用于该功能。其他 writer 需有专用契约、OOXML 对象要求或同契约实测的资源收益。 |
| W02 | 修改已有文件的值、公式、样式、批注、打印属性 | ExcelEditPlan + ExcelModifier；现有已验证引擎按整批保留契约选择；未来 Microsoft Excel/WPS COM 仅为 Windows 独立候选。 |
| W03/W04 | 追加；插删行列、移动区域、sheet 改名/重排/删除 | 修改操作；检查其他 sheet、公式、名称、表格、打印区域和外链依赖。重建 sheet 不算保真追加。 |
| W05–W07 | 跨簿按值提取、复制指定样式、整表保真复制 | 每项定义 preserve；按值复制不能称为格式/整表复制。 |
| P01/P02 | 写打印页面设置；PDF/渲染 | P01 只写页面设置；P02 当前不支持。 |
| Q01/Q02 | 读旧公式缓存；按本轮输入普通/完整重算 | Q01 标记“文件已保存结果”；Q02 由独立 ExcelCalculator 默认关闭，显式选择已验收计算后端并验证缓存、公式、依赖和保存保真。 |
| C01/C02 | 按值重建为 xlsx；同类格式转换 | C01 必须明示丢弃范围；C02 显式转换并校验。 |
| D01 | Word 同类格式转换 | 仅已登记方向、显式 sidecar 或验收通过的 Rust/Go 转换器；PDF/跨类不支持。 |
| T01/T02 | CSV 读写、分组/连接/透视 | Rust CSV 与业务计算；生成 XLSX 另执行 W01。 |

## 6. 能力登记与历史实现快照

### 6.1 状态

以下能力状态是引擎技术证据，不是业务功能状态；具体问题按需使用。能力按“操作子类 + 格式 + 对象特征 + 保留项 + 平台 + 版本范围”登记，状态仅为 **passed / unverified / unsupported**（通过/未验收/不支持）。代码存在、库文档、mock 或能打开简单文件，都不等于 passed。对拍对象逐项记“保留 / 丢失 / 损坏 / 不适用 / 未验证”；未测量写“未测量”。

登记字段至少包括 operation/feature_id、engine/version/适配器入口、输入输出格式、write_pattern、preserve/drop_allowed、platform、样本、报告、性能范围、默认值和状态。可维护在现有架构文档、静态表或 JSON，不新增插件框架。业务配置、夹具和报告分开。

候选顺序与生产路由必须分开登记：安装了库、基准更快、或出现在候选顺序中都不等于当前功能/格式/对象/平台已经 `passed`。实际生产路由仅能来自具体功能的验收证据和现有默认记录；通用基准不得自动改写默认值。

| 当前组件 | 现状证据 | 可声明范围 |
|---|---|---|
| excel-reader / calamine | fastpath_read_compare 对 synthetic_probe.xlsx：sheet names 1/1、values 65/65、formulas 1/1、merges 1/1、comments 1/1；批注由 OOXML 补读 | 仅上述合成样本/字段可 passed；真实模板及 XLS/XLSM/XLSB/CSV 矩阵仍 unverified。 |
| excel-writer / rust_xlsxwriter | 当前为库类型导出；新建合成工作簿 9/9 | 仅样本断言 passed；业务写入 API、低内存能力和真实版式仍 unverified。 |
| excel-editor / umya | 薄封装；原始合成修改报告 17/18，列定义保留失败；结合窄列补丁后的样本 18/18，补丁范围 3/3 | 这些结果只覆盖该合成样本和列定义操作；pytools 机构名称规范化也仅为合成候选。通用 W02 unverified，无统一路由/事务边界。 |
| ooxml-patch | 有未验证部件拒绝及窄列定义/列宽补丁 | 仅报告内定点操作可 passed；通用单元格、批注、公式、页面设置仍未验收。 |
| probe | 读取/JSON 导出和部分比较 | 修改计划与全引擎路由探针未形成。 |
| Excelize worker | 当前没有 Go module/worker/IPC | 当前 unsupported（未实现）；达到 §4.3 门槛后再立项。 |
| openpyxl worker | `crates/excel-editor/python/openpyxl_worker.py` + Rust `OpenpyxlXlsxBackend`；交互菜单不要求外部 ExcelModifier capability 文件；支持已接入的 XLSX EditPlan 批处理 | 已有合成 worker/进程集成测试；菜单各功能按对应定点回归记录。运行时或格式不支持仍须明确报错；openpyxl 重写 package 后继续执行数据、结构、快照和保留项校验。该状态不代表全部 XLSX/XLSM 对象均已覆盖。 |
| `pytools-core/src/conversion_engine.rs` / LibreOffice | 当前仅由 Linux 1-4 调用；Rust/OOXML 先复制校验区域，LO Calc 在批次复用 UNO worker 中打开隔离副本、计算并保存，随后 Calamine 提取缓存。LO 不执行其他业务的常规读取、修改、新建或格式转换。 | WSL 合成工作簿上的公式计算、保存及 Calamine 缓存提取测试通过；Linux/deb 打包和真实工作簿仍未验证。其他业务未接入 LO 计算。格式转换禁止走 LO。 |
| `pytools-core/python/pywin32_excel_conversion_worker.py` | 随包 CPython + pywin32 Worker；4-4 Office 格式转换、1-4 Windows 校验区域重算、4-5 `.xls` 工作表改名，以及 5-1/5-7 Windows `.xls` 专用业务 COM 路径已接入；5-2 至 5-6 尚待逐项开发/验收。不得据此提供通用编辑能力。普通 Excel 文件先尝试 `Excel.Application`，Excel 不可用时尝试 `KET.Application`；ET/ETT 优先 KET；Word/WPS Writer 同样走 Worker；COM 批次复用宿主会话。 | 5-1、5-7 在限定 `.xls` 样本上分别使用 Excel 与 WPS 完成 PyCode 结果对拍；工作簿级比较结果、源文件保护和局限见 `pytools-core/MIGRATION_INVENTORY.md`。WPS 会改动交给它打开的文件字节，因此引用文件须先建立私有暂存副本；Linux 没有该 COM 路由，相关 `.xls` 业务操作需在 Windows 完成。其他 5-x 不能引用本次结果作为已验收证据。 |
| ExcelModifier 的 Excel COM / WPS COM 编辑后端 | 未实现通用既有工作簿 COM 修改路由；仅允许按本表列出的功能专用 COM 操作，不开放通用编辑选项 | 不支持 COM；分别按具体功能记录 Linux 路由或限制。 |
| ExcelCalculator | 不是通用工作簿修改 API；Windows 由 pywin32 调用 Excel/WPS COM，Linux 由 LibreOffice 承担公式计算。 | 当前 1-4 已接入两平台重算，并完成 Windows 编译及 Linux WSL 合成端到端验证；Linux/deb 和真实业务样本仍待验收。其他功能未接入时不得宣称可计算，也不得用 LO 转换文件。 |

当前功能状态和已执行验证见 `pytools-core/MIGRATION_INVENTORY.md` 与 `pytools-core/TEST_COVERAGE.md`；性能证据见本文件引用的基准报告。报告数字只适用于其中样本和断言。已经通过的窄范围实现继续保留，不因本规则接入而重构。

### 6.2 限定范围的性能证据（报告时间：2026-10-23）

证据来源：[Direct OOXML 与 Rust Excel 引擎独立基准报告](../../excel-engine-benchmarks/20261022-ooxml-vs-rust/REPORT.md)（SHA-256：`D33A786C90D1942FB66DE33E3203A5D9A136414DA89D92A96B6935E04B84A596`）。测试环境为 Ubuntu/WSL2 x86_64、ext4、Release Rust；Rust 1.99.1、calamine 0.37.0、rust_xlsxwriter 0.101.0；Python 3.12.3、python-calamine 0.9.0、openpyxl 3.1.5。有效组合均为通过正确性检查后的至少五次串行测量。证据限定于报告中的 XLSX 输入、实现版本、规模和明确输出契约，不外推到 Windows、麒麟、统信、XLS 或 XLSB。

本节性能证据只能按 `(data_owner, 操作/output_kind, input_format/output_format, 输出契约, 平台, 引擎版本)` 匹配。它用于候选评估，不登记为任何业务功能已通过，也不替代真实模板、目标平台或完整调用链验收。

#### 读取证据适用范围

- **适用路由**：`data_owner=rust`、`input_format=xlsx`，读取普通存储值；只有当输出坐标、类型、空值和值语义与报告 typed digest 契约一致时，Calamine 才是该具体功能的首选 reader 候选。`pure_same_range` 中 Calamine operation 中位数为 132.100 ms、进程墙钟 138.571 ms、峰值 RSS 中位数 17.3 MiB；`pure_full_sheet` 分别为 134.064 ms、142.904 ms、18.1 MiB。对不同范围、类型解释或返回结构须重新验收。
- Python `ooxml_external_prepare` 对应实现中，same-range 为 2161.393 ms operation / 2354.164 ms process wall / 183.3 MiB RSS；full-sheet 为 2148.653 ms / 2448.143 ms / 183.3 MiB。它是本报告内的具体 Python ZIP/XML 实现，不能代表所有 Direct OOXML 实现。
- 报告**没有**比较 Python `python-calamine` 与 Python Direct OOXML 的同进程端到端性能。`mixed_source_reader` 还包含额外预加载和元数据构造，不得据此推广 Python-owned 流程的 reader 排序；Python-owned 迁移期流程维持原项目同契约验收路由，调整前须测完整调用链。
- 条件格式规则元数据是单独的 `metadata-only` 请求：`ooxml_cf_public` operation 中位数为 2.467 ms、含启动/导入/校验摘要的 process wall 为 88.579 ms；Calamine 没有对等 API，因此不作 reader 排名。批注、条件格式、验证等 OOXML 专属字段应由已针对该对象和契约验收的专用适配器提供。Direct OOXML 普通读取路径拒绝部分含合并、条件格式、数据验证或公式的样本；能力不满足时报告缺项，不作为 Calamine 失败后的普通值 fallback。

#### 新建工作簿证据适用范围

- **适用路由**：`data_owner=rust`、`output_kind=new_report`、从零生成 `output_format=xlsx` 的纯数据报表。基准数据预先生成；被测规模为 10,000、100,000、300,000 行 × 20 列，另有一行表头，并分别包含纯数据/格式化输出。operation 包含数据文件加载/模型适配、writer 和 fsync，不含数据生成；需要评估具体功能时，应以该功能的输出和必要保留契约为准；普通外观差异不自动阻塞交付。
- `rust_xlsxwriter` 使用默认内存工作表模式，报告没有测试 constant-memory。30 万行 × 20 列 same-level 纯数据样本：Python `ooxml_s5` 为 17.406 s / 288.2 MiB RSS，Rust `rust_xlsxwriter` 为 8.729 s / 1376.7 MiB；same-level 格式化样本分别为 19.801 s / 354.0 MiB 与 7.129 s / 1376.7 MiB。Rust writer 的 1,376.7 MiB 约为 1.34 GiB，是该规模/模式的实测观测，不是通用阈值或任意大文件的内存承诺。delivery-default 是另一配置组，须单独比较，不能与 same-level 数据混算。
- 报告中的 `ooxml_s5` / `ooxml_s2` 是 Python 专用 writer，不是 Rust Direct OOXML writer。报告没有比较 `rust_xlsxwriter` 与 Rust Direct OOXML writer，也没有测试通用 OOXML writer、既有工作簿编辑、任意精度 Decimal、图表/宏或 constant-memory 模式。Direct OOXML 只在存在已定义的专用输出契约、明确 OOXML 对象要求，或同契约实测证明资源收益时进入候选；不同实现不能仅凭名称、压缩配置或 ZIP 写入方式视为等价。

#### 环境边界与路由决定

- 环境为 Ubuntu/WSL2 x86_64、ext4、Release Rust；报告记录 Rust 1.99.1、calamine 0.37.0、rust_xlsxwriter 0.101.0，以及 Python 3.12.3、python-calamine 0.9.0、openpyxl 3.1.5。该证据只覆盖报告中的 XLSX 样本和契约，不外推到 Windows、麒麟、统信、XLS、XLSB 或其他未测平台/格式。
- 候选顺序必须与已验收路由分开登记。库已安装、基准更快或候选顺序靠前都不等于具体业务功能、格式、对象、平台和输出契约已通过；本报告不改写现有生产默认。若为性能原因评估修改既有默认，可先定义该功能的时间、内存和测试规模，再用同机、同输入、同契约的完整调用链测量；针对功能缺陷的修复无需先开展性能项目。迁移期间如存在 Python/Rust 桥接，计入进程启动/导入、序列化、传输、数据转换、IPC 和校验成本。

### 6.3 第二份独立性能证据（报告时间：2026-10-23）

证据来源：历史 Open-Data-Audit Direct OOXML 与 Rust 引擎基准报告（原文件不属于本仓库，不作为当前运行路径）（SHA-256：`9BE7045378F1CDF2B45288D478933560DF915553CF3AD08A2D4490876552128A`；逐轮数据见同目录 `benchmark_runs.csv`，机器结果及环境见 `benchmark_results.json`）。这是独立于 §6.2 的第二轮测量：Ubuntu/WSL2 x86_64、Release、每项 5 次串行；Rust calamine 0.37.0、rust_xlsxwriter 0.100.0，Python 3.12.3、openpyxl 3.1.5。任务说明列出 `python-calamine 0.9.0`，但同目录 `benchmark_results.json` 的 `environment.python_calamine` 和基准 `README.md` 均记录为 0.3.1；本节暂按可追溯的机器结果字段登记为 0.3.1，实际运行时版本仍待基准环境复核，不把版本冲突外推为性能结论。`rust_xlsxwriter` 与 §6.2 的 0.101.0 不同，两份报告不合并、不平均，也不据此解释或修订彼此数字。


本节证据必须按 `(data_owner, 操作/output_kind, input_format/output_format, 输出契约, 平台, 引擎版本)` 匹配；下列数字只是对应样本和实现的观测，不是通用阈值、默认路由验收或所有 XLSX 文件的性能承诺。读取计时是报告所列 reader 对同一 XLSX 输入的测量，不含 Rust 写入路径中的 Python JSON 数据桥接；所有 Rust writer 的端到端数字均明确包含报告所述 Python JSON 编码、子进程启动、Rust JSON 解析及 XLSX 写入阶段，不能用 Rust writer 内部时间替代。

#### 读取证据适用范围

- **适用路由匹配键**：`data_owner=rust`、`operation=R01 ordinary_value_read`、`input_format=xlsx`、`output_format=JSON coordinate/value digest`、输出坐标/值摘要且范围与报告一致、Ubuntu/WSL2 x86_64、Release、calamine 0.37.0；对比实现是本报告的 Python Direct OOXML value provider，不代表所有 Direct OOXML reader。`synthetic_full_sheet` 为 50,000 行 × 20 列数据（含表头的报告 bounds 为 50,001 × 20）：Calamine 中位 1.043 s / 73.7 MB 峰值 RSS，Direct OOXML 中位 12.524 s / 770.1 MB 峰值 RSS；坐标/值摘要和 bounds 一致。**这些读取数字不含 Python JSON 数据桥接端到端时间**，且仅限该 reader 计时口径和合成 XLSX。（来源：报告“读取结果”与“读取一致性”。）
- `real_s2_xlsx_full_sheet` 的实际 S2 XLSX 全表读值分别为 Calamine 0.012 s / 4.0 MB 与 Direct OOXML 0.210 s / 29.8 MB；其 bounds 为 729 × 16，坐标/值摘要一致。匹配键仍限 `data_owner=rust`、`operation=R01 ordinary_value_read`、`input_format=xlsx`、`output_format=JSON coordinate/value digest`、同一普通值全表读取契约、Ubuntu/WSL2 x86_64、Release 和上述 reader 版本，不作为业务功能完整调用链证据。此项同样是 reader 计时，不含 Python JSON 数据桥接。（来源：报告“读取结果”与“读取一致性”。）
- `type_error_edge_probe` 的差异按 §9.1 处理，不因普通数据样本摘要一致而视为类型语义等价；格式/类型契约不同的读取路径必须单独验收。

#### 新建工作簿证据适用范围

- 报告的新建写入矩阵覆盖 `data_values_base_style` 与 `common_format` 两类 10,000、100,000、300,000 行 × 20 列样本；本节摘录 30 万行指标作为大样本观测，并单列真实 S2 三表导出。其余规模的逐项数据见报告“写入结果”，不能把本节摘录视为全矩阵结果。
- **适用路由匹配键**：`data_owner=Python-side payload via JSON bridge`、`operation=W01`、`output_kind=new_report`、`input_format=pre-generated in-memory data`、`output_format=xlsx`、报告对应的数据值或共同格式契约、Ubuntu/WSL2 x86_64、Release；Rust 候选为 rust_xlsxwriter 0.100.0 默认内存模式，Direct 对照为报告中的 Python S1 writer。该键标明迁移桥接边界，不等同于 Rust-owned 产品数据已验收。
- 30 万行 × 20 列 `data_values_base_style`：rust_xlsxwriter 两组 Rust 写入数字均为**含 Python JSON 桥接端到端口径**；ZIP level 1 中位时间为 16.631 s / 1,588.2 MB 峰值进程树 RSS，Python Direct 为 20.356 s / 376.6 MB；`common_format` 的 Rust 数字同为含 Python JSON 桥接端到端口径，为 14.722 s / 1,594.0 MB，Python Direct 为 19.121 s / 379.8 MB。Rust 端到端口径包含 Python JSON 编码、子进程启动、Rust JSON 解析及写入；Direct 数字是该 Python 专用 S1 writer 的端到端结果。内存数字只是这两个 30 万 × 20 样本的观测，不是通用上限或任意大文件承诺；报告未测 constant-memory。（来源：报告“写入结果”中 `data_values_base_style_300000x20` 与 `common_format_300000x20` 项。）
- **适用路由匹配键**：`data_owner=Python S2 report payload`、`operation=actual_s2_three_sheet_report`、`input_format=in-memory payload transported by JSON bridge`、三表新建 `output_kind=new_report` / `output_format=xlsx`、同报告 payload/样式/值摘要、Ubuntu/WSL2 x86_64、Release、Rust rust_xlsxwriter 0.100.0。首次结果中 Rust ZIP level 1 的 0.040 s 为**含 Python JSON 桥接端到端口径**，Python Direct S2 ZIP level 6 为 0.758 s；但 Direct 的该次总时间包含每个新 worker 首次导入 exporter，且压缩级别不同，故不可据此作稳定运行时间胜负结论。报告热路径复测为 Direct 0.076 s（0.071–0.082）、Rust 0.050 s（0.049–0.054）；该 Rust 数字同样为**含 Python JSON 桥接端到端口径**。Rust 输出仍有 17,580 格中的 23 格浮点值相差 1 ULP，严格摘要不等。（来源：报告“写入结果”及“补充复测：实际 S2 写入热路径”；1 ULP 差异见附录“实际 S2 写入：浮点末位（1 ULP）序列化差异”。）
- S2 三表输出的业务/值保留契约尚未通过严格浮点对拍；S3 是修改既有工作簿并保留模板部件，未由本报告与新建 rust_xlsxwriter 作同类写入比较。保留 S2/S3 专用路径，不由新建 writer 的性能数据推导替换结论。

#### 候选评估注记与范围

- 本轮读值数据使 Calamine 值得按 §9.2、§9.3 的完整调用链、输出契约及资源门槛继续评估；它是候选注记，不改任何既有读取默认。
- 新建写入暂不因本报告切换默认：30 万行样本中的 Rust 峰值 RSS 明显高于 Python Direct，且 S2 浮点严格值摘要未通过。S2/S3 专用实现按其不同输出/模板契约保留；任何替换仍须先定义节省时间、加速比例、内存上限及规模，并按 §9.2、§9.3 完成同机同输入完整流程对拍。
- 本证据限 Ubuntu/WSL2 x86_64、报告列出的 Release 版本、XLSX、样本规模及契约；不外推到 Windows、麒麟、统信、XLS、XLSB、ET、constant-memory 或既有工作簿通用编辑。XLS 侧测只用于 §7 的伪 `.xls` 预期拒绝和 §9.1 的类型/空值口径观察，不纳入 Direct OOXML 与 Calamine 的读取排名。

## 7. 路由、临时输出与安全回退

1. 解析契约、真实格式/对象、公式口径与外部引擎要求。
2. 按 §5 匹配操作及子类。
3. 只把适配器已实现且 passed、格式/平台/版本/保留项/资源范围匹配的后端加入候选。
4. 手动模式只执行所选合格后端；不可用或失败即停止。自动模式仅从合格候选中选择。
5. 执行到临时文件，按契约验证后才提交最终路径。

读取路由特例：R01 普通存储值读取遇到 Calamine 错误时，不得转为 Direct OOXML 通用重试；OOXML 专属字段是单独请求的元数据能力，只能由对应对象已验收的专用适配器提供。除该限制外，自动回退只允许执行前发现能力/平台不匹配，或引擎在写出前明确报告不支持该文件特性，且有另一完全合格后端；回退必须从原始输入或干净副本重跑并记录原因。输入损坏、加密、权限/占用、空间不足、输出验证差异、公式错误、缓存未刷新、超时或未知异常均停止，不换引擎掩盖问题。不得在部分写入或疑似损坏的中间文件上继续执行。覆盖原文件须有明确授权与备份/哈希保护。日志记录引擎、版本、耗时、路由/回退原因和未满足项。

格式签名无效的伪 `.xls`（包括实测到的 HTML 伪装导出，CFB/OLE 签名无效）必须在路由前以“格式不支持”显式拒绝并记录为预期拒绝；不得计为引擎失败，也不得触发备用引擎回退。依据：第二份独立基准报告附录“real_xls_side_test”复核。

## 8. OOXML 安全边界

OOXML 是补丁/缺失元数据层，不是主力完整读写库。必须使用 namespace-aware XML parser；禁止正则替换 XML 全文。

- 由 workbook.xml 和 .rels 解析真实 sheet 身份到 worksheet part 的关系；不假设 sheet1.xml。
- 维护操作涉及的 Relationship、[Content_Types].xml、sheet 身份/顺序、sharedStrings、styles、defined names 和公式引用。
- 只对明确、已验收子集修改；区分单元格类型、共享/内联字符串、样式 ID、缺失单元格、共享/数组公式及缓存。
- 批注可能关联 comments、VML、threaded comments、relationships 和 Content Types；未经专项验收不得改写或新增。
- sheet/行列/合并/名称/超链接/页面设置等变更须检查跨部件引用。宏、签名、外链、加密、未知扩展或非标准变体不因“未主动修改”就自动安全；超出白名单则拒绝。
- 按本次操作和必要保留契约检查相关 ZIP entry 的解压后语义，未检查的范围注明未验证；整个 ZIP 字节哈希可因重压缩而变化。
- 每个适配器须覆盖该操作涉及的全部保留项；未验证部件直接拒绝。先做小能力和拒绝规则，不实现完整编辑器。

## 9. 对拍、性能与默认切换

### 9.1 按需功能验证

本项目进行用户要求的 PyCode 对齐或引擎调整时，遵循以下顺序，先建立基线再改实现：

1. **圈定范围**：列出本轮 Rust-Tools 实际存在且可达的功能及菜单入口，按用户指定范围执行；忽略返回/退出项，不把 PyCode 有而 Rust-Tools 没有的功能扩展成迁移任务。PyNoComPyTools 不作为对照基线，除非用户另行指定。Rust 独有功能或 PyCode 无对应实现的功能，标为“无直接基线”，按现有业务契约测试。
2. **准备基线**：优先使用用户提供的最新配置、测试用例和已有验收数据。对会覆盖、删除、改名、归档或追加的输入，先在独立测试目录备份并核对副本；PyCode 与 Rust-Tools 使用内容相同但互相隔离的副本及输出目录，禁止用生产原件直接对拍。为每次运行留存实际选中的输入绝对路径及 SHA-256、配置文件路径及 SHA-256；仅凭同名文件、准备目录的清单或用户选择提示，不能证明运行时读到的是那份副本。
3. **先跑两侧现状**：在修改 Rust 实现前，分别运行 PyCode 和当前 Rust-Tools，捕获实际命令、版本、配置、入口、引擎、终端输出、退出状态和产物，并将实际输入路径/哈希清单关联到该次运行记录。没有执行的侧或环境不可用的场景标为“未验证”，不得用模拟输出冒充真实结果。
   **按实际入口、引擎和扩展名分别记录**：旧 PyCode 的自动路由入口会将 `.xls` 分到 COM 分支、将 `.xlsx`/`.xlsm` 分到非 COM 分支（相关功能通常由 Openpyxl 处理）；但部分功能同时保留普通入口和自动路由入口，不能仅凭扩展名推断某次运行实际调用了哪个引擎。逐功能核实入口及运行记录，并把真实执行的 COM 结果作为独立业务参照，和 Openpyxl 结果分开保存、比较；不要求没有对应入口的功能人为补跑另一引擎。4-4 Office 格式转换是例外，`.xls` 和 `.xlsx` 都通过 COM `SaveAs`。不得把一种引擎或扩展名的结果冒充另一种；若 Rust-Tools 对某种扩展名没有同类路由，明确记为“未对齐/未验证”，不要以另一格式的结果代替。
4. **按契约比较**：比较该功能实际承诺的值、类型、行列/排序、公式文本及可用缓存、工作表名称/顺序、合并区域、样式和输出文件；只对相关对象做检查。日期、浮点、空值等先按明确业务语义归一，避免类型表示差异造成误报。分别记录 Rust 独有、PyCode 独有、相同、失败和未验证项。
5. **再定点对齐**：以 PyCode 作为业务行为参考，修复 Rust-Tools 的差异后，用相同副本、配置和场景重跑两侧，确认差异收敛且源文件安全。若 PyCode 的行为会破坏本项目明确要求保护的公式、合并区域、样式或业务数据，不照搬破坏行为；记录该差异并保持保护，必要时交由用户决定。
6. **留存证据**：保存测试覆盖矩阵、真实输出/关键摘要、输入备份位置、产物比较结果、引擎路由、已知差异及未运行项目。报告要区分“实现/静态检查”“合成测试”“用户样本对拍”和“人工验收”；不能将菜单可达、编译成功或已有测试用例等同于功能对齐。

平常的定点问题仍只验证用户指定功能及必要依赖；不因此自动扩大为全项目审计。普通字体、边框等细微外观差异不是普遍阻塞项，但不得把公式丢失、合并区域丢失、样式整体损坏或业务数据变化归类为普通外观差异。未运行的验证如实记录；Rust 编译、运行和测试须按根 AGENTS.md 获得明确授权。

- **跨引擎浮点值不得要求位精确相等**：§6.3 的真实 S2 对拍中，rust_xlsxwriter 与 Python `repr` 序列化存在 1 ULP 差异（约 17,580 格中 23 格，集中于一列）。跨引擎/跨实现的浮点值按相对容差不大于 `1e-15` 比较，或先按业务约定归一为十进制字符串再比较；容差与十进制精度须写入该功能契约。该规则适用于 `excel-reader`、`excel-editor` 及其他跨引擎对拍与验收。SHA-256 逐格摘要只用于同一引擎的前后状态比较，不得用于跨引擎浮点值判等。依据：第二份独立基准报告附录“实际 S2 写入：浮点末位（1 ULP）序列化差异”。
- **读取适配器必须声明值解释口径**：Direct OOXML `CellValueProvider` 返回原始缓存值（bool 为 1/0、日期为序列号、错误值为缓存字符串）；Calamine 返回解释后的 bool/datetime/error 类型。适配器登记必须写明原始缓存或解释类型、日期/错误值及空值语义；声明前不得把另一口径作为替换读取路径的对拍基线。依据：第二份独立基准报告 `type_error_edge_probe` 及附录“取值口径差异”。
- **跨语言空单元格先归一再比较**：probe/对拍管道比较 python-calamine 与 Rust calamine 前，须把 Python `''` 和 Rust `Empty` 映射为已声明的同一空值状态，再按坐标比较；禁止将绑定层空值记号差异误报为数据差异。依据：第二份独立基准报告附录“real_xls_side_test”中的 14 × 21 样本，Python 侧共返回 294 个非 `None` 值，其中 214 个非空字符串、80 个空字符串；Rust 侧为 214 个 string + 80 个 Empty，实际非空字符串一致。

### 9.2 性能优化

仅当用户要求性能优化或当前明确问题确有必要时参考以下测量方法，不自动启动性能任务。

1. 先固定现有实现、输入、输出契约和基线，再定位当前功能的性能热点。
2. 只检查当前操作及直接依赖的重复打开/保存、整表副本、逐单元格 IPC/COM、重复矩阵构造、不必要转换、宽泛异常回退和语义缺口；不扩展为无授权全仓重写。
3. 同机、同输入、同输出契约预热后至少五次，报告中位数、范围、机器/库版本、样本规模、峰值内存和输出大小；测量必须覆盖该功能完整调用链，同时提供分阶段计时，并计入 Python/Rust 桥接的进程启动/导入、序列化、数据传输、数据转换、进程/IPC、隔离 runtime 和校验成本。允许 deb 内置 openpyxl worker 专用 Python runtime；不得依赖系统 Python。
4. 写入使用同一份预生成数据；分别记录同压缩配置和实际发布配置。单次微基准或仅 Rust 内部耗时不能支持业务提速结论。
5. 因性能原因考虑修改现有默认路由时，测量前先定义该功能的绝对节省时间、加速比例、峰值内存上限和测试规模，再按这些指标逐项判断；不能测后改门槛或编造比例。保留项/业务语义未通过时，不得因性能切换默认。
6. 兼容修复凭缺陷复现、目标契约和回归证据决策，不虚构性能收益。未有证据时保留已验收路径。

### 9.3 调整默认路由

改变具体功能的生产路由时，按该功能的操作、格式、平台和必须保留的对象核实能力，保留现有已验证路径，不能依据通用基准或菜单存在自动切换。验证范围以当前功能及风险为准，不设全项目对拍或固定覆盖率门槛；未获授权运行的测试明确记为未验证。原位编辑无法安全处理时应拒绝或明确报告，不得覆盖原始文件掩盖失败。

## 10. 按需排查与留痕

用户指出具体问题时，只核对相关入口、配置、输入输出、实际引擎、关键保留项和必要依赖；保护已有行为和证据，不启动全量业务盘点。发现后端能力缺口时在本仓库记录并按用户需求处理，无须等待其他项目或执行分支同步。

涉及 Excel/CSV 或引擎路由时，参照本规范和根 AGENTS.md；功能范围、状态及已知限制见 pytools-core/MIGRATION_PLAN.md 和 pytools-core/MIGRATION_INVENTORY.md。历史规则冲突分析、旧交接和完整阶段说明仅见 docs/archive/。保留真实配置、业务文件、测试夹具、历史报告和人工验收记录。
