# Rust-Tools Excel 读写、修改与引擎路由规范

> 当前项目范围：仅适用于独立 Rust-Tools 仓库。用户本轮明确要求和根 AGENTS.md 优先；本文件保留具体功能所需的 Excel 安全与能力边界。历史基准和性能数据仅供按需参考，不是全量迁移、全面样式对拍或跨项目同步任务。
> 旧联合工作区规范原文已保存在 docs/archive/Excel规则_旧共享工作区_20261027.md；旧分支、旧路径和 Agent 分工不再生效。

## 1. 项目目标与已确认边界

Rust-Tools 独立维护现有 PyTools 业务和仓库内 Excel-Core。Rust 负责业务核心、读取、规则计算、新建工作簿及 ExcelEditPlan/ExcelModifier 调度；既有工作簿修改按具体能力和证据选择本地后端。隔离 openpyxl worker 仅可用于受限的既有 XLSX 修改，未来 Windows COM 仅为单独验证且显式选择的可选能力，均不得承载 PyTools 业务逻辑。原 Python 项目只读，仅在用户需要时作为行为参考。公式计算与工作簿修改分开，默认关闭。

目标交付为麒麟 V10 x86_64 和统信 UOS V20 x86_64 离线 deb，兼容 glibc 2.28；核心路径不得要求用户预装 Python、Go、Office、WPS、LibreOffice 或 JVM。启用 openpyxl 引擎时可在 deb 内置隔离 Python runtime 与依赖，不得使用系统 Python、用户 PATH 或联网 pip。可选 LibreOffice 计算能力须明确独立运行时来源；Windows COM 不进入 Linux deb 或其必要依赖。ABI 和打包约束见根 AGENTS.md。

### 1.1 主路径和旁路

- 核心主路径不依赖 Office、WPS、LibreOffice、COM 或 UNO。
- 公式计算、同类别格式转换及未来 Windows COM 修改只能作为独立、显式选择的可选能力；其不可用不影响读取、新建和其他业务路径。
- 公式计算通过独立 ExcelCalculator 提供，默认关闭。用户可显式请求普通重算或完整重算；必须验证公式语义、结果和依赖关系，并检查引擎保存造成的公式、样式或结构变化。不得用旧缓存、零值或静态重建冒充重算；未安装/不可用时报告“该功能需要外部引擎”。
- 同类转换由用户明确触发，可选择经对拍的保真转换，或经说明损失范围的 Rust/Go 转换。已确认方向包括 Excel .et/.xls → .xlsx、Word .doc → .docx；.xls 可作为改工作表名等操作前的显式转换。具体方向/引擎未登记通过前仍属未验证；结果必须验证 sheet、数据和契约要求保留的对象。
- PDF、纸张渲染、打印机输出和跨类别转换不在当前迁移范围。打印只写页面设置。
- sidecar/worker 是独立进程，主程序启动时不检查外部引擎。用户显式触发时才检查依赖；缺失/失败不能阻塞纯 Rust 主路径，未验证输出不得报告成功。
- 未来 ExcelModifier 可选接入 Microsoft Excel COM、WPS 表格 COM 两个独立 Windows 后端：优先使用 pywin32 / `win32com.client` 与工作簿级 JSON IPC，Excel 使用 `Excel.Application`，WPS 优先验证 `KET.Application`。二者使用独立 EngineId 和 capability 证据，按真实 COM 激活及具体操作验收；不能仅凭软件已安装判定可用。显式选择不可用后端须报错，不静默切换。COM 不提供 Reader/Writer，也不进入 Linux 依赖；不使用 Rust COM、PyO3 或 FFI。只有端到端测量证实 IPC 为瓶颈后才评估替代通信方式。
- LibreOffice Calc 可作为未来独立 ExcelCalculator 后端（Windows/Linux）；Microsoft Excel COM、WPS COM 计算候选仅限 Windows，WPS 单独验收。ExcelCalculator 不纳入 ExcelModifier；LibreOffice 暂不纳入修改后端。计算适配器可使用适合引擎的通信接口，不强制 JSON IPC。
- PDF、纸张渲染、打印机输出和跨类别转换不在当前迁移范围。打印只写页面设置。旧 COM 业务工作流不得整体搬入 sidecar/worker 以规避 Rust 迁移。
- 可选 LibreOffice 计算能力不得假设 Linux 目标机预装；发布前须说明独立离线组件/安装包或受支持外部安装等运行时来源，并完成目标平台验证。

### 1.2 已确认模板与打印范围

历史已检查模板不含图表、透视表、切片器、图片、工作表保护或 VBA；这不代表其他输入没有这些对象。修改前按当前操作确认实际文件中相关的必要对象。

工作表提取保留普通无批注版，带批注的 COM 分支已明确排除。PyCode 与 NoComPyTools 的其他差异保留为历史记录，仅按用户具体需求核实。

打印只写 orientation、fitToPage、fitToWidth、fitToHeight、清除 scale、paperSize、printArea、horizontalCentered、verticalCentered 和页脚页码。旧代码没有写 margins、printTitleRows/Columns、gridLines、headings、blackAndWhite、draft、firstPageNumber；新增属性须另立需求，不得顺手加入。

## 2. 项目结构与依赖边界

本仓库的业务逻辑位于 pytools-core/，Excel 原语位于本仓库 crates/，转换辅助位于本仓库 sidecar/lo-convert/。它们是 Rust-Tools 自己维护的实现，不依赖旧 Main、其他项目或旧共享 workspace，不通过合并、cherry-pick 或自动同步取得更新。

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
| reference_app / engine_mode | none、Microsoft Excel、WPS、LibreOffice；Excel/WPS 作为独立参考结果，不假设一致。Windows COM 脚本可为读取/新建测试生成参考数据或文件，但不进入 Reader/Writer 生产路由。仅提供已实现、环境可用且 capability 已登记通过的选项；用户手动选择失败时不得静默换引擎。 |
| workload_limits | 记录覆盖的文件/sheet/行列/对象数量、耗时和内存；没有用户阈值先测，不编造硬限制。 |

修改已有工作簿时，按本次实际操作和文件特征保护值/类型、公式、合并、工作表、批注、名称、VBA 及其他约定必须保留的对象和结构。对可能受影响但未检查的对象注明未验证，不假设其不存在；输入特征变化时应重新判断能力是否匹配。普通字体、边框等外观差异不自动阻塞交付。

## 4. 技术分工

### 4.1 读取

- 生产读取使用纯 Rust reader；普通值读取优先评估纯 Rust calamine，且必须按数据归属、文件格式、输出契约和对象验收。禁止 python-calamine、PyO3，openpyxl worker 不得用于读取路由。
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

- 对 Rust 持有结果、从零生成纯数据 XLSX 的功能，可评估 rust_xlsxwriter；按当前功能的输出契约和实际验证范围使用，它不是既有工作簿编辑器。不要为未指定功能自动开展对拍或性能测试。
- 按本次业务需要核验值/类型、公式、必要结构和文件可打开性；格式或性能比较仅在用户要求或当前问题确有必要时针对相关范围进行。
- 基准报告中的 Direct OOXML 写入是 Python S5/S2 专用 writer；报告没有比较 rust_xlsxwriter 与 Rust Direct OOXML writer。不得把该实现称为 Rust writer，也不得仅凭名称、压缩设置或 ZIP 行写方式推断不同 writer 等价。
- Direct OOXML 仅在存在专用输出契约、明确 OOXML 对象要求，或同契约实测证明资源收益时进入候选；constant-memory 模式、既有工作簿编辑和通用 OOXML writer 均须单独实现与验收。

### 4.3 修改已有工作簿

业务层生成 ExcelEditPlan，由 ExcelModifier 按“操作子类 + 格式 + 对象特征 + 保留项 + 平台/版本”选已验收引擎。未来 COM 后端仅修改既有工作簿；不得为统一形式而引入通用 ExcelWritePlan 或完整工作簿对象模型。

- **Rust OOXML Patch**：适合明确部件、明确位置、依赖关系可证明的窄修改；仅对已验收白名单执行，无法证明安全时拒绝。
- **Rust umya-spreadsheet**：复杂 XLSX 对象修改候选。按具体操作和需要保护的对象验证；能打开/保存不等于全部必要结构保留。
- **Go Excelize worker**：条件性候选。仅当 umya/定点补丁对真实模板有大面积失败或维护成本过高，且 Excelize 对同一整批操作验证适用时才评估。Rust 以一个批量 IPC 请求传文件路径和完整计划，不逐单元格跨进程。若采用，随包交付，CGO_ENABLED=0。
- **隔离 openpyxl worker**：可选第四修改引擎。Rust 按每本 workbook 提交一份结构化 ExcelEditPlan，由内部绝对路径启动隔离 runtime 中的 Python/openpyxl。worker 仅执行 XLSX 修改操作，不做读取业务、统计、规则计算、公式重算或新建报表。对非空公式缓存、未知/未验收 package part、超出 capability 的格式/对象或操作批次，必须在写入前拒绝。openpyxl 会重写 XLSX package，必须验证未修改区域及计划要求保留的对象；行插入/删除当前 unsupported。运行时缺失可作为写入前不可用处理，只有存在另一条已验证路由才允许重新从原始输入执行。崩溃、超时、写入后失败或验证不一致不得自动切换。
- **未来 Microsoft Excel/WPS COM worker**：仅为可选 Windows 修改候选，分别以 Excel `Excel.Application`、WPS `KET.Application`（需验证）激活；优先通过隔离 Python runtime 中的 pywin32 / `win32com.client` 接收一份工作簿级 ExcelEditPlan 和 JSON IPC，不依赖系统 Python/PATH。每种应用、版本、操作组合、对象、保留契约分别验收。不是 Reader/Writer，不作为当前实现或默认；Windows 产品若交付该 worker，须单独验证 runtime 与 pywin32 依赖。
- 修改引擎不预设全局优先级，也不要求全部启用。只围绕当前功能检查相关能力；引擎注册和合成测试不等于业务能力或用户人工验收通过。
- umya 已覆盖且未证明 Go 有保真、性能或维护收益时，不引入 Go。

### 4.4 业务计算、公式计算与禁止项

规则、期间、校验、统计、去重、映射和审核结果由 Rust 业务逻辑计算；财务数值、Decimal、舍入、日期、空值、错误和文本前导零策略按当前功能契约明确并验证。业务计算、公式文本、旧缓存读取、Excel 公式重算是不同能力，不能互相冒充。

openpyxl 仅可通过隔离 worker 用于已验收的既有 XLSX 修改；未来 pywin32 / `win32com.client` 仅可通过独立 Windows worker 用于已验收的 COM 修改或公式计算。它们不得用于生产读取、新建写入、业务规则、统计或通用业务逻辑。worker 不依赖系统 Python/PATH，修改按工作簿批量处理，不逐格启动；对拍/测试 Python 仍可在 tests/compare 使用。禁止 python-calamine、PyO3、pandas、xlrd、pyxlsb 进入生产。不得无证据重写已验收模块；不得以完整 OOXML writer 为默认目标；不得无测量串联 Rust 读取→中间文件→Go 修改→Go 写回。

`ExcelCalculator` 是与 ExcelModifier 分离的未来能力；本阶段计算目标为 LibreOffice Calc（Windows/Linux）和 Microsoft Excel COM（Windows）。此前列出的 WPS COM 是历史候选，不属于当前计算范围，除非用户另行确认。默认关闭，计算模式区分普通重算与完整重算。LibreOffice 基础公式计算的可行性已由用户实际测试确认，不重复泛泛的可行性调研；具体公式覆盖率、与 Excel 的公式语义一致性、保存保真和 Linux 离线运行时仍未验收。不得宣称计算结果完全等价。显式选择不可用或未经该功能验收的引擎时返回明确错误；界面仅列出已实现、可用且有匹配能力证据的选项。

### 4.5 本项目的 Excel-Core 维护

- 按用户明确提出的业务问题维护本地 Excel-Core，保留已有实现和兼容路径，不为其他项目统一而覆盖或重写。
- 读取、新建和既有工作簿修改分别按当前操作的实际契约选择已有能力；改变生产默认路由应具备对应操作、格式和必要保留项的证据。
- 保留现有 openpyxl worker、接口、测试和兼容依赖；未经测试，不得移除旧功能。Go/Excelize、COM 或其他扩展不是默认任务。
- 既有区域复制默认保留公式原文；引用平移和公式重算须单独明确。用户未提出相关功能时，不主动开展引擎扩展或全面验收。

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
| openpyxl worker | `crates/excel-editor/python/openpyxl_worker.py` + Rust `OpenpyxlXlsxBackend`；pytools capability dispatcher 已注册；支持有限 XLSX EditPlan 批处理 | 已实现并通过合成 worker/进程集成测试；Rust 独立读回覆盖已测样本的样式语义、显式行高列宽、填充色、页面设置和 OOXML Print_Area。真实项目模板、操作级 capability、Windows 原生 Rust、glibc 2.28 隔离 runtime、deb 打包和目标机均未验证。无生产 capability JSON、无默认路由；不支持行插删与 ImportRange 尺寸复制，非空公式缓存/未知 parts 写前拒绝。 |
| sidecar/lo-convert | 同类格式转换已接入；`recalculate` 命令已实现，含缓存清理/完整计算标记、输出校验和 Ubuntu/WSL2 LibreOffice 合成 `SUM` 测试 | 仅候选实现；用户已实际验证 LibreOffice 基础公式计算可行性，不重复泛泛可行性调研；真实业务公式、Excel/WPS 一致性、复杂依赖、Windows COM、超时、目标 Linux/glibc 2.28 和离线运行时来源均未验收；不能宣称通用公式能力或生产默认。 |
| Excel COM / WPS COM worker | 当前未实现；现有 ExcelModifier 无对应已验收后端 | 尚未实现。未来分别登记 Windows-only Microsoft Excel 与 WPS capability；不得展示为可用选项或用于 Linux 依赖。 |
| ExcelCalculator | 尚未形成统一公共计算接口 | 未来独立能力；LibreOffice Calc、Excel COM、WPS COM 为候选，按后端/公式/平台验收，不改变当前读写修改默认。 |

细节和报告见根目录 EXCEL_ENGINE_ARCHITECTURE.md、PYTHON_EXCEL_OPERATION_INVENTORY.md 及各自 probe/reports。报告数字只适用于其中样本和断言。已经通过的窄范围实现继续保留，不因本规则接入而重构。

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

仅针对用户指定的功能和必要依赖选择验证范围。优先核对业务结果、源文件安全、公式及该操作必要的结构；用户明确要求或具体问题需要时，再比较相关格式和引擎差异。不执行全项目对拍，也不以普通字体、边框等外观差异作为普遍交付门槛。未运行的验证如实记录，合成样本与用户真实样本分开；Rust 编译、运行和测试须按根 AGENTS.md 获得明确授权。

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
