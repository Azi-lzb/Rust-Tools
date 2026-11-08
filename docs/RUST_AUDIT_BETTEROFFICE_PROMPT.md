# Rust-Audit 接入 betteroffice 提示词（与 Rust-Tools 计算同步）

> 使用方式：把下方代码块整段复制给 Rust-Audit 仓库里的 agent 会话。
> 本文件与代码块内容一致；Rust-Tools 侧实现与验收数据见
> `docs/BETTEROFFICE_VENDOR_INTEGRATION.md`（提交 919750e、925630b、6f52d2b）。

---

```
任务：Rust-Audit 的 S1F1 校验区域重算接入 betteroffice 计算引擎（纯 Rust），
引擎路由、回退机制、配置项与 Rust-Tools 完全对齐，GUI 高级配置暴露全部开关。

【背景与参考实现】
Rust-Tools 已完成同源改造并验收（提交 919750e、925630b、6f52d2b）。
- 参考文档：C:\Users\AZI\Desktop\VibeCode\Rust-Tools\docs\BETTEROFFICE_VENDOR_INTEGRATION.md
- vendor 源码：C:\Users\AZI\Desktop\VibeCode\Rust-Tools\vendor\
  （betteroffice-xlsx-model、betteroffice-xlsx-calc 0.5.1 修改版，Apache-2.0）
- 外链机制区分文档：C:\Users\AZI\Desktop\VibeCode\Rust-Tools\docs\EXTERNAL_REF_DANGLING_MECHANISM.md
先读这三份再动手。禁止自己另写一版 vendor 修改。

【必须复用的 vendor 修改（四处，改动已在源码里，直接 copy 目录即可）】
1. xlsx-model 新增 Sheet::content_range()：只统计有值/有公式的存储格；
   CellProvider 的 used_rows/used_cols 改用它；原 used_range 行为不变。
   原因：上游把纯样式空格算进 extent，"格式刷满全列"的报送模板会把整列引用
   求值范围从内容行（42~68）撑到样式行（16,415）。
2. content_range() 加 OnceLock 记忆化 + set_cell 精确失效（仅"有无内容"转换时
   清缓存，重算改写公式值不击穿；remap_cells 无条件失效）。
   原因：不做记忆化时 12,162 公式 × 771,721 存储格 ≈ 94 亿次扫描，单轮 112 秒；
   做了之后 3.5 秒（32×），正确性零变化。
3. xlsx-calc 的 MAX_RECALCULATION_CELL_VISITS 常量改 AtomicU64 + 运行时
   set_max_recalculation_visits()/max_recalculation_visits()。
   原因：上游写死 1000 万且不可配置。
4. vendor 自带测试必须全绿（model 101、calc 257+集成），新增的 vendor 测试
   （content_range 三条 + whole_column_criteria + max_recalculation_visits）
   一并保留。

【接入要求】
- 根 Cargo.toml 加 [patch.crates-io] 指向 vendor path（版本保持 0.5.1），
  两个 vendor 目录加入 workspace.exclude；确保 --offline 构建通过
  （Cargo.lock 里这两条不应再有 source/checksum）。
- 新建等价于 Rust-Tools pytools-core/src/bo_calc.rs 的模块：
  open → 设预算 → recalculate_all(N 轮) → save → 统计 budget_exhausted /
  limited_cells / #NAME? 数 / 循环引用数。
- now_serial 必须显式注入（系统时钟换算 Excel 1900 序列号）。传 None 会让
  NOW()/TODAY() 及其下游全变 #VALUE!。
- 默认 2 轮重算（1~5 可配）：betteroffice 的依赖图看不到 INDIRECT 的动态引用，
  第 1 轮跨表动态链可能读到陈旧值；S1F1 模板实测第 2 轮即 changed=0 收敛。

【回退机制（核心，必须与 Rust-Tools 逐条一致）】
- 路由优先级：betteroffice → COM(Windows)/LibreOffice(Linux) → 明确报错。
- 触发回退的条件（任一命中即整簿回退，禁止部分采用）：
  a. budget_exhausted=true 或 limited_cells 非空（保留旧值不可信）
  b. 出现 #NAME?（betteroffice 未实现的函数）
  c. 存在循环引用格
- 回退引擎整簿重算，不依赖 betteroffice 留下的半新半旧缓存。
- 回退原因必须进运行日志（actual_engine 或等价字段），格式示例：
  "Excel.Application；betteroffice 回退（重算预算耗尽（1329 个公式格未求值，保留旧值不可信））"
- FALLBACK=none：不回退，整次任务明确失败并给出原因；绝不静默采用受限结果。
- COM 侧保留既有保护：Excel COM 打不开先跑 Excel 修复（补全角括号表名引号/
  修正冻结窗格矛盾，即 excel_open_repair），再按既有门控回 WPS。

【配置与 GUI 高级配置（两边统一）】
- 配置优先级：GUI/配置存储 > 环境变量 > 内置默认（与 Rust-Tools 的
  config.xlsx 全局配置 > env > default 完全一致）。
- GUI 高级配置界面必须暴露 5 个设置，命名/语义与 Rust-Tools 一致：
  1. 校验计算引擎：betteroffice（默认）｜com｜off —— 下拉
     说明：betteroffice=纯Rust优先（快，推荐）；com=直接用Excel/WPS
  2. 校验计算预算上限：留空=默认 1000 万 —— 数字输入
  3. 校验计算轮数：1~5，默认 2 —— 数字输入
     说明：第2轮起跨表动态引用（INDIRECT）才收敛；保持2即可
  4. 校验计算剥离样式空格：是/否，默认 否 —— 开关
     说明：是=重算前删掉工作副本里只有格式没有内容的空格子（大表更快）
  5. 校验计算回退策略：auto（默认）｜com｜libreoffice｜none —— 下拉
     说明：auto=Windows回Excel/WPS、Linux回LibreOffice；none=不回退直接报错
- 环境变量同名兜底：PYTOOLS_BO_CALC、PYTOOLS_BO_CALC_MAX_VISITS、
  PYTOOLS_BO_CALC_ROUNDS、PYTOOLS_BO_CALC_STRIP_STYLE、PYTOOLS_BO_CALC_FALLBACK。
- GUI 保存的值必须闭环生效：改 com 重跑 S1F1 应看到路由切到 COM，
  改回 betteroffice 应切回（两端都要验证）。

【已知坑（Rust-Tools 实测踩过，不要重踩）】
- betteroffice 把求值错误写成 <v>#ERROR!</v>——OOXML 只允许 7 种标准错误码，
  Excel 会以 0x80020009 拒开整个文件。S1F1 模板求值正常时不产生，但任何
  "落盘前统计到 #ERROR! 字面"都必须视为不可采用并回退。
- ooxml 的 strip 样式空格如果实现，重建 <row> 时不要多输出 '>' 或重放 '</row>'
  （首版缺陷，产物直接打不开；Rust-Tools ooxml-patch 44/44 有回归测试可抄）。
- 引擎替代品已评估淘汰，不要引入：IronCalc 0.9.1（输出被 Excel 拒开 +
  定义名求值全败）、visi-core 0.3.0（中文 UTF-8 字节截断 panic + 慢 640s）。

【禁止】
- 改业务算法、改菜单结构、为测试通过删断言。
- 动 [N] 外链 fail-closed 逻辑（模板含外链 → 拒绝，是 4-7 口径；
  ！70 模板在两条路由下报同样错误是预期）。
- 混淆两种复制来源：Rust-Audit 的 excel_open_repair 处理"外部文件导入"，
  S1F1 是"模板自身复制"，机制文档 EXTERNAL_REF_DANGLING_MECHANISM.md §1.1。

【验收标准（全部用 测试数据/S1F1_新旧对比_20261105_231620/business_templates 的 20 个模板）】
1. 20 模板逐个双路由（betteroffice / COM）跑 S1F1，比对最终校验结果表：
   命中条数相同、排序后 (工作表名, 定位单元格) 集合相同、逐字段只允许
   "15 位 vs 17 位浮点文本"差异（Excel 的 & 拼接截 15 位有效数字），
   实质差异必须为 0。
   已知基线（Rust-Tools 实测）：18/20 完全对齐；！20（引用不存在的 REF 表）与
   ！70（模板含 [1] 外链）在计算前被模板校验拦下且两条路由报同样错误——
   这是预期，不是回归。
2. 配置矩阵三分支：off（或 com）回到纯 COM 行为；MAX_VISITS=10 复现预算耗尽
   且自动回退、结果不变；FALLBACK=none 明确报错不静默。
3. GUI 闭环：5 个设置逐个改值 → 重跑 → 行为切换 → 改回。
4. 全量回归：不得有新增失败（Rust-Tools 基线 557/0/15）。
5. 报告：20 模板对拍表（命中数/实质差异/浮点差异/两引擎耗时）、三条配置分支
   实测、GUI 闭环证据、全量回归数字。

【未验证项如实记录（两边同口径）】
Linux 实机与 deb 交叉编译；循环引用与 #NAME? 拒绝分支的构造样本；
xlsm/图表/条件格式簿的重算保真；多轮迭代在 3 轮以上是否继续收敛。
完成后明确列出哪些项未验证，不得把自动化测试冒充人工终端验收。
```
