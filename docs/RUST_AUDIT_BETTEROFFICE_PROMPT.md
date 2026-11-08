# Rust-Audit 接入 betteroffice 的提示词（可直接复制给另一个仓库的 agent）

本文件只保存提示词正文；对应的实现与验收数据见
`docs/BETTEROFFICE_VENDOR_INTEGRATION.md`（Rust-Tools 已完成同源改造，
提交 `919750e`、`925630b`、`a23b2ca`）。

---

```
任务：把 S1F1 校验区域重算的引擎路由改为 betteroffice 优先，COM/LibreOffice 兜底

背景：Rust-Tools 已完成同源改造（提交 919750e、925630b、a23b2ca），
参考实现与验收数据在
C:\Users\AZI\Desktop\VibeCode\Rust-Tools\docs\BETTEROFFICE_VENDOR_INTEGRATION.md，
vendor 源码在 C:\Users\AZI\Desktop\VibeCode\Rust-Tools\vendor\
（betteroffice-xlsx-model、betteroffice-xlsx-calc 两个 0.5.1 修改版，Apache-2.0）。
先读那份文档再动手。

必须复用的 vendor 修改（不要自己另写）：

1. xlsx-model 新增 Sheet::content_range()：只统计有值/有公式的存储格；
   impl CellProvider for Workbook 的 used_rows/used_cols 改用它，原 used_range 不动。
   原因：上游把纯样式空格算进 extent，"格式刷满全列"的模板会把整列引用撑大数百倍。

2. content_range() 加 OnceLock 记忆化 + set_cell 精确失效（仅"有无内容"转换时清缓存）。
   原因：不做记忆化时 12,162 公式 × 771,721 存储格 ≈ 94 亿次扫描，单轮 112 秒；
   做了之后 3.5 秒，且正确性零变化。

3. xlsx-calc 把 MAX_RECALCULATION_CELL_VISITS 改成 AtomicU64 + 运行时
   set_max_recalculation_visits()。原因：上游写死 1000 万不可调。

接入要求：

- 通过根 Cargo.toml 的 [patch.crates-io] 指向 vendor path，版本号保持 0.5.1，
  并把两个 vendor 目录加进 workspace.exclude。确保 --offline 构建通过。

- 新建等价于 Rust-Tools pytools-core/src/bo_calc.rs 的模块，封装
  open → 设预算 → recalculate_all(N 轮) → save → 统计受限格/错误格。

- now_serial 必须显式注入（用系统时钟换算 Excel 1900 序列号）。
  传 None 会让 NOW()/TODAY() 全变 #VALUE!。

- 默认 2 轮重算：betteroffice 的依赖图看不到 INDIRECT 的动态引用，
  第一轮可能读到陈旧值；S1F1 模板实测第 2 轮即 changed=0 收敛。

- 必须 fail-closed 回退：budget_exhausted 或 limited_cells 非空、
  出现 #NAME?、存在循环引用 → 一律回退 COM（Windows）/LibreOffice（Linux），
  绝不能把"保留旧值"当"算对了"。

- 路由优先级：betteroffice → COM → LibreOffice → 明确报错。
  回退原因写进运行日志的 actual_engine 字段（形如
  Excel.Application；betteroffice 回退（重算预算耗尽（N 个公式格未求值…）））。

- 配置项（建议与 Rust-Tools 同名，便于两边统一）：
  PYTOOLS_BO_CALC（auto/off）、PYTOOLS_BO_CALC_MAX_VISITS、
  PYTOOLS_BO_CALC_ROUNDS（默认 2）、PYTOOLS_BO_CALC_STRIP_STYLE（默认 off）、
  PYTOOLS_BO_CALC_FALLBACK（auto/com/libreoffice/none）。

验收标准（用 测试数据/S1F1_新旧对比_20261105_231620/business_templates 全部 20 个模板）：

- 每个模板用自己的副本作源，分别跑 betteroffice 与 COM 两条路由，
  比对最终校验结果表：命中条数相同、排序后 (工作表名, 定位单元格) 集合相同、
  逐字段比对时只允许"15 位 vs 17 位浮点文本"差异，实质差异必须为 0。

- 已知基线：18/20 完全对齐；！20（公式引用不存在的 REF 表）与
  ！70（模板含 [1] 外链）会在计算前被模板校验拦下，两条路由报同样错误——
  这是预期结果，不是回归。

- 额外验证三条配置分支：PYTOOLS_BO_CALC=off 回到纯 COM 行为；
  把 MAX_VISITS 设成 10 复现预算耗尽并确认自动回退且结果不变；
  FALLBACK=none 时明确报错不静默。

- 全量回归：不得有新增失败。

禁止：改业务算法；改 Excel 引擎选型（除本次新增的 betteroffice 优先路由）；
改菜单结构；为让测试通过而删断言；动 [N] 外链 fail-closed 逻辑。

注意：Rust-Audit 若已有自己的 excel_open_repair 相关外链处理，那是外部文件导入
场景，与 1-4 的"模板自身复制"不同，不要混改（详见 Rust-Tools
docs/EXTERNAL_REF_DANGLING_MECHANISM.md §1.1）。

完成后报告：20 模板逐条对拍结果表、三条配置分支实测、全量回归数字、
以及 betteroffice 与 COM 的逐模板耗时对比。
```
