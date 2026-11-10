# Excel XLSX 修改引擎技术选型建议

数据来源：`EXCEL_MODIFY_ENGINE_CAPABILITY.md`、`EXCEL_MODIFY_ENGINE_FIDELITY.md`、
`EXCEL_MODIFY_ENGINE_BENCHMARK.md` 及 `benchmark/excel-modify-engine/results/` 原始 JSON。
原则：保真度优先，性能其次，功能覆盖第三，维护成本与部署复杂度纳入决策。

## 一、逐操作推荐矩阵（M01–M14 + 补充）

| 操作 | 推荐引擎 | 依据 |
| --- | --- | --- |
| M01–M03 单格/批量值写入 | **umya**（性能首选）或 rxls | umya 每簿 274–430ms、保真无值丢失；rxls 批量 API 后同档且部件保留 |
| M04 公式文本 | umya | 198ms vs rxls 747ms / betteroffice 890ms；umya 该操作零样式风险 |
| M05 数字格式 | umya 或 openpyxl | rxls 无 API；betteroffice 可用但 890ms；umya 直接 |
| M06 字体/填充/边框/对齐 | umya 或 openpyxl | rxls 无 API |
| M07 行高列宽 | umya / rxls / betteroffice 均可 | 性能 umya 最快 |
| M08 合并增删 | umya | 同上 |
| M09 批注 | **openpyxl 或 rxls** | betteroffice 无批注表示；umya 保存丢批注部件（保真 FAIL）|
| M10 命名区域 | rxls / umya / openpyxl | betteroffice 可用但无优势 |
| M11 打印设置 | umya / openpyxl | betteroffice/rxls 无页面方向 API |
| M12 表增删重命名 | umya / rxls | rxls rename 会重写公式内引用（加分）|
| M13 复杂模板单点 | **rxls**（部件保留）或 umya | rxls 未触碰部件字节不变；umya 快但丢批注/外链（本夹具即受影响）|
| M14 混合修改 | umya / openpyxl | betteroffice 缺批注子操作 |
| 行列插入删除 | umya / openpyxl / betteroffice | rxls 官方不支持 |
| XLSM 编辑 | **openpyxl + keep_vba（维持现状）** | 其他三引擎编辑后 Excel 打不开或丢宏（COM 冒烟证实）|
| 大文件（≥50k 行）单点修改 | umya | 11s@50k / 53s@100k，唯一 100k 可完成的 Rust 引擎 |

## 二、任务七的 12 问回答

1. **rxls 比 openpyxl 快吗？** 是（批量 API 用法下）：w1 4.7x、w100 2.9x、w1000 2.7x、
   批量 100 簿常驻 3.9x。但逐格 `set_cell_value` 用法是数量级慢路径（官方文档区分的用法差异）。
2. **BetterOffice 比 openpyxl 快吗？** 混合：w1 2.4x 快；w1000 0.3x（慢 3 倍+）；
   批量 100 簿 1.15x（慢）。`apply_ops` 的公式图重建成本随规模超线性。
3. **rxls vs BetterOffice 谁性能好？** rxls 全线占优（批量 100 簿 3.9x vs 1.15x；
   w1000 2.7x vs 0.3x），且内存更低（rxls 46MB vs betteroffice 39–87MB，100k 时 3.6GB）。
4. **保真度谁最高？** rxls（53/56，唯一部件字节保留兑现）≈ openpyxl raw（53/56，
   但生产 worker 有额外缓存恢复机制实际更高）> betteroffice（50/56 + 重算语义）> umya（45/56，丢批注/外链）。
5. **谁适合复杂模板？** rxls（部件保留 + StyleLoss 类型化报告）与 openpyxl（成熟）。
   betteroffice 重算语义会改公式缓存；umya 丢批注/外链。
6. **谁适合批量修改？** umya（274ms/簿）与 rxls（356ms/簿）第一梯队；
   betteroffice（1159ms/簿）与 openpyxl 桥接不具竞争力。
7. **rxls 特别适合的操作**：复杂模板单点修改（M13，部件保留）、值矩形批量写入、
   不需样式的数据表编辑、XLSM no-op 场景。
8. **BetterOffice 特别适合的操作**：已有 `apply_ops` 生态的值+样式+公式混合编辑
   （`PatchRangeStyle` 能力完整），以及**需要公式重算作为副产品的场景**（本项目 1-4 已在用）。
9. **能否替代 openpyxl？** 不能整体替代：rxls 缺样式/数字格式 API，betteroffice 缺批注；
   且 XLSM 编辑三者都不可靠。openpyxl 仍是唯一全能力 + 唯一可靠 XLSM 编辑。
10. **能否替代手写 OOXML？** 部分可以：rxls 的部件保留契约可覆盖 4-7 修复类的
    「未触碰部件不动」需求雏形，但其行/列操作缺失与 100k 打开上限使其
    **无法替代** `ooxml-patch` 的修复与窄范围补丁（那些操作需要字节级手术）。
11. **减少运行时依赖？** 若仅替换值/样式类编辑可去掉 Python（rxls/betteroffice 均纯 Rust）；
    但 XLSM 编辑仍必须保留 openpyxl worker，Python 运行时无法删除，只能缩小使用面。
12. **立即迁移还是观察？** 见下。

## 三、明确建议

### 主结论：**D（暂时保留 openpyxl）+ 定向试点 C（按操作选型）**

- **生产默认引擎维持 openpyxl worker 不变**。理由：唯一 14/14 能力覆盖 + 唯一可靠
  XLSM 编辑 + 生产链路的缓存恢复/守卫体系都是围绕它建成的。
- **umya 维持现状**（2-3 追加候选）：性能冠军，但其批注/外链丢失使它只能用于
  「无批注、无外链」的目标——现有 2-3 路由已隐式满足，扩大使用面需先加结构守卫。
- **rxls 列入观察名单，不建议生产使用**：保真验证最优（94.6%）+ 性能第一梯队，
  但 0.1.x 成熟度、100k 打开超限、失败安全缺陷（损坏输入仍产出输出）、无样式 API
  四项叠加，尚不满足生产引擎门槛。**等 0.2+ 稳定版补齐样式 API 与大文件上限后重评。**
- **betteroffice 维持现有 1-4 计算角色，不进入修改路径**：修改能力可用但性能
  （w1000 0.3x、批量 1.15x）与内存（100k 3.6GB）不具优势，且缺批注。
- **手写 OOXML 保留**：rxls 现状无法覆盖 4-7 修复的字节级手术需求。

### 可立即试点（不影响生产默认）

- **无样式数据表批量值写入**（如 2-3 追加候选的值模式）可试点 rxls
  `set_cell_range_values`：性能 2.7–2.9x、部件保留、无 VBA 场景。试点须带
  rxls 现有的 calcChain 移除行为核查 + 失败安全补丁（拒绝损坏输入）。

### 迁移成本与风险提示

- XLSM 编辑是 openpyxl 的不可替代项 → Python 运行时无法完全移除，`减少依赖`的收益上限有限。
- rxls 0.1.x（12 stars、24 open PRs、0.1.4 尚未在 README 正式宣布）API 变动风险高，
  vendoring 成本（本项目已有 betteroffice vendor 先例）可控但需锁版本+补丁管理。
- betteroffice vendor 已在本仓库维护 4 处补丁，新增引擎会复制这份维护负担。

## 四、后续独立任务建议（本轮不做）

1. rxls 失败安全补丁验证（损坏输入拒绝）——向上游提 issue 或本地包装。
2. umya 批注/外链丢失的上游确认与规避补丁评估。
3. rxls 0.2 发布后重跑本 benchmark（夹具与驱动脚本已入库可复现）。
4. 2-3 追加候选 rxls 试点的定向对拍（真实业务样本）。
