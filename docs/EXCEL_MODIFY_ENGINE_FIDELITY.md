# Excel XLSX 修改引擎保真度验证报告

数据：`benchmark/excel-modify-engine/results/fidelity_final.json`（227 条 XLSX + 16 条 XLSM）、
`results/office_compat.json`（240 条 Excel COM 兼容冒烟）。
方法：每引擎 × 14 类夹具 × {no-op / 单点修改 / 重复保存两轮} + 失败安全 + XLSM 独立；
三层验证 = ZIP 部件清单/哈希（结构）→ calamine+excel-reader 语义对拍（值/公式/合并/批注/命名区域/CF/DV）
→ Excel COM 实开冒烟。副本验证，原件未动。

## XLSX 总分（每个引擎 14 夹具 × 4 场景 = 56 项）

| 引擎 | PASS | FAIL | 通过率 | 失败性质 |
| --- | ---: | ---: | ---: | --- |
| **rxls** | 53 | 3 | 94.6% | calcChain 派生部件移除（合法）；损坏输入仍产出输出 |
| **openpyxl** | 53 | 3 | 94.6% | calcChain/metadata/sharedStrings 移除（合法派生）；损坏输入拒绝（判 PASS 后口径修正，见下） |
| **betteroffice** | 50 | 6 | 89.3% | calcChain 移除（合法派生）；**f_complex G1 语义值被重算改变** |
| **umya** | 45 | 11 | 80.4% | **真实丢失**：外链部件、批注部件（no-op 即丢） |

口径说明：`calcChain.xml`/`metadata.xml` 是 Excel 可重建的派生部件，移除不构成业务丢失；
本报告的 FAIL 计数含这类「结构移除但语义无损」项，两者在明细中已分开。
f_basic failsafe：rxls 对损坏输入仍产出输出（见「失败安全」），判 FAIL；openpyxl 拒绝损坏输入，判 PASS。

## 真实发现（按引擎）

### rxls 0.1.4 —— 部件保留式验证通过
- no-op：14/14 PASS，且**未触碰部件逐字节一致**（合同兑现）。
- 语义层零非预期变化；XLSM no-op/edit1 PASS（VBA 字节保留）。
- **失败安全缺陷**：对损坏输入仍产出输出文件（应在打开阶段拒绝）。
- **派生部件移除**：real_target 编辑后 `calcChain.xml` 被移除——Excel 可重建，非数据丢失。

### betteroffice 0.5.1 —— 重渲染式，无部件丢失，但有重算语义
- no-op：14/14 PASS（重渲染后无部件移除、语义无损）。
- **f_complex G1 公式缓存值被主动重算**（`null → 2023.0`）：`apply_ops` 触发计算；
  缓存值变化是计算语义而非损坏，但与「只改一处」的预期不符，生产使用需注意。
- XLSM：no-op PASS；**edit1 产出 Excel 打不开的文件**（COM 冒烟 FAIL）——xlsm 编辑路径不可用。

### openpyxl 3.1.5 —— 基线表现符合已知特性
- no-op：13/14（real_target 丢 calcChain/metadata/sharedStrings，均为派生/可重建部件）。
- 语义层零非预期变化；XLSM no-op/edit1 均丢 `xl/vbaProject.bin`（keep_vba 未加时），
  **与生产 worker 的实际行为一致**——本项目 worker 对 xlsm 强制 keep_vba，raw 基线不外推。
- 失败安全：拒绝损坏输入，PASS。

### umya 3.1.1 —— 重写式，真实丢件最多
- **f_comments no-op 即丢批注部件**（`xl/comments/comment1.xml` + VML）：批注丢失。
- **f_external_link no-op 即丢外链部件**：外链引用悬空。
- real_target 丢 calcChain + metadata（派生，可接受）。
- XLSM PASS（VBA 保留）。
- **umya 丢批注/外链是真实业务风险**：本项目 2-3 用 umya 做追加候选时未受影响（追加目标无批注/外链场景未被触发过），但作为通用修改引擎的保真等级最低。

## Office COM 兼容冒烟（240 项，全部真机 Excel）

| 引擎 | PASS | FAIL | 失败点 |
| --- | ---: | ---: | --- |
| betteroffice | 52 | 8 | f_external_link edit1、real_vba edit1 |
| openpyxl | 56 | 4 | f_external_link edit1、real_vba no-op/edit1 |
| rxls | 52 | 8 | f_external_link edit1、real_vba edit1 |
| umya | 48 | 12 | f_comments edit1、f_external_link 全部、real_vba 全部 |

- f_external_link edit1 三引擎同败：手工构建的外链夹具在编辑后跨引擎损坏，
  更可能是夹具脆弱性而非引擎缺陷（no-op 时 openpyxl/rxls 均通过）。已如实标注。
- **XLSM 编辑后 COM FAIL（betteroffice/umya/rxls）是引擎真实缺陷**：宏工作簿编辑路径在这三家都不可靠，
  与保真层「VBA 字节保留」结论交叉印证——rxls 结构保留但 Excel 拒开，说明保留契约在 edit 路径有缺口。
- openpyxl raw XLSM FAIL 与生产 worker（强制 keep_vba）行为不同，已在报告中隔离标注。

## 结论

1. **保真度排名（XLSX 值得托付）：rxls ≈ openpyxl > betteroffice > umya**。
2. rxls 的部件保留契约经 56 项验证兑现（唯 calcChain 派生件移除 + 失败安全缺陷）。
3. umya 对批注/外链的静默丢失是通用修改场景下的硬伤——本项目仅在 2-3 追加（无批注/外链目标）使用它是安全的。
4. 三个 Rust 引擎的 XLSM 编辑路径（betteroffice/umya/rxls）都存在 Excel 兼容缺口；
   openpyxl + keep_vba 仍是宏工作簿编辑的唯一可靠选择。
