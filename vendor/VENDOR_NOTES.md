# vendor：betteroffice-xlsx-model / betteroffice-xlsx-calc 0.5.1（本地修改版）

来源：crates.io `betteroffice-xlsx-model 0.5.1`、`betteroffice-xlsx-calc 0.5.1`
上游：https://github.com/openooxml/betteroffice
许可证：Apache-2.0（http://www.apache.org/licenses/LICENSE-2.0）
       本目录为修改后再分发，保留上游版权与 Apache-2.0 授权声明；
       修改点逐条登记于下表。上游发布新版本时不自动升级，需人工评估后同步。

接入方式：根 `Cargo.toml` 的 `[patch.crates-io]` 以 path 指向本目录，
版本号保持 `0.5.1`，上层 `betteroffice-xlsx 0.5.1` 的依赖约束照常匹配。

## 修改点

| # | crate / 文件 | 修改 | 理由 |
| --- | --- | --- | --- |
| 1 | xlsx-model `src/workbook.rs` | 新增 `Sheet::content_range()`：只统计**有值或有公式**的存储格；原 `used_range()` 行为不变 | `used_range` 把纯样式空格也算进 extent，"格式刷满全列"的报送模板会把整列引用的求值范围撑大数百倍（实测样式把 extent 从 68 行撑到 16,415 行） |
| 1b | xlsx-model `src/workbook.rs` | `content_range()` 加 `OnceLock` 记忆化（`ContentExtent`，同 `CellStyles` 范式）；`set_cell` 仅在内容有无转换时失效，`remap_cells` 无条件失效 | 每次调用都全表扫描 `cells`，大簿 12,162 公式 × 771,721 存储格 ≈ 94 亿次遍历；省信贷 1 轮由 112.5s 降至 3.5s（32×），正确性零变化 |
| 2 | xlsx-model `src/workbook.rs` | `impl CellProvider for Workbook` 的 `used_rows/used_cols` 改为基于 `content_range`（数组公式 spill 逻辑保留） | 计算 provider 是整列引用截断的唯一入口 |
| 3 | xlsx-calc `src/eval.rs` | `MAX_RECALCULATION_CELL_VISITS` 常量改为 `AtomicU64` 全局默认值 + `pub fn set_max_recalculation_visits(u64)` | 上限原为编译期 const 10,000,000 不可配置；业务侧需按 `PYTOOLS_BO_CALC_MAX_VISITS` 配置 |
| 4 | xlsx-calc `src/eval.rs` | `format_number` 改为 Excel 的 **15 位有效数字**（原用 Rust 最短往返表示，最多 17 位） | Excel 的 `&` 文本拼接只保留 15 位有效数字；17 位会让校验描述文字出现尾数差异（实测 8/8 样本按 15 位格式化即与 Excel 一致） |

## 配套测试（登记于各 crate tests）

- xlsx-model：样式空格不进入 `content_range`；`used_range` 行为与上游一致（回归）。
- xlsx-calc：整列 SUMIFS 在"内容 5000 行 + 样式刷到 16400 行"场景下按内容行求值且预算不耗尽；
  `set_max_recalculation_visits` 设小值后 `budget_exhausted=true`。
- xlsx-model：`content_range` 缓存只随内容增减失效（样式格与值改写不击穿），见
  `content_range_cache_invalidates_only_on_content_change`。

## 同步策略

锁定 0.5.1。上游若修复 `used_range` 语义（或合并本修改方向），评估后摘除 patch 回归 crates.io 版本。
