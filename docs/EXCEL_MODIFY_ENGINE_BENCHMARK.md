# Excel XLSX 修改引擎性能 Benchmark 报告

数据：`results/perf_single.jsonl`（115 条）、`results/perf_scale.jsonl`（56 条）、`results/perf_batch.jsonl`（84 条）。
口径：完整链路 打开→解析→修改→保存→关闭；中位数；每组合 ≥3–5 次重复；
峰值内存 = Windows Job Object `PeakJobMemoryUsed`；输出大小已记录。
批量模式区分冷启动（每文件一次进程）与常驻（一进程处理全部）；
openpyxl 分「Python 原生」与「Rust 桥接」两列。
**公平性修正**：w100/w1000 使用各引擎官方批量 API（betteroffice `apply_ops` 批量、
rxls `set_cell_range_values`、umya/openpyxl 自然循环）——逐格调用会把批量场景
降级为每格一次原子提交，不属于引擎能力差异。

## 1. 单文件（1k 行 × 20 列，5 次重复取中位）

| 引擎 | w1(1格) | w100 | w1000 | w-multi(5表) | w-formula | w-mixed | 峰值内存 |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| **umya** | 170 | 199 | 193 | 223 | 198 | 201 | 9.8 MB |
| **rxls** | **167** | **340** | **346** | **461** | 747 | **683** | 45–62 MB |
| betteroffice | 330 | 699 | 3086 | 437 | 890 | 1284 | 37–39 MB |
| openpyxl 原生 | 781 | 974 | 931 | — | 1299 | 943 | 26.9 MB |
| openpyxl 桥接 | — | — | — | — | — | — | （见批量） |

（w-multi 需 5 表夹具，openpyxl 原生未单独跑该负载，以 w-mixed 为参照。）

**相对 openpyxl 原生倍数（total 中位）**：

| 引擎 | w1 | w100 | w1000 |
| --- | ---: | ---: | ---: |
| umya | 4.6x | 4.9x | 4.8x |
| rxls | 4.7x | **2.9x** | **2.7x** |
| betteroffice | 2.4x | 1.1x | 0.3x |

## 2. 规模梯度（w1 单格修改，3 次重复取中位）

| 规模 | umya | rxls | betteroffice | openpyxl |
| --- | ---: | ---: | ---: | ---: |
| 1k | 170 | 167 | 330 | 781 |
| 10k | 2559 | **4318** | 7524 | 9328 |
| 50k | **11028** | **11601** | 30462 | 42774 |
| 100k | **52997** | Unsupported* | 62771 | **144047** |

\* rxls 100k：`invalid ZIP package: OOXML package entry is too large`——11MB sheet XML
超过其单部件上限，**打开即拒绝**（计入能力限制，不计性能失败）。

**峰值内存（w1）**：umya 9.8→859.6MB（随规模线性增长且 100k 巨大）、
openpyxl 26.9→997.1MB、betteroffice 37→3610MB（内存最重）、
rxls 45→1705MB（50k）/打开超限（100k）。

## 3. 批量 10/50/100 工作簿（1k 行，w100 负载，3 次重复取中位）

| 引擎 | 模式 | 10 簿墙钟 | 50 簿 | 100 簿 | 每簿中位 | 峰值 |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| umya | 常驻 | 4325 | 17551 | **27437** | 274 | 87.8 MB |
| rxls | 常驻 | 5261 | 20114 | 35639 | 356 | 91.0 MB |
| betteroffice | 常驻 | 16704 | 88709 | 115994 | 1159 | 87.8 MB |
| openpyxl 原生 | 冷启动 | 24729 | 153414 | 231567 | 1250 | 87.8 MB |
| openpyxl 原生 | 常驻 | 13624 | 83948 | 134353 | 1272 | 103.0 MB |
| openpyxl 桥接 | 冷启动 | （生产 worker 单请求协议无常驻对应物，见批量原始数据缺失说明） | | | | |

**批量相对倍数（100 簿常驻/冷启动墙钟 vs openpyxl 对应模式）**：
umya 常驻 4.9x、rxls 常驻 3.9x、betteroffice 常驻 1.15x（慢）。

## 4. 数据解读

1. **umya 是小/中型文件批量修改的吞吐冠军**（每簿 274–430ms），内存也最低（常驻 100 簿 87.8MB
   含峰值累计）。其保真缺陷（批注/外链丢失）是独立问题——性能与保真在本报告分开呈现。
2. **rxls 在单格/小批量场景与 umya 同档**（w1 167ms 最快单文件），批量矩形写入走官方
   `set_cell_range_values` 后达到 w100 340ms / w1000 346ms，也属第一梯队。
3. **betteroffice 的 modify 阶段随格数超线性增长**（w1000 3.1s，比 umya 慢 16 倍）：
   `apply_ops` 每步做公式图重建 + 样式派生，是其设计成本。
4. **openpyxl 的总耗时被 Python 启动/导入（约 250–330ms）与保存（约 350–400ms）主导**，
   修改阶段本身极快（0.1–6ms）——桥接模式的固定成本即在此。
5. 100k 行场景只有 umya（53s/860MB）与 openpyxl（144s/997MB）能完成；
   betteroffice 62.8s/3.6GB 内存代价过高；rxls 打开即超限。
6. 输出文件大小：umya 92–94KB > rxls/betteroffice 76–80KB ≈ openpyxl 76KB（同内容）。

## 5. 口径与限制

- 桥接模式无法分段计时（生产协议不回传阶段耗时），报告为总耗时；
  原生与桥接的对比只用于「生产形态」参照，不与 Rust 引擎直接比出优劣。
- perf-batch 的 openpyxl-bridge resident 无对应物（生产 worker 为单请求协议），跳过。
- 100k × w100/w1000 组合因单次成本（betteroffice 预计 >60s/rep、openpyxl 预计 >300s/rep）
  仅做 w1 试运行，未计入完整矩阵（已按要求标注，不冒充完整结果）。
- 冷启动/常驻的墙钟含进程创建与文件拷贝，per-file 中位数为主要对比口径。
