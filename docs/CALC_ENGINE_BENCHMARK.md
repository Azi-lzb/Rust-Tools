# 计算引擎对比：IronCalc vs betteroffice vs Excel COM vs LibreOffice

日期：2026-11-08
探针位置：`C:\Users\AZI\AppData\Local\Temp\engine-bench`（临时，未进仓库）
全部样本**先清空公式缓存值**（去掉每个公式格的 `<v>`，并置 `fullCalcOnLoad="1"`），
确保没有引擎能靠文件里的旧结果蒙混过关。

## 0. 一句话结论

**IronCalc 不能用在业务管线上**——它算得不错（99.81%），但落盘文件会被 **Excel 直接拒绝打开**，而且会**销毁全部批注**。
betteroffice 包保真最好，但大簿预算耗尽后大面积 `#NUM!`（仅 3% 正确）。
LibreOffice 结果最准（99.91%），但仍需重写样式表。Excel COM 是唯一结果与文件双保真的基准。

## 1. 速度对比

单位：秒。样本按公式量递增。所有数字为本机单次实测（Windows 11，i7 级 CPU，SSD）。

### 1.1 总耗时（打开 + 计算 + 保存）

| 样本 | 公式格 | 文件大小 | IronCalc | betteroffice | Excel COM | LibreOffice |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| battery（合成） | 51 | 7 KB | 0.007 | 0.008 | 1.00 | 3.52 ¹ |
| mubiao（目标用例） | 175 | 40 KB | 0.014 | 0.047 | 1.39 | 6.75 ¹ |
| guangda（光大银行月报） | 614 | 99 KB | 0.050 | 0.430 | 1.54 | 2.84 ¹ |
| shengxindai（省信贷，7.1 MB） | 73807 | 6.8 MB | **148.7** | 14.28 ² | **8.11** | 22.69 ¹ |

¹ LibreOffice 含 soffice 进程启动（1.4~4.8s）。去掉启动只算 load+calc+save 分别是 0.70 / 2.30 / 1.43 / 17.87 秒。
² betteroffice 此处的"完成"是假的：预算耗尽，71095 个格没算。

### 1.2 纯计算阶段

| 样本 | IronCalc | betteroffice | Excel COM | LibreOffice |
| --- | ---: | ---: | ---: | ---: |
| battery | 0.001 | 0.001 | 0.01 | 0.002 |
| mubiao | 0.000 | 0.001 | 0.04 | 0.004 |
| guangda | 0.001 | 0.005 | 0.49 | 0.005 |
| shengxindai | **146.05** | 3.92 ² | **2.51** | 2.24 |

**纯 Rust 引擎在计算阶段并不快**：
- IronCalc 在大簿上 146 秒，是 Excel COM 的 **58 倍慢**、LibreOffice 的 **65 倍慢**。
- betteroffice 3.92 秒接近 COM/LibreOffice，但那个数字是"算不完就放弃"的结果。

### 1.3 速度结论

- 小簿（≤1000 公式）：IronCalc / betteroffice 微秒级，碾压 COM（1~1.5s 主要是 Excel 进程启动）。
- 大簿（7 万公式）：**Excel COM 最快（8.11s）**，LibreOffice 次之（22.69s），IronCalc 最慢（148.7s）。
- 纯 Rust 的速度优势只在"进程启动"上，一旦公式量上来，解释执行的开销立刻反超。

## 2. 计算结果对齐（基准 = Excel COM）

### 2.1 总体一致率

| 样本 | 公式格 | IronCalc | betteroffice | LibreOffice |
| --- | ---: | ---: | ---: | ---: |
| battery | 51 | 42/51 (82.35%) | 49/51 (96.08%) | 49/51 (96.08%) |
| mubiao | 175 | **175/175 (100%)** | **175/175 (100%)** | **175/175 (100%)** |
| guangda | 614 | 604/614 (98.37%) | 604/614 (98.37%) | **614/614 (100%)** |
| shengxindai | 73807 | 73664/73807 (99.81%) | 2255/73807 (**3.06%**) | 73744/73807 (**99.91%**) |

### 2.2 差异归因

**battery（合成打靶样本）**

| 引擎 | 差异 | 原因 |
| --- | ---: | --- |
| IronCalc | 9 | `RAND`/`RANDBETWEEN` 随机（3）；`NOW` 微差（2）；**定义名 4 个全 `#ERROR!`**；`SUMIF(B2:B6,">=C",...)` 算成 60（正确 120）——文本比较语义错 |
| betteroffice | 2 | 仅 `RAND`/`RANDBETWEEN` 随机 |
| LibreOffice | 2 | 仅 `RAND`/`RANDBETWEEN` 随机 |

> 排除天然随机的 `RAND`/`RANDBETWEEN` 后：**betteroffice 49/49、LibreOffice 49/49、IronCalc 42/47**。

**guangda（含悬空外链 `[1]`）**

| 引擎 | 差异 | 原因 |
| --- | ---: | --- |
| IronCalc | 10 | 全部源于 `C7` 公式 `'[1]A1411_2023-5421-20260630'!$E$660`——外链文件不存在，IronCalc 解析为 `#REF!` 并级联 8 格 |
| betteroffice | 10 | 同上（一模一样的 10 格） |
| LibreOffice | **0** | 与 Excel 完全一致 |

> Excel 保留该外链的**缓存值**并算出结果。两个纯 Rust 引擎都无法复现——它们不做外链缓存 fallback。
> 注意：这正是我们 4-7/1-4 修过的 `[N]` 场景，业务上应被 fail-closed 拦截，而非引擎能力问题。

**shengxindai（7 万公式大簿）**

| 引擎 | 差异 | 原因 |
| --- | ---: | --- |
| IronCalc | 143 | 78 格是 `=RANK(D14,(D$5:D$24,D$26))`——**逗号联合区间语法不支持**；65 格是错误码语义差（`#REF!`→`#NA`、`#N/A`→`#NUM!`） |
| betteroffice | 71552 | **预算耗尽**：63087 格 `#NUM!`、7878 格 `IFERROR` 回退成空、574 格算成 0 |
| LibreOffice | 63 | 全是错误码语义差（`#N/A` vs `#VALUE!`），无一处数值错误 |

## 3. IronCalc 的三个致命问题

### 3.1 输出文件被 Excel 拒绝打开 ⛔

**IronCalc 落盘的文件，Excel COM `Workbooks.Open` 报 `0x80020009`（"类 Workbooks 的 Open 方法无效"）。**

根因已定位并实验证实：IronCalc 把内部错误写进 `<v>` 标签，产生 **`<v>#ERROR!</v>`**。
OOXML 规范只允许 7 种错误码（`#NULL!` `#DIV/0!` `#VALUE!` `#REF!` `#NAME?` `#NUM!` `#N/A`），
`#ERROR!` 是 IronCalc 内部枚举名，不是合法 Excel 错误。

验证实验：

```
把 ic-battery.xlsx 里 4 处 "#ERROR!" 字面替换为 "#VALUE!"
→ Excel COM 立刻打开成功，3 张表
```

只要工作表里**任何一格**触发了 IronCalc 的内部错误（定义名、`RANK` 联合区间等），
整个文件就变成 Excel 打不开的废文件。

### 3.2 销毁全部批注 ⛔

| 样本 | 原件批注 | IronCalc 输出 | betteroffice 输出 |
| --- | ---: | ---: | ---: |
| guangda | 12 条（2 个部件 + 2 个 VML） | **0 条** | 12 条 |
| shengxindai | 9 条（5 个部件 + 5 个 VML） | **0 条** | 9 条 |

IronCalc 还丢失：`docProps/custom.xml`、`xl/externalLinks/*`。
它是**整簿重建**（11 部件 → 19 部件，含非法目录条目），不是包保真补丁。

### 3.3 定义名与联合区间不支持

- 定义名在 `Model::evaluate()` 阶段全部返回 `#ERROR!`（`SUM(单价区域)`、`INDEX(单价区域,2)`、`汇总值`、`合计` 4/4 失败）。
- 注意：定义名**能正确写出到 `workbook.xml`**，是求值不支持。
- `RANK(D14,(D$5:D$24,D$26))` 这种逗号联合区间语法 → `#ERROR!`。

## 4. 四引擎总评

| 维度 | IronCalc | betteroffice | Excel COM | LibreOffice |
| --- | --- | --- | --- | --- |
| 计算结果准确 | 中（99.81%，定义名/联合区间不支持） | 小簿优 / 大簿废 | **基准** | **优（99.91%）** |
| 大簿可行性 | 慢（148s） | **不可用（预算耗尽）** | **优（8s）** | 良（23s） |
| 输出 Excel 可打开 | **否（非法错误码）** | 是 | 是 | 是 |
| 批注保真 | **否（全丢）** | **是（全留）** | 是 | 是 |
| 包结构保真 | 差（重建，丢部件） | **优（仅补丁改动格）** | 优 | 差（重写样式/主题） |
| 公式文本保真 | 未测（重建） | **73807 条逐字零变化** | 是 | 会重写 |
| 外部依赖 | 无 | 无 | 需 Excel | 需 soffice + python3-uno |
| 许可证 | MIT OR Apache-2.0 | Apache-2.0 | 商业 | MPL 2.0 |
| 离线构建 | 是 | 是 | — | — |
| 进程启动开销 | 无 | 无 | ~0.9s | ~1.4~4.8s |

## 5. 各引擎的适用面

**IronCalc**：只能用于"纯计算、不落盘"的场景，且公式集必须落在它支持的子集内。
一旦要保存成 Excel 文件，就必须拦截并修正 `#ERROR!`（不能简单替换成 `#VALUE!`——那会把"引擎不支持"伪装成"业务出错"），
还要接受批注全丢。当前不建议接入。

**betteroffice**：包保真和公式文本保真是全场最好，小簿（1-4 的实际负载，614 格）100% 正确且速度极快。
但大簿会预算耗尽并**大面积返回 `#NUM!` 或 0**，且预算常量写死不可调。可用，但必须通过 `budget_exhausted` 硬拦截。

**LibreOffice**：准确率最高（大簿 99.91%，唯一差异是错误码语义），跨平台，但要重写样式/主题、
新增部件、`sharedStrings` 膨胀 9 倍（30KB→285KB）。已在项目里跑通（`libreoffice_calc_worker.py`）。

**Excel COM**：最快的大簿引擎，结果与文件双保真。但依赖 Excel/WPS 安装、有进程启动开销、不适合批量小文件。

## 6. 未验证项（如实记录）

- IronCalc 只测了 `0.9.1`；未测 git 主分支是否已修 `#ERROR!` 问题。
- 未验证 `.xlsm`、图表、条件格式、数据验证的重算保真（只验证了批注和包结构）。
- 未做多轮重复计时（单次实测，未取中位数；进程启动项波动较大）。
- 未在 Linux 实机跑 IronCalc / betteroffice（仅确认无 C 依赖、可交叉编译）。
- `shengxindai` 的 LibreOffice 63 处错误码差异未逐格确认业务等价。
- 未验证 IronCalc 输出的 `#ERROR!` 是否在其他版本/配置下不出现（例如全部公式都在支持子集内时）。

## 7. 建议

按当前证据，**没有纯 Rust 引擎可以无条件替换 COM/LibreOffice**：

1. **IronCalc：暂不接入。** 输出文件 Excel 打不开是硬伤，且丢批注。
   若要评估，需先确认上游是否修了 `#ERROR!` 序列化和定义名求值。
2. **betteroffice：可以给 1-4 用**（小簿场景已验证 100% + 极快），但必须：
   - `budget_exhausted` / `limited_cells` 非空 → fail-closed 回退；
   - `now_serial` 显式注入；
   - 重算后校验区域不得出现 `#NAME?` / `#NUM!`；
   - 保留 COM/LO 作为回退路径。
3. **LibreOffice**：维持现状（Linux 1-4 主路由），准确率最高。
4. **Excel COM**：保留为 Windows 主力与大簿首选。

需要我继续做哪一步？（例如：给 betteroffice 写正式的预算拦截接入、或深入 IronCalc 上游 issue 跟踪）
