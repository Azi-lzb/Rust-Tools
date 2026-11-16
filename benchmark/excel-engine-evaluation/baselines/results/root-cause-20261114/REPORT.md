# XLS 公式与 UMYA 数据验证根因核实及修复

> 后续更正（2026-11-14）：本报告对 calamine 公式解码与 UMYA 双上下限缺陷的独立核实和修复结论不变。但附带的早版 customXml 哨兵只是包级未知部件，不能证明合法 Workbook 自定义 XML 丢失。后续四维实验已修正关系和属性，Aspose 对合法夹具保留通过；原证据保留，不再沿用早版归因。

日期：2026-11-14。工作区：Rust-Tools main，HEAD `34dfcd905b1157dca531908083b156b41f440a03`。本轮没有提交、push、合并、更新 dist 或改生产/Linux 路由。

## 结论

两项差异均已独立复现为固定版本库的缺陷，既不是原夹具非法，也不是比较脚本制造的差异。

- calamine 0.37.0：xlwt、Aspose 两种独立 XLS 生成器，各含 44 条严格公式期望。未修补程序分别不匹配 41 条和 40 条；同内容 XLSX 控制组不匹配 0 条。Aspose 直接读取两个 XLS 和 XLSX 的全部 44 条都匹配期望。原始 BIFF token、XTI 与源公式保存，未用“忽略 $ / Sheet / 运算符”来消除差异。
- UMYA 3.1.1：直接 reader→writer→reader，完全不经过旧修改探针。源 XML 与 reader 内存都正确含 formula1="0"、formula2="1000"，writer 后两者消失。另两组负数、小数双上下限同样复现；单公式规则不丢，定位到 writer 的空节点判断。
- 原始 `.crate` SHA256 与此前 Cargo.lock 一致，两处相关全局缓存源码逐字节与原包相同。全局缓存没有被本轮改写。

证据：`crate-provenance.json`、`global-cache-audit.json`、`aspose-independent.jsonl`、`xlwt-biff-records.json`、`aspose-biff-records.json`、`upstream-report.json`。保留未修补 EXE `F:/RustBuild/pytools/root-cause-upstream.exe` 及原始输出。供应商评估 Sheet/标记未删除；比较通过原 Sheet 名取有效公式。

## 实际修复

以原始包为基础引入两个固定版本 vendor 副本并保留 MIT 版权和许可证。751 个原包文件中仅两个源码文件有改动，见 `vendor-source-audit.json` 与上游/修补差异文件。

1. Calamine `src/xls.rs`：BIFF8 的 3D 引用经 XTI 正确定位 Sheet；列号掩码与行列相对标志分离；普通引用也修正混合 $ 标志；特殊 Sheet 名正确转义；修正 > 与 >= 操作码反置。BIFF8 引用行号用 u32 表达，避免第 65536 行加一溢出；该极限行及 BIFF2-5 未单独验收。没有添加公式计算器。
2. UMYA `src/structs/data_validation.rs`：存在任一公式就输出子节点，只有两者均空才输出空节点。严格 XML 字段比较还发现提示文字的 & 被重复转义，因此仅对同一个 DataValidation 对象的 errorTitle/error/promptTitle/prompt 属性在读取时解码一次；不改全局 XML driver。
3. 根及独立探针 Cargo 清单使用同一副本，版本号不变。根 Cargo.lock 只去掉这两个包的 registry source/checksum，其他包版本未变。没有切换引擎路由。
4. 报告生成器原先含有针对首次结果写死的失败描述；现改为根据实际解析诊断和差异分类生成，避免修复后仍误报旧失败。旧报告和原始失败证据保留。
5. 实验 XLS 公式转换仍为 UNSUPPORTED，只更新说明为“有意不开放”，没有因公式文本解析通过就开放 XLS 写入或计算。生产 COM 及其他路径未调整。

## 针对性回归结果

以下命令均使用 `CARGO_TARGET_DIR=F:\RustBuild\pytools`；没有全量测试。

- `cargo test --offline -p excel-reader --test xls_formula_references --test formula_range_empty_cells --test unknown_error_tokens`：3/3 通过，含新增主项目合成 XLS 夹具和既有两个控制回归。
- `cargo test --offline -p excel-editor --test data_validation_roundtrip`：2/2 通过。使用 rust_xlsxwriter 作为第二种 XLSX 生成器，验证 0、负数、小数、between/notBetween、单公式、list/custom，以及四个自由文本属性的 XML 实体只解码一次。输入字节不变。
- `cargo test --offline --manifest-path benchmark/excel-engine-evaluation/baselines/adapters/rust/Cargo.toml --bin root_cause`：3/3 通过，验证两个 XLS 生成器和 XLSX 控制、7 组数据验证、拒绝已有输出和原件不变。
- `py -3.12 -m unittest discover -s benchmark/excel-engine-evaluation/baselines/tests -p test_root_cause.py -v`：6/6 通过。核对修补前后的直接 API 证据、12 个修改输出、4 个显式计算结果、旧报告不再误报、此前 209 个工作簿原件及配置/dist 哈希保护。
- 合计 14 个测试函数通过；12 个修改输出和 132 条公式观测是覆盖量，不重复计成测试函数数量。

直接前后对拍入口：

`py -3.12 benchmark/excel-engine-evaluation/baselines/root_cause_check.py --root benchmark/excel-engine-evaluation/baselines/results/root-cause-20261114 --rust F:/RustBuild/pytools/release/root_cause.exe --label patched --expect-fixed`

通过：两份 XLS 与 XLSX 均 44/44；7 组数据验证的已检查 XML 字段一致。这里是解析后的验证节点字段对比，不代表整张 Sheet 或 ZIP 字节不变。字段包含 sqref、类型、运算符、两个公式及提示文字。

格式检查：独立适配器 cargo fmt --check、新增两份主项目测试和 Calamine 修改文件 rustfmt --check、git diff --check 均通过。UMYA 仅保留上游 nightly 风格并定点改行，没有宣称整个第三方库或仓库格式检查通过。复用库仍有既有 unused 类警告，没有进行无关清理。

开发期曾出现最小程序 API 名错误及 xlwt Sheet 创建顺序错误，已修正后重建；旧源码与失败日志/终端记录保留。这些测试工具开发失败不用于证明库缺陷，库缺陷结论来自成功运行的直接 API 对拍。

## 计算和修改究竟测试了什么

之前已经测试，首份速度表仅展示 I/O，未把这两组正确性结果说明充分。本轮又复跑了它们，原始评估输出位于 `capabilities-postfix/`：

- 计算：4 个 BIFF 已确认空缓存的 XLS 公式（SUMIF、SUMPRODUCT、IF/AND/ROUND、INDEX/MATCH）。Aspose 显式调用 CalculateFormula，保存后由 openpyxl 独立回读，得到 8、68、6、5，并保留公式。Rust 明确拒绝，无输出且源文件不变；计算能力差异仍标 EXPECTED_CAPABILITY_GAP，而解析诊断已变为 MATCH。
- 修改：openpyxl 的 no-op/单值/加粗/页眉 4 个、UMYA 的 no-op/单值/加粗 3 个、Aspose 的 4 个，加 ooxml-patch 页眉窄操作 1 个，共 12 个输出。核对目标修改、其他单元格值/类型/公式、合并和验证上下限，保留全包部件差异及评估过滤视图。
- UMYA 本轮已无验证上下限丢失；剩余列宽补写/Sheet 引号变化继续如实记录。整簿保存丢失 customXml 的既有能力边界仍在，不能据此直接替代定点 OOXML 补丁。
- 尚未做计算与修改各自的 3 预热＋10 正式轮性能基准；之前的多轮速度表只覆盖读/新建写/往返。没有将正确性测试冒称这两类性能验收。

## 修改文件与此前修改的区分

此前已有：整个 `baselines/` 初版 I/O 适配器、首轮速度/能力结果及既有工作区未跟踪资料，均保留。

本轮：

- 根 Cargo.toml/Cargo.lock；新增 `vendor/calamine-0.37.0/`、`vendor/umya-spreadsheet-3.1.1/` 和 `vendor/EXCEL_LIBRARY_FIXES.md`。
- 新增主项目 `crates/excel-reader/tests/xls_formula_references.rs`、合成 XLS/JSON、说明和精确忽略例外；新增 `crates/excel-editor/tests/data_validation_roundtrip.rs`。
- 独立适配器 Cargo.toml/Cargo.lock、XLS 拒绝原因文字；新增 `src/bin/root_cause.rs`。
- 新增 C# `RootCause.cs` 和 Program.cs 的独立诊断入口；Aspose 仍只存在于外部探针工程。
- 新增 `root_cause_fixtures.py`、`root_cause_check.py`、`tests/test_root_cause.py`；更新 README、capabilities.py 的诊断说明和 report.py 的证据驱动描述。
- 本目录内保存本轮新证据。没有覆盖此前工作簿、首轮速度表或其他工作目录。

## 未验证与安全确认

NOT_TESTED：Linux/WSL/Docker/麒麟、Excel/WPS COM、BIFF2-5、外部链接、共享/数组公式、命名公式/所有公式语法、极限行列、授权版 Aspose、真实业务样本、修补后性能回归与完整主项目回归。所有新夹具均为合成，不使用业务原件做测试。

此前 209 个工作簿和 root config.xlsx、dist/config/config.xlsx、dist/RustTools.exe 前后 SHA256 一致，见 Python 保护回归。dist 旧 EXE 未更新，不能用于验证这次源码修补。HEAD 未变、暂存区无改动，没有提交、push、合并、打包、安装或卸载；没有操作 feature/test 工作目录。
