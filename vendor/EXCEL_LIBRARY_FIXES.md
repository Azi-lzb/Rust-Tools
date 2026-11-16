# 定点 Excel 库修补（2026-11-14）

来源为本机缓存的原始 crates.io 包，解包逐文件比较后修补，不改全局 Cargo 缓存。两个库均保留 MIT 许可证、上游版权、版本及其他原始源码。

- calamine 0.37.0：上游 https://github.com/tafia/calamine，包 SHA256 `c9ca341b94bc4b29e08970b1a03914fad2dfe177e2ecb67fcd6de9046004aafb`。
- umya-spreadsheet 3.1.1：上游 https://github.com/MathNya/umya-spreadsheet，包 SHA256 `032a20340d2eac6e2e6c940e6b793cd3b68223d5f9412724cb7e32ef17c584d0`。

根 Cargo.toml 与独立基准适配器的 `[patch.crates-io]` 指向这两个固定副本；没有升级库、改默认引擎、改生产路由或引入 Aspose 运行时。根 Cargo.lock 只将这两个同版本包改为 path 来源。dist 未构建，旧 EXE 不包含本修补。

## 修改范围

`calamine-0.37.0/src/xls.rs`：

1. BIFF8 PtgRef3d/PtgArea3d 经 ExternSheet XTI 定位实际 Sheet，不能把 ixti 当工作表数组下标。
2. 列号掩码为低 14 位，bit 14 是列相对标志，bit 15 是行相对标志；修正 3D 及普通 BIFF8 引用的混合绝对/相对标志。
3. 3D 引用特殊 Sheet 名加引号并转义单引号，引用操作数长度不足时明确失败。
4. 修正二元比较操作码 0x0C（>=）与 0x0D（>）的反置。
5. BIFF8 普通引用的行号提升 u32 后加 1，避免合法第 65536 行的 u16 溢出。

不新增公式计算；不声称覆盖 BIFF2-5、外部链接、共享/数组公式、定义名称或所有 3D 引用语法。原有 BIFF2-5 引用分支保留。实验 XLS 公式转换仍明确拒绝，不因解析修复自动开放写入。

`umya-spreadsheet-3.1.1/src/structs/data_validation.rs`：

1. 只在两个公式都为空时写空节点；有任一公式均写子节点，保留 between/notBetween 的上下限以及单公式规则。
2. 同一最小夹具的严格 XML 回归揭示 prompt 中 & 被二次转义。仅对 DataValidation 的四个自由文本属性在读取时解码一次 XML 实体，写出仍由原写入器转义；不改全局 XML 驱动。

## 证据与复现

证据位于 `benchmark/excel-engine-evaluation/baselines/results/root-cause-20261114/`。
上游未修补程序保存在 `F:/RustBuild/pytools/root-cause-upstream.exe`，上下限内存读取值和写出后值分别记录，原始 XML/BIFF 与 Aspose 独立解析结果均保留。

两种 XLS 生成器（xlwt、Aspose）的 44 条公式与同内容 XLSX 控制组，7 种数据验证组合，从同一输入分别运行未修补和修补后的直接库 API 程序。比较只允许去掉简单 Sheet 名冗余引号，不忽略 $、Sheet 身份、坐标或运算符。保存后直接解析 ZIP XML，不能用库内存对象正确代替写出正确。

必要命令和最终结果见该目录 `REPORT.md`。原速度表使用未修补二进制，历史计时与原件未改写，不能将其自动标作修补后的性能成绩。后续上游版本升级须重新核实并移除已被覆盖的补丁，不能盲目叠加。

规范依据：
- https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/b8bc555a-09b0-4b22-835f-20e2028c814a （XTI 与 3D 引用）
- https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/bccb00ac-67a1-490a-9f0b-5ce59978eb9a （ColRelU 相对位）
