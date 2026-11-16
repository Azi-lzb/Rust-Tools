# 合成 BIFF8 公式夹具

`xls_formula_refs.xls` 由本仓库 `benchmark/excel-engine-evaluation/baselines/root_cause_fixtures.py` 使用 xlwt 1.3.0 生成，包含 5 个合成 Sheet 和 44 个公式，不含真实业务数据。

`xls_formula_refs.json` 为生成规格与严格期望：Sheet/XTI 映射、相对/绝对引用、包含空格和单引号的 Sheet 名、Z/AA 列、9/10/99/100 行和六种比较运算。源 XLS 与独立 Aspose 生成的对应 XLS 都经 Aspose 26.10.0 读取核对；XLSX 控制组读取同样通过。

该文件中的公式缓存不用于计算验收。本测试只验证公式文本读取，不能据此宣称有公式计算或保真 XLS 转换能力。

重建：在一个新的隔离目录运行上面的 Python 生成器，再将生成的 `xlwt-references.xls` 和 `spec.json` 作为本夹具及 JSON 规格；不要覆盖未经核对的业务文件。原始对拍记录保存在基准 `results/root-cause-20261114/`。
