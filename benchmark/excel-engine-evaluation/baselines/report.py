"""Render only observed data. No guessed or cross-platform performance values."""
import argparse
import json
import statistics
from collections import Counter
from pathlib import Path


def formula_parser_note(capability):
    status=capability['formula'].get('formula_text_diagnostic',{}).get('status','NOT_TESTED')
    if status=='UNEXPECTED_PARSER_DIFFERENCE':
        return '- 另有独立的 UNEXPECTED_PARSER_DIFFERENCE：calamine 对该 XLS 的跨表公式解码与已确认源公式不符；它不能归入预期计算缺口。该实验拒绝 XLS 公式转换，原期望与解析文本保存在能力报告。'
    if status=='MATCH':
        return '- 公式文本解析与本夹具期望匹配（MATCH）；实验仍明确拒绝 XLS 公式转换、不实现计算，不能把解析通过当作转换或重算已启用。'
    return f'- XLS 公式文本解析证据：{status}；不推断已通过。'


def umya_validation_note(capability):
    cases=[x for x in capability['edits'] if x['engine']=='umya']
    if not cases: return '- UMYA：NOT_TESTED。'
    lost=any(d.get('classification')=='DATA_VALIDATION_BOUNDS_LOST' for x in cases for d in x['semantic_unexpected'])
    if lost: return '- UMYA：数据验证的 formula1/formula2 上下限丢失，属于实际规则变化；另有列宽补写和等价的 Sheet 引用引号变化。'
    return '- UMYA：本批夹具未发现验证上下限丢失；其他差异按实际分类保留，不能据此宣称完整保真。'


def run(root, capabilities):
    root=Path(root); cap=Path(capabilities)
    rows=json.loads((root/'summary.json').read_text('utf-8'))
    records=[json.loads(line) for line in (root/'rounds.jsonl').read_text('utf-8').splitlines()]
    cold=json.loads((root/'cold-start.json').read_text('utf-8'))
    env=json.loads((root/'environment.json').read_text('utf-8'))
    host=json.loads((root.parents[1]/'host.json').read_text('utf-8-sig'))
    capability=json.loads((cap/'report.json').read_text('utf-8'))
    assert len(rows)==18 and all(r['n']==r['total']==10 and r['status']=='PASS' for r in rows)
    assert len(records)==234 and all(r['ok'] and r['validation']['ok'] for r in records)
    assert len(list((root/'raw').glob('aspose-*.xlsx')))==52
    assert max(r.get('workbook_count',0) for r in records)==78
    lines=['# Windows 独立 Excel 引擎实测（2026-11-14）','',
           '基线：main `34dfcd905b1157dca531908083b156b41f440a03`。仅独立探针，生产路由、配置及 dist 未改。','',
           f"主机：{host['cpu'].strip()}，{host['logical_processors']} 逻辑处理器，Windows {host['version']}，约 16 GiB 内存；输入/输出在 C: NTFS，构建产物在 F:。后台负载保留，不是冷磁盘或裸机基准。",'',
           '计时：3 轮实际预热 + 10 轮正式测量，串行轮换引擎顺序。下表为进程内处理时间（毫秒）的中位数 [最小–最大]，不含进程启动、矩阵预载、摘要、IPC及独立回读校验；包含实际文件 I/O。全部慢轮保留。','',
           '样本为无公式混合标量（整数、文本/前导零、布尔、空值）。']
    for f in env['fixtures']:
        lines += [f"- {f['id']} 档：{f['bytes']:,} 字节（{f['bytes']/1048576:.4f} MiB），{f['rows']:,} 行 × {f['cols']} 列，{f['nonempty_cells']:,} 个非空单元格，{f['positions']:,} 个含空值位置；SHA256 `{f['sha256']}`。"]
    labels=dict(rust='Rust 原生',openpyxl='openpyxl 3.1.5',aspose='Aspose.Cells 26.10.0 评估版')
    for op,title in [('read','完整读取'),('write','从零新建写入'),('roundtrip','独立实测往返')]:
        lines += ['',f'## {title}','','| 引擎 / 模式 | 1 MiB 档 ms | 10 MiB 档 ms |', '| --- | ---: | ---: |']
        for engine in ('rust','openpyxl','aspose'):
            name=labels[engine]
            if engine=='rust': name={'read':'calamine 0.37.0','write':'rust_xlsxwriter 0.101.0','roundtrip':'calamine + rust_xlsxwriter（值重建）'}[op]
            if engine=='openpyxl': name+=' / '+{'read':'read_only 全消费','write':'write_only','roundtrip':'normal 加载/保存'}[op]
            cells=[]
            for size in ('1MiB','10MiB'):
                r=next(r for r in rows if r['size']==size and r['operation']==op and r['engine']==engine)
                cells.append(f"{r['median_ms']:.2f} [{r['min_ms']:.2f}–{r['max_ms']:.2f}]")
            lines += ['| '+name+' | '+' | '.join(cells)+' |']
    lines += ['', 'Rust 往返是纯值重建，不保留样式/结构；Aspose/openpyxl 往返是加载后保存，工作量不同，不能用该表推断保真能力。calamine 不提供写入，rust_xlsxwriter 不提供读取。原生 openpyxl 结果不等同生产安全 Worker。', '', '## 启动与首次调用','']
    for engine in ('rust','openpyxl','aspose'):
        vals=[r['launch_to_ready_ms'] for r in cold if r['engine']==engine]
        first=next(r for r in records if r['engine']==engine)
        lines += [f"- {engine}：新进程启动至 READY，3 次中位数 {statistics.median(vals):.2f} ms，范围 {min(vals):.2f}–{max(vals):.2f} ms；首次 1 MiB 读取 API {first['processing_ms']:.2f} ms（预热轮）。矩阵预载单列 {env[engine]['preload']['request_e2e_ms']:.2f} ms，不计主表。"]
    lines += ['', '驻留 API 计时直接排除启动，不做 HelloWorld 的不可靠常数相减。READY 不证明 JIT 全部预热；正式轮的分层优化/GC差异留在范围中。Rust 初始 JSON 预载使用非缓冲 File 解析，约 30 秒，是独立夹具预载开销；没有计入 Excel I/O。不得把该预载当成 calamine 读取时间或隐瞒为启动成本。','',
              '## 校验与评估输出','',
              '- 234/234 请求成功并通过校验：54 预热 + 180 正式记录，18 组各 10 个正式样本。全部 156 个保存结果均独立完整回读；读取结果完整物化后核对类型和值摘要。',
              '- Aspose 性能进程共 78 个 Workbook 请求，52 个原始评估输出全部保留。仅按原输入 Sheet 身份及供应商标记排除新增 Evaluation Warning 的比较/计数视图。有效数据量没有评估 Sheet。实际输出字节及水印标记另存每轮记录。',
              '- 保存耗时包含评估版生成水印的实际开销，不估算扣除、不推断授权版速度。没有修改评估输出、许可证或 SDK 限制。',
              '- 两个输入的前后 SHA256 相同，见 input-integrity.json。没有用真实业务文件。','',
              '## 公式降级与修改能力','',
              f"- XLS 合成夹具的 4 个公式具有经 BIFF 字节确认的空缓存；Rust 明确 UNSUPPORTED，无输出且输入不变。Aspose 显式 CalculateFormula 后缓存为 {capability['formula']['aspose_actual']}，公式仍保留：EXPECTED_CAPABILITY_GAP。未声称 Aspose 打开/保存必然自动计算，也没有给 Rust 添加计算器。",
              formula_parser_note(capability), '',
              '| 引擎 | 操作 | 观察到的差异条数（含等价编码） | 原包定点部件不变条件 |','| --- | --- | ---: | --- |']
    for r in capability['edits']:
        lines += [f"| {r['engine']} | {r['operation']} | {len(r['semantic_unexpected'])} | {'满足本夹具条件' if r['precise_part_patch'] else '不满足'} |"]
    lines += ['', '差异分类与结论：', '',
              '- openpyxl：本夹具选定语义均符合预期，但不透明 customXml 被丢弃。',
              umya_validation_note(capability),
              '- Aspose：默认行高显式化、命名引用引号、RGB alpha、验证默认 between 编码、空公式缓存类型编码等变化不应全叫作数据损坏。此早版夹具的包级未知 customXml 部件未保留，但它不能证明合法 Workbook 自定义 XML 丢失；正确关系夹具的后续四维实验中 Aspose 保留通过。评估 Sheet 过滤后仍有整簿重写。',
              '- Aspose 相对自身 no-op，过滤评估 Sheet 后的单值修改仅改变对应工作表，样式修改改变工作表和 styles；这描述其序列化基线，不能证明原始包精确补丁能力。',
              '- ooxml-patch 页眉替换：只有目标 worksheet 部件变化，其他全部未压缩部件哈希不变，语义检查通过。该结论仅针对本窄操作和夹具。','',
              '原始部件增删和未压缩 SHA256、规范化 XML差异、评估过滤视图以及 Aspose 相对自身 no-op 的原始/过滤差异见能力 report.json。语义项包括值/类型/公式、缓存编码、合并、批注、链接、数值格式、加粗、填色、行列尺寸、命名区域、打印和页眉页脚/数据验证；不是覆盖所有 Excel 对象的通用保真认证。不透明 customXml 部件的丢失不可归因于水印。ooxml-patch 仅做它实际支持的页眉替换，不伪造通用单元格修改能力。','',
              '## 未验证与边界','',
              '- NOT_TESTED：Linux/WSL/Docker、麒麟、Windows Excel/WPS COM、XLS 大文件速度、XLSM、真实业务样本、授权版 Aspose、并发/峰值内存、冷磁盘及复杂公式全集。',
              '- Rust 降级为保守实验口径：所有 XLS 公式拒绝；OOXML 空/歧义/非数值公式缓存拒绝，OOXML 数值缓存保留分支本轮未单独验收。旧 rxls 实验命令未改，不能将新适配器的保护推广到旧命令；不代表生产 COM 不支持 XLS 追加或计算。',
              '- 本轮不是主项目全量回归或发布验收，不更新 dist、不提交、不 push。','']
    (root/'REPORT.md').write_text('\n'.join(lines),'utf-8')
    (root/'evidence-audit.json').write_text(json.dumps(dict(requests=len(records),phase_counts=dict(Counter(r['phase'] for r in records)),
            groups=len(rows),all_groups_pass=True,raw_saved_outputs=156,aspose_raw_outputs=52,aspose_workbook_count=78),indent=2),'utf-8')
    print(root/'REPORT.md')


if __name__=='__main__':
    p=argparse.ArgumentParser(); p.add_argument('results');p.add_argument('capabilities');a=p.parse_args();run(a.results,a.capabilities)
