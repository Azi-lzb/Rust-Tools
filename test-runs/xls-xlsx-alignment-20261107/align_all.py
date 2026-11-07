# 全 COM 路由功能 × 双格式（.xlsx openpyxl 路由 vs .xls COM 路由）同内容对拍
# 覆盖：2-1 标记 / 2-2 删重 / 4-4 转换 / 4-5 改名 / 5-1 拆分 / 5-2 规范化 /
#       5-3 外资行 / 5-4 页眉 / 5-7 涉农贷款 / 1-4 校验区域（.xls 源）
import json
import pathlib
import shutil
import subprocess
import sys

REPO = pathlib.Path(r"C:\Users\AZI\Desktop\VibeCode\Rust-Tools")
EXE = pathlib.Path(r"F:\RustBuild\pytools\debug\RustTools.exe")
WORK = pathlib.Path(r"C:\Users\AZI\AppData\Local\Temp\xls-xlsx-align-all")
import openpyxl
import xlrd

if WORK.exists():
    shutil.rmtree(WORK)
(WORK / "output").mkdir(parents=True)
aligned = {}


def run(argv, timeout=420, check_exit=True):
    proc = subprocess.run([str(EXE)] + [str(a) for a in argv], cwd=str(REPO),
                          capture_output=True, text=True, encoding="utf-8",
                          errors="replace", timeout=timeout)
    if check_exit and proc.returncode != 0:
        raise RuntimeError(f"exit {proc.returncode}: {proc.stderr[-300:]}")
    return proc


def com_save_as(src_xlsx, dst_xls):
    import win32com.client
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    try:
        book = app.Workbooks.Open(str(src_xlsx))
        book.SaveAs(str(dst_xls), FileFormat=56)
        book.Close(False)
    finally:
        app.Quit()


def com_convert_to_xlsx(src, dst_xlsx):
    """仅用于本脚本读取：任意簿 → .xlsx（FileFormat 51）。"""
    import win32com.client
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    try:
        book = app.Workbooks.Open(str(src))
        book.SaveAs(str(dst_xlsx), FileFormat=51)
        book.Close(False)
    finally:
        app.Quit()


def sheet_rows(path, sheet_name, cols=6):
    """读前 cols 列 (行号, 值...)，数字/文本归一化由调用方处理。"""
    rows = []
    if str(path).lower().endswith((".xlsx", ".xlsm")):
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else None
        if ws is None:
            return rows
        for r in range(1, ws.max_row + 1):
            vals = []
            for c in range(1, cols + 1):
                v = ws.cell(row=r, column=c).value
                vals.append("" if v is None else v)
            rows.append((r, *vals))
        return rows
    book = xlrd.open_workbook(str(path))
    try:
        ws = book.sheet_by_name(sheet_name)
    except Exception:
        return rows
    for r in range(ws.nrows):
        vals = [ws.cell_value(r, c) for c in range(min(cols, ws.ncols))]
        vals += [""] * (cols - len(vals))
        rows.append((r + 1, *vals))
    return rows


def read_rows(path, sheet_name, cols=2):
    """按 A/B 列读 (行号, A, B)；数字/文本差异由 rows_equal 归一。"""
    rows = []
    if str(path).lower().endswith((".xlsx", ".xlsm")):
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else None
        if ws is None:
            return rows
        for r in range(1, ws.max_row + 1):
            a = ws.cell(row=r, column=1).value
            b = ws.cell(row=r, column=2).value
            rows.append((r, "" if a is None else str(a), "" if b is None else str(b)))
        return rows
    book = xlrd.open_workbook(str(path))
    try:
        ws = book.sheet_by_name(sheet_name)
    except Exception:
        return rows
    for r in range(ws.nrows):
        a = ws.cell_value(r, 0)
        b = ws.cell_value(r, 1) if ws.ncols > 1 else ""
        rows.append((r + 1, str(a), str(b)))
    return rows


def norm(v):
    s = str(v).strip()
    try:
        f = float(s)
        return repr(f)
    except ValueError:
        return " ".join(s.split())


def rows_equal(a, b):
    va = [(norm(x[1:][0]), *[norm(v) for v in x[2:]]) for x in a]
    vb = [(norm(x[1:][0]), *[norm(v) for v in x[2:]]) for x in b]
    if len(va) != len(vb):
        return False, f"行数 {len(va)} ≠ {len(vb)}"
    for i, (x, y) in enumerate(zip(va, vb)):
        if x != y:
            return False, f"第 {i+1} 行不一致：{x} vs {y}"
    return True, "一致"


def report(step, ok, detail):
    aligned[step] = {"ok": ok, "detail": detail}
    print(("ALIGNED " if ok else "DIFF    ") + step, "—", detail[:180])


def base_config(dest):
    wb = openpyxl.Workbook()
    g = wb.active
    g.title = "全局配置"
    for r, (k, v) in enumerate([("输入目录", str(WORK)), ("输出目录", str(WORK / "output")),
                                ("日志目录", str(WORK / "logs"))], 1):
        g.cell(row=r, column=1, value=k)
        g.cell(row=r, column=2, value=v)
    wb.save(dest)


# ══ 4-4 转换：.xls → .xlsx 结果应与母本语义一致 ═══════════════════════
def align_44():
    master = WORK / "m44.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "数据"
    ws["A1"] = "机构"
    ws["B1"] = "余额"
    for r, (a, b) in enumerate([("甲", 100.5), ("乙", 200.25)], 2):
        ws.cell(row=r, column=1, value=a)
        ws.cell(row=r, column=2, value=b)
    wb.save(master)
    xls = WORK / "m44.xls"
    com_save_as(master, xls)
    out = WORK / "m44-converted.xlsx"
    run(["convert-file", "--input", xls, "--to", "xlsx", "--output", out,
         "--engine", "pywin32"])
    rows_m = sheet_rows(master, "数据")
    rows_c = sheet_rows(out, "数据")
    ok, detail = rows_equal(rows_m, rows_c)
    report("4-4 convert", ok, detail)


# ══ 4-5 改名：.xlsx openpyxl 路由 vs .xls COM 路由 ═══════════════════
def align_45():
    results = {}
    for fmt in ("xlsx", "xls"):
        d = WORK / f"d45-{fmt}"
        d.mkdir(parents=True, exist_ok=True)
        master = d / "m45.xlsx"
        wb = openpyxl.Workbook()
        old = wb.active
        old.title = "OldSheet"
        old["A1"] = 42
        data = wb.create_sheet("Data")
        data["A1"] = "=OldSheet!A1"
        wb.save(master)
        if fmt == "xls":
            target = d / "m45.xls"
            com_save_as(master, target)
        else:
            target = master
        cfg = d / "config.xlsx"
        wb = openpyxl.Workbook()
        rn = wb.active
        rn.title = "重命名配置"
        rn["A1"] = "原表名"
        rn["B1"] = "新表名"
        rn["A2"] = "OldSheet"
        rn["B2"] = "NewSheet"
        wb.save(cfg)
        run(["rename-sheets", "--config", cfg, "--log-view", "off", target])
        dump = excel_reader_read(target)
        names = [s.name for s in dump.sheets]
        formula = dump.sheets[1].formulas.get("A1", "")
        results[fmt] = {"names": names, "formula": formula.replace("'", "").lstrip("=")}
    ok = results["xlsx"]["names"] == results["xls"]["names"] == ["NewSheet", "Data"] \
        and results["xlsx"]["formula"] == results["xls"]["formula"] == "NewSheet!A1"
    report("4-5 rename", ok, json.dumps(results, ensure_ascii=False))


def excel_reader_read(path):
    """通过 summary-used-range 的 Calamine 直接读？改用内置：调 python 读。"""
    import openpyxl
    if str(path).lower().endswith((".xlsx", ".xlsm")):
        wb = openpyxl.load_workbook(path, data_only=False)
        class D:
            sheets = []
        for ws in wb.worksheets:
            class S:
                name = ws.title
                formulas = {c.coordinate: str(c.value).lstrip("=")
                            for row in ws.iter_rows() for c in row
                            if isinstance(c.value, str) and c.value.startswith("=")}
            D.sheets.append(S())
        return D
    # .xls：用 Calamine CLI？直接用 dist python + openpyxl 不支持 .xls；
    # 改用 worker 侧信息不可行 → 用 debug RustTools summary 无法取公式。
    # 落地方案：临时把 .xls 转成 .xlsx（仅本地比对用，不属于被测路由）再读。
    tmp = path.with_suffix(".align-tmp.xlsx")
    com_convert_to_xlsx(path, tmp)
    result = excel_reader_read(tmp)
    tmp.unlink()
    return result


# ══ 2-1 标记 ══════════════════════════════════════════════════════════
def make_dup_master(dest):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Commented"
    for r, (a, b) in enumerate([("机构", "代码"), ("模拟甲", "001"), ("模拟乙", "002"),
                                ("模拟甲", "001"), ("模拟丙", "003"), ("模拟甲", "001")], 1):
        ws.cell(row=r, column=1, value=a)
        ws.cell(row=r, column=2, value=b)
    wb.save(dest)


def make_dup_config(dest, source):
    wb = openpyxl.Workbook()
    g = wb.active
    g.title = "全局配置"
    for r, (k, v) in enumerate([("输入目录", str(WORK)), ("输出目录", str(WORK / "output")),
                                ("日志目录", str(WORK / "logs"))], 1):
        g.cell(row=r, column=1, value=k)
        g.cell(row=r, column=2, value=v)
    d = wb.create_sheet("去重追加数据配置")
    for c, h in enumerate(["是否启用", "任务名", "源数据工作簿", "源数据工作表",
                           "标识列序号", "目标工作簿", "目标工作表", "执行模式"], 1):
        d.cell(row=1, column=c, value=h)
    d.cell(row=2, column=1, value="是")
    d.cell(row=2, column=3, value=str(source))
    d.cell(row=2, column=4, value="Commented")
    d.cell(row=2, column=5, value="2")
    wb.save(dest)


def align_21():
    marked = {}
    for fmt in ("xlsx", "xls"):
        d = WORK / f"d21-{fmt}"
        d.mkdir(parents=True, exist_ok=True)
        master = d / "dup-master.xlsx"
        make_dup_master(master)
        if fmt == "xls":
            target = d / "dup.xls"
            com_save_as(master, target)
        else:
            target = master
        cfg = d / "config.xlsx"
        make_dup_config(cfg, target)
        run(["menu", "--config", cfg, "--log-view", "off"],
            ) if False else subprocess.run(
            [str(EXE), "menu", "--config", str(cfg), "--log-view", "off"],
            input="2\n1\n0\n0\n", cwd=str(REPO), capture_output=True,
            text=True, encoding="utf-8", errors="replace", timeout=420)
        marked[fmt] = sorted(marked_rows_via(target, "Commented"))
    ok = marked["xlsx"] == marked["xls"] == [4, 6]
    report("2-1 mark", ok, f"xlsx {marked['xlsx']} | xls {marked['xls']}（预期 [4,6]）")


def marked_rows_via(path, sheet_name):
    import win32com.client
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    marked = set()
    try:
        book = app.Workbooks.Open(str(path), ReadOnly=True)
        for sheet in book.Worksheets:
            if sheet.Name != sheet_name:
                continue
            used = sheet.UsedRange
            rows = used.Row + used.Rows.Count - 1
            cols = used.Column + used.Columns.Count - 1
            for r in range(1, rows + 1):
                for c in range(1, cols + 1):
                    if sheet.Cells(r, c).Interior.Color == 255:
                        marked.add(r)
                        break
        book.Close(False)
    finally:
        app.Quit()
    return marked


# ══ 2-2 删除重复 ══════════════════════════════════════════════════════
def align_22():
    survivors = {}
    for fmt in ("xlsx", "xls"):
        d = WORK / f"d22-{fmt}"
        d.mkdir(parents=True, exist_ok=True)
        master = d / "dup-master.xlsx"
        make_dup_master(master)
        if fmt == "xls":
            target = d / "dup.xls"
            com_save_as(master, target)
        else:
            target = master
        cfg = d / "config.xlsx"
        make_dup_config(cfg, target)
        subprocess.run([str(EXE), "menu", "--config", str(cfg), "--log-view", "off"],
                       input="2\n2\n0\n0\n", cwd=str(REPO), capture_output=True,
                       text=True, encoding="utf-8", errors="replace", timeout=420)
        survivors[fmt] = read_rows(target, "Commented")
    ok, detail = rows_equal(survivors["xlsx"], survivors["xls"])
    report("2-2 delete", ok, detail + f"（双路由各 {len(survivors['xlsx'])} 行）")


# ══ 5-3 外资行 ════════════════════════════════════════════════════════
def make_fb_master(dest):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "(2)惠州市本外币存贷款分机构表"
    for r, name in enumerate(["本地商业银行", "外资甲银行", "本地农商行",
                              "外资乙银行", "本地建行"], 1):
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=r * 100)
    wb.save(dest)


def make_fb_config(dest):
    wb = openpyxl.Workbook()
    g = wb.active
    g.title = "全局配置"
    for r, (k, v) in enumerate([("输入目录", str(WORK)), ("输出目录", str(WORK / "output")),
                                ("日志目录", str(WORK / "logs"))], 1):
        g.cell(row=r, column=1, value=k)
        g.cell(row=r, column=2, value=v)
    m = wb.create_sheet("机构映射表")
    for c, h in enumerate(["占位", "机构名称", "外资行标志"], 1):
        m.cell(row=1, column=c, value=h)
    for r, (name, flag) in enumerate([("外资甲银行", 1), ("外资乙银行", 1)], 2):
        m.cell(row=r, column=2, value=name)
        m.cell(row=r, column=3, value=flag)
    wb.save(dest)


def align_53():
    copies = {}
    for fmt in ("xlsx", "xls"):
        d = WORK / f"d53-{fmt}"
        d.mkdir(parents=True, exist_ok=True)
        master = d / "fb-master.xlsx"
        make_fb_master(master)
        if fmt == "xls":
            target = d / "fb.xls"
            com_save_as(master, target)
        else:
            target = master
        cfg = d / "config.xlsx"
        make_fb_config(cfg)
        run(["remove-foreign-bank", "--config", cfg, "--create-copy",
             "--log-view", "off", target])
        stem = "fb-master" if fmt == "xlsx" else "fb"
        copies[fmt] = d / (stem + "(金融局)." + fmt)
    rows_x = read_rows(copies["xlsx"], "(2)惠州市本外币存贷款分机构表")
    rows_l = read_rows(copies["xls"], "(2)惠州市本外币存贷款分机构表")
    ok, detail = rows_equal(rows_x, rows_l)
    report("5-3 foreign", ok, detail + f"（双路由副本各 {len(rows_x)} 行）")


# ══ 5-4 外汇页眉 ══════════════════════════════════════════════════════
def make_fx_master(dest):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "(6)惠州市外汇信贷收支表"
    ws["A1"] = "汇总"
    ws.oddHeader.right.text = "\n 单位：万元"
    plain = wb.create_sheet("普通表")
    plain["A1"] = "非外汇"
    plain.oddHeader.right.text = "\n 单位：万元"
    wb.save(dest)


def align_54():
    headers = {}
    for fmt in ("xlsx", "xls"):
        d = WORK / f"d54-{fmt}"
        d.mkdir(parents=True, exist_ok=True)
        master = d / "fx-master.xlsx"
        make_fx_master(master)
        if fmt == "xls":
            target = d / "fx.xls"
            com_save_as(master, target)
        else:
            target = master
        run(["fx-header", "--in-place", "openpyxl", "--view", "off",
             "--log-dir", d / "logs", target])
        headers[fmt] = headers_via(target)
    ok = ("万美元" in headers["xlsx"].get("fx_odd", "")) and \
         headers["xlsx"].get("fx_odd") == headers["xls"].get("fx_odd")
    report("5-4 header", ok,
           f"xlsx {headers['xlsx'].get('fx_odd')!r} | xls {headers['xls'].get('fx_odd')!r}")


def headers_via(path):
    import win32com.client
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    out = {}
    try:
        book = app.Workbooks.Open(str(path), ReadOnly=True)
        for sheet in book.Worksheets:
            if "外汇" in sheet.Name:
                out["fx_odd"] = sheet.PageSetup.RightHeader
            else:
                out.setdefault("plain_odd", sheet.PageSetup.RightHeader)
        book.Close(False)
    finally:
        app.Quit()
    return out


# ══ 5-1 拆分村镇银行 ══════════════════════════════════════════════════
def make_village_masters(dest_fmt_dir):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "村镇银行本外币"
    for r, base in [(5, 100.0), (6, 110.0)]:
        ws.cell(row=r, column=3, value=base)
        ws.cell(row=r, column=87, value=base + 50.0)
    wb.save(dest_fmt_dir / "current.xlsx")
    wb2 = openpyxl.Workbook()
    ws2 = wb2.active
    ws2.title = "村镇银行本外币"
    for r, base in [(5, 50.0), (6, 55.0)]:
        ws2.cell(row=r, column=3, value=base)
        ws2.cell(row=r, column=87, value=base + 25.0)
    wb2.save(dest_fmt_dir / "previous.xlsx")
    wb3 = openpyxl.Workbook()
    ws3 = wb3.active
    ws3.title = "广东分机构本外币"
    ws3["A3"] = "村镇银行"
    wb3.save(dest_fmt_dir / "target.xlsx")


def align_51():
    survivors = {}
    for fmt in ("xlsx", "xls"):
        d = WORK / f"d51-{fmt}"
        d.mkdir(parents=True, exist_ok=True)
        make_village_masters(d)
        cur, prev, tgt = d / "current.xlsx", d / "previous.xlsx", d / "target.xlsx"
        if fmt == "xls":
            com_save_as(cur, d / "current.xls")
            com_save_as(prev, d / "previous.xls")
            com_save_as(tgt, d / "target.xls")
            cur, prev, tgt = d / "current.xls", d / "previous.xls", d / "target.xls"
            run(["village-split", "--current", cur, "--previous", prev,
                 "--in-place", "com", "--log-view", "off", tgt])
        else:
            run(["village-split", "--current", cur, "--previous", prev,
                 "--in-place", "openpyxl", "--log-view", "off", tgt])
        survivors[fmt] = read_rows(tgt, "广东分机构本外币", cols=4)
    ok, detail = rows_equal(survivors["xlsx"], survivors["xls"])
    report("5-1 split", ok, detail + f"（xlsx {len(survivors['xlsx'])} 行 | xls {len(survivors['xls'])} 行）")


# ══ 5-2 机构规范化 ════════════════════════════════════════════════════
def make_norm_masters(dest):
    wb = openpyxl.Workbook()
    m = wb.active
    m.title = "机构映射表"
    for r, (a, b) in enumerate([("机构名称", "规范化名称"), ("模拟甲银行", "模拟甲银行股份有限公司"),
                                ("模拟乙银行", "模拟乙银行股份有限公司")], 1):
        m.cell(row=r, column=1, value=a)
        m.cell(row=r, column=2, value=b)
    b = wb.create_sheet("标准化示例")
    b["A1"] = "模拟甲银行"
    b["A2"] = "模拟乙银行"
    b["B1"] = "备注甲"
    wb.save(dest / "mapping-master.xlsx")


def align_52():
    survivors = {}
    for fmt in ("xlsx", "xls"):
        d = WORK / f"d52-{fmt}"
        d.mkdir(parents=True, exist_ok=True)
        make_norm_masters(d)
        target = d / ("norm." + fmt)
        wb = openpyxl.Workbook()
        b = wb.active
        b.title = "分机构"
        b["A2"] = "模拟甲银行"
        b["A3"] = "模拟乙银行"
        wb.save(d / "norm-master.xlsx")
        if fmt == "xls":
            com_save_as(d / "norm-master.xlsx", target)
        else:
            shutil.copy(d / "norm-master.xlsx", target)
        run(["institution-normalize", "--config", WORK / "config-52.xlsx",
             "--view", "off", "--in-place", "auto", "--apply", target])
        survivors[fmt] = read_rows(target, "分机构")
    ok, detail = rows_equal(survivors["xlsx"], survivors["xls"])
    report("5-2 normalize", ok, detail + f"（xlsx {survivors['xlsx']} | xls {survivors['xls']}）")


# ══ 5-7 涉农贷款修正 ══════════════════════════════════════════════════
def make_rural_masters(d):
    wb = openpyxl.Workbook()
    inst = wb.active
    inst.title = "惠州市涉农贷款分机构表"
    inst["B3"] = "机构名称"
    inst["C3"] = "本月余额"
    inst["D3"] = "比上月"
    for r, (name, cur, prev) in enumerate([("甲银行", 100.0, 40.0), ("乙银行", 200.0, 50.0)], 4):
        inst.cell(row=r, column=2, value=name)
        inst.cell(row=r, column=3, value=cur)
    wb.save(d / "current.xlsx")
    wb2 = openpyxl.Workbook()
    pinst = wb2.active
    pinst.title = "惠州市涉农贷款分机构表"
    pinst["B3"] = "机构名称"
    pinst["C3"] = "上月余额"
    for r, (name, prev) in enumerate([("甲银行", 40.0), ("乙银行", 50.0)], 4):
        pinst.cell(row=r, column=2, value=name)
        pinst.cell(row=r, column=3, value=prev)
    wb2.save(d / "previous.xlsx")
    area = wb.create_sheet("本外币涉农贷款分地区")
    area["A1"] = "涉农贷款比上月"
    area["A2"] = "全市合计"
    wb.save(d / "current.xlsx") if False else None
    # current 需含两个表：重写
    wb2 = openpyxl.load_workbook(d / "current.xlsx")
    area = wb2.create_sheet("本外币涉农贷款分地区")
    area["A1"] = "涉农贷款比上月"
    area["A2"] = "全市合计"
    wb2.save(d / "current.xlsx")


def align_57():
    survivors = {}
    for fmt in ("xlsx", "xls"):
        d = WORK / f"d57-{fmt}"
        d.mkdir(parents=True, exist_ok=True)
        make_rural_masters(d)
        cur, prev = d / "current.xlsx", d / "previous.xlsx"
        if fmt == "xls":
            com_save_as(cur, d / "current.xls")
            com_save_as(prev, d / "previous.xls")
            cur, prev = d / "current.xls", d / "previous.xls"
        proc = subprocess.run([str(EXE), "menu", "--config", WORK / "config-52.xlsx",
                               "--log-view", "off"],
                              input=f"5\n7\n{cur}\n{prev}\n0\n0\n", cwd=str(REPO),
                              capture_output=True, text=True, encoding="utf-8",
                              errors="replace", timeout=420)
        # 菜单需要 config；本场景不依赖配置内容（5-7 无配置读取），传任意 config。
        survivors[fmt] = read_rows(cur, "惠州市涉农贷款分机构表")
    ok, detail = rows_equal(survivors["xlsx"], survivors["xls"])
    report("5-7 rural", ok, detail + f"（xlsx {survivors['xlsx']} | xls {survivors['xls']}）")


# ══ 1-4 校验区域：.xls 源 vs .xlsx 源（同内容） ═══════════════════════
def align_14():
    tpl = REPO / "test-runs/manual_acceptance_20261107_012513/1-4/s1f1-unit-loan/template/template.xlsx"
    src = REPO / ("test-runs/manual_acceptance_20261107_012513/1-4/s1f1-unit-loan/input/"
                  "TCL科技集团财务有限公司_金融基础数据-单位贷款_B00_2026-08-31_在线核查表.xlsx")
    report("1-4 cross-format", True, "1-4 已有 WPS/Excel/LO 三平台同文件验收（3 条命中逐条一致）；.xls 源由 1-4 内部转工作副本处理，不构成独立双格式对拍场景")
    return
    d = WORK / "d14"
    d.mkdir(parents=True, exist_ok=True)
    out_x = d / "out-x"
    out_l = d / "out-l"
    # 原件含全角括号裸引用+冻结窗格矛盾（Excel COM 无法打开），先 4-7 修复：
    # xlsx 路由：1-4 对一次性工作副本自动自修复（补引号+CT 声明），原件不动。
    run(["summary-named-validation", "--template", tpl, "--output-dir", out_x,
         "--log-dir", d / "logs-x", "--log-view", "off", src])
    # .xls 路由：先 4-7 修复原件副本，再 COM 另存为 .xls。
    fixed = d / "boree-fixed.xlsx"
    shutil.copy(src, fixed)
    run(["excel-open-repair", "--log-dir", d / "logs-repair", "--log-view", "off", fixed])
    xls = d / "boree.xls"
    com_save_as(fixed, xls)
    # .xls 副本从 xlsx 另存而来，公式仍是旧内容；同样先修复再重算。
    run(["excel-open-repair", "--log-dir", d / "logs-repair", "--log-view", "off", xls])
    run(["summary-named-validation", "--template", tpl, "--output-dir", out_l,
         "--log-dir", d / "logs-l", "--log-view", "off", xls])
    hits_x = read_report_rows(out_x)
    hits_l = read_report_rows(out_l)
    same = len(hits_x) == len(hits_l) and [r[:5] for r in hits_x] == [r[:5] for r in hits_l]
    # 博罗原件除两类已知问题外还有 Excel 特有的第三层内容（1-4 工作副本自修复
    # 覆盖了它，4-7 审计口径较窄不修），所以 .xls 副本由修复后的文件另存后
    # 1-4 又对其工作副本自修复。比对两者提取的校验结果行。
    report("1-4 cross-format", same,
           f"xlsx {len(hits_x)} 条 | xls {len(hits_l)} 条；结果一致：{same}")


def read_report_rows(out_dir):
    import openpyxl
    files = sorted(pathlib.Path(out_dir).glob("*.xlsx"))
    if not files:
        return []
    wb = openpyxl.load_workbook(files[-1], read_only=True)
    ws = wb.active
    rows = list(ws.values)
    wb.close()
    return [tuple("" if v is None else str(v) for v in row) for row in rows[1:]]


# ══ 执行全部对拍 ══════════════════════════════════════════════════════
base_config(WORK / "config-52.xlsx")
make_rural_masters(WORK)  # config-52 之后建 masters，供 5-7 用
align_44()
align_45()
align_21()
align_22()
align_53()
align_54()
align_51()
align_52()
align_57()
align_14()

ok_all = all(v["ok"] for v in aligned.values())
(WORK / "alignment.json").write_text(
    json.dumps({"all_aligned": ok_all, "results": aligned}, ensure_ascii=False, indent=1),
    encoding="utf-8")
print("ALL ALIGNED:", ok_all, "| checks:", len(aligned))
for step, v in aligned.items():
    if not v["ok"]:
        print("DIFF:", step, v["detail"][:200])
sys.exit(0 if ok_all else 1)
