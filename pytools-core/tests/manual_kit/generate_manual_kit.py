# -*- coding: utf-8 -*-
"""Generate the manual menu testing kit under testdata/manual-menu/.

All data is synthetic; no real institution or business data is used. The
generated config.xlsx follows the sheet schema implemented by
pytools-core/src/config.rs (全局配置 / 时序提取规则 / 路径标准化映射 /
去重追加数据配置 / 打印配置 / 机构映射表 / 工作表提取).

Re-run any time to rebuild the kit:
    python pytools-core/tests/manual_kit/generate_manual_kit.py
"""

from __future__ import annotations

import shutil
from pathlib import Path

import openpyxl
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font

ROOT = Path(__file__).resolve().parents[3]
KIT = ROOT / "testdata" / "manual-menu"
FIXTURES = ROOT / "pytools-core" / "tests" / "fixtures"


def add_sheet(wb: Workbook, name: str, rows: list[list], header_bold: bool = True) -> None:
    sheet = wb.create_sheet(name)
    for row_index, row in enumerate(rows, start=1):
        for col_index, value in enumerate(row, start=1):
            cell = sheet.cell(row=row_index, column=col_index, value=value)
            if row_index == 1 and header_bold and value not in (None, ""):
                cell.font = Font(bold=True)
    return sheet


def save_workbook(path: Path, wb: Workbook) -> None:
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def build_config() -> None:
    wb = Workbook()

    # 表头与真实 config.xlsx 对齐（键/值/备注三列、Y/N、数字布尔、执行模式）。
    add_sheet(wb, "全局配置", [
        ["键", "值", "备注"],
        ["输入目录", "input", ""],
        ["输出目录", "output", ""],
        ["日志目录", "logs", ""],
        ["错误策略", "continue", ""],
        ["默认编码", "utf-8", ""],
        ["源文件扩展名", ".xlsx;.xlsm;.csv;.et", ""],
        ["时序规则命中策略", "all_match", ""],
        ["输出附带计划文件", "是", ""],
    ])

    add_sheet(wb, "时序提取规则", [
        ["是否启用", "规则名称", "工作簿关键字", "工作表关键字", "行头列", "列表头行",
         "必含列头", "必含行头", "数据起始行", "数据结束行", "数据起始列", "数据结束列",
         "跳过关键字", "目标工作簿路径", "目标工作表", "启用目标写入", "set区域",
         "目标去重列"],
        ["Y", "测试余额时序", "月度报表", "余额表", "A", "1",
         "", "存款,贷款", "2", "", "B", "", "合计",
         "input/timeline/历史宽表.xlsx", "宽表", "1", "", ""],
    ])

    add_sheet(wb, "路径标准化映射", [
        ["是否启用", "映射名称", "适用规则名", "工作簿关键字", "工作表关键字",
         "作用对象", "匹配方式", "原始路径", "标准路径", "备注"],
        ["Y", "机构别名归一", "", "月度报表", "余额表", "两者", "精确",
         "测试机构甲", "甲机构", ""],
    ])

    add_sheet(wb, "去重追加数据配置", [
        ["是否启用", "源数据工作簿", "源数据工作表", "标识列序号",
         "目标工作簿", "目标工作表", "执行模式", "备注"],
        ["Y", "input/dedup/报送数据.xlsx", "数据", "1",
         "input/dedup/历史汇总.xlsx", "汇总", 1, ""],
    ])

    add_sheet(wb, "打印配置", [
        ["是否启用", "打印模式", "源工作簿", "源工作表", "源工作表打印区域",
         "目标工作簿", "目标工作表", "FitToPagesWide", "FitToPagesTall",
         "打印方向", "水平居中", "垂直居中", "不输出批注", "零值不输出", "备注"],
        ["Y", "1", "input/print/打印源表.xlsx", "数据", "",
         "output/print/静态打印.xlsx", "静态报表", 1, 1,
         "纵向", 1, 1, 1, 1, "合成样本"],
    ])

    add_sheet(wb, "机构映射表", [
        ["原始机构名称", "映射后机构名称", "是否为外资行"],
        ["测试村镇银行甲", "甲村镇银行", 1],
        ["测试外资银行", "示范外资银行", 1],
        ["测试机构乙", "乙机构", 0],
        ["测试机构丙", "丙机构", 0],
        ["测试机构丁", "丁机构", 0],
    ])

    add_sheet(wb, "工作表提取", [
        ["是否启用", "工作簿关键字", "工作表关键字", "排除关键字", "提取区域",
         "多区域拼接方式", "输出文件夹", "文件名", "备注"],
        ["Y", "省级月度报表", "数据", "", "A1:C5", "",
         "提取结果", "提取输出.xlsx", "合成样本"],
    ])

    save_workbook(KIT / "config" / "config.xlsx", wb)


def build_timeline_sources() -> None:
    for institution in ("甲机构", "乙机构"):
        wb = Workbook()
        sheet = add_sheet(wb, "余额表", [
            ["科目", "2026-01-01", "2026-02-01", "2026-03-01"],
            ["存款", 1200.5, 1310.0, 0.0],
            ["贷款", 800.0, 815.25, 830.0],
            ["合计", 2000.5, 2125.25, 830.0],
        ])
        sheet.cell(row=1, column=1).comment = Comment("合成样本：表头行", "PyTools 测试")
        save_workbook(KIT / "input" / "timeline" / f"月度报表_{institution}.xlsx", wb)


def build_dedup_sources() -> None:
    wb = Workbook()
    add_sheet(wb, "数据", [
        ["机构", "存款", "贷款"],
        ["测试机构乙", 100.0, 50.0],
        ["测试机构乙", 100.0, 50.0],   # 完全相同的记录
        ["测试机构丙", 100.0, 50.0],   # 不同机构、相同指标：不是重复
        ["测试机构丁", 200.0, 90.0],
    ])
    save_workbook(KIT / "input" / "dedup" / "报送数据.xlsx", wb)

    wb = Workbook()
    add_sheet(wb, "汇总", [
        ["机构", "存款", "贷款"],
        ["测试机构乙", 100.0, 50.0],   # 历史已有记录：追加时应跳过
    ])
    save_workbook(KIT / "input" / "dedup" / "历史汇总.xlsx", wb)


def build_print_sources() -> None:
    wb = Workbook()
    sheet = add_sheet(wb, "数据", [
        ["项目", "数值", "说明"],
        ["数值零", 0, "数值 0（零值不输出时应转空白）"],
        ["文本零", "0", '文本 "0"（不应被误清空）'],
        ["微小非零", 0.0001, "非零小数应保留"],
        ["正常值", 123.45, "普通数值"],
    ])
    sheet.merge_cells("A6:D6")
    sheet["A6"] = "合并区域标题"
    sheet["A1"].comment = Comment("合成样本：源批注不应出现在静态打印结果", "PyTools 测试")
    save_workbook(KIT / "input" / "print" / "打印源表.xlsx", wb)

    wb = Workbook()
    sheet = add_sheet(wb, "数据", [
        ["项目", "数值"],
        ["公式无缓存", "=1/0"],
    ])
    save_workbook(KIT / "input" / "print" / "打印源表_公式无缓存.xlsx", wb)


def build_file_processing() -> None:
    base = KIT / "input" / "files"

    # 4-1 / 5-2：使用原始（未标准化）机构名的文件
    for raw, _standard in [
        ("测试村镇银行甲", "甲村镇银行"),
        ("测试机构乙", "乙机构"),
        ("测试机构丙", "丙机构"),
    ]:
        wb = Workbook()
        add_sheet(wb, "报表", [["机构", raw], ["数值", 1.0]])
        save_workbook(base / "机构原名" / f"{raw} 2026-01.xlsx", wb)

    # 4-2：带编号排序前缀的文件（添加/移除）
    for name in ("01_已排序.xlsx", "02_已排序.xlsx"):
        wb = Workbook()
        add_sheet(wb, "报表", [["前缀样本", name]])
        save_workbook(base / "排序前缀" / name, wb)

    # 4-3：内容重命名映射表 + 一个待重命名文件
    wb = Workbook()
    add_sheet(wb, "重命名", [
        ["当前文件名", "新文件名"],
        ["测试机构乙 2026-01.xlsx", "乙机构_2026-01.xlsx"],
    ])
    save_workbook(base / "内容重命名" / "重命名映射.xlsx", wb)
    wb = Workbook()
    add_sheet(wb, "报表", [["内容", "用于内容重命名测试"]])
    save_workbook(base / "内容重命名" / "测试机构乙 2026-01.xlsx", wb)

    # 4-4：转换样本来自既有合成 legacy .xls 夹具；转换引擎不可用时应明确报错
    (base / "待转换").mkdir(parents=True, exist_ok=True)
    shutil.copyfile(
        FIXTURES / "timeline" / "wide_summary_legacy.xls",
        base / "待转换" / "示例报表.xls",
    )

    # 4-5：工作表重命名样本
    wb = Workbook()
    add_sheet(wb, "旧表名", [["数据", 1]])
    add_sheet(wb, "辅助表", [["数据", 2]])
    save_workbook(base / "工作表重命名" / "工作表重命名样本.xlsx", wb)

    # 4-6：报送完整性检查——五个机构只报送了三个
    for raw in ("测试村镇银行甲", "测试机构乙", "测试机构丙"):
        wb = Workbook()
        add_sheet(wb, "报表", [["机构", raw], ["数值", 2.0]])
        save_workbook(base / "报送检查" / "已报送" / f"{raw}.xlsx", wb)


def build_special_reports() -> None:
    base = KIT / "input" / "special"

    # 5-3：删除外资行——外资行数据只应删除所在报表行，不删除整表
    wb = Workbook()
    add_sheet(wb, "机构明细", [
        ["机构", "存款", "贷款"],
        ["测试机构乙", 100.0, 50.0],
        ["测试外资银行", 300.0, 150.0],
        ["测试机构丙", 120.0, 60.0],
        ["测试外资银行", 310.0, 155.0],
    ])
    save_workbook(base / "分机构报表.xlsx", wb)

    # 5-4：外汇页眉样本来自既有合成夹具
    shutil.copyfile(FIXTURES / "fx_header" / "source.xlsx", base / "外汇报表.xlsx")

    # 5-5：地区总分——balanced 应通过核对，mismatch 应报告差异
    shutil.copyfile(FIXTURES / "region_sum" / "balanced.xlsx", base / "地区总分_平衡.xlsx")
    shutil.copyfile(FIXTURES / "region_sum" / "mismatch.xlsx", base / "地区总分_差异.xlsx")

    # 5-6：工作表数据提取样本来自既有合成夹具
    shutil.copyfile(
        FIXTURES / "sheet_extract_area" / "省级月度报表_A.xlsx",
        base / "省级月度报表_测试.xlsx",
    )

    # 5-7 / 5-8：涉农贷款最简形状样本（正式自动化断言见 tests 内 parity 用例）
    wb = Workbook()
    add_sheet(wb, "涉农贷款分机构", [
        ["机构名称", "本月余额", "比上月增减"],
        ["测试机构乙", 500.0, 20.0],
        ["测试机构丙", 420.0, 15.0],
        ["合计", 920.0, 35.0],
    ])
    add_sheet(wb, "涉农贷款分地区统计", [
        ["地区", "本月余额", "比上月增减"],
        ["城区", 500.0, 20.0],
        ["县域", 420.0, 15.0],
    ])
    save_workbook(base / "涉农贷款_月度.xlsx", wb)

    wb = Workbook()
    add_sheet(wb, "涉农贷款分机构", [
        ["机构名称", "本季余额", "比上季增减"],
        ["测试机构乙", 1500.0, 60.0],
        ["测试机构丙", 1260.0, 45.0],
        ["合计", 2760.0, 105.0],
    ])
    add_sheet(wb, "涉农贷款分地区统计", [
        ["地区", "本季余额", "比上季增减"],
        ["城区", 1500.0, 60.0],
        ["县域", 1260.0, 45.0],
    ])
    save_workbook(base / "涉农贷款_季度.xlsx", wb)


def build_archive_tree() -> None:
    base = KIT / "input" / "archive" / "待归档"
    files = [
        ("2026-01-15_月度报表_甲机构.xlsx", "报表"),
        ("2026-01-20_报送清单.csv", "清单"),
        ("2026-02-01_月度报表_乙机构.xlsx", "报表"),
        ("说明文档.docx", "文档"),
    ]
    for name, _kind in files:
        wb = Workbook()
        add_sheet(wb, "报表", [["样本", name]])
        # .csv/.docx 占位样本仅用于归档分类演示，内容不是有效对应格式
        path = base / name
        path.parent.mkdir(parents=True, exist_ok=True)
        if path.suffix == ".xlsx":
            save_workbook(path, wb)
        else:
            path.write_text("synthetic placeholder", encoding="utf-8")
    (KIT / "input" / "archive" / "归档目标").mkdir(parents=True, exist_ok=True)


def main() -> None:
    if KIT.exists():
        shutil.rmtree(KIT)
    build_config()
    build_timeline_sources()
    build_dedup_sources()
    build_print_sources()
    build_file_processing()
    build_special_reports()
    build_archive_tree()
    for empty in ("output", "logs"):
        (KIT / empty).mkdir(parents=True, exist_ok=True)
    print(f"manual kit generated at {KIT}")


if __name__ == "__main__":
    main()
