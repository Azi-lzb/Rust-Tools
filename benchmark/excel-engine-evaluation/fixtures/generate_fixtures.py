"""Generate shared synthetic phase-2 fixtures; this is not an engine under test."""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.page import PageMargins
from openpyxl.workbook.defined_name import DefinedName
import xlwt


ROOT = Path(__file__).resolve().parent
ROOT.mkdir(parents=True, exist_ok=True)

wb = Workbook()
ws = wb.active
ws.title = "Data"
ws.append(["Item", "Quantity", "Unit Price", "Amount", "Rate", "Date"])
for row in range(2, 12):
    ws.cell(row, 1, f"item-{row - 1}")
    ws.cell(row, 2, row - 1)
    ws.cell(row, 3, 1.25 * row)
    ws.cell(row, 4, f"=B{row}*C{row}")
    ws.cell(row, 5, 0.125)
    ws.cell(row, 6, f"2026-01-{row:02d}")

header_fill = PatternFill("solid", fgColor="17365D")
header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
thin_blue = Side(style="thin", color="4472C4")
for cell in ws[1]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = Border(bottom=thin_blue)
ws[2][2].number_format = '#,##0.00;[Red]-#,##0.00'
for row in range(2, 12):
    ws.cell(row, 3).number_format = '0.000'
    ws.cell(row, 4).number_format = '0.00'
    ws.cell(row, 5).number_format = '0.0%'
ws[2][0].comment = Comment("synthetic capability fixture", "rustTools evaluation")
ws[3][0].hyperlink = "https://example.invalid/item-2"
ws[3][0].style = "Hyperlink"
ws.merge_cells("H1:I1")
ws["H1"] = "Merged fixture header"
ws["H1"].font = Font(bold=True, color="C00000")
ws.column_dimensions["A"].width = 18
ws.column_dimensions["C"].width = 14
ws.row_dimensions[1].height = 28
ws.row_dimensions[4].hidden = True
ws.column_dimensions["I"].hidden = True

dv = DataValidation(type="whole", operator="between", formula1="0", formula2="100")
dv.promptTitle = "Quantity"
dv.prompt = "Use a non-negative integer"
ws.add_data_validation(dv)
dv.add("B2:B11")

chart = BarChart()
chart.title = "Amount by item"
chart.add_data(Reference(ws, min_col=4, min_row=1, max_row=11), titles_from_data=True)
chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=11))
ws.add_chart(chart, "K2")

ws.print_area = "A1:F11"
ws.print_title_rows = "1:1"
ws.print_title_cols = "A:A"
ws.page_setup.paperSize = ws.PAPERSIZE_A4
ws.page_setup.orientation = "landscape"
ws.page_setup.scale = 83
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_margins = PageMargins(left=0.35, right=0.4, top=0.55, bottom=0.6, header=0.2, footer=0.25)
ws.print_options.horizontalCentered = True
ws.print_options.verticalCentered = False
ws.oddHeader.center.text = "&BPhase 2 fixture"
ws.oddFooter.right.text = "Page &P of &N"
ws.page_setup.pageOrder = "overThenDown"

calc = wb.create_sheet("Calc")
calc["A1"] = "=SUM(Data!B2:B11)"
calc["A2"] = "=SUMIF(Data!A2:A11,\"item-2\",Data!D2:D11)"
calc["A3"] = "=COUNTIF(Data!B2:B11,\">5\")"
calc["A4"] = "=IF(AND(Data!B2>0,Data!C2>0),ROUND(Data!B2*Data!C2,2),0)"
calc["A5"] = "=VLOOKUP(\"item-2\",Data!A2:D11,4,FALSE)"
calc["A6"] = "=INDEX(Data!D2:D11,MATCH(\"item-3\",Data!A2:A11,0))"
calc["A7"] = "=DATE(2026,1,1)"
calc["A8"] = "=SUM(Data!D2:D11)+Data!B2"
calc.page_setup.paperSize = calc.PAPERSIZE_A4
calc.page_setup.orientation = "portrait"
calc.print_area = "A1:A8"

wb.defined_names.add(DefinedName("InputValues", attr_text="Data!$B$2:$B$11"))
wb.calculation.fullCalcOnLoad = True
wb.calculation.forceFullCalc = True
wb.calculation.calcMode = "auto"
wb.save(ROOT / "capabilities.xlsx")

xls = xlwt.Workbook(encoding="utf-8")
xdata = xls.add_sheet("Data")
xcalc = xls.add_sheet("Calc")
header = xlwt.easyxf(
    "font: name Arial, height 220, bold on, colour_index white;"
    "pattern: pattern solid, fore_colour dark_blue;"
    "align: horiz center, vert center, wrap on;"
    "borders: bottom thin, bottom_colour blue;"
)
money = xlwt.easyxf(num_format_str="#,##0.00;[Red]-#,##0.00")
percent = xlwt.easyxf(num_format_str="0.0%")
for col, value in enumerate(["Item", "Quantity", "Unit Price", "Amount", "Rate"]):
    xdata.write(0, col, value, header)
for row in range(1, 11):
    xdata.write(row, 0, f"item-{row}")
    xdata.write(row, 1, row)
    xdata.write(row, 2, 1.25 * (row + 1), money)
    xdata.write(row, 3, xlwt.Formula(f"B{row + 1}*C{row + 1}"), money)
    xdata.write(row, 4, 0.125, percent)
xdata.write_merge(0, 0, 7, 8, "Merged fixture header", header)
xdata.col(0).width = 18 * 256
xdata.col(2).width = 14 * 256
xdata.row(0).height = 28 * 20
xdata.row(3).hidden = True
xdata.col(8).hidden = True
xcalc.write(0, 0, xlwt.Formula("SUM(Data!B2:B11)"))
xcalc.write(1, 0, xlwt.Formula('SUMIF(Data!A2:A11,"item-2",Data!D2:D11)'))
xls.save(str(ROOT / "capabilities.xls"))

