import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.utils import get_column_letter
from pathlib import Path

OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)
OUTPUT_FILE = OUTPUT_DIR / "存储器市场分析报告.xlsx"

COLORS = {
    "title_bg": "1F4E79", "title_fg": "FFFFFF",
    "header_bg": "2E75B6", "header_fg": "FFFFFF",
    "sub_bg": "D6E4F0",
    "accent1": "4472C4", "accent2": "ED7D31", "accent3": "70AD47",
    "light_gray": "F2F2F2", "white": "FFFFFF",
}

thin_border = Border(
    left=Side(style="thin", color="B4C6E7"),
    right=Side(style="thin", color="B4C6E7"),
    top=Side(style="thin", color="B4C6E7"),
    bottom=Side(style="thin", color="B4C6E7"),
)

def title_font(size=18):
    return Font(name="微软雅黑", size=size, bold=True, color=COLORS["title_fg"])

def header_font(size=11):
    return Font(name="微软雅黑", size=size, bold=True, color=COLORS["header_fg"])

def body_font(size=10, bold=False, color="000000"):
    return Font(name="微软雅黑", size=size, bold=bold, color=color)

def title_fill():
    return PatternFill(start_color=COLORS["title_bg"], end_color=COLORS["title_bg"], fill_type="solid")

def header_fill():
    return PatternFill(start_color=COLORS["header_bg"], end_color=COLORS["header_bg"], fill_type="solid")

def sub_fill():
    return PatternFill(start_color=COLORS["sub_bg"], end_color=COLORS["sub_bg"], fill_type="solid")

def stripe_fill(row_idx):
    if row_idx % 2 == 0:
        return PatternFill(start_color=COLORS["light_gray"], end_color=COLORS["light_gray"], fill_type="solid")
    return PatternFill(start_color=COLORS["white"], end_color=COLORS["white"], fill_type="solid")

def set_col_widths(ws, widths):
    for col_letter, w in widths.items():
        ws.column_dimensions[col_letter].width = w

def write_title_row(ws, row, text, merge_end_col, height=40):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=merge_end_col)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = title_font(18)
    cell.fill = title_fill()
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = height

def write_subtitle_row(ws, row, text, merge_end_col, height=28):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=merge_end_col)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = Font(name="微软雅黑", size=12, bold=True, color=COLORS["title_bg"])
    cell.fill = sub_fill()
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = height

def write_header_row(ws, row, headers, height=30):
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.font = header_font(11)
        cell.fill = header_fill()
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws.row_dimensions[row].height = height

def write_data_row(ws, row, values, formats=None, height=24):
    for col, v in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=v)
        cell.font = body_font(10)
        cell.fill = stripe_fill(row)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        if formats and col <= len(formats) and formats[col - 1]:
            cell.number_format = formats[col - 1]
    ws.row_dimensions[row].height = height

def add_pie_chart(ws, title, data_rows, categories_col, values_col, anchor, width=14, height=10):
    chart = PieChart()
    chart.title = title
    chart.style = 10
    chart.width = width
    chart.height = height
    data = Reference(ws, min_col=values_col, min_row=data_rows[0] - 1, max_row=data_rows[-1])
    cats = Reference(ws, min_col=categories_col, min_row=data_rows[0], max_row=data_rows[-1])
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showPercent = True
    chart.dataLabels.showCatName = True
    ws.add_chart(chart, anchor)

def add_bar_chart(ws, title, data_rows, categories_col, values_cols, anchor, width=16, height=10):
    chart = BarChart()
    chart.type = "col"
    chart.title = title
    chart.style = 10
    chart.width = width
    chart.height = height
    cats = Reference(ws, min_col=categories_col, min_row=data_rows[0], max_row=data_rows[-1])
    for vc in values_cols:
        data = Reference(ws, min_col=vc, min_row=data_rows[0] - 1, max_row=data_rows[-1])
        chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    ws.add_chart(chart, anchor)

# Sheet builders are in the full local file
# Run: python generate_storage_market_report.py

def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    print("Generating storage market report...")
    # build_overview_sheet(wb) ... etc
    wb.save(str(OUTPUT_FILE))
    print(f"[OK] Report saved: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
