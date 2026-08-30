"""
Excel 大屏展示（Dashboard）生成器
演示如何用 openpyxl 创建包含 KPI 卡片、多图表、数据表格的
专业级数据大屏，适用于月度经营、项目进度等场景
"""

import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.utils import get_column_letter
from pathlib import Path

# ── 输出路径 ──
OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)
OUTPUT_FILE = OUTPUT_DIR / "经营数据大屏.xlsx"

# ════════════════════════════════════════════════════════════════
# 配色方案 - 深色大屏风格
# ════════════════════════════════════════════════════════════════
DARK = {
    # 背景
    "bg":          "0D1B2A",   # 深海蓝背景
    "card_bg":     "1B2838",   # 卡片背景
    "header_bg":   "1B3A5C",   # 表头背景
    "row_even":    "152238",   # 偶数行
    "row_odd":     "0D1B2A",   # 奇数行
    # 文字
    "title_fg":    "FFFFFF",   # 标题白色
    "subtitle_fg": "8BB4E0",   # 副标题浅蓝
    "body_fg":     "C8D6E5",   # 正文浅灰蓝
    "dim_fg":      "6B8299",   # 次要文字
    # 强调色
    "accent_blue": "4FC3F7",   # 蓝色强调
    "accent_green":"66BB6A",   # 绿色强调（增长）
    "accent_red":  "EF5350",   # 红色强调（下降）
    "accent_amber":"FFB74D",   # 琥珀色强调
    "accent_cyan": "26C6DA",   # 青色强调
    "accent_purple":"AB47BC",  # 紫色强调
    # 图表色板
    "chart1":      "42A5F5",
    "chart2":      "66BB6A",
    "chart3":      "FFA726",
    "chart4":      "AB47BC",
    "chart5":      "26C6DA",
    "chart6":      "EF5350",
    "chart7":      "78909C",
    "chart8":      "FFEE58",
    # 边框
    "border":      "1E3A5F",
    "divider":     "2A4A6B",
}

# 浅色方案（备选）
LIGHT = {
    "bg":          "FFFFFF",
    "card_bg":     "F5F7FA",
    "header_bg":   "2E75B6",
    "row_even":    "F2F2F2",
    "row_odd":     "FFFFFF",
    "title_fg":    "1F4E79",
    "subtitle_fg": "2E75B6",
    "body_fg":     "333333",
    "dim_fg":      "999999",
    "accent_blue": "2E75B6",
    "accent_green":"008000",
    "accent_red":  "C00000",
    "accent_amber":"ED7D31",
    "accent_cyan": "0097A7",
    "accent_purple":"7B1FA2",
    "border":      "B4C6E7",
    "divider":     "CCCCCC",
}

# 当前配色
C = DARK

# ── 通用样式 ──
thin_border = Border(
    left=Side(style="thin", color=C["border"]),
    right=Side(style="thin", color=C["border"]),
    top=Side(style="thin", color=C["border"]),
    bottom=Side(style="thin", color=C["border"]),
)


def set_sheet_bg(ws, color):
    """设置工作表背景色"""
    ws.sheet_properties.tabColor = color
    # 遍历可见区域填充背景
    for row in ws.iter_rows(min_row=1, max_row=60, max_col=20):
        for cell in row:
            cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")


def set_col_widths(ws, widths: dict):
    for col_letter, w in widths.items():
        ws.column_dimensions[col_letter].width = w


def write_kpi_card(ws, start_row, start_col, title, value, unit="", change=None, width=4, height=4):
    """
    绘制 KPI 卡片（合并单元格模拟）
    - title: 指标名称
    - value: 主数值
    - unit: 单位
    - change: 同比/环比变化，正数绿色，负数红色
    - width: 卡片占几列
    - height: 卡片占几行
    """
    end_col = start_col + width - 1
    end_row = start_row + height - 1

    # 卡片背景
    card_fill = PatternFill(start_color=C["card_bg"], end_color=C["card_bg"], fill_type="solid")
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = card_fill
            cell.border = Border(
                left=Side(style="thin", color=C["border"]),
                right=Side(style="thin", color=C["border"]),
                top=Side(style="thin", color=C["border"]),
                bottom=Side(style="thin", color=C["border"]),
            )

    # 标题行
    ws.merge_cells(start_row=start_row, start_column=start_col,
                   end_row=start_row, end_column=end_col)
    title_cell = ws.cell(row=start_row, column=start_col, value=title)
    title_cell.font = Font(name="微软雅黑", size=10, color=C["subtitle_fg"])
    title_cell.fill = card_fill
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    # 数值行（合并2行）
    val_row = start_row + 1
    ws.merge_cells(start_row=val_row, start_column=start_col,
                   end_row=val_row + 1, end_column=end_col)
    display = f"{value}{unit}" if unit else str(value)
    val_cell = ws.cell(row=val_row, column=start_col, value=display)
    val_cell.font = Font(name="微软雅黑", size=22, bold=True, color=C["accent_blue"])
    val_cell.fill = card_fill
    val_cell.alignment = Alignment(horizontal="center", vertical="center")

    # 变化行
    if change is not None:
        chg_row = end_row
        ws.merge_cells(start_row=chg_row, start_column=start_col,
                       end_row=chg_row, end_column=end_col)
        arrow = "^" if change >= 0 else "v"
        color = C["accent_green"] if change >= 0 else C["accent_red"]
        chg_cell = ws.cell(row=chg_row, column=start_col,
                           value=f"{arrow} {abs(change):.1f}%")
        chg_cell.font = Font(name="微软雅黑", size=9, bold=True, color=color)
        chg_cell.fill = card_fill
        chg_cell.alignment = Alignment(horizontal="center", vertical="center")

    # 行高
    ws.row_dimensions[start_row].height = 22
    ws.row_dimensions[start_row + 1].height = 28
    ws.row_dimensions[start_row + 2].height = 28
    ws.row_dimensions[end_row].height = 20


def write_title_bar(ws, row, text, merge_end_col, height=50):
    """大屏顶部标题栏"""
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=merge_end_col)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = Font(name="微软雅黑", size=22, bold=True, color=C["title_fg"])
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = height
    # 底部装饰线
    for c in range(1, merge_end_col + 1):
        ws.cell(row=row, column=c).border = Border(
            bottom=Side(style="medium", color=C["accent_blue"])
        )


def write_section_title(ws, row, text, merge_end_col, height=32):
    """区域标题"""
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=merge_end_col)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = Font(name="微软雅黑", size=13, bold=True, color=C["accent_cyan"])
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = height
    # 左侧装饰线
    ws.cell(row=row, column=1).border = Border(
        left=Side(style="medium", color=C["accent_cyan"]),
        bottom=Side(style="thin", color=C["divider"]),
    )


def write_header_row(ws, row, headers, height=28):
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.font = Font(name="微软雅黑", size=10, bold=True, color=C["title_fg"])
        cell.fill = PatternFill(start_color=C["header_bg"], end_color=C["header_bg"], fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws.row_dimensions[row].height = height


def write_data_row(ws, row, values, formats=None, height=22):
    for col, v in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=v)
        cell.font = Font(name="微软雅黑", size=10, color=C["body_fg"])
        bg = C["row_even"] if row % 2 == 0 else C["row_odd"]
        cell.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        if formats and col <= len(formats) and formats[col - 1]:
            cell.number_format = formats[col - 1]
    ws.row_dimensions[row].height = height


def write_trend_cell(ws, row, col, value, threshold=0):
    """带颜色趋势指示的单元格"""
    cell = ws.cell(row=row, column=col, value=value)
    if isinstance(value, (int, float)):
        if value > threshold:
            cell.font = Font(name="微软雅黑", size=10, bold=True, color=C["accent_green"])
        elif value < 0:
            cell.font = Font(name="微软雅黑", size=10, bold=True, color=C["accent_red"])
        else:
            cell.font = Font(name="微软雅黑", size=10, color=C["body_fg"])
    bg = C["row_even"] if row % 2 == 0 else C["row_odd"]
    cell.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border


def add_pie_chart(ws, title, data_rows, categories_col, values_col, anchor, width=14, height=10):
    """饼图"""
    chart = PieChart()
    chart.title = title
    chart.style = 26  # 暗色风格
    chart.width = width
    chart.height = height
    data = Reference(ws, min_col=values_col, min_row=data_rows[0] - 1, max_row=data_rows[-1])
    cats = Reference(ws, min_col=categories_col, min_row=data_rows[0], max_row=data_rows[-1])
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showPercent = True
    chart.dataLabels.showCatName = True
    chart.dataLabels.showVal = False
    ws.add_chart(chart, anchor)


def add_bar_chart(ws, title, data_rows, categories_col, values_cols, anchor, width=16, height=10):
    """柱状图"""
    chart = BarChart()
    chart.type = "col"
    chart.title = title
    chart.style = 26
    chart.width = width
    chart.height = height
    cats = Reference(ws, min_col=categories_col, min_row=data_rows[0], max_row=data_rows[-1])
    for vc in values_cols:
        data = Reference(ws, min_col=vc, min_row=data_rows[0] - 1, max_row=data_rows[-1])
        chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    ws.add_chart(chart, anchor)


def add_line_chart(ws, title, data_rows, categories_col, values_cols, anchor, width=16, height=10):
    """折线图"""
    chart = LineChart()
    chart.title = title
    chart.style = 26
    chart.width = width
    chart.height = height
    cats = Reference(ws, min_col=categories_col, min_row=data_rows[0], max_row=data_rows[-1])
    for vc in values_cols:
        data = Reference(ws, min_col=vc, min_row=data_rows[0] - 1, max_row=data_rows[-1])
        chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    # 平滑曲线
    for s in chart.series:
        s.smooth = True
    ws.add_chart(chart, anchor)


# ════════════════════════════════════════════════════════════════
# Sheet 1: 经营总览大屏
# ════════════════════════════════════════════════════════════════
def build_dashboard_sheet(wb):
    ws = wb.create_sheet("经营总览", 0)
    set_sheet_bg(ws, C["bg"])

    set_col_widths(ws, {
        "A": 14, "B": 14, "C": 14, "D": 14,
        "E": 14, "F": 14, "G": 14, "H": 14,
        "I": 14, "J": 14, "K": 14, "L": 14,
    })

    # ── 1. 标题栏 ──
    write_title_bar(ws, 1, "2025年度经营数据大屏", 12)

    # 日期信息
    ws.merge_cells("A2:L2")
    date_cell = ws.cell(row=2, column=1,
                        value="数据更新: 2025-12-31 | 数据来源: 企业ERP/CRM系统 | 报告周期: 年度")
    date_cell.font = Font(name="微软雅黑", size=9, color=C["dim_fg"])
    date_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    # ── 2. KPI 卡片行 ──
    write_kpi_card(ws, 4, 1, "年度营收", "12.8", "亿", 15.3)
    write_kpi_card(ws, 4, 5, "净利润", "2.1", "亿", 8.7)
    write_kpi_card(ws, 4, 9, "订单量", "38,650", "单", 22.1)

    # 第二行 KPI
    write_kpi_card(ws, 9, 1, "客户数", "1,285", "家", 12.5)
    write_kpi_card(ws, 9, 5, "人效", "86.5", "万/人", -3.2)
    write_kpi_card(ws, 9, 9, "回款率", "94.6", "%", 2.1)

    # ── 3. 月度营收趋势（折线图数据区） ──
    write_section_title(ws, 14, "月度营收趋势", 12)

    month_headers = ["月份", "营收(万元)", "成本(万元)", "利润(万元)", "利润率"]
    write_header_row(ws, 15, month_headers)

    month_data = [
        ["1月",  8500, 6800, 1700, 0.20],
        ["2月",  7200, 5900, 1300, 0.18],
        ["3月",  9800, 7600, 2200, 0.224],
        ["4月", 10500, 8100, 2400, 0.229],
        ["5月", 11200, 8500, 2700, 0.241],
        ["6月", 12000, 9000, 3000, 0.25],
        ["7月", 10800, 8400, 2400, 0.222],
        ["8月", 11500, 8800, 2700, 0.235],
        ["9月", 12800, 9500, 3300, 0.258],
        ["10月",13500, 9900, 3600, 0.267],
        ["11月",14200,10200, 4000, 0.282],
        ["12月",15000,10600, 4400, 0.293],
    ]
    fmts = [None, "#,##0", "#,##0", "#,##0", "0.0%"]
    for i, row_data in enumerate(month_data):
        write_data_row(ws, 16 + i, row_data, fmts)

    # 折线图
    add_line_chart(ws, "月度营收/成本/利润趋势",
                   list(range(16, 28)), 1, [2, 3, 4],
                   "G14", width=22, height=12)

    # ── 4. 业务构成（饼图） ──
    write_section_title(ws, 30, "业务收入构成", 12)

    biz_headers = ["业务线", "收入(万元)", "占比"]
    write_header_row(ws, 31, biz_headers)

    biz_data = [
        ["企业软件",  45000, 0.352],
        ["云服务",    32000, 0.250],
        ["数据解决方案", 22000, 0.172],
        ["技术咨询",  15000, 0.117],
        ["运维服务",   9000, 0.070],
        ["其他",      5000,  0.039],
    ]
    fmts2 = [None, "#,##0", "0.0%"]
    for i, row_data in enumerate(biz_data):
        write_data_row(ws, 32 + i, row_data, fmts2)

    add_pie_chart(ws, "业务收入占比",
                  list(range(32, 38)), 1, 2,
                  "G30", width=18, height=12)

    # ── 5. 区域业绩对比（柱状图） ──
    write_section_title(ws, 45, "区域业绩对比", 12)

    region_headers = ["区域", "营收(万元)", "目标(万元)", "达成率"]
    write_header_row(ws, 46, region_headers)

    region_data = [
        ["华东", 42000, 40000, 1.05],
        ["华南", 35000, 38000, 0.921],
        ["华北", 28000, 30000, 0.933],
        ["西南", 15000, 16000, 0.938],
        ["华中", 12000, 14000, 0.857],
        ["东北",  8000, 10000, 0.800],
        ["西北",  6000,  8000, 0.750],
    ]
    fmts3 = [None, "#,##0", "#,##0", "0.0%"]
    for i, row_data in enumerate(region_data):
        r = 47 + i
        write_data_row(ws, r, row_data[:3], fmts3[:3])
        # 达成率用颜色标识
        write_trend_cell(ws, r, 4, row_data[3], threshold=1.0)
        ws.cell(row=r, column=4).number_format = "0.0%"

    add_bar_chart(ws, "区域营收 vs 目标",
                  list(range(47, 54)), 1, [2, 3],
                  "G45", width=20, height=12)

    return ws


# ════════════════════════════════════════════════════════════════
# Sheet 2: 项目进度看板
# ════════════════════════════════════════════════════════════════
def build_project_sheet(wb):
    ws = wb.create_sheet("项目进度看板", 1)
    set_sheet_bg(ws, C["bg"])

    set_col_widths(ws, {
        "A": 6, "B": 22, "C": 14, "D": 12, "E": 12,
        "F": 12, "G": 14, "H": 12, "I": 14, "J": 18,
    })

    write_title_bar(ws, 1, "项目进度看板", 10)

    # ── 项目进度状态统计 ──
    write_kpi_card(ws, 3, 1, "进行中", "12", "个", None, width=3)
    write_kpi_card(ws, 3, 4, "已完成", "28", "个", None, width=3)
    write_kpi_card(ws, 3, 7, "逾期项目", "3", "个", None, width=3)

    # ── 项目明细表 ──
    write_section_title(ws, 8, "重点项目明细", 10)

    headers = ["序号", "项目名称", "负责人", "状态", "进度",
               "预算(万)", "已用(万)", "进度偏差", "风险等级"]
    write_header_row(ws, 9, headers)

    projects = [
        [1, "ERP系统升级",     "张三", "进行中", 0.75, 500, 350, 0.05, "低"],
        [2, "数据中台建设",     "李四", "进行中", 0.60, 800, 520, -0.10, "中"],
        [3, "客户门户重构",     "王五", "进行中", 0.45, 300, 180, 0.02, "低"],
        [4, "AI推荐引擎",      "赵六", "延期",   0.30, 600, 420, -0.25, "高"],
        [5, "移动端App v3.0",  "钱七", "进行中", 0.85, 200, 165, 0.08, "低"],
        [6, "安全合规改造",     "孙八", "进行中", 0.50, 150,  90, -0.05, "中"],
        [7, "供应链优化",       "周九", "已完成", 1.00, 400, 380, 0.03, "低"],
        [8, "BI报表平台",      "吴十", "进行中", 0.35, 250, 120, -0.08, "中"],
    ]
    fmts = [None, None, None, None, "0%", "#,##0", "#,##0", "0.0%", None]

    for i, proj in enumerate(projects):
        r = 10 + i
        write_data_row(ws, r, proj[:7], fmts[:7])
        # 进度偏差 - 正值绿色，负值红色
        write_trend_cell(ws, r, 8, proj[7])
        ws.cell(row=r, column=8).number_format = "0.0%"
        # 风险等级着色
        risk = proj[8]
        risk_cell = ws.cell(row=r, column=9, value=risk)
        risk_colors = {"低": C["accent_green"], "中": C["accent_amber"], "高": C["accent_red"]}
        risk_cell.font = Font(name="微软雅黑", size=10, bold=True,
                              color=risk_colors.get(risk, C["body_fg"]))
        bg = C["row_even"] if r % 2 == 0 else C["row_odd"]
        risk_cell.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
        risk_cell.alignment = Alignment(horizontal="center", vertical="center")
        risk_cell.border = thin_border

    # ── 进度甘特条（用单元格背景色模拟） ──
    write_section_title(ws, 20, "项目进度可视化", 10)

    # 简化版：用数据条效果展示
    gantt_headers = ["项目名称", "进度", "1月", "2月", "3月", "4月", "5月", "6月"]
    write_header_row(ws, 21, gantt_headers)

    gantt_data = [
        ["ERP系统升级",    0.75, 0.10, 0.15, 0.20, 0.15, 0.10, 0.05],
        ["数据中台建设",   0.60, 0.05, 0.10, 0.15, 0.15, 0.10, 0.05],
        ["客户门户重构",   0.45, 0.00, 0.05, 0.10, 0.15, 0.10, 0.05],
        ["AI推荐引擎",    0.30, 0.05, 0.05, 0.05, 0.05, 0.05, 0.00],
        ["移动端App v3.0",0.85, 0.10, 0.15, 0.20, 0.20, 0.15, 0.05],
    ]

    gantt_fill = PatternFill(start_color=C["accent_blue"], end_color=C["accent_blue"], fill_type="solid")
    empty_fill = PatternFill(start_color=C["card_bg"], end_color=C["card_bg"], fill_type="solid")

    for i, gd in enumerate(gantt_data):
        r = 22 + i
        # 项目名称
        name_cell = ws.cell(row=r, column=1, value=gd[0])
        name_cell.font = Font(name="微软雅黑", size=10, color=C["body_fg"])
        name_cell.fill = PatternFill(start_color=C["card_bg"], end_color=C["card_bg"], fill_type="solid")
        name_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        name_cell.border = thin_border
        # 进度值
        pct_cell = ws.cell(row=r, column=2, value=gd[1])
        pct_cell.font = Font(name="微软雅黑", size=10, bold=True,
                             color=C["accent_green"] if gd[1] >= 0.6 else C["accent_amber"])
        pct_cell.number_format = "0%"
        pct_cell.fill = PatternFill(start_color=C["card_bg"], end_color=C["card_bg"], fill_type="solid")
        pct_cell.alignment = Alignment(horizontal="center", vertical="center")
        pct_cell.border = thin_border
        # 月度进度条
        for j in range(6):
            cell = ws.cell(row=r, column=3 + j)
            if gd[2 + j] > 0:
                cell.fill = gantt_fill
                cell.value = gd[2 + j]
                cell.font = Font(name="微软雅黑", size=8, color=C["title_fg"])
                cell.number_format = "0%"
            else:
                cell.fill = empty_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        ws.row_dimensions[r].height = 24

    return ws


# ════════════════════════════════════════════════════════════════
# Sheet 3: 人员效能看板
# ════════════════════════════════════════════════════════════════
def build_team_sheet(wb):
    ws = wb.create_sheet("人员效能看板", 2)
    set_sheet_bg(ws, C["bg"])

    set_col_widths(ws, {
        "A": 6, "B": 16, "C": 14, "D": 14, "E": 14,
        "F": 14, "G": 14, "H": 14, "I": 14, "J": 14,
    })

    write_title_bar(ws, 1, "人员效能看板", 10)

    # ── 部门人效对比 ──
    write_section_title(ws, 3, "部门人效对比", 10)

    dept_headers = ["序号", "部门", "人数", "营收(万)", "人均营收(万)",
                    "项目数", "人均项目", "满意度", "评级"]
    write_header_row(ws, 4, dept_headers)

    dept_data = [
        [1, "研发一部", 45, 5800, 128.9, 18, 0.40, 0.92, "A"],
        [2, "研发二部", 38, 4200, 110.5, 14, 0.37, 0.88, "B+"],
        [3, "咨询部",   25, 4500, 180.0, 22, 0.88, 0.95, "A+"],
        [4, "运维部",   30, 2800,  93.3, 35, 1.17, 0.85, "B"],
        [5, "销售部",   20, 8500, 425.0,  8, 0.40, 0.90, "A+"],
        [6, "测试部",   18, 1200,  66.7, 25, 1.39, 0.82, "B-"],
        [7, "产品部",   12, 2000, 166.7, 10, 0.83, 0.91, "A"],
    ]
    fmts = [None, None, "#,##0", "#,##0", "#,##0.0",
            "#,##0", "0.00", "0%", None]

    for i, dept in enumerate(dept_data):
        r = 5 + i
        write_data_row(ws, r, dept, fmts)
        # 评级着色
        grade = dept[8]
        grade_cell = ws.cell(row=r, column=9, value=grade)
        grade_colors = {
            "A+": C["accent_green"], "A": C["accent_green"],
            "B+": C["accent_blue"],  "B": C["accent_amber"],
            "B-": C["accent_red"],
        }
        grade_cell.font = Font(name="微软雅黑", size=10, bold=True,
                               color=grade_colors.get(grade, C["body_fg"]))
        bg = C["row_even"] if r % 2 == 0 else C["row_odd"]
        grade_cell.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
        grade_cell.alignment = Alignment(horizontal="center", vertical="center")
        grade_cell.border = thin_border

    # 部门人效柱状图
    add_bar_chart(ws, "部门人均营收对比(万元)",
                  list(range(5, 12)), 2, [5],
                  "A14", width=22, height=12)

    # ── 技能矩阵 ──
    write_section_title(ws, 28, "团队技能分布", 10)

    skill_headers = ["技能领域", "高级", "中级", "初级", "合计"]
    write_header_row(ws, 29, skill_headers)

    skill_data = [
        ["Java/后端",   15, 22, 18, 55],
        ["前端/Vue",    10, 16, 12, 38],
        ["数据分析",     8, 12,  6, 26],
        ["DevOps",       6, 10,  8, 24],
        ["AI/ML",        4,  6,  5, 15],
        ["安全",         3,  5,  4, 12],
    ]
    fmts2 = [None, "#,##0", "#,##0", "#,##0", "#,##0"]
    for i, sd in enumerate(skill_data):
        write_data_row(ws, 30 + i, sd, fmts2)

    add_bar_chart(ws, "技能分布(人数)",
                  list(range(30, 36)), 1, [2, 3, 4],
                  "F28", width=20, height=12)

    return ws


# ════════════════════════════════════════════════════════════════
# Sheet 4: 浅色大屏（备选方案）
# ════════════════════════════════════════════════════════════════
def build_light_dashboard(wb):
    """浅色主题大屏 - 适合打印"""
    global C
    old_C = C
    C = LIGHT

    ws = wb.create_sheet("经营概览(浅色)", 3)
    set_sheet_bg(ws, C["bg"])

    set_col_widths(ws, {"A": 18, "B": 16, "C": 16, "D": 16, "E": 16, "F": 16, "G": 16, "H": 16})

    # 标题
    write_title_bar(ws, 1, "2025年度经营概览", 8)

    # KPI
    write_kpi_card(ws, 3, 1, "年度营收", "12.8", "亿", 15.3)
    write_kpi_card(ws, 3, 5, "净利润", "2.1", "亿", 8.7)

    # 季度数据
    write_section_title(ws, 8, "季度经营数据", 8)
    q_headers = ["季度", "营收(万元)", "同比增速", "净利润(万元)", "净利率", "订单量", "客户数", "NPS"]
    write_header_row(ws, 9, q_headers)

    q_data = [
        ["Q1", 25500, 0.18, 5200, 0.204, 8200, 310, 72],
        ["Q2", 33700, 0.22, 8100, 0.240, 9800, 345, 75],
        ["Q3", 35100, 0.19, 8400, 0.239, 9600, 330, 74],
        ["Q4", 42700, 0.25, 11300, 0.265, 11050, 300, 78],
    ]
    fmts = [None, "#,##0", "0.0%", "#,##0", "0.0%", "#,##0", "#,##0", "0"]
    for i, qd in enumerate(q_data):
        write_data_row(ws, 10 + i, qd, fmts)

    # 柱状图
    add_bar_chart(ws, "季度营收与利润",
                  list(range(10, 14)), 1, [2, 4],
                  "A15", width=20, height=12)

    C = old_C
    return ws


# ════════════════════════════════════════════════════════════════
# 主入口
# ════════════════════════════════════════════════════════════════
def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # 删除默认 Sheet

    print("生成 Sheet 1: 经营总览大屏 ...")
    build_dashboard_sheet(wb)

    print("生成 Sheet 2: 项目进度看板 ...")
    build_project_sheet(wb)

    print("生成 Sheet 3: 人员效能看板 ...")
    build_team_sheet(wb)

    print("生成 Sheet 4: 经营概览(浅色) ...")
    build_light_dashboard(wb)

    wb.save(OUTPUT_FILE)
    print(f"大屏报告已生成: {OUTPUT_FILE}")
    print("包含 4 个 Sheet:")
    print("  1. 经营总览 - 深色大屏(KPI卡片 + 折线图 + 饼图 + 柱状图)")
    print("  2. 项目进度 - 甘特条 + 风险着色 + 状态表格")
    print("  3. 人员效能 - 部门对比 + 技能矩阵 + 评级着色")
    print("  4. 经营概览 - 浅色主题(适合打印)")


if __name__ == "__main__":
    main()
