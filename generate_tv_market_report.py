"""
电视机市场分析报告生成器
覆盖全球电视品牌份额、面板厂商、显示技术路线、
竞争格局、发展趋势
"""

import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.utils import get_column_letter
from pathlib import Path

# ── 输出路径 ──
OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)
OUTPUT_FILE = OUTPUT_DIR / "电视机市场分析报告.xlsx"

# ── 配色方案 ──
C = {
    "title_bg":   "2C3E50",
    "title_fg":   "FFFFFF",
    "header_bg":  "34495E",
    "header_fg":  "FFFFFF",
    "sub_bg":     "D5DBDB",
    "deep_blue":  "2C3E50",
    "accent1":    "3498DB",
    "accent2":    "E74C3C",
    "accent3":    "2ECC71",
    "light_gray": "F2F2F2",
    "white":      "FFFFFF",
}

# ── 通用样式 ──
thin_border = Border(
    left=Side(style="thin", color="B4C6E7"),
    right=Side(style="thin", color="B4C6E7"),
    top=Side(style="thin", color="B4C6E7"),
    bottom=Side(style="thin", color="B4C6E7"),
)


def _font(name="微软雅黑", size=10, bold=False, color="000000"):
    return Font(name=name, size=size, bold=bold, color=color)


def _fill(hex_color):
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")


def stripe_fill(row_idx):
    return _fill(C["light_gray"]) if row_idx % 2 == 0 else _fill(C["white"])


def set_col_widths(ws, widths: dict):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def write_title_row(ws, row, text, merge_end, height=48):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=merge_end)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = _font(size=18, bold=True, color=C["title_fg"])
    cell.fill = _fill(C["title_bg"])
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = height


def write_source_row(ws, row, text, merge_end):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=merge_end)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = _font(size=9, color="666666")
    cell.alignment = Alignment(horizontal="center", vertical="center")


def write_subtitle_row(ws, row, text, merge_end, height=28):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=merge_end)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = _font(size=12, bold=True, color=C["deep_blue"])
    cell.fill = _fill(C["sub_bg"])
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = height


def write_header_row(ws, row, headers, height=30):
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.font = _font(size=11, bold=True, color=C["header_fg"])
        cell.fill = _fill(C["header_bg"])
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws.row_dimensions[row].height = height


def write_data_row(ws, row, values, fmts=None, height=24, bold_cols=None):
    for col, v in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=v)
        cell.font = _font(size=10, bold=(bold_cols and col in bold_cols))
        cell.fill = stripe_fill(row)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        if fmts and col <= len(fmts) and fmts[col - 1]:
            cell.number_format = fmts[col - 1]
    ws.row_dimensions[row].height = height


def add_pie(ws, title, data_rows, cat_col, val_col, anchor, w=14, h=10):
    chart = PieChart()
    chart.title = title
    chart.style = 10
    chart.width = w
    chart.height = h
    data = Reference(ws, min_col=val_col, min_row=data_rows[0] - 1, max_row=data_rows[-1])
    cats = Reference(ws, min_col=cat_col, min_row=data_rows[0], max_row=data_rows[-1])
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showPercent = True
    chart.dataLabels.showCatName = True
    chart.dataLabels.showVal = False
    ws.add_chart(chart, anchor)


def add_bar(ws, title, data_rows, cat_col, val_cols, anchor, w=16, h=10):
    chart = BarChart()
    chart.type = "col"
    chart.title = title
    chart.style = 10
    chart.width = w
    chart.height = h
    cats = Reference(ws, min_col=cat_col, min_row=data_rows[0], max_row=data_rows[-1])
    for vc in val_cols:
        data = Reference(ws, min_col=vc, min_row=data_rows[0] - 1, max_row=data_rows[-1])
        chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    ws.add_chart(chart, anchor)


# ═══════════════════════════════════════════════════════════════
# Sheet 1: 市场概览
# ═══════════════════════════════════════════════════════════════
def build_overview(wb):
    ws = wb.create_sheet("市场概览", 0)
    set_col_widths(ws, {"A": 20, "B": 18, "C": 18, "D": 14, "E": 20, "F": 26, "G": 30})

    write_title_row(ws, 1, "全球电视机市场分析报告（2024-2025）", 7)
    write_source_row(ws, 2, "数据来源：Omdia、Counterpoint、AVC Revo、TrendForce 等公开报告 | 制表日期：2026年5月", 7)

    # ── 细分市场规模 ──
    write_subtitle_row(ws, 4, "一、电视市场整体规模", 7)
    write_header_row(ws, 5, [
        "指标", "2022年", "2023年", "2024年", "2025年预测", "同比变化", "说明"
    ])
    data = [
        ["全球出货量(百万台)",  202,   197,   210,   218,   0.038, "大尺寸化+新兴市场驱动复苏"],
        ["全球出货额(亿美元)",  1020,  960,   1050,  1120,  0.067, "高端化(OLED/MiniLED)拉动ASP"],
        ["平均尺寸(英寸)",      49.5,  51.2,  53.0,  55.0,  0.038, "65寸成为新主力，75寸快速渗透"],
        ["平均单价(美元)",      505,   487,   500,   514,   0.028, "大尺寸+高端技术推动单价回升"],
        ["4K+渗透率",          0.68,  0.73,  0.78,  0.83,  0.064, "8K在高端市场试水，4K为主流"],
        ["智能电视渗透率",      0.82,  0.86,  0.90,  0.93,  0.033, "智能系统成标配，内容生态为王"],
    ]
    fmts = [None, "#,##0", "#,##0", "#,##0", "#,##0", "0.0%", None]
    for i, d in enumerate(data):
        write_data_row(ws, 6 + i, d, fmts)

    # ── 各尺寸段市场 ──
    write_subtitle_row(ws, 13, "二、各尺寸段出货份额与趋势", 7)
    write_header_row(ws, 14, [
        "尺寸段", "2023年份额", "2024年份额", "趋势", "主流技术", "主力品牌", "市场特征"
    ])
    size_data = [
        ["32寸及以下",  0.18, 0.14, -0.222, "LCD",        "TCL/海信/小米",   "份额持续萎缩，新兴市场低端需求"],
        ["40-43寸",    0.15, 0.12, -0.200, "LCD",        "三星/TCL/海信",   "被50+寸替代，价格战激烈"],
        ["50-55寸",    0.25, 0.23, -0.080, "LCD/QLED",   "三星/TCL/海信",   "性价比区间，QLED渗透提升"],
        ["65寸",       0.22, 0.25,  0.136, "QLED/MiniLED", "三星/海信/TCL/LG", "新黄金尺寸，各品牌主战场"],
        ["75寸",       0.12, 0.15,  0.250, "MiniLED/OLED", "三星/海信/TCL/Sony", "快速增长，价格快速下探"],
        ["85寸+",      0.04, 0.06,  0.500, "MiniLED/OLED/MicroLED", "三星/Sony/LG", "高端市场，百寸电视成热点"],
        ["98寸/100寸+",0.01, 0.02,  1.000, "MiniLED/MicroLED", "海信/TCL/红米",  "百寸大战，价格降至万元内"],
    ]
    fmts2 = [None, "0.0%", "0.0%", "+0.0%;-0.0%", None, None, None]
    for i, d in enumerate(size_data):
        write_data_row(ws, 15 + i, d, fmts2)

    # ── 各区域市场 ──
    write_subtitle_row(ws, 23, "三、各区域电视市场特征", 7)
    write_header_row(ws, 24, [
        "区域", "出货占比", "主流尺寸", "主流技术", "增长驱动力", "头部品牌", "市场特征"
    ])
    region_data = [
        ["中国",       0.24, "65-75寸",   "MiniLED/QLED",  "以旧换新+大尺寸升级", "海信/TCL/小米/创维", "竞争最激烈，价格战主战场"],
        ["北美",       0.21, "65-75寸",   "QLED/OLED",     "高端需求+体育赛事",   "三星/LG/Vizio/Sony", "ASP最高，高端产品接受度好"],
        ["欧洲",       0.19, "55-65寸",   "OLED/QLED",     "能效标准升级",        "三星/LG/Philips/Sony","OLED偏好强，节能法规严格"],
        ["亚太(不含中日韩)", 0.16, "43-55寸", "LCD",      "新兴市场首购+换机",   "三星/LG/TCL/中国品牌", "价格敏感，LCD为主"],
        ["拉美",       0.09, "43-55寸",   "LCD",           "城镇化+消费升级",     "三星/LG/TCL/海信",    "中低端为主，增速较快"],
        ["中东非",     0.07, "32-50寸",   "LCD",           "人口增长+基础设施",   "三星/LG/TCL/海信",    "入门级市场，价格驱动"],
        ["日韩",       0.04, "55-65寸",   "OLED/QLED",     "高端换机",            "Sony/LG/三星/松下",   "OLED渗透率全球最高"],
    ]
    for i, d in enumerate(region_data):
        write_data_row(ws, 25 + i, d, [None, "0.0%", None, None, None, None, None])

    add_bar(ws, "各尺寸段出货份额变化",
            range(15, 22), 1, [2, 3], "A33", w=20, h=12)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 2: 品牌份额分析
# ═══════════════════════════════════════════════════════════════
def build_brand_sheet(wb):
    ws = wb.create_sheet("品牌份额分析")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 14, "D": 14, "E": 22, "F": 32, "G": 32})

    write_title_row(ws, 1, "电视品牌厂商份额与竞争力分析", 7)

    # ── 全球品牌份额 ──
    write_subtitle_row(ws, 3, "一、全球电视品牌出货量份额（2024年）", 7)
    write_header_row(ws, 4, [
        "品牌", "出货量份额", "营收份额", "总部", "主力产品线", "技术特点", "核心优势"
    ])
    brand_data = [
        ["三星Samsung",   0.195, 0.30, "韩国",
         "Neo QLED/QLED\nThe Frame/The Serif\nOLED(S95D/S90D)",
         "QLED量子点技术领先\nNeo QLED MiniLED背光矩阵\nQD-OLED自研面板\nTizen OS智能生态",
         "全球电视龙头18年连续第一\n品牌溢价+全技术路线覆盖\nNeo QLED高端MiniLED标杆"],
        ["海信Hisense",   0.135, 0.12, "中国",
         "ULED X/ULED\n激光电视/VIDAA\n东芝TVS品牌",
         "ULED X分区背光控制\n激光电视(三色)独有品类\n信芯X1画质芯片\n收购东芝TVS+夏普美洲",
         "全球第二，中国第一\n激光电视开创者\n体育营销(欧洲杯/世界杯)"],
        ["TCL",           0.125, 0.10, "中国",
         "C系列(MiniLED)\nQLED/雷鸟子品牌\n75/85/98寸大屏",
         "MiniLED背光技术领先\n华星光电面板垂直整合\n分区数同价位最高\n自研芯片+鸿鹄画质引擎",
         "MiniLED性价比之王\n面板-整机垂直整合降本\n大尺寸市场增速最快"],
        ["LG",            0.115, 0.16, "韩国",
         "OLED evo/C Series\nQNED/QNED MiniLED\nwebOS智能生态",
         "WOLED面板自研自供\nOLED evo亮度提升技术\nQNED(量子点+NANO Cell)\nwebOS开机无广告",
         "OLED电视绝对领导者\n面板自供成本优势\n欧洲/北美高端市场稳固"],
        ["小米Xiaomi",    0.065, 0.03, "中国",
         "Redmi电视/A系列\n小米电视S系列\nES/EA系列",
         "极致性价比策略\nPatchWall智能系统\n小爱同学AI语音\n生态链产品协同",
         "线上份额中国第一\nIoT生态协同获客\n价格优势突出"],
        ["索尼Sony",      0.045, 0.08, "日本",
         "BRAVIA XR系列\nA95L(QD-OLED)\nX95L(MiniLED)",
         "XR认知芯片画质标杆\nQD-OLED/MiniLED双线\nAcoustic Surface Audio+\nPS5/蓝光生态协同",
         "画质口碑行业标杆\nXR芯片独有优势\n影音发烧友首选"],
        ["创维Skyworth",  0.04, 0.03, "中国",
         "壁纸电视/Q系列\n酷开系统/COOLITA\nOLED/WALLPAPER",
         "壁纸电视差异化\nOLED早期布局者\n酷开系统海外版COOLITA\n自研AI画质引擎",
         "国内OLED先锋\n壁纸电视品类创新"],
        ["Vizio",         0.035, 0.03, "美国",
         "M-Series/P-Series\nV-Series/D-Series",
         "美国市场性价比品牌\nFull Array Local Dimming\nWatchFree+免费流媒体",
         "北美市场深耕\n沃尔玛渠道优势\n被沃尔玛收购整合中"],
        ["其他",          0.245, 0.15, "—",
         "飞利浦/松下/夏普\nAOC/康佳/长虹等",
         "各区域品牌差异大\n飞利浦Ambilight差异化\n松下日本/欧洲高端",
         "区域性品牌为主\n份额分散"],
    ]
    fmts = [None, "0.0%", "0.0%", None, None, None, None]
    for i, d in enumerate(brand_data):
        write_data_row(ws, 5 + i, d, fmts, height=64)

    # ── 中国市场品牌份额 ──
    write_subtitle_row(ws, 15, "二、中国市场电视品牌出货量份额（2024年）", 7)
    write_header_row(ws, 16, [
        "品牌", "出货量份额", "营收份额", "主力尺寸", "主力技术", "渠道特征", "核心策略"
    ])
    cn_data = [
        ["海信",    0.20, 0.20, "65-85寸",  "ULED X/激光",  "线上线下均衡\n体育营销",      "全品类+激光电视差异化"],
        ["小米",    0.16, 0.08, "50-65寸",  "LCD/QLED",     "线上为主\n电商爆款",          "极致性价比+IoT生态"],
        ["TCL",    0.14, 0.12, "65-98寸",  "MiniLED/QLED", "线上+线下大屏\n大尺寸优势",   "MiniLED+大尺寸双引擎"],
        ["创维",   0.10, 0.09, "55-75寸",  "OLED/壁纸",    "线下传统渠道\n工程市场",      "OLED+壁纸电视细分"],
        ["海尔",   0.07, 0.05, "43-65寸",  "LCD",          "线下+下乡\n渠道下沉",         "三四线城市+以旧换新"],
        ["长虹",   0.06, 0.04, "55-75寸",  "LCD/QLED",     "线下+运营商\n集采",           "运营商合作+区域市场"],
        ["康佳",   0.05, 0.03, "43-55寸",  "LCD",          "低价+运营商",                 "入门级+性价比"],
        ["华为",   0.05, 0.06, "55-75寸",  "QLED/智慧屏",  "线上+华为门店\n鸿蒙生态",     "鸿蒙智慧屏+多屏协同"],
        ["索尼",   0.03, 0.07, "65-85寸",  "OLED/MiniLED", "线下高端\n影音渠道",          "画质旗舰+PS5生态"],
        ["三星",   0.03, 0.06, "65-85寸",  "QLED/OLED",    "线下高端\n品牌溢价",          "全球品牌+Neo QLED"],
        ["其他",   0.11, 0.08, "—",        "—",            "—",                           "区域品牌/白牌"],
    ]
    for i, d in enumerate(cn_data):
        write_data_row(ws, 17 + i, d, [None, "0.0%", "0.0%", None, None, None, None])

    # 饼图
    add_pie(ws, "全球电视品牌出货量份额", list(range(5, 14)), 1, 2, "A29", w=16, h=11)
    add_pie(ws, "中国电视品牌出货量份额", list(range(17, 28)), 1, 2, "I29", w=16, h=11)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 3: 面板厂商分析
# ═══════════════════════════════════════════════════════════════
def build_panel_sheet(wb):
    ws = wb.create_sheet("面板厂商分析")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 14, "D": 16, "E": 24, "F": 32, "G": 32})

    write_title_row(ws, 1, "电视面板厂商份额与技术分析", 7)

    # ── LCD面板份额 ──
    write_subtitle_row(ws, 3, "一、LCD 电视面板厂商出货面积份额（2024Q3）", 7)
    write_header_row(ws, 4, [
        "厂商", "面积份额", "出货量份额", "世代线", "主力产品", "技术特点", "核心优势"
    ])
    lcd_data = [
        ["京东方BOE",     0.25, 0.24, "G10.5/G8.6\n/G8.5/G6",
         "ADS Pro面板\n大尺寸4K/8K\nOLED(柔性/刚屏)",
         "ADS Pro硬屏技术\n广视角+高色域+低反射\nG10.5大尺寸产能最大\nOLED柔性屏国内唯一量产",
         "全球面板出货面积第一\n大尺寸产能最大\n客户覆盖全球主流品牌"],
        ["华星光电CSOT",  0.20, 0.20, "G11/G8.6\n/G8.5/G6",
         "HVA面板\nMiniLED背光面板\n印刷OLED开发",
         "HVA软屏高对比度\nMini LED背光规模化\n与TCL垂直整合\nt7/t9产线大尺寸",
         "TCL集团垂直整合\nMiniLED面板份额领先\n大尺寸成本优势"],
        ["惠科HKC",       0.15, 0.16, "G8.6/G8.5",
         "主流尺寸LCD\n高性价比面板",
         "G8.6产能规模大\n成本控制能力强\n快速扩张中",
         "第三大LCD面板厂\n价格竞争力强\n中小尺寸出货量大"],
        ["群创Innolux",   0.10, 0.10, "G8.5/G6",
         "GOA面板\n高端商用/车载",
         "GOA窄边框技术\n车用面板占比高\n利基型产品线",
         "车用/商用面板差异化\n利基市场盈利能力好"],
        ["友达AUO",       0.08, 0.08, "G8.5/G6",
         "AERO面板\n高端电竞/商用",
         "AERO轻薄设计\nAmLED MiniLED背光\n电竞高刷新率面板",
         "高端/利基差异化\n电竞面板份额高\nAmLED技术领先"],
        ["LGD(LG Display)",0.10, 0.08, "G8.5/P8\n/P9(OLED)",
         "WOLED电视面板\nIPS面板",
         "WOLED面板全球唯一量产\nOLED evo技术(MLA)\nIPS硬屏传统优势\nP10产线转型OLED",
         "OLED电视面板垄断供应\n大尺寸OLED技术最成熟\n逐步退出LCD聚焦OLED"],
        ["SDC(三星显示)",  0.07, 0.06, "G8.5/Q1\n(QD-OLED)",
         "QD-OLED面板\n(已退出LCD)",
         "QD-OLED自发光技术\n蓝色OLED+量子点转换\n高色域+高亮度\n已全面退出LCD生产",
         "QD-OLED技术独家\n色域覆盖行业最广\n三星集团品牌协同"],
        ["其他",           0.05, 0.08, "—",
         "夏普(堺)等",
         "夏普IGZO技术\n堺10代线部分运营",
         "利基/特殊需求"],
    ]
    fmts = [None, "0.0%", "0.0%", None, None, None, None]
    for i, d in enumerate(lcd_data):
        write_data_row(ws, 5 + i, d, fmts, height=64)

    # ── OLED面板 ──
    write_subtitle_row(ws, 14, "二、OLED 电视面板技术对比", 7)
    write_header_row(ws, 15, [
        "技术路线", "代表厂商", "发光方式", "峰值亮度", "色域", "寿命", "特点"
    ])
    oled_data = [
        ["WOLED(白光OLED)",  "LG Display", "白光OLED+彩色滤光片",
         "~2100nits(MLA)", "DCI-P3 ~99%", "10万小时+",
         "技术最成熟，量产良率高\n亮度逐年提升(MLA微透镜)\n大尺寸量产100英寸\nLG独家供应"],
        ["QD-OLED(量子点OLED)", "三星显示", "蓝色OLED+量子点转换",
         "~2500nits", "BT.2020 ~90%", "10万小时+",
         "色域更广，色彩更鲜艳\n蓝色OLED衰减挑战\n三星/Sony品牌使用\n技术迭代快，亮度提升大"],
        ["印刷OLED(Printed OLED)", "京东方/华星", "溶液法印刷RGB发光材料",
         "~1500nits(研发)", "DCI-P3 ~98%", "验证中",
         "大幅降低制造成本\n材料利用率远高于蒸镀\n大尺寸量产优势\n尚在研发/试产阶段"],
    ]
    for i, d in enumerate(oled_data):
        write_data_row(ws, 16 + i, d, height=48)

    # ── 面板供应链 ──
    write_subtitle_row(ws, 20, "三、电视面板供应链关键环节", 7)
    write_header_row(ws, 21, [
        "环节", "说明", "代表厂商", "国产化程度", "壁垒", "价值占比", "趋势"
    ])
    chain_data = [
        ["玻璃基板",    "面板核心基材，决定世代线面积",  "康宁/AGC/电气硝子\n东旭光电/彩虹股份", "低(高端)\n中(常规)", "高",   "10-15%",
         "国产化加速，10.5代突破中"],
        ["偏光片",      "控制光线偏振，影响对比度",      "日东/住友/三星SDI\n三利谱/杉金光电",     "中",               "中高", "5-8%",
         "国产化率超40%，大尺寸突破中"],
        ["背光模组",    "LCD光源系统，MiniLED核心升级点", "瑞仪/东贝/聚积\n兆驰/鸿利智汇",         "高",               "中",   "15-20%",
         "MiniLED驱动IC/灯珠是关键"],
        ["驱动IC",      "面板驱动与时序控制",             "联咏/奇景/天鈺\n集创北方/韦尔股份",     "中",               "高",   "5-8%",
         "4K/8K+高刷驱动IC需求增长"],
        ["LED芯片",     "背光/显示发光芯片",              "三安/兆驰/兆元\n首尔半导体/Nichia",    "高",               "中",   "3-5%",
         "MiniLED用芯片需求爆发"],
        ["OLED材料",    "有机发光材料(红绿蓝)",           "默克/UDC/出光\n莱特光电/奥来德",         "低",               "极高", "8-12%",
         "国产材料替代加速，蓝光材料是难点"],
    ]
    for i, d in enumerate(chain_data):
        write_data_row(ws, 22 + i, d, height=36)

    add_pie(ws, "LCD电视面板面积份额", list(range(5, 13)), 1, 2, "A29", w=16, h=11)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 4: 显示技术对比
# ═══════════════════════════════════════════════════════════════
def build_tech_sheet(wb):
    ws = wb.create_sheet("显示技术对比")
    set_col_widths(ws, {"A": 16, "B": 16, "C": 16, "D": 16, "E": 16, "F": 20, "G": 20, "H": 20})

    write_title_row(ws, 1, "电视显示技术综合对比", 8)

    # ── 技术基本参数 ──
    write_subtitle_row(ws, 3, "一、各显示技术基本参数对比", 8)
    write_header_row(ws, 4, [
        "参数", "LCD(QLED)", "MiniLED", "WOLED", "QD-OLED", "MicroLED", "激光电视", "说明"
    ])
    params = [
        ["发光方式",     "背光+滤光", "背光+滤光\n(分区控光)", "自发光(白光)", "自发光(蓝+QD)", "自发光(InGaN)", "反射式投影", "自发光对比度优势大"],
        ["峰值亮度",     "500-800nits","1000-4000nits","1500-2100nits","1500-2500nits","4000+nits","300-500nits", "HDR亮度Mini/MicroLED领先"],
        ["对比度",       "3000:1(VA)\n1200:1(IPS)","50000:1+\n(千级分区)","无穷:1\n(像素级)","无穷:1\n(像素级)","无穷:1\n(像素级)","300:1\n(环境光影响)","自发光无穷对比度"],
        ["色域(DCI-P3)", "~95%",  "~97%",  "~99%",  "~99%+\n(BT.2020 90%)","~99%+\n(BT.2020 95%)","~107%\n(激光纯度高)","激光色域最广"],
        ["可视角度",     "IPS:178度\nVA:160度","同LCD",  "178度", "178度", "178度", "受限(投影)", "自发光视角均匀"],
        ["响应时间",     "4-8ms", "4-8ms", "0.1ms", "0.1ms", "<0.01ms", "8-16ms", "OLED/MicroLED极速"],
        ["刷新率",       "60-120Hz","120-240Hz","120Hz","120-240Hz","120Hz+", "60-120Hz","高刷对游戏/体育重要"],
        ["使用寿命",     "10万小时+","10万小时+","6-10万小时","6-10万小时","10万小时+","2万小时\n(光源)", "OLED烧屏风险改善中"],
        ["典型功耗",     "80-150W","100-200W","80-150W","80-150W","150-300W","200-350W","OLED低亮度省电"],
        ["65寸参考价格", "2000-4000","3000-8000","7000-15000","8000-18000","50万+","6000-15000","MiniLED性价比突出"],
    ]
    for i, d in enumerate(params):
        write_data_row(ws, 5 + i, d, height=36)

    # ── 技术路线演进 ──
    write_subtitle_row(ws, 16, "二、各显示技术路线演进（2024-2027）", 8)
    write_header_row(ws, 17, [
        "技术", "2024现状", "2025进展", "2026-2027展望", "核心厂商", "挑战", "确定性", "定位"
    ])
    route_data = [
        ["QLED(光致)",  "LCD背光+量子点膜\n主流中高端标配",
         "量子点材料升级\n色域+亮度持续提升",
         "量子点效率再提升\n与MiniLED深度结合",
         "三星/TCL/海信", "物理极限(背光)\n非自发光", "极高", "中高端主流"],
        ["MiniLED",     "千级分区量产\n三星/海信/TCL主推",
         "万级分区量产\n灯珠/驱动IC降本",
         "分区数持续提升\n价格下探至3000元内",
         "三星/海信/TCL\n/小米", "光晕控制\n成本下降速度", "极高", "高端性价比"],
        ["WOLED",       "LGD唯一供应\nMLA微透镜技术",
         "MLA 2.0亮度提升\n新产线产能增加",
         "印刷OLED试产\n成本有望下降30%",
         "LG Display", "蓝光材料寿命\n成本高于LCD", "高", "高端品质"],
        ["QD-OLED",     "SDC供应\n三星/Sony品牌",
         "2nd Gen亮度+30%\n寿命改善",
         "3rd Gen成本下降\n更多品牌采用",
         "三星显示", "蓝光衰减\n产能有限", "中高", "高端旗舰"],
        ["MicroLED",    "三星The Wall商用\n110寸/百万元级",
         "小尺寸(76寸)试产\n巨量转移良率提升",
         "成本下降50%+\n中大尺寸(50-65寸)突破",
         "三星/錼创/富采", "巨量转移良率\n成本极高", "中低", "超高端/商用"],
        ["激光电视",    "三色激光普及\n100寸+主流",
         "双色/三色降本\n超短焦4K方案",
         "激光+菲涅尔屏升级\n百寸内价格竞争",
         "海信/峰米/长虹", "环境光干扰\n安装复杂", "中", "百寸大屏"],
    ]
    for i, d in enumerate(route_data):
        write_data_row(ws, 18 + i, d, height=48)

    # ── HDR标准对比 ──
    write_subtitle_row(ws, 25, "三、HDR 标准对比", 8)
    write_header_row(ws, 26, [
        "标准", "发起方", "峰值亮度要求", "色域要求", "位深", "元数据", "支持品牌", "市场地位"
    ])
    hdr_data = [
        ["HDR10",     "CTA/OpenHDR", "1000+nits",  "BT.709",   "10bit", "静态",  "所有品牌", "基础标准，兼容性最好"],
        ["HDR10+",    "三星/亚马逊",  "1000-4000nits","BT.2020","10bit", "动态",  "三星/TCL/松下", "三星生态，动态元数据"],
        ["Dolby Vision","杜比",       "4000+nits",  "BT.2020",  "12bit", "动态",  "LG/Sony/海信/TCL", "高端标配，授权费争议"],
        ["HLG",       "BBC/NHK",     "1000+nits",  "BT.2020",  "10bit", "无",    "广播/日韩品牌", "广播电视专用"],
        ["HDR Vivid", "CUVA(中国)",   "1000+nits",  "BT.2020",  "10/12bit","动态", "海信/TCL/华为/小米", "国产HDR标准，政策推动"],
    ]
    for i, d in enumerate(hdr_data):
        write_data_row(ws, 27 + i, d, height=28)

    add_bar(ws, "各显示技术65寸参考价格区间",
            [5, 6, 7, 8, 9, 10], 1, [6], "A33", w=18, h=10)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 5: 高端电视技术分析
# ═══════════════════════════════════════════════════════════════
def build_highend_sheet(wb):
    ws = wb.create_sheet("高端电视技术分析")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 14, "D": 16, "E": 22, "F": 30, "G": 30})

    write_title_row(ws, 1, "高端电视技术产品线深度分析", 7)

    # ── MiniLED 产品线 ──
    write_subtitle_row(ws, 3, "一、MiniLED 电视主力产品对比（2024-2025）", 7)
    write_header_row(ws, 4, [
        "产品", "品牌", "尺寸", "背光分区", "峰值亮度", "核心特点", "参考价格"
    ])
    mini_led = [
        ["QN90D",           "三星", "65/75/85寸", "~2000分区", "~2000nits",
         "Neo QLED矩阵控光\nAI画质升级\n抗反射屏\nDolby Atmos",
         "65寸: ~8000元"],
        ["U8N",             "海信", "65/75寸",    "~5000分区", "~3000nits",
         "ULED X分区精控\n信芯X1 Pro\n黑曜屏Pro低反射\n144Hz高刷",
         "65寸: ~6000元"],
        ["C855/C755",       "TCL",  "65/75/85寸", "~5000分区", "~3000nits",
         "Turbo MiniLED\n万象分区II\n超薄一体化\n2.1.2声道",
         "65寸: ~5000元"],
        ["S Pro 85",        "小米",  "85寸",      "~2000分区", "~2000nits",
         "大师画质2.0\n4K 144Hz\n双路HDMI 2.1\n智能联动",
         "85寸: ~7000元"],
        ["X95L",            "索尼",  "65/75/85寸", "~1000分区", "~1500nits",
         "XR认知芯片\nXR Backlight Master Drive\nAcoustic Multi-Audio\nPS5完美搭档",
         "65寸: ~12000元"],
        ["QM891G",          "Vizio","75寸",       "~2000分区", "~3000nits",
         "Quantum Pro\nActive Full Array\n120Hz+VRR\nWatchFree+",
         "75寸: ~$1500"],
    ]
    for i, d in enumerate(mini_led):
        write_data_row(ws, 5 + i, d, height=36)

    # ── OLED 产品线 ──
    write_subtitle_row(ws, 12, "二、OLED 电视主力产品对比（2024-2025）", 7)
    write_header_row(ws, 13, [
        "产品", "品牌", "尺寸", "面板类型", "峰值亮度", "核心特点", "参考价格"
    ])
    oled_products = [
        ["S95D",            "三星", "65/75寸",  "QD-OLED 2nd Gen", "~2500nits",
         "QD-OLED最亮\nOLED Glare Free防眩\n144Hz游戏增强\nDolby Atmos",
         "65寸: ~15000元"],
        ["C4",              "LG",   "65/75/83寸","WOLED evo(MLA)",  "~1800nits",
         "a9 AI处理器\nwebOS 24\n4K 120Hz/144Hz(VRR)\n0.1ms响应",
         "65寸: ~12000元"],
        ["G4",              "LG",   "65/77/83寸","WOLED evo(MLA+)",  "~2100nits",
         "MLA微透镜阵列\nOne Wall设计(壁挂)\na11 AI处理器\n5年webOS更新",
         "65寸: ~18000元"],
        ["A95L",            "索尼", "65/77寸",  "QD-OLED 2nd Gen", "~2000nits",
         "XR认知芯片+QD-OLED\nAcoustic Surface Audio+\nXR Triluminos Max\nPS5/蓝光旗舰",
         "65寸: ~20000元"],
        ["U8N OLED",        "海信", "65/75寸",  "WOLED evo",       "~1500nits",
         "信芯X1 OLED优化\n黑曜屏Pro\nHDR Vivid支持\n中国OLED性价比",
         "65寸: ~10000元"],
    ]
    for i, d in enumerate(oled_products):
        write_data_row(ws, 14 + i, d, height=36)

    # ── 画质芯片对比 ──
    write_subtitle_row(ws, 20, "三、电视画质芯片对比", 7)
    write_header_row(ws, 21, [
        "芯片", "品牌", "代际", "AI算力", "核心功能", "差异化", "搭载产品"
    ])
    chip_data = [
        ["XR认知芯片",     "索尼", "3rd Gen", "~22 TOPS",
         "认知智能画质\n人眼焦点模拟\nXR OLED对比度增强\nXR 4K倍线",
         "画质口碑行业标杆\n人眼认知模型独有\n跨品类(PS5/蓝光)协同",
         "A95L/X95L/A80L"],
        ["信芯X1 Pro",     "海信", "2nd Gen", "~18 TOPS",
         "AI画质感知\nULED X分区控光\nAI色彩增强\n运动补偿",
         "MiniLED控光算法优化\nHDR Vivid核心支持\n体育模式优化",
         "U8N/U7K/100L8K"],
        ["NQ4 AI Gen2",    "三星", "2nd Gen", "~16 TOPS",
         "AI画质升级\n4K AI倍线\n量子点HDR\n抗反射优化",
         "三星生态整合\nTizen OS深度优化\nQ-Symphony音频",
         "QN90D/S95D/QN85D"],
        ["a11 AI",         "LG",   "2nd Gen", "~14 TOPS",
         "AI画质Pro\nOLED动态调光\nAI声音增强\nwebOS优化",
         "OLED面板算法深度调优\nwebOS无广告\n游戏增强模式",
         "G4/B4/C4"],
        ["鸿鹄868",        "华为", "1st Gen", "~10 TOPS",
         "AI HDR增强\n鸿鹄画质引擎\n多屏协同优化\n鸿蒙系统",
         "鸿蒙生态深度整合\n手机/平板/PC多屏协同\n分布式能力",
         "智慧屏V5系列"],
        ["Pentonic 700",   "联发科","—",      "—",
         "通用电视SoC\n4K 120Hz解码\nAI SR超分\nMEMC运动补偿",
         "电视芯片市占率超50%\n多数品牌中端产品使用\n成本/性能均衡",
         "中端电视通用平台"],
    ]
    for i, d in enumerate(chip_data):
        write_data_row(ws, 22 + i, d, height=40)

    # ── 智能系统对比 ──
    write_subtitle_row(ws, 29, "四、电视智能系统生态对比", 7)
    write_header_row(ws, 30, [
        "系统", "品牌", "基础", "应用数量", "开机广告", "核心差异", "趋势"
    ])
    os_data = [
        ["Tizen OS",    "三星", "自研Linux",  "丰富",   "无",  "三星生态互联\nSmartThings物联网\n多设备协同",
         "与Google合作引入Play"],
        ["webOS",       "LG",   "自研Linux",  "丰富",   "无",  "webOS开源\n5年系统更新承诺\nNVIDIA G-Sync",
         "webOS对外授权给其他品牌"],
        ["Google TV",   "TCL/海信等","Android TV","丰富",  "有/可关闭","Google生态\nChromecast内置\n语音助手",
         "全球占比持续提升"],
        ["VIDAA",       "海信", "自研Linux",  "中等",   "有",  "轻量快速启动\n海信内容聚合\n聚好看教育",
         "海外市场逐步转向Google TV"],
        ["鸿蒙OS",      "华为", "自研微内核", "增长中", "无",  "多屏协同\n分布式能力\n超级终端\nHarmonyOS NEXT",
         "鸿蒙生态扩展中"],
        ["PatchWall",   "小米", "Android TV", "丰富",   "有",  "小米IoT生态\n小爱同学\n内容推荐引擎",
         "与米家深度绑定"],
        ["酷开系统",    "创维", "Android",    "中等",   "有",  "酷开影视内容\n海外COOLITA版",
         "海外推出COOLITA系统"],
    ]
    for i, d in enumerate(os_data):
        write_data_row(ws, 31 + i, d, height=32)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 6: 发展趋势与建议
# ═══════════════════════════════════════════════════════════════
def build_trends_sheet(wb):
    ws = wb.create_sheet("发展趋势与建议")
    set_col_widths(ws, {"A": 8, "B": 20, "C": 36, "D": 20, "E": 14, "F": 14, "G": 28})

    write_title_row(ws, 1, "电视机市场发展趋势与投资建议", 7)

    # ── 核心趋势 ──
    write_subtitle_row(ws, 3, "一、2025-2027年核心趋势", 7)
    write_header_row(ws, 4, [
        "序号", "趋势方向", "详细说明", "受益领域", "确定性", "影响", "关键变量"
    ])
    trends = [
        ["1", "MiniLED爆发普及",
         "MiniLED价格快速下探，2025年3000元内65寸MiniLED\n分区数从千级向万级演进，光晕控制改善\n海信/TCL/三星/小米全线布局",
         "MiniLED面板/驱动IC/灯珠", "极高", "★★★★★",
         "驱动IC成本下降速度\n灯珠良率/一致性"],
        ["2", "大尺寸化加速",
         "65寸成为新基准，75寸快速渗透\n85寸价格下探至5000元内\n98/100寸百寸大战价格降至万元内\n大尺寸驱动面板厂G10.5/G11产能消化",
         "大尺寸面板/背光/整机", "极高", "★★★★★",
         "以旧换新政策力度\n大尺寸运输/安装"],
        ["3", "OLED成本下降与渗透提升",
         "WOLED产线折旧逐步完成\n印刷OLED量产有望2026年落地\nQD-OLED 2nd/3rd Gen降本\nOLED渗透率从10%向15%+提升",
         "OLED面板/材料/设备", "高", "★★★★",
         "印刷OLED量产进度\nOLED面板降本速度"],
        ["4", "AI+电视深度融合",
         "AI画质芯片成高端标配\nAI语音/手势交互升级\nAI内容推荐与智能家居控制\nAI实时翻译/字幕生成",
         "AI芯片/智能系统/内容", "高", "★★★★",
         "AI端侧算力提升\n用户使用习惯改变"],
        ["5", "游戏电视细分赛道",
         "4K 144Hz+VRR成高端标配\nHDMI 2.1+ALLM+eARC\nNVIDIA G-Sync/AMD FreeSync\n主机(PS5/Xbox)生态绑定",
         "高刷面板/游戏芯片/接口", "高", "★★★",
         "游戏主机市场规模\n云游戏发展速度"],
        ["6", "MicroLED技术储备",
         "巨量转移良率从99.9%向99.999%突破\n三星The Wall商用线扩展\n预计2027-2028年50-65寸试产\n长期替代OLED的技术方向",
         "MicroLED芯片/巨量转移/封装", "中高", "★★★",
         "巨量转移良率\n芯片微缩化进度"],
    ]
    for i, d in enumerate(trends):
        write_data_row(ws, 5 + i, d, height=56)

    # ── 品牌竞争格局演进 ──
    write_subtitle_row(ws, 12, "二、品牌竞争格局演进预判", 7)
    write_header_row(ws, 13, [
        "品牌", "当前定位", "2024份额", "2027预测份额", "增长引擎", "核心挑战", "策略方向"
    ])
    brand_trend = [
        ["三星",   "全球第一，全技术覆盖",  "19.5%", "17-18%",
         "Neo QLED MiniLED持续领先\nQD-OLED高端差异化",
         "中国品牌大尺寸价格战\n中端市场份额被蚕食",
         "技术领先+品牌溢价维持高端"],
        ["海信",   "全球第二，中国第一",    "13.5%", "15-16%",
         "ULED X MiniLED性价比\n激光电视百寸独有\n体育营销全球化",
         "海外品牌认知度\n高端与三星/LG差距",
         "MiniLED+激光电视双驱动\n全球化品牌升级"],
        ["TCL",   "全球第三，MiniLED先锋",  "12.5%", "14-15%",
         "MiniLED份额全球第一\n华星光电垂直整合成本优势\n大尺寸(85/98寸)增速最快",
         "品牌溢价不足\n高端市场突破困难",
         "MiniLED+大尺寸规模领先\n品牌升级(子品牌雷鸟)"],
        ["LG",    "全球第四，OLED王者",     "11.5%", "10-11%",
         "OLED evo持续升级\nwebOS系统对外授权\n游戏电视专业定位",
         "LCD市场收缩\nWOLED成本高于MiniLED",
         "OLED技术护城河\nwebOS平台化"],
        ["小米",   "性价比线上之王",        "6.5%",  "7-8%",
         "MiniLED下探2000元内\nIoT生态获客\nRedmi子品牌价格覆盖",
         "品牌固化为低端\n利润率极低",
         "价格战获客+生态变现\nS Pro系列冲击中高端"],
        ["索尼",   "高端画质标杆",          "4.5%",  "3.5-4%",
         "XR芯片画质口碑\nPS5/蓝光影音生态\nQD-OLED旗舰",
         "价格高份额小\n中端产品竞争力不足",
         "坚守高端影音定位\nPS5/影音生态绑定"],
    ]
    for i, d in enumerate(brand_trend):
        write_data_row(ws, 14 + i, d, height=48)

    # ── 产业链投资方向 ──
    write_subtitle_row(ws, 21, "三、产业链投资关注方向", 7)
    write_header_row(ws, 22, [
        "方向", "逻辑", "标的类型", "时间窗口", "确定性", "风险", "风险提示"
    ])
    invest = [
        ["MiniLED产业链",
         "MiniLED渗透率从15%向40%+提升\n驱动IC/灯珠/背光模组需求爆发\n2025年MiniLED电视出货量预计翻倍",
         "驱动IC/LED芯片/背光模组", "2025-2027", "极高", "中低",
         "价格战导致利润压缩\n技术迭代速度"],
        ["大尺寸面板",
         "65寸+出货占比超50%\nG10.5/G11产线产能价值释放\n大尺寸面板ASP相对稳定",
         "面板厂/玻璃基板/偏光片", "2025-2027", "高", "中",
         "面板周期性波动\n产能过剩风险"],
        ["OLED材料与设备",
         "OLED渗透率持续提升\n印刷OLED量产带动材料/设备需求\n国产OLED材料替代空间大",
         "OLED材料/蒸镀设备/检测", "2025-2028", "中高", "中高",
         "印刷OLED量产进度不确定\n技术路线变更风险"],
        ["AI画质芯片",
         "AI芯片成高端电视差异化核心\n自研芯片品牌溢价明显\nNPU集成化趋势",
         "电视SoC/AI芯片IP", "2025-2027", "中高", "中",
         "联发科通用SoC竞争\n自研芯片投入大周期长"],
        ["激光电视/百寸大屏",
         "百寸市场快速增长\n激光电视价格下探\n菲涅尔屏/超短焦技术升级",
         "光机/镜头/屏幕/整机", "2025-2028", "中", "中高",
         "百寸LCD/MiniLED价格竞争\n激光电视安装限制"],
        ["MicroLED技术储备",
         "长期替代OLED方向\n巨量转移/芯片微缩化突破\n2027-2028年中小尺寸试产",
         "MicroLED芯片/巨量转移/封装", "2027-2030", "中低", "高",
         "技术成熟度不确定\n成本下降速度慢于预期"],
    ]
    for i, d in enumerate(invest):
        write_data_row(ws, 23 + i, d, height=48)

    # ── 综合竞争力评分 ──
    write_subtitle_row(ws, 30, "四、全球电视品牌综合竞争力评分（1-10分）", 7)
    write_header_row(ws, 31, [
        "品牌", "技术先进性", "品牌溢价", "产品线广度",
        "渠道覆盖", "面板供应链", "综合评分"
    ])
    score_data = [
        ["三星Samsung", 10, 10, 10, 9, 8, 9.4],
        ["海信Hisense", 8,  6,  9,  8, 7, 7.6],
        ["TCL",         8,  5,  8,  8, 9, 7.6],
        ["LG",          9,  8,  7,  7, 9, 8.0],
        ["小米Xiaomi",  6,  3,  6,  7, 4, 5.2],
        ["索尼Sony",    9,  9,  5,  5, 4, 6.4],
        ["创维Skyworth",6,  4,  6,  6, 5, 5.4],
        ["华为Huawei",  8,  7,  5,  6, 4, 6.0],
    ]
    for i, d in enumerate(score_data):
        write_data_row(ws, 32 + i, d, bold_cols={7})
        ws.cell(row=32+i, column=7).font = _font(size=11, bold=True, color=C["deep_blue"])

    add_bar(ws, "全球电视品牌综合竞争力评分",
            range(32, 40), 1, [2, 3, 4, 5, 6], "A41", w=22, h=13)

    return ws


# ═══════════════════════════════════════════════════════════════
# 主入口
# ═══════════════════════════════════════════════════════════════
def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    sheets = [
        ("市场概览",       build_overview),
        ("品牌份额分析",   build_brand_sheet),
        ("面板厂商分析",   build_panel_sheet),
        ("显示技术对比",   build_tech_sheet),
        ("高端电视技术分析", build_highend_sheet),
        ("发展趋势与建议", build_trends_sheet),
    ]
    for name, builder in sheets:
        print(f"  Generating: {name} ...")
        builder(wb)

    wb.save(str(OUTPUT_FILE))
    size_kb = OUTPUT_FILE.stat().st_size / 1024
    print(f"\n[OK] Report saved: {OUTPUT_FILE}")
    print(f"     Sheets: {len(sheets)}  |  Size: {size_kb:.1f} KB")


if __name__ == "__main__":
    main()
