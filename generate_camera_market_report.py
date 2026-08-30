"""
数码相机市场分析报告生成器
覆盖 无反相机、单反、卡片机、运动相机、全景/VR相机、电影机、镜头 等细分市场
厂商份额、技术特点、竞争格局、发展趋势
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
OUTPUT_FILE = OUTPUT_DIR / "数码相机市场分析报告.xlsx"

# ── 配色方案 ──
C = {
    "title_bg":    "1F3A2E",   # 深绿(相机主题)
    "title_fg":    "FFFFFF",
    "header_bg":   "2E6E4E",   # 草绿
    "header_fg":    "FFFFFF",
    "sub_bg":      "D0E8DC",   # 浅绿
    "accent1":     "4472C4",
    "accent2":     "ED7D31",
    "accent3":     "70AD47",
    "light_gray":  "F0F5F2",
    "white":       "FFFFFF",
    "deep_green":  "1F3A2E",
    "red":         "C00000",
    "green":       "008000",
}

# ── 通用样式工厂 ──
thin_border = Border(
    left=Side(style="thin", color="A8C8B6"),
    right=Side(style="thin", color="A8C8B6"),
    top=Side(style="thin", color="A8C8B6"),
    bottom=Side(style="thin", color="A8C8B6"),
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
    cell.font = _font(size=12, bold=True, color=C["deep_green"])
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
    set_col_widths(ws, {"A": 18, "B": 16, "C": 16, "D": 14, "E": 18, "F": 24, "G": 28})

    write_title_row(ws, 1, "全球数码相机市场分析报告（2024-2025）", 7)
    write_source_row(ws, 2, "数据来源：CIPA、IDC、Counterpoint、BCN、DSCM 等公开报告 | 制表日期：2026年7月", 7)

    # ── 细分市场规模 ──
    write_subtitle_row(ws, 4, "一、数码相机细分市场规模", 7)
    write_header_row(ws, 5, [
        "细分市场", "2023年(万台)", "2024年(万台)", "同比增速",
        "2025年预测(万台)", "主要应用", "市场特征"
    ])
    data = [
        ["无反相机(Mirrorless)", 410, 440, 0.073, 460, "摄影/视频/专业",  "索尼/佳能/尼康三足鼎立,主力增长"],
        ["单反相机(DSLR)",      180, 120, -0.333, 80,  "摄影/传统专业",   "加速退场,存量为主"],
        ["卡片机(Compact)",     300, 280, -0.067, 260, "入门/便携",       "高端复古卡片机回暖"],
        ["运动相机(Action)",    420, 460, 0.095, 500, "运动/Vlog/户外",   "GoPro/Insta360/DJI主导"],
        ["全景/VR相机",         80,  110, 0.375, 150, "VR/全景直播",       "Insta360独大,增长最快"],
        ["电影机(Cinema)",      25,  28,  0.120, 32,  "影视制作/广播",     "RED/索尼/ARRI主导,高价值"],
        ["拍立得(Instant)",    900, 950, 0.056, 980, "社交/趣味",         "富士/宝丽来回暖,长青"],
        ["中画幅(Medium Fmt)",  12,  14,  0.167, 16,  "商业/风光摄影",     "富士/哈苏小众高利润"],
    ]
    fmts = [None, "#,##0", "#,##0", "0.0%", "#,##0", None, None]
    for i, d in enumerate(data):
        write_data_row(ws, 6 + i, d, fmts)

    write_data_row(ws, 14, ["合计", 2327, 2402, 0.032, 2478, "—", "无反+运动+全景三引擎"],
                   fmts, bold_cols={1, 2, 3, 4, 5})

    # ── 市场集中度 ──
    write_subtitle_row(ws, 16, "二、各细分市场集中度（CR3）", 7)
    write_header_row(ws, 17, [
        "细分市场", "CR1", "CR2", "CR3", "头部厂商", "竞争格局", "进入壁垒"
    ])
    cr = [
        ["无反相机",     "32%", "62%", "88%", "索尼/佳能/尼康",     "日系三巨头主导",     "极高(光学+CMOS+生态)"],
        ["单反相机",     "40%", "70%", "92%", "佳能/尼康/宾得",     "佳能尼康双寡头退场","极高(存量生态)"],
        ["卡片机",       "30%", "50%", "68%", "佳能/索尼/富士",     "高端复古差异化",     "中(光学+品牌)"],
        ["运动相机",     "45%", "70%", "85%", "GoPro/Insta360/DJI", "Insta360挑战GoPro",   "中高(防抖+生态)"],
        ["全景/VR相机",  "75%", "90%", "96%", "Insta360/Ricoh",     "Insta360绝对主导",    "高(算法+拼接)"],
        ["电影机",       "45%", "70%", "85%", "RED/索尼/ARRI",      "三足鼎立各定位",      "极高(影视生态)"],
        ["拍立得",       "70%", "90%", "98%", "富士/宝丽来",        "双寡头复古潮流",      "中(胶片+品牌)"],
        ["中画幅",       "55%", "80%", "95%", "富士/哈苏/Phase One","富士性价比,哈苏高端", "极高(传感器+光学)"],
    ]
    for i, d in enumerate(cr):
        write_data_row(ws, 18 + i, d)

    # ── 关键趋势 ──
    write_subtitle_row(ws, 26, "三、2025年关键趋势", 7)
    write_header_row(ws, 27, ["序号", "趋势", "说明", "主要推进方", "确定性", "影响", "关键变量"])
    trends = [
        ["1", "无反全面取代单反", "单反退场加速,无反成唯一可换镜头形态\n佳能R5II/R1、尼康Z9/Z8扛起", "佳能/尼康/索尼", "极高", "★★★★★", "存量单反用户迁移速度"],
        ["2", "视频能力视频化",   "8K/6K录制,ProRes RAW\nCinemaLine与照片机融合",                  "索尼/佳能/松下", "高", "★★★★", "散热/存储/编码生态"],
        ["3", "AI对焦与识别",     "AI主体识别(人/动物/车/飞机)\n深度学习对焦预测",                    "索尼/佳能/尼康", "高", "★★★★", "AI模型/算力/数据"],
        ["4", "复古卡片机回暖",   "富士X100系列/G3X/Ricoh GR\n社交内容创作驱动复古潮流",             "富士/理光/佳能","中高","★★★",  "产能/内容创作生态"],
        ["5", "全景/VR增长最快",   "Insta360独大,运动相机融合全景\nVR/直播/自媒体驱动",                "Insta360/理光", "高", "★★★★", "VR生态/拼接算法"],
        ["6", "国产相机崛起",     "大疆/小米/松典在运动/全景/卡片\n国产替代加速",                     "大疆/小米/松典","中高","★★★",  "光学/CMOS/品牌"],
    ]
    for i, d in enumerate(trends):
        write_data_row(ws, 28 + i, d, height=44)

    add_bar(ws, "数码相机细分市场规模对比（万台）",
            range(6, 14), 1, [2, 4], "A36", w=22, h=12)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 2: 无反相机市场
# ═══════════════════════════════════════════════════════════════
def build_mirrorless_sheet(wb):
    ws = wb.create_sheet("无反相机市场分析")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 14, "D": 16, "E": 22, "F": 34, "G": 34})

    write_title_row(ws, 1, "无反相机市场厂商份额与技术分析", 7)

    # ── 无反份额 ──
    write_subtitle_row(ws, 3, "一、无反相机厂商份额（2024年）", 7)
    write_header_row(ws, 4, [
        "厂商", "出货量份额", "营收份额", "卡口系统", "主力产品", "技术特点", "核心优势"
    ])
    ml = [
        ["索尼Sony",      0.32, 0.34, "E卡口",
         "A7R V/A7 IV\nA1 II/A9 III\nZV-E10 II",
         "全画幅堆栈式Exmor RS\nAI对焦(主体识别+预测)\n3300万-5010万像素\n4K120p/8K30p\nBIONZ XR处理器",
         "全画幅无反开创者\nCMOS自研领先\nE卡口生态最开放"],
        ["佳能Canon",     0.30, 0.33, "RF卡口",
         "EOS R5 II/R5\nR1/R3/R6 II\nR50/R100",
         "背照堆栈式CMOS\n双像素对焦二代\nDIGIC X/XL处理器\n8K60 RAW/4K120\n眼控对焦回归",
         "单反霸主迁移无反\nRF卡口镜头群\n视频与照片双强"],
        ["尼康Nikon",     0.18, 0.18, "Z卡口",
         "Z9/Z8/Z6 III\nZ5 II/Zf\nZ50II/Zfc",
         "部分堆栈式CMOS\nEXPEED 7处理器\n8K60/4K120 N-RAW\n预录制功能\nZ卡口大口径",
         "Z9/Z8视频旗舰\n大直径Z卡口\n风光/新闻口碑"],
        ["富士Fujifilm",  0.10, 0.08, "X卡口/GFX",
         "X-T5/X-H2S\nX100VI\nGFX100II/50SII",
         "APS-C X-Trans CMOS V\n堆栈式/背照式\n胶片模拟19种\nGFX中画幅1.02亿像素\n6K30p 4:2:2",
         "复古设计+胶片模拟\nX100系列爆款\n中画幅平民化"],
        ["松下Panasonic", 0.04, 0.04, "L卡口",
         "LUMIX S5 II/S9\nG9 II/GH6",
         "相位差对焦升级\nL²联合(松下/徕卡/适马)\nProRes RAW内录\nM43与全画幅双线",
         "视频相机标杆\nL卡口联盟生态\nVlog用户群"],
        ["OM System",    0.03, 0.02, "M43卡口",
         "OM-1 II/OM-3\nOM-5",
         "堆栈式Live MOS\n计算摄影(三脚架高解像)\nAI检测对焦\n防抖8.5档\n紧凑轻量",
         "M43计算摄影标杆\n微距/打鸟利器\n便携高倍率"],
        ["徕卡Leica",    0.02, 0.06, "L卡口/M",
         "Q3/Q3 43\nSL3/M11\nM11-P",
         "全画幅背照式\nMaestro III处理器\n8K30p(Q3)\n手动对焦M旁轴\n德式精密光学",
         "奢侈品定位\n德式工艺+光学\n收藏级品牌溢价"],
        ["其他(适马/哈苏)",0.01,0.01,"L/H", "适马fp/SD\n哈苏X2D", "—", "利基/中画幅细分"],
    ]
    fmts = [None, "0.0%", "0.0%", None, None, None, None]
    for i, d in enumerate(ml):
        write_data_row(ws, 5 + i, d, fmts, height=64)

    # ── 全画幅旗舰对比 ──
    write_subtitle_row(ws, 14, "二、2025年全画幅旗舰对比", 7)
    write_header_row(ws, 15, [
        "规格", "A1 II", "EOS R5 II", "Z8", "X-H2S", "说明"
    ])
    flagship = [
        ["传感器",    "5010万堆栈式", "4500万背照堆栈", "4571万部分堆栈", "2616万堆栈式", "堆栈成旗舰标配"],
        ["处理器",    "BIONZ XR",     "DIGIC XL",        "EXPEED 7",         "X-Processor 5", "新一代处理器"],
        ["连拍(电子)","30fps",         "30fps",           "20-120fps",        "40fps",         "电子快门爆发"],
        ["最高视频",  "8K30p 4:2:2",   "8K60 RAW",        "8K60 N-RAW",        "6K60/4K120",    "8K成旗舰门槛"],
        ["对焦系统",  "AI主体识别",    "双像素二代+眼控", "3D跟踪+主体识别",  "AI检测+预测",   "AI对焦全面落地"],
        ["防抖",      "8.5档(协同)",   "8.5档(协同)",     "6档(机身)",        "7档(协同)",      "协同防抖成趋势"],
        ["价格$",     "6500",          "4299",            "5499",             "2499",          "旗舰价位分层"],
    ]
    for i, d in enumerate(flagship):
        write_data_row(ws, 16 + i, d, height=32)

    # ── 卡口生态对比 ──
    write_subtitle_row(ws, 24, "三、可换镜头卡口生态对比", 7)
    write_header_row(ws, 25, [
        "卡口", "厂商", "法兰距(mm)", "卡口直径", "原厂镜头数", "开放策略", "趋势"
    ])
    mount = [
        ["E卡口",   "索尼",    18, "46.1mm", "70+", "开放(适马/腾龙等)", "APS-C+全画幅,生态最开放"],
        ["RF卡口",  "佳能",    20, "54mm",   "40+", "部分开放(2024起)",   "近期放开第三方自动对焦"],
        ["Z卡口",   "尼康",    16, "55mm",   "45+", "部分开放",            "大直径短法兰距,光学潜力大"],
        ["X卡口",   "富士",    17.7, "44mm", "40+", "开放",                "APS-C主力,中画幅GFX独立"],
        ["L卡口",   "松下/徕卡/适马",20,"51.6mm","60+","联盟开放",         "L²联盟,适马/徕马协同"],
        ["M43卡口", "OM/松下", 19.25,"38mm", "60+", "完全开放",           "M43联盟,便携长焦优势"],
        ["M卡口",   "徕卡",    27.8,"48mm",  "30+", "封闭(徕卡)",          "手动旁轴,收藏级"],
    ]
    for i, d in enumerate(mount):
        write_data_row(ws, 26 + i, d, height=32)

    add_pie(ws, "无反相机出货量份额", [5, 6, 7, 8, 9, 10, 11, 12], 1, 2, "A34", w=16, h=11)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 3: 运动与全景相机
# ═══════════════════════════════════════════════════════════════
def build_action_panorama_sheet(wb):
    ws = wb.create_sheet("运动与全景相机")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 14, "D": 18, "E": 24, "F": 34, "G": 30})

    write_title_row(ws, 1, "运动相机与全景相机市场分析", 7)

    # ── 运动相机份额 ──
    write_subtitle_row(ws, 3, "一、运动相机厂商份额（2024年）", 7)
    write_header_row(ws, 4, [
        "厂商", "出货量份额", "营收份额", "主力平台", "主力产品", "技术特点", "核心优势"
    ])
    action = [
        ["GoPro",       0.32, 0.38, "GoPro生态",
         "HERO13 Black\nHERO12/Max2",
         "1/1.9英寸传感器\nHyperSmooth 6.0防抖\n5.3K60p/4K120\nGP2定制芯片\n空气动力学磁吸配件",
         "运动相机开创者\n北美/Vlog忠诚\n生态+剪辑软件"],
        ["Insta360",   0.28, 0.30, "全景+运动",
         "X4/X3/X5\nAce Pro 2/GO 3S\nFlow云台",
         "8K全景拼接\nFlowState防抖\nAI自动剪辑(锋芒)\n隐形自拍杆\n模块化设计",
         "全景绝对领先\nAI剪辑差异化\n运动+全景双线"],
        ["大疆DJI",    0.22, 0.20, "Osmo生态",
         "Osmo Action 5 Pro\nOsmo Pocket 3\nDJI Mic",
         "1/1.3英寸大底\n双原生ISO\nRockSteady/Hyperlapse\n磁吸快拆\n云台+相机一体",
         "无人机+相机协同\n大底画质领先\n云台技术迁移"],
        ["AKASO",      0.08, 0.04, "性价比运动",
         "Brave 7/8 Elite\nSea系列水下",
         "4K60p双屏\nIPX8防水\n性价比方案",
         "中低端量大\n性价比路线\n水下细分"],
        ["理光Ricoh",  0.05, 0.04, "Theta全景",
         "Theta X/X2\nWG系列三防",
         "360°全景轻量化\n大底1/2英寸\n防抖+直播\n三防卡片",
         "全景老牌\n商务/地产直播\n三防细分"],
        ["索尼Sony",    0.03, 0.04, "RX0/Vlog",
         "RX0 II\nZV系列Vlog",
         "1英寸大底\n蔡司镜头\n专业级画质",
         "画质路线\n专业Vlog细分"],
        ["其他",        0.02, 0.00, "—", "—", "—", "白牌/区域品牌"],
    ]
    fmts = [None, "0.0%", "0.0%", None, None, None, None]
    for i, d in enumerate(action):
        write_data_row(ws, 5 + i, d, fmts, height=56)

    # ── 全景相机份额 ──
    write_subtitle_row(ws, 13, "二、全景/VR 相机厂商份额（2024年）", 7)
    write_header_row(ws, 14, [
        "厂商", "出货量份额", "营收份额", "主力产品", "分辨率", "技术特点", "核心优势"
    ])
    pano = [
        ["Insta360",   0.75, 0.78, "X4/X3/Evo\nPro 2/Titan",
         "8K30p/11K(拼接)\n双镜头360°",
         "FlowState防抖\nAI自动剪辑/隐形自拍杆\nFlow云台+剪辑生态\nPro级拼接",
         "全景绝对领先\nAI剪辑差异化\n全价位覆盖"],
        ["理光Ricoh",  0.12, 0.10, "Theta X/X2",
         "5.7K/6K双镜头",
         "全景轻量化\n大底1/2英寸\n防抖+直播",
         "全景老牌\n商务/地产直播"],
        ["Garmin/Vuze",0.05, 0.04, "Vuze XR/VIRB",
         "5.7K拼接",
         "三防全景\n无人机协同",
         "细分三防\n户外专业"],
        ["三星Samsung",0.04, 0.04, "Gear 360(停)\n360 Round",
         "4K双镜头",
         "商务/流媒体\nIP防尘",
         "企业级应用"],
        ["其他",       0.04, 0.04, "—", "—", "—", "区域/利基品牌"],
    ]
    for i, d in enumerate(pano):
        write_data_row(ws, 15 + i, d, fmts, height=44)

    # ── 运动相机对比 ──
    write_subtitle_row(ws, 21, "三、2025年运动相机旗舰对比", 7)
    write_header_row(ws, 22, [
        "规格", "HERO13 Black", "Ace Pro 2", "Action 5 Pro", "X4(全景)", "说明"
    ])
    act_cmp = [
        ["传感器",   "1/1.9英寸",    "1/1.28英寸",  "1/1.3英寸",    "1/2英寸双镜头",  "大底成趋势"],
        ["最高视频", "5.3K60/4K120",  "8K30/4K120",  "4K120 HDR",    "8K30全景",       "8K逐步下放"],
        ["防抖",     "HyperSmooth 6", "FlowState",   "RockSteady",   "FlowState全景",  "电子防抖标配"],
        ["续航",     "70min(5.3K)",   "150min(4K)",  "180min(4K)",   "75min(8K)",      "大电池+功耗优化"],
        ["防水",     "10m裸机",       "10m裸机",     "15m裸机",      "10m裸机",        "裸机防水标配"],
        ["AI功能",   "AutoHighlight", "AI自动剪辑",  "AI构图",       "AI剪辑隐形杆",   "AI剪辑成差异化"],
        ["价格$",   "399",           "499",         "429",          "499",            "旗舰价位400-500"],
    ]
    for i, d in enumerate(act_cmp):
        write_data_row(ws, 23 + i, d, height=32)

    add_pie(ws, "运动相机出货量份额", [5, 6, 7, 8, 9, 10, 11], 1, 2, "A30", w=15, h=11)
    add_pie(ws, "全景相机出货量份额", [15, 16, 17, 18, 19], 1, 2, "I30", w=15, h=11)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 4: 电影机与专业视频
# ═══════════════════════════════════════════════════════════════
def build_cinema_sheet(wb):
    ws = wb.create_sheet("电影机市场分析")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 16, "D": 18, "E": 24, "F": 34, "G": 28})

    write_title_row(ws, 1, "电影机与专业视频市场分析", 7)

    # ── 电影机厂商份额 ──
    write_subtitle_row(ws, 3, "一、电影机厂商份额（2024年）", 7)
    write_header_row(ws, 4, [
        "厂商", "市场份额", "产品定位", "主力产品", "技术特点", "核心优势", "主要市场"
    ])
    cinema = [
        ["索尼Sony",     0.35, "全画幅电影\n摄影机", "FX9/Venice 2\nBURANO/FX6 II",
         "全画幅8.6K双基础ISO\n双原生ISO(800/3200)\nXAVC/ProRes RAW\nCineAltaD祖传色彩",
         "电影机份额第一\nCineAlta色彩科学\n租赁生态成熟", "院线/剧集/广告"],
        ["RED",         0.20, "数字电影\nRAW旗舰",  "V-Raptor/V-Raptor XL\nKomodo/Komodo X",
         "8K VV全画幅\nREDCODE RAW\n19.5档动态范围\nIPP2色彩管线\n模块化紧凑",
         "RAW旗舰标杆\nKomoda普及化\n电影/广告主力", "院线/广告/剧集"],
        ["ARRI",        0.18, "顶级电影机",         "Alexa 35/LF/Mini LF\nAlexa 35多规格",
         "ALEV IV CMOS 4.6K\n17档动态范围(行业最高)\nARRIRAW/ProRes\n色域科学标杆",
         "院线电影标准\n色彩与动态范围标杆",          "院线电影/高端剧集"],
        ["松下Panasonic",0.10,"视频/M43/全画幅",     "LUMIX S1H II/GH7\nVariCam系列",
         "全画幅6K open gate\nProRes RAW内录\n相位差对焦\nV-Log/V-Gamut",
         "视频相机标杆\nL卡口生态\n性价比专业",      "剧集/广告/Vlog"],
        ["佳能Canon",    0.08, "Cinema EOS",         "C70/C300 III\nC500 II/R5C",
         "Super 35/全画幅双规格\n双像素对焦\nCinema RAW Light\n4K120",
         "照片+视频生态协同\nOEM关系",              "剧集/新闻/广告"],
        ["Blackmagic",  0.05, "性价比电影机",        "Cinema 6K/12K\nPYXIS/Pocket",
         "6K/12K全画幅\nBlackmagic RAW\nDaVinci Resolve\n性价比",
         "性价比电影机\nResolve生态\n独立制片",     "独立电影/广告"],
        ["其他(Z CAM/Kinefinity)",0.04,"模块化/国产","Z CAM E2/Kinefinity","—","模块化/国产替代","利基/国产"],
    ]
    fmts = [None, "0.0%", None, None, None, None, None]
    for i, d in enumerate(cinema):
        write_data_row(ws, 5 + i, d, fmts, height=56)

    # ── 电影机旗舰对比 ──
    write_subtitle_row(ws, 13, "二、电影机旗舰传感器对比", 7)
    write_header_row(ws, 14, [
        "规格", "Venice 2", "V-Raptor XL", "Alexa 35", "FX9 II", "说明"
    ])
    cam_cmp = [
        ["传感器尺寸", "全画幅8.6K", "8K VV全画幅",  "S35 4.6K",   "全画幅6K",   "全画幅成主流"],
        ["分辨率",     "8.6K/6K",    "8K VV",         "4.6K",       "6K",         "8K电影机普及"],
        ["动态范围",   "16档+",      "19.5档",        "17档",       "15档+",      "RED/ARRI领先"],
        ["双原生ISO",  "800/3200",   "800/6400",      "单ISO(800)", "800/12800",  "双原生ISO成标配"],
        ["编码",       "XAVC/ProRes","REDCODE RAW",   "ARRIRAW/ProRes","XAVC/ProRes RAW","RAW格式分立"],
        ["价格$",     "58000",      "40000",         "75000",       "16000",      "ARRI溢价最高"],
    ]
    for i, d in enumerate(cam_cmp):
        write_data_row(ws, 15 + i, d, height=32)

    # ── 视频格式对比 ──
    write_subtitle_row(ws, 22, "三、专业视频编码/格式对比", 7)
    write_header_row(ws, 23, [
        "编码/格式", "提出方", "压缩", "画质", "后期弹性", "代表机型", "趋势"
    ])
    codec = [
        ["REDCODE RAW",    "RED",       "小波",  "极高", "最高(IPP2)",  "V-Raptor/Komodo", "RED独有RAW"],
        ["ARRIRAW",        "ARRI",      "无压缩","极高", "最高",        "Alexa系列",       "院线标准RAW"],
        ["Cinema RAW Light","佳能",     "轻压缩","高",   "高",          "C300/C500",       "佳能轻量RAW"],
        ["Blackmagic RAW",  "Blackmagic","轻压缩","高",   "高",          "Cinema 6K/12K",   "性价比RAW+Resolve"],
        ["ProRes RAW",      "苹果",      "轻压缩","高",   "高",          "FX6/R5C/Pocket",  "生态广但苹果控制"],
        ["XAVC/XAVC HS",    "索尼",      "帧内/HEVC","高","中",          "FX9/Venice",      "索尼通用编码"],
        ["ProRes",          "苹果",      "帧内",  "高",   "中高",        "全系列",          "后期生态最广"],
        ["N-RAW",            "尼康",      "轻压缩","高",  "高",          "Z9/Z8",           "尼康电影机RAW"],
    ]
    for i, d in enumerate(codec):
        write_data_row(ws, 24 + i, d, height=28)

    add_pie(ws, "电影机市场份额", [5, 6, 7, 8, 9, 10, 11], 1, 2, "A32", w=15, h=11)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 5: 卡片机与拍立得
# ═══════════════════════════════════════════════════════════════
def build_compact_instant_sheet(wb):
    ws = wb.create_sheet("卡片机与拍立得")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 14, "D": 18, "E": 24, "F": 34, "G": 28})

    write_title_row(ws, 1, "卡片机与拍立得市场分析", 7)

    # ── 卡片机份额 ──
    write_subtitle_row(ws, 3, "一、卡片机厂商份额（2024年）", 7)
    write_header_row(ws, 4, [
        "厂商", "出货量份额", "营收份额", "主力品类", "主力产品", "技术特点", "核心优势"
    ])
    compact = [
        ["富士Fujifilm",  0.30, 0.38, "高端复古卡片",
         "X100VI/X100V\nX20/X30",
         "APS-C X-Trans V\n固定大光圈镜头(f/2)\n胶片模拟19种\n混合取景器(光/电)\n复古旁轴造型",
         "X100系列爆款\n复古潮流标杆\n社交内容创作驱动"],
        ["佳能Canon",     0.25, 0.16, "入门卡片/长焦",
         "PowerShot V10\nSX740 HS/G7 X III",
         "1英寸CMOS(G7X)\nVlog手持卡片\n4K30p\n长焦40x",
         "入门量大\nVlog卡片差异化"],
        ["索尼Sony",       0.18, 0.16, "RX高端卡片",
         "RX100 VII/RX1R II\nZV-1 II",
         "1英寸堆栈式CMOS\n24-200mm/f2.8-4.5\n蔡司镜头\n4K HDR\nVlog翻转屏",
         "RX系列标杆\n画质+便携双优"],
        ["理光Ricoh",     0.10, 0.10, "GR系列卡片",
         "GR III/GR IIIx\nWG系列三防",
         "APS-C大底卡片\n28/40mm定焦\n高画质街拍\n三防卡片线",
         "GR街拍文化\n小众忠诚市场"],
        ["松下Panasonic", 0.08, 0.06, "LX/长焦卡片",
         "LX15/LX100 II\nTZ系列长焦",
         "1英寸大底\nLeica镜头\n30x长焦便携",
         "性价比卡片\nLeica光学协同"],
        ["徕卡Leica",     0.04, 0.10, "D-Lux/Q卡片",
         "D-Lux 8/Q3 43\nV-Lux 5",
         "全画幅/1英寸大底\nLeica调校\n德式工艺",
         "奢侈品卡片\n收藏级品牌"],
        ["其他(松典/海鸟)",0.05, 0.04, "国产卡片",
         "松典/海鸟/柯达贴牌",
         "CMOS方案+复古外观\n4K/直播卡片\n性价比",
         "国产复古卡片崛起\n性价比路线"],
    ]
    fmts = [None, "0.0%", "0.0%", None, None, None, None]
    for i, d in enumerate(compact):
        write_data_row(ws, 5 + i, d, fmts, height=56)

    # ── 拍立得份额 ──
    write_subtitle_row(ws, 13, "二、拍立得(即时成像)厂商份额（2024年）", 7)
    write_header_row(ws, 14, [
        "厂商", "出货量份额", "营收份额", "主力产品", "胶片规格", "技术特点", "核心优势"
    ])
    instant = [
        ["富士Fujifilm",  0.70, 0.72, "Instax Mini 12/90\nInstax Wide/Link",
         "Instax Mini卡式\nWide宽幅",
         "彩色即时胶片\nmini卡式最普及\nLink手机打印\n复古潮流造型",
         "拍立得绝对第一\nZ世代社交爆款\n胶片+相机双收"],
        ["宝丽来Polaroid", 0.25, 0.22, "Now+/Now Gen 2\ni-Type/600型",
         "i-Type/600型方幅",
         "经典方幅胶片\n蓝牙连接+App\n即时打印\nNostalgic品牌",
         "拍立得鼻祖回归\n复古潮流品牌\n艺术/收藏"],
        ["柯达Kodak",      0.03, 0.04, "Printomatic/Step\nSmile Classic",
         "Zink无墨打印",
         "Zink热敏无墨\n手机照片打印\n性价比",
         "无墨打印细分\n性价比路线"],
        ["其他",           0.02, 0.02, "—", "—", "—", "区域/利基品牌"],
    ]
    for i, d in enumerate(instant):
        write_data_row(ws, 15 + i, d, fmts, height=44)

    # ── 卡片机趋势 ──
    write_subtitle_row(ws, 20, "三、卡片机/拍立得趋势", 7)
    write_header_row(ws, 21, [
        "趋势", "说明", "代表厂商/产品", "驱动力", "时间线", "影响", "挑战"
    ])
    ci_trend = [
        ["复古卡片回暖",  "X100VI一机难求/溢价\n复古造型+胶片模拟",   "富士/理光",         "社交内容/复古潮","2024-2027","中高","产能/需求波动"],
        ["Vlog卡片化",    "PowerShot V10/ZV-1手持卡片\n竖拍+收音一体化","佳能/索尼",         "短视频/Vlog",   "2024-2026","中",  "手机竞争/价位"],
        ["国产卡片崛起",  "松典/海鸟复古卡片机\nCMOS方案+4K直播",        "松典/海鸟",         "国产替代/性价比","2025-2027","中",  "光学/品牌/CMOS"],
        ["拍立得潮流",    "Instax Mini年轻人爆款\n社交即时分享",          "富士/宝丽来",       "Z世代社交",     "持续",      "中",  "胶片成本/数字替代"],
        ["即时打印融合",  "拍立得+数字(打印手机照片)\nLink/蓝牙",          "富士/柯达",         "数字+实体融合",  "2024-2027","中",  "胶片耗材生态"],
    ]
    for i, d in enumerate(ci_trend):
        write_data_row(ws, 22 + i, d, height=44)

    add_pie(ws, "卡片机出货量份额", [5, 6, 7, 8, 9, 10, 11], 1, 2, "A28", w=15, h=11)
    add_pie(ws, "拍立得出货量份额", [15, 16, 17, 18], 1, 2, "I28", w=15, h=11)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 6: 镜头市场
# ═══════════════════════════════════════════════════════════════
def build_lens_sheet(wb):
    ws = wb.create_sheet("镜头市场分析")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 14, "D": 16, "E": 22, "F": 34, "G": 28})

    write_title_row(ws, 1, "可换镜头市场分析", 7)

    # ── 镜头厂商份额 ──
    write_subtitle_row(ws, 3, "一、可换镜头厂商份额（2024年）", 7)
    write_header_row(ws, 4, [
        "厂商", "出货量份额", "营收份额", "主力卡口", "主力产品", "技术特点", "核心优势"
    ])
    lens = [
        ["佳能Canon",     0.30, 0.32, "RF/EF",
         "RF24-70 f2.8L\nRF50 f1.2L\nRF100-500L",
         "BR镜片(蓝光折射)\nNano USM马达\nL红圈工艺\nRF大光圈\nEF存量巨大",
         "原厂镜头份额第一\nL红圈品牌溢价\nEF生态延续"],
        ["索尼Sony",       0.25, 0.24, "E卡口",
         "FE24-70 GM II\nFE50 f1.2 GM\nFE70-200 GM II",
         "XA镜片(超高平滑)\n直驱SSM马达\nG Master系列\n呼吸效应抑制\n11片光圈圆形焦外",
         "GM系列品牌\nE卡口原厂标杆\n视频对焦优化"],
        ["尼康Nikon",      0.15, 0.15, "Z/F",
         "Z24-70 f2.8S\nZ50 f1.2S\nZ180-600",
         "ARNEO/纳米结晶涂层\nZ Noct大光圈\nSTM/SSV马达\nS-Line高端线",
         "Z卡口大直径光学\nNoct光学潜力\n风光/新闻口碑"],
        ["适马Sigma",       0.12, 0.10, "E/L/RF/Z/EF",
         "Art 35/50/85 f1.4\n28-45 f2.8 DG DN\n150-600 Sport",
         "Art系列画质标杆\nDG DN无反优化\n防尘防滴\n高性价比大光圈\n多卡口通用",
         "副厂第一\nArt画质口碑\n性价比+多卡口"],
        ["腾龙Tamron",     0.08, 0.06, "E/RF/Z/F/EF",
         "28-75 f2.8 G2\n70-180 f2.8 G2\n35-150 f2-2.8",
         "XD线性马达\n轻量化设计\n大变焦比\nDi系列\nVC光学防抖",
         "轻量化标杆\n性价比路线\n独特焦段组合"],
        ["富士Fujifilm",    0.05, 0.06, "X/GFX",
         "XF16-55 f2.8 R\nXF50 f1.0\nGF32-64 f4",
         "LM线性马达\n胶片色彩科学\n红标XM系列\n中画幅GF",
         "X卡口原厂\n胶片色彩协同\nGFX中画幅"],
        ["松下/徕卡",       0.03, 0.05, "L卡口",
         "LUMIX S 24-70 f2.8\nSummilux 50 f1.4",
         "徕卡光学设计\n双效防抖\nPana Leica联合",
         "L²联盟协同\n徕卡光学血统"],
        ["其他(松下/三阳/老蛙等)",0.02,0.02,"多卡口",
         "三阳/老蛙/唯卓",
         "手动镜头/电影镜头\n国产光学",
         "国产替代/电影镜头"],
    ]
    fmts = [None, "0.0%", "0.0%", None, None, None, None]
    for i, d in enumerate(lens):
        write_data_row(ws, 5 + i, d, fmts, height=56)

    # ── 镜头技术演进 ──
    write_subtitle_row(ws, 14, "二、镜头光学技术演进", 7)
    write_header_row(ws, 15, [
        "技术", "说明", "代表厂商/产品", "价值", "时间线", "影响", "挑战"
    ])
    lens_tech = [
        ["特殊镜片",   "XA/BR/萤石/低色散\n提升解析力抑制色散",  "索尼XA/佳能BR/尼康萤石", "画质提升",  "持续",  "高",   "工艺/成本"],
        ["线性马达",   "XD/STM/LM线性马达\n对焦速度+静音",        "适马XD/索尼SSM/富士LM",   "对焦提升",  "2024+",  "高",   "成本/体积"],
        ["呼吸抑制",   "光学对焦组设计\n视频对焦画面稳定",          "索尼GM II/佳能RF",         "视频化",    "2024+",  "中高", "光学设计复杂度"],
        ["协同防抖",   "镜头OIS+机身IBIS协同\n最高8档防抖",         "佳能/索尼/富士",           "手持提升",  "2024+",  "中高", "算法/兼容性"],
        ["大光圈化",   "f1.2/f0.95超大光圈\n景深+进光",             "RF50 f1.2/Noct f0.95",     "差异化",    "持续",  "中",   "体积/重量/成本"],
        ["国产光学",   "老蛙/唯卓/三阳\n电影镜头+大光圈",            "老蛙/唯卓/三阳",           "国产替代",  "2025+",  "中",   "光学积累/品牌"],
    ]
    for i, d in enumerate(lens_tech):
        write_data_row(ws, 16 + i, d, height=40)

    # ── 卡口开放趋势 ──
    write_subtitle_row(ws, 23, "三、副厂镜头与卡口开放趋势", 7)
    write_header_row(ws, 24, [
        "卡口", "原厂", "副厂支持", "代表副厂镜头", "开放程度", "趋势", "影响"
    ])
    openness = [
        ["E卡口",   "索尼",  "完全开放", "适马Art/腾龙G2/三阳",   "完全开放", "生态最开放",     "副厂丰富,E卡口受益"],
        ["L卡口",   "松下/徕卡/适马","联盟开放","适马Art/徕卡/松下","联盟开放","L²联盟协同",     "联盟内互通"],
        ["RF卡口",  "佳能",  "2024起部分开放", "适马/腾龙(自动对焦)","部分开放","RF逐步放开自动对焦","副厂镜头即将丰富"],
        ["Z卡口",   "尼康",  "部分开放",  "适马/腾龙(部分)",      "部分开放","Z卡口逐步放开",    "副厂镜头增加"],
        ["X卡口",   "富士",  "开放",      "适马/腾龙/三阳",       "开放",     "X卡口生态开放",   "APS-C副厂丰富"],
        ["M43卡口", "OM/松下","完全开放", "适马/腾龙/三阳",       "完全开放","M43联盟开放",     "便携长焦副厂丰富"],
    ]
    for i, d in enumerate(openness):
        write_data_row(ws, 25 + i, d, height=32)

    add_pie(ws, "可换镜头出货量份额", [5, 6, 7, 8, 9, 10, 11, 12], 1, 2, "A32", w=15, h=11)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 7: 发展趋势与建议
# ═══════════════════════════════════════════════════════════════
def build_trends_sheet(wb):
    ws = wb.create_sheet("发展趋势与建议")
    set_col_widths(ws, {"A": 8, "B": 20, "C": 36, "D": 20, "E": 14, "F": 14, "G": 28})

    write_title_row(ws, 1, "数码相机市场发展趋势与投资建议", 7)

    # ── 核心趋势 ──
    write_subtitle_row(ws, 3, "一、2025-2027年核心趋势", 7)
    write_header_row(ws, 4, [
        "序号", "趋势方向", "详细说明", "受益领域", "确定性", "影响", "关键变量"
    ])
    trends = [
        ["1", "无反全面取代单反",
         "单反退场加速,无反成唯一可换镜头形态\n佳能R5II/R1、尼康Z9/Z8扛旗\n2025年单反出货跌破100万台",
         "无反相机/镜头", "极高", "★★★★★",
         "存量单反用户迁移速度"],
        ["2", "视频化与8K普及",
         "8K录制成旗舰标配\nProRes RAW/RAW内录普及\n散热+存储成关键\n电影机与照片机融合",
         "无反/电影机/镜头", "高", "★★★★",
         "散热/存储/编码生态"],
        ["3", "AI对焦与计算摄影",
         "AI主体识别(人/动物/车/飞机)\n深度学习对焦预测\n计算摄影(全景合成/HDR/堆栈)",
         "无反/运动/中画幅", "高", "★★★★",
         "AI模型/算力/数据标注"],
        ["4", "复古卡片机回暖",
         "X100VI一机难求\n复古造型+胶片模拟\n社交内容创作驱动",
         "卡片机/拍立得", "中高", "★★★",
         "产能/复古潮流持续性"],
        ["5", "全景与运动相机融合",
         "Insta360独大,运动相机融合全景\nVR/直播/自媒体驱动\n全景增长最快",
         "运动/全景相机", "高", "★★★★",
         "VR生态/拼接算法/AI剪辑"],
        ["6", "国产相机崛起",
         "大疆(运动/云台)/松典/海鸟(卡片)\n国产镜头(老蛙/唯卓)\n国产CMOS/光学突破",
         "运动/卡片/镜头/全景", "中高", "★★★",
         "光学/CMOS/品牌积累"],
    ]
    for i, d in enumerate(trends):
        write_data_row(ws, 5 + i, d, height=56)

    # ── 产业链价值分布 ──
    write_subtitle_row(ws, 12, "二、数码相机产业链价值分布", 7)
    write_header_row(ws, 13, [
        "环节", "说明", "代表厂商", "价值占比", "壁垒", "国产化", "趋势"
    ])
    chain = [
        ["品牌与整机", "相机品牌/整机设计",   "佳能/索尼/尼康/富士/大疆", "40-45%", "极高", "低(高端品牌)",
         "日系主导,大疆突破"],
        ["图像传感器", "CMOS/堆栈式/背照式",   "索尼半导体/佳能/松下",     "20-25%", "极高", "低(高端堆栈)",
         "索尼半导体CMOS垄断"],
        ["光学镜头",   "原厂镜头/镜片",        "佳能/索尼/适马/腾龙",       "15-20%", "高",   "中(国产镜头崛起)",
         "适马/腾龙副厂壮大"],
        ["图像处理器", "ISP/处理器/算法",     "佳能DIGIC/索尼BIONZ",       "5-8%",   "高",   "低",
         "AI算法价值提升"],
        ["机械与防抖", "快门/IBIS/对焦马达",   "各原厂/精工部件",           "5-8%",   "中高", "低",
         "协同防抖成差异化"],
        ["存储介质",   "SD/CFexpress",        "雷克沙/铠侠/西部数据",       "3-5%",   "中",   "中(国产存储崛起)",
         "CFexpress成8K刚需"],
    ]
    for i, d in enumerate(chain):
        write_data_row(ws, 14 + i, d, height=36)

    # ── 投资关注方向 ──
    write_subtitle_row(ws, 21, "三、投资关注方向", 7)
    write_header_row(ws, 22, [
        "方向", "逻辑", "标的类型", "时间窗口", "确定性", "风险", "风险提示"
    ])
    invest = [
        ["无反高端升级",
         "单反用户迁移+8K视频升级\n旗舰无反持续提价",
         "佳能/索尼/尼康供应链", "2025-2027", "中高", "中",
         "手机冲击/存量饱和"],
        ["运动/全景相机",
         "运动+全景融合+AI剪辑\nInsta360/大疆双雄",
         "Insta360/大疆供应链", "2025-2027", "高", "中",
         "手机运动相机竞争/创意周期"],
        ["国产相机与镜头",
         "大疆运动相机/松典卡片/老蛙镜头\n国产替代加速",
         "大疆/老蛙/唯卓", "2025-2028", "中", "中高",
         "光学/CMOS技术差距/品牌"],
        ["视频生态",
         "8K RAW/ProRes RAW生态\n电影机下沉专业用户",
         "RED/索尼/Blackmagic", "2025-2028", "中高", "中",
         "8K内容稀缺/手机冲击"],
        ["复古潮流品类",
         "X100VI/拍立得潮流\n社交内容创作红利",
         "富士/拍立得供应链", "2025-2027", "中", "中",
         "潮流持续性/产能"],
    ]
    for i, d in enumerate(invest):
        write_data_row(ws, 23 + i, d, height=48)

    # ── 全球厂商综合评分 ──
    write_subtitle_row(ws, 29, "四、全球主要相机厂商综合竞争力评分（1-10分）", 7)
    write_header_row(ws, 30, [
        "厂商", "品牌力", "影像技术", "镜头生态", "视频能力", "市场地位", "综合评分"
    ])
    score = [
        ["索尼Sony",     9, 10, 9,  9,  10, 9.4],
        ["佳能Canon",   10, 9,  10, 9,  9,  9.4],
        ["尼康Nikon",    9, 8,  8,  9,  7,  8.2],
        ["富士Fujifilm",  9, 8,  7,  7,  8,  7.8],
        ["大疆DJI",       8, 8,  5,  8,  8,  7.4],
        ["GoPro",         8, 6,  4,  7,  7,  6.4],
        ["Insta360",      7, 7,  4,  8,  8,  6.8],
        ["松下Panasonic", 7, 7,  7,  8,  6,  7.0],
        ["RED",          7, 8,  5,  9,  6,  7.0],
        ["ARRI",         9, 9,  5,  9,  6,  7.6],
    ]
    for i, d in enumerate(score):
        write_data_row(ws, 31 + i, d, bold_cols={7})
        ws.cell(row=31+i, column=7).font = _font(size=11, bold=True, color=C["deep_green"])

    add_bar(ws, "全球主要相机厂商综合竞争力评分",
            range(31, 41), 1, [2, 3, 4, 5, 6], "A42", w=22, h=13)

    return ws


# ═══════════════════════════════════════════════════════════════
# 主入口
# ═══════════════════════════════════════════════════════════════
def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    sheets = [
        ("市场概览",           build_overview),
        ("无反相机市场分析",   build_mirrorless_sheet),
        ("运动与全景相机",     build_action_panorama_sheet),
        ("电影机市场分析",     build_cinema_sheet),
        ("卡片机与拍立得",     build_compact_instant_sheet),
        ("镜头市场分析",       build_lens_sheet),
        ("发展趋势与建议",     build_trends_sheet),
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
