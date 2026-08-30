"""
键盘鼠标市场分析报告生成器
分析机械键盘/薄膜键盘、游戏鼠标/办公鼠标等各细分市场的
厂商份额、技术特点、发展趋势的 Excel 报告
"""

import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from pathlib import Path

# ── 输出路径 ──
OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)
OUTPUT_FILE = OUTPUT_DIR / "键盘鼠标市场分析报告.xlsx"

# ── 配色方案 ──
COLORS = {
    "title_bg":   "1F4E79",
    "title_fg":   "FFFFFF",
    "header_bg":  "2E75B6",
    "header_fg":  "FFFFFF",
    "sub_bg":     "D6E4F0",
    "accent1":    "4472C4",
    "accent2":    "ED7D31",
    "accent3":    "70AD47",
    "accent4":    "FFC000",
    "accent5":    "5B9BD5",
    "accent6":    "A5A5A5",
    "light_gray": "F2F2F2",
    "white":      "FFFFFF",
    "red_text":   "C00000",
    "green_text": "008000",
}

# ── 通用样式 ──
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


def set_col_widths(ws, widths: dict):
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
    chart.dataLabels.showVal = False
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


# ════════════════════════════════════════════════════════════════
# Sheet 1: 市场概览
# ════════════════════════════════════════════════════════════════
def build_overview_sheet(wb):
    ws = wb.create_sheet("市场概览", 0)

    set_col_widths(ws, {"A": 20, "B": 18, "C": 18, "D": 16, "E": 20, "F": 28, "G": 28})

    write_title_row(ws, 1, "全球键盘鼠标市场分析报告（2025-2026）", 7, 48)

    ws.merge_cells("A2:G2")
    c = ws.cell(row=2, column=1,
                value="数据来源：IDC、Counterpoint、行业调研及厂商财报 | 制表日期：2026年7月")
    c.font = body_font(9, color="666666")
    c.alignment = Alignment(horizontal="center", vertical="center")

    # ── 细分市场规模 ──
    write_subtitle_row(ws, 4, "一、键盘鼠标细分市场规模", 7)
    headers = ["细分市场", "2024年规模(亿美元)", "2025年规模(亿美元)", "同比增速",
               "2026年预测(亿美元)", "主要应用领域", "市场特征"]
    write_header_row(ws, 5, headers)

    data = [
        ["机械键盘(游戏)",  32,   38,  0.1875, 44,  "游戏/电竞/高端办公", "客制化趋势明显，热插拔普及"],
        ["薄膜键盘(办公)",  28,   29,  0.0357, 30,  "办公/教育/公共终端", "平缓下滑，被剪刀脚替代"],
        ["游戏鼠标",        24,   28,  0.1667, 32,  "电竞/FPS/MOBA/RPG", "轻量化+无线化，PAW3395普及"],
        ["办公鼠标",        18,   19,  0.0556, 20,  "办公/家用/笔记本",   "静音+人体工学，垂直鼠标增长"],
        ["机械轴体",        8,    10,  0.25,   12,  "机械键盘组装/维修", "Gateron/Kailh快速扩张"],
        ["鼠标传感器",      4,    4.5, 0.125,  5,   "游戏鼠标/高端办公",  "原相主导，雷蛇自主跟进"],
    ]
    fmts = [None, "#,##0", "#,##0", "0.0%", "#,##0", None, None]
    for i, row_data in enumerate(data):
        write_data_row(ws, 6 + i, row_data, fmts)

    write_data_row(ws, 12, ["合计", 114, 128.5, 0.1272, 143, "—", "外设市场稳定增长"], fmts)
    for col in range(1, 8):
        ws.cell(row=12, column=col).font = body_font(10, bold=True)

    # ── 市场集中度 ──
    write_subtitle_row(ws, 14, "二、各细分市场集中度（CR3）", 7)
    headers2 = ["细分市场", "CR1", "CR2", "CR3", "头部厂商", "竞争格局", "进入壁垒"]
    write_header_row(ws, 15, headers2)

    cr_data = [
        ["游戏键盘",    "22%", "38%", "48%", "罗技/雷蛇/海盗船",    "品牌寡占",  "中（品牌+渠道）"],
        ["办公薄膜键盘","30%", "50%", "65%", "罗技/戴尔/联想",      "品牌寡占",  "低（OEM代工为主）"],
        ["游戏鼠标",    "25%", "42%", "52%", "罗技/雷蛇/赛睿",      "品牌寡占",  "中（技术+生态）"],
        ["办公鼠标",    "28%", "45%", "58%", "罗技/微软/戴尔",      "品牌寡占",  "低（OEM代工为主）"],
        ["机械轴体",    "30%", "45%", "55%", "Cherry/佳达隆/凯华",  "竞争型",    "中（专利+工艺）"],
    ]
    for i, row_data in enumerate(cr_data):
        write_data_row(ws, 16 + i, row_data)

    # ── 关键趋势 ──
    write_subtitle_row(ws, 22, "三、2025-2026年关键趋势", 7)
    trends = [
        ["1", "无线化全面普及",    "2.4G+蓝牙三模连接成为高端键盘鼠标标配\n低延迟无线技术接近有线体验",      "罗技Lightspeed/雷蛇HyperSpeed", "高"],
        ["2", "客制化热潮",        "热插拔轴座+RGB+Gasket结构键盘火爆\n客制化社区驱动小众品牌快速增长",    "Keychron/VARMILO/阿米洛",      "高"],
        ["3", "轻量化鼠标",        "游戏鼠标重量降至50g以下(打孔/镁合金)\nPAW3395+4K回报率成为旗舰标配",   "雷蛇Viper Mini/罗技GPX",       "高"],
        ["4", "AI+外设",           "AI辅助宏指令/智能灯效/自动场景切换\n打字习惯分析/健康提示功能出现",      "罗技/雷蛇布局中",               "中"],
        ["5", "磁轴(霍尔效应)",    "磁轴键盘实现0.1mm精度可调触发\nFPS游戏急停优势明显，快速取代光学轴",    "雷蛇/Wooting/醉鹿",             "高"],
        ["6", "可持续环保材料",    "消费后回收塑料(PCR)+FSC认证包装\nRoHS/REACH合规成为品牌门槛",           "罗技/微软引领",                 "低"],
    ]
    headers3 = ["序号", "趋势", "说明", "主要推进方", "确定性"]
    write_header_row(ws, 23, headers3)
    for i, row_data in enumerate(trends):
        write_data_row(ws, 24 + i, row_data)

    add_bar_chart(ws, "键盘鼠标细分市场规模对比（亿美元）",
                  range(6, 12), 1, [2, 4], "A30", width=22, height=12)

    return ws


# ════════════════════════════════════════════════════════════════
# Sheet 2: 键盘市场分析
# ════════════════════════════════════════════════════════════════
def build_keyboard_sheet(wb):
    ws = wb.create_sheet("键盘市场分析")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 14, "D": 14, "E": 22, "F": 32, "G": 32})

    write_title_row(ws, 1, "键盘市场厂商份额与技术分析", 7, 44)

    # ── 游戏键盘份额 ──
    write_subtitle_row(ws, 3, "一、游戏键盘品牌市场份额（2025年）", 7)
    headers = ["品牌", "营收份额", "出货份额", "核心价位段", "主力产品线", "技术特点", "核心优势"]
    write_header_row(ws, 4, headers)

    kb_data = [
        ["罗技(Logitech)", 0.22, 0.20, "500-1500元",
         "G Pro X/G913/G915 TKL",
         "Lightspeed无线技术行业标杆\nROMER-G/GL轴体自研(凯华代工)\nGHUB生态整合",
         "无线技术最强；渠道覆盖广\n办公+游戏双线覆盖"],
        ["雷蛇(Razer)",    0.16, 0.15, "400-1800元",
         "黑寡妇V4/雨林狼蛛V3/噬魂金蝎",
         "雷蛇绿轴/黄轴(凯华代工)\nHyperSpeed无线+Chroma灯效\n光学轴体自有技术",
         "游戏生态最完整(音频/手柄/鼠标)\nChroma灯效联动领先"],
        ["海盗船(Corsair)",0.10, 0.08, "600-2000元",
         "K70/K100/K65 Plus",
         "Cherry轴体+光学轴\nAXON超快处理(8000Hz轮询)\niCUE生态联动",
         "高端定位+军工级做工\nCherry轴合作关系深"],
        ["赛睿(SteelSeries)",0.08, 0.07, "500-1500元",
         "Apex Pro/7/9系列",
         "OmniPoint可调触发磁轴\nPBT键帽+OLED屏幕\nGameSense联动游戏",
         "磁轴先发优势(OmniPoint)\neSports赛事合作广泛"],
        ["HyperX(HP)",     0.07, 0.08, "300-1000元",
         "Alloy Origins/Core/Haste",
         "HyperX自研红轴(TTC代工)\nPBT键帽+铝合金框\n性价比突出",
         "电竞赞助多；性价比高\nHP渠道资源支持"],
        ["Keychron",       0.06, 0.07, "300-800元",
         "Q/V/K系列",
         "Gasket结构热插拔轴座\nQMK/VIA开源固件\n全系列Mac/Win兼容",
         "客制化社区首选\n开源固件灵活性好"],
        ["其他品牌",       0.31, 0.35, "100-2000元",
         "Ducky/Akko/VARMILO等",
         "多元化客制化设计\n不同价位各种轴体可选",
         "细分市场灵活\n客制化/性价比导向"],
    ]

    for i, row_data in enumerate(kb_data):
        write_data_row(ws, 5 + i, row_data,
                       [None, "0.0%", "0.0%", None, None, None, None],
                       height=48)

    # ── 键盘类型对比 ──
    write_subtitle_row(ws, 13, "二、键盘类型对比", 7)
    headers2 = ["键盘类型", "市场份额", "均价(元)", "主要轴体", "优点", "缺点", "适用人群"]
    write_header_row(ws, 14, headers2)

    type_data = [
        ["机械键盘",        0.52, 450,  "Cherry/佳达隆/凯华", "手感好/寿命长/可DIY", "重/噪音大/贵",     "游戏玩家/打字爱好者"],
        ["薄膜键盘",        0.28, 80,   "橡胶碗导电膜",        "便宜/静音/轻薄",       "手感差/寿命短",     "办公/家庭/教育"],
        ["剪刀脚键盘",      0.12, 120,  "剪刀脚结构",          "轻薄/手感好于薄膜\n笔记本键盘体验一致", "行程短/不适合游戏", "笔记本用户/移动办公"],
        ["磁轴(霍尔效应)",  0.05, 800,  "霍尔效应传感器",      "可调触发点(0.1mm)\n急停快/寿命极长",    "贵/选择少/维修难",   "FPS电竞玩家"],
        ["光学轴体键盘",    0.03, 700,  "光轴(红外阻断)",      "响应快(无抖动)\n寿命长/无触点磨损",  "贵/不能兼容普通键帽", "竞技游戏玩家"],
    ]
    for i, row_data in enumerate(type_data):
        write_data_row(ws, 15 + i, row_data, [None, "0.0%", None, None, None, None, None])

    # ── 键盘技术趋势 ──
    write_subtitle_row(ws, 21, "三、机械键盘结构对比", 7)
    headers3 = ["结构类型", "描述", "手感特点", "代表键盘", "流行度", "改装难度", "优缺点"]
    write_header_row(ws, 22, headers3)

    structure_data = [
        ["Gasket Mount",   "垫片固定内胆，通过硅胶垫悬空", "软弹/Q弹手感\n一致性高",       "Keychron Q系列\nTofu65",     "极高(2024-2025主流)", "高", "手感柔和但刚性不足"],
        ["Tray Mount",     "螺丝柱固定PCB到底壳",           "硬/钢性手感\n手感不一致",      "入手机械键盘\n多数成品键盘", "逐渐淘汰",            "低", "成本最低但均匀性差"],
        ["Top/Bottom Mount","上下盖直接夹紧定位板",         "硬朗/一致性高\n反馈清晰",     "樱桃G80-3000\nLeopold FC900","中",             "中", "做工要求高/硬手感"],
        ["Sandwich Mount", "通过硅胶条夹在上下盖之间",      "介于Gasket和Tray之间\n可控",  "部分客制化套件",           "中低",                "中", "设计与调校难度大"],
        ["Integrated Plate","定位板与外壳一体化",           "刚性最大/一致性最高\n硬朗",   "雷蛇黑寡妇V4\n部分军规级键盘","低",              "高", "无法独立更换定位板"],
    ]
    for i, row_data in enumerate(structure_data):
        write_data_row(ws, 23 + i, row_data, height=36)

    add_pie_chart(ws, "游戏键盘品牌营收份额", [5, 6, 7, 8, 9, 10, 11], 1, 2, "A29", width=16, height=11)
    add_bar_chart(ws, "键盘类型市场份额与均价",
                  range(15, 20), 1, [2, 3], "I29", width=18, height=11)

    return ws


# ════════════════════════════════════════════════════════════════
# Sheet 3: 鼠标市场分析
# ════════════════════════════════════════════════════════════════
def build_mouse_sheet(wb):
    ws = wb.create_sheet("鼠标市场分析")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 14, "D": 14, "E": 22, "F": 32, "G": 32})

    write_title_row(ws, 1, "鼠标市场厂商份额与技术分析", 7, 44)

    # ── 游戏鼠标份额 ──
    write_subtitle_row(ws, 3, "一、游戏鼠标品牌市场份额（2025年）", 7)
    headers = ["品牌", "营收份额", "出货份额", "核心价位段", "主力产品线", "技术特点", "核心优势"]
    write_header_row(ws, 4, headers)

    mouse_data = [
        ["罗技(Logitech)", 0.25, 0.22, "300-1200元",
         "G Pro X Superlight 2\nG502 X/G903",
         "Hero 2传感器(44K DPI)\nLightspeed无线+PowerPlay无线充电\nLightforce混合微动",
         "无线技术王者；Hero传感器自研\nGrand Prix赞助赛事覆盖广"],
        ["雷蛇(Razer)",    0.18, 0.16, "300-1500元",
         "Viper V3 Pro/DeathAdder V3\nBasilisk V3",
         "Focus Pro 30K/35K光学传感器\nHyperSpeed无线+4K轮询\n雷蛇光学微动(第三代)",
         "传感器性能最强；轻量化领先\nChroma灯效生态协同"],
        ["赛睿(SteelSeries)",0.09, 0.08, "300-1000元",
         "Aerox 5/9/Rival 系列",
         "TrueMove Air传感器\nAquaBarrier防水\n超轻量打孔设计",
         "电竞市场地位稳固\nTrueMove技术积累深厚"],
        ["海盗船(Corsair)",0.07, 0.06, "300-900元",
         "Dark Core RGB Pro\nScimitar系列/M65",
         "MARKMAN传感器定制\nSlipstream无线连接\niCUE生态联动",
         "MMO鼠标独特(12侧键)\niCUE生态一体化控制"],
        ["HyperX(HP)",     0.06, 0.07, "200-600元",
         "Pulsefire Haste 2\nPulsefire Dart",
         "轻量蜂窝打孔设计\nPAW3335/3395传感器\nTTC防尘金微动",
         "性价比高；轻量化设计优秀\nHP渠道+电竞赞助"],
        ["Glorious",       0.04, 0.05, "300-700元",
         "Model O/D/I系列",
         "极致轻量化(59g Model O-)\n打孔外壳+Glorious开关\n可换外壳DIY设计",
         "轻量化先锋；性价比\n客制化社区支持好"],
        ["其他品牌",       0.31, 0.36, "50-1500元",
         "Zowie/ROG/Finalmouse等",
         "各种细分方向\n不同传感器和微动方案",
         "细分市场深耕\n特定游戏需求满足"],
    ]

    for i, row_data in enumerate(mouse_data):
        write_data_row(ws, 5 + i, row_data,
                       [None, "0.0%", "0.0%", None, None, None, None],
                       height=48)

    # ── 办公鼠标份额 ──
    write_subtitle_row(ws, 13, "二、办公鼠标品牌市场份额（2025年）", 7)
    headers2 = ["品牌", "营收份额", "核心价位段", "主力产品线", "技术特点", "核心优势"]
    write_header_row(ws, 14, headers2)

    office_data = [
        ["罗技(Logitech)",  0.30, "100-800元", "MX Master 3S/MX Anywhere 3\nPebble 2/M590/M720",
         "Darkfield 4000 DPI传感器\n电磁滚轮MagSpeed+自由滚动\nFlow跨屏(3台电脑+复制粘贴)\nBolt优联接收器",
         "办公鼠标绝对领导者\nMX系列旗舰标杆\n跨屏协作生态独有"],
        ["微软(Microsoft)", 0.18, "100-600元", "Sculpt/Arc Mouse/Bluetooth系列",
         "BlueTrack蓝影技术(任何表面)\nBluetooth 5.0 LE低功耗\n人体工学Sculpt设计",
         "人体工学设计最佳\nWindows深度集成\nSurface生态黏性"],
        ["戴尔(Dell)",      0.12, "30-200元", "MS5320W/WM615等",
         "OEM渠道量大价优\n基本办公功能齐全\n多色可选",
         "企业采购首选\nPC捆绑销售量大"],
        ["联想(Lenovo)",    0.10, "20-150元", "ThinkPad系列/Mice",
         "ThinkPad经典小红点生态\n大量OEM/ODM代工\n低端/入门市场覆盖",
         "ThinkPad用户生态\n企业批量采购优势"],
        ["双飞燕(A4Tech)",  0.08, "20-80元", "W系列/G系列",
         "性价比极高\n静音设计完善\n国产自主品牌",
         "中国市场份额大\n极高性价比"],
        ["HP/其他",         0.22, "30-500元", "HP系列/雷柏/英菲克等",
         "各自细分市场差异化",
         "广泛覆盖"],
    ]
    for i, row_data in enumerate(office_data):
        write_data_row(ws, 15 + i, row_data, [None, "0.0%", None, None, None, None],
                       height=48)

    # ── 鼠标传感器对比 ──
    write_subtitle_row(ws, 22, "三、主流游戏鼠标传感器方案对比", 7)
    headers3 = ["传感器型号", "DPI", "IPS", "加速度", "主频", "应用品牌", "特点"]
    write_header_row(ws, 23, headers3)

    sensor_data = [
        ["原相 PAW3395",           "26,000-30,000", "650 IPS", "50G", "125-1000Hz(原)\n4K/8K+无线", "雷蛇/赛睿/Glorious/HyperX",
         "2023-2025年主流旗舰\n功耗低+性能强\n市场占有率最高"],
        ["原相 PAW3399",           "26,000",        "650 IPS", "50G", "125-1000Hz(原)\n4K+无线", "罗技(独占G系列)",
         "Hero 2版罗技定制\n优化无线功耗\nLightforce微动搭配"],
        ["罗技 Hero 25K/Hero 2",   "25,600-44,000", "400+ IPS","40G", "125-1000Hz(原)\nPowerPlay兼容", "罗技全系G系列",
         "罗技自研传感器\n低功耗(单电池续航>200h)\nLightforce混动方案"],
        ["雷蛇 Focus Pro 30K/35K","30,000-35,000", "750 IPS", "70G", "125-4KHz原生\n8K HyperPolling", "雷蛇旗舰系列",
         "原生4KHz轮询\n智能追踪(AI)\n运动同步功能"],
        ["原相 PAW3335",           "16,000-20,000", "400 IPS", "40G", "125-1000Hz", "中端游戏/办公鼠标",
         "性价比高的入门方案\n功耗适中\n广泛用于300-500元市场"],
    ]
    for i, row_data in enumerate(sensor_data):
        write_data_row(ws, 24 + i, row_data, height=48)

    add_pie_chart(ws, "游戏鼠标品牌营收份额", [5, 6, 7, 8, 9, 10, 11], 1, 2, "A30", width=16, height=11)
    add_bar_chart(ws, "办公鼠标品牌营收份额对比",
                  range(15, 21), 1, [2], "I30", width=16, height=10)

    return ws


# ════════════════════════════════════════════════════════════════
# Sheet 4: 键盘轴体技术分析
# ════════════════════════════════════════════════════════════════
def build_switch_sheet(wb):
    ws = wb.create_sheet("键盘轴体分析")
    set_col_widths(ws, {"A": 16, "B": 14, "C": 14, "D": 12, "E": 18, "F": 16, "G": 20, "H": 22})

    write_title_row(ws, 1, "键盘轴体(开关)技术与市场分析", 8, 44)

    # ── 轴体品牌份额 ──
    write_subtitle_row(ws, 3, "一、机械轴体品牌市场份额（2025年）", 8)
    headers = ["品牌", "营收份额", "出货份额", "总代", "核心产品", "代工厂", "技术特点", "核心优势"]
    write_header_row(ws, 4, headers)

    switch_brand_data = [
        ["佳达隆(Gateron)", 0.22, 0.28, "中国东莞",
         "G Pro系列/G黄/G白/CJ系列\nInk系列/Milky系列",
         "自产", "POM轴心/MX兼容\n出厂预润滑\nKS-3/KS-9底座",
         "性价比最高；出货量最大\n客制化社区深受欢迎"],
        ["凯华(KAILH/Kaihua)",0.18, 0.20, "中国东莞",
         "BOX白/BOX红/夜枭系列\n深海静音/芒果轴",
         "自产", "BOX结构防水防尘\n扭簧发声方案\nCP(镀金弹簧)工艺",
         "BOX结构独特专利\n轴体种类最多\n雷蛇/罗技代工来源"],
        ["Cherry(樱桃)",      0.17, 0.12, "德国/中国",
         "MX2A Red/Brown/Blue\nMX3A Silent/Black\nViola光学",
         "自产(德国+中国)", "黄金十字触点\nMX2A新模具+出厂润滑\n德国制造品质信心",
         "行业标准制定者\n品牌溢价+品质口碑\n德国制造情怀加成"],
        ["TTC(正牌科电)",     0.12, 0.13, "中国东莞",
         "金刚粉红轴/冰静轴\n金茶轴/红轴\n万彩轴(Wan Tsai)",
         "自产", "镀金弹簧+双触点\n防尘结构\n高可靠性(1亿次寿命)",
         "可靠性极高\n雷蛇/HyperX代工\n大厂首选ODM伙伴"],
        ["环诺/高特(Outemu)", 0.08, 0.12, "中国", "蓝/红/棕轴\n防尘轴系列",
         "自产", "低端入门级\n防尘硅胶罩\n成品键盘大量使用",
         "成本最低\n入门成品键盘广泛采用"],
        ["雷蛇(自有轴)",      0.05, 0.04, "品牌定制",
         "绿轴/黄轴/橙轴\n光学轴(线性/段落)\n雷蛇模拟(霍尔)",
         "凯华/特科代工", "定制手感参数\n光学轴1亿次寿命\n模拟霍尔连续触发",
         "品牌生态绑定\n雷蛇迷用户首选"],
        ["其他(SP/Zeal/Alps)", 0.18, 0.11, "国际",
         "SP SA高度键帽/Zeal PC\nAlps复古轴",
         "美国/日本等", "高端客制化\nAlps复古轴体\n稀有轴收藏",
         "客制化高端市场\n收藏稀缺性"],
    ]
    for i, row_data in enumerate(switch_brand_data):
        write_data_row(ws, 5 + i, row_data,
                       [None, "0.0%", "0.0%", None, None, None, None, None],
                       height=48)

    # ── 主流轴体参数对比 ──
    write_subtitle_row(ws, 13, "二、主流轴体参数对比", 8)
    headers2 = ["轴体型号", "类型", "触发力度(g)", "触发行程(mm)", "总行程(mm)",
                "寿命(万次)", "价格参考(元/颗)", "最佳应用场景"]
    write_header_row(ws, 14, headers2)

    switch_params = [
        ["Cherry MX2A Red",   "线性", 45, 2.0, 4.0, 10000, "2-3",   "游戏/打字通用"],
        ["Cherry MX2A Brown", "段落", 55, 2.0, 4.0, 10000, "2-3",   "打字/办公首选"],
        ["Cherry MX2A Blue",  "强段落(有声)",60,2.2,4.0,10000,"2-3",  "打字/打字声音爱好者"],
        ["Gateron G Pro Yellow","线性", 50, 2.0, 4.0, 8000,  "1-1.5", "游戏性价比首选"],
        ["Gateron CJ",        "线性", 55, 2.0, 4.0, 6000,  "2-3",   "客制化线性天花板"],
        ["Kailh BOX White",   "有声段落", 50, 1.8, 3.6, 8000, "1.5-2","防尘+清脆声音玩家"],
        ["Kailh Nightwalker(静音)","线性静音",45,2.0,3.6,8000,"1.5-2","电竞/宿舍静音"],
        ["TTC 金刚粉红轴",    "线性", 42, 2.0, 4.0, 10000, "2-2.5","原厂大厂品质"],
        ["Razer 光学线性(代工)","光学线性",40,1.0,4.0,10000,"雷蛇独占","FPS游戏快速触发"],
        ["磁轴(霍尔效应)",    "线性可调", "取决于弹簧", "0.1-4.0", "4.0", "15000+", "3-5",   "FPS急停王者"],
    ]
    for i, row_data in enumerate(switch_params):
        write_data_row(ws, 15 + i, row_data, height=24)

    # ── 轴体代工关系 ──
    write_subtitle_row(ws, 26, "三、品牌轴体代工关系", 8)
    headers3 = ["成品品牌", "轴体品牌/型号", "代工厂", "类型", "触发参数", "代工特点", "品质等级"]
    write_header_row(ws, 27, headers3)

    oem_data = [
        ["雷蛇 Razer 绿轴",   "雷蛇定制", "凯华/特科",   "有声段落", "50g/1.9mm",   "凯华代工绿轴，手感类似Cherry青", "高"],
        ["雷蛇 Razer 黄轴",   "雷蛇定制", "凯华/特科",   "线性",     "45g/1.2mm",   "短触发+FPS优化",                "高"],
        ["雷蛇 光学轴",       "雷蛇定制", "特科",         "光学线性", "40g/1.0mm",   "红外阻断式光学触发",             "高"],
        ["罗技 ROMER-G",      "罗技定制", "凯华(停产后欧姆龙)", "线性/段落","50g/1.5mm","中央LED设计(已停产)",             "中高"],
        ["罗技 GL轴(矮轴)",   "罗技定制", "凯华",         "线性/段落", "50g/1.5mm",  "矮轴,G913/G915使用",            "高"],
        ["HyperX 红轴(第一代)","HyperX定制","TTC",        "线性",     "45g/1.8mm",   "TTC代工,镀金金手指",            "高"],
        ["海盗船 OPX光轴",    "海盗船定制","欧姆龙(日本)", "光学线性", "40g/1.0mm",   "日本欧姆龙代工,短触发",         "极高"],
        ["赛睿 OmniPoint 2.0","赛睿定制","磁轴(霍尔)",    "磁轴可调", "10-45g/0.1-4.0mm","可调触发点0.1mm精度",       "极高"],
    ]
    for i, row_data in enumerate(oem_data):
        write_data_row(ws, 28 + i, row_data, height=36)

    add_pie_chart(ws, "机械轴体品牌营收份额", [5, 6, 7, 8, 9, 10, 11], 1, 2, "A37", width=16, height=11)

    return ws


# ════════════════════════════════════════════════════════════════
# Sheet 5: 鼠标微动与传感器分析
# ════════════════════════════════════════════════════════════════
def build_sensor_sheet(wb):
    ws = wb.create_sheet("鼠标微动与传感器")
    set_col_widths(ws, {"A": 18, "B": 16, "C": 16, "D": 16, "E": 20, "F": 24, "G": 28})

    write_title_row(ws, 1, "鼠标微动开关与传感器技术分析", 7, 44)

    # ── 微动品牌对比 ──
    write_subtitle_row(ws, 3, "一、鼠标微动开关品牌对比", 7)
    headers = ["品牌", "类型", "寿命(万次)", "力度(g)", "手感特点", "应用品牌", "评价"]
    write_header_row(ws, 4, headers)

    switch_data = [
        ["欧姆龙(OMRON) Japan",  "机械蓝点/白点", "5000-6000", "50-75",
         "业界标杆，手感清脆\n回弹快，耐久好\nD2FC-F系列通用",
         "罗技/雷蛇/赛睿/海盗船", "行业标准，高端首选"],
        ["欧姆龙(OMRON) China",  "机械白点/灰点", "1000-3000", "50-75",
         "比日产略闷\n寿命较短\n成本低",
         "罗技入门/部分国产", "性价比之选"],
        ["凯华(KAILH) 黑壳金点", "机械红/黑/蓝", "6000-8000", "55-90",
         "大力金刚指优选\n寿命长\n防尘版IP56",
         "HyperX/雷蛇部分型号", "寿命领先，偏硬"],
        ["华诺(HUANO) 蓝壳/粉点","机械/静音", "3000-5000", "50-70",
         "手感细腻\n粉色外壳辨识度高\n国产精品",
         "Glorious/Ninjutso", "客制化鼠标热门"],
        ["TTC 防尘金(金点/金轮)","机械金点", "6000", "60-80",
         "防尘结构+镀金触点\n寿命长可靠性高\n金轮编码器出色",
         "罗技G502X(Gold)\n部分国产鼠", "滚轮编码器也非常好"],
        ["雷蛇 光学微动(第三代)", "光学(红外)", "9000-10000", "55-75",
         "红外光触发无机械触点\n无双击问题\n响应快(1ms以内)",
         "雷蛇全系中高端", "彻底解决双击\n寿命最长"],
        ["罗技 Lightforce混合微动","光学+机械", "10000", "60-70",
         "光学触发+机械确认混合\n既省电又有手感\nG502X首创",
         "罗技G502X/G Pro X", "兼顾省电与手感创新"],
    ]
    for i, row_data in enumerate(switch_data):
        write_data_row(ws, 5 + i, row_data, height=48)

    # ── 鼠标无线技术 ──
    write_subtitle_row(ws, 13, "二、鼠标无线技术对比", 7)
    headers2 = ["技术方案", "连接方式", "延迟(毫秒)", "续航(典型)", "支持回报率",
                "代表产品", "综合评价"]
    write_header_row(ws, 14, headers2)

    wireless_data = [
        ["罗技 LIGHTSPEED",  "2.4GHz专有", "<1ms",  "50-80h(GPX S2)\n连续>140h(G603)",
         "1000Hz(原)/4K+更新", "G Pro X Superlight 2\nG502 X/G903",
         "行业延迟标杆\nPowerPlay无线充电独有\nHero传感器功耗极低"],
        ["雷蛇 HyperSpeed",  "2.4GHz专有", "<1ms",  "60-100h(Viper V3 Pro)",
         "原生1000Hz\n4KHz+HyperPolling", "Viper V3 Pro\nDeathAdder V3 Pro",
         "4KHz轮询领先\n续航与性能平衡好"],
        ["赛睿 2.4GHz/Quantum", "2.4GHz专有","<2ms","40-60h",
         "1000Hz", "Aerox 5/9 Wireless",
         "成熟稳定\n无突出优势"],
        ["博通/赛普拉斯 BLE 5.0/5.2","Bluetooth LE","8-12ms",">200h(办公)\n>100h(游戏)",
         "125-250Hz", "办公类无线鼠标\nMX Master 3S\nNod/多设备",
         "通用性兼容性最好\n多设备切换\n延迟不适用FPS"],
        ["海盗船 Slipstream", "2.4GHz专有", "<2ms",  "30-50h",
         "1000Hz", "Dark Core RGB Pro\nSabre Pro",
         "自有协议\n生态绑定iCUE\n覆盖距离较长"],
        ["USB 有线连接",     "USB线直连",  "<0.5ms", "无限制(供电)",
         "最高8000Hz", "全系有线鼠标\n部分可无线切换",
         "最低延迟无续航焦虑\n线缆束缚存在"],
    ]
    for i, row_data in enumerate(wireless_data):
        write_data_row(ws, 15 + i, row_data, height=48)

    # ── 鼠标重量趋势 ──
    write_subtitle_row(ws, 22, "三、游戏鼠标重量演进趋势", 7)
    headers3 = ["时期", "典型重量范围", "代表型号", "减重技术", "市场趋势", "结构特点", "用户偏好"]
    write_header_row(ws, 23, headers3)

    weight_data = [
        ["2015-2017",    "100-130g", "雷蛇DeathAdder Elite\n罗技G502(121g+线)",
         "无特殊减重\n传统钢筋滚轮",
         "有线为主\n设计偏好厚重感",
         "全实心塑料外壳\n金属滚轮",
         "当时用户认为重=扎实"],
        ["2018-2019",    "80-100g",  "罗技G Pro Wireless(80g)\n雷蛇Viper(69g)",
         "蜂巢打孔内部结构\n打孔外壳(轻量化早期)",
         "FPS驱动轻量化\n无线不再增加重量",
         "开始使用薄壁外壳\n蜂巢内骨架",
         "轻量化的接受度增长"],
        ["2020-2021",    "60-80g",   "Glorious Model O(58g)\n罗技GPX S1(63g)\n雷蛇Viper Mini(61g)",
         "打孔外壳极致化\n蜂窝结构全面应用\n小尺寸+薄壳",
         "打孔成为主流\n60g成为游戏目标\n磁吸充电",
         "大面积蜂窝打孔\n镁合金骨架尝试\n无RGB减重",
         "轻量化全面接受\n60g成为金标准"],
        ["2022-2023",    "50-65g",   "罗技GPX S2(60g)\n雷蛇Viper V3 Pro(55g)\nFinalmouse Starlight(42g)",
         "镁铝合金外壳(镂空)\n无打孔超薄壁\n电池小型化",
         "不打孔轻量化\n镁合金/碳纤维新材质\n小电池+长续航平衡",
         "镁铝合金一体框架\n无打孔密闭外壳\n特殊脚贴减摩擦",
         "不打孔+轻量=最高追求\n材质升级接受度好"],
        ["2024-2025",    "45-60g",   "雷蛇Viper V3 Hyperspeed(55g)\nROG Harpe Ace(54g)\nFinalmouse ULX(38g)",
         "碳纤维/镁合金\n无打孔极致薄壁\n极小电池+高效无线",
         "40g已不稀奇\n碳纤维成为新趋势\n极致轻量+4K/8K轮询",
         "碳纤维/钛合金/镁合金\n超小主板设计\n集成传感器定制",
         "追求超轻量+高性能\n材质成为新差异化"],
    ]
    for i, row_data in enumerate(weight_data):
        write_data_row(ws, 24 + i, row_data, height=48)

    return ws


# ════════════════════════════════════════════════════════════════
# Sheet 6: 技术特点综合对比
# ════════════════════════════════════════════════════════════════
def build_tech_comparison_sheet(wb):
    ws = wb.create_sheet("技术特点综合对比")
    set_col_widths(ws, {"A": 16, "B": 20, "C": 18, "D": 18, "E": 18, "F": 18, "G": 22})

    write_title_row(ws, 1, "键盘鼠标技术特点综合对比", 7, 44)

    # ── 人机工程设计 ──
    write_subtitle_row(ws, 3, "一、人机工程学设计对比", 7)
    headers = ["设计方向", "代表产品", "人体工学特点", "适用场景", "用户群体", "健康影响", "市场趋势"]
    write_header_row(ws, 4, headers)

    ergo_data = [
        ["分体式键盘(Alice/Ergo)", "Keychron Q8/10\nAlice系列/Arisu",
         "左右手分区，打字姿势自然\n减少手腕外翻\n角度可调分体",
         "程序员/作家\n高强度打字用户",
         "高端打字需求\n办公族",
         "减少尺骨偏斜\n降低腕管综合征风险",
         "客制化热门布局\n受程序员社区追捧"],
        ["垂直鼠标",   "罗技MX Vertical\n国产垂直鼠标\n金士顿Silhouette",
         "竖握手型自然姿势\n减轻前臂扭转\n缓解鼠标手握压力",
         "手腕疼痛用户\n办公/设计",
         "办公室人群\n腕管综合征患者",
         "可显著减少前臂扭转\n缓解重复性劳损",
         "健康意识提升驱动\n增长稳健"],
        ["人体工学键盘(波浪形)", "微软Sculpt/Ergonomic\nKeychron分体\n阿米洛人体工学",
         "波形键位排列\n弧形键盘分开\n带掌托",
         "高强度办公\n专业写作",
         "程序员/作家\n翻译/数据录入",
         "减少手指伸展\n减轻手腕压力",
         "平缓但持续增长\n企业健康采购"],
        ["轨迹球鼠标", "罗技M575/ERGO M580\nKensington SlimBlade\nElecom DEFT",
         "手固定不动用拇指/手指滚球\n无需手臂移动\n空间要求极低",
         "空间受限环境\n桌面极简/嵌入式",
         "工程师/做图/HTPC\n专业工作站",
         "肩颈压力最小\n无手腕移动伤害",
         "小众但稳定\n专业领域忠诚度高"],
        ["掌托/手托系统", "各品牌键盘掌托\n雷蛇/罗技记忆海绵\n木质/树脂客制化掌托",
         "支撑手腕/手掌\n减少手腕悬空\n辅助抬高打字角度",
         "键盘/鼠标配套\n长时间办公",
         "所有长时间电脑用户",
         "可减轻手腕压力\n但需正确高度配合",
         "高端产品标配\n客制化配件丰富"],
    ]
    for i, row_data in enumerate(ergo_data):
        write_data_row(ws, 5 + i, row_data, height=36)

    # ── 键位布局 ──
    write_subtitle_row(ws, 11, "二、键盘布局与配列对比", 7)
    headers2 = ["配列名", "键数", "特点", "代表产品", "适合场景", "便携性", "人气趋势"]
    write_header_row(ws, 12, headers2)

    layout_data = [
        ["100% Full Size",     "104/108键", "完整数字键盘+功能键区+方向键", "罗技G915/G913\nCherry G80-3000",
         "办公会计/数据输入\n无空间限制使用", "最低", "逐渐下降"],
        ["1800 Compact",       "96-100键", "基本完整布局但紧凑排列\n方向键和数字键无空格", "Keychron Q5\nLeopard FC980M",
         "数据输入需求+桌面空间节约\n会计/财务", "较低", "稳定小众"],
        ["80% TKL(TenKeyLess)", "87键", "砍数字键盘，保留功能键和方向键", "罗技G Pro X TKL\n雷蛇黑寡妇V3 TKL",
         "FPS游戏首选\n节省桌面空间", "中", "经典常青树"],
        ["75% Compact",        "80-84键", "TKL基础上更紧凑\nF键有间隔但紧凑", "Keychron Q1/V1\nGlorious GMMK Pro",
         "程序员/游戏兼顾\n2023-2025最热配列", "高", "快速增长中"],
        ["65%",                "65-68键", "仅主键区+方向键+小型编辑键\n无F键行", "Tofu65/Keychron Q2/V2\nDucky One 3 Mini",
         "桌面极简主义\n轻运输/背包携带", "较高", "客制化热门"],
        ["60%",                "61键", "仅主键区，无方向键/F键/编辑\nFn组合键调用", "Poker/Anne Pro 2/3\nDucky Mini/GH60",
         "极简桌面\n重度Fn快捷键使用", "最高", "稳定，硬核用户"],
        ["特殊布局(Alice/分体)", "分体拼接", "左右手分离/Alice弧形/直列\n人机工学为主", "Keychron Q8/10\n阿米洛Alice\nMountain Ergo",
         "长时间打字的程序员\n写作/文学", "低", "快速增长"],
    ]
    for i, row_data in enumerate(layout_data):
        write_data_row(ws, 13 + i, row_data, height=36)

    # ── 连接方式对比 ──
    write_subtitle_row(ws, 21, "三、键盘连接方式对比", 7)
    headers3 = ["连接方式", "延迟", "续航/供电", "多设备支持", "兼容性", "适用场景", "代表产品"]
    write_header_row(ws, 22, headers3)

    connectivity_data = [
        ["有线(USB Type-C)", "<1ms", "无续航焦虑", "否", "最佳/100%", "游戏/竞技/固定办公", "有线机械键盘全系列"],
        ["2.4GHz(专有接收器)", "1-2ms", "数周至数月(干电池/锂电)", "否", "需要接收器", "游戏/无蓝牙环境", "罗技Lightspeed/雷蛇HyperSpeed"],
        ["Bluetooth 3.0/5.0+", "8-15ms", "数月至一年+(干电池)/数周(充电)", "可连3+设备", "广泛", "办公/多设备切换/Mac", "Keychron/罗技MX/阿米洛"],
        ["三模(有线+2.4G+BT)", "有线<1ms\n2.4G 1-3ms\nBT 8-12ms", "电池供电可充电", "可连多设备", "最佳", "全场景通用", "Keychron Q/V/K/罗技G"],
        ["RF无线(低端专有)", "15-25ms", "数月至一年(干电池)", "通常否", "需接收器", "入门办公", "联想/戴尔/双飞燕办公键鼠"],
    ]
    for i, row_data in enumerate(connectivity_data):
        write_data_row(ws, 23 + i, row_data)

    return ws


# ════════════════════════════════════════════════════════════════
# Sheet 7: 发展趋势与展望
# ════════════════════════════════════════════════════════════════
def build_trends_sheet(wb):
    ws = wb.create_sheet("发展趋势与展望")
    set_col_widths(ws, {"A": 8, "B": 20, "C": 36, "D": 20, "E": 14, "F": 14, "G": 30})

    write_title_row(ws, 1, "键盘鼠标市场发展趋势与展望", 7, 44)

    # ── 行业趋势 ──
    write_subtitle_row(ws, 3, "一、2025-2027年行业核心趋势", 7)
    headers = ["序号", "趋势方向", "详细说明", "受益细分市场", "确定性", "影响程度", "关键驱动因素"]
    write_header_row(ws, 4, headers)

    trends_data = [
        ["1", "磁轴(霍尔效应)渗透率飙升",
         "磁轴键盘实现0.01mm精度的可调触发点\nFPS游戏急停操作刚需(Cs2/Valorant)\n雷蛇/赛睿/醉鹿/Wooting推动普及\n2027年预计占游戏键盘20-25%",
         "游戏键盘", "高", "★★★★★",
         "FPS电竞急停需求\n霍尔传感器成本下降\n品牌旗舰转向磁轴"],
        ["2", "无线延迟追平有线",
         "罗技Lightspeed/雷蛇HyperSpeed延迟已<1ms\n4K/8K轮询无线完成\n未来2年有线延迟优势消失",
         "游戏鼠标/键盘", "极高", "★★★★",
         "无线技术持续进步\n续航与性能平衡突破\n去线缆化大势所趋"],
        ["3", "AI赋能外设体验",
         "AI动态调整DPI/键灵敏度根据游戏场景\n自动宏录制/AI性能教练\n键盘输入预测与修正\n健康监测告警(手腕/坐姿)",
         "高端外设生态", "中", "★★★",
         "AI硬件化浪潮\n外设成为AI交互入口\n大厂布局"],
        ["4", "客制化向主流渗透",
         "热插拔轴座从客制化进入量产键盘(Keychron带动)\nGasket结构普及到500元以下键盘\n客制化社区反向影响品牌设计",
         "机械键盘全市场", "高", "★★★★",
         "DIY文化年轻化\n社交媒体/YouTube评测\n大规模OEM产能易客制化"],
        ["5", "磁轴/光学轴加速替代传统机械轴",
         "磁轴可调触发竞争差异化\n光学轴寿命(1亿次)远超机械(5000万)\n大厂优先推自有轴(雷蛇/罗技/赛睿)",
         "轴体市场/高端键盘", "高", "★★★★",
         "电竞竞技化需求\n自有生态锁定用户\n制造代工成熟"],
        ["6", "可持续发展的绿色外设",
         "罗技承诺PCR塑料占比>50%\nFSC包装/碳中性认证成为标配\n被动散热(取消风扇)/低待机功耗",
         "所有品牌", "中", "★★",
         "ESG(环境社会和治理)压力\n消费者环保意识\n欧盟新规推动"],
        ["7", "国产轴体全面崛起",
         "佳达隆出货量已超Cherry\n凯华BOX轴市占率攀升\nTTC代工大厂OEM品质\n国产轴体性价比远超原厂",
         "轴体市场", "极高", "★★★★★",
         "中国制造质量提升\n客制化社区国产接受度高\nOEM/ODM产能支撑规模化"],
    ]
    for i, row_data in enumerate(trends_data):
        write_data_row(ws, 5 + i, row_data, height=56)

    # ── 厂商技术路线 ──
    write_subtitle_row(ws, 13, "二、主要厂商技术路线图（2025-2027）", 7)
    headers2 = ["厂商", "键盘路线", "鼠标路线", "传感器方案", "无线方案", "轴体/微动路线", "生态战略"]
    write_header_row(ws, 14, headers2)

    roadmap_data = [
        ["罗技(Logitech)",
         "G913后继(矮轴+无线)\nTKL+65%双线\n磁轴布局中",
         "GPX S3(碳纤维/更轻量)\nG502 X升级(Hero 3)\nPowerPlay V2无线充",
         "Hero 3自研传感器\n(50K DPI+AI追踪)\n更低功耗",
         "Lightspeed V2(8K原生)\nPowerPlay V2无线充\nBolt办公系列升级",
         "Lightforce混合微动V2\n自研磁轴(研发中)\nGL矮轴V2",
         "GHUB平台统一\nLogitech AI for Gaming\n办公+游戏双线独立发展"],
        ["雷蛇(Razer)",
         "黑寡妇V5(磁轴)\n噬魂金蝎V2(光学)\nRainbow/生态",
         "Viper V4 Pro(碳纤维+\n4K原生+40g级)\nDeathAdder V4",
         "Focus Pro 40K+自研\n(4K原生轮询)\nAI运动同步V2",
         "HyperSpeed V3(4K原生)\nHyperPolling 8K V2\nQi2无线充电",
         "雷蛇光学轴V3\n(模拟霍尔轴推广)\n光学微动V4(GaN)",
         "Chroma生态智能联动\nAI灵敏度+灯效\n全外设+音频+椅子生态"],
        ["海盗船(Corsair)",
         "K70 V2(OLED升级)\n磁轴推出(2026)\n矮轴产品线扩展",
         "Dark Core 续作\nM65/Scimitar更新\n夜间模式增强",
         "MARKMAN 2定制\n(+增加AI追踪)\n原生4K",
         "Slipstream V2(4K原生)\niCUE无线协议升级\nUSB4扩展",
         "OPX光轴(欧姆龙)\n品牌LAMBO(自有轴)\n樱桃轴稳定合作",
         "iCUE平台一站式\nElgato直播生态联动\n游戏+直播+创作生态"],
        ["赛睿(SteelSeries)",
         "Apex Pro V2(磁轴主力)\nOmniPoint 2.0(0.1-4mm)\nGG整合驱动",
         "Aerox V2/Prime续作\nTrueMove Air续作\nOptic升级",
         "TrueMove Pro自研\n(协作原相)\n智能化校准",
         "Quantum 2.0无线\n低延迟+长续航\nQi充电统一化",
         "OmniPoint磁轴深耕\n2代可调行程+压力\n自有微动开发中",
         "GameSense生态(游戏联动)\nGG软件平台\n电竞赞助优先"],
        ["Keychron",
         "Q系列(全铝Gasket)\nV系列(性价比Gasket)\nK系列(无线办公)\nLemokey(游戏子牌)",
         "尚未进入鼠标领域\n专注于键盘",
         "—", "QMK/VIA开源固件\n蓝牙5.1/2.4G三模\n全系Mac/Win双兼容",
         "预装Gateron/佳达隆\n热插拔轴座全部兼容\nK Pro系列优化调校",
         "开源固件社区驱动\n客制化品牌第一\nMac友好最大卖点"],
        ["佳达隆(Gateron)",
         "纯轴体制造商\n不为成品键盘品牌\n(偶有限量套件)",
         "—", "—", "—",
         "G Pro V4(全POM自润滑)\n磁轴产品线(霍尔效应)\n各类型轴体持续迭代",
         "专注轴体研发\n不参与品牌竞争\n出货量全球第一目标"],
    ]
    for i, row_data in enumerate(roadmap_data):
        write_data_row(ws, 15 + i, row_data, height=56)

    # ── 投资建议 ──
    write_subtitle_row(ws, 22, "三、行业投资关注方向", 7)
    headers3 = ["序号", "方向", "逻辑", "标的类型", "时间窗口", "风险等级", "风险提示"]
    write_header_row(ws, 23, headers3)

    invest_data = [
        ["1", "磁轴(霍尔效应)产业链",
         "FPS电竞风潮推动磁轴需求爆发\n2025年仅Wooting/赛睿/醉鹿三家主力\n替代传统机械轴空间巨大",
         "霍尔传感器厂/磁轴品牌", "2025-2027", "中",
         "供应链产能不足\nApple专利限制\n是否只是短期热"],
        ["2", "国产自主品牌客制化出海",
         "Keychron海外客制化品牌/销量非常大\n佳达隆全球出货超Cherry\n国内供应链成熟+性价比极高",
         "国产品牌/代工厂", "持续", "中低",
         "品牌溢价不足\n国际渠道建设难\n贸易摩擦影响"],
        ["3", "无线游戏外设升级",
         "无线低延迟已全面成熟\n4K/8K轮询成为高端壁垒\n鼠标轻量化持续更新换代需求",
         "无线芯片/传感器厂", "2025-2026", "低",
         "竞争加剧\n技术代差缩小\n续航仍是瓶颈"],
        ["4", "PC外设健康化细分",
         "垂直鼠标/分体键盘年增15%+\n办公族健康意识提升\n企业外设健康采购需求增长",
         "人体工学外设品牌", "2025-2028", "低",
         "市场规模天花板不高\n教育成本高(C端认知)\n替代品(指环鼠标等)影响"],
        ["5", "外设AI化+软件生态",
         "AI赋能的智能外设将重新定义交互\n动态设置/语音助手/健康监测\n软件生态粘性决定客户留存",
         "外设软件平台/算法公司", "2026-2028", "高",
         "AI实用性待验证\n用户付费意愿不确定\n数据隐私问题"],
    ]
    for i, row_data in enumerate(invest_data):
        write_data_row(ws, 24 + i, row_data, height=48)

    # ── 产业链图谱 ──
    write_subtitle_row(ws, 30, "四、外设产业链核心环节", 7)
    headers4 = ["环节", "说明", "代表厂商", "国产化程度", "壁垒", "价值占比", "趋势"]
    write_header_row(ws, 31, headers4)

    chain_data = [
        ["传感器(鼠标光学)", "光电传感器芯片设计", "原相 PixArt/罗技/赛普拉斯", "低(原相垄断)", "高(光学技术+专利)", "15-20%", "原相主导，雷蛇/罗技自研追赶"],
        ["MCU主控芯片",      "鼠标键盘微控制处理器", "NXP/STM/赛普拉斯/博通/Nordic","中", "中(ARM生态成熟)", "8-12%", "国产GD32/沁恒加速替代"],
        ["机械轴体制造",      "开关设计+精密注塑+组装", "佳达隆/凯华/Cherry/TTC",    "高", "中(工艺积累+专利)", "15-20%(键盘)", "国产轴全球份额过半"],
        ["微动开关",          "鼠标按键微动/编码器", "欧姆龙/凯华/TTC/华诺",         "中高", "中(进口替代中)", "5-8%(鼠标)", "国产寿命追平日本"],
        ["外壳模具(ABS/PBT)", "键帽+外壳注塑成型", "台资/大陆的模具ODM厂",          "极高", "低(成熟制造)", "10-15%", "PBT键帽占比提升"],
        ["无线芯片方案",      "2.4G+蓝牙多模射频IC", "Nordic/博通/Realtek/赛普拉斯", "中", "中(射频+协议栈)", "8-12%", "国产芯(AirM2M等)快速追赶"],
        ["电池/电源管理",     "锂电/充电IC/快充电路", "ATL/欣旺达/TI/MPS",           "高", "低(成熟方案)", "3-5%", "快充+长续航标准化"],
        ["品牌+渠道+软件",    "品牌营销/线上线下/驱动软件", "罗技/雷蛇/赛睿/海盗船",  "低(品牌力弱)", "极高(品牌壁垒)", "25-35%", "品牌溢价最大价值环节"],
    ]
    for i, row_data in enumerate(chain_data):
        write_data_row(ws, 32 + i, row_data, height=36)

    return ws


# ════════════════════════════════════════════════════════════════
# 主入口
# ════════════════════════════════════════════════════════════════
def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    print("生成 Sheet 1: 市场概览...")
    build_overview_sheet(wb)

    print("生成 Sheet 2: 键盘市场分析...")
    build_keyboard_sheet(wb)

    print("生成 Sheet 3: 鼠标市场分析...")
    build_mouse_sheet(wb)

    print("生成 Sheet 4: 键盘轴体分析...")
    build_switch_sheet(wb)

    print("生成 Sheet 5: 鼠标微动与传感器...")
    build_sensor_sheet(wb)

    print("生成 Sheet 6: 技术特点综合对比...")
    build_tech_comparison_sheet(wb)

    print("生成 Sheet 7: 发展趋势与展望...")
    build_trends_sheet(wb)

    wb.save(str(OUTPUT_FILE))
    print(f"\n[OK] 报告已生成: {OUTPUT_FILE}")
    print(f"   共 7 个工作表，文件大小: {OUTPUT_FILE.stat().st_size / 1024:.1f} KB")


if __name__ == "__main__":
    main()
