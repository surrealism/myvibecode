"""
音响市场分析报告生成器
覆盖 TWS耳机、智能音箱、蓝牙音箱、家庭影院、专业音响、汽车音响、Hi-Fi、头戴耳机 等细分市场
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
OUTPUT_FILE = OUTPUT_DIR / "音响市场分析报告.xlsx"

# ── 配色方案 ──
C = {
    "title_bg":    "3E2A1F",   # 深棕(音响主题)
    "title_fg":    "FFFFFF",
    "header_bg":   "8B5E3C",   # 暖棕
    "header_fg":   "FFFFFF",
    "sub_bg":      "F0E0D0",   # 浅米
    "accent1":     "4472C4",
    "accent2":     "ED7D31",
    "accent3":     "70AD47",
    "light_gray":  "F5F0EC",
    "white":       "FFFFFF",
    "deep_brown":  "3E2A1F",
    "red":         "C00000",
    "green":       "008000",
}

# ── 通用样式工厂 ──
thin_border = Border(
    left=Side(style="thin", color="D4B896"),
    right=Side(style="thin", color="D4B896"),
    top=Side(style="thin", color="D4B896"),
    bottom=Side(style="thin", color="D4B896"),
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
    cell.font = _font(size=12, bold=True, color=C["deep_brown"])
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

    write_title_row(ws, 1, "全球音响市场分析报告（2024-2025）", 7)
    write_source_row(ws, 2, "数据来源：Futuresource、Strategy Analytics、Counterpoint、IDC、Omdia 等公开报告 | 制表日期：2026年7月", 7)

    # ── 细分市场规模 ──
    write_subtitle_row(ws, 4, "一、音响细分市场规模", 7)
    write_header_row(ws, 5, [
        "细分市场", "2023年(亿美元)", "2024年(亿美元)", "同比增速",
        "2025年预测(亿美元)", "主要应用", "市场特征"
    ])
    data = [
        ["TWS真无线耳机",  380, 420, 0.105, 450, "消费级耳机",   "苹果/索尼/三星主导,白牌承压"],
        ["智能音箱",       110, 95, -0.136,  90, "智能家居",     "亚马逊/谷歌/百度三足鼎立,增长放缓"],
        ["蓝牙音箱",       220, 240, 0.091, 260, "便携/户外",    "JBL/索尼/BOSE主导,细分化"],
        ["家庭影院/回音壁", 140, 155, 0.107, 170, "客厅影音",     "索尼/三星/LG主导,全景声成卖点"],
        ["专业音响",        85,  92,  0.082, 100, "演出/录音/商用","Sennheiser/Shure/STM主导,高壁垒"],
        ["汽车音响",       320, 350, 0.094, 380, "前装/后装车机", "哈曼/Bose/B&O主导OEM,新能源拉动"],
        ["Hi-Fi高保真",     45,  48,  0.067,  52, "发烧/家用",     "细分忠诚市场,流媒体Hi-Res拉动"],
        ["头戴耳机(非TWS)", 180, 185, 0.028, 190, "通勤/游戏/监听","索尼/BOSE/苹果,降噪是核心"],
    ]
    fmts = [None, "#,##0", "#,##0", "0.0%", "#,##0", None, None]
    for i, d in enumerate(data):
        write_data_row(ws, 6 + i, d, fmts)

    write_data_row(ws, 14, ["合计", 1380, 1585, 0.149, 1592, "—", "TWS+汽车音响是双引擎"],
                   fmts, bold_cols={1, 2, 3, 4, 5})

    # ── 市场集中度 ──
    write_subtitle_row(ws, 16, "二、各细分市场集中度（CR3）", 7)
    write_header_row(ws, 17, [
        "细分市场", "CR1", "CR2", "CR3", "头部厂商", "竞争格局", "进入壁垒"
    ])
    cr = [
        ["TWS真无线耳机",   "30%", "55%", "70%", "苹果/三星/索尼",      "苹果领跑,安卓阵营激烈", "中(芯片+算法+声学)"],
        ["智能音箱",        "28%", "55%", "75%", "亚马逊/谷歌/百度",    "寡头+生态绑定",          "高(语音AI生态)"],
        ["蓝牙音箱",        "25%", "45%", "60%", "JBL/索尼/BOSE",       "JBL领先,品牌集中",       "中(品牌+渠道)"],
        ["家庭影院/回音壁", "20%", "40%", "60%", "索尼/三星/LG",         "日韩主导,全景声竞赛",    "中高(声学+视频协同)"],
        ["专业音响",        "20%", "38%", "55%", "Sennheiser/Shure/STM","细分领域多,壁垒高",      "极高(声学+行业认证)"],
        ["汽车音响",        "35%", "60%", "78%", "哈曼/Bose/B&O",        "OEM绑定深,集中度高",     "高(车规+OEM关系)"],
        ["Hi-Fi高保真",     "12%", "22%", "32%", "天龙/马兰士/雅马哈",   "极度细分,品牌林立",      "中(调音文化+口碑)"],
        ["头戴耳机",        "35%", "55%", "72%", "苹果/BOSE/索尼",      "降噪三巨头主导",         "中(降噪算法+声学)"],
    ]
    for i, d in enumerate(cr):
        write_data_row(ws, 18 + i, d)

    # ── 关键趋势 ──
    write_subtitle_row(ws, 26, "三、2025年关键趋势", 7)
    write_header_row(ws, 27, ["序号", "趋势", "说明", "主要推进方", "确定性", "影响", "关键变量"])
    trends = [
        ["1", "空间音频普及",   "杜比全景声/索尼360RA下沉到中端\n头追+动态跟踪成标配", "苹果/索尼/杜比",   "高",  "★★★★", "内容生态/算法精度"],
        ["2", "AI降噪与通话",    "AI麦克风阵列+实时降噪\n通话/会议耳机升级",          "苹果/Bose/QT",     "高",  "★★★★", "端侧算力/AI模型"],
        ["3", "自研声学芯片",    "苹果H2/H3、索尼V2、Bose自研\nANC/算力/续航一体化",    "苹果/索尼/Bose",   "中高","★★★★", "芯片研发投入/规模"],
        ["4", "无损+LE Audio",   "蓝牙LE Audio+LC3编码\n高码率无损逐步落地",         "蓝牙SIG/高通",     "高",  "★★★",  "设备互通/编码普及"],
        ["5", "新能源车音响升级","多声道/主动声学/空间音频\n新能源车标配高端音响",      "哈曼/Bose/比亚迪", "高",  "★★★★", "新能源车销量/座舱升级"],
        ["6", "TWS高端反弹",     "开放式/带屏AirPods/健康监测\nTWS向可穿戴延伸",       "苹果/三星/华为",   "中高","★★★",  "健康传感器/形态创新"],
    ]
    for i, d in enumerate(trends):
        write_data_row(ws, 28 + i, d, height=44)

    add_bar(ws, "音响细分市场规模对比（亿美元）",
            range(6, 14), 1, [2, 4], "A36", w=22, h=12)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 2: TWS 真无线耳机市场
# ═══════════════════════════════════════════════════════════════
def build_tws_sheet(wb):
    ws = wb.create_sheet("TWS耳机市场分析")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 14, "D": 18, "E": 24, "F": 34, "G": 34})

    write_title_row(ws, 1, "TWS 真无线耳机市场厂商份额与技术分析", 7)

    # ── TWS份额 ──
    write_subtitle_row(ws, 3, "一、TWS 真无线耳机厂商份额（2024年）", 7)
    write_header_row(ws, 4, [
        "厂商", "出货量份额", "营收份额", "主力芯片/平台", "主力产品", "技术特点", "核心优势"
    ])
    tws = [
        ["苹果Apple",   0.30, 0.45, "Apple H2/H3",
         "AirPods Pro 2\nAirPods 4\nAirPods Max",
         "自研H2/H3芯片；主动降噪V2\n计算音频+自适应均衡\n空间音频+头追\nUSB-C/Find My",
         "品牌+生态绝对领先\n营收份额远超出货\niOS无缝体验"],
        ["三星Samsung", 0.18, 0.14, "三星/博通",
         "Buds3/Buds3 Pro\nBuds FE/Buds2 Pro",
         "自研+博通芯片；24bit音频\n360音频+杜比全景声\nAI实时降噪/通话\nGalaxy生态联动",
         "安卓阵营龙头\nGalaxy设备深度协同\nOLED屏+带柄新形态"],
        ["索尼Sony",    0.08, 0.10, "索尼V2/MTK",
         "WF-1000XM5\nLinkBuds系列",
         "自研V2芯片；双降噪麦克风\nLDAC无损+DSEE\n8.4g轻量化\n泡沫耳塞被动降噪",
         "降噪口碑标杆\nLDAC高解析音质\n发烧/通勤双覆盖"],
        ["BOSE",        0.06, 0.08, "Bose自研/通用",
         "QuietComfort Ultra\nQC Earbuds/Open",
         "CustomTune声学校准\n自适应ANC 10级\n开放式新形态\nAware入耳侦听",
         "降噪鼻祖品牌\n舒适度口碑\n商务/通勤忠诚用户"],
        ["华为Huawei",   0.06, 0.04, "海思/恒玄",
         "FreeBuds Pro 4\nFreeBuds 6i/Studio",
         "海思Kirin A2/Lossless\n星闪连接+L2HC 4.0\n双发音单元\n鸿蒙生态",
         "国内份额领先\n无损+低延迟(星闪)\n鸿蒙多设备协同"],
        ["小米Xiaomi",  0.05, 0.02, "恒玄/BES",
         "Buds 5 Pro\nRedmi Buds 系列",
         "恒玄/BES芯片；LE Audio\nLHDC 5.0高码率\n主动降噪55dB\n性价比旗舰",
         "性价比+销量双优\n印度/欧洲增长\n生态+价格双线"],
        ["其他(漫步者等)",0.27, 0.17, "恒玄/BES/络达",
         "漫步者/声阔/QCY/JBL TWS",
         "恒玄/络达/恒玄方案\nANC+LE Audio普及\n白牌方案成熟",
         "中低端量大\n方案公司支撑\n全球白牌承压"],
    ]
    fmts = [None, "0.0%", "0.0%", None, None, None, None]
    for i, d in enumerate(tws):
        write_data_row(ws, 5 + i, d, fmts, height=64)

    # ── 旗舰TWS对比 ──
    write_subtitle_row(ws, 13, "二、2025年旗舰 TWS 对比", 7)
    write_header_row(ws, 14, [
        "规格", "AirPods Pro 3", "Buds3 Pro", "WF-1000XM6", "QC Ultra 2", "FreeBuds Pro 4", "说明"
    ])
    flagship = [
        ["降噪芯片",   "Apple H3",   "三星Exynos",  "索尼V3",      "Bose自研",     "海思A3",    "自研芯片成趋势"],
        ["降噪深度",   "~45dB",      "~45dB",       "~50dB",       "~50dB",       "~50dB",     "旗舰均45dB+"],
        ["音频编码",   "AAC/LDAC?",  "24bit SSC",   "LDAC",        "aptX Lossless","L2HC 4.0",  "无损编码百家争鸣"],
        ["续航(单次)", "6h(开ANC)",  "6h(开ANC)",   "8h(开ANC)",   "5.5h(开ANC)", "6.5h(开ANC)","续航向6h+靠拢"],
        ["空间音频",   "杜比+头追",  "360+杜比",    "360RA",       "杜比+头追",    "空间音频",   "中端逐步下放"],
        ["连接方式",   "蓝牙5.3",    "蓝牙5.4",     "蓝牙5.3",     "蓝牙5.3",      "星闪+BT5.4","华为星闪差异化"],
        ["价格区间$",  "249",        "249",         "279",         "299",         "199(国内)", "安卓旗舰压价"],
    ]
    for i, d in enumerate(flagship):
        write_data_row(ws, 15 + i, d, height=32)

    # ── TWS芯片平台 ──
    write_subtitle_row(ws, 23, "三、TWS 蓝牙音频芯片平台对比", 7)
    write_header_row(ws, 24, [
        "芯片厂商", "代表方案", "主要客户", "音频编码", "ANC能力", "定位", "趋势"
    ])
    chip = [
        ["恒玄科技BES",  "BES2700/2800",   "小米/华为/漫步者", "LHDC 5.0", "自适应ANC",   "中高端安卓",  "国内TWS芯片龙头"],
        ["络达Lofarsil", "AB1565/1568",    "索尼/Anker/小米",   "AAC/SBC",  "Hybrid ANC",  "中端主流",    "联发科子公司,量大"],
        ["高通Qualcomm", "QCC5181/S5",     "Bose/JBL/索尼",     "aptX Lossless","自适应ANC","高端+车规", "aptX Lossless生态"],
        ["苹果Apple",   "H1/H2/H3",        "苹果自用",          "AAC",      "ANC V2/V3",   "iOS自研",    "自研自用不外供"],
        ["博通Broadcom", "BCM63xx",        "三星/部分高端",     "AAC/LDAC", "混合ANC",     "三星自研协同","定制化路线"],
        ["海思HiSilicon","Kirin A1/A2/A3", "华为自用",          "L2HC/LDAC","自适应ANC",  "华为自研",    "星闪+无损路线"],
    ]
    for i, d in enumerate(chip):
        write_data_row(ws, 25 + i, d, height=32)

    add_pie(ws, "TWS出货量份额", [5, 6, 7, 8, 9, 10, 11], 1, 2, "A32", w=16, h=11)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 3: 头戴与降噪耳机
# ═══════════════════════════════════════════════════════════════
def build_headphone_sheet(wb):
    ws = wb.create_sheet("头戴耳机市场分析")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 14, "D": 18, "E": 24, "F": 34, "G": 34})

    write_title_row(ws, 1, "头戴式耳机市场厂商份额与技术分析", 7)

    # ── 头戴份额 ──
    write_subtitle_row(ws, 3, "一、头戴耳机（含降噪）厂商份额（2024年）", 7)
    write_header_row(ws, 4, [
        "厂商", "出货量份额", "营收份额", "主力平台", "主力产品", "技术特点", "核心优势"
    ])
    hp = [
        ["苹果Apple",  0.20, 0.35, "H2/H3",
         "AirPods Max\n(USB-C更新版)",
         "Apple H2芯片\n计算音频+自适应均衡\n空间音频+动态头追\n不锈钢/铝合金机身",
         "高端价位标杆\niOS生态绑定\n奢侈品定位"],
        ["索尼Sony",   0.25, 0.22, "V2/V3+MTK",
         "WH-1000XM6\nWH-1000XM5",
         "自研V3芯片；8麦克风降噪\n动态降噪+智能场景\nLDAC+DSEE Extreme\n折叠便携设计",
         "降噪头戴标杆\n音质+降噪双优\n续航30h+口碑"],
        ["BOSE",       0.15, 0.18, "Bose自研",
         "QuietComfort Ultra\nNC700/降噪头盔",
         "Bose自适应ANC\nAware入耳侦听\n20级降噪调节\n航空降噪专利",
         "降噪鼻祖品牌\n商务/航空忠诚\n舒适度行业标杆"],
        ["森海Sennheiser",0.08, 0.08,"高通",
         "Momentum 4\nHD 660S2",
         "高通芯片；60h续航\n自适应降噪\n声音调校发烧级\n德式Hi-Fi传承",
         "Hi-Fi血统\n续航最长\n发烧用户忠诚"],
        ["Bowers&Wilkins",0.04, 0.06,"—",
         "Px8/Px7 S2\n700系列耳机",
         "高品质DAC+功放\n24bit高解析\n铝合金+真皮材质",
         "奢侈品定位\n调音/做工标杆\n高端利基"],
        ["Beats(Apple)",0.06, 0.05, "H2/W1",
         "Studio Pro\nSolo 4/Fit Pro",
         "Apple H2芯片\nUSB-C/Lossless\n运动+潮流定位",
         "潮流品牌+苹果生态\n年轻用户忠诚\n运动场景强"],
        ["其他(漫步者/铁三角等)",0.22, 0.06,"恒玄/BES",
         "漫步者W820NB+/铁三角M50x",
         "国产方案+性价比ANC\n监听/Hi-Fi细分",
         "中低端+专业细分\n性价比路线"],
    ]
    fmts = [None, "0.0%", "0.0%", None, None, None, None]
    for i, d in enumerate(hp):
        write_data_row(ws, 5 + i, d, fmts, height=64)

    # ── 降噪技术演进 ──
    write_subtitle_row(ws, 13, "二、主动降噪（ANC）技术演进", 7)
    write_header_row(ws, 14, [
        "代际", "时间", "技术特征", "降噪深度", "代表产品", "麦克风数量", "局限"
    ])
    anc = [
        ["第一代", "2016-2018", "前馈式单麦ANC\n固定滤波",        "20-25dB", "索尼WI-1000X\nBose QC35 II", "1-2",   "低频残留/啸叫"],
        ["第二代", "2019-2021", "混合反馈式ANC\n自适应滤波",      "30-35dB", "AirPods Pro\nSony XM4",       "2-4",   "风噪/通话差"],
        ["第三代", "2022-2024", "计算音频+AI\n多麦阵列+学习",      "40-45dB", "AirPods Pro 2\nXM5",         "4-6",   "环境适应仍有限"],
        ["第四代", "2024-2025", "AI实时降噪\n个性化声学校准",      "45-55dB", "Buds3 Pro/FreeBuds Pro 4",   "5-8",   "功耗/算力挑战"],
        ["趋势",   "2026+",     "AI神经网络降噪\n空间感知+健康",   "55dB+",  "AirPods Pro 3/XM6",          "6-10",  "隐私/续航"],
    ]
    for i, d in enumerate(anc):
        write_data_row(ws, 15 + i, d, height=40)

    # ── 音频编码 ──
    write_subtitle_row(ws, 21, "三、蓝牙音频编码技术对比", 7)
    write_header_row(ws, 22, [
        "编码", "提出方", "最高码率", "延迟", "无损", "使用厂商", "趋势"
    ])
    codec = [
        ["SBC",        "蓝牙SIG",    "328 kbps",  "高",   "否", "通用",         "基础兼容,逐步被LC3替代"],
        ["AAC",        "Fraunhofer", "256 kbps",  "中高", "否", "苹果/索尼",     "苹果主力编码"],
        ["aptX HD",    "高通",       "576 kbps",  "中",   "否", "Bose/铁三角",   "安卓高解析"],
        ["aptX Lossless","高通",     "1.2 Mbps",  "中",   "是", "Bose/高品安卓", "高通无损主推"],
        ["LDAC",       "索尼",       "990 kbps",  "中高", "近无损","索尼/小米",  "安卓Hi-Res事实标准"],
        ["LHDC 5.0",   "盛微/恒玄",  "1 Mbps+",   "中",   "近无损","小米/漫步者", "国产高码率路线"],
        ["L2HC 4.0",   "华为",       "1.5 Mbps",  "低",   "是",  "华为",         "华为无损+星闪"],
        ["LC3/LC3+",   "蓝牙SIG",    "345 kbps",  "低",   "否",  "LE Audio通用", "LE Audio标配编码"],
    ]
    for i, d in enumerate(codec):
        write_data_row(ws, 23 + i, d, height=28)

    add_pie(ws, "头戴耳机营收份额", [5, 6, 7, 8, 9, 10, 11], 1, 3, "A31", w=16, h=11)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 4: 智能音箱与蓝牙音箱
# ═══════════════════════════════════════════════════════════════
def build_smart_bluetooth_sheet(wb):
    ws = wb.create_sheet("智能与蓝牙音箱")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 14, "D": 18, "E": 24, "F": 34, "G": 30})

    write_title_row(ws, 1, "智能音箱与蓝牙音箱市场分析", 7)

    # ── 智能音箱份额 ──
    write_subtitle_row(ws, 3, "一、智能音箱厂商份额（2024年）", 7)
    write_header_row(ws, 4, [
        "厂商", "出货量份额", "营收份额", "语音平台", "主力产品", "技术特点", "核心优势"
    ])
    smart = [
        ["亚马逊Amazon", 0.28, 0.30, "Alexa",
         "Echo Dot/Pop/Show\nEcho Studio",
         "远场麦克风阵列6+\nZigbee/Matter网关\nAlexa大模型升级\n带屏Echo Show",
         "智能音箱鼻祖\nAlexa生态最广\n电商渠道绑定"],
        ["谷歌Google",   0.27, 0.25, "Google Assistant",
         "Nest Mini/Audio\nNest Hub",
         "远场麦克风\nGoogle Assistant+Gemini\nCast投屏\n音频识别领先",
         "Gemini AI加持\n搜索+影音协同\n安卓生态联动"],
        ["百度Baidu",    0.20, 0.15, "小度",
         "小度智能音箱\n添添/带屏版",
         "远场阵列+小度\n文心大模型升级\n带屏智能音箱主力\n儿童教育场景",
         "国内份额第一\n带屏音箱开创者\n教育/家庭场景"],
        ["阿里Ali",     0.08, 0.06, "天猫精灵",
         "天猫精灵/X5\n带屏CC/IN",
         "远场阵列+通义\n购物+IoT网关\nZigble/Matter",
         "电商+IoT协同\n国内第二\n购物场景差异化"],
        ["苹果Apple",   0.05, 0.10, "Siri",
         "HomePod mini\nHomePod(2023)",
         "Apple Music高解析\n空间音频+Siri\n计算音频+感知\nHomeKit中枢",
         "音质口碑回流\nHomeKit生态\nApple Music绑定"],
        ["哈曼/JBL",    0.05, 0.06, "Alexa/OK Google",
         "JBL Link\nHarman Kardon系列",
         "JBL声学+Alexa\n高保真单元\n便携智能结合",
         "音质+智能结合\nOEM代工能力强"],
        ["小米Xiaomi",  0.04, 0.03, "小爱",
         "小爱音箱\nRedmi带屏",
         "远场+小爱同学\n米家IoT联动\n性价比",
         "米家IoT生态\n性价比路线"],
        ["其他",        0.03, 0.05, "—", "—", "—", "利基/区域品牌"],
    ]
    fmts = [None, "0.0%", "0.0%", None, None, None, None]
    for i, d in enumerate(smart):
        write_data_row(ws, 5 + i, d, fmts, height=56)

    # ── 蓝牙音箱份额 ──
    write_subtitle_row(ws, 14, "二、便携蓝牙音箱厂商份额（2024年）", 7)
    write_header_row(ws, 15, [
        "厂商", "出货量份额", "营收份额", "主打品类", "主力产品", "技术特点", "核心优势"
    ])
    bt = [
        ["JBL(哈曼)",   0.25, 0.22, "便携户外",
         "Flip 6/Charge 5\nBoombox 3/Go 4",
         "被动辐射器低音\nIP67防水防尘\nPartyBoost多连\n蓝牙5.3+LE Audio",
         "便携音箱全球第一\n品牌+渠道双优\n哈曼供应链"],
        ["索尼Sony",    0.12, 0.14, "高解析便携",
         "SRS-XB系列\nULT 1/Field系列",
         "EXTRA BASS低音\nIP67防水\n360声场\nLDAC无损",
         "音质+防水双线\nLDAC差异化\n发烧用户群"],
        ["BOSE",        0.10, 0.12, "便携高端",
         "SoundLink Flex\nSoundLink Max",
         "PositionIQ方位感知\nIP67防水\n单声道高保真\n定制驱动单元",
         "高端便携口碑\n商务+户外双覆盖\n品牌溢价"],
        ["马歇尔Marshall",0.08, 0.10,"复古便携",
         "Emberton II\nWillen/Kilburn",
         "复古吉他音箱外观\n360声场\nIP67防水\n布鲁姆林声学",
         "复古设计溢价\n年轻潮流用户\n礼物场景强"],
        ["漫步者Edifier", 0.10, 0.06, "性价比便携",
         "MR series/M228\nW系列",
         "国产方案+性价比\nANC+蓝牙5.3\n多场景细分",
         "国内量大\n性价比+多品类\n国产替代"],
        ["Ultimate Ears", 0.06, 0.05,"户外便携",
         "BOOM/MEGABOOM\nWONDERBOOM",
         "360声场+IP67\nMagSafe磁性\nPartyUp多连",
         "户外场景标杆\n罗技旗下\n潮流用户"],
        ["B&O",          0.03, 0.06, "高端便携",
         "Beosound A1/A5\nEdge",
         "Bang&Olufsen调音\n铝合金+真皮\n空间声学\nAirPlay 2",
         "奢侈品定位\n设计+音质标杆\n高端利基"],
        ["其他(小米/Anker等)",0.26,0.25,"—",
         "小米方盒子/Anker Soundcore",
         "国产方案+性价比\nIPX7+大容量电池",
         "中低端量大\n性价比路线"],
    ]
    for i, d in enumerate(bt):
        write_data_row(ws, 16 + i, d, fmts, height=48)

    # ── 智能音箱AI能力 ──
    write_subtitle_row(ws, 25, "三、智能音箱 AI 大模型能力对比", 7)
    write_header_row(ws, 26, [
        "平台", "模型", "能力", "内容生态", "多模态", "智能家庭", "趋势"
    ])
    ai = [
        ["Alexa",     "Alexa LLM",  "对话/任务执行",    "Amazon Music",     "带屏/语音", "Zigbee/Matter",   "生成式AI重做交互"],
        ["Google",    "Gemini",     "多模态/搜索优势",  "YouTube/Spotify",  "强",         "Matter/Google Home","Gemini Nano端侧"],
        ["小度",      "文心一言",    "中文对话/教育",     "百度音乐/爱奇艺",  "带屏",       "度家IoT",          "文心大模型升级"],
        ["天猫精灵",  "通义千问",    "购物/助手",        "淘宝/优酷",        "带屏",       "阿里IoT",          "通义模型赋能"],
        ["Siri",      "Apple Intelligence","生成式AI",   "Apple Music",      "弱",         "HomeKit",          "Apple Intelligence落地"],
        ["小爱同学",  "小米大模型",  "中文对话/IoT",      "小米音乐/视频",     "带屏",       "米家IoT",          "米家生态强绑定"],
    ]
    for i, d in enumerate(ai):
        write_data_row(ws, 27 + i, d, height=32)

    add_pie(ws, "智能音箱出货量份额", [5, 6, 7, 8, 9, 10, 11, 12], 1, 2, "A34", w=15, h=10)
    add_pie(ws, "蓝牙音箱营收份额", [16, 17, 18, 19, 20, 21, 22, 23], 1, 3, "I34", w=15, h=10)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 5: 家庭影院与回音壁
# ═══════════════════════════════════════════════════════════════
def build_home_theater_sheet(wb):
    ws = wb.create_sheet("家庭影院市场分析")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 14, "D": 18, "E": 24, "F": 34, "G": 30})

    write_title_row(ws, 1, "家庭影院与回音壁市场分析", 7)

    # ── 回音壁份额 ──
    write_subtitle_row(ws, 3, "一、回音壁/家庭影院厂商份额（2024年）", 7)
    write_header_row(ws, 4, [
        "厂商", "出货量份额", "营收份额", "音频平台", "主力产品", "技术特点", "核心优势"
    ])
    sb = [
        ["索尼Sony",    0.20, 0.22, "360RA/杜比",
         "HT-A9/A7000\nHT-S2000",
         "360 Spatial Sound Mapper\n杜比全景声/DTS:X\n无线环绕+低音炮\nSA-RS5后环绕",
         "全景声+PS5协同\n360RA独家\n索尼影音生态"],
        ["三星Samsung", 0.22, 0.20, "Q-Symphony",
         "HW-Q990D/Q800D\nHW-S601",
         "Q-Symphony电视协同\n杜比全景声11.1.4\nSpace Sound Fit\n无线后环绕",
         "三星电视协同\nSoundbar份额第一\n空间音频下放"],
        ["LG",          0.12, 0.12, "WOW Orchestra",
         "S95TR/S90TY\nSC9S",
         "WOW Orchestra电视协同\n杜比全景声/IMAX Enhanced\nAI音质校准\nMeridian调音",
         "LG OLED协同\nMeridian音质\nOLED电视捆绑"],
        ["Bose",        0.10, 0.12, "Bose自有",
         "Smart Soundbar 900\nUltra Soundbar",
         "Bose空间音频\n自适应音质\nQuietPort低音\nAirPlay2",
         "高端口碑\n简洁设计\n品牌溢价"],
        ["Sonos",       0.10, 0.14, "Sonos平台",
         "Arc Ultra/Beam\nEra 300",
         "多房间音频平台\n杜比全景声11.1.4\nSonos无线生态\nTrueplay调音",
         "多房间生态标杆\n流媒体平台\n苹果HomeKit协同"],
        ["JBL/Harman",  0.08, 0.06, "杜比",
         "Bar 1000/500\nBar 9.1",
         "可拆卸环绕\n杜比全景声/DTS:X\nMultiBeam空间声\n5.1+Harman调音",
         "可拆卸环绕创新\n哈曼供应链\n性价比高端"],
        ["天龙Denon/马兰士",0.05,0.08,"杜比/DTS",
         "DHT-S series\nAV Receiver",
         "AV功放血统\n杜比全景声/DTS:X\nHEOS多房间\n8K HDMI",
         "AV功放继承\n发烧用户忠诚\nHi-Fi血统"],
        ["其他(雅马哈/小米等)",0.13,0.06,"—",
         "小米Sound/YAS系列",
         "性价比+杜比\n部分多房间",
         "中低端+细分\n性价比路线"],
    ]
    fmts = [None, "0.0%", "0.0%", None, None, None, None]
    for i, d in enumerate(sb):
        write_data_row(ws, 5 + i, d, fmts, height=56)

    # ── 全景声技术 ──
    write_subtitle_row(ws, 14, "二、空间音频/全景声技术对比", 7)
    write_header_row(ws, 15, [
        "技术", "提出方", "声道规格", "渲染方式", "授权情况", "代表产品", "趋势"
    ])
    spatial = [
        ["杜比全景声",   "杜比",     "5.1.4-11.1.4", "基于对象",  "需授权",     "Sony/Samsung/LG Soundbar","客厅全景声事实标准"],
        ["DTS:X",       "DTS",     "灵活布局",      "基于对象+元数据","需授权",  "天龙/JBL/三星",          "与杜比全景声并行"],
        ["360 Reality Audio","索尼","9.1声道",       "基于对象",  "开放+授权",  "Sony耳机/音箱",          "索尼独有,生态小"],
        ["IMAX Enhanced","IMAX/DTS","增强版DTS:X",   "基于对象",  "授权认证",   "LG/天龙/海信",           "影院级体验"],
        ["MPEG-H 3D Audio","Fraunhofer","22.2声道",   "基于场景+对象","开放",   "韩国ATSC 3.0",          "广播标准,消费少"],
        ["Apple 空间音频","苹果",    "动态+固定",     "基于对象+头追","iOS自有","AirPods/HomePod",        "移动端最普及"],
    ]
    for i, d in enumerate(spatial):
        write_data_row(ws, 16 + i, d, height=32)

    # ── 家庭影院趋势 ──
    write_subtitle_row(ws, 23, "三、家庭影院技术趋势", 7)
    write_header_row(ws, 24, [
        "趋势", "说明", "代表厂商/产品", "驱动力", "时间线", "影响", "挑战"
    ])
    ht_trend = [
        ["全景声下沉",   "杜比全景声/DTS:X下放到中端\n5.1.2成起步配置",       "三星HW-S系列\n索尼HT-S2000",  "内容生态",     "2024-2026","高",   "授权费/算法实现"],
        ["无线环绕",     "无线后环绕+无线低音炮\n减少布线",                   "三星Q-Symphony\nBose Soundbar","用户体验",     "2024-2027","中高", "延迟/供电/抗干扰"],
        ["电视协同",     "Soundbar+电视扬声器协同发声\n提升声场",             "三星/LG/索尼",                 "电视厂商绑定",  "2024-2026","中高", "生态封闭/兼容性"],
        ["AI音质校准",   "AI房间声学校准\n自适应频响",                        "LG WOW Calibration\nSonos Trueplay","AI+声学","2025-2027","中","麦克风精度/算力"],
        ["8K+HDMI 2.1", "8K直通+VRR/ALLM\n游戏场景",                        "天龙/马兰士AV功放",            "次世代游戏",    "2025-2026","中",  "8K内容稀缺/成本"],
    ]
    for i, d in enumerate(ht_trend):
        write_data_row(ws, 25 + i, d, height=44)

    add_pie(ws, "回音壁出货量份额", [5, 6, 7, 8, 9, 10, 11, 12], 1, 2, "A32", w=15, h=11)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 6: 专业音响
# ═══════════════════════════════════════════════════════════════
def build_pro_audio_sheet(wb):
    ws = wb.create_sheet("专业音响市场分析")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 16, "D": 18, "E": 24, "F": 34, "G": 28})

    write_title_row(ws, 1, "专业音响市场厂商份额与技术分析", 7)

    # ── 专业麦克风/监听 ──
    write_subtitle_row(ws, 3, "一、专业麦克风/监听耳机厂商份额（2024年）", 7)
    write_header_row(ws, 4, [
        "厂商", "市场份额", "主力品类", "主力产品", "技术特点", "核心优势", "主要市场"
    ])
    pro = [
        ["Sennheiser",   0.18, "动圈/电容麦\n监听耳机", "MK4/MK8\nHD 25/280 Pro",
         "德式精密声学\n射频无线话筒领航\nEvolution无线系列\nNeumann合并",
         "专业录音/演出标杆\nNeumann高端调音传承", "录音棚/演出/广电"],
        ["Shure",        0.16, "动圈/无线话筒",        "SM58/SM57\nBeta系列/SLX-D",
         "SM58行业经典60年\n无线话筒射频领航\nQLX-D/SLX-D数字无线\nMXW会议系统",
         "现场演出标配\nSM58不可替代",               "演出/会议/广播"],
        ["Audio-Technica",0.12,"电容/动圈\n监听耳机",  "AT2020/AT4040\nATH-M50x",
         "性价比电容麦\nATH-M50x监听标杆\nUNIPACK无线话筒",
         "性价比+品质均衡\n监听耳机销量王",            "录音/直播/DJ"],
        ["Sony Pro",     0.10, "无线话筒\n专业监听",    "C100/C800G\nMDR-7506",
         "C800G电子管电容麦\nMDR-7506监听标杆\n数字无线DWX系列",
         "录音棚标杆\n监听耳机经典",                   "录音棚/广电/演出"],
    ]
    fmts = [None, "0.0%", None, None, None, None, None]
    rows_data = [
        ["Sennheiser",   0.18, "动圈/电容麦\n监听耳机", "MK4/MK8\nHD 25/280 Pro",
         "德式精密声学\n射频无线话筒领航\nEvolution无线系列\nNeumann合并",
         "专业录音/演出标杆\nNeumann高端调音传承", "录音棚/演出/广电"],
        ["Shure",        0.16, "动圈/无线话筒",        "SM58/SM57\nBeta系列/SLX-D",
         "SM58行业经典60年\n无线话筒射频领航\nQLX-D/SLX-D数字无线\nMXW会议系统",
         "现场演出标配\nSM58不可替代",               "演出/会议/广播"],
        ["Audio-Technica",0.12,"电容/动圈\n监听耳机",  "AT2020/AT4040\nATH-M50x",
         "性价比电容麦\nATH-M50x监听标杆\nUNIPACK无线话筒",
         "性价比+品质均衡\n监听耳机销量王",            "录音/直播/DJ"],
        ["Sony Pro",     0.10, "无线话筒\n专业监听",    "C100/C800G\nMDR-7506",
         "C800G电子管电容麦\nMDR-7506监听标杆\n数字无线DWX系列",
         "录音棚标杆\n监听耳机经典",                   "录音棚/广电/演出"],
        ["Bose Pro",     0.06, "便携舞台系统",         "L1 Pro\nS1 Pro+",
         "L1线阵列便携舞台\nS1 Pro+多场景\n被动全频",
         "便携演出标杆\n小型场馆首选",                  "小型演出/演讲"],
        ["Yamaha Pro",   0.06, "调音台/监听音箱",      "QL/CL系列\nDXR/DZR",
         "数字调音台QL/CL\nDante网络音频\nDXR系列有源监听",
         "广电/演出调音台标准\n网络音频领航",            "演出/广电/会议"],
        ["Beyerdynamic", 0.05, "动圈话筒\n监听耳机",     "TG V70/TG V50\nDT 770/990",
         "德式动圈话筒\nDT系列监听口碑\nREV无线系列",
         "德式监听血统\n直播/录音忠诚",                  "录音/直播/广播"],
        ["其他(Rode/AKG等)",0.17,"—","Rode NT1/AKGC414","—",
         "细分领域多\nRode/AKg/Neumann等",
         "细分忠诚市场"],
    ]
    for i, d in enumerate(rows_data):
        write_data_row(ws, 5 + i, d, fmts, height=56)

    # ── 数字无线音频 ──
    write_subtitle_row(ws, 14, "二、专业数字无线音频技术对比", 7)
    write_header_row(ws, 15, [
        "技术", "提出方", "频段/带宽", "延迟", "音质", "代表产品", "趋势"
    ])
    wireless = [
        ["UHF模拟",    "通用",        "470-608MHz",   "近零", "中",   "Shure SM58无线版",   "逐步被数字替代"],
        ["UHF数字",    "Shure/Senn", "470-608MHz",   "极低", "高",   "Shure QLX-D/SLX-D",  "专业演出主流"],
        ["2.4GHz数字","Shure/Senn",  "2.4GHz ISM",   "中低", "中高", "Shure BLX/DLX-D",    "免许可,便携但易干扰"],
        ["Dante",     "Audinate",    "以太网",        "极低", "极高", "Yamaha QL/CL\nShure MXW", "网络音频事实标准"],
        ["AES67/RAVENNA","AES",      "以太网",        "极低", "极高","广电/大型安装",      "跨厂商互通标准"],
        ["Auracast",  "蓝牙SIG",     "2.4GHz BLE",   "低",   "中",   "LE Audio广播音频",   "广播音频新方向"],
    ]
    for i, d in enumerate(wireless):
        write_data_row(ws, 16 + i, d, height=32)

    add_pie(ws, "专业麦克风/监听份额", [5, 6, 7, 8, 9, 10, 11, 12], 1, 2, "A23", w=15, h=11)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 7: 汽车音响
# ═══════════════════════════════════════════════════════════════
def build_auto_audio_sheet(wb):
    ws = wb.create_sheet("汽车音响市场分析")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 16, "D": 20, "E": 24, "F": 34, "G": 28})

    write_title_row(ws, 1, "汽车音响市场厂商份额与技术分析", 7)

    # ── OEM汽车音响份额 ──
    write_subtitle_row(ws, 3, "一、OEM 前装汽车音响厂商份额（2024年）", 7)
    write_header_row(ws, 4, [
        "厂商", "市场份额", "音响类型", "主力品牌/客户", "技术特点", "核心优势", "主要市场"
    ])
    auto = [
        ["哈曼Harman",   0.35, "OEM全套\n后装",  "JBL/AKG/HK\nBMW/Mercedes/BYD",
         "多声道空间音频\n主动声学ANC\nQuantumLogic环绕\n车载生态整合",
         "OEM份额绝对第一\n三星旗下供应链\n中高端车型覆盖广", "欧美/中国高端车"],
        ["Bose Auto",   0.20, "OEM高端",       "Bose\nNissan/Mazda/Cadillac",
         "Bose空间音频\nAudioPilot噪声补偿\nCenterpoint环绕\n主动降噪ANC",
         "高端OEM口碑\n舒适性口碑标杆",         "日美/部分中国车"],
        ["Bang&Olufsen",0.10, "OEM奢侈",       "B&O/B&O Play\nAudi/Lambo/Ferrari",
         "Acoustic Lens声学透镜\n主动分频\n铝合金扬声器\n设计+音质双标杆",
         "奢侈品牌溢价\n设计美学",               "欧洲超豪华车"],
        ["Burmester",   0.05, "OEM顶级",       "Burmester\nMercedes S/Porsche",
         "高端分立功放\n镀膜高音\n航空级铝合金\n手工调校",
         "顶级Hi-Fi血统\n德国精密",              "德系豪华车"],
        ["Meridian",    0.04, "OEM高端",       "Meridian\nJaguar/Land Rover",
         "British声学调校\nTrifield 3D\nDSP数字分频",
         "英式Hi-Fi传承",                       "英系豪华车"],
        ["Dynaudio",    0.03, "OEM高端",       "丹拿\nVW/Bugatti",
         "丹麦手工单元\n软球顶高音\nMSP低音单元",
         "丹麦Hi-Fi血统",                        "德系/北欧豪华车"],
        ["雅马哈Yamaha", 0.04, "OEM高端",       "Yamaha\nLexus/Toyota",
         "Hi-Fi血统\n矩阵声场\n主动声学",
         "乐器+Hi-Fi传承",                       "日系高端车"],
        ["Dynaudio/infinity",0.06,"OEM中端",  "Infinity/JBL\n大众/通用",
         "哈曼中端品牌\n多声道方案",
         "性价比+哈曼供应链",                    "中端车型"],
        ["自研系(特斯拉/比亚迪等)",0.13,"自研",  "Tesla/BYD/NIO",
         "多声道主动声学\n空间音频\n软件定义音响\nAI声场校准",
         "新能源车自研趋势\nOTA升级音质",        "新能源车"],
    ]
    fmts = [None, "0.0%", None, None, None, None, None]
    for i, d in enumerate(auto):
        write_data_row(ws, 5 + i, d, fmts, height=56)

    # ── 新能源车音响趋势 ──
    write_subtitle_row(ws, 15, "二、新能源车音响技术趋势", 7)
    write_header_row(ws, 16, [
        "趋势", "说明", "代表厂商/产品", "驱动力", "时间线", "影响", "挑战"
    ])
    auto_trend = [
        ["多声道空间音频", "7.1.4/11.1.4多声道\n杜比全景声车内落地", "奔驰/比亚迪/特斯拉", "座舱升级",   "2024-2026","高",   "扬声器数量/功放/调音"],
        ["主动声学ANC",    "路噪主动降噪\nRNC发动机/路噪消除",       "哈曼/Bose/比亚迪",  "EV安静度反差","2024-2027","中高", "算法/参考麦克风"],
        ["软件定义音响",   "OTA升级音质\n场景化调音\n用户自定义",     "特斯拉/蔚来/小鹏",  "智能化",     "2025-2028","中高", "DSP算力/算法生态"],
        ["座舱融合",       "音响+HUD+氛围灯+座椅联动\n沉浸式座舱",     "比亚迪/华为/理想",  "新势力竞争", "2025-2027","中",   "跨域融合架构"],
        ["自研音响",       "新势力自研音响系统\n摆脱Tier1依赖",       "特斯拉/比亚迪/蔚来","降本+定制",  "2025-2028","中",   "声学积累/调音能力"],
    ]
    for i, d in enumerate(auto_trend):
        write_data_row(ws, 17 + i, d, height=44)

    # ── 车载音频生态 ──
    write_subtitle_row(ws, 23, "三、车载音频平台与生态对比", 7)
    write_header_row(ws, 24, [
        "平台", "提出方", "能力", "内容生态", "生态绑定", "代表车机", "趋势"
    ])
    plat = [
        ["CarPlay",    "苹果",  "投屏+空间音频", "Apple Music",     "iOS生态",     "几乎所有车型",      "下一代CarPlay接管整车"],
        ["Android Auto","谷歌", "投屏+Gemini",   "YouTube/Spotify", "Android生态", "安卓阵营车型",      "Gemini AI车内集成"],
        ["Harman OS",  "哈曼",  "车机+音响融合", "聚合应用",        "Tier1生态",   "BMW/Mercedes",     "Tier1车机音响一体"],
        ["鸿蒙座舱",   "华为",  "多设备协同",    "华为音乐/视频",   "鸿蒙生态",    "问界/智界",         "国产车机一体化"],
        ["比亚迪DiLink","比亚迪","自研座舱+音响",  "DiLink应用商店",  "自研生态",    "王朝/海洋系列",     "新势力自研代表"],
    ]
    for i, d in enumerate(plat):
        write_data_row(ws, 25 + i, d, height=32)

    add_pie(ws, "OEM汽车音响厂商份额", [5, 6, 7, 8, 9, 10, 11, 12, 13], 1, 2, "A31", w=16, h=11)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 8: 发展趋势与建议
# ═══════════════════════════════════════════════════════════════
def build_trends_sheet(wb):
    ws = wb.create_sheet("发展趋势与建议")
    set_col_widths(ws, {"A": 8, "B": 20, "C": 36, "D": 20, "E": 14, "F": 14, "G": 28})

    write_title_row(ws, 1, "音响市场发展趋势与投资建议", 7)

    # ── 核心趋势 ──
    write_subtitle_row(ws, 3, "一、2025-2027年核心趋势", 7)
    write_header_row(ws, 4, [
        "序号", "趋势方向", "详细说明", "受益领域", "确定性", "影响", "关键变量"
    ])
    trends = [
        ["1", "空间音频全面普及",
         "杜比全景声/360RA下沉中端\n头追/动态跟踪成标配\nTWS+音箱+车机全覆盖",
         "TWS/音箱/汽车音响", "高", "★★★★",
         "内容生态/算法精度"],
        ["2", "AI驱动音频体验",
         "AI降噪+AI通话+AI声学校准\n生成式AI进入智能音箱\n端侧AI芯片成关键",
         "耳机/智能音箱/汽车", "高", "★★★★",
         "端侧算力/AI模型"],
        ["3", "自研声学芯片",
         "苹果H2/H3、索尼V2/V3、华为海思\n芯片+算法+续航一体化\n大厂自研趋势加速",
         "TWS/头戴耳机", "中高", "★★★★",
         "芯片研发投入/规模"],
        ["4", "无损音频普及",
         "LE Audio/LDAC/L2HC无损\n高码率蓝牙编码标准化\n流媒体Hi-Res催化",
         "全品类无线音频", "高", "★★★",
         "编码互通/内容供给"],
        ["5", "汽车音响升级",
         "新能源车标配高端音响\n多声道/主动声学/OTA\n新势力自研替代Tier1",
         "汽车音响OEM", "高", "★★★★",
         "新能源车销量/座舱升级"],
        ["6", "TWS向可穿戴延伸",
         "开放式/带屏AirPods/健康监测\nTWS向耳穿戴+健康延伸\n存量升级周期",
         "TWS/可穿戴", "中高", "★★★",
         "健康传感器/形态创新"],
    ]
    for i, d in enumerate(trends):
        write_data_row(ws, 5 + i, d, height=56)

    # ── 价值链分布 ──
    write_subtitle_row(ws, 12, "二、音响产业链价值分布", 7)
    write_header_row(ws, 13, [
        "环节", "说明", "代表厂商", "价值占比", "壁垒", "国产化", "趋势"
    ])
    chain = [
        ["品牌与设计",   "产品定义/声学调校/品牌",  "苹果/索尼/Bose/JBL",     "35-40%", "极高", "低(高端品牌)",
         "中国品牌向上突破"],
        ["芯片/平台",    "蓝牙SoC/ANC芯片/功放",   "恒玄/络达/高通/苹果",    "15-20%", "高",   "中(恒玄/络达)",
         "自研芯片成护城河"],
        ["声学单元",     "扬声器/麦克风/振膜",      "歌尔/瑞声/声学电子",    "15-20%", "中高", "高",
         "歌尔/瑞声全球领先"],
        ["整机制造",     "OEM/ODM代工",             "歌尔/瑞声/万魔/通力",   "20-25%", "中",   "高",
         "中国OEM主导全球"],
        ["软件与算法",   "ANC/空间音频/AI算法",     "苹果/索尼/恒玄/杜比",   "5-10%",  "高",   "中",
         "AI算法价值提升"],
    ]
    for i, d in enumerate(chain):
        write_data_row(ws, 14 + i, d, height=36)

    # ── 投资关注方向 ──
    write_subtitle_row(ws, 20, "三、投资关注方向", 7)
    write_header_row(ws, 21, [
        "方向", "逻辑", "标的类型", "时间窗口", "确定性", "风险", "风险提示"
    ])
    invest = [
        ["TWS高端升级",
         "AirPods/Buds3 Pro带屏+健康\n高端TWS存量升级周期",
         "苹果/三星/索尼供应链", "2025-2027", "中高", "中",
         "创新不及预期/价格战"],
        ["汽车音响",
         "新能源车标配高端音响\n新势力自研+Tier1双线",
         "哈曼/Bose/国产Tier1", "2025-2028", "高", "中",
         "新能源车增速放缓/自研替代"],
        ["声学芯片",
         "自研芯片成差异化核心\n恒玄/络达国产替代加速",
         "恒玄/络达/海思", "2025-2027", "中高", "中",
         "消费电子需求波动/竞争加剧"],
        ["空间音频生态",
         "杜比全景声/360RA生态扩张\n内容+硬件双向催化",
         "杜比/索尼/内容方", "2025-2028", "中高", "中",
         "内容供给/编码互通"],
        ["AI音频",
         "AI降噪+AI通话+生成式音箱\n端侧AI芯片驱动",
         "苹果/恒玄/算法公司", "2025-2028", "中", "中高",
         "AI落地节奏/隐私合规"],
    ]
    for i, d in enumerate(invest):
        write_data_row(ws, 22 + i, d, height=48)

    # ── 全球厂商综合评分 ──
    write_subtitle_row(ws, 28, "四、全球主要音响厂商综合竞争力评分（1-10分）", 7)
    write_header_row(ws, 29, [
        "厂商", "品牌力", "声学技术", "芯片自研", "生态绑定", "市场地位", "综合评分"
    ])
    score = [
        ["苹果Apple",   10, 9, 10, 10, 9,  9.6],
        ["索尼Sony",     9, 9,  8,  7, 8,  8.2],
        ["Bose",         9, 8,  6,  5, 7,  7.0],
        ["三星Samsung",  8, 7,  6,  8, 8,  7.4],
        ["哈曼JBL",      8, 8,  5,  6, 9,  7.2],
        ["Sennheiser",   8, 9,  4,  4, 6,  6.2],
        ["华为Huawei",   7, 7,  9,  8, 7,  7.6],
        ["小米Xiaomi",   7, 6,  4,  7, 7,  6.2],
        ["漫步者Edifier", 6,6,   4,  4, 7,  5.4],
    ]
    for i, d in enumerate(score):
        write_data_row(ws, 30 + i, d, bold_cols={7})
        ws.cell(row=30+i, column=7).font = _font(size=11, bold=True, color=C["deep_brown"])

    add_bar(ws, "全球主要音响厂商综合竞争力评分",
            range(30, 39), 1, [2, 3, 4, 5, 6], "A40", w=22, h=13)

    return ws


# ═══════════════════════════════════════════════════════════════
# 主入口
# ═══════════════════════════════════════════════════════════════
def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    sheets = [
        ("市场概览",           build_overview),
        ("TWS耳机市场分析",    build_tws_sheet),
        ("头戴耳机市场分析",   build_headphone_sheet),
        ("智能与蓝牙音箱",     build_smart_bluetooth_sheet),
        ("家庭影院市场分析",   build_home_theater_sheet),
        ("专业音响市场分析",   build_pro_audio_sheet),
        ("汽车音响市场分析",   build_auto_audio_sheet),
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
