"""
智能手表市场分析报告生成器
覆盖 旗舰智能手表、运动健康手表、轻智能手表、儿童手表、健康监测 等细分市场
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
OUTPUT_FILE = OUTPUT_DIR / "智能手表市场分析报告.xlsx"

# ── 配色方案 ──
C = {
    "title_bg":    "2A2A4A",   # 深紫(可穿戴主题)
    "title_fg":    "FFFFFF",
    "header_bg":   "4A4A7A",   # 钢蓝紫
    "header_fg":    "FFFFFF",
    "sub_bg":      "DEDEEC",   # 浅紫灰
    "accent1":     "4472C4",
    "accent2":     "ED7D31",
    "accent3":     "70AD47",
    "light_gray":  "F0F0F6",
    "white":       "FFFFFF",
    "deep_purple": "2A2A4A",
    "red":         "C00000",
    "green":       "008000",
}

# ── 通用样式工厂 ──
thin_border = Border(
    left=Side(style="thin", color="B0B0CC"),
    right=Side(style="thin", color="B0B0CC"),
    top=Side(style="thin", color="B0B0CC"),
    bottom=Side(style="thin", color="B0B0CC"),
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
    cell.font = _font(size=12, bold=True, color=C["deep_purple"])
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

    write_title_row(ws, 1, "全球智能手表市场分析报告（2024-2025）", 7)
    write_source_row(ws, 2, "数据来源：IDC、Counterpoint、Canalys、Strategy Analytics 等公开报告 | 制表日期：2026年7月", 7)

    # ── 细分市场规模 ──
    write_subtitle_row(ws, 4, "一、智能手表细分市场规模", 7)
    write_header_row(ws, 5, [
        "细分市场", "2023年(百万台)", "2024年(百万台)", "同比增速",
        "2025年预测(百万台)", "主要应用", "市场特征"
    ])
    data = [
        ["旗舰智能手表",   95,  105, 0.105, 118, "全功能/生态联动",  "苹果/华为/三星主导,生态壁垒高"],
        ["运动健康手表",   78,  88,  0.128, 100, "运动/户外/健康",    "Garmin/Amazfit/博能,专业赛道"],
        ["轻智能手表",     55,  62,  0.127, 70,  "传统腕表+智能",     "卡西欧/天梭/佳明Fenix,长续航"],
        ["儿童手表",       42,  46,  0.095, 50,  "儿童安全/定位",     "小天才/华为/360主导国内"],
        ["健康监测表带",   18,  22,  0.222, 28,  "医疗/慢病监测",     "Withings/Fitbit,医疗化方向"],
        ["入门手环/表",   110, 105, -0.045, 100, "基础运动/计步",     "小米/华为手环,被手表替代"],
    ]
    fmts = [None, "#,##0", "#,##0", "0.0%", "#,##0", None, None]
    for i, d in enumerate(data):
        write_data_row(ws, 6 + i, d, fmts)

    write_data_row(ws, 12, ["合计", 398, 428, 0.075, 466, "—", "旗舰+运动是双引擎"],
                   fmts, bold_cols={1, 2, 3, 4, 5})

    # ── 市场集中度 ──
    write_subtitle_row(ws, 14, "二、各细分市场集中度（CR3）", 7)
    write_header_row(ws, 15, [
        "细分市场", "CR1", "CR2", "CR3", "头部厂商", "竞争格局", "进入壁垒"
    ])
    cr = [
        ["旗舰智能手表", "32%", "55%", "75%", "苹果/华为/三星",       "苹果领跑,安卓追赶",   "极高(生态+芯片+健康)"],
        ["运动健康手表", "30%", "50%", "68%", "Garmin/Amazfit/华为",  "Garmin领先,国产品牌追赶","高(传感器+算法+户外)"],
        ["轻智能手表",   "20%", "38%", "55%", "卡西欧/佳明/天梭",     "传统+智能融合",        "中(腕表文化+续航)"],
        ["儿童手表",     "55%", "80%", "92%", "小天才/华为/360",      "小天才国内绝对第一",   "高(儿童生态+安全)"],
        ["健康监测表带", "25%", "45%", "65%", "Withings/Fitbit/欧姆龙","医疗化细分",           "高(医疗认证+算法)"],
        ["入门手环/表", "30%", "55%", "72%", "小米/华为/Fitbit",      "手环被手表替代",        "中(性价比+量)"],
    ]
    for i, d in enumerate(cr):
        write_data_row(ws, 16 + i, d)

    # ── 关键趋势 ──
    write_subtitle_row(ws, 23, "三、2025年关键趋势", 7)
    write_header_row(ws, 24, ["序号", "趋势", "说明", "主要推进方", "确定性", "影响", "关键变量"])
    trends = [
        ["1", "健康监测医疗化", "ECG/血压/血糖/体温连续监测\n慢病管理+医疗认证升级",  "苹果/华为/Withings","高",  "★★★★★","传感器精度/医疗认证"],
        ["2", "自研芯片",       "苹果S系列/海思/瑞昱\n芯片+算法+续航一体化",          "苹果/华为/瑞昱",    "中高","★★★★","芯片研发投入/规模"],
        ["3", "AI助手入腕",     "端侧AI语音助手\nLLM问答+健康分析",                   "苹果/华为/三星",    "中高","★★★",  "端侧算力/AI模型"],
        ["4", "卫星通信",       "北斗/苹果紧急SOS\nGarmin inReach户外救援",            "苹果/Garmin/北斗", "高",  "★★★",  "卫星生态/功耗"],
        ["5", "续航与快充",     "长续航(7-14天)+快充\n轻智能路线回潮",                 "华为/Amazfit/卡西欧","高","★★★","电池能量密度/功耗优化"],
        ["6", "无感健康生态",   "手表+耳机+戒指+秤联动\n健康数据中枢",                 "苹果/三星/华为",    "中高","★★★", "跨设备协同/隐私"],
    ]
    for i, d in enumerate(trends):
        write_data_row(ws, 25 + i, d, height=44)

    add_bar(ws, "智能手表细分市场规模对比（百万台）",
            range(6, 12), 1, [2, 4], "A33", w=22, h=12)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 2: 旗舰智能手表市场
# ═══════════════════════════════════════════════════════════════
def build_flagship_sheet(wb):
    ws = wb.create_sheet("旗舰智能手表")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 14, "D": 16, "E": 22, "F": 34, "G": 34})

    write_title_row(ws, 1, "旗舰智能手表市场厂商份额与技术分析", 7)

    # ── 旗舰份额 ──
    write_subtitle_row(ws, 3, "一、旗舰智能手表厂商份额（2024年）", 7)
    write_header_row(ws, 4, [
        "厂商", "出货量份额", "营收份额", "主力芯片/OS", "主力产品", "技术特点", "核心优势"
    ])
    flag = [
        ["苹果Apple",    0.32, 0.45, "S10芯片\nwatchOS",
         "Apple Watch S10\nApple Watch Ultra 2\nSE 2",
         "自研S10 SiP\n双指捏合手势\n睡眠呼吸暂停检测(FDA)\n心电图ECG+血氧\n双层OLED(Ultra)\nIP6X防水",
         "营收份额远超出货\n健康认证最全\niOS生态绑定"],
        ["华为Huawei",   0.18, 0.15, "海思/玄玑OS\nHarmonyOS",
         "WATCH GT 7/Pro 3\nWATCH Ultimate\nFIT 4",
         "玄玑感知健康\n血压/心电/体温\n双向北斗卫星\n14天续航\n蓝宝石+钛金属\n玄玑OS全栈",
         "国内份额领先\n续航+健康双线\n鸿蒙生态联动"],
        ["三星Samsung",  0.12, 0.11, "Exynos W1000\nWearOS",
         "Galaxy Watch7/Ultra\nGalaxy Watch FE",
         "3nm Exynos W1000\nGalaxy AI端侧\n体温/心电/BP(韩国)\n双层OLED(Ultra)\n4nm能效核",
         "安卓阵营标杆\n3nm芯片领先\nGalaxy生态协同"],
        ["小米Xiaomi",   0.08, 0.05, "恒玄/高通\nHyperOS",
         "Watch S4/S3\nRedmi Watch 5",
         "高通W5/恒玄芯片\nHyperOS生态\n血压(实验)/心电\n14-21天续航\n性价比",
         "性价比+销量双优\nHyperOS联动\n印度/欧洲增长"],
        ["谷歌Google",   0.05, 0.05, "Snapdragon W5+\nWearOS",
         "Pixel Watch 3\nPixel Watch 2",
         "高通W5+ Gen1\nFitbit健康算法\nAI助手Gemini\nAMOLED 120Hz\n心电/血氧",
         "Pixel生态+Fitbit\nGemini AI助手\n安卓原生体验"],
        ["Garmin",       0.04, 0.05, "自研\nGarminOS",
         "Fenix 8/Epix\nForerunner 970\nVenu 3",
         "自研多频GPS\n太阳能充电\n急救联动\n48天续航\n运动算法最全",
         "运动/户外标杆\n专业用户忠诚\n高利润高价位"],
        ["Amazfit(华米)",0.06, 0.03, "恒玄/瑞昱\nZepp OS",
         "GTR Mini/GTR 4\nT-Rex Ultra\nActive系列",
         "恒玄/瑞昱芯片\nAI健康评估\nZepp Flow AI助手\n21天续航\n性价比",
         "性价比+海外\n华米供应链\n轻智能路线"],
        ["其他(OPPO/vivo/一加)",0.15,0.06,"高通/恒玄",
         "OPPO Watch X\nvivo Watch GT\nOnePlus Watch 2",
         "高通W5/恒玄\nWearOS/HyperOS\n性价比旗舰",
         "国产二线+性价比\nWearOS/自研OS双线"],
    ]
    fmts = [None, "0.0%", "0.0%", None, None, None, None]
    for i, d in enumerate(flag):
        write_data_row(ws, 5 + i, d, fmts, height=64)

    # ── 旗舰对比 ──
    write_subtitle_row(ws, 14, "二、2025年旗舰智能手表对比", 7)
    write_header_row(ws, 15, [
        "规格", "Watch S10", "Galaxy Watch Ultra", "WATCH GT 7 Pro", "Watch S4", "说明"
    ])
    flag_cmp = [
        ["芯片",    "Apple S10",     "Exynos W1000(3nm)", "海思/玄玑",  "高通W5+",      "苹果/三星自研领先"],
        ["屏幕",    "LTPO OLED",     "双层OLED 3000nit", "AMOLED LTPO", "AMOLED",        "LTPO/双层OLED成旗舰"],
        ["续航",    "18h(36h低功耗)", "100h极限",          "14天",        "21天",          "安卓向长续航靠拢"],
        ["健康",    "ECG/血氧/呼吸暂停","体温/心电/BP(韩)","血压/心电/体温","心电/血氧/血压","医疗化竞赛"],
        ["卫星",    "紧急SOS",       "—",                 "双向北斗",     "—",             "北斗/苹果差异化"],
        ["AI",      "双指捏合+AI",   "Galaxy AI",         "玄玑AI助手",  "HyperOS AI",    "AI助手入腕"],
        ["价格$",  "399-799",       "649+",              "400-700",      "200-300",       "苹果价位最高"],
    ]
    for i, d in enumerate(flag_cmp):
        write_data_row(ws, 16 + i, d, height=32)

    # ── 操作系统生态 ──
    write_subtitle_row(ws, 24, "三、智能手表操作系统生态对比", 7)
    write_header_row(ws, 25, [
        "OS", "厂商", "生态规模", "应用数", "跨手机", "AI助手", "趋势"
    ])
    os_list = [
        ["watchOS",    "苹果",   "iOS封闭",     "10万+", "仅iPhone",   "Siri/Apple Intelligence", "生态最完善"],
        ["WearOS",     "谷歌",   "安卓开放",    "5万+",  "跨安卓",      "Gemini",                   "谷歌+三星联合推动"],
        ["HarmonyOS",  "华为",   "鸿蒙开放",    "万+",   "华为/安卓",   "玄玑AI",                   "国产生态崛起"],
        ["HyperOS",    "小米",   "米家生态",    "千+",   "跨安卓",      "小米大模型",                "小米生态绑定"],
        ["GarminOS",   "Garmin", "运动专业",    "千+",   "跨平台",      "Garmin AI",                "运动专业生态"],
        ["Zepp OS",    "华米",   "轻量自研",    "千+",   "跨平台",      "Zepp Flow AI",              "轻量长续航"],
        ["Tizen(停)", "三星",   "已转WearOS",  "—",     "三星",        "—",                        "2021起转WearOS"],
    ]
    for i, d in enumerate(os_list):
        write_data_row(ws, 26 + i, d, height=32)

    add_pie(ws, "旗舰智能手表出货量份额", [5, 6, 7, 8, 9, 10, 11, 12], 1, 2, "A34", w=16, h=11)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 3: 运动健康手表市场
# ═══════════════════════════════════════════════════════════════
def build_sport_health_sheet(wb):
    ws = wb.create_sheet("运动健康手表")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 14, "D": 16, "E": 22, "F": 34, "G": 30})

    write_title_row(ws, 1, "运动健康手表市场厂商份额与技术分析", 7)

    # ── 运动健康份额 ──
    write_subtitle_row(ws, 3, "一、运动健康手表厂商份额（2024年）", 7)
    write_header_row(ws, 4, [
        "厂商", "出货量份额", "营收份额", "主力芯片/OS", "主力产品", "技术特点", "核心优势"
    ])
    sport = [
        ["Garmin",       0.30, 0.45, "自研SoC\nGarminOS",
         "Fenix 8/Epix Pro\nForerunner 970/265\nVenu 3/Vivomove",
         "多频GPS(L1+L5)\n太阳能充电\nBody Battery体能\n48天续航\n急救联动\n跑步功率计",
         "运动专业绝对领先\n铁三/越野忠诚\n高价位高利润"],
        ["Amazfit(华米)",0.20, 0.10, "恒玄/瑞昱\nZepp OS",
         "GTR 4/Mini\nT-Rex Ultra/Pro\nActive系列",
         "恒玄芯片\nAI健康评估\n21天续航\n150+运动模式\n双频GPS(T-Rex)",
         "性价比+海外量\n华米方案能力\n轻智能路线"],
        ["华为Huawei",   0.18, 0.14, "海思/玄玑OS",
         "WATCH GT 7\nFIT 4/Mini\nWATCH Runner",
         "玄玑感知健康\nTruSeen 6.0心率\n双频五星GPS\n14天续航\n130+运动模式",
         "国内份额领先\n健康+续航双线\n鸿蒙生态"],
        ["博能Polar",    0.08, 0.06, "自研\nPolarOS",
         "Vantage V3\nPacer Pro\nIgnite 3",
         "光学心率精度标杆\n跑步功率(腕式)\nElixia充电\n双频GPS",
         "心率算法标杆\n欧洲运动市场\n专业研究合作"],
        ["Suunto",       0.06, 0.05, "—",
         "Race/Race S\nVertical/Core\n9 Baro",
         "AMOLED/双频GPS\n钛合金外壳\n自由潜/登山专业\n地图离线",
         "芬兰户外传承\n潜水/登山细分"],
        ["COROS高驰",    0.06, 0.04, "—",
         "Apex 2 Pro\nVertix 2s\nPace 3",
         "双频GPS\n触摸+旋钮\n长续航(60h GPS)\n运动功率\n越野/铁三",
         "性价比专业路线\n铁三/越野细分\n口碑驱动"],
        ["小米/Redmi",   0.08, 0.04, "恒玄/瑞昱",
         "手环9 Pro/Watch\nRedmi Watch 5",
         "恒玄/瑞昱方案\n基础运动+健康\n性价比旗舰\n长续航",
         "性价比+量\n入门运动+健康"],
        ["其他(Wahoo/Withings)",0.04,0.12,"—","—","—","细分/医疗化"],
    ]
    fmts = [None, "0.0%", "0.0%", None, None, None, None]
    for i, d in enumerate(sport):
        write_data_row(ws, 5 + i, d, fmts, height=64)

    # ── 运动传感器对比 ──
    write_subtitle_row(ws, 14, "二、核心运动传感器与健康算法对比", 7)
    write_header_row(ws, 15, [
        "传感器/算法", "厂商", "精度", "差异化", "代表产品", "趋势", "壁垒"
    ])
    sensor = [
        ["多频GPS",     "L1+L5双频",  "高",     "城市/峡谷精度提升",     "Garmin/华为/COROS", "成旗舰标配", "中(算法+天线)"],
        ["光学心率",    "PPG多波长",  "中高",   "TruSeen 6.0/Precision Prime", "华为/Polar",        "持续提升精度", "中(算法+LED)"],
        ["心电ECG",     "导联心电",   "高",     "FDA/NMPA认证",            "苹果/华为/三星",    "医疗化标配", "高(认证+算法)"],
        ["血氧SpO2",    "PPG红光",   "中",     "24h连续监测",             "全厂商",            "普及化",      "低"],
        ["血压",        "光学/袖带",  "中",     "无袖带光学血压",          "华为/三星/Withings","医疗化前沿",  "极高(算法+认证)"],
        ["体温监测",    "皮肤+环境",  "中高",   "连续体温+经期预测",       "苹果/华为/三星",    "女性健康切入","中"],
        ["体脂/水分",   "BIA电极",   "中",     "腕部BIA+汗液",            "Withings/Garmin",   "细分前沿",    "高(算法+认证)"],
        ["血糖(无创)",  "光学+AI",    "低-中",  "无创血糖(研发中)",        "苹果/华为(实验)",     "医疗化圣杯",  "极高(精度/认证)"],
    ]
    for i, d in enumerate(sensor):
        write_data_row(ws, 16 + i, d, height=36)

    # ── 运动场景细分 ──
    write_subtitle_row(ws, 25, "三、运动手表场景细分定位", 7)
    write_header_row(ws, 26, [
        "场景", "代表产品", "核心卖点", "续航", "价位$", "主要厂商", "趋势"
    ])
    scene = [
        ["铁人三项", "Forerunner 970", "跑步功率+多频GPS+续航", "23h GPS", "649+", "Garmin", "专业铁三标杆"],
        ["越野/登山","Fenix 8/Epix", "钛合金+地图+急救",      "48天/16h GPS", "899+", "Garmin/Suunto", "户外高利润"],
        ["跑步训练","Forerunner 265", "跑步功率+恢复",         "13天/39h GPS","399+", "Garmin/COROS", "跑步主流量"],
        ["潜水",    "Fenix 8潜水/Descent", "100m+潜水+气瓶",  "48天",      "999+", "Garmin",      "潜水细分"],
        ["健身/生活","Venu 3/Active",  "Body Battery+150运动", "14天",     "299+", "Garmin/Amazfit","生活+健身双线"],
        ["马拉松",   "Pace 3/Apex 2",   "轻量+续航+GPS",        "60h GPS",   "199+", "COROS",       "性价比马拉松"],
    ]
    for i, d in enumerate(scene):
        write_data_row(ws, 27 + i, d, height=32)

    add_pie(ws, "运动健康手表营收份额", [5, 6, 7, 8, 9, 10, 11, 12], 1, 3, "A34", w=16, h=11)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 4: 轻智能与儿童手表
# ═══════════════════════════════════════════════════════════════
def build_light_kids_sheet(wb):
    ws = wb.create_sheet("轻智能与儿童手表")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 14, "D": 16, "E": 22, "F": 34, "G": 30})

    write_title_row(ws, 1, "轻智能手表与儿童手表市场分析", 7)

    # ── 轻智能份额 ──
    write_subtitle_row(ws, 3, "一、轻智能手表厂商份额（2024年）", 7)
    write_header_row(ws, 4, [
        "厂商", "出货量份额", "营收份额", "主力品类", "主力产品", "技术特点", "核心优势"
    ])
    light = [
        ["卡西欧Casio",   0.22, 0.20, "G-Shock/Protrek",
         "G-Shock DW-H5600\nMRG系列\nProtrek",
         "MIP内存屏(超低功耗)\n太阳能充电\n蓝牙连手机\n多频GPS(Protrek)\n轻智能+长续航",
         "腕表+智能融合\n超长续航\n户外潮流品牌"],
        ["佳明Garmin",   0.18, 0.18, "Vivomove/instinct",
         "Vivomove系列\ninstinct 2/3\nLily 2",
         "MIP隐屏\n太阳能充电\n传统指针+智能\n21-48天续航",
         "Garmin运动能力下放\n隐屏+长续航"],
        ["天梭/斯沃奇",   0.15, 0.10, "传统腕表智能款",
         "T-Touch Connect\nSwatch Touch",
         "SwAlpineTouch OS\n瑞士机械+智能\n长续航\n奢华定位",
         "瑞士钟表品牌智能款\n传统+智能融合"],
        ["华为Huawei",   0.12, 0.10, "FIT/Watch GT",
         "FIT 4/Mini\nWATCH GT 7\nWATCH FIT系列",
         "AMOLED屏\n长续航14天\n健康监测\n轻智能路线",
         "国内轻智能领先\n健康+续航双线"],
        ["Amazfit",     0.10, 0.06, "GTR Mini",
         "GTR Mini\nPop系列",
         "AMOLED\n长续航14-21天\n性价比轻智能",
         "性价比+海外\n轻智能路线"],
        ["Withings",    0.06, 0.10, "ScanWatch/Move",
         "ScanWatch Nova/2\nMove/Echo",
         "PMO隐屏\n心率/ECG/SpO2\n医疗认证\n长续航30天",
         "医疗化轻智能\nNMPA/FDA认证"],
        ["其他(华为/小米/Fitbit)",0.17,0.26,"—","—","—","细分+性价比"],
    ]
    fmts = [None, "0.0%", "0.0%", None, None, None, None]
    for i, d in enumerate(light):
        write_data_row(ws, 5 + i, d, fmts, height=56)

    # ── 儿童手表份额 ──
    write_subtitle_row(ws, 13, "二、儿童手表厂商份额（2024年）", 7)
    write_header_row(ws, 14, [
        "厂商", "出货量份额", "营收份额", "主力产品", "操作系统", "技术特点", "核心优势"
    ])
    kids = [
        ["小天才(步步高)",0.55, 0.60, "Z10/Z9\nZ7/Z6A\nZ6巅峰版",
         "自研OS\n(基于RTOS)",
         "视频通话+定位\nAI问答助手\n紧急SOS\n运动健康\n双摄+AI学习\n儿童社交",
         "国内儿童手表绝对第一\n儿童社交生态\n家长信任度高"],
        ["华为Huawei",   0.20, 0.15, "儿童手表 5 Pro\n儿童手表 4",
         "HarmonyOS",
         "5G双卡\nAI学习\n视频通话\n健康监测\n14天续航",
         "国内二线\n鸿蒙生态联动\n运营商渠道"],
        ["360",          0.08, 0.06, "360儿童手表\n8XS/M2",
         "自研OS",
         "视频通话\nAI问答\n定位+SOS\n性价比",
         "性价比+互联网渠道\n儿童安全"],
        ["小米/米兔",    0.07, 0.04, "米兔手表\n小米儿童手表",
         "HyperOS",
         "视频通话\n定位+SOS\n性价比\n米家联动",
         "性价比+米家生态"],
        ["读书郎/小寻", 0.05, 0.03, "读书郎G100\n小寻Y5",
         "自研OS",
         "视频通话+学习\n定位\n教育内容",
         "教育内容差异化\n细分市场"],
        ["其他",         0.05, 0.12, "—", "—", "—", "海外/利基"],
    ]
    for i, d in enumerate(kids):
        write_data_row(ws, 15 + i, d, fmts, height=48)

    # ── 轻智能趋势 ──
    write_subtitle_row(ws, 22, "三、轻智能与儿童手表趋势", 7)
    write_header_row(ws, 23, [
        "趋势", "说明", "代表厂商/产品", "驱动力", "时间线", "影响", "挑战"
    ])
    lk_trend = [
        ["隐屏长续航", "MIP隐屏+太阳能\n30-48天续航",       "卡西欧/Garmin/Withings", "续航焦虑",   "持续",  "中高","屏显技术/功耗"],
        ["传统腕表智能","瑞士机械+智能模块\n传统+智能融合",    "天梭/卡西欧",             "腕表文化复兴","2025-2027","中","成本/品牌调性"],
        ["医疗化轻智能","隐屏+ECG/血压医疗认证\n慢病管理",       "Withings",               "慢病管理",   "2025-2028","中高","医疗认证/算法"],
        ["儿童AI问答",  "AI学习+儿童社交\n紧急救援",            "小天才/华为",             "AI教育",     "2025-2027","中高","内容合规/隐私"],
        ["儿童5G+视频", "5G双卡+视频通话\n高清定位",             "小天才/华为",             "通信升级",   "2024-2027","中",  "资费/功耗"],
    ]
    for i, d in enumerate(lk_trend):
        write_data_row(ws, 24 + i, d, height=44)

    add_pie(ws, "轻智能手表出货量份额", [5, 6, 7, 8, 9, 10, 11], 1, 2, "A31", w=15, h=11)
    add_pie(ws, "儿童手表出货量份额", [15, 16, 17, 18, 19, 20], 1, 2, "I31", w=15, h=11)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 5: 发展趋势与建议
# ═══════════════════════════════════════════════════════════════
def build_trends_sheet(wb):
    ws = wb.create_sheet("发展趋势与建议")
    set_col_widths(ws, {"A": 8, "B": 20, "C": 36, "D": 20, "E": 14, "F": 14, "G": 28})

    write_title_row(ws, 1, "智能手表市场发展趋势与投资建议", 7)

    # ── 核心趋势 ──
    write_subtitle_row(ws, 3, "一、2025-2027年核心趋势", 7)
    write_header_row(ws, 4, [
        "序号", "趋势方向", "详细说明", "受益领域", "确定性", "影响", "关键变量"
    ])
    trends = [
        ["1", "健康监测医疗化",
         "ECG/血压/体温/血糖无创监测\nFDA/NMPA认证升级\n慢病管理方向",
         "旗舰/医疗化细分", "高", "★★★★★",
         "传感器精度/医疗认证"],
        ["2", "自研芯片成护城河",
         "苹果S系列/海思/瑞昱\n芯片+算法+续航一体化\n大厂自研加速",
         "旗舰智能手表", "中高", "★★★★",
         "芯片研发投入/规模"],
        ["3", "AI助手入腕",
         "端侧LLM语音助手\n健康问答+分析\nGemini/玄玬AI/Apple Intelligence",
         "全品类智能手表", "中高", "★★★",
         "端侧算力/AI模型/隐私"],
        ["4", "卫星通信与户外",
         "北斗双向/苹果紧急SOS\nGarmin inReach救援\n户外场景刚需",
         "旗舰/运动手表", "高", "★★★",
         "卫星生态/功耗"],
        ["5", "续航与形态创新",
         "长续航(7-48天)+快充\n轻智能回潮\n隐屏/MIP技术",
         "轻智能/运动手表", "高", "★★★",
         "电池能量密度/功耗优化"],
        ["6", "无感健康生态",
         "手表+耳机+戒指+秤联动\n健康数据中枢\nApple/三星/华为生态",
         "全品类可穿戴", "中高", "★★★",
         "跨设备协同/隐私合规"],
    ]
    for i, d in enumerate(trends):
        write_data_row(ws, 5 + i, d, height=56)

    # ── 产业链价值分布 ──
    write_subtitle_row(ws, 12, "二、智能手表产业链价值分布", 7)
    write_header_row(ws, 13, [
        "环节", "说明", "代表厂商", "价值占比", "壁垒", "国产化", "趋势"
    ])
    chain = [
        ["品牌与整机",   "手表品牌/整机设计",   "苹果/华为/三星/Garmin",  "35-40%", "极高", "低(高端品牌)",
         "国产品牌向上突破"],
        ["芯片/SoC",     "AP/协处理/通信",      "苹果/海思/瑞昱/高通",   "15-20%", "极高", "中(海思/瑞昱)",
         "自研芯片成护城河"],
        ["传感器",       "光学/IMU/生物电极",   "AMS/欧姆龙/华米自研",   "10-15%", "高",   "中",
         "光学传感器价值高"],
        ["显示屏幕",     "AMOLED/MIP隐屏",      "京东方/天马/三星/SDC",   "10-12%", "中高", "中(京东方/天马)",
         "LTPO/双层OLED"],
        ["电池与封装",   "电池/SiP封装/外壳",   "欣旺达/歌尔/立讯",       "8-10%",  "中",   "高",
         "SiP封装成趋势"],
        ["软件与算法",   "OS/健康算法/AI",      "苹果/华为/Garmin",       "8-10%",  "高",   "中",
         "健康算法+AI价值提升"],
    ]
    for i, d in enumerate(chain):
        write_data_row(ws, 14 + i, d, height=36)

    # ── 投资关注方向 ──
    write_subtitle_row(ws, 21, "三、投资关注方向", 7)
    write_header_row(ws, 22, [
        "方向", "逻辑", "标的类型", "时间窗口", "确定性", "风险", "风险提示"
    ])
    invest = [
        ["健康医疗化",
         "ECG/血压/血糖医疗认证\n慢病管理+保险联动",
         "苹果/华为/Withings供应链", "2025-2028", "中高", "中高",
         "传感器精度/认证周期"],
        ["自研芯片",
         "苹果/华为/瑞昱自研\n芯片+算法+续航护城河",
         "苹果/海思/瑞昱", "2025-2027", "中", "中",
         "研发投入/规模经济"],
        ["运动专业赛道",
         "Garmin/COROS高利润高价位\n户外+运动忠诚用户",
         "Garmin/COROS供应链", "2025-2028", "中高", "中",
         "细分天花板/竞争加剧"],
        ["轻智能回潮",
         "隐屏长续航+传统腕表智能\n续航焦虑+腕表文化",
         "卡西欧/Garmin/Withings", "2025-2027", "中", "中",
         "潮流持续性/技术路线"],
        ["儿童AI化",
         "AI学习+儿童社交+救援\n国内儿童手表刚需",
         "小天才/华为供应链", "2025-2027", "中", "中",
         "内容合规/海外受限"],
    ]
    for i, d in enumerate(invest):
        write_data_row(ws, 23 + i, d, height=48)

    # ── 全球厂商综合评分 ──
    write_subtitle_row(ws, 29, "四、全球主要智能手表厂商综合竞争力评分（1-10分）", 7)
    write_header_row(ws, 30, [
        "厂商", "品牌力", "健康医疗", "生态", "AI能力", "市场地位", "综合评分"
    ])
    score = [
        ["苹果Apple",   10, 9, 10, 9,  10, 9.6],
        ["华为Huawei",   8, 9,  8,  7,  8,  8.2],
        ["三星Samsung",  8, 7,  8,  7,  8,  7.4],
        ["Garmin",       8, 8,  7,  6,  6,  7.0],
        ["小米Xiaomi",   7, 6,  7,  6,  8,  6.8],
        ["谷歌Google",   7, 7,  6,  9,  5,  6.8],
        ["Amazfit华米",  6, 6,  5,  5,  7,  5.8],
        ["小天才",        7, 5,  6,  5,  7,  6.0],
        ["卡西欧Casio",   8, 4,  4,  3,  5,  5.0],
        ["Withings",      6, 8,  4,  4,  4,  5.2],
    ]
    for i, d in enumerate(score):
        write_data_row(ws, 31 + i, d, bold_cols={7})
        ws.cell(row=31+i, column=7).font = _font(size=11, bold=True, color=C["deep_purple"])

    add_bar(ws, "全球主要智能手表厂商综合竞争力评分",
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
        ("旗舰智能手表",       build_flagship_sheet),
        ("运动健康手表",       build_sport_health_sheet),
        ("轻智能与儿童手表",   build_light_kids_sheet),
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
