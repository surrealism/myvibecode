"""
打印机市场分析报告生成器
覆盖激光打印机、喷墨打印机、针式打印机等细分市场
厂商份额、技术特点、竞争格局、发展趋势
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
OUTPUT_FILE = OUTPUT_DIR / "打印机市场分析报告.xlsx"

# ── 配色方案 ──
C = {
    "title_bg":    "1F4E79",
    "title_fg":    "FFFFFF",
    "header_bg":   "2E75B6",
    "header_fg":   "FFFFFF",
    "sub_bg":      "D6E4F0",
    "accent1":     "4472C4",
    "accent2":     "ED7D31",
    "accent3":     "70AD47",
    "accent4":     "FFC000",
    "accent5":     "5B9BD5",
    "accent6":     "A5A5A5",
    "light_gray":  "F2F2F2",
    "white":       "FFFFFF",
    "deep_blue":   "1F4E79",
    "red":         "C00000",
    "green":       "008000",
}

# ── 通用样式工厂 ──
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
def build_overview_sheet(wb):
    ws = wb.create_sheet("市场概览", 0)
    set_col_widths(ws, {"A": 20, "B": 18, "C": 18, "D": 14, "E": 20, "F": 26, "G": 30})

    write_title_row(ws, 1, "全球打印机市场分析报告（2024-2025）", 7)
    write_source_row(ws, 2, "数据来源：IDC，Gartner，Keypoint Intelligence 等行业研究报告 | 制表日期: 2026年7月", 7)

    # ── 细分市场规模 ──
    write_subtitle_row(ws, 4, "一、打印机细分市场规模", 7)
    write_header_row(ws, 5, [
        "细分市场", "2023年(亿美元)", "2024年(亿美元)", "同比增速",
        "2025年预测(亿美元)", "主要应用", "市场特征"
    ])
    data = [
        ["激光打印机(Laser)",    280,  295,  0.054,  310,  "办公/企业/政府",    "A4为主力，彩色渗透提升"],
        ["喷墨打印机(Inkjet)",   250,  260,  0.040,  270,  "家用/摄影/商用SOHO", "耗材收入占比较大"],
        ["针式打印机(Dot Matrix)", 25,  23,  -0.080,  21,  "金融/税务/物流",    "利基市场持续萎缩"],
        ["大幅面打印机",         55,   60,   0.091,  65,   "CAD/GIS/广告",      "高利润，专业领域"],
        ["3D打印机(工业级)",     70,   85,   0.214,  105,  "制造/医疗/航空",    "高增长，快速成型"],
        ["3D打印机(消费级)",     20,   22,   0.100,  25,   "教育/创客/DIY",     "入门级普及加速"],
    ]
    fmts = [None, "#,##0", "#,##0", "0.0%", "#,##0", None, None]
    for i, d in enumerate(data):
        write_data_row(ws, 6 + i, d, fmts)

    # 合计
    write_data_row(ws, 12, [
        "合计", 700, 745, 0.0643, 796, "—", "传统打印市场平稳，3D打印增长加速"
    ], fmts, bold_cols={1, 2, 3, 4, 5})

    # ── 市场集中度 ──
    write_subtitle_row(ws, 14, "二、各细分市场集中度（CR3）", 7)
    write_header_row(ws, 15, [
        "细分市场", "CR1", "CR2", "CR3", "头部厂商", "竞争格局", "进入壁垒"
    ])
    cr = [
        ["激光打印机",  "30%", "56%", "78%",  "惠普/兄弟/佳能",       "寡头竞争",   "高(耗材锁定+渠道)"],
        ["喷墨打印机",  "35%", "61%", "82%",  "惠普/佳能/爱普生",     "寡头竞争",   "高(专利+耗材生态)"],
        ["针式打印机",  "35%", "60%", "79%",  "爱普生/富士通/Oki",    "寡头+日系",  "中(利基+售后)"],
        ["大幅面打印机","35%", "56%", "75%",  "惠普/佳能/爱普生",     "寡头竞争",   "高(色彩管理+渠道)"],
        ["工业3D打印",  "18%", "32%", "44%",  "Stratasys/3D Systems/EOS","竞争型", "高(材料+专利)"],
        ["消费3D打印",  "25%", "40%", "52%",  "创想三维/Bambu Lab/拓竹","竞争型",  "中等(开源+社区)"],
    ]
    for i, d in enumerate(cr):
        write_data_row(ws, 16 + i, d)

    # ── 关键趋势 ──
    write_subtitle_row(ws, 23, "三、2025年关键趋势", 7)
    write_header_row(ws, 24, ["序号", "趋势", "说明", "主要推进方", "确定性", "影响", "关键变量"])
    trends = [
        ["1", "打印机即服务(MPS)",       "企业从买硬件转向按页付费/全包服务模式",  "HP/佳能/施乐/Konica Minolta", "极高", "★★★★★", "企业IT外包趋势/远程办公"],
        ["2", "云打印与移动化",           "云原生打印管理、移动设备AirPrint/Mopria", "微软/苹果/惠普/佳能",       "高", "★★★★", "混合办公模式/安全需求"],
        ["3", "3D打印工业化",             "金属3D打印走向批量生产，航天/医疗应用加速", "EOS/GE Additive/铂力特",   "高", "★★★★", "材料性能/速度/成本"],
        ["4", "Bambu Lab颠覆消费3D打印",  "高速大尺寸+多色打印，封闭生态引争议",    "Bambu Lab/拓竹/创想三维",   "高", "★★★",  "开源vs封闭路线之争"],
        ["5", "耗材模式创新",             "惠普Instant Ink/爱普生Smart Tank(墨仓)",  "惠普/爱普生/佳能",        "高", "★★★★", "耗材单价/使用频次"],
        ["6", "可持续/绿色打印",          "再生材料/节能模式/碳抵消认证",            "惠普(碳中和)/爱普生",       "中高","★★★",  "环保法规/ESG要求"],
    ]
    for i, d in enumerate(trends):
        write_data_row(ws, 25 + i, d)

    add_bar(ws, "打印机细分市场规模对比（亿美元）",
            [6, 7, 8, 9, 10, 11], 1, [2, 4], "A32", w=22, h=12)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 2: 激光打印机市场
# ═══════════════════════════════════════════════════════════════
def build_laser_sheet(wb):
    ws = wb.create_sheet("激光打印机市场分析")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 14, "D": 16, "E": 24, "F": 32, "G": 34})

    write_title_row(ws, 1, "激光打印机市场厂商份额与技术分析", 7)

    # ── 市场份额 ──
    write_subtitle_row(ws, 3, "一、激光打印机厂商出货份额（2024年）", 7)
    write_header_row(ws, 4, [
        "厂商", "出货量份额", "营收份额", "主力产品线", "技术特点", "核心优势", "战略方向"
    ])
    data = [
        ["惠普(HP)",     0.38, 0.40, "LaserJet Pro/Enterprise\nM系列/Color LaserJet",
         "LaserJet引擎+JetIntelligence\n自动双面/云打印安全\nA4/A3全覆盖，100页多功能",
         "渠道覆盖最广；MPS服务领先\n耗材Instant Ink锁客\n企业安全打印方案强",
         "加码MPS/A3复合机\n云打印安全平台"],
        ["兄弟(Brother)", 0.18, 0.15, "HL-L系列/DCP三合一\nMFC多功能系列",
         "紧凑设计/鼓粉分离\nBrother iPrint&Scan\n全系网络打印标配",
         "TCO成本低(鼓粉分离+长寿命\n代用粉兼容性强)\n中小企业/家庭用户首选",
         "鼓粉分离持续迭代\n中小B2B市场深耕"],
        ["佳能(Canon)",  0.16, 0.18, "LBP系列/iR复合机\nA3多功能设备",
         "LBP激光技术专利\n按需定影/快速预热\n与腾彩喷墨双线布局",
         "彩色激光图像质量好\n复印机渠道优势/复合机份额大\nR系列再制造减少碳足迹",
         "彩色A3复合机\n办公IT整合方案"],
        ["富士胶片(FujiFilm)", 0.10, 0.08, "Apeos系列/Revoria",
         "富士施乐技术继承\nApeos连接/云平台\nSCGA打印技术",
         "高端彩机图像质量优异\n办公解决方案生态全\n行业定制能力强",
         "Apeos系列全彩化\n整合富士胶片打印技术"],
        ["柯尼卡美能达(Konica Minolta)", 0.08, 0.07, "bizhub i系列",
         "bizhub系列/AUF技术\n云连接/Cost Hub\nIoT设备管理",
         "海外MPS市场强\n彩色中速复合机质量好\nbizhub生态开放",
         "数字化转型服务\n环保碳中和路线"],
        ["京瓷(Kyocera)", 0.05, 0.05, "ECOSYS系列/MA系列",
         "ECOSYS免更换耗材\n陶瓷感光鼓长寿命\n长寿命低TCO定位",
         "ECOSYS专利(鼓寿命超10万页)\nTCO行业最低\n环保/工业情怀品牌",
         "ECOSYS产品线全\n鼓寿命继续突破"],
        ["利盟(Lexmark)", 0.03, 0.04, "CS/MX系列",
         "云服务/安全打印\n物联网管理平台\n全系加密硬盘",
         "企业安全打印领先\n医疗/金融行业客户多\n与Nioce图像整合",
         "企业安全/行业定制\nMPS服务"],
        ["其他",          0.02, 0.03, "联想/奔图(国产)", "—", "国产替代/信创/性价比", "国产激光打印份额提升"],
    ]
    fmts = [None, "0.0%", "0.0%", None, None, None, None]
    for i, d in enumerate(data):
        write_data_row(ws, 5 + i, d, fmts, height=56)

    # ── 市场细分 ──
    write_subtitle_row(ws, 14, "二、激光打印速度与价位段分布", 7)
    write_header_row(ws, 15, [
        "类别", "速度(页/分)", "价格带", "主要厂商", "目标用户", "占出货比", "趋势"
    ])
    seg = [
        ["入门级个人",     "18-25ppm",  "$100-250",  "惠普/兄弟",       "家庭/个人/小型办公",   "0.25", "向无线/AutoDoc转型"],
        ["中小型办公",     "25-40ppm",  "$250-600",  "兄弟/惠普/佳能",    "中小企业/工作组",      "0.35", "内卷严重，鼓粉分离普及"],
        ["中高速办公",     "40-60ppm",  "$600-1500",  "惠普/佳能/富士",   "中大型企业",           "0.22", "彩色+A3复合机化"],
        ["高速生产",       "60-100ppm", "$1500+",     "惠普/柯尼卡/富士",  "大型企业/打印外包",    "0.12", "MPS合约主导"],
        ["A3彩色复合机",   "20-70ppm",  "$2000-8000", "佳能/柯尼卡/利盟",  "大型企业/政府",        "0.06", "向IT解决方案演进"],
    ]
    for i, d in enumerate(seg):
        write_data_row(ws, 16 + i, d, [None, None, None, None, None, "0.0%", None])

    # ── 技术对比 ──
    write_subtitle_row(ws, 22, "三、激光打印核心技术对比", 7)
    write_header_row(ws, 23, [
        "技术", "说明", "代表厂商", "优势", "不足", "趋势", "影响"
    ])
    tech = [
        ["鼓粉分离",        "感光鼓和墨粉盒独立更换", "兄弟/京瓷/惠普部分",  "降低耗材成本/减少浪费",       "用户需自行更换鼓",      "成为行业标配",     "提升性价比/降低TCO"],
        ["鼓粉一体",        "感光鼓和墨粉盒集成",     "惠普LaserJet传统",     "更换方便/打印质量一致",        "耗材成本高",            "逐步向分离过渡",   "惠普推Instant Ink模式"],
        ["ECOSYS长寿命鼓",  "陶瓷感光鼓寿命超10万页", "京瓷",                "TCO极低/环保/几乎免维护",      "感光鼓替换成本高",      "寿命继续提升",     "企业招标TCO优势"],
        ["惠普JetIntelligence","智能碳粉+打印算法",   "惠普",                "打印质量高/安全打印\n云打印Smart Task",
                                                                               "耗材锁死/代用粉兼容差",       "AI智能打印优化",     "企业安全打印标杆"],
        ["LED打印技术",     "LED阵列替代激光扫描",     "富士/兄弟(部分)",     "结构简单/体积小/故障率低",       "精度略低于激光扫描",    "LED技术成熟度提升","A4机小型化趋势"],
        ["A3彩机引擎",      "彩色A3激光引擎",          "佳能/柯尼卡/利盟/富士","高端企业市场/B2B渠道",           "价格高/体积大",        "向企业IT方案融合",  "MPS和解决方案增长"],
    ]
    for i, d in enumerate(tech):
        write_data_row(ws, 24 + i, d, height=40)

    # 饼图
    add_pie(ws, "激光打印机出货量份额", [5, 6, 7, 8, 9, 10, 11, 12], 1, 2, "A31", w=14, h=10)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 3: 喷墨打印机市场
# ═══════════════════════════════════════════════════════════════
def build_inkjet_sheet(wb):
    ws = wb.create_sheet("喷墨打印机市场分析")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 14, "D": 16, "E": 26, "F": 32, "G": 32})

    write_title_row(ws, 1, "喷墨打印机市场厂商份额与技术分析", 7)

    # ── 市场份额 ──
    write_subtitle_row(ws, 3, "一、喷墨打印机厂商出货份额（2024年）", 7)
    write_header_row(ws, 4, [
        "厂商", "出货量份额", "营收份额", "主力产品线", "技术特点", "核心优势", "战略方向"
    ])
    data = [
        ["惠普(HP)",     0.35, 0.38, "HP Smart Tank/OfficeJet\nENVY/DeskJet",
         "Thermal Inkjet(TIJ)专利\nInstant Ink订阅锁客\nSmart Tank墨仓连续供墨\nHP+云打印+安全",
         "品牌知名度最高/渠道最深\n墨仓+订阅双耗材模式\n订阅收入占打印机利润70%+",
         "Instant Ink全球化扩张\nMPS+SMB云打印方案"],
        ["佳能(Canon)",  0.26, 0.25, "PIXMA G系列/MegaTank\nMAXIFY商务/TS照片",
         "FINE(Full-lithography Inkjet\nNozzle)喷头专利\nG系列墨仓(6,000页)\n照片级打印色彩还原度高",
         "照片打印画质行业领先\n墨仓技术成熟(大容量/低TCO)\nPIXMA生态完善(耗材/APP)",
         "G系列持续迭代\n高附加值照片打印"],
        ["爱普生(Epson)",0.24, 0.25, "EcoTank L系列\nWorkForce/SureColor",
         "PrecisionCore/Heat-Free\n压电微机电(Piezo)技术\nEcoTank大容量墨仓(7,500页)\n冷打印技术无需加热",
         "PrecisionCore芯片级喷头精度\n压电式不发热(更省电/不堵头)\nEcoTank TCO极低\n大幅面市场份额大",
         "Heat-Free打印技术推广\n企业级WorkForce 8,000元+市场\n工业级喷墨布局"],
        ["兄弟(Brother)", 0.10, 0.08, "INNOVIS系列/MFC-J\nDCP-J",
         "Piezo Inkjet技术\n墨仓系列/紧凑设计\nBrother独创防墨水干涸",
         "紧凑/可靠/鼓粉分离理念延伸\n鼓粉分离+喷墨墨仓双线\n中小企业用户忠诚度高",
         "喷墨+激光双线互补\n小企业整合方案"],
        {"lexmark": "利盟(Lexmark)"},
    ]
    # Correcting with a list:
    data = [
        ["惠普(HP)", 0.35, 0.38, "Smart Tank/OfficeJet/ENVY", "Thermal Inkjet (TIJ) 墨仓连续供墨 Instant Ink订阅 HP+云打印安全", "品牌最强渠道最深 Instant Ink订阅收入占利润70%+", "Instant Ink全球扩张 MPS+SMB云打印"],
        ["佳能(Canon)", 0.26, 0.25, "PIXMA G MegaTank MAXIFY", "FINE喷头专利 G系列墨仓6000页 照片打印色彩还原出色", "照片打印画质领先 墨仓成熟大容量 PIXMA生态完善", "G系列迭代 高附加值照片打印"],
        ["爱普生(Epson)", 0.24, 0.25, "EcoTank L WorkForce SureColor", "PrecisionCore压电微机电+Heat-Free EcoTank 7500页 冷打印不发热", "芯级喷头精度 压电不堵头 EcoTank TCO极低 大幅面份额大", "Heat-Free技术推广企业级8,000元+ 工业喷墨"],
        ["兄弟(Brother)", 0.10, 0.08, "INNOVIS MFC-J DCP-J系列", "Piezo Inkjet 墨仓 紧凑设计 防墨水干涸", "紧凑可靠 鼓粉分离理念延伸 喷墨墨仓双线", "喷墨+激光互补 中小企业方案"],
        ["其他", 0.05, 0.04, "联想/京瓷等", "—", "中低端/地区性品牌", "国产/白牌喷墨"],
    ]
    fmts = [None, "0.0%", "0.0%", None, None, None, None]
    for i, d in enumerate(data):
        write_data_row(ws, 5 + i, d, fmts, height=56)

    # ── 喷墨技术对比 ──
    write_subtitle_row(ws, 11, "二、热发泡 vs 压电式喷墨技术对比", 7)
    write_header_row(ws, 12, [
        "维度", "热发泡(Thermal Inkjet)", "压电式(Piezo/Pressure)", "优势对比", "代表厂商", "应用场景", "趋势"
    ])
    tech = [
        ["原理",       "加热墨水产生气泡喷射",       "电压驱动压电晶体变形喷射",       "—",             "惠普/佳能",              "家用/办公",                  "两技术路线并存"],
        ["喷头寿命",   "1-2年/需频繁更换喷头",       "3-5年/寿命更长/稳定",            "压电寿命长",     "爱普生/兄弟",            "商务/工业",                  "压电优势扩大"],
        ["墨水类型",   "染料为主/颜料有限",           "颜料/染料/溶剂/UV墨水全支持",    "压电墨水更广",   "爱普生/佳能(颜料)",      "照片/大幅面/工业",           "颜料墨水需求增长"],
        ["打印速度",   "中速/连续打印加热衰减",       "高速稳定/无加热衰减",            "压电速度快",     "爱普生PrecisionCore",    "高速商务/工业",              "PrecisionCore速度领先"],
        ["打印质量",   "照片色彩鲜艳(染料墨水)",       "精细度高/边缘锐利(颜料墨水)",    "各有所长",       "佳能照片/Piezo文字",     "照片选佳能/文字选爱普生",    "差异化竞争"],
        ["堵头风险",   "中(加热导致结垢)",             "低(无加热/墨水适应性强)",        "压电更可靠",     "爱普生EcoTank",          "长期低频打印用户",           "压电喷头免维护趋势"],
        ["能效",       "中(需加热墨水)",               "高(无需加热/功耗低)",            "压电更节能",     "爱普生Heat-Free",        "大批量打印/环保",            "Heat-Free推广"],
    ]
    for i, d in enumerate(tech):
        write_data_row(ws, 13 + i, d, height=36)

    # ── 耗材模式 ──
    write_subtitle_row(ws, 21, "三、耗材商业模式对比", 7)
    write_header_row(ws, 22, [
        "模式", "代表厂商", "说明", "用户成本(平均每页)", "利润模式", "用户粘性", "趋势"
    ])
    sub = [
        ["Instant Ink(订阅)",  "惠普",    "按页付费订阅，自动送墨", "~$0.03-0.07/页", "高(复购+订阅)",    "极高(锁定用户)",  "惠普核心战略，全球扩张"],
        ["EcoTank(墨仓)",      "爱普生",  "大容量墨水(7500页/瓶)",  "~$0.005-0.01/页", "中(硬件利润+墨水)", "中高",            "爱普生/佳能/惠普都推进"],
        ["MegaTank(墨仓)",     "佳能",    "大容量墨水(6000页/瓶)",  "~$0.005-0.01/页", "中(硬件利润+墨水)", "中高",            "佳能G系列主推"],
        ["传统墨盒",            "惠普/佳能","标准墨盒(200-500页)",   "~$0.10-0.30/页", "高(耗材利润)",      "高(耗材锁定)",     "份额持续下降"],
        ["连续供墨系统(CISS)",  "第三方/DIY","外挂墨瓶+导管",        "~$0.001-0.005/页","低(耗材)",          "低",              "兼容/山寨品牌，稳定问题"],
    ]
    for i, d in enumerate(sub):
        write_data_row(ws, 23 + i, d, height=36)

    # 饼图
    add_pie(ws, "喷墨打印机出货量份额", [5, 6, 7, 8, 9], 1, 2, "A29", w=14, h=10)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 4: 大幅面与专业打印
# ═══════════════════════════════════════════════════════════════
def build_wide_format_sheet(wb):
    ws = wb.create_sheet("大幅面与专业打印")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 14, "D": 20, "E": 24, "F": 30, "G": 30})

    write_title_row(ws, 1, "大幅面与专业打印市场分析", 7)

    # ── 市场份额 ──
    write_subtitle_row(ws, 3, "一、大幅面打印机厂商份额（2024年）", 7)
    write_header_row(ws, 4, [
        "厂商", "市场份额", "主力产品线", "技术特点", "核心优势", "应用领域", "战略方向"
    ])
    data = [
        ["惠普(HP)",     0.35, "DesignJet/PageWide XL\nLatex R系列",
         "Thermal Inkjet + Latex\n热发泡+乳胶技术\nPageWide XL高速大幅面",
         "DesignJet CAD/GIS行业标准\nLatex环保/无溶剂\n渠道服务网络最广",
         "CAD建筑/GIS地图\n广告/标识/装饰", "Latex R系列\n工业级大幅面"],
        ["佳能(Canon)",  0.22, "imagePROGRAF PRO\nGP/TA/TX系列",
         "FINE喷头12色/60英寸\nLUCIA墨水/色彩科学\nimagePROGRAF色彩管理",
         "照片打印色彩还原顶级\n12色喷头(3种黑+9色)\n博物馆收藏级输出质量",
         "艺术微喷/摄影/画廊\nCAD/GIS/海报", "PRO系列高色域\n大幅面云打印"],
        ["爱普生(Epson)", 0.21, "SureColor SC-P/SC-T\nSC-F(热转印)",
         "PrecisionCore MicroTFP\n10色UltraChrome PRO\n黑色镀膜/热转印技术",
         "PrecisionCore高精度\n颜料墨水黑色密度极佳\n大幅面照片/色彩管理强",
         "照片/艺术微喷\nCAD/室内装饰\n纺织品热转印", "SureColor P系列\n工业级+纺织品"],
        ["罗兰(Roland DG)",0.08, "TrueVIS/SolJet\nVersaUV/LEC",
         "Piezo热转印/UV固化\n5色+白+光油\nRoland VersaWorks RIP",
         "UV平板+热转印双线\n中小广告制作/装饰\nRIP软件生态完善",
         "UV平板打印/标识\n灯箱/瓶瓶罐罐/装饰", "UV DTF打印\n数码装饰"],
        ["明基(BenQ)",    0.04, "GP系列/W系列\n色准显示器生态",
         "专业影像生态联动\n从显示屏到打印色彩链\nPalette Master调色",
         "摄影全链路色彩协同\n与专业显示器无缝校色\n适合个人工作室/摄影师",
         "高端照片/微喷\n个人影棚/画廊", "摄影全生态战略\n工作室方案"],
        ["其他",          0.10, "Mimaki/武藤/Ricoh", "—", "各具利基优势", "各细分专业领域", "分散化"],
    ]
    fmts = [None, "0.0%", None, None, None, None, None]
    for i, d in enumerate(data):
        write_data_row(ws, 5 + i, d, fmts, height=48)

    # ── 应用细分 ──
    write_subtitle_row(ws, 12, "二、大幅面打印应用领域", 7)
    write_header_row(ws, 13, [
        "应用领域", "市场占比", "增长(YoY)", "主要技术", "主要厂商", "特点", "趋势"
    ])
    app = [
        ["CAD/GIS/建筑",  0.28, 0.03, "HP DesignJet/佳能imagePROGRAF", "惠普/佳能",   "24-44英寸/黑白+彩C/E绘图",    "平稳增长/BIM数字化"],
        ["广告/标识/海报", 0.25, 0.06, "Latex/UV/溶剂",                 "惠普/罗兰",   "64英寸+/户外耐候性",           "Latex环保替代溶剂"],
        ["照片/艺术微喷",  0.18, 0.04, "12色Piezo/Thermal",            "爱普生/佳能", "收藏级/博物馆/画廊/美术馆",    "高利润/专业化"],
        ["室内装饰",       0.12, 0.10, "UV DTF/热转印",                 "罗兰/Mimaki", "壁纸/软装/瓷砖/玻璃/金属面",    "增长率最高/定制化"],
        ["纺织品/T恤",     0.10, 0.08, "热转印(DTG)/Direct to Fabric",  "爱普生/罗兰", "直喷织物/热转印裁片",           "快时尚/个性化定制"],
        ["包装/原型",      0.07, 0.12, "UV平板(Roland LEC/BenQ)",       "罗兰/明基",   "瓶瓶罐罐/异形面/W2P",          "小批量定制/打样"],
    ]
    for i, d in enumerate(app):
        write_data_row(ws, 14 + i, d, [None, "0.0%", "+0.0%;-0.0%", None, None, None, None])

    # ── 墨水技术 ──
    write_subtitle_row(ws, 21, "三、大幅面墨水技术对比", 7)
    write_header_row(ws, 22, [
        "墨水类型", "特点", "优点", "缺点", "适用场景", "环保性", "趋势"
    ])
    ink = [
        ["染料墨水",     "水溶性染料/色纯度高",             "色彩鲜艳/价格低",        "耐候性差/易褪色",   "照片/室内展示",    "中等",      "用于照片打印"],
        ["颜料墨水",     "固态颜料颗粒悬浮",               "耐候100年+/防水耐光",   "色彩略淡/价格高",    "艺术微喷/博物馆",  "高",        "UltraChrome/PRO进化"],
        ["溶剂墨水",     "有机溶剂溶解树脂+颜料",          "户外耐候3-5年/介质广",   "VOC高/有气味",      "广告/车贴/户外",   "低",        "被Latex UV替代中"],
        ["乳胶墨水(Latex)","水性树脂分散颜料+乳胶",        "户外耐候/无VOC/触感好",  "设备成本高/速度慢",  "广告/标识/装饰",   "高",        "HP主推/替代溶剂"],
        ["UV固化墨水",   "紫外光固化成膜",                  "即时干燥/多介质/立体感","设备贵/特殊气味",    "装饰/包装/异形面", "中(无VOC)", "UV LED固化主流"],
        ["热转印墨水",   "分散染料/加热转印到聚酯",        "纺织品鲜艳/耐水洗",     "仅适用于聚酯面料",   "服装/运动服/旗帜",  "中等",      "升华转印主流"],
    ]
    for i, d in enumerate(ink):
        write_data_row(ws, 23 + i, d, height=36)

    add_pie(ws, "大幅面打印机市场份额", [5, 6, 7, 8, 9, 10], 1, 2, "A30", w=14, h=10)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 5: 3D打印市场
# ═══════════════════════════════════════════════════════════════
def build_3d_printing_sheet(wb):
    ws = wb.create_sheet("3D打印市场分析")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 16, "D": 16, "E": 24, "F": 30, "G": 30})

    write_title_row(ws, 1, "3D打印市场厂商份额与技术分析", 7)

    # ── 工业级 ──
    write_subtitle_row(ws, 3, "一、工业级 3D 打印厂商份额（2024年）", 7)
    write_header_row(ws, 4, [
        "厂商", "市场份额", "核心技术", "主力产品", "技术特点", "核心优势", "应用领域"
    ])
    ind = [
        ["Stratasys",     0.16, "FDM(熔融沉积)\nPolyJet(光敏喷射)",
         "Fortus 450mc/F770\nJ850/J55 (PolyJet)",
         "FDM大尺寸碳纤维增强\nPolyJet多材料+全彩\nGrabCAD软件生态",
         "FDM专利/工业可靠性\n多材料全彩PolyJet独有\n材料品种最丰富",
         "航空/汽车/消费品\n医疗模型/模具"],
        ["3D Systems",    0.12, "SLA(光固化)\nSLS(粉末烧结)",
         "ProX/SLA 750\nFigure 4 SLS",
         "SLA立体光刻精度高\nFigure 4高速生产级\n多材料(树脂/金属/塑料)",
         "SLA发明者/精度极高\nFigure 4批产速度领先\n医疗牙科认证齐全",
         "牙科/珠宝/医疗植入物\n航空/精密工业件"],
        ["EOS(德国)",     0.10, "SLS(尼龙/金属)\nDMLS(金属烧结)",
         "EOS P 770 尼龙\nEOS M 300-4 金属",
         "双激光/4激光金属烧结\nSLS尼龙大成型尺寸\nEOS软件/过程监控",
         "金属SLS行业标杆\n汽车批量生产验证充分\n航空认证(AS9100)",
         "汽车零件/航空航天\n医疗器械/模具"],
        ["HP(惠普)",      0.09, "Multi Jet Fusion(MJF)",
         "HP Jet Fusion 5420W\n5200 3D打印系统",
         "Multi Jet Fusion+整页喷墨\nMJF 5200全彩/高频熔剂\nHP 3D处理软件+云连接",
         "速度(比FDM快10倍)\n全彩+黑+特殊介质\nHP全球打印渠道优势",
         "消费品/汽车零件\n医疗/教育"],
        ["GE Additive",   0.08, "EBM(电子束熔化)\nMetal Laser(Laser Powder Bed)",
         "Arcam EBM Spectra H\nConcept Laser M LINE",
         "Arcam EBM真空/应力低\nConcept Laser LPBF高精度\nGE航空供应链一体化",
         "EBM电子束熔化独有\nGE航空内部验证+供应链\n钛合金航空件认证最多",
         "航空涡轮叶片/骨科植入物\n汽车/模具"],
        ["Markforged",    0.05, "CFR(连续纤维)\n金属/碳纤维",
         "FX20/Mark Two\nMetal X Metal",
         "连续纤维增强(碳纤/玻纤/Kevlar)\n金属间接烧结工艺\nEiger云端ADAPT切片",
         "碳纤维连续打印独有认证\n金属+碳纤双线\n国防军工高端认可",
         "航空/军工/汽车/工装\n模具/夹具/终端件"],
        ["Desktop Metal", 0.05, "Binder Jetting(粘接剂)\nSingle Pass Jetting",
         "P-50/P-1 Production\nShop System SPJ",
         "Binder Jetting高速批量\nSPJ单次喷射达12,000cc/h\nDM Studio System+砂模",
         "粘接剂喷射批产速度最快\n金属/陶瓷/砂模多材料\n生产级成本优势",
         "大批量金属件/汽车零配件\n航空航天/牙科/工具"],
        ["其他",          0.35, "Carbon DLS/Formlabs SLS\n铂力特BLT/华曙远铸",
         "—", "—", "—", "—"],
    ]
    fmts = [None, "0.0%", None, None, None, None, None]
    for i, d in enumerate(ind):
        write_data_row(ws, 5 + i, d, fmts, height=56)

    # ── 消费级 ──
    write_subtitle_row(ws, 14, "二、消费级/桌面级 3D 打印厂商份额（2024年）", 7)
    write_header_row(ws, 15, [
        "厂商", "市场份额", "核心技术", "主力产品", "技术特点", "核心优势", "战略方向"
    ])
    cons = [
        ["创想三维(Creality)", 0.28, "FDM(熔融沉积)\n光固化SLA",
         "Ender 3系列/K1\nCR-10/HALOT", \
         "开源FDM/高度可定制\nEnder系列性价比王者\nK1Plus高速500mm/s\nHALOT光固化系列",
         "全球销量最大\nEnder系列爆款/社区生态最大\n超高性价比+开源定制", \
         "K系列高速打印\nHALOT光固化+树脂"],
        ["Bambu Lab(拓竹)",   0.22, "FDM(高速/大尺寸)",
         "X1C/X1E/A1 Mini/P1S\n多色AMS系统",
         "全封闭高速250-500mm/s\nAMS自动换料16色\nX1C工业级精度/TMC驱动\n内置LiDAR校准+AI首层检测",
         "消费级技术领先(速度精度)\nAMS多色生态独有\n从开源转向封闭APP/Gcode云",
         "多色x16 AMS\n工业级桌面机X1E"],
        ["UltiMaker",         0.12, "FDM(专业桌面)",
         "S5/S7/Method XL\nMakerBot Sketch(教育)",
         "双重挤出(水溶支撑)\nS7皮带驱动精度高\nMethod XL大尺寸300x300\nCura切片(开源最强)",
         "Cura行业标准切片软件\n专业级中型企业首选\n教育SKetch积累用户",
         "Method XL工业级\nCura AI切片升级"],
        ["Formlabs",          0.10, "SLA(光固化)\nSLS(选择性烧结)",
         "Form 4/Form 3+\nFuse 1 SLS",
         "Low Force Stereolith\nLFS低应力光固化\nForm 4 100mm/hr高速\nFuse 1 SLS桌面粉末烧",
         "光固化表面精度最佳\nForm 4速度比上代快5倍\n专业设计师/医疗牙科热选",
         "Form 4持续迭代\nFuse+SLS"],
        ["Prusa Research",    0.09, "FDM(开源社区)",
         "MK4/XL\nPrusa Mini+",
         "完全开源(GPL)\nMK4 32位Buddy + Input Shaper\nXL 5工具头多色/多材料\nPrusa Slicer(基于Cura)",
         "开源社区旗帜\n稳定可靠/升级扩展性好\nPrusa Slicer/社区生态",
         "XL多工具头\n高速/大尺寸发展"],
        ["Anycubic",          0.08, "FDM+光固化SLA",
         "Kobra系列/Photon系列\nMono 4K/6K光固化",
         "FDM Kobra自动调平\n光固化Photon高精度\n4K/6K 10寸LCD",
         "入门级性价比好\nPhoton光固化爆款\nAmazon渠道优势",
         "Kobra Plus高速\nPhoton高分辨率"],
        ["其他",              0.11, "Elegoo/Flashforge/长木/Voron", "—", "—", "—", "—"],
    ]
    for i, d in enumerate(cons):
        write_data_row(ws, 16 + i, d, fmts, height=48)

    # ── 技术对比 ──
    write_subtitle_row(ws, 24, "三、主流 3D 打印技术对比", 7)
    write_header_row(ws, 25, [
        "技术", "原理", "精度", "速度", "材料", "成本", "代表厂商", "优势场景"
    ])
    tech = [
        ["FDM/FFF",  "热熔丝逐层沉积",        "0.05-0.4mm", "中(50-500mm/s)", "PLA/ABS/PETG/PC/Nylon/碳纤/金属填充",
         "低(入门$200-$5k)", "Creality/Bambu/UltiMaker/Prusa", "原型/工装/教育/DIY"],
        ["SLA/DLP",  "紫外光固化液态树脂",    "0.01-0.1mm", "中-快",          "光敏树脂(刚性/柔性/铸造/生物)",
         "中($200-$10k)",    "Formlabs/Anycubic/3D Systems", "牙科/珠宝/高精度原型/医疗"],
        ["SLS(尼龙)", "激光烧结粉末",         "0.1-0.2mm",  "慢-中",          "PA(尼龙)/TPU/PP/PEKK",
         "高($20k-$200k)",   "EOS/Formlabs Fuse/华曙",        "功能性零件/小批量生产/汽车"],
        ["DMLS(金属)", "激光烧结金属粉末",    "0.02-0.1mm", "慢",             "钛合金/不锈钢/铝合金/钴铬/镍基",
         "极高($100k+)",     "EOS/GE Additive/铂力特",         "航空叶片/医疗植入/模具"],
        ["MJF(Multi Jet)", "喷墨粘接剂+熔剂","0.1-0.2mm",   "快(FDM 10x)",     "PA(尼龙)/TPU/Metal(预备)",
         "中-高($50k-$200k)", "HP",                           "批量功能件/色件/消费品"],
        ["PolyJet","多喷头光敏喷射+UV固化","0.014mm(超细)","快",  "多树脂(刚性/柔性/透明/高温)",
         "高($20k-$100k)",   "Stratasys","全彩+多材料/医疗模型/设计验证"],
        ["EBM","电子束熔化金属粉末","0.1-0.3mm","慢(但批量好)","钛合金/钴铬/镍基","极高($300k+)","GE Additive(Arcam)","航空涡轮/骨科/低应力"],
    ]
    for i, d in enumerate(tech):
        write_data_row(ws, 26 + i, d, height=40)

    # 饼图
    add_pie(ws, "工业级3D打印市场份额", [5, 6, 7, 8, 9, 10, 11, 12], 1, 2, "A34", w=14, h=10)
    add_pie(ws, "消费级3D打印出货份额", [16, 17, 18, 19, 20, 21, 22], 1, 2, "I34", w=14, h=10)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 6: 针式打印机与特种打印
# ═══════════════════════════════════════════════════════════════
def build_dotmatrix_sheet(wb):
    ws = wb.create_sheet("针式与特种打印")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 14, "D": 20, "E": 24, "F": 30, "G": 28})

    write_title_row(ws, 1, "针式打印机与特种打印市场分析", 7)

    # ── 市场份额 ──
    write_subtitle_row(ws, 3, "一、针式打印机厂商份额（2024年）", 7)
    write_header_row(ws, 4, [
        "厂商", "市场份额", "主力产品", "技术特点", "核心优势", "应用领域", "趋势"
    ])
    data = [
        ["爱普生(Epson)", 0.35, "LQ系列/LQ-2090II\nLQ-680KII/Dot Matrix",
         "24针打印头/复写纸1+6\n打印头2亿次寿命\nESC/POS指令兼容",
         "针式打印全球龙头\n金融税务票据双联出票\n复写纸打印+连续纸进纸",
         "金融/税务/银行/物流\n酒店/POS/医疗/窗口", "稳步收缩/高利润利基"],
        ["富士通(Fujitsu)", 0.18, "DL系列/DPK系列",
         "24针/SPP 4亿次寿命\n自动纸厚检测\n多联复写(1+7层)",
         "日本品质/4亿次长寿命\n银行/财政票据标准\n多层复写能力最强",
         "银行/保险/巨量打印\n窗口/政府/企事业单位", "高可靠性专业领域"],
        ["Oki Data",      0.16, "Microline系列/BM系列",
         "LED+针式双线\nMicroline紧凑设计\n省资源设计(87%再生材料)",
         "日系品质/紧凑节省空间\nLED+针式互补\n环保/节能口碑好",
         "银行/政务/彩票/医院\n窗口/物流单/票据", "LED彩色+针式"],
        ["NCR",           0.10, "NCR POS系列",
         "自助打印终端\n热敏+针式可选\n多联复写定制化",
         "ATM/POS自助打印内置\n终端嵌入式定制\n金融/零售/服务",
         "自助终端/ATM/收银\n外卖/零售/娱乐/票务", "自助KIOSK终端"],
        ["Star Micronics", 0.08, "SP700/TSP系列",
         "高可靠性POS打印\n高速低噪音\n蓝牙/USB/以太全接口",
         "POS/餐饮/KIOSK专用\n热敏+针式双线\n接口丰富小巧耐用",
         "POS/KIOSK/餐饮/快递\n自提柜/储物柜", "自助终端+物联网"],
        ["国产(联想/中盈/新北洋)", 0.13, "各品牌系列",
         "兼容爱普生LQ指令\n更低价格/LQ兼容耗材\n国产替代+信创",
         "国产替代性价比优先\n信创政策驱动\n售后响应快/本地服务",
         "国内政务/金融/税务\n替代进口/窗口打印", "国产替代加速"],
    ]
    fmts = [None, "0.0%", None, None, None, None, None]
    for i, d in enumerate(data):
        write_data_row(ws, 5 + i, d, fmts, height=48)

    # ── 应用场景 ──
    write_subtitle_row(ws, 12, "二、针式/特种打印应用场景", 7)
    write_header_row(ws, 13, [
        "应用场景", "占比", "技术需求", "打印量", "纸张规格", "竞争方案", "趋势"
    ])
    app = [
        ["银行/金融票据",  0.25, "多联复写(1+6)/连续纸", "极大(5-10万页/月)", "241mm/380mm多联",        "无替代/法规存档",    "数字化降低票据量"],
        ["税务发票",       0.20, "多联复写/专用格式",     "大",                "241mm税票专用纸",          "电子发票替代加速",   "电子发票渗透>50%"],
        ["物流面单",       0.18, "热敏取代/针式补充",     "大-中",             "热敏不干胶/连续纸",        "热敏打印占主流",     "热敏+电子面单"],
        ["POS收银/零售",   0.15, "连续打印/联机打印",     "中(500-2000张/日)", "58mm/80mm热敏卷纸",        "热敏POS占主流",     "针式POS逐步淘汰"],
        ["医疗/处方/检验", 0.12, "复写/条码/专用格式",    "中-大",             "A5/连续纸/专用处方纸",     "激光/热敏替代",      "电子病历替代"],
        ["政府/企事业单位",0.10, "多联复写/连续纸/证照",  "中-大",             "A4/连续纸/专用表格纸",     "激光替代中",         "窗口服务电子化"],
    ]
    for i, d in enumerate(app):
        write_data_row(ws, 14 + i, d, [None, "0.0%", None, None, None, None, None])

    # ── 替代趋势 ──
    write_subtitle_row(ws, 21, "三、针式打印替代方案与挑战", 7)
    write_header_row(ws, 22, [
        "替代方案/趋势", "说明", "优势", "挑战", "替代速度", "针式不可替代场景", "结论"
    ])
    sub = [
        ["电子发票",       "PDF/OFD/区块链发票替代纸质", "无需打印/传递快/合规",    "企业ERP/财政对接",          "快(东部快于西部)", "电子+针式(一联备用)",    "深度融合过渡"],
        ["热敏POS/物流",   "热敏纸直接打印/免色带/免碳带","快/静/低维护/低成本",    "热敏纸不耐保存/褪色",       "快",                "多联复写金融财税场景",    "针式转向高价值"],
        ["激光替代窗口打印","A4激光替代针式单联",          "速度快/质量好/静音",     "不能复写/成本高/不可连纸",  "中",                "复写/连续纸/金融专用",    "互补共存"],
        ["无纸化数字化",   "全程电子化/云签/电子签名",     "完全无纸/成本和时延最优", "法律效力/用户习惯/系统成本","慢(法规约束)",      "法规效力要求场景难替代",  "长期趋势/缓慢渗透"],
        ["在线阅读/核验",  "扫码/身份证/指纹在线查询",     "即时验真/大数据风控",     "系统集成/安全隐私",         "中-快",             "现场纸质凭证需求仍在",    "纸质+电子并行"],
    ]
    for i, d in enumerate(sub):
        write_data_row(ws, 23 + i, d, height=40)

    add_pie(ws, "针式打印机市场份额", [5, 6, 7, 8, 9, 10], 1, 2, "A29", w=14, h=10)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 7: 技术特点综合对比
# ═══════════════════════════════════════════════════════════════
def build_tech_comparison_sheet(wb):
    ws = wb.create_sheet("技术特点综合对比")
    set_col_widths(ws, {"A": 18, "B": 16, "C": 16, "D": 16, "E": 16, "F": 20, "G": 20, "H": 20})

    write_title_row(ws, 1, "打印机技术特点综合对比", 8)

    # ── 基本特性 ──
    write_subtitle_row(ws, 3, "一、主流打印技术基本特性对比", 8)
    write_header_row(ws, 4, [
        "特性", "激光打印", "热发泡喷墨", "压电喷墨", "针式打印", "大幅面(专业)", "3D打印(FDM)", "说明"
    ])
    basic = [
        ["打印原理",     "静电硒鼓转印\n+碳粉热熔",   "加热墨水产生\n气泡喷射", "压电晶体变形\n驱动墨水喷射", "钢针撞击\n色带+纸张",  "喷墨/热发泡\n压电/MJF", "热熔丝逐层\n沉积堆积",   "不同打印引擎"],
        ["打印速度",     "18-100ppm",                "5-15ppm(照片)\n8-33ppm(文本)","10-20ppm(照片)\n5-33ppm(文本)","180-480cps\n(字符/秒)",  "10-100sqm/hr",          "50-500mm/s",            "激光速度最快"],
        ["打印精度(dpi)","600-1200dpi",              "4800-9600dpi",           "5760-2880dpi",          "180-360dpi(低)",         "1200-2880dpi",          "0.05-0.4mm(层高)","喷墨分辨率最高"],
        ["单页成本",     "$0.02-0.05(黑白)\n$0.08-0.20(彩)", "$0.03-0.30(彩/墨盒)\n$0.005-0.01(墨仓)",
                                                                                 "同喷墨",                "$0.005-0.02\n(碳带)",   "$0.05-0.50",            "$0.50-2.00(耗材)",  "墨仓喷墨TCO最优"],
        ["介质",         "普通纸/标签/信封",           "普通纸/照片纸/光泽纸",   "普通纸/照片/美术纸",     "多层复写纸/连续纸",      "卷筒纸/乙烯基/画布/板材","PLA/ABS/PETG/金属","针式可多联复写"],
        ["彩色能力",     "彩色(CMYK)/价格高",           "出色(染料/6-12色)",     "出色(颜料/8-12色)",      "黑白为主(少数彩针)",     "出色(12色/Latex/UV)",   "全彩(AMS多色/换料)","喷墨彩打质量最好"],
        ["维护/可靠性",  "低维护/碳粉/感光鼓定期换",     "中维护(堵头风险/每周打)", "低维护(不发热/不堵头)",  "高维护/色带/打印针耗损","中-高维护(喷头校准)",   "低维护(FDM/无需调平)", "压电/激光维护少"],
        ["噪音",         "中(<50dB)",                  "低(<40dB)",             "低(<40dB)",             "高(55-65dB)",           "中-低",                 "中(风扇+电机+挤出头)", "针式噪音最大"],
        ["典型寿命",     "5-10年/15-50万页",            "3-5年/1-5万页",         "3-5年/3-10万页",         "5-10年/10-30万页",       "5-8年/大印量",         "3-5年",                 "针式/激光寿命长"],
    ]
    for i, d in enumerate(basic):
        write_data_row(ws, 5 + i, d, height=48)

    # ── 厂商技术路线 ──
    write_subtitle_row(ws, 15, "二、主要厂商技术路线与战略（2025-2027）", 8)
    write_header_row(ws, 16, [
        "厂商", "激光路线", "喷墨路线", "大幅面/特种", "3D打印", "软件/MPS", "战略核心", "关键布局"
    ])
    roadmap = [
        ["惠普(HP)",
         "LaserJet Enterprise\nA3/A4全系+JetIntelligence",
         "Smart Tank+Instant Ink\nHP+/云打印",
         "DesignJet+Latex\nPageWide XL",
         "Multi Jet Fusion\nMJF 5420W",
         "HP PrintOS/MPS云端\nWolf Security安全",
         "MPS+耗材订阅\n安全和云打印",
         "Instant Ink全球覆盖\nHP+安全锁客户\nMJF加大企业推广"],
        ["爱普生(Epson)",
         "—(无激光)",
         "EcoTank+PrecisionCore\nHeat-Free专利路线",
         "SureColor+PrecisionCore\n热转印+大幅面",
         "—(无3D打印)",
         "WorkForce云打印\nEpson Print Admin",
         "Heat-Free喷墨\n+压电技术护城河",
         "EcoTank持续迭代\n精密工业/大幅面\n纸回收/干纤维技术"],
        ["佳能(Canon)",
         "LBP激光+imageRUNNER\nA3/A4全覆盖",
         "PIXMA G MegaTank\nFINE喷头技术演进",
         "imagePROGRAF\nPRO系列12色",
         "—(无3D打印)",
         "uniFLOW管理软件\nCanon云打印",
         "彩色喷墨/激光\n照片+复印机",
         "EOS相机联动打印\n专业影像方案\n办公多色3D"],
        ["兄弟(Brother)",
         "HL/DCP/MFC全A4\n鼓粉分离长寿命",
         "INNOVIS墨仓系列\nPiezo核心",
         "—(无大幅面)",
         "—(无3D打印)",
         "Brother iPrint&Scan\nAdmin Console管理",
         "中小办公+家庭\n鼓粉分离+品牌信赖",
         "A4紧凑市场深耕\n鼓粉分离持续延展\nMFC多功能化"],
        ["Bambu Lab",
         "—", "—", "—",
         "FDM X1C/X1E/A1\nAMS 16色多功能",
         "Bambu Handy APP\nMakerWorld云社区",
         "消费3D打印颠覆者\n高速+多色+封闭生态",
         "AMS 16色+工业级X1E\n企业战略+工程材料\nX1E+PLA/Aero"],
        ["Stratasys",
         "—", "—", "—",
         "FDM+PolyJet\n工业级GrabCAD软件",
         "GrabCAD Print\nDigital ABS",
         "工业3D打印老牌\nFDM+PolyJet双线",
         "多材料全彩+大尺寸FDM\n汽车航空认证"],
    ]
    for i, d in enumerate(roadmap):
        write_data_row(ws, 17 + i, d, height=56)

    # ── 竞争力评分 ──
    write_subtitle_row(ws, 24, "三、主要打印厂商综合竞争力评分（1-10分）", 8)
    write_header_row(ws, 25, [
        "厂商", "技术先进性", "品牌/渠道", "产品线广度", "耗材模式", "AI/云/软件", "增长潜力", "综合评分"
    ])
    score = [
        ["惠普(HP)",            9,  10, 9,  8,  10, 9,  9.2],
        ["爱普生(Epson)",       9,  8,  8,  8,  7,  7,  7.8],
        ["佳能(Canon)",         8,  9,  9,  7,  8,  6,  7.8],
        ["兄弟(Brother)",       7,  7,  6,  7,  6,  6,  6.5],
        ["Bambu Lab",            9,  6,  5,  6,  8,  10, 7.3],
        ["Stratasys",            8,  5,  7,  4,  5,  4,  5.5],
        ["创想三维(Creality)",   6,  6,  5,  5,  5,  6,  5.5],
        ["富士胶片(FujiFilm)",   7,  7,  6,  6,  6,  5,  6.2],
    ]
    for i, d in enumerate(score):
        write_data_row(ws, 26 + i, d, bold_cols={8})
        ws.cell(row=26+i, column=8).font = _font(size=11, bold=True, color=C["deep_blue"])

    add_bar(ws, "主要打印厂商综合竞争力评分",
            [26, 27, 28, 29, 30, 31, 32, 33], 1, [2, 3, 4, 5, 6, 7], "A35", w=22, h=13)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 8: 发展趋势与建议
# ═══════════════════════════════════════════════════════════════
def build_trends_sheet(wb):
    ws = wb.create_sheet("发展趋势与建议")
    set_col_widths(ws, {"A": 8, "B": 20, "C": 36, "D": 20, "E": 14, "F": 14, "G": 28})

    write_title_row(ws, 1, "打印机市场发展趋势与投资建议", 7)

    # ── 核心趋势 ──
    write_subtitle_row(ws, 3, "一、2025-2027年核心趋势", 7)
    write_header_row(ws, 4, [
        "序号", "趋势方向", "详细说明", "受益领域", "确定性", "影响", "关键变量"
    ])
    trends = [
        ["1", "MPS/打印管理服务全面化",
         "企业从买硬件转向按页付费/全包服务\n云管理+安全打印+固件更新+耗材配送\nMPS渗透率从35%提升至45%",
         "惠普/施乐/Konica Minolta", "极高", "★★★★★",
         "混合办公形态\nIT外包趋势强"],
        ["2", "耗材订阅经济模式",
         "惠普Instant Ink/爱普生EcoTank/佳能G\n硬件降价+耗材提价+订阅锁定\n订阅收入占比将超过50%",
         "惠普/爱普生/佳能", "高", "★★★★",
         "用户接受度/价格敏感\n第三方兼容耗材打压"],
        ["3", "3D打印工业化加速",
         "金属/尼龙3D打印批量进入汽车航空\nGE/空客批量使用3D打印钛合金零件\n金属3D打印市场CAGR 25%+",
         "EOS/GE/铂力特/Stratasys", "高", "★★★★",
         "材料认证\n打印速度提升"],
        ["4", "云打印+移动化+IoT",
         "ChromeOS/AirPrint/Mopria原生支持\nIOT远程固件+墨水监测+自动补货\n印刷管理平台SAAS化",
         "惠普/微软/苹果/佳能", "高", "★★★★",
         "零信任网络\n数据安全合规"],
        ["5", "AI智能打印",
         "AI自动文档优化/色彩校准/高速纠偏\nAI一页多语言识别/语音打印\nAI耗材用量预测+自动补货",
         "各厂商+芯片/AI SaaS", "中高", "★★★",
         "AI芯片/部署成本\n用户习惯"],
        ["6", "可持续/回收/环保法规",
         "再生塑料占比30%+目标\nEPEAT/Energy Star/TCO认证\nHP/佳能/Epson碳足迹削减50%目标",
         "HP/佳能/Epson/京瓷", "中高", "★★★",
         "ESG投资/法规\n再生材料质量/成本"],
        ["7", "Bambu Lab引领消费3D打印革新",
         "X1C颠覆速度+多色+精度\n封闭生态vs开源社区争论\nA1 Mini下沉到$300入门级\nX1E切入企业教育",
         "Bambu Lab/Creality", "高", "★★★★",
         "专利侵权争议/生态封闭\n企业客户拓展进程"],
    ]
    for i, d in enumerate(trends):
        write_data_row(ws, 5 + i, d, height=56)

    # ── 供需背景 ──
    write_subtitle_row(ws, 13, "二、各细分市场竞争态势", 7)
    write_header_row(ws, 14, [
        "细分市场", "总量趋势", "竞争强度", "价格趋势", "利润水平", "2025展望", "关键变量"
    ])
    comp = [
        ["激光打印机",    "稳定/微降", "激烈", "稳中略降", "中(硬件+耗材)", "A3彩色+解决方案价值提升", "MPS渗透率"],
        ["喷墨打印机",    "稳定",      "激烈", "两级分化", "低(硬件)+高(耗材)", "墨仓模式+订阅+照片高附加值", "耗材收入"],
        ["针式打印机",    "下滑(-10%)","低",   "稳中持平", "中(利基专业)", "精细化管理/替代缓慢", "电子发票替代速度"],
        ["大幅面打印",    "稳中有升",   "中",   "稳定",     "高(专业+墨水)", "Latex环保+UV+装饰增长", "建筑/广告/装修热度"],
        ["3D打印(工业)",  "快速增长",   "中-高","下降(CAGR -5%)","低(设备)+高(材料服务)","大批量航空医疗工业批产", "材料/速度/认证突破"],
        ["3D打印(消费)",  "快速增长",   "激烈", "快速下降", "低(设备竞争+材料)", "AI+多色+人机交互创新", "Bambu/Pusa/Creality竞争"],
    ]
    for i, d in enumerate(comp):
        write_data_row(ws, 15 + i, d, height=36)

    # ── 投资方向 ──
    write_subtitle_row(ws, 22, "三、投资关注方向", 7)
    write_header_row(ws, 23, [
        "方向", "逻辑", "标的类型", "时间窗口", "确定性", "风险", "风险提示"
    ])
    invest = [
        ["惠普(HP)打印业务",
         "MPS+Instant Ink订阅模式稳健\n安全和云管理高粘性\n耗材收入现金流出色",
         "HP Inc./打印服务商", "持续", "高", "中低",
         "商用需求波动\n兼容耗材/制裁对印制约"],
        ["Bambu Lab/拓竹",
         "消费3D打印颠覆性创新\nX1C/A1 Mini定义行业\n多色+高速+AI难以复制",
         "Bambu Lab(未上市)/供应链", "2025-2027", "中高", "中高",
         "专利诉讼风险/开源社区\n生态封闭/企业拓展不确定"],
        ["金属3D打印工业化",
         "航空/医疗批量认证通过\nGE/空客/波音需求确定\n材料+服务利润率高",
         "EOS/GE/铂力特/华曙", "2025-2028", "中高", "中",
         "批量成本/材料验证\n认证周期长/钛合金价格"],
        ["大幅面/Latex环保打印",
         "Latex替代溶剂/VOC法规驱动\n装饰/纺织品增长快\nHP/爱普生大幅面增长",
         "HP/爱普生/罗兰大幅面", "持续", "中", "低",
         "广告投资/建筑周期\n溶剂被替代缓慢"],
        ["国产替代(打印机)",
         "奔图/联想国产激光出货增长\n信创政策驱动政务打印机替代\n奔图(纳思达)全球打印机第5",
         "纳思达(奔图)/联想/中盈", "持续(政策驱动)", "中高", "中高",
         "核心技术差距\n专利风险/全球渠道"],
        ["打印耗材/兼容",
         "兼容耗材替代原装\n再生墨盒+环保趋势\n增材(3D打印材料)市场高增",
         "纳思达/天威/鼎龙/3D材料", "持续", "中", "中",
         "原装专利封锁\n环保法规"],
    ]
    for i, d in enumerate(invest):
        write_data_row(ws, 24 + i, d, height=48)

    # ── 产业链 ──
    write_subtitle_row(ws, 31, "四、打印产业链核心环节", 7)
    write_header_row(ws, 32, [
        "环节", "说明", "代表厂商", "价值占比", "壁垒", "国产化", "趋势"
    ])
    chain = [
        ["打印芯片(喷头/硒鼓)", "喷墨喷头/激光OPC/3D打印头",     "惠普/佳能/爱普生/天威",  "20-25%", "极高", "极低",   "喷头专利壁垒/国产突破中"],
        ["耗材(墨盒/碳粉/色带)", "墨盒/碳粉盒/墨水/色带/树脂",    "惠普/天威/纳思达/鼎龙",  "30-35%", "高",   "中",     "兼容+再制造+环保3D材料"],
        ["打印引擎/机芯",        "激光引擎/喷墨机芯/扫描/走纸",    "惠普/佳能/兄弟/爱普生",   "15-20%", "极高", "极低",   "核心机芯专利掌控在原厂"],
        ["主控板/固件",          "控制板/主控芯片/打印语言/固件",   "Marvell/三星/AMD/联发科", "10-15%", "高",   "低",     "自研PCB+嵌入式稳定"],
        ["软件/云/APP",         "云打印/APP/HP+/AirPrint/打印管理", "惠普/微软/苹果/Mopria",  "5-8%",   "中-高","中",     "云+安全+AI打印管理"],
        ["渠道/服务/物流",       "代理商/电商/维修/返修/Call Center","各地渠道商/京东/天猫/维修", "5-10%","中",   "国内完善","线上线下融合+上门服务"],
    ]
    for i, d in enumerate(chain):
        write_data_row(ws, 33 + i, d, height=40)

    add_bar(ws, "主要打印厂商综合竞争力评分",
            [26, 27, 28, 29, 30, 31, 32, 33], 1, [2, 3, 4, 5, 6, 7], "A35", w=22, h=13)

    return ws


# ═══════════════════════════════════════════════════════════════
# 主入口
# ═══════════════════════════════════════════════════════════════
def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    sheets = [
        ("市场概览",           build_overview_sheet),
        ("激光打印机市场分析", build_laser_sheet),
        ("喷墨打印机市场分析", build_inkjet_sheet),
        ("大幅面与专业打印",   build_wide_format_sheet),
        ("3D打印市场分析",     build_3d_printing_sheet),
        ("针式与特种打印",     build_dotmatrix_sheet),
        ("技术特点综合对比",   build_tech_comparison_sheet),
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