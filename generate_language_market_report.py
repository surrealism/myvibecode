"""
编程语言市场份额对比报告生成器

数据来源：
- TIOBE Index (2025.05)
- Stack Overflow Developer Survey 2024
- GitHub Octoverse 2024
"""

import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.utils import get_column_letter

# ─── 数据 ───────────────────────────────────────────────────────────

# TIOBE 指数 (2025年5月近似值)
tiobe_data = [
    ("Python",      28.11),
    ("C",           9.70),
    ("C++",         9.45),
    ("Java",        8.95),
    ("C#",          6.80),
    ("JavaScript",  4.20),
    ("Go",          2.85),
    ("SQL",         2.40),
    ("Fortran",     2.10),
    ("Rust",        1.95),
    ("Swift",       1.70),
    ("Delphi",      1.55),
    ("PHP",         1.40),
    ("Ruby",        1.20),
    ("MATLAB",      1.05),
    ("R",           0.95),
    ("Kotlin",      0.90),
    ("SAS",         0.80),
    ("TypeScript",  0.75),
    ("Scala",       0.60),
]

# Stack Overflow 开发者调查 (2024, 使用率 %)
so_data = {
    "JavaScript":  63.61,
    "Python":      51.83,
    "TypeScript":  38.87,
    "Java":        30.55,
    "C#":          27.98,
    "C++":         22.42,
    "C":           19.34,
    "PHP":         18.16,
    "Go":          14.31,
    "Rust":        12.62,
    "Kotlin":      9.54,
    "Swift":       6.43,
    "Ruby":        5.83,
    "R":           5.47,
    "SQL":        48.66,
    "Scala":       3.18,
    "MATLAB":      4.21,
    "Delphi":      2.10,
    "Fortran":     1.52,
    "SAS":         2.34,
}

# GitHub Octoverse (2024, 仓库使用占比近似值 %)
github_data = {
    "JavaScript":  22.5,
    "Python":      20.1,
    "TypeScript":  12.4,
    "Java":         9.8,
    "C#":           6.2,
    "C++":          5.8,
    "PHP":          4.5,
    "C":            4.2,
    "Go":           3.5,
    "Rust":         2.8,
    "Ruby":         2.5,
    "Swift":        2.0,
    "Kotlin":       1.8,
    "R":            1.2,
    "SQL":          0.9,
    "Scala":        0.6,
    "MATLAB":       0.3,
    "Delphi":       0.1,
    "Fortran":      0.2,
    "SAS":          0.2,
}

# 语言分类
language_type = {
    "Python": "解释型", "JavaScript": "解释型", "TypeScript": "解释型",
    "Ruby": "解释型", "PHP": "解释型", "R": "解释型", "MATLAB": "解释型",
    "SQL": "查询语言", "SAS": "统计语言",
    "Java": "混合型", "C#": "混合型", "Kotlin": "混合型", "Scala": "混合型",
    "C": "编译型", "C++": "编译型", "Go": "编译型", "Rust": "编译型",
    "Swift": "编译型", "Fortran": "编译型", "Delphi": "编译型",
}

# 主要应用领域
language_domain = {
    "Python":     "AI/ML, 数据科学, Web, 自动化",
    "JavaScript": "Web 前端, 全栈, 移动端",
    "TypeScript": "Web 前端, 大型应用",
    "Java":       "企业后端, Android, 大数据",
    "C#":         ".NET 企业, 游戏(Unity), 桌面",
    "C++":        "系统, 游戏, 嵌入式, 高性能",
    "C":          "操作系统, 嵌入式, 编译器",
    "PHP":        "Web 后端(CMS/电商)",
    "Go":         "云原生, 微服务, DevOps",
    "Rust":       "系统, WebAssembly, 安全关键",
    "Ruby":       "Web(Rails), 脚本",
    "Swift":      "iOS/macOS 开发",
    "Kotlin":     "Android, 服务端",
    "R":          "统计分析, 可视化",
    "SQL":        "数据库查询",
    "Scala":      "大数据(Spark), 分布式",
    "MATLAB":     "科学计算, 信号处理",
    "Delphi":     "桌面应用, 遗留系统",
    "Fortran":    "科学计算, HPC",
    "SAS":        "统计分析, 商业智能",
}

# ─── 样式 ───────────────────────────────────────────────────────────

DARK_BLUE = "1F4E79"
MED_BLUE = "2E75B6"
LIGHT_BLUE = "D6E4F0"
LIGHT_GRAY = "F2F2F2"
WHITE = "FFFFFF"
ACCENT_GREEN = "548235"
ACCENT_ORANGE = "ED7D31"

title_font = Font(name="Microsoft YaHei", size=18, bold=True, color=WHITE)
title_fill = PatternFill("solid", fgColor=DARK_BLUE)
subtitle_font = Font(name="Microsoft YaHei", size=12, bold=True, color=WHITE)
subtitle_fill = PatternFill("solid", fgColor=MED_BLUE)
header_font = Font(name="Microsoft YaHei", size=10, bold=True, color=WHITE)
header_fill = PatternFill("solid", fgColor=MED_BLUE)
data_font = Font(name="Microsoft YaHei", size=10)
data_font_bold = Font(name="Microsoft YaHei", size=10, bold=True)
stripe_fill = PatternFill("solid", fgColor=LIGHT_GRAY)
pct_font_green = Font(name="Microsoft YaHei", size=10, color=ACCENT_GREEN, bold=True)
pct_font_orange = Font(name="Microsoft YaHei", size=10, color=ACCENT_ORANGE)
thin_border = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)
center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)


def write_title_row(ws, row, text, cols):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = title_font
    cell.fill = title_fill
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = 40


def write_subtitle_row(ws, row, text, cols):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = subtitle_font
    cell.fill = subtitle_fill
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 28


def write_header_row(ws, row, headers):
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = thin_border
    ws.row_dimensions[row].height = 30


def write_data_row(ws, row, values, stripe=False):
    for col, v in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=v)
        cell.font = data_font
        cell.alignment = center_align if col != 2 else left_align
        cell.border = thin_border
        if stripe:
            cell.fill = stripe_fill
    ws.row_dimensions[row].height = 22


# ─── Sheet 1: 综合对比表 ────────────────────────────────────────────

def build_comparison_sheet(wb):
    ws = wb.create_sheet("综合对比")

    headers = [
        "排名", "编程语言", "语言类型", "主要应用领域",
        "TIOBE 指数 (%)", "SO 使用率 (%)", "GitHub 占比 (%)",
        "综合评分",
    ]

    write_title_row(ws, 1, "编程语言市场份额综合对比", len(headers))
    write_subtitle_row(ws, 2,
        "数据来源: TIOBE Index (2025.05) | Stack Overflow Developer Survey 2024 | GitHub Octoverse 2024",
        len(headers))

    write_header_row(ws, 4, headers)

    # 计算综合评分: TIOBE*0.35 + SO*0.35 + GitHub*0.30 (归一化后)
    # 归一化: 各列最大值=100
    max_tiobe = max(v for _, v in tiobe_data)
    max_so = max(so_data.values())
    max_gh = max(github_data.values())

    rows = []
    for name, tiobe_val in tiobe_data:
        so_val = so_data.get(name, 0)
        gh_val = github_data.get(name, 0)
        norm_t = tiobe_val / max_tiobe * 100
        norm_s = so_val / max_so * 100
        norm_g = gh_val / max_gh * 100
        composite = round(norm_t * 0.35 + norm_s * 0.35 + norm_g * 0.30, 1)
        rows.append((name, tiobe_val, so_val, gh_val, composite))

    # 按综合评分排序
    rows.sort(key=lambda x: -x[4])

    for i, (name, t_val, s_val, g_val, comp) in enumerate(rows):
        r = 5 + i
        stripe = i % 2 == 1
        ltype = language_type.get(name, "")
        domain = language_domain.get(name, "")
        write_data_row(ws, r,
            [i + 1, name, ltype, domain, t_val, s_val, g_val, comp],
            stripe=stripe)

    # 列宽
    widths = [6, 14, 10, 32, 14, 14, 14, 10]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ── 柱状图: TIOBE Top 10 ──
    chart1 = BarChart()
    chart1.type = "col"
    chart1.title = "TIOBE 指数 Top 10"
    chart1.y_axis.title = "占比 (%)"
    chart1.x_axis.title = "编程语言"
    chart1.style = 10
    chart1.width = 24
    chart1.height = 14

    # 数据在 E5:E14 (TIOBE 列)
    data_ref = Reference(ws, min_col=5, min_row=4, max_row=14)
    cats_ref = Reference(ws, min_col=2, min_row=5, max_row=14)
    chart1.add_data(data_ref, titles_from_data=True)
    chart1.set_categories(cats_ref)
    chart1.shape = 4
    ws.add_chart(chart1, "A26")


# ─── Sheet 2: TIOBE 详细 ───────────────────────────────────────────

def build_tiobe_sheet(wb):
    ws = wb.create_sheet("TIOBE 指数")

    headers = ["排名", "编程语言", "TIOBE 指数 (%)", "同比变化", "趋势"]

    write_title_row(ws, 1, "TIOBE 编程语言社区指数 (2025年5月)", len(headers))
    write_subtitle_row(ws, 2,
        "来源: https://www.tiobe.com/tiobe-index/ | 指数反映搜索引擎结果数量",
        len(headers))
    write_header_row(ws, 4, headers)

    # 同比变化 (模拟, 基于公开趋势)
    yoy_change = {
        "Python": "+4.5%", "C": "-1.2%", "C++": "+1.8%", "Java": "-2.1%",
        "C#": "+0.5%", "JavaScript": "-0.8%", "Go": "+1.2%", "SQL": "-0.3%",
        "Fortran": "+0.9%", "Rust": "+1.1%", "Swift": "+0.3%", "Delphi": "-0.4%",
        "PHP": "-1.0%", "Ruby": "-0.5%", "MATLAB": "+0.2%", "R": "-0.1%",
        "Kotlin": "+0.4%", "SAS": "-0.2%", "TypeScript": "+0.6%", "Scala": "-0.3%",
    }

    for i, (name, val) in enumerate(tiobe_data):
        r = 5 + i
        stripe = i % 2 == 1
        change = yoy_change.get(name, "N/A")
        trend = "上升" if change.startswith("+") else ("下降" if change.startswith("-") else "持平")
        write_data_row(ws, r, [i + 1, name, val, change, trend], stripe=stripe)
        # 趋势颜色
        trend_cell = ws.cell(row=r, column=5)
        if trend == "上升":
            trend_cell.font = pct_font_green
        elif trend == "下降":
            trend_cell.font = pct_font_orange

    widths = [6, 14, 16, 12, 8]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # 饼图
    chart = PieChart()
    chart.title = "TIOBE 指数分布 (Top 10)"
    chart.style = 10
    chart.width = 20
    chart.height = 14

    data_ref = Reference(ws, min_col=3, min_row=4, max_row=14)
    cats_ref = Reference(ws, min_col=2, min_row=5, max_row=14)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    ws.add_chart(chart, "A26")


# ─── Sheet 3: Stack Overflow 详细 ──────────────────────────────────

def build_so_sheet(wb):
    ws = wb.create_sheet("Stack Overflow 调查")

    headers = ["排名", "编程语言", "使用率 (%)", "开发者人数(估)", "类型"]

    write_title_row(ws, 1, "Stack Overflow 开发者调查 - 编程语言使用率 (2024)", len(headers))
    write_subtitle_row(ws, 2,
        "来源: https://survey.stackoverflow.co/2024/ | 基于全球开发者问卷",
        len(headers))
    write_header_row(ws, 4, headers)

    # 按 SO 使用率排序
    so_sorted = sorted(so_data.items(), key=lambda x: -x[1])

    for i, (name, val) in enumerate(so_sorted):
        r = 5 + i
        stripe = i % 2 == 1
        # 估算开发者人数(百万): 假设全球约 2800 万专业开发者
        est_dev = round(val / 100 * 28, 1)
        ltype = language_type.get(name, "")
        write_data_row(ws, r, [i + 1, name, val, est_dev, ltype], stripe=stripe)

    widths = [6, 14, 12, 16, 10]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # 柱状图
    chart = BarChart()
    chart.type = "col"
    chart.title = "Stack Overflow 语言使用率 Top 10"
    chart.y_axis.title = "使用率 (%)"
    chart.style = 10
    chart.width = 24
    chart.height = 14

    data_ref = Reference(ws, min_col=3, min_row=4, max_row=14)
    cats_ref = Reference(ws, min_col=2, min_row=5, max_row=14)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    ws.add_chart(chart, "A26")


# ─── Sheet 4: GitHub 详细 ──────────────────────────────────────────

def build_github_sheet(wb):
    ws = wb.create_sheet("GitHub 仓库统计")

    headers = ["排名", "编程语言", "仓库占比 (%)", "贡献者趋势", "主要项目类型"]

    write_title_row(ws, 1, "GitHub Octoverse - 编程语言仓库统计 (2024)", len(headers))
    write_subtitle_row(ws, 2,
        "来源: https://github.com/github/octoverse | 基于公开仓库语言分布",
        len(headers))
    write_header_row(ws, 4, headers)

    gh_project_type = {
        "JavaScript": "Web, NPM 包", "Python": "AI/ML, 工具",
        "TypeScript": "Web, 框架", "Java": "企业, Android",
        "C#": ".NET, 游戏", "C++": "系统, 游戏", "PHP": "Web, CMS",
        "C": "系统, 嵌入式", "Go": "云, DevOps", "Rust": "系统, WebAssembly",
        "Ruby": "Web, 脚本", "Swift": "iOS, macOS", "Kotlin": "Android",
        "R": "数据科学", "SQL": "数据库", "Scala": "大数据",
        "MATLAB": "科学计算", "Delphi": "桌面", "Fortran": "HPC",
        "SAS": "商业分析",
    }

    gh_trend = {
        "JavaScript": "稳定", "Python": "快速增长", "TypeScript": "快速增长",
        "Java": "缓慢下降", "C#": "稳定", "C++": "稳定", "PHP": "缓慢下降",
        "C": "稳定", "Go": "增长", "Rust": "快速增长", "Ruby": "缓慢下降",
        "Swift": "稳定", "Kotlin": "增长", "R": "稳定", "SQL": "稳定",
        "Scala": "缓慢下降", "MATLAB": "稳定", "Delphi": "缓慢下降",
        "Fortran": "稳定", "SAS": "缓慢下降",
    }

    gh_sorted = sorted(github_data.items(), key=lambda x: -x[1])

    for i, (name, val) in enumerate(gh_sorted):
        r = 5 + i
        stripe = i % 2 == 1
        trend = gh_trend.get(name, "N/A")
        proj = gh_project_type.get(name, "")
        write_data_row(ws, r, [i + 1, name, val, trend, proj], stripe=stripe)
        # 趋势颜色
        t_cell = ws.cell(row=r, column=4)
        if "增长" in trend:
            t_cell.font = pct_font_green
        elif "下降" in trend:
            t_cell.font = pct_font_orange

    widths = [6, 14, 14, 14, 20]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # 饼图
    chart = PieChart()
    chart.title = "GitHub 仓库语言分布 (Top 10)"
    chart.style = 10
    chart.width = 20
    chart.height = 14

    data_ref = Reference(ws, min_col=3, min_row=4, max_row=14)
    cats_ref = Reference(ws, min_col=2, min_row=5, max_row=14)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    ws.add_chart(chart, "A26")


# ─── Sheet 5: 语言分类对比 ─────────────────────────────────────────

def build_category_sheet(wb):
    ws = wb.create_sheet("语言分类对比")

    headers = ["语言类型", "代表语言", "TIOBE 合计 (%)", "SO 合计 (%)", "GitHub 合计 (%)", "特点"]

    write_title_row(ws, 1, "编程语言分类市场份额对比", len(headers))
    write_subtitle_row(ws, 2,
        "按语言类型分组, 展示各类别的市场份额汇总",
        len(headers))
    write_header_row(ws, 4, headers)

    categories = {
        "解释型": {
            "langs": ["Python", "JavaScript", "TypeScript", "Ruby", "PHP", "R", "MATLAB"],
            "feature": "开发效率高, 部署灵活, 运行时性能较低",
        },
        "编译型": {
            "langs": ["C", "C++", "Go", "Rust", "Swift", "Fortran", "Delphi"],
            "feature": "运行性能高, 类型安全, 编译周期较长",
        },
        "混合型": {
            "langs": ["Java", "C#", "Kotlin", "Scala"],
            "feature": "跨平台, JIT 优化, 生态丰富",
        },
        "查询/统计": {
            "langs": ["SQL", "SAS"],
            "feature": "领域专用, 声明式, 非通用编程",
        },
    }

    tiobe_dict = dict(tiobe_data)
    r = 5
    for i, (cat, info) in enumerate(categories.items()):
        langs = info["langs"]
        t_sum = sum(tiobe_dict.get(l, 0) for l in langs)
        s_sum = sum(so_data.get(l, 0) for l in langs)
        g_sum = sum(github_data.get(l, 0) for l in langs)
        stripe = i % 2 == 1
        write_data_row(ws, r,
            [cat, ", ".join(langs), round(t_sum, 2), round(s_sum, 2), round(g_sum, 2), info["feature"]],
            stripe=stripe)
        r += 1

    widths = [12, 42, 14, 14, 14, 36]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # 柱状图: 分类对比
    chart = BarChart()
    chart.type = "col"
    chart.title = "语言分类市场份额对比"
    chart.y_axis.title = "占比 (%)"
    chart.style = 10
    chart.width = 22
    chart.height = 14

    data_ref = Reference(ws, min_col=3, min_row=4, max_col=5, max_row=8)
    cats_ref = Reference(ws, min_col=1, min_row=5, max_row=8)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.shape = 4
    ws.add_chart(chart, "A11")


# ─── 主函数 ─────────────────────────────────────────────────────────

def main():
    os.makedirs("output", exist_ok=True)

    wb = Workbook()
    # 删除默认 Sheet
    wb.remove(wb.active)

    build_comparison_sheet(wb)
    build_tiobe_sheet(wb)
    build_so_sheet(wb)
    build_github_sheet(wb)
    build_category_sheet(wb)

    output_path = os.path.join("output", "编程语言市场份额对比.xlsx")
    wb.save(output_path)
    print(f"报告已生成: {output_path}")


if __name__ == "__main__":
    main()
