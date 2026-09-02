# 项目架构说明

> 本文档描述 `py_excel` 仓库的实际结构与运行方式。
> 所有结论均在 Python 3.14.3 / openpyxl 3.1.5 / pandas 3.0.5 实测确认，未经验证的部分已明确标注。

## 1. 项目定位

两件互不相关的事被放在同一个仓库里：

| 用途 | 入口 | 依赖 |
|---|---|---|
| **A. Excel 批量读取工具库** | `excel_processor.py` | pandas |
| **B. 行业分析报告生成器** | `generate_*.py`（13 个） | openpyxl |

**关键点：这两条路径在代码层面零交集。** 没有任何生成器 `import excel_processor`，也没有任何生成器读取外部数据。把它们当成两个独立子项目理解，比试图找出统一架构更准确。

## 2. 目录结构

```
.
├── excel_processor.py          # A: ExcelProcessor 类，6 个方法
├── example.py                  # A: 使用示例（5 个演示块）
├── create_sample.py            # A: 生成 data/ 下的样例 xlsx
├── generate_*.py               # B: 13 个报告生成器，各自独立
├── docs/ARCHITECTURE.md        # 本文档
├── requirements.txt
├── CLAUDE.md                   # 给 AI 助手的项目约定
├── data/                       # 运行时生成，已 gitignore
└── output/                     # 运行时生成，已 gitignore
```

`data/` 和 `output/` 都由代码用 `mkdir(exist_ok=True)` / `os.makedirs` 自建，不需要手动创建，也不入库。

## 3. 路径 A：Excel 读取工具库

`ExcelProcessor`（`excel_processor.py`，131 行）是一个无状态的薄封装，6 个方法：

| 方法 | 作用 |
|---|---|
| `read_single_file(path, sheet_name=0, header=0)` | 读单文件，返回 DataFrame |
| `read_multiple_files(directory, ...)` | 读目录下所有 `.xlsx`/`.xls`，返回 `{filename: df}`；单文件失败只跳过不中断 |
| `read_multiple_sheets(path, header=0)` | 读单文件全部工作表，返回 `{sheet_name: df}` |
| `save_dataframe(df, output_path, ...)` | 写出 xlsx，会自动创建父目录 |
| `merge_dataframes(dfs, axis=0, ignore_index=True)` | `pd.concat` 封装 |
| `get_file_info(path)` | 返回工作表数量、名称、各表列名 |

设计特征：所有方法在失败时抛 `Exception`（字符串包装原始错误），`__init__` 里的 `self.data_store = {}` 目前没有任何方法使用。

跑通方式：
```bash
python create_sample.py     # 先生成 data/example1.xlsx 等 3 个文件
python example.py           # 依赖上一步的产物
```
`example.py` 必须在 `create_sample.py` 之后运行，它读的是 `data/example1.xlsx`。

## 4. 路径 B：报告生成器

### 4.1 统一的四段式结构

每个生成器都是**自包含单文件**，从上到下固定四段：

```python
# ① 输出路径常量
OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)
OUTPUT_FILE = OUTPUT_DIR / "电视机市场分析报告.xlsx"

# ② 样式常量：配色 dict + Border 对象
C = {"title_bg": "1F4E79", ...}

# ③ 样式/写入/图表 helper 函数
def write_data_row(ws, row, values, ...): ...
def add_pie(ws, ...): ...

# ④ 每个工作表一个 build_*_sheet(wb) 函数，main() 依次调用
def build_overview(wb):
    ws = wb.create_sheet("市场概览", 0)
    ...
def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)        # 删掉默认 Sheet
    for name, builder in sheets:
        builder(wb)
    wb.save(str(OUTPUT_FILE))
```

`wb.remove(wb.active)` 这一步很关键 —— 它删掉 openpyxl 的默认工作表，因此**后续必须至少 `create_sheet` 一次**，否则保存时报 `IndexError: At least one sheet must be visible`（见 §6）。

### 4.2 数据全部硬编码

全仓库 `grep read_excel|requests|urllib|json.load|csv` 在 `generate_*.py` 中**零匹配**。所有市场数据（份额、价格、品牌、技术参数）都以字面量写在 `build_*_sheet` 函数里。

这意味着：
- 生成器是**确定性**的，同一份代码永远产出同样的报告，跑多少次都一样；
- 更新数据 = 改代码，没有配置层也没有数据层；
- 数据的准确性和时效性由写入时的人负责，仓库本身不做校验。

### 4.3 报告清单

| 脚本 | 输出文件 | 工作表 | 状态 |
|---|---|---|---|
| `generate_audio_market_report.py` | 音响市场分析报告.xlsx | 市场概览 / TWS耳机 / 头戴耳机 / 智能与蓝牙音箱 / 家庭影院 / 专业音响 / 汽车音响 / 发展趋势 | 实测通过 |
| `generate_camera_market_report.py` | 数码相机市场分析报告.xlsx | 市场概览 / 无反相机 / 运动与全景 / 电影机 / 卡片机与拍立得 / 镜头 / 发展趋势 | 实测通过 |
| `generate_keyboard_mouse_report.py` | 键盘鼠标市场分析报告.xlsx | 市场概览 / 键盘 / 鼠标 / 键盘轴体 / 鼠标微动与传感器 / 技术特点对比 / 发展趋势 | 实测通过 |
| `generate_printer_market_report.py` | 打印机市场分析报告.xlsx | 市场概览 / 激光 / 喷墨 / 大幅面 / 3D打印 / 针式与特种 / 技术对比 / 发展趋势 | 实测通过 |
| `generate_router_market_report.py` | 路由器市场分析报告.xlsx | 市场概览 / 家用Wi-Fi / 企业级 / 运营商 / 5G与工业 / 发展趋势 | 实测通过 |
| `generate_tv_market_report.py` | 电视机市场分析报告.xlsx | 市场概览 / 品牌份额 / 面板厂商 / 显示技术对比 / 高端电视技术 / 发展趋势 | 实测通过 |
| `generate_smartwatch_market_report.py` | 智能手表市场分析报告.xlsx | 市场概览 / 旗舰 / 运动健康 / 轻智能与儿童 / 发展趋势 | 实测通过 |
| `generate_language_market_report.py` | 编程语言市场份额对比.xlsx | 综合对比 / TIOBE 指数 / Stack Overflow 调查 / GitHub 仓库统计 / 语言分类对比 | 实测通过 |
| `generate_dashboard.py` | 经营数据大屏.xlsx | 经营总览 / 项目进度看板 / 人员效能看板 / 经营概览(浅色) | 实测通过 |
| `generate_chip_market_report.py` | 芯片市场分析报告.xlsx | — | **骨架，跑不通** |
| `generate_drone_market_report.py` | 航拍无人机市场分析报告.xlsx | — | **骨架，跑不通** |
| `generate_monitor_market_report.py` | 显示器市场分析报告.xlsx | — | **骨架，跑不通** |
| `generate_storage_market_report.py` | 存储器市场分析报告.xlsx | — | **骨架，跑不通** |

`generate_dashboard.py` 是唯一的异类：它不是市场分析报告，而是经营数据看板，有深浅两套配色常量（`DARK` / `LIGHT`）、KPI 卡片（`write_kpi_card`）、趋势单元格（`write_trend_cell`）和折线图（`add_line_chart`），这些在其他生成器里都不存在。

### 4.4 图表

用到 3 种 openpyxl 图表：`PieChart`、`BarChart`（`type="col"`, `shape=4`）、`LineChart`（仅 dashboard）。饼图统一开 `showPercent` + `showCatName`。

## 5. helper 函数的四套命名（技术债）

生成器之间**没有共享模块** —— helper 是复制粘贴的，且在复制过程中发生了命名漂移，目前存在四套互不兼容的命名：

| 世代 | 命名风格 | 使用者 |
|---|---|---|
| **A** | `_font` `_fill` `stripe_fill` `set_col_widths` `write_title_row` `write_source_row` `write_subtitle_row` `write_header_row` `write_data_row` `add_pie` `add_bar` | audio, camera, tv, printer, router, smartwatch |
| **A⁻** | 同 A，但无 `write_source_row` | chip（骨架） |
| **B** | `title_font` `header_font` `body_font` `title_fill` `header_fill` `sub_fill` `stripe_fill` `set_col_widths` `write_*_row` `add_pie_chart` `add_bar_chart` | keyboard_mouse, storage（骨架） |
| **C** | 缩写式：`_f` `_fl` `stripe` `setw` `title_row` `sub_row` `hdr_row` `data_row` `add_pie` `add_bar` | drone（骨架）, monitor（骨架） |
| **D** | 只有 4 个 `write_*_row` | language |
| **E** | dashboard 专用：`set_sheet_bg` `write_kpi_card` `write_title_bar` `write_section_title` `write_trend_cell` `add_line_chart` | dashboard |

配色常量名也不统一：多数用 `C`，keyboard_mouse/storage 用 `COLORS`，dashboard 用 `DARK`/`LIGHT`，language 没有顶层配色 dict。

**影响**：改一处样式需要在多个文件里分别改，且改法因文件而异。抽取共享模块（如 `report_style.py`）是明显的改进方向，但会触及全部 13 个文件，属于一次性大改动。

## 6. 已知问题

**① 4 个生成器是未完成的骨架。** `chip` / `drone` / `monitor` / `storage` 只有常量和 helper，`build_*_sheet` 函数完全不存在，`main()` 里是注释：

```python
# Sheet builders are in the full local file
def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)        # 删掉唯一的 sheet
    # build_overview_sheet(wb) ... etc     ← 没有实现
    wb.save(str(OUTPUT_FILE))   # 零个 sheet → IndexError
```

运行会抛 `IndexError: At least one sheet must be visible`，**且崩溃前 openpyxl 已在 `output/` 留下约 3.3KB 的残缺 xlsx** —— 这个文件是坏的，看到时直接删。

注意 `CLAUDE.md` 的"运行命令"里把 `generate_storage_market_report.py` 列为示例，那条命令现在跑不通。

**② 控制台编码。** 本机 `sys.stdout.encoding` 是 `gbk`，脚本里的中文 `print` 会输出乱码。加环境变量解决：
```bash
PYTHONIOENCODING=utf-8 python generate_tv_market_report.py
```
不影响生成的 xlsx，只影响终端显示。

**③ `ExcelProcessor.data_store` 是死代码**，`__init__` 里初始化后无人使用。

## 7. 环境与运行

```bash
pip install -r requirements.txt        # openpyxl>=3.1, pandas>=2.1
```

依赖用的是下界而非精确版本，因为原先钉的 `pandas==2.1.4` 在 Python 3.14 上没有 wheel，会退化成源码编译并失败。3.14 环境下实际解析到 **pandas 3.0.5 + openpyxl 3.1.5 + numpy 2.5.2**，这是本文档所有结论的验证环境。`>=2.1` 这个下界沿用 CLAUDE.md 的既有声明，未在 pandas 2.x 上实测。

生成器只需要 openpyxl，工具库只需要 pandas —— 想跑报告不装 pandas 也可以。

```bash
PYTHONIOENCODING=utf-8 python generate_tv_market_report.py    # 单个报告
python create_sample.py && python example.py                  # 工具库演示
```

## 8. 新增一个报告生成器

复制一个已通过实测的同类文件（推荐 `generate_tv_market_report.py`，763 行，结构最清晰），然后：

1. 改 `OUTPUT_FILE` 的文件名；
2. 替换 `build_*_sheet` 里的硬编码数据；
3. 更新 `main()` 的 `sheets` 列表（工作表名 + builder 函数配对）；
4. 确认至少有一个 `create_sheet` 调用，否则触发 §6① 的 `IndexError`。

沿用被复制文件的 helper 命名世代（见 §5），不要混用。
