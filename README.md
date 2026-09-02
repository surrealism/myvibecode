# Python Excel 处理工具

一个简单易用的Python库，用于批量读取和处理Excel表格文件，以及一组行业分析报告生成器。

> 项目结构、报告清单、已知问题详见 [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)

## 功能特性

- 读取单个Excel文件
- 批量读取目录中的所有Excel文件
- 读取单个文件中的所有工作表
- 合并多个DataFrame
- 保存DataFrame为Excel文件
- 获取Excel文件信息

## 安装依赖

```bash
pip install -r requirements.txt
```

## 主要依赖

- openpyxl: 读写 .xlsx 文件
- pandas: 数据处理和分析
