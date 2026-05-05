import pandas as pd
import os


def create_sample_excel():
    """创建示例Excel文件"""
    os.makedirs("data", exist_ok=True)

    data1 = {
        "ID": [1, 2, 3, 4, 5],
        "姓名": ["张三", "李四", "王五", "赵六", "钱七"],
        "年龄": [25, 30, 28, 35, 22],
        "部门": ["销售", "技术", "市场", "人事", "销售"]
    }

    data2 = {
        "ID": [6, 7, 8, 9, 10],
        "姓名": ["孙八", "周九", "吴十", "郑十一", "王十二"],
        "年龄": [29, 31, 27, 33, 26],
        "部门": ["技术", "市场", "销售", "技术", "人事"]
    }

    df1 = pd.DataFrame(data1)
    df2 = pd.DataFrame(data2)

    with pd.ExcelWriter("data/example1.xlsx", engine="openpyxl") as writer:
        df1.to_excel(writer, sheet_name="员工表", index=False)
        df2.to_excel(writer, sheet_name="员工表2", index=False)

    with pd.ExcelWriter("data/example2.xlsx", engine="openpyxl") as writer:
        df2.to_excel(writer, sheet_name="Sheet1", index=False)

    df3 = pd.DataFrame({
        "产品": ["A", "B", "C"],
        "数量": [100, 200, 150],
        "价格": [10.5, 20.0, 15.8]
    })
    df3.to_excel("data/product.xlsx", index=False)

    print("示例文件创建已就绪:")
    print("  - data/example1.xlsx (多工作表)")
    print("  - data/example2.xlsx")
    print("  - data/product.xlsx")


if __name__ == "__main__":
    create_sample_excel()
