from excel_processor import ExcelProcessor


def example_usage():
    """使用示例"""
    processor = ExcelProcessor()

    print("=" * 50)
    print("示例1: 读取单个文件")
    print("=" * 50)
    try:
        df = processor.read_single_file("data/example.xlsx")
        print("\n前5行数据:")
        print(df.head())
        print("\n列名:")
        print(df.columns.tolist())
    except Exception as e:
        print(f"错误: {str(e)}")

    print("\n" + "=" * 50)
    print("示例2: 批量读取目录中的文件")
    print("=" * 50)
    try:
        results = processor.read_multiple_files("data/")
        print("\n读取的文件:")
        for filename, df in results.items():
            print(f"- {filename}: {df.shape[0]} 行 x {df.shape[1]} 列")
    except Exception as e:
        print(f"错误: {str(e)}")

    print("\n" + "=" * 50)
    print("示例3: 读取所有工作表")
    print("=" * 50)
    try:
        sheets = processor.read_multiple_sheets("data/example.xlsx")
        print("\n读取的工作表:")
        for sheet_name, df in sheets.items():
            print(f"- {sheet_name}: {df.shape[0]} 行 x {df.shape[1]} 列")
    except Exception as e:
        print(f"错误: {str(e)}")

    print("\n" + "=" * 50)
    print("示例4: 批量读取并合并数据")
    print("=" * 50)
    try:
        results = processor.read_multiple_files("data/")
        if results:
            dfs = list(results.values())
            merged_df = processor.merge_dataframes(dfs, axis=0)
            print("\n合并后的数据:")
            print(merged_df.head())
            print(f"\n总行数: {merged_df.shape[0]}")
            processor.save_dataframe(merged_df, "output/merged_result.xlsx")
    except Exception as e:
        print(f"错误: {str(e)}")

    print("\n" + "=" * 50)
    print("示例5: 获取文件信息")
    print("=" * 50)
    try:
        info = processor.get_file_info("data/example.xlsx")
        print(f"\n文件路径: {info['file_path']}")
        print(f"工作表数量: {info['sheet_count']}")
        print(f"工作表列表: {info['sheet_names']}")
        print("\n各工作表列名:")
        for sheet_name, sheet_info in info['sheets'].items():
            print(f"- {sheet_name}: {sheet_info['columns']}")
    except Exception as e:
        print(f"错误: {str(e)}")


if __name__ == "__main__":
    example_usage()
