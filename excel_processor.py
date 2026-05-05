import os
import pandas as pd
from typing import List, Dict, Optional, Union


class ExcelProcessor:
    def __init__(self):
        self.data_store = {}

    def read_single_file(
        self,
        file_path: str,
        sheet_name: Optional[Union[str, int, List[str]]] = 0,
        header: Union[int, List[int], None] = 0
    ) -> pd.DataFrame:
        """读取单个Excel文件"""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=header)
            print(f"成功读取文件: {file_path}")
            print(f"形状: {df.shape[0]} 行 x {df.shape[1]} 列")
            return df
        except Exception as e:
            raise Exception(f"读取文件失败: {str(e)}")

    def read_multiple_files(
        self,
        directory: str,
        file_pattern: str = "*.xlsx",
        sheet_name: Optional[Union[str, int, List[str]]] = 0,
        header: Union[int, List[int], None] = 0
    ) -> Dict[str, pd.DataFrame]:
        """批量读取目录中的Excel文件"""
        if not os.path.exists(directory):
            raise FileNotFoundError(f"目录不存在: {directory}")
        results = {}
        files = [f for f in os.listdir(directory) if f.endswith(('.xlsx', '.xls'))]
        if not files:
            print(f"警告: 目录 {directory} 中没有找到Excel文件")
            return results
        print(f"找到 {len(files)} 个Excel文件")
        for filename in files:
            file_path = os.path.join(directory, filename)
            try:
                df = self.read_single_file(file_path, sheet_name, header)
                results[filename] = df
            except Exception as e:
                print(f"跳过文件 {filename}: {str(e)}")
        print(f"成功读取 {len(results)} 个文件")
        return results

    def read_multiple_sheets(
        self,
        file_path: str,
        header: Union[int, List[int], None] = 0
    ) -> Dict[str, pd.DataFrame]:
        """读取单个文件中的所有工作表"""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")
        try:
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            results = {}
            for sheet_name in sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=header)
                results[sheet_name] = df
                print(f"工作表 '{sheet_name}': {df.shape[0]} 行 x {df.shape[1]} 列")
            print(f"成功读取文件: {file_path}")
            return results
        except Exception as e:
            raise Exception(f"读取文件失败: {str(e)}")

    def save_dataframe(
        self,
        df: pd.DataFrame,
        output_path: str,
        sheet_name: str = "Sheet1",
        index: bool = False
    ) -> None:
        """将DataFrame保存为Excel文件"""
        try:
            df.to_excel(output_path, sheet_name=sheet_name, index=index)
            print(f"成功保存文件: {output_path}")
        except Exception as e:
            raise Exception(f"保存文件失败: {str(e)}")

    def merge_dataframes(
        self,
        dataframes: List[pd.DataFrame],
        axis: int = 0,
        ignore_index: bool = True
    ) -> pd.DataFrame:
        """合并多个DataFrame"""
        try:
            merged = pd.concat(dataframes, axis=axis, ignore_index=ignore_index)
            print(f"合并完成，结果形状: {merged.shape[0]} 行 x {merged.shape[1]} 列")
            return merged
        except Exception as e:
            raise Exception(f"合并失败: {str(e)}")

    def get_file_info(self, file_path: str) -> Dict:
        """获取Excel文件信息"""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")
        try:
            excel_file = pd.ExcelFile(file_path)
            info = {
                "file_path": file_path,
                "sheet_names": excel_file.sheet_names,
                "sheet_count": len(excel_file.sheet_names)
            }
            sheet_info = {}
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=1)
                sheet_info[sheet_name] = {
                    "columns": list(df.columns),
                    "column_count": len(df.columns)
                }
            info["sheets"] = sheet_info
            return info
        except Exception as e:
            raise Exception(f"获取文件信息失败: {str(e)}")


if __name__ == "__main__":
    processor = ExcelProcessor()
    print("请取消注释相应的代码行来使用这些功能")
