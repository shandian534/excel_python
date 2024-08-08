import pandas as pd
import numpy as np
from xlsxwriter import Workbook

# 读取Excel文件

if __name__ == '__main__':
    file_path = r'C:\Users\Acer\Desktop\重复1.xlsx'
    file_path_out = r'C:\Users\Acer\Desktop\重复2.xlsx'
    sheet_name = 'Sheet1'  # 替换为你的工作表名称

    # 使用pandas读取Excel文件
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # 指定单元格区域，例如 A1 到 C3
    start_row, start_col = 0, 3  # A1 的行列索引
    end_row, end_col = 301, 13  # C3 的行列索引

    # 遍历指定的单元格区域
    for row in range(start_row, end_row + 1):
        col_val = {}
        for col in range(start_col, end_col + 1):
            cell_value = df.iloc[row, col]
            if col_val.get(cell_value):
                print("重复值")
                df.iat[row, col] = np.nan
            else:
                col_val[cell_value] = cell_value
            print(f'Row {row + 1}, Column {col + 1}: {cell_value}')
    df.to_excel(file_path_out, sheet_name='Sheet1', index=False)


