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

    result = pd.DataFrame()
    with pd.ExcelWriter(file_path_out, engine='xlsxwriter') as writer:
        # 数据传给Excel的writer
        result.to_excel(writer, index=False, sheet_name='Sheet1')
        # 再从writer加载回该sheet
        worksheet = writer.sheets[sheet_name]
        # 指定列宽设置的范围
        worksheet.set_column("D:N", 20)
        # 循环每一列的列序号，设置列宽为20个单位（指Excel中的列宽单位）  这种方法也行
        for idx in range(result.shape[1]):
            worksheet.set_column(idx, idx, 20)
        # 设置缩放比例
        worksheet.set_zoom(zoom=125)
        # writer保存数据到本地
        writer._save()
