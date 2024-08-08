
import pandas as pd


if __name__ == '__main__':
    sheet_name = 'Sheet1'  # 替换为你的工作表名称
    file_path_out = r'C:\Users\Acer\Desktop\重复2.xlsx'
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