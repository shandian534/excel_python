import os
import openpyxl

if __name__ == '__main__':
    path = r"C:\Users\Acer\Desktop"
    os.chdir(path)  # 修改工作路径

    workbook = openpyxl.load_workbook('重复1.xlsx')  # 返回一个workbook数据类型的值
    print(workbook.sheetnames)  # 打印Excel表中的所有表
    sheet = workbook.active  # 获取活动表
    print(sheet)
    # 结果：
    # ['Sheet1', 'Sheet2']



