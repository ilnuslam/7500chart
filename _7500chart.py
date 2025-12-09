import numpy as np
import matplotlib.pyplot as plt
import re
import xlrd
import tkinter as tk
from tkinter import N, filedialog
import sys
import os


def choose_xls_doc():
    #选择文件
    print("please choose document")
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', 1)
    excel_path = filedialog.askopenfilename(
        title="选择qPCR数据表格",
        filetypes=[("excel", "*.xls"), ("所有文件", "*.*")]
    )
    if not excel_path:
        print("未选择文件，程序退出")
        sys.exit(0)
    print(excel_path)
    file_type = os.path.splitext(excel_path)
    return excel_path, file_type 

def main():
    #选择Excel文件
    input_source, file_type = choose_xls_doc()
    path = os.path.dirname(input_source)
    file_name = os.path.basename(input_source)

    workbook = xlrd.open_workbook(input_source)        # 打开Excel文件
    sheet = workbook.sheet_by_name('Multicomponent Data')        # 获取工作表

    cell_dict = {}
    letters = [chr(ord('A') + i) for i in range(8)]  # A-H

    y_max = 0
    j = 0
    for letter in letters:
        for number in range(1, 13):  # 1-12
            key = f"{letter}{number}"
            cell = []
            for row in range(8 + j, 2888 + j, 96):
                cell_value = sheet.cell_value(row, 2)
                if cell_value != '':
                    if float(cell_value) > y_max:
                        y_max = float(cell_value)
                cell.append(cell_value)
            j += 1
            cell_dict[key] = cell
    print(cell_dict)

    # 创建图形和坐标轴
    plt.figure(figsize=(16, 9))
    
    keys = list(cell_dict.keys())
    x_values = range(1, 31)
    # 绘制折线图
    for key in keys:
        y_values = cell_dict[key]
        plt.plot(x_values, y_values,
                 linewidth=2)

    # 设置刻度标记大小
    plt.tick_params(axis='both', which='major', labelsize=10)

    plt.yticks(range(0, int(y_max) + 2000, 1000), range(0, int(y_max) + 2000, 1000))

    # 显示图表
    plt.show()

    


if __name__ == "__main__":
    main()