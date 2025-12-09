import numpy as np
import matplotlib.pyplot as plt
from matplotlib.widgets import Button
from matplotlib.gridspec import GridSpec
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

class ButtonGrid:
    def __init__(self, rows=8, cols=12):
        self.rows = rows
        self.cols = cols
        self.fig, self.ax = plt.subplots(figsize=(cols*1.2, rows*1.2))
        self.ax.axis('off')
        self.buttons = []
        self.create_grid()
    

    def create_grid(self):
        for i in range(self.rows):
            for j in range(self.cols):
                # 计算按钮位置
                ax_button = plt.axes([0.05 + j*0.08, 0.9 - i*0.1, 0.07, 0.07])
                btn = Button(ax_button, f'{chr(j + 65)}{i + 1}', color='lightblue', hovercolor='skyblue')
                btn.label.set_fontsize(15)
                # 绑定点击事件
                btn.on_clicked(lambda event, row=i, col=j: self.on_click(row, col))
                self.buttons.append(btn)
        #plt.suptitle(f'{self.rows}x{self.cols} Button Grid', y=0.95, fontsize=14)
    
    # 按钮点击事件处理
    def on_click(self, row, col):
        print(f'Button at row {row}, column {col} was clicked')
        # 改变被点击按钮的颜色
        idx = row * self.cols + col
        self.buttons[idx].color = 'lightgreen'
        self.buttons[idx].hovercolor = 'limegreen'
        self.fig.canvas.draw()

    # 按钮释放事件处理
    def on_release(self, row, col):
        # 改变被点击按钮的颜色
        idx = row * self.cols + col
        self.buttons[idx].color = 'lightblue'
        self.buttons[idx].hovercolor = 'skyblue'
        self.fig.canvas.draw()

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
    #print(cell_dict)

    # 创建图形和坐标轴
    fig = plt.figure(figsize=(26, 12))
    gs = fig.add_gridspec(1, 2, width_ratios=[2.5, 1], wspace=0.15)
    
    # 左侧主图区域
    ax = fig.add_subplot(gs[0, 0])
    ax.set_position([0.03, 0.1, 0.38, 0.8])  # 手动调整位置和大小
    
    keys = list(cell_dict.keys())
    x_values = range(1, 31)
    lines = []
    # 绘制折线图
    for key in keys:
        y_values = cell_dict[key]
        line, = ax.plot(x_values, y_values, linewidth=2)
        lines.append(line)

    # 设置刻度标记大小
    ax.tick_params(axis='both', which='major', labelsize=10)
    ax.set_yticks(range(0, int(y_max) + 2000, 1000))
    ax.set_yticklabels(range(0, int(y_max) + 2000, 1000))

    # 右侧按钮区域
    ax_buttons = fig.add_subplot(gs[0, 1])
    ax_buttons.set_position([0.4, 0.1, 0.25, 0.8])
    ax_buttons.axis('off')  # 隐藏坐标轴


    # 按钮点击回调函数
    button_states = {}
    for letter in letters:
        for number in range(1, 13):
            key = f"{letter}{number}"
            cell = True
            button_states[key] = cell

    def toggle_line(line_index, index):
        def callback(event):
            button_states[line_index] = not button_states[line_index]
            print(line_index)
            if button_states[line_index]:
                lines[index].set_alpha(1.0)
            else:
                lines[index].set_alpha(0.0)
            fig.canvas.draw_idle()
        return callback


    # 创建8x12按钮网格
    buttons = []
    button_size = 0.1
    start_x, start_y = 0.43, 0.85  # 调整起始位置

    for row in range(8):
        for col in range(12):
            index = row * 12 + col
            ax_button = fig.add_axes([
                start_x + col * button_size * 0.46,
                start_y - (row + 0.5) * button_size,
                button_size * 0.415,
                button_size * 0.9
            ])
            button = Button(ax_button, f'{chr(row + 65)}{col + 1}', color='lightblue', hovercolor='skyblue')
            button.label.set_fontsize(18)
            button.on_clicked(toggle_line(f'{chr(row + 65)}{col + 1}', index))
            buttons.append(button)


    #grid = ButtonGrid(8, 12)
    # 显示图表
    plt.show()

    


if __name__ == "__main__":
    main()