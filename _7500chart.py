import numpy as np
import matplotlib.pyplot as plt
from matplotlib.widgets import Button
from matplotlib.gridspec import GridSpec
import re
import xlrd
import tkinter as tk
from tkinter import BUTT, N, filedialog
import sys
import os

#button_states = [True] * 96
text_obj = None

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

    hide_index = []
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
                else:
                    if j not in hide_index:
                        hide_index.append(j)
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
    scatters = []
    # 绘制折线图
    for key in keys:
        y_values = cell_dict[key]
        line, = ax.plot(x_values, y_values, linewidth=2)
        lines.append(line)
        # 绘制最后一个点的marker
        scatter = ax.scatter(x_values[-1], y_values[-1], color=line.get_color(), 
                            marker='<', s = 100, zorder=3)
        scatters.append(scatter)

    for scatter in scatters:
            scatter.set_visible(False)

    # 设置刻度标记大小
    ax.tick_params(axis='both', which='major', labelsize=10)
    ax.set_yticks(range(0, int(y_max) + 2000, 1000))
    ax.set_yticklabels(range(0, int(y_max) + 2000, 1000))



    # 右侧按钮区域
    ax_buttons = fig.add_subplot(gs[0, 1])
    ax_buttons.set_position([0.4, 0.1, 0.25, 0.8])
    ax_buttons.axis('off')  # 隐藏坐标轴


    # 按钮点击回调函数
    button_states = [True] * 96
    for index in hide_index:
        button_states[index] = False

    
    nc = []
    nc_num = []

    def toggle_line(index, alp):
        def callback(event):
            global text_obj
            if event.button == 1 and index not in nc:
                button_states[index] = not button_states[index]
                if button_states[index] and index not in hide_index:
                    lines[index].set_alpha(1.0)
                    if not nc:
                        buttons[index].color = 'lightblue'
                        buttons[index].hovercolor = '#7fb1b3'
                    elif float(cell_dict[alp][29]) / np.average(nc_num) >= 2:
                        buttons[index].color = 'red'
                        buttons[index].hovercolor = 'lightcoral'
                    elif float(cell_dict[alp][29]) / np.average(nc_num) >= 1.5:
                        buttons[index].color = 'orange'
                        buttons[index].hovercolor = 'wheat'
                elif index not in hide_index:
                    lines[index].set_alpha(0.0)
                    buttons[index].color = 'cadetblue'
                    buttons[index].hovercolor = '#7fb1b3'
                    buttons[index].ax.set_facecolor('#7fb1b3')
            elif event.button == 3:  
                # 平均值计算
                if index not in hide_index and index not in nc:
                    nc.append(index)
                    nc_num.append(cell_dict[alp][29])
                    scatters[index].set_visible(True)
                    buttons[index].color = 'yellow'
                    buttons[index].hovercolor = 'lightyellow'
                    buttons[index].ax.set_facecolor('lightyellow')
                    buttons[index].canvas.draw()
                elif index not in hide_index and index in nc:
                    nc.remove(index)
                    nc_num.remove(cell_dict[alp][29])
                    scatters[index].set_visible(False)
                    if button_states[index]:
                        buttons[index].color = 'lightblue'
                    else:
                        buttons[index].color = 'cadetblue'
                    buttons[index].hovercolor = '#7fb1b3'
                    buttons[index].ax.set_facecolor('#7fb1b3')

                i = 0
                for row in range(8):
                    for col in range(12):
                        alph = f'{chr(row + 65)}{col + 1}'
                        if i not in hide_index and i not in nc:
                            if float(cell_dict[alph][29]) / np.average(nc_num) >= 2:
                                print(cell_dict[alph][29])
                                buttons[i].color = 'red'
                                buttons[i].hovercolor = 'lightcoral'
                                #button.ax.set_facecolor('lightcoral' if button.ax.get_facecolor() != (1.0, 1.0, 0.6274509803921569, 1.0) else button.ax.get_facecolor())
                            elif float(cell_dict[alph][29]) / np.average(nc_num) >= 1.5:
                                print(cell_dict[alph][29])
                                buttons[i].color = 'orange'
                                buttons[i].hovercolor = 'wheat'
                                buttons[i].ax.set_facecolor('orange')
                            else:
                                buttons[i].color = 'lightblue'
                                buttons[i].hovercolor = '#7fb1b3'
                        i += 1
                button.canvas.draw()
                # 文本显示
                if text_obj:
                    text_obj.remove()
                text_obj = ax.text(0.5, 0.5, np.average(nc_num) if nc_num else '', fontsize=30)

                
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
            button = Button(ax_button, f'{chr(row + 65)}{col + 1}',
                            color = 'lightblue' if index not in hide_index else 'dimgray',
                            hovercolor = '#7fb1b3' if index not in hide_index else 'dimgray')
            button.label.set_fontsize(18)
            button.on_clicked(toggle_line(index, f'{chr(row + 65)}{col + 1}'))
            buttons.append(button)
            lines[index].set_alpha(0.0) if index in hide_index else lines[index].set_alpha(1.0)


    #grid = ButtonGrid(8, 12)
    # 显示图表
    plt.show()

    


if __name__ == "__main__":
    main()