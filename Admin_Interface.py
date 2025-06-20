import PySimpleGUI as sg
import os
from pathlib import Path
from reconciles import process_inventory_data, process_requisition_data, compare_data, mark_items_with_colors

layout = [
    [sg.Text('请选择需要比对的Excel文件：'), sg.Input(key='-EXCEL-', enable_events=True), sg.FileBrowse('选择', file_types=(("Excel Files", "*.xlsx;*.xls"),))],
    [sg.Button('开始比对并标色'), sg.Button('退出')],
    [sg.Multiline(size=(80, 20), key='-OUTPUT-', autoscroll=True, disabled=True, font=('Consolas', 10))]
]

window = sg.Window('盘点与领用单自动比对标色', layout, finalize=True)

def print_to_window(window, text):
    window['-OUTPUT-'].print(text)

while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, '退出'):
        break
    if event == '开始比对并标色':
        window['-OUTPUT-'].update('')
        excel_file = values['-EXCEL-'].strip()
        if not excel_file:
            print_to_window(window, "请先选择Excel文件！")
            continue
        try:
            inventory_sheets = ["电类盘点", "水类盘点"]
            requisition_sheet = "领用单"
            # 处理盘点表
            print_to_window(window, f"正在处理盘点表：{excel_file}")
            inventory_data = process_inventory_data(excel_file, inventory_sheets)
            # 处理领用单
            print_to_window(window, f"正在处理领用单：{excel_file}")
            requisition_data = process_requisition_data(excel_file, requisition_sheet)
            # 比较
            consistent, inconsistent, only_in_inventory, only_in_requisition = compare_data(inventory_data, requisition_data)
            print_to_window(window, "\n一致的品名及数量:")
            for item, qty in consistent:
                print_to_window(window, f"{item}: {qty}")
            print_to_window(window, "\n不一致的品名:")
            for detail in inconsistent:
                print_to_window(window, f"{detail['品名']} | 盘点数量:{detail['盘点数量']} | 领用单数量:{detail['领用单数量']} | 差异:{detail['差异']}")
            print_to_window(window, "\n只在盘点表有的品名:")
            print_to_window(window, "，".join(sorted(only_in_inventory)) if only_in_inventory else "无")
            print_to_window(window, "\n只在领用单有的品名:")
            print_to_window(window, "，".join(sorted(only_in_requisition)) if only_in_requisition else "无")
            # 标色并输出
            mark_items_with_colors(
                file_path=excel_file,
                sheet_names=inventory_sheets,
                inconsistent_names=set(item['品名'] for item in inconsistent),
                unique_names=only_in_inventory,
                consistent_names=set(item for item, _ in consistent),
                name_col="品名",
                qty_col="本次领用",
                diff_dict=[
                    {'品名': d['品名'], '领用单数量': d['领用单数量'], '盘点数量': d['盘点数量']}
                    for d in inconsistent
                ]
            )
            mark_items_with_colors(
                file_path=excel_file,
                sheet_names=[requisition_sheet],
                inconsistent_names=set(item['品名'] for item in inconsistent),
                unique_names=only_in_requisition,
                consistent_names=set(item for item, _ in consistent),
                name_col="品名",
                qty_col="数量",
                diff_dict=[
                    {'品名': d['品名'], '领用单数量': d['领用单数量'], '盘点数量': d['盘点数量']}
                    for d in inconsistent
                ]
            )
            print_to_window(window, "\n处理完成！标色后的文件已输出在原文件夹，文件名前缀为“标色_”。")
        except Exception as e:
            print_to_window(window, f"发生错误：{e}")

window.close()