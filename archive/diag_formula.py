# -*- coding: utf-8 -*-
"""看 4 张油耗 sheet 的完整公式, 理解"当日未加油"与达标率的判定机制"""
import openpyxl
CUR = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
wb = openpyxl.load_workbook(CUR, data_only=False)
ws = wb['华东油耗']
print('=== 华东油耗 表头 6 行 (A~T):')
for c in range(1, 21):
    v = ws.cell(6, c).value
    if v is not None: print(f'   {ws.cell(6,c).coordinate} = {v}')
print()
for coord in ('D7', 'E7', 'F7', 'G7', 'K7', 'L7', 'M7', 'N7', 'O7', 'S7', 'T7'):
    v = ws.cell(7, openpyxl.utils.column_index_from_string(coord[0])).value
    t = v.text if hasattr(v, 'text') else v
    print(f'  {coord} = {str(t)[:230]}')
print()
# 第二行(粤BT802F)对比
print('=== 行8 (粤BT802F):')
for coord in ('D8', 'F8', 'G8', 'N8'):
    v = ws.cell(8, openpyxl.utils.column_index_from_string(coord[0])).value
    t = v.text if hasattr(v, 'text') else v
    print(f'  {coord} = {str(t)[:230]}')
print()
print('=== 行9 (粤BY2J27, 有数字达标率):')
for coord in ('F9', 'G9', 'N9'):
    v = ws.cell(9, openpyxl.utils.column_index_from_string(coord[0])).value
    t = v.text if hasattr(v, 'text') else v
    print(f'  {coord} = {str(t)[:230]}')
wb.close()
