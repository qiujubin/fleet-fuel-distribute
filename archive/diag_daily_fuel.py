# -*- coding: utf-8 -*-
import openpyxl
CUR = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
wb = openpyxl.load_workbook(CUR, data_only=True)
print('全部 sheet:', wb.sheetnames)
ws = wb['每日油耗']
print(f'\n=== 每日油耗: max_row={ws.max_row} max_col={ws.max_column}')
for i, row in enumerate(ws.iter_rows(min_row=1, max_row=4, values_only=True), start=1):
    print(f'  行{i}:', [str(c)[:14] if c is not None else '' for c in row[:16]])
print('  ...尾部:')
tail = []
for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
    if row and any(c is not None for c in row[:10]): tail.append((i, row))
for i, row in tail[-4:]:
    print(f'  行{i}:', [str(c)[:14] if c is not None else '' for c in row[:16]])
print(f'  数据行数: {len(tail)}')
# D 列车牌 + 是否有 09-20 的记录(B车)
print('\n  BY2J27/BT802F 在每日油耗里的行:')
for i, row in tail:
    if row[3] and str(row[3]).strip() in ('粤BY2J27', '粤BT802F', '京LPW138'):
        print(f'  行{i}: 序号={row[0]} 类型={row[1]} B={row[2]} D={row[3]} E={row[4]} F={row[5]} G={row[6]} I={row[8]}')
wb.close()
