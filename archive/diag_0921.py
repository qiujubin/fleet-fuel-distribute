# -*- coding: utf-8 -*-
"""诊断: ①D14 粤BT802F 09-20 双"是"的处理 ②4张油耗sheet为何只收集到6行旧数据"""
import openpyxl, os
from openpyxl.worksheet.formula import ArrayFormula
TPL = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板'
CUR = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'

print('========== 1) D14 粤BT802F 09-20 记录与模板 ==========')
wb = openpyxl.load_workbook(os.path.join(TPL, 'D14-4.2米-粤BT802F.xlsx'), data_only=False)
wst = wb['计算模板']; wsd = wb['加油数据']
def txt(c):
    v = wst[c].value
    return v.text if isinstance(v, ArrayFormula) else v
# 加油数据里 09-20 相关记录
for rr in range(2, wsd.max_row + 1):
    t = wsd.cell(rr, 13).value
    if t is not None and str(t)[:10] == '2026-09-20':
        print(f'  加油数据行{rr}: {wsd.cell(rr,1).value} {str(t)[:19]} 加满[{wsd.cell(rr,9).value}] ¥{wsd.cell(rr,6).value}')
print('  模板 199~202 行:')
for r in range(199, 203):
    print(f'    行{r}: F={txt(f"F{r}")} G={txt(f"G{r}")}')
wb.close()

print('========== 2) 考核表里 D14 的 09-20 行 ==========')
wb = openpyxl.load_workbook(CUR, read_only=True, data_only=True)
ws = wb['每日油耗数据']
for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
    if row and row[3] and str(row[3]).strip() == '粤BT802F' and row[6] and str(row[6])[:10] == '2026-09-20':
        print(f'  考核行{i}: F={row[5]} G={row[6]} I={row[8]} J={row[9]}')
print('========== 3) 每日加油记录里 D14 ==========')
ws2 = wb['每日加油记录']
for i, row in enumerate(ws2.iter_rows(min_row=2, values_only=True), start=2):
    if row and row[1] and str(row[1]).strip() == '粤BT802F':
        print(f'  行{i}: {row[0]} {str(row[12])[:19]} 加满[{row[8]}] ¥{row[5]}')
print('========== 4) 4张油耗sheet 华东样本(BY2J27)当前值与达标率 ==========')
ws3 = wb['华东油耗']
hdr = None
for i, row in enumerate(ws3.iter_rows(min_row=6, max_row=6, values_only=True)):
    hdr = row
print('  表头 N 列(14):', hdr[13] if hdr else '?')
for i, row in enumerate(ws3.iter_rows(min_row=7, max_row=12, values_only=True), start=7):
    if row and row[1]:
        print(f'  行{i}: 车牌={row[1]} 时间F={str(row[4])[:19] if row[4] else None} 金额G={row[6]} 达标率N={row[13]!r} 类型={type(row[13]).__name__}')
wb.close()
