# -*- coding: utf-8 -*-
import openpyxl
SUM = r'C:\Users\Jubin\Desktop\源数据\油耗分析汇总.xlsm'
wb = openpyxl.load_workbook(SUM, read_only=True, data_only=True)
ws = wb['Sheet1']
rows = []
for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
    if row and any(c is not None for c in row): rows.append((i, row))
print(f'汇总总行数: {len(rows)}')
print('=== 09-20/09-21 的行 (重点看时间A、里程):')
for i, row in rows[-14:]:
    g = str(row[6])[:10] if row[6] else ''
    if g in ('2026-09-20', '2026-09-21'):
        print(f'  行{i}: 序号={row[0]} 区域={row[1]} 车牌={row[2]} | 时间A={str(row[5])[:19]} 时间B={str(row[6])[:19]} '
              f'| 金额={row[7]} 升={row[8]} | 行驶里程J={row[9]!r} GPS油耗K={row[10]!r} 达标率O={row[14]!r}')
wb.close()
