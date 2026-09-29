# -*- coding: utf-8 -*-
import openpyxl, os
TPL = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板'
for fn, want in [('B7-9.6米-京AEP303.xlsx', '2026-09-18'), ('C6-7.6米-粤BDW641.xlsx', '2026-08-21')]:
    wb = openpyxl.load_workbook(os.path.join(TPL, fn), read_only=True, data_only=True)
    ws = wb['加油数据']
    print(f'=== {fn[:2]} 加油数据里 {want} 的记录:')
    hit = False
    from collections import Counter
    types = Counter()
    strs = []
    anchor = None
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if row and row[12] is not None:
            types[type(row[12]).__name__] += 1
            if not hasattr(row[12], 'year') and not isinstance(row[12], (int, float)):
                strs.append((i, str(row[12])[:19], row[4]))
            if str(row[12])[:10] == want and anchor is None:
                anchor = (i, type(row[12]).__name__, str(row[12])[:19], row[1], row[4], row[3], row[8])
                hit = True
    print('  锚点日期首条:', anchor if hit else '(无此日期记录!)')
    print(f'  M列类型统计: {dict(types)}')
    print(f'  字符串时间记录: {strs[:6]} (共{len(strs)})')
    wb.close()
