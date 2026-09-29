# -*- coding: utf-8 -*-
"""实验: normalize_com_times 是否会破坏模板 F 列(用副本, 安全)"""
import os, shutil, sys, datetime
sys.path.insert(0, r'E:\WorkSpace\workbuddy\2026-09-18-16-38-15')
TE = r'E:\WorkSpace\workbuddy\2026-09-18-16-38-15\test_norm'
if os.path.exists(TE):
    for f in os.listdir(TE): os.remove(os.path.join(TE, f))
else:
    os.makedirs(TE)
SRC = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板\D14-4.2米-粤BT802F.xlsx'
DST = os.path.join(TE, 'D14_copy.xlsx')
shutil.copy2(SRC, DST)

import pythoncom
import win32com.client as wc
pythoncom.CoInitialize()
app = None
try:
    app = wc.DispatchEx('Ket.Application')
    app.Visible = False; app.DisplayAlerts = False
    wb = app.Workbooks.Open(DST, 0, False)
    ws = wb.Worksheets('计算模板')

    def dump(tag):
        vals = ws.Range(ws.Cells(195, 6), ws.Cells(202, 7)).Value2
        print(f'  [{tag}] 行195~202 F/G:')
        for i, row in enumerate(vals, start=195):
            print(f'    行{i}: F={row[0]!r} G={row[1]!r}')

    dump('原始')
    # 复刻 normalize_com_times
    def normalize(wsx, cols, last_row):
        fixed = 0
        for c in cols:
            if last_row < 2: continue
            vals = wsx.Range(wsx.Cells(2, c), wsx.Cells(last_row, c)).Value2
            if vals is None: continue
            if not isinstance(vals, tuple): vals = ((vals,),)
            out = []
            changed = False
            for row in vals:
                v = row[0] if isinstance(row, tuple) else row
                nv = v
                if isinstance(v, str):
                    fixed += 1
                out.append((nv,) if isinstance(row, tuple) else nv)
            print(f'    列{c}: 读回 {len(vals)} 行, out 元素示例 {out[:3]}')
            # 原脚本只在 changed 时写回; 这里强制写回以放大问题
            wsx.Range(wsx.Cells(2, c), wsx.Cells(last_row, c)).Value = tuple(out)
        return fixed

    normalize(ws, (6, 7), 202)
    dump('normalize 后')
    wb.Close(False)
finally:
    if app is not None:
        try: app.Quit()
        except Exception: pass
    pythoncom.CoUninitialize()
