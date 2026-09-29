# -*- coding: utf-8 -*-
"""单独重跑宏「同步油耗到汇总」（前一次调用因 ActiveWorkbook 状态异常失败）"""
import os, atexit

KAOHE = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'

import pythoncom
import win32com.client as wc

pythoncom.CoInitialize()
app = None
opened = []


def cleanup():
    global app
    try:
        if app is None: return
        for wb in list(opened):
            try: wb.Close(False)
            except Exception: pass
        opened.clear()
        try: app.Quit()
        except Exception: pass
        app = None
    except Exception:
        pass


atexit.register(cleanup)

for pid in ('Ket.Application', 'et.Application', 'Excel.Application'):
    try:
        app = wc.DispatchEx(pid); print('COM 应用:', pid); break
    except Exception:
        continue
app.Visible = False
app.DisplayAlerts = False
try: app.ScreenUpdating = False
except Exception: pass

wb = app.Workbooks.Open(KAOHE, 0, False)
if wb.ReadOnly:
    raise SystemExit('考核表被占用(只读)')
opened.append(wb)
wb.Activate()
try: app.Calculate()
except Exception: pass
print('调用宏...')
try:
    ret = app.Run('同步油耗到汇总')
    print('  返回:', ret)
except Exception as e:
    print('  宏失败:', e)

# 区域填充兜底
try:
    import sys
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from region_map import REGION_MAP
    before = set(w.Name for w in app.Workbooks)
except Exception:
    before = set()

for w2 in [w for w in app.Workbooks if w.Name not in before and '汇总' in w.Name]:
    try:
        p2r = {}
        for reg, plates in REGION_MAP.items():
            for p in plates: p2r[p.strip().upper()] = reg
        ws3 = w2.Worksheets('Sheet1')
        ur = ws3.UsedRange
        last2 = ur.Row + ur.Rows.Count - 1
        vals = ws3.Range(ws3.Cells(2, 3), ws3.Cells(last2, 3)).Value2
        if vals is None: vals = ()
        if not isinstance(vals, tuple): vals = ((vals,),)
        filled = 0
        for i2, cv in enumerate(vals):
            r2 = 2 + i2
            pv2 = cv[0] if isinstance(cv, tuple) else cv
            pl = str(pv2 or '').strip().upper()
            if not pl: continue
            if ws3.Cells(r2, 2).Value2 not in (None, ''): continue
            reg = p2r.get(pl)
            if reg:
                ws3.Cells(r2, 2).Value = reg; filled += 1
        if filled: w2.Save()
        print(f'  区域填充补 {filled} 行 ({w2.Name})')
        w2.Close(True)
    except Exception as e2:
        print('  收尾告警:', e2)

wb.Close(False)
if wb in opened: opened.remove(wb)
print('完成')
