# -*- coding: utf-8 -*-
"""补丁: 删除 尿素D14-4.2米-粤BT802F.xlsx「计算模板」里 260928000587 的那一行。

背景: fix_bt802f_0928.py 删「加油数据」行后, WPS UsedRange 缓存导致该行 A 列
      读回 0(公式缓存被清), 按系统编号定位失效 -> 未删。改为按 G 列(加油时间B)末行定位。

用法: python fix_bt802f_0928b.py
"""
import os, atexit, datetime
from python_calamine import CalamineWorkbook

UREA_F = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板\尿素D14-4.2米-粤BT802F.xlsx'
TARGET_DAY = datetime.date(2026, 9, 28)
PLATE = '粤BT802F'

import pythoncom
import win32com.client as wc

pythoncom.CoInitialize()
app = None
_opened = []


def cleanup():
    global app
    try:
        if app is None: return
        for wb in list(_opened):
            try: wb.Close(False)
            except Exception: pass
        _opened.clear()
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


def true_last_row(ws, col=1):
    ur = ws.UsedRange
    end = min(ur.Row + ur.Rows.Count - 1, ws.Rows.Count)
    while end >= 1:
        start = max(1, end - 1999)
        vals = ws.Range(ws.Cells(start, col), ws.Cells(end, col)).Value2
        if vals is None:
            end = start - 1; continue
        if not isinstance(vals, tuple): vals = ((vals,),)
        for i in range(len(vals) - 1, -1, -1):
            v = vals[i][0] if isinstance(vals[i], tuple) else vals[i]
            if v is not None and str(v).strip() != '':
                return start + i
        end = start - 1
    return 0


def from_serial(v):
    if v is None: return None
    if isinstance(v, datetime.datetime): return v
    from openpyxl.utils.datetime import from_excel
    return from_excel(float(v))


wb = app.Workbooks.Open(UREA_F, 0, False)
if wb.ReadOnly:
    raise SystemExit('文件被占用')
_opened.append(wb)
wst = None
for sn in ('计算模板', '油耗计算'):
    try: wst = wb.Worksheets(sn); break
    except Exception: continue

rt = true_last_row(wst, 7)
print(f'计算模板 G 列最后有值行 = {rt}')
gv = wst.Cells(rt, 7).Value2
pl = str(wst.Cells(rt, 4).Value or '')
dt = from_serial(gv)
print(f'  车牌=[{pl}] 加油时间B={dt}')

if pl.strip().upper() == PLATE and dt and dt.date() == TARGET_DAY:
    # 记录公式现场(删除前后对比 #REF!)
    last_before = true_last_row(wst, 2)
    wst.Rows(rt).Delete()
    print(f'  已删除行 {rt}')
    # 序号重排: B 列 = 行号 - 1
    last_after = true_last_row(wst, 2)
    for r in range(2, last_after + 1):
        wst.Cells(r, 2).Value = r - 1
    try: app.Calculate()
    except Exception: pass
    # 校验: 检查 I/J 列公式是否出现 #REF!
    bad = []
    for r in range(2, last_after + 1):
        for c, cl in ((9, 'I'), (10, 'J')):
            try:
                f = wst.Cells(r, c).Formula
            except Exception:
                continue
            if isinstance(f, str) and '#REF!' in f:
                bad.append((r, cl, f[:60]))
    print(f'  I/J 公式 #REF! 检查: {len(bad)} 处')
    for b in bad[:10]: print('   ', b)
    wb.Save()
    print('  已保存')
else:
    print('  未匹配目标, 不操作')
try: wb.Close(False)
except Exception: pass

print('\n--- 复查 ---')
wbc = CalamineWorkbook.from_path(UREA_F)
for sn in ['加油数据', '计算模板']:
    rows = wbc.get_sheet_by_name(sn).to_python()
    print(f'[{sn}] rows={len(rows)}')
print('尾部 4 行 (计算模板):')
rows = wbc.get_sheet_by_name('计算模板').to_python()
for i, row in enumerate(rows, 1):
    if i > len(rows) - 4:
        print(' ', i, '|', '|'.join('' if x is None else str(x)[:20] for x in row[:11]))
