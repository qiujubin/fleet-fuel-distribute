# -*- coding: utf-8 -*-
"""全车队骨架余量补齐: 余量(骨架末行 - 最后已填G行)不足 15 行的, 扩充到 15 行。
每消耗一行补一行的策略由主脚本日常维护, 本脚本做一次性盘点补齐。

⚠️ 2026-09-28 用户明令禁止 openpyxl 保存文件 —— 本脚本原用 openpyxl 写回车辆模板，
   现改为 WPS COM（整行 Copy 骨架末行，公式由 WPS 自动偏移，不手工改公式文本）。
"""
import os, glob, atexit

TPL = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板'
FLOOR = 15

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
        app = wc.DispatchEx(pid); break
    except Exception:
        continue
if app is None:
    raise SystemExit('无法启动 COM')
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
            if v not in (None, ''):
                return start + i
        end = start - 1
    return 0


def tpl_sheet(wb):
    for sn in ('计算模板', '油耗计算'):
        try:
            return wb.Worksheets(sn)
        except Exception:
            continue
    return None


done = []
for fp in sorted(glob.glob(os.path.join(TPL, '*.xlsx'))):
    fn = os.path.basename(fp)
    if fn.startswith('~$') or fn.startswith('模板'): continue
    wb = app.Workbooks.Open(fp, 0, False)
    if wb.ReadOnly:
        print(f'[跳过] {fn}: 被占用(只读)'); continue
    opened.append(wb)
    try:
        ws = tpl_sheet(wb)
        if ws is None:
            print(f'[跳过] {fn}: 无计算模板sheet')
            continue
        last_filled = true_last_row(ws, 7)
        skel_last = true_last_row(ws, 2)
        if skel_last < last_filled:
            skel_last = last_filled
        free = skel_last - last_filled
        if free >= FLOOR:
            continue
        plate_v = str(ws.Cells(2, 4).Value or '')
        type_v = str(ws.Cells(2, 5).Value or '')
        need = FLOOR - free
        src = skel_last
        for i in range(1, need + 1):
            r = src + i
            ws.Rows(src).Copy(ws.Rows(r))       # 整行复制, 公式由 WPS 自动偏移
            ws.Cells(r, 2).Value = r - 1
            ws.Cells(r, 4).Value = plate_v
            ws.Cells(r, 5).Value = type_v
        wb.Save()
        done.append(f'{fn}: 余量{free} -> {FLOOR} (补{need}行, 已填至{last_filled})')
    finally:
        try: wb.Close(False)
        except Exception: pass
        if wb in opened: opened.remove(wb)

print(f'共补齐 {len(done)} 个文件:')
for d in done: print('  ', d)
