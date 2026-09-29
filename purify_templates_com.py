# -*- coding: utf-8 -*-
"""把被 openpyxl 写过的车辆模板用 WPS 原生「打开→重算→保存」洗回标准格式。

背景: rebuild_formulas / retro_fix 历史上用 openpyxl 保存过这几个文件
      (2026-09-28 用户明令: 写入一律只准 win32/COM)。openpyxl 会重新序列化
      styles.xml / sheetN.xml, 虽未丢部件, 但格式应回归 WPS 原生。

用法: python purify_templates_com.py 文件名1 [文件名2 ...]
"""
import os, sys, atexit, datetime

TPL = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板'
TARGETS = sys.argv[1:] or [
    'C3-7.6米-粤BPV550.xlsx',
    'D15-4.2米-粤BY2J27.xlsx',
    'D17-4.2米-京LPW138.xlsx',
]

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
if app is None:
    raise SystemExit('无法启动 COM')
app.Visible = False
app.DisplayAlerts = False
try: app.ScreenUpdating = False
except Exception: pass

for fn in TARGETS:
    fp = os.path.join(TPL, fn)
    if not os.path.exists(fp):
        print('  跳过(不存在):', fn); continue
    before = os.path.getsize(fp)
    wb = app.Workbooks.Open(fp, 0, False)
    if wb.ReadOnly:
        print(f'  !! {fn} 被占用(只读), 跳过'); continue
    opened.append(wb)
    try:
        app.Calculate()
    except Exception:
        pass
    wb.Save()
    opened.remove(wb)
    wb.Close(False)
    after = os.path.getsize(fp)
    print(f'  {fn}: {before} -> {after} bytes  (WPS 原生保存)')

print('\n完成')
