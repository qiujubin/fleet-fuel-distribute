# -*- coding: utf-8 -*-
"""隔离测试改进版宏: 单次调用, flush 输出定位卡点, 验证去重+末行定位"""
import pythoncom, shutil, os, sys
import win32com.client as wc

TE = r'E:\WorkSpace\workbuddy\2026-09-18-16-38-15\test_env5'
K = os.path.join(TE, '考核测试.xlsm')
S = os.path.join(TE, '汇总测试.xlsm')
print('1) 初始化 COM...', flush=True)
pythoncom.CoInitialize()
app = None
try:
    app = wc.DispatchEx('Ket.Application')
    app.Visible = False
    app.DisplayAlerts = False
    print('2) COM 就绪, 打开考核表...', flush=True)
    wb = app.Workbooks.Open(K, 0, False)
    wb.Activate()
    print('3) 调用宏 同步油耗到汇总 (第1次)...', flush=True)
    ret = app.Run('同步油耗到汇总')
    print('   返回:', ret, flush=True)
    wb.Close(False)
    print('4) 第2次调用 (验证去重)...', flush=True)
    wb = app.Workbooks.Open(K, 0, False)
    wb.Activate()
    ret2 = app.Run('同步油耗到汇总')
    print('   返回:', ret2, flush=True)
    wb.Close(False)
    print('5) 检查测试汇总...', flush=True)
finally:
    if app is not None:
        try: app.Quit()
        except Exception: pass
    pythoncom.CoUninitialize()

from collections import Counter
from python_calamine import CalamineWorkbook
cwb = CalamineWorkbook.from_path(S)
rows = cwb.get_sheet_by_name('Sheet1').to_python(skip_empty_area=False)
keys, n = [], 0
for row in rows[1:]:
    if row and any(c is not None for c in row):
        n += 1
        keys.append('|'.join(str(row[i])[:16] for i in (2, 5, 6, 7, 13) if i < len(row)))
dup = {k: c for k, c in Counter(keys).items() if c > 1}
print(f'测试汇总数据行: {n} | 重复键: {len(dup)}', flush=True)
print('完成', flush=True)
