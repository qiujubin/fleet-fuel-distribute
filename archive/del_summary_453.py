# -*- coding: utf-8 -*-
"""汇总: 删除 453 的旧行(京LPW138, 时间A=09-20 03:04:50 时间B=09-21 08:14:17), 重排序号"""
import pythoncom
import win32com.client as wc

SUM = r'C:\Users\Jubin\Desktop\源数据\油耗分析汇总.xlsm'
pythoncom.CoInitialize()
app = None
try:
    app = wc.DispatchEx('Ket.Application')
    app.Visible = False; app.DisplayAlerts = False
    wbs = app.Workbooks.Open(SUM, 0, False)
    assert not wbs.ReadOnly
    wss = wbs.Worksheets('Sheet1')
    last = wss.Cells(wss.Rows.Count, 3).End(-4162).Row
    del_row = None
    for i in range(2, last + 1):
        if str(wss.Cells(i, 3).Value2).strip() != '京LPW138': continue
        ta, tb = wss.Cells(i, 6).Value2, wss.Cells(i, 7).Value2
        if ta is None or tb is None: continue
        try:
            ta_f, tb_f = float(ta), float(tb)
        except (TypeError, ValueError):
            continue
        if abs(ta_f - 46285.1283564815) < 0.0001 and abs(tb_f - 46286.3432523148) < 0.0001:
            del_row = i; break
    if del_row:
        wss.Rows(del_row).Delete()
        last2 = wss.Cells(wss.Rows.Count, 3).End(-4162).Row
        for i in range(del_row, last2 + 1):
            wss.Cells(i, 1).Value2 = i - 1
        print(f'汇总已删除 453 旧行(原行 {del_row}), 序号重排至 {last2}, 当前末行 {last2}')
    else:
        print('未找到 453 的旧行')
    wbs.Save(); wbs.Close(False)
finally:
    if app is not None:
        try: app.Quit()
        except Exception: pass
    pythoncom.CoUninitialize()
print('完成')
