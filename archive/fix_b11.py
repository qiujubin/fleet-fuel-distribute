# -*- coding: utf-8 -*-
def sid_key(v):
    s = str(v).strip()
    try:
        f = float(s)
        if f == int(f): return str(int(f))
    except ValueError:
        pass
    return s
"""修 B11: 删除重复的 440(第187行), A186 的数字型 sid 转回文本"""
import pythoncom
import win32com.client as wc

P = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板\B11-9.6米-粤BFE485.xlsx'
pythoncom.CoInitialize()
app = None
try:
    app = wc.DispatchEx('Ket.Application')
    app.Visible = False; app.DisplayAlerts = False
    wb = app.Workbooks.Open(P, 0, False)
    assert not wb.ReadOnly
    ws = wb.Worksheets('加油数据')
    # 确认 187 是重复 440
    a186, a187 = ws.Cells(186, 1).Value2, ws.Cells(187, 1).Value2
    assert sid_key(a186) == '260920000440' and sid_key(a187) == '260920000440', (a186, a187)
    ws.Rows(187).Delete()
    # A186 数字转文本
    ws.Cells(186, 1).NumberFormat = '@'
    ws.Cells(186, 1).Value = '260920000440'
    wb.Save()
    print('B11 已修复: 删187重复行, A186 转文本')
    wb.Close(False)
finally:
    if app is not None:
        try: app.Quit()
        except Exception: pass
    pythoncom.CoUninitialize()
