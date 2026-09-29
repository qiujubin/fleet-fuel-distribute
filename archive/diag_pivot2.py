# -*- coding: utf-8 -*-
"""诊断 v2: 「每日油耗」完整内容 + 日期字段项 + 考核表今日行 L/M"""
import pythoncom
import win32com.client as wc
pythoncom.CoInitialize()
app = None
try:
    app = wc.DispatchEx('Ket.Application')
    app.Visible = False; app.DisplayAlerts = False
    wb2 = app.Workbooks.Open(r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm', 0, True)
    ws2 = wb2.Worksheets('每日油耗')
    print('=== 「每日油耗」A1:K16 完整内容:')
    vr = ws2.Range('A1:K16').Value2
    for i, row in enumerate(vr, start=1):
        print(f'  行{i}:', [str(x)[:16] if x is not None else '' for x in row])
    pt = ws2.PivotTables(1)
    pf = pt.PivotFields('日期')
    print(f'\n=== 「日期」字段: orientation={pf.Orientation}, 项数={pf.PivotItems().Count}')
    for j in range(1, pf.PivotItems().Count + 1):
        it = pf.PivotItems(j)
        print(f'  项{j}: {it.Name!r} Visible={it.Visible}')
    # 考核表今日行 L/M
    ws3 = wb2.Worksheets('每日油耗数据')
    print('\n=== 考核表今日考核行的 L/M (col12/13):')
    targets = {'260920000440', '260920000441', '260920000443', '260920000448',
               '260921000449', '260921000451', '260921000453', '260921000454', '260921000455'}
    for i in range(3300, 3360):
        v = ws3.Cells(i, 1).Value2
        if v is None: continue
        sv = str(v).strip()
        if sv.endswith('.0'): sv = sv[:-2]
        if sv in targets:
            print(f'  行{i}: {sv} 车牌={ws3.Cells(i,4).Value2} | F(时间A)={str(ws3.Cells(i,6).Value2)[:19]} '
                  f'| L(GPS起)={ws3.Cells(i,12).Value2!r} M(仪表起)={ws3.Cells(i,13).Value2!r} '
                  f'| N(GPS终)={ws3.Cells(i,14).Value2!r} P(GPS周期)={ws3.Cells(i,16).Value2!r}')
    wb2.Close(False)
finally:
    if app is not None:
        try: app.Quit()
        except Exception: pass
    pythoncom.CoUninitialize()
