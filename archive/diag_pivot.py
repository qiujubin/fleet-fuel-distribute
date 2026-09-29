# -*- coding: utf-8 -*-
"""诊断: ①汇总里 09-21 的行与空起始字段 ②「每日油耗」透视表结构 ③考核表今日行 L/M"""
import pythoncom
import win32com.client as wc
pythoncom.CoInitialize()
app = None
try:
    app = wc.DispatchEx('Ket.Application')
    app.Visible = False; app.DisplayAlerts = False

    # ---- 1) 汇总: 09-21 的行 ----
    wb1 = app.Workbooks.Open(r'C:\Users\Jubin\Desktop\源数据\油耗分析汇总.xlsm', 0, True)
    ws = wb1.Worksheets('Sheet1')
    ur = ws.UsedRange
    last = ur.Row + ur.Rows.Count - 1
    print('=== 汇总 Sheet1 (末行', last, ') 09-20/09-21 的行:')
    vals = ws.Range(ws.Cells(max(2, last - 30), 1), ws.Cells(last, 20)).Value2
    for i, row in enumerate(vals):
        rr = max(2, last - 30) + i
        g = row[6] if len(row) > 6 else None  # G=加油时间B (col7)
        gs = str(g)[:10] if g else ''
        if gs in ('2026-09-20', '2026-09-21'):
            print(f'  行{rr}: 序号={row[0]} 车牌={row[2]} 时间A={str(row[5])[:19]} 时间B={str(row[6])[:19]} '
                  f'金额={row[7]} 升={row[8]} 里程J={row[9]!r} GPS油耗K={row[10]!r}')
    wb1.Close(False)

    # ---- 2) 考核表: 「每日油耗」透视表结构 + 今日考核行 L/M ----
    wb2 = app.Workbooks.Open(r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm', 0, True)
    ws2 = wb2.Worksheets('每日油耗')
    pts = ws2.PivotTables()
    print(f'\n=== 「每日油耗」透视表数量: {pts.Count}')
    for i in range(1, pts.Count + 1):
        pt = pts.Item(i)
        print(f'  透视表{i}: 名称={pt.Name} 源={pt.SourceData if hasattr(pt,"SourceData") else "?"}')
        print(f'    位置: {pt.TableRange1.Address}')
        # 页字段/筛选字段
        for j in range(1, pt.PivotFields().Count + 1):
            pf = pt.PivotFields(j)
            orient = pf.Orientation  # 1行 2列 3页 4数据 0隐藏
            if orient in (3, 2, 1):
                cur = ''
                try: cur = pf.CurrentPage
                except Exception: pass
                print(f'    字段[{pf.Name}] orientation={orient} CurrentPage={cur!r}')
    # 透视表数据区域内容
    tr = pts.Item(1).TableRange1
    print('  数据区前几行:')
    vr = tr.Resize(min(tr.Rows.Count, 8), min(tr.Columns.Count, 10)).Value2
    if tr.Rows.Count == 1:
        print('   ', vr)
    else:
        for r_ in vr: print('   ', [str(x)[:14] if x is not None else '' for x in r_])

    # 考核表今日行 L/M (col12/13)
    ws3 = wb2.Worksheets('每日油耗数据')
    print('\n=== 考核表今日行的 L/M (起始GPS/仪表):')
    for i in range(3330, 3350):
        v = ws3.Cells(i, 1).Value2
        if v is not None and str(v).strip() in ('260920000440', '260920000441', '260920000443',
                                                 '260920000448', '260921000451', '260921000453',
                                                 '260921000454', '260921000455', '260921000449'):
            print(f'  行{i}: {str(v).strip()} 车牌={ws3.Cells(i,4).Value2} L(GPS起)={ws3.Cells(i,12).Value2!r} '
                  f'M(仪表起)={ws3.Cells(i,13).Value2!r} N={ws3.Cells(i,14).Value2!r} O={ws3.Cells(i,15).Value2!r}')
    wb2.Close(False)
finally:
    if app is not None:
        try: app.Quit()
        except Exception: pass
    pythoncom.CoUninitialize()
