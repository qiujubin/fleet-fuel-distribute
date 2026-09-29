# -*- coding: utf-8 -*-
"""修复 fix_d15_tz.py 写坏的那一行(车牌为空 + 时间偏 -8h)
正确值从 D15 模板行151 用 Value2(序列数) 重读, 全程使用序列数避免时区转换"""
import os, datetime
import pythoncom, win32com.client as wc
from openpyxl.utils.datetime import from_excel

D15 = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板\D15-4.2米-粤BY2J27.xlsx'
KAOHE = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
DT_FMT = 'yyyy-mm-dd hh:mm:ss'

pythoncom.CoInitialize()
app = None
try:
    app = wc.DispatchEx('Ket.Application')
    app.Visible = False
    app.DisplayAlerts = False

    # ---- 1) 从模板读行151 的正确值(Value2 = 序列数, 无时区风险) ----
    wbv = app.Workbooks.Open(D15, 0, False)
    wsv = wbv.Worksheets('计算模板')
    app.Calculate()
    rng = wsv.Range(wsv.Cells(151, 6), wsv.Cells(151, 22)).Value2
    if not isinstance(rng, tuple):
        rng = (rng,)
    if rng and isinstance(rng[0], tuple):
        rng = rng[0]
    plate_v = str(wsv.Cells(2, 4).Value2 or '').strip()
    type_v = str(wsv.Cells(2, 5).Value2 or '').strip()
    driver_v = wsv.Cells(151, 3).Value2
    print(f'模板行151: 车牌={plate_v} 类型={type_v} 司机={driver_v}')
    print('F..V:', [round(x, 4) if isinstance(x, float) else x for x in rng])
    wbv.Close(False)

    def n_(x):
        try: return float(x)
        except (TypeError, ValueError): return None

    F_s, G_s = n_(rng[0]), n_(rng[1])
    print(f'F(序列数)={F_s} -> {from_excel(F_s)} | G(序列数)={G_s} -> {from_excel(G_s)}')
    assert F_s and G_s, '模板行151 的 F/G 为空'

    # ---- 2) 定位坏行: 车牌为空且金额≈832.9 ----
    wbk = app.Workbooks.Open(KAOHE, 0, False)
    wsd = wbk.Worksheets('每日油耗数据')
    target = None
    for i in range(2, 4000):
        p = str(wsd.Cells(i, 4).Value2 or '').strip()
        if p:
            continue
        amt = n_(wsd.Cells(i, 9).Value2)
        if amt is not None and abs(amt - 832.9) < 0.5:
            target = i
    if not target:
        print('!!! 未找到坏行')
        wbk.Close(False)
    else:
        print(f'找到坏行: {target}')
        print('  修复前:', [wsd.Cells(target, c).Value2 for c in (1, 4, 5, 6, 7, 9)])

        # ---- 3) 重写(序列数) ----
        rowvals = [None, target - 1, driver_v, plate_v, type_v,
                   F_s, G_s, '是',
                   n_(rng[3]), n_(rng[4]), n_(rng[5]),
                   n_(rng[6]), n_(rng[7]), n_(rng[8]), n_(rng[9]),
                   n_(rng[10]), n_(rng[11]), n_(rng[12]), n_(rng[13]), n_(rng[14]),
                   n_(rng[15]), n_(rng[16]),
                   float(int(G_s)), 0]
        # A 列(sid)从导出记录取: 粤BY2J27 09-25 10:38
        rowvals[0] = wsd.Cells(target, 1).Value2 or '260925000544'
        wsd.Cells(target, 1).NumberFormat = '@'
        wsd.Range(wsd.Cells(target, 1), wsd.Cells(target, 24)).Value = tuple([tuple(rowvals)])
        for c in (6, 7):
            wsd.Cells(target, c).NumberFormat = DT_FMT
        wsd.Cells(target, 23).NumberFormat = 'yyyy-mm-dd'
        wbk.Save()

        # ---- 4) 读回验证 ----
        chk = [wsd.Cells(target, c).Value2 for c in (1, 4, 5, 6, 7, 9, 10, 16, 17)]
        print('  修复后:', chk)
        print(f'    车牌={chk[1]} 类型={chk[2]}')
        print(f'    时间A={from_excel(chk[3]) if isinstance(chk[3],float) else chk[3]}')
        print(f'    时间B={from_excel(chk[4]) if isinstance(chk[4],float) else chk[4]}')
        wbk.Close(False)
finally:
    if app is not None:
        try: app.Quit()
        except Exception: pass
    pythoncom.CoUninitialize()
print('完成')
