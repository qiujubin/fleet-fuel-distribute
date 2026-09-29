# -*- coding: utf-8 -*-
"""修 D17 京LPW138: 453(09-21 08:14 是) 降级, 466(09-21 21:18 是) 升为有效是
①模板 ②rebuild ③COM重算 ④考核行替换 ⑤汇总删除453旧行+重排序号"""
import sys, os, datetime
sys.path.insert(0, r'E:\WorkSpace\workbuddy\2026-09-18-16-38-15')
import openpyxl
from rebuild_formulas import rebuild

TPL = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板'
KAOHE = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
SUM = r'C:\Users\Jubin\Desktop\源数据\油耗分析汇总.xlsm'
D17 = os.path.join(TPL, 'D17-4.2米-京LPW138.xlsx')

print('模板/考核表步骤已完成, 只重跑汇总删除')

import pythoncom
import win32com.client as wc
pythoncom.CoInitialize()
app = None
try:
    app = wc.DispatchEx('Ket.Application')
    app.Visible = False; app.DisplayAlerts = False

    v = [46285.1283564815, 46286.888125, '是', 907.69, 114.35, 7.95, 205388.6, 200150.0, 206055.8, 200794.0, 667.2, 644.0]

    wbk = None
    if False:
        wbk = app.Workbooks.Open(KAOHE, 0, False)
    wsd = wbk.Worksheets('每日油耗数据')
    target = None
    for i in range(2, 6000):
        x = wsd.Cells(i, 1).Value2
        if x is not None and str(x).strip() in ('260921000453', '260921000453.0'):
            target = i; break
    assert target, '找不到 453 考核行'
    def num_(x):
        try: return float(x)
        except (TypeError, ValueError): return None
    rowvals = ['260921000466', wsd.Cells(target, 2).Value2, wsd.Cells(target, 3).Value2,
               wsd.Cells(target, 4).Value2, wsd.Cells(target, 5).Value2,
               v[0], v[1], '是', num_(v[3]), num_(v[4]), wsd.Cells(target, 11).Value2,
               num_(v[6]), num_(v[7]), num_(v[8]), num_(v[9]), num_(v[10]), num_(v[11]),
               num_(v[12]), num_(v[13]), num_(v[14]), num_(v[15]), num_(v[16]),
               float(int(v[1])), 0]
    wsd.Cells(target, 1).NumberFormat = '@'
    wsd.Range(wsd.Cells(target, 1), wsd.Cells(target, 24)).Value = tuple([tuple(rowvals)])
    for c in (6, 7): wsd.Cells(target, c).NumberFormat = 'yyyy-mm-dd hh:mm:ss'
    wsd.Cells(target, 23).NumberFormat = 'yyyy-mm-dd'
    print(f'考核行 {target} 已替换为 466 (I={num_(v[3])} J={num_(v[4])})')
    wbk.Save(); wbk.Close(False)

    # ---- 4) 汇总: 删除 453 旧行(时间A=09-20 03:04:50 时间B=09-21 08:14:17), 重排序号 ----
    wbs = app.Workbooks.Open(SUM, 0, False)
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
            continue   # 字符串日期行, 跳过
        if abs(ta_f - 46285.1283564815) < 0.0001 and abs(tb_f - 46286.3432523148) < 0.0001:
            del_row = i; break
    if del_row:
        wss.Rows(del_row).Delete()
        # 重排序号
        last2 = wss.Cells(wss.Rows.Count, 3).End(-4162).Row
        for i in range(del_row, last2 + 1):
            wss.Cells(i, 1).Value2 = i - 1
        print(f'汇总已删除 453 旧行(原行{del_row}), 序号重排至 {last2}')
    else:
        print('汇总里未找到 453 的旧行(可能未被同步过)')
    wbs.Save(); wbs.Close(False)
finally:
    if app is not None:
        try: app.Quit()
        except Exception: pass
    pythoncom.CoUninitialize()
print('完成 (466 的新行将由下次宏同步写入汇总)')
