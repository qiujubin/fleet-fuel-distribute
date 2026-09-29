# -*- coding: utf-8 -*-
"""修 B7/C6: ①加油数据 M/C 列字符串时间规范化为序列数 ②重建公式 ③COM 重算 ④更新考核行 L/M/P..V"""
import sys, os, datetime
sys.path.insert(0, r'E:\WorkSpace\workbuddy\2026-09-18-16-38-15')
import openpyxl
from rebuild_formulas import rebuild

TPL = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板'
KAOHE = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
EPOCH = datetime.datetime(1899, 12, 30)

def to_serial(v):
    if v is None: return None
    if isinstance(v, datetime.datetime): return (v - EPOCH).total_seconds() / 86400.0
    if isinstance(v, datetime.date): return (datetime.datetime(v.year, v.month, v.day) - EPOCH).total_seconds() / 86400.0
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip()
    for f in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d'):
        try: return (datetime.datetime.strptime(s, f) - EPOCH).total_seconds() / 86400.0
        except ValueError: pass
    return v

FILES = [
    ('B7-9.6米-京AEP303.xlsx', '260921000455', 3343),
    ('C6-7.6米-粤BDW641.xlsx', '260921000454', 3344),
]
latest = {}
for fn, sid, kao_row in FILES:
    p = os.path.join(TPL, fn)
    wb = openpyxl.load_workbook(p, data_only=True)
    wsd = wb['加油数据']
    # 找锚点行(M 字符串) 与最新有效是行
    anchor_serial = None
    for r in range(2, wsd.max_row + 1):
        t = wsd.cell(r, 13).value
        if t is not None and str(to_serial(t))[:10] != 'None':
            pass
    wb.close()
    # 用 data_only=False 读 M 列原值并规范化
    wb = openpyxl.load_workbook(p, data_only=False)
    wsd = wb['加油数据']
    n_fix = 0
    for r in range(2, wsd.max_row + 1):
        v = wsd.cell(r, 13).value
        if v is not None and not hasattr(v, 'year') and not isinstance(v, (int, float)):
            srl = to_serial(v)
            if isinstance(srl, float):
                wsd.cell(r, 13).value = srl
                wsd.cell(r, 13).number_format = 'yyyy-mm-dd hh:mm:ss'
                n_fix += 1
        d = wsd.cell(r, 3).value
        if d is not None and not hasattr(d, 'year') and not isinstance(d, (int, float)):
            srl = to_serial(d)
            if isinstance(srl, float):
                wsd.cell(r, 3).value = srl
                wsd.cell(r, 3).number_format = 'yyyy-mm-dd'
                n_fix += 1
    wb.save(p); wb.close()
    print(f'{fn[:2]}: 加油数据时间规范化 {n_fix} 格')
    # 重建公式(模板 F/G 若为字符串, rebuild 的锚点匹配用字符串对比, 不受影响;
    # 但模板 G 列若是字符串, 与加油数据 M(序列数)的 XLOOKUP 匹配失败 → 模板 G 也要规范化)
    wb = openpyxl.load_workbook(p, data_only=False)
    wst = wb['计算模板']
    n_fix2 = 0
    for r in range(2, wst.max_row + 1):
        for c in (6, 7):
            v = wst.cell(r, c).value
            if v is not None and not hasattr(v, 'year') and not isinstance(v, (int, float)):
                import re as _re
                m = _re.match(r'(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})', str(v).strip())
                if m:
                    dt = datetime.datetime(*map(int, m.groups()))
                    wst.cell(r, c).value = (dt - EPOCH).total_seconds() / 86400.0
                    wst.cell(r, c).number_format = 'yyyy-mm-dd hh:mm:ss'
                    n_fix2 += 1
    wb.save(p); wb.close()
    print(f'{fn[:2]}: 模板 F/G 规范化 {n_fix2} 格')
    n2, filled = rebuild(p)
    print(f'{fn[:2]}: rebuild {n2} 公式 / {filled} 行G')

# ---- COM: 重算并读最新有效是行实时值, 更新考核行 ----
import pythoncom
import win32com.client as wc
pythoncom.CoInitialize()
app = None
try:
    app = wc.DispatchEx('Ket.Application')
    app.Visible = False; app.DisplayAlerts = False
    wbk = app.Workbooks.Open(KAOHE, 0, False)
    assert not wbk.ReadOnly
    wsd2 = wbk.Worksheets('每日油耗数据')
    for fn, sid, kao_row in FILES:
        p = os.path.join(TPL, fn)
        w2 = app.Workbooks.Open(p, 0, False)
        wst2 = w2.Worksheets('计算模板')
        app.Calculate()
        # 最新有效是行 = G 列最后非空行
        XL_UP = -4162
        last_g = wst2.Range('G' + str(wst2.Rows.Count)).End(XL_UP).Row
        rng = wst2.Range(wst2.Cells(last_g, 6), wst2.Cells(last_g, 22)).Value2
        vals = rng[0] if isinstance(rng[0], tuple) else rng
        def n(v):
            try: return float(v)
            except (TypeError, ValueError): return None
        print(f'{fn[:2]} 模板行{last_g}: F={vals[0]} G={vals[1]} I={vals[3]} J={vals[4]} L={vals[6]} M={vals[7]} P={vals[10]} Q={vals[11]}')
        # 更新考核行: L/M/N/O(col12~15) P/Q(col16/17) R/S/T/U/V(col18~22) I/J(col9/10)
        upd = {9: n(vals[3]), 10: n(vals[4]), 12: n(vals[6]), 13: n(vals[7]),
               14: n(vals[8]), 15: n(vals[9]), 16: n(vals[10]), 17: n(vals[11]),
               18: n(vals[12]), 19: n(vals[13]), 20: n(vals[14]), 21: n(vals[15]), 22: n(vals[16])}
        for c, v in upd.items():
            wsd2.Cells(kao_row, c).Value = v
        print(f'  考核行{kao_row} 已更新 (L={upd[12]} M={upd[13]} P={upd[16]} Q={upd[17]})')
        w2.Close(False)
    wbk.Save()
    print('考核表已保存')
    wbk.Close(False)
finally:
    if app is not None:
        try: app.Quit()
        except Exception: pass
    pythoncom.CoUninitialize()
print('完成')
