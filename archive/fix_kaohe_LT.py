# -*- coding: utf-8 -*-
"""补全考核表中 415/424 两行的 L~T 列 (首次运行时 recF 用了旧缓存导致为空)"""
import os, datetime
import openpyxl

TPL = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板'
KAOHE = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'

def get_rec(path, sid):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['加油数据']
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(r, 1).value or '').strip() == sid:
            return {'mile': ws.cell(r, 4).value, 'gps': ws.cell(r, 5).value,
                    'amt': ws.cell(r, 6).value, 'vol': ws.cell(r, 7).value,
                    'price': ws.cell(r, 8).value, 'person': ws.cell(r, 12).value}
    return None

def num(v):
    try: return float(v)
    except: return None

targets = [
    # (考核sid, F基准sid, 模板文件, 新记录sid, F时间)
    ('260918000415', '260918000408', 'D14-4.2米-粤BT802F.xlsx'),
    ('260919000424', '260918000409', 'B13-9.6米-粤BNT993.xlsx'),
]
wbk = openpyxl.load_workbook(KAOHE, keep_vba=True)
w1 = wbk['每日油耗数据']
def find_row(ws, sid):
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(r, 1).value or '').strip() == sid: return r
for sid_new, sid_f, fn in targets:
    rec_n = get_rec(os.path.join(TPL, fn), sid_new)
    rec_f = get_rec(os.path.join(TPL, fn), sid_f)
    L, M = num(rec_f['gps']), num(rec_f['mile'])
    N, O = num(rec_n['gps']), num(rec_n['mile'])
    I, J = num(rec_n['amt']), num(rec_n['vol'])
    P = round(N - L, 4) if L is not None and N is not None else None
    Q = round(O - M, 4) if M is not None and O is not None else None
    Rv = round(J / P * 100, 6) if P else None
    Sv = round(I / P, 6) if P else None
    Tv = round(J / Q * 100, 6) if Q else None
    rr = find_row(w1, sid_new)
    assert rr, sid_new
    vals = [11, L, M, N, O, P, Q, Rv, Sv, Tv]   # col11=单价, 12..20
    for off, v in enumerate(vals):
        w1.cell(rr, 11 + off, v)
    print(f'{sid_new} (考核第{rr}行): L={L} M={M} N={N} O={O} P={P} Q={Q} R={Rv} S={Sv} T={Tv}')
wbk.save(KAOHE)
print('考核表已补全保存')
