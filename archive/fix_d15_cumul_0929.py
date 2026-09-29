# -*- coding: utf-8 -*-
"""修复 2026-09-29 D15-粤BY2J27 行155 漏累加

根因: 新版 retro_fix 把「改 F」与「重写 I/J」合并在一次 COM 打开里，
      而 rebuild_ij 当时从**磁盘**读 F/G → 读到旧状态（行154 F 还在、行155 F 为空）
      → 行155 被判为非有效是 → I/J 无累加项。
      已修 com_formula.read_plan_com（从内存态读）。

本脚本修数据（纯 COM）:
  1) D15 计算模板: 重写 I/J
  2) 考核表「每日油耗数据」: 按 sid 重写该行 24 列
  3) 汇总: 删错误行 + 重跑宏重新追加
"""
import os, shutil, atexit, datetime

TPL = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板'
D15 = os.path.join(TPL, 'D15-4.2米-粤BY2J27.xlsx')
KAOHE = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
SUM_F = r'C:\Users\Jubin\Desktop\源数据\油耗分析汇总.xlsm'
BAK = r'E:\WorkSpace\workbuddy\2026-09-18-16-38-15\backup_fix_d15_0929'
SID = '260928000598'
SUM_ROW = 511
DAY_N1 = datetime.date(2026, 9, 28)
DT_FMT = 'yyyy-mm-dd hh:mm:ss'

os.makedirs(BAK, exist_ok=True)
for p in (D15, KAOHE, SUM_F):
    shutil.copy2(p, os.path.join(BAK, os.path.basename(p)))
print('备份 ->', BAK)

import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from com_formula import rebuild_ij, read_plan_com
from region_map import REGION_MAP

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


def to_serial(dt):
    if isinstance(dt, datetime.date) and not isinstance(dt, datetime.datetime):
        dt = datetime.datetime(dt.year, dt.month, dt.day)
    return (dt - datetime.datetime(1899, 12, 30)).total_seconds() / 86400.0


# ===== 1) D15: 重写 I/J =====
w1 = app.Workbooks.Open(D15, 0, False)
if w1.ReadOnly:
    raise SystemExit('D15 被占用(只读)')
opened.append(w1)
plan = read_plan_com(w1)
pv = dict(plan)
print(f'\n[D15] 内存态计划: 行155 累加 {len(pv.get(155, []))} 项 -> {pv.get(155)}')
n = rebuild_ij(app, D15, wb=w1)
print(f'  重写 {n} 行 I/J')
try: app.Calculate()
except Exception: pass
ws = None
for sn in ('计算模板', '油耗计算'):
    try: ws = w1.Worksheets(sn); break
    except Exception: continue
rng = ws.Range(ws.Cells(155, 6), ws.Cells(155, 22)).Value2
v = list(rng[0]) if isinstance(rng[0], tuple) else list(rng)
nm = ['F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V']
print('  行155: ' + '  '.join(f'{nm[i]}={v[i]}' for i in range(len(nm))))
w1.Save(); w1.Close(False)
if w1 in opened: opened.remove(w1)

# ===== 2) 考核表: 按 sid 重写该行 =====
wbk = app.Workbooks.Open(KAOHE, 0, False)
if wbk.ReadOnly:
    raise SystemExit('考核表被占用(只读)')
opened.append(wbk)
wsd = wbk.Worksheets('每日油耗数据')
row_kao = None
for i in range(2, 6000):
    x = wsd.Cells(i, 1).Value2
    if x is not None and str(x).strip().rstrip('.0') == SID:
        row_kao = i; break
print(f'\n[考核表] sid {SID} 在行 {row_kao}')
if row_kao:
    old = [wsd.Cells(row_kao, c).Value2 for c in (3, 4, 5, 11, 2)]
    print('  保留 司机/车牌/车型/单价/序号 =', old)
    def _n(x):
        try: return float(x)
        except (TypeError, ValueError): return None
    F_s, G_s = _n(v[0]), _n(v[1])
    F_d = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=F_s)
    G_d = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=G_s)
    rowvals = [SID, old[4], old[0], old[1], old[2],
               F_s, G_s, '是', _n(v[3]), _n(v[4]), old[3],
               _n(v[6]), _n(v[7]), _n(v[8]), _n(v[9]), _n(v[10]), _n(v[11]),
               _n(v[12]), _n(v[13]), _n(v[14]), _n(v[15]), _n(v[16]),
               to_serial(G_d.date()), 0]
    wsd.Cells(row_kao, 1).NumberFormat = '@'
    wsd.Range(wsd.Cells(row_kao, 1), wsd.Cells(row_kao, 24)).Value = tuple([tuple(rowvals)])
    wsd.Cells(row_kao, 6).NumberFormat = DT_FMT
    wsd.Cells(row_kao, 7).NumberFormat = DT_FMT
    wsd.Cells(row_kao, 23).NumberFormat = 'yyyy-mm-dd'
    print(f'  已重写: I={_n(v[3])} J={_n(v[4])} P={_n(v[10])} R={_n(v[12])}')
try: app.Calculate()
except Exception: pass
# 刷新「每日油耗」透视表 + 选 09-28
try:
    wsp = wbk.Worksheets('每日油耗')
    sd = to_serial(datetime.datetime(DAY_N1.year, DAY_N1.month, DAY_N1.day))
    tg = {f'{DAY_N1.year}/{DAY_N1.month}/{DAY_N1.day}', DAY_N1.strftime('%Y-%m-%d'), str(int(sd)), str(sd)}
    for i in range(1, wsp.PivotTables().Count + 1):
        pt = wsp.PivotTables(i); pt.RefreshTable()
        pf = pt.PivotFields('日期')
        for j in range(1, pf.PivotItems().Count + 1):
            it = pf.PivotItems(j)
            vis = str(it.Name).strip() in tg
            if it.Visible != vis: it.Visible = vis
    print('  「每日油耗」透视表已刷新, 日期 2026-09-28')
except Exception as e:
    print('  透视表告警:', e)
try: app.Calculate()
except Exception: pass
wbk.Save(); wbk.Close(False)
if wbk in opened: opened.remove(wbk)

# ===== 3) 汇总: 删错误行 =====
wsum = app.Workbooks.Open(SUM_F, 0, False)
if not wsum.ReadOnly:
    opened.append(wsum)
    wss = wsum.Worksheets('Sheet1')
    plate = str(wss.Cells(SUM_ROW, 3).Value2 or '')
    print(f'\n[汇总] 行{SUM_ROW} 车牌={plate}')
    if plate.strip().upper() == '粤BY2J27':
        wss.Rows(SUM_ROW).Delete()
        wsum.Save()
        print('  已删除并保存')
    else:
        print('  !! 车牌不符, 未删除')
    wsum.Close(False)
    if wsum in opened: opened.remove(wsum)

# ===== 4) 重跑宏 =====
print('\n[宏] 重跑 同步油耗到汇总')
try:
    before = set(w.Name for w in app.Workbooks)
    ret = app.Run('同步油耗到汇总')
    print('  ', ret)
    for w2 in [w for w in app.Workbooks if w.Name not in before]:
        try:
            nm2 = w2.Name
            if '汇总' in nm2:
                p2r = {}
                for reg, plates in REGION_MAP.items():
                    for p in plates: p2r[p.strip().upper()] = reg
                ws3 = w2.Worksheets('Sheet1')
                ur = ws3.UsedRange
                last2 = ur.Row + ur.Rows.Count - 1
                vals = ws3.Range(ws3.Cells(2, 3), ws3.Cells(last2, 3)).Value2
                if vals is None: vals = ()
                if not isinstance(vals, tuple): vals = ((vals,),)
                filled = 0
                for i2, cv in enumerate(vals):
                    r2 = 2 + i2
                    pv2 = cv[0] if isinstance(cv, tuple) else cv
                    pl = str(pv2 or '').strip().upper()
                    if not pl: continue
                    if ws3.Cells(r2, 2).Value2 not in (None, ''): continue
                    reg = p2r.get(pl)
                    if reg:
                        ws3.Cells(r2, 2).Value = reg; filled += 1
                if filled: w2.Save()
                print(f'  区域填充补 {filled} 行 ({nm2})')
            w2.Close(True)
        except Exception as e2:
            print('  收尾告警:', e2)
except Exception as e:
    print('  宏调用失败:', e)

print('\n完成')
