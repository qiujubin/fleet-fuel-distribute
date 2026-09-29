# -*- coding: utf-8 -*-
"""粤BPV550 单据 260927000582 全链路修复（I/J 全量累加 bug 的下游）

上游已修: C3-7.6米-粤BPV550.xlsx 计算模板行271
  I 586719.09 -> 3465.65   J 82760.16 -> 418.3   (1767.01+1698.64 / 213.15+205.15)

本脚本修下游:
  1) 考核表「每日油耗数据」行 3393 的 24 列 (从模板行271 实时读回重组)
  2) 刷新「每日油耗」透视表 + 日期选 2026-09-27
  3) 4 张油耗 sheet 公式重算(其中 干线油耗!行9 = 粤BPV550)
  4) 汇总 Sheet1 删掉错误行 501, 重跑宏「同步油耗到汇总」重新追加正确行
  5) 区域填充兜底(后台宏跨簿守卫)

用法: python fix_bpv550_chain_0928.py [--check]
"""
import os, sys, shutil, atexit, datetime

KAOHE = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
SUM_F = r'C:\Users\Jubin\Desktop\源数据\油耗分析汇总.xlsm'
TPL_C3 = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板\C3-7.6米-粤BPV550.xlsx'
BAK = r'E:\WorkSpace\workbuddy\2026-09-18-16-38-15\backup_fix_bpv550_0928'
MACRO_SYNC = '同步油耗到汇总'

TPL_ROW = 271        # C3 模板里的有效是行
KAOHE_ROW = 3393     # 每日油耗数据 里的错误行
SUM_ROW = 501        # 汇总 Sheet1 里的错误行
DAY_N1 = datetime.date(2026, 9, 27)

CHECK = '--check' in sys.argv

os.makedirs(BAK, exist_ok=True)
for p in (KAOHE, SUM_F, TPL_C3):
    shutil.copy2(p, os.path.join(BAK, os.path.basename(p)))
print('备份 ->', BAK)
print('  ' + ', '.join(os.listdir(BAK)))

if CHECK:
    print('体检模式: 仅备份, 不写入')
    sys.exit(0)

import pythoncom
import win32com.client as wc

pythoncom.CoInitialize()
app = None
_opened = []


def cleanup():
    global app
    try:
        if app is None: return
        for wb in list(_opened):
            try: wb.Close(False)
            except Exception: pass
        _opened.clear()
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

DT_FMT = 'yyyy-mm-dd hh:mm:ss'
D_FMT = 'yyyy-mm-dd'


def to_serial(v):
    if v is None: return None
    if isinstance(v, datetime.datetime):
        return (v - datetime.datetime(1899, 12, 30)).total_seconds() / 86400.0
    if isinstance(v, datetime.date):
        return (datetime.datetime(v.year, v.month, v.day) - datetime.datetime(1899, 12, 30)).total_seconds() / 86400.0
    return v


# ============ 1. 从模板行 271 读实时值 ============
wt = app.Workbooks.Open(TPL_C3, 0, False)
if wt.ReadOnly:
    raise SystemExit('模板被占用(只读): ' + TPL_C3)
_opened.append(wt)
wst = wt.Worksheets('计算模板')
app.Calculate()
rng = wst.Range(wst.Cells(TPL_ROW, 6), wst.Cells(TPL_ROW, 22)).Value2
v = list(rng[0]) if isinstance(rng[0], tuple) else list(rng)
name = ['F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V']
tv = {}
for i, nm in enumerate(name):
    tv[nm] = v[i]
print('\n[模板行271] ' + '  '.join(f'{k}={tv[k]}' for k in name))
# 兜底: Q=0 时 WPS 返回 #DIV/0!(-2146826281); 同 sheet 其他粤BPV550 行均用 GPS 值替代
if tv['Q'] in (0, None) or (isinstance(tv['Q'], (int, float)) and float(tv['Q']) <= 0):
    print('  Q=0 -> T/U/V 用 R/S/R 兜底(与同车历史行一致)')
    tv['T'], tv['U'], tv['V'] = tv['R'], tv['S'], tv['R']
for k in ('T', 'U', 'V'):
    if isinstance(tv[k], (int, float)) and float(tv[k]) < -1e6:
        tv[k] = tv['R'] if k != 'U' else tv['S']
sid = str(wst.Cells(TPL_ROW, 1).Value or '')
person = str(wst.Cells(TPL_ROW, 3).Value or '')
plate = str(wst.Cells(TPL_ROW, 4).Value or '')
vtype = str(wst.Cells(TPL_ROW, 5).Value or '')
print(f'  单据={sid} 司机={person} 车牌={plate} 车型={vtype}')
wt.Close(False)

# ============ 2. 考核表: 重写每日油耗数据行 + 刷新透视表 ============
wbk = app.Workbooks.Open(KAOHE, 0, False)
if wbk.ReadOnly:
    raise SystemExit('考核表被占用(只读)')
_opened.append(wbk)
ws = wbk.Worksheets('每日油耗数据')
cur = str(ws.Cells(KAOHE_ROW, 1).Value or '')
print(f'\n[考核表] 行{KAOHE_ROW} 当前系统编号={cur}')
if str(int(float(cur))) != sid:
    print('  !! 行号与预期不符, 按系统编号重新定位')
    found = None
    for r in range(2, 3406):
        c = ws.Cells(r, 1).Value
        if c is not None and str(c).strip() in (sid, sid + '.0'):
            found = r; break
    if found:
        KAOHE_ROW = found
        print(f'  定位到行 {KAOHE_ROW}')
    else:
        raise SystemExit('找不到该考核行')

F_s, G_s = tv['F'], tv['G']
F_d = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(F_s))
G_d = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(G_s))
rowvals = [sid, TPL_ROW - 1, person, plate, vtype,
           to_serial(F_d), to_serial(G_d), '是', tv['I'], tv['J'], tv['K'],
           tv['L'], tv['M'], tv['N'], tv['O'], tv['P'], tv['Q'],
           tv['R'], tv['S'], tv['T'], tv['U'], tv['V'],
           to_serial(G_d.date()), 0]
ws.Cells(KAOHE_ROW, 1).NumberFormat = '@'
for c, val in enumerate(rowvals, start=1):
    ws.Cells(KAOHE_ROW, c).Value = val
ws.Cells(KAOHE_ROW, 6).NumberFormat = DT_FMT
ws.Cells(KAOHE_ROW, 7).NumberFormat = DT_FMT
ws.Cells(KAOHE_ROW, 23).NumberFormat = D_FMT
print('  已重写 24 列: 金额%s 升数%s 里程%s 油耗%s' % (tv['I'], tv['J'], tv['P'], tv['R']))

app.Calculate()

# 刷新「每日油耗」透视表 + 选 n-1
try:
    wsp = wbk.Worksheets('每日油耗')
    serial_d = to_serial(datetime.datetime(DAY_N1.year, DAY_N1.month, DAY_N1.day))
    targets = {f'{DAY_N1.year}/{DAY_N1.month}/{DAY_N1.day}', DAY_N1.strftime('%Y-%m-%d'),
               str(int(serial_d)), str(serial_d)}
    for i in range(1, wsp.PivotTables().Count + 1):
        pt = wsp.PivotTables(i)
        pt.RefreshTable()
        pf = pt.PivotFields('日期')
        for j in range(1, pf.PivotItems().Count + 1):
            it = pf.PivotItems(j)
            vis = str(it.Name).strip() in targets
            if it.Visible != vis:
                it.Visible = vis
    print(f'  「每日油耗」透视表已刷新, 日期选 {DAY_N1}')
except Exception as e:
    print('  透视表刷新告警:', e)

app.Calculate()
# 验证 4 张油耗 sheet
for sn in ('干线油耗', '华东油耗', '华南油耗', '华北油耗'):
    try:
        w = wbk.Worksheets(sn)
        for r in range(1, 40):
            if plate in str(w.Cells(r, 2).Value or ''):
                vals = [w.Cells(r, c).Value for c in range(2, 12)]
                print(f'  [{sn}] 行{r}: ' + '|'.join(str(x)[:14] for x in vals))
    except Exception:
        pass

wbk.Save()
print('  考核表已保存')

# ============ 3. 汇总: 删错误行 ============
print(f'\n[汇总] 删除行 {SUM_ROW}')
wsum = app.Workbooks.Open(SUM_F, 0, False)
if wsum.ReadOnly:
    print('  !! 汇总被占用(只读), 跳过删行')
    wsum = None
else:
    _opened.append(wsum)
    ws2 = wsum.Worksheets('Sheet1')
    plate_c = str(ws2.Cells(SUM_ROW, 3).Value or '')
    print(f'  行{SUM_ROW} 车牌={plate_c}')
    if plate_c.strip().upper() == plate.upper():
        ws2.Rows(SUM_ROW).Delete()
        wsum.Save()
        print('  已删除并保存')
    else:
        print('  !! 车牌不符, 未删除')
    wsum.Close(False)
    if wsum in _opened: _opened.remove(wsum)

# ============ 4. 重跑宏 ============
print('\n[宏] 重跑', MACRO_SYNC)
try:
    before = set(w.Name for w in app.Workbooks)
    ret = app.Run(MACRO_SYNC)
    print('  ', ret)
    for w2 in [w for w in app.Workbooks if w.Name not in before]:
        try:
            nm = w2.Name
            if '汇总' in nm:
                from region_map import REGION_MAP
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
                    pv = cv[0] if isinstance(cv, tuple) else cv
                    pl = str(pv or '').strip().upper()
                    if not pl: continue
                    bv = ws3.Cells(r2, 2).Value2
                    if bv not in (None, ''): continue
                    reg = p2r.get(pl)
                    if reg:
                        ws3.Cells(r2, 2).Value = reg; filled += 1
                if filled: w2.Save()
                print(f'  区域填充补 {filled} 行 ({nm})')
            w2.Close(True)
        except Exception as e2:
            print('  汇总收尾告警:', e2)
except Exception as e:
    print('  宏调用失败:', e)

wbk.Close(False)
if wbk in _opened: _opened.remove(wbk)

print('\n修复完成')
