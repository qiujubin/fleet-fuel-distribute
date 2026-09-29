# -*- coding: utf-8 -*-
"""一次性修正: 260928000587 (粤BT802F 09-28 06:56) —— 单价 2.28 -> 8.28 (司机笔误)

判断依据(用户确认 + 里程证据):
  - 油车文件 D14-4.2米-粤BT802F 上一笔有效"是" = 09-26 16:10:09 -> 里程 308.6 km, 油耗 ~15.9L/100km (正常)
  - 尿素文件 上一笔有效"是" = 2026-02-27 21:04 -> 里程 60096.7 km (荒唐)
  -> 该笔实为加柴油, 被单价 2.28 误判进尿素链路

动作:
  1) 导出文件 加注单价(H) 2.28 -> 8.28
  2) 尿素D14-4.2米-粤BT802F.xlsx: 「加油数据」删该行 + 「计算模板」删该行
  3) 考核表「每日尿素数据」: 删该行 + 刷新"每日尿素"透视表
  4) 随后重跑 run_daily.py -> 自动写入油车文件 + 每日油耗数据 + 透视表 + 宏

用法: python fix_bt802f_0928.py --check   # 只体检, 不写入
      python fix_bt802f_0928.py           # 执行修正
"""
import os, sys, shutil, atexit, datetime

SID       = '260928000587'
PLATE     = '粤BT802F'
NEW_PRICE = 8.28

EXPORT = r'C:\Users\Jubin\Desktop\数据\车辆加油_管理(2026-09-28).xlsx'
UREA_F = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板\尿素D14-4.2米-粤BT802F.xlsx'
FUEL_F = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板\D14-4.2米-粤BT802F.xlsx'
KAOHE  = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
SUM_F  = r'C:\Users\Jubin\Desktop\源数据\油耗分析汇总.xlsm'
BAK    = r'E:\WorkSpace\workbuddy\2026-09-18-16-38-15\backup_fix_bt802f_0928'

CHECK_ONLY = '--check' in sys.argv


# ---------------- 工具 ----------------
def sid_key(v):
    s = str(v).strip()
    try:
        f = float(s)
        if f == int(f): return str(int(f))
    except ValueError:
        pass
    return s


def true_last_row(ws, col=1):
    ur = ws.UsedRange
    end = min(ur.Row + ur.Rows.Count - 1, ws.Rows.Count)
    while end >= 1:
        start = max(1, end - 1999)
        vals = ws.Range(ws.Cells(start, col), ws.Cells(end, col)).Value2
        if vals is None:
            end = start - 1; continue
        if not isinstance(vals, tuple): vals = ((vals,),)
        for i in range(len(vals) - 1, -1, -1):
            v = vals[i][0] if isinstance(vals[i], tuple) else vals[i]
            if v is not None and str(v).strip() != '':
                return start + i
        end = start - 1
    return 0


def find_sid_row(ws, sid, col=1):
    last = true_last_row(ws, col)
    if last < 2: return None
    vals = ws.Range(ws.Cells(2, col), ws.Cells(last, col)).Value2
    if not isinstance(vals, tuple): vals = ((vals,),)
    for i, row in enumerate(vals, start=2):
        v = row[0] if isinstance(row, tuple) else row
        if sid_key(v) == sid:
            return i
    return None


# ---------------- 备份 ----------------
os.makedirs(BAK, exist_ok=True)
for src in (EXPORT, UREA_F, FUEL_F, KAOHE):
    if os.path.exists(src):
        shutil.copy2(src, os.path.join(BAK, os.path.basename(src)))
print(f'备份 -> {BAK}')
print('  ' + ', '.join(os.listdir(BAK)))

import pythoncom
import win32com.client as wc
from python_calamine import CalamineWorkbook

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
    raise SystemExit('无法启动 WPS/Excel COM')
app.Visible = False
app.DisplayAlerts = False
try: app.ScreenUpdating = False
except Exception: pass

report = []


def open_wb(path):
    wb = app.Workbooks.Open(path, 0, False)
    if wb.ReadOnly:
        raise RuntimeError(f'文件被占用(只读打开): {path}')
    _opened.append(wb)
    return wb


# ========== 1. 导出文件: 单价 -> 8.28 ==========
wb = open_wb(EXPORT)
ws = wb.Worksheets(1)
r = find_sid_row(ws, SID)
if r:
    old = ws.Cells(r, 8).Value
    print(f'[导出] 行{r} 加注单价 {old} -> {NEW_PRICE}')
    if not CHECK_ONLY:
        ws.Cells(r, 8).Value = NEW_PRICE
        wb.Save()
        report.append(f'导出 行{r} 单价 {old} -> {NEW_PRICE}')
else:
    print('[导出] 未找到该记录')
try: wb.Close(False)
except Exception: pass

# ========== 2. 尿素文件: 删行 ==========
wb = open_wb(UREA_F)
wsd = wb.Worksheets('加油数据')
rd = find_sid_row(wsd, SID)
print(f'[尿素] 加油数据 该行 = {rd}')
if rd and not CHECK_ONLY:
    wsd.Rows(rd).Delete()
wst = None
for sn in ('计算模板', '油耗计算'):
    try: wst = wb.Worksheets(sn); break
    except Exception: continue
rt = find_sid_row(wst, SID)
print(f'[尿素] 计算模板 该行 = {rt}')
if rt and not CHECK_ONLY:
    wst.Rows(rt).Delete()
if (rd or rt) and not CHECK_ONLY:
    wb.Save()
    report.append(f'尿素文件 删 加油数据行{rd} + 计算模板行{rt}')
try: wb.Close(False)
except Exception: pass

# ========== 3. 考核表: 删每日尿素数据行 ==========
wb = open_wb(KAOHE)
wsu = wb.Worksheets('每日尿素数据')
ru = find_sid_row(wsu, SID)
print(f'[考核] 每日尿素数据 该行 = {ru}')
if ru and not CHECK_ONLY:
    wsu.Rows(ru).Delete()
    try: app.Calculate()
    except Exception: pass
    # 刷新每日尿素透视表
    try:
        for ws2 in wb.Worksheets:
            try:
                pcs = ws2.PivotTables()
                for i in range(1, pcs.Count + 1):
                    pt = pcs.Item(i)
                    pt.RefreshTable()
                    print(f'  透视表已刷新: {ws2.Name} / {pt.Name}')
            except Exception:
                pass
    except Exception as e:
        print('  透视表刷新告警:', e)
    wb.Save()
    report.append(f'考核表 每日尿素数据 删行{ru}')
try: wb.Close(False)
except Exception: pass

# ========== 4. 体检: 汇总里是否残留 ==========
print('\n--- 汇总残留检查 ---')
if os.path.exists(SUM_F):
    wbs = CalamineWorkbook.from_path(SUM_F)
    hit = 0
    for sn in wbs.sheet_names:
        try: rows = wbs.get_sheet_by_name(sn).to_python()
        except Exception: continue
        for i, row in enumerate(rows, 1):
            txt = '|'.join(str(x) for x in row)
            if '260928000587' in txt:
                print(f'  [残留] {sn} 行{i}: {txt[:120]}'); hit += 1
    print('  汇总无该单据号残留' if not hit else f'  共 {hit} 处')
else:
    print('  汇总文件不存在:', SUM_F)

print('\n===== 修正' + ('体检' if CHECK_ONLY else '完成') + ' =====')
for x in report: print('  ✓', x)
if not CHECK_ONLY:
    print('\n下一步: 重跑 run_daily.py 把该笔写入油车文件并同步考核/透视表/汇总')
