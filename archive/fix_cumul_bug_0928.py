# -*- coding: utf-8 -*-
"""修复 rebuild_formulas 锚点匹配 bug 造成的 I/J 全量累加（2026-09-28）

扫描判据: 「计算模板」有效是行的 I 列公式累加项 > 80 项
受影响:
  C3-7.6米-粤BPV550.xlsx   行271  (累加 269 项, 行2~270)
  D15-4.2米-粤BY2J27.xlsx  行112  (累加 110 项)
  D17-4.2米-京LPW138.xlsx  行199  (累加 197 项)

根因: retro_fix 把 F 锚点写成**序列数**(46291.42), 而 rebuild_formulas 用
      str(F)[:19] 去比对 g_rows 里的 datetime 字符串 -> 锚行找不到 ->
      退化成"首周期"从第 2 行全量累加。已修 tkey() 统一成整数秒键。

动作: 备份 -> rebuild(修好匹配后) -> COM 重算 -> 读回校验
用法: python fix_cumul_bug_0928.py [--check]
"""
import os, sys, shutil, atexit, datetime

sys.path.insert(0, r'E:\WorkSpace\workbuddy\2026-09-18-16-38-15')
from rebuild_formulas import rebuild

TPL = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板'
BAK = r'E:\WorkSpace\workbuddy\2026-09-18-16-38-15\backup_fix_cumul_0928'
TARGETS = [
    ('C3-7.6米-粤BPV550.xlsx', 271),
    ('D15-4.2米-粤BY2J27.xlsx', 112),
    ('D17-4.2米-京LPW138.xlsx', 199),
]
CHECK = '--check' in sys.argv

if CHECK:
    print('体检模式: 只 rebuild 到临时副本? -> 直接跳过, 只打印计划')
    for fn, r in TARGETS:
        print(f'  将修复 {fn} 行{r}')
    sys.exit(0)

os.makedirs(BAK, exist_ok=True)
for fn, _ in TARGETS:
    shutil.copy2(os.path.join(TPL, fn), os.path.join(BAK, fn))
print('备份 ->', BAK)
print('  ' + ', '.join(os.listdir(BAK)))

print('\n--- rebuild ---')
for fn, r in TARGETS:
    fp = os.path.join(TPL, fn)
    n, filled = rebuild(fp)
    print(f'  {fn}: rebuild {n} 公式, 已填G行 {filled}')

# ---------- COM 重算 + 校验 ----------
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
        app = wc.DispatchEx(pid); print('\nCOM 应用:', pid); break
    except Exception:
        continue
if app is None:
    raise SystemExit('无法启动 COM')
app.Visible = False
app.DisplayAlerts = False
try: app.ScreenUpdating = False
except Exception: pass

print('\n--- 重算校验 ---')
for fn, r in TARGETS:
    fp = os.path.join(TPL, fn)
    wb = app.Workbooks.Open(fp, 0, False)
    if wb.ReadOnly:
        print(f'!! {fn} 被占用(只读)'); continue
    _opened.append(wb)
    wst = None
    for sn in ('计算模板', '油耗计算'):
        try: wst = wb.Worksheets(sn); break
        except Exception: continue
    app.Calculate()
    rng = wst.Range(wst.Cells(r, 6), wst.Cells(r, 22)).Value2
    v = list(rng[0]) if isinstance(rng[0], tuple) else list(rng)
    name = ['F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V']
    print(f'  {fn} 行{r}:')
    for i, nm in enumerate(name):
        val = v[i]
        if nm in ('F', 'G') and isinstance(val, float):
            try: val = f'{val:.6f} ({datetime.datetime(1899,12,30)+datetime.timedelta(days=val):%Y-%m-%d %H:%M:%S})'
            except Exception: pass
        print(f'      {nm} = {val}')
    wb.Save(); wb.Close(False)

print('\n完成。若 I/J 已回到单笔/周期累计量级, 再跑修考核表与汇总的脚本。')
