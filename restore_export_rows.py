# -*- coding: utf-8 -*-
"""把误删/误改的导出记录，从车辆模板的「加油数据」反查并原样写回导出文件。
（车辆模板里的「加油数据」是这份数据的第二副本，可用于反向恢复）

⚠️ 2026-09-28 用户明令禁止 openpyxl 保存文件 —— 本脚本原用 openpyxl 写导出，
   现改为：读取用 python-calamine（只读），写入用 WPS COM（原生保存）。

用法：改下面 TARGETS（sid, 车辆模板文件名, 导出文件中的目标行号）后运行。
"""
import os, shutil, datetime

TPL = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板'
EXP = r'C:\Users\Jubin\Desktop\数据\车辆加油_管理(2026-09-22).xlsx'

# sid -> (模板文件, 导出文件中的空行号)
TARGETS = [
    ('260921000452', 'B13-9.6米-粤BNT993.xlsx', 23),
    ('260921000453', 'D17-4.2米-京LPW138.xlsx', 24),
    ('260921000464', '尿素B16-9.6米-粤BNV030.xlsx', 33),
]

D_FMT, DT_FMT = 'yyyy-mm-dd', 'yyyy-mm-dd hh:mm:ss'
EPOCH = datetime.datetime(1899, 12, 30)


def clean_sid(v):
    s = str(v or '').strip()
    return s[:-2] if s.endswith('.0') else s


def to_serial(v):
    if v is None: return None
    if isinstance(v, (int, float)) and not isinstance(v, bool): return float(v)
    if isinstance(v, datetime.datetime): return (v - EPOCH).total_seconds() / 86400.0
    if isinstance(v, datetime.date):
        return (datetime.datetime(v.year, v.month, v.day) - EPOCH).total_seconds() / 86400.0
    s = str(v).strip()
    for f in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d'):
        try:
            return (datetime.datetime.strptime(s, f) - EPOCH).total_seconds() / 86400.0
        except ValueError:
            pass
    return None


# ---------- 1) 从模板收集原始 14 列（calamine 只读） ----------
from python_calamine import CalamineWorkbook

recs = {}
for sid, fn, exp_row in TARGETS:
    fp = os.path.join(TPL, fn)
    cwb = CalamineWorkbook.from_path(fp)
    rows = cwb.get_sheet_by_name('加油数据').to_python(skip_empty_area=False)
    for row in rows[1:]:
        if row and clean_sid(row[0]) == sid:
            recs[sid] = list(row[:14]) + [None] * max(0, 14 - len(row))
            break
    assert sid in recs, f'{sid} 未在 {fn} 找到'

print('=== 待回填的记录(取自车辆文件) ===')
for sid, fn, exp_row in TARGETS:
    v = recs[sid]
    print(f'  导出行{exp_row} <- {sid} {v[1]} '
          f'{v[12]} 满[{v[8]}] {v[9]} ¥{v[5]} {v[6]}L 备注[{v[13] or "-"}]')

# ---------- 2) 备份导出 ----------
bak = EXP + '.bak_before_restore'
shutil.copy2(EXP, bak)
print(f'\n导出已备份 -> {os.path.basename(bak)}')

# ---------- 3) COM 写回 ----------
import pythoncom
import win32com.client as wc

pythoncom.CoInitialize()
app = None
try:
    for pid in ('Ket.Application', 'et.Application', 'Excel.Application'):
        try:
            app = wc.DispatchEx(pid); break
        except Exception:
            continue
    if app is None:
        raise SystemExit('无法启动 COM')
    app.Visible = False
    app.DisplayAlerts = False
    try: app.ScreenUpdating = False
    except Exception: pass
    wb = app.Workbooks.Open(EXP, 0, False)
    if wb.ReadOnly:
        raise SystemExit('导出文件被占用(只读)')
    ws = wb.Worksheets(1)
    for sid, fn, exp_row in TARGETS:
        v = list(recs[sid])
        existing = [ws.Cells(exp_row, c).Value2 for c in range(1, 15)]
        assert all(x in (None, '') for x in existing), f'行{exp_row} 非空, 中止: {existing}'
        v[2] = to_serial(v[2])      # 登记日期
        v[12] = to_serial(v[12])    # 登记时间
        if v[13] is None: v[13] = ''
        ws.Cells(exp_row, 1).NumberFormat = '@'
        ws.Range(ws.Cells(exp_row, 1), ws.Cells(exp_row, 14)).Value = tuple([tuple(v)])
        ws.Cells(exp_row, 3).NumberFormat = D_FMT
        ws.Cells(exp_row, 13).NumberFormat = DT_FMT
        print(f'  行{exp_row} 已写入: {v[0]} {v[1]}')
    wb.Save()
    wb.Close(False)
    print('\n导出文件已保存（WPS 原生保存）')
finally:
    if app is not None:
        try: app.Quit()
        except Exception: pass
    pythoncom.CoUninitialize()
