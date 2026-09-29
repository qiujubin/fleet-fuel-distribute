# -*- coding: utf-8 -*-
"""一次性清理: 油耗分析汇总 Sheet1 的 11 条历史重复行(去重键=车牌+时间A天+时间B天+金额+升数),
删除后重排 A 列序号。COM 操作, 先备份。"""
import pythoncom, shutil, datetime
import win32com.client as wc

S = r'C:\Users\Jubin\Desktop\源数据\油耗分析汇总.xlsm'
BAK = r'E:\WorkSpace\workbuddy\2026-09-18-16-38-15\backup_daily\油耗分析汇总_清理前备份.xlsm'

def day_key(v):
    if v is None or v == '': return ''
    from openpyxl.utils.datetime import from_excel
    try: return str(from_excel(float(v)))[:10]
    except (TypeError, ValueError): return str(v)[:10]

def num_key(v):
    if v is None or v == '': return ''
    try: return str(round(float(v) * 1e4) / 1e4)
    except (TypeError, ValueError): return str(v).strip()

try:
    f = open(S, 'r+b'); f.close(); print('汇总未占用 ✓')
except PermissionError:
    print('汇总被占用, 请关闭 WPS'); raise SystemExit(1)
shutil.copy2(S, BAK)
print('清理前备份 ->', BAK)

pythoncom.CoInitialize()
app = None
try:
    app = wc.DispatchEx('Ket.Application')
    app.Visible = False; app.DisplayAlerts = False
    try: app.ScreenUpdating = False
    except Exception: pass
    wb = app.Workbooks.Open(S, 0, False)
    assert not wb.ReadOnly, '只读打开'
    ws = wb.Worksheets('Sheet1')
    ur = ws.UsedRange
    last = ur.Row + ur.Rows.Count - 1
    print('末行:', last)
    # 读全部数据行键 (C=3车牌 F=6 G=7 H=8 I=9)
    vals = ws.Range(ws.Cells(2, 3), ws.Cells(last, 9)).Value2
    n = last - 1
    seen = set()
    dup_rows = []
    for i in range(n):
        row = vals[i] if isinstance(vals[i], tuple) else (vals[i],)
        # row 对应 C..I 7列: 索引 0=车牌 3=时间A 4=时间B 5=金额 6=升数
        k = '|'.join([str(row[0] or '').strip().upper(), day_key(row[3]), day_key(row[4]), num_key(row[5]), num_key(row[6])])
        if k.replace('|', '') and k in seen:
            dup_rows.append(2 + i)
        else:
            seen.add(k)
    print('重复行:', dup_rows)
    # 从后往前整行删除
    for r in reversed(dup_rows):
        ws.Rows(r).Delete()
    # 重排 A 列序号
    ur2 = ws.UsedRange
    last2 = ur2.Row + ur2.Rows.Count - 1
    n_data = last2 - 1
    seq = list(range(1, n_data + 1))
    ws.Range(ws.Cells(2, 1), ws.Cells(n_data + 1, 1)).Value = tuple((s,) for s in seq)
    wb.Save()
    print(f'已删除 {len(dup_rows)} 行, 重排序号 1~{n_data}, 已保存')
    wb.Close(False)
finally:
    if app is not None:
        try: app.Quit()
        except Exception: pass
    pythoncom.CoUninitialize()
print('完成')
