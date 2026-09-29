# -*- coding: utf-8 -*-
"""一次性清理: 真实考核表「每日加油记录」中旧版 openpyxl 脚本遗留的残留行
(2..11 = 09-19 的 10 条正确数据; 12..16 = 旧版 cell(None) 无操作 bug 留下的 09-18 残留)"""
import pythoncom, shutil, datetime
import win32com.client as wc

KAOHE = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
BAK = r'E:\WorkSpace\workbuddy\2026-09-18-16-38-15\backup_daily\考核表_清理前备份.xlsm'

try:
    f = open(KAOHE, 'r+b'); f.close(); print('未占用 ✓')
except PermissionError:
    print('考核表被占用, 请关闭 WPS'); raise SystemExit(1)

shutil.copy2(KAOHE, BAK)
print('清理前备份 ->', BAK)

pythoncom.CoInitialize()
app = None
try:
    app = wc.DispatchEx('Ket.Application')
    app.Visible = False; app.DisplayAlerts = False
    try: app.ScreenUpdating = False
    except Exception: pass
    wb = app.Workbooks.Open(KAOHE, 0, False)
    assert not wb.ReadOnly, '只读打开'
    ws = wb.Worksheets('每日加油记录')
    # 确认 2..11 是 09-19 的 10 条, 12..16 是 09-18 残留
    v = ws.Range(ws.Cells(2, 1), ws.Cells(16, 3)).Value2
    d11 = v[9][2]; d12 = v[10][2]   # A11/C11, A12/C12 (0基)
    from openpyxl.utils.datetime import from_excel
    d11s, d12s = str(from_excel(d11))[:10], str(from_excel(d12))[:10]
    print(f'行11日期={d11s} 行12日期={d12s}')
    assert d11s == '2026-09-19' and d12s == '2026-09-18', '状态与预期不符, 中止'
    for r in range(12, 17):
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 14)).ClearContents()
    wb.Save()
    print('已清除 12~16 行残留并保存')
    wb.Close(False)
finally:
    if app is not None:
        try: app.Quit()
        except Exception: pass
    pythoncom.CoUninitialize()
print('完成')
