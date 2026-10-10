# -*- coding: utf-8 -*-
"""10-10 补删：考核表「每日尿素数据」里的误入行 261009000184（粤BT802F）

前次删除因 Open(..., not CHECK) 实际为只读而静默失败，这次：
  - 正确以可写方式打开
  - 删行后 Save，再【重新打开只读复核】，确认行确已消失
"""
import os, sys

KAOHE = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
UID = '261009000184'
SHEET = '每日尿素数据'


def disp(v):
    s = str(v or '').strip()
    return s[:-2] if s.endswith('.0') else s


def main():
    check = '--check' in sys.argv
    import pythoncom, win32com.client as wc
    pythoncom.CoInitialize()
    app = None
    for pid in ('Ket.Application', 'et.Application'):
        try:
            app = wc.DispatchEx(pid); break
        except Exception:
            continue
    app.Visible = False
    app.DisplayAlerts = False
    try:
        wb = app.Workbooks.Open(KAOHE, 0, check)
        print('ReadOnly =', wb.ReadOnly)
        ws = wb.Worksheets(SHEET)
        ur = ws.UsedRange
        last = ur.Row + ur.Rows.Count - 1
        target = None
        for r in range(last, 1, -1):
            if disp(ws.Cells(r, 1).Value) == UID:
                target = r
                print('  命中 行%d 车牌=%s 单价=%s' % (r, ws.Cells(r, 4).Value, ws.Cells(r, 11).Value))
                break
        if target is None:
            print('  未找到，无需处理')
        elif check:
            print('  [--check] 未删除')
        else:
            ws.Rows(target).Delete()
            wb.Save()
            print('  已删除行 %d 并保存' % target)
        wb.Close(False)
    finally:
        try: app.Quit()
        except Exception: pass
        pythoncom.CoUninitialize()

    # ---- 复核：重新只读打开 ----
    from python_calamine import CalamineWorkbook
    wb2 = CalamineWorkbook.from_path(KAOHE)
    rows = wb2.get_sheet_by_name(SHEET).to_python(skip_empty_area=False)
    n = sum(1 for r in rows[1:] if r and r[0] not in (None, ''))
    hit = [i for i, r in enumerate(rows, 1) if i > 1 and r and UID in str(r[0])]
    print('复核: %s 数据行=%d 含%s行=%s' % (SHEET, n, UID, hit or '无'))
    return 0 if not hit else 3


if __name__ == '__main__':
    sys.exit(main())
