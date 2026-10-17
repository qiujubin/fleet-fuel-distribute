# -*- coding: utf-8 -*-
"""粤BT802F 261018000360 单价 2.28 -> 8.28 误分流修正 (2026-10-18)
司机又漏打"8"（647.24/78.17=8.2800），被"单价<5 归尿素"规则误分到尿素文件，
并在尿素侧形成 2026-02-27 -> 10-18 的荒唐周期。
步骤：改导出单价 -> 回滚尿素D14（独立副本） -> 删考核表每日尿素数据误入行 -> 重读复核
"""
import os, shutil, sys
import pythoncom, win32com.client as wc

HERE = os.path.dirname(os.path.abspath(__file__))
CHECK = '--check' in sys.argv
EXPORT = r'C:\Users\Jubin\Desktop\数据\车辆加油_管理(2026-10-18).xlsx'
UREA = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板\尿素D14-4.2米-粤BT802F.xlsx'
UREA_ROLLBACK = os.path.join(HERE, 'backup_fix_bt802f_1018', '尿素D14_回滚源.xlsx')
KAOHE = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
SID = '261018000360'


def disp(v):
    s = str(v or '').strip()
    return s[:-2] if s.endswith('.0') else s


ok = True
pythoncom.CoInitialize()
app = None
for pid in ('Ket.Application', 'et.Application'):
    try:
        app = wc.DispatchEx(pid)
        break
    except Exception:
        continue
app.Visible = False
app.DisplayAlerts = False
try:
    # ---- 1) 导出文件单价 ----
    wb = app.Workbooks.Open(EXPORT, 0, CHECK)
    ws = wb.Worksheets(1)
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    hit1 = None
    for r in range(2, last + 1):
        if disp(ws.Cells(r, 1).Value) == SID:
            hit1 = r
            print('  导出 行%d: 金额=%s 油量=%s 单价=%s' % (r, ws.Cells(r, 6).Value, ws.Cells(r, 7).Value, ws.Cells(r, 8).Value))
            if not CHECK:
                ws.Cells(r, 8).Value = 8.28
            break
    if hit1 is None:
        print('  !! 导出未找到', SID)
        ok = False
    else:
        if CHECK:
            print('  [check] 将把导出 行%d 单价 -> 8.28' % hit1)
        else:
            wb.Save()
            # 只读复核
            wbc = app.Workbooks.Open(EXPORT, 0, True)
            v = wbc.Worksheets(1).Cells(hit1, 8).Value
            wbc.Close(False)
            print('  导出复核: 行%d 单价=%r %s' % (hit1, v, '✓' if abs(float(v) - 8.28) < 1e-9 else '✗ 未落盘!'))
            if abs(float(v) - 8.28) > 1e-9:
                ok = False
    wb.Close(False)

    # ---- 2) 尿素D14 回滚（文件替换，需先关 WPS 句柄）----
    if not CHECK:
        app.Quit()
        app = None
        shutil.copy2(UREA_ROLLBACK, UREA)
        print('  尿素D14 已回滚 <- backup_fix_bt802f_1018/尿素D14_回滚源.xlsx')
    else:
        from python_calamine import CalamineWorkbook
        wb2 = CalamineWorkbook.from_path(UREA_ROLLBACK)
        d2 = wb2.get_sheet_by_name('加油数据').to_python(skip_empty_area=False)
        hit = [i for i, r in enumerate(d2, 1) if r and SID in str(r[0])]
        print('  [check] 回滚源 加油数据=%d 行, 含%s的行=%s（应为无）' % (len(d2), SID, hit or '无'))
        if hit:
            ok = False

    # ---- 3) 考核表 每日尿素数据 删误入行 ----
    if app is None:
        pythoncom.CoInitialize()
        for pid in ('Ket.Application', 'et.Application'):
            try:
                app = wc.DispatchEx(pid)
                break
            except Exception:
                continue
        app.Visible = False
        app.DisplayAlerts = False
    wbk = app.Workbooks.Open(KAOHE, 0, CHECK)
    wsu = wbk.Worksheets('每日尿素数据')
    lastu = wsu.UsedRange.Row + wsu.UsedRange.Rows.Count - 1
    hit3 = None
    for r in range(lastu, 1, -1):
        if disp(wsu.Cells(r, 1).Value) == SID:
            hit3 = r
            print('  每日尿素数据 行%d: 车牌=%s A=%s' % (r, wsu.Cells(r, 4).Value, wsu.Cells(r, 6).Value))
            break
    if hit3 is None:
        print('  每日尿素数据 无 %s 行（已删或未写入）' % SID)
    else:
        if str(wsu.Cells(hit3, 4).Value or '').strip() == '粤BT802F' and not CHECK:
            wsu.Rows(hit3).Delete()
            wbk.Save()
            # 只读复核
            wbc = app.Workbooks.Open(KAOHE, 0, True)
            wsc = wbc.Worksheets('每日尿素数据')
            still = False
            for r in range(2, wsc.UsedRange.Row + wsc.UsedRange.Rows.Count - 1 + 1):
                if disp(wsc.Cells(r, 1).Value) == SID:
                    still = True
                    break
            wbc.Close(False)
            print('  删除行%d 并复核: %s' % (hit3, '✗ 仍存在!' if still else '✓ 已落盘'))
            if still:
                ok = False
        elif CHECK:
            print('  [check] 将删除 每日尿素数据 行%d' % hit3)
        else:
            print('  !! 行%d 车牌校验失败，未删' % hit3)
            ok = False
    wbk.Close(False)
finally:
    try:
        if app is not None:
            app.Quit()
    except Exception:
        pass
    pythoncom.CoUninitialize()

print('RESULT:', 'OK' if ok else 'FAIL')
