# -*- coding: utf-8 -*-
"""删除汇总里指定车牌 + 时间B 日期的行(用于回溯修正后的清理)"""
import sys, datetime
import pythoncom, win32com.client as wc

EPOCH = datetime.datetime(1899, 12, 30)


def from_excel(v):
    """序列数 -> datetime（自实现，不依赖 openpyxl）"""
    return EPOCH + datetime.timedelta(days=float(v))

SUM = r'C:\Users\Jubin\Desktop\源数据\油耗分析汇总.xlsm'
PLATE = sys.argv[1] if len(sys.argv) > 1 and not sys.argv[1].startswith('--') else '粤BY2J27'
_day = sys.argv[2] if len(sys.argv) > 2 else '2026-09-24'
DAY_B = datetime.date(*[int(x) for x in _day.split('-')])   # 时间B 的日期
PRICE_OLD = float(sys.argv[3]) if len(sys.argv) > 3 else 584.5   # 旧行金额(精确匹配)
CHECK_ONLY = '--check' in sys.argv


def true_last_row(ws, col=1):
    ur = ws.UsedRange
    last_used = ur.Row + ur.Rows.Count - 1
    vals = ws.Range(ws.Cells(1, col), ws.Cells(last_used, col)).Value2
    if vals is None:
        return 1
    if not isinstance(vals, tuple):
        vals = ((vals,),)
    for i in range(len(vals) - 1, -1, -1):
        row = vals[i]
        v = row[0] if isinstance(row, tuple) else row
        if v not in (None, ''):
            return 1 + i
    return 1


pythoncom.CoInitialize()
app = None
try:
    app = wc.DispatchEx('Ket.Application')
    app.Visible = False
    app.DisplayAlerts = False
    wbs = app.Workbooks.Open(SUM, 0, False)
    wss = wbs.Worksheets('Sheet1')
    last = true_last_row(wss, 3)
    print(f'汇总数据末行: {last}')

    targets = []
    for i in range(2, last + 1):
        if str(wss.Cells(i, 3).Value2 or '').strip() != PLATE:
            continue
        tb = wss.Cells(i, 7).Value2
        amt = wss.Cells(i, 8).Value2
        try:
            tb_dt = from_excel(float(tb))
        except (TypeError, ValueError):
            continue
        try:
            amt_f = float(amt)
        except (TypeError, ValueError):
            amt_f = None
        mark = '  <<< 待删' if (tb_dt.date() == DAY_B and amt_f is not None
                                and abs(amt_f - PRICE_OLD) < 1) else ''
        print(f'  行{i}: 时间A={wss.Cells(i,6).Value2} 时间B={tb_dt} 金额={amt_f}{mark}')
        if mark:
            targets.append(i)

    if not targets:
        print('未找到待删行')
    elif CHECK_ONLY:
        print(f'[仅检查] 将删除 {len(targets)} 行: {targets}')
    else:
        for r in sorted(targets, reverse=True):
            wss.Rows(r).Delete()
            print(f'  已删除行{r}')
        last2 = true_last_row(wss, 3)
        for i in range(2, last2 + 1):
            wss.Cells(i, 1).Value2 = i - 1
        wbs.Save()
        print(f'汇总已保存, 现有 {last2 - 1} 行数据')
    wbs.Close(False)
finally:
    if app is not None:
        try:
            app.Quit()
        except Exception:
            pass
    pythoncom.CoUninitialize()
