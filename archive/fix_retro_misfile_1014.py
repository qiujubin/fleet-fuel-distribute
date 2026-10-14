# -*- coding: utf-8 -*-
"""10-14 数据修正（不含待确认项）

上游 bug：retro_fix.py 第 190 行写死 Worksheets('每日油耗数据')，处理**尿素文件**时
把修正行也写进了「每日油耗数据」。今天 261014000272（粤BNV030 尿素）被错写进去，
且"追加"分支还漏填 司机/车牌/类型/单价。retro_fix.py 已修（按文件名选表 + 补全表头字段）。

本脚本只处理**已经确定的两件事**：
  A. 删掉「每日油耗数据」里的错位行 261014000272
  B. 修正 粤BPV550 261013000267 的单价 0.28 -> 8.28
     （金额 1417.87 ÷ 油量 171.24 = 8.2800；同车其它记录 8.29/8.31）
     影响面：**考核表与汇总不受影响**（它们的单价/油价都是 金额÷升数 推出来的，已核过
     考核行3542 单价=8.29、汇总行640 油价=8.29），所以只需把这一个字段改对。

**不动**：每日尿素数据里 261014000271 那行 —— 271/272 两条数值完全相同（36.29/15.25，
相隔 13 分钟，同一登记人），疑似重复录入，等用户确认后再处理。

用法：python fix_retro_misfile_1014.py [--check]
"""
import os, sys, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
KAOHE = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
TPL_DIR = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板'
EXPORT = r'C:\Users\Jubin\Desktop\数据\车辆加油_管理(2026-10-14).xlsx'
C3 = os.path.join(TPL_DIR, 'C3-7.6米-粤BPV550.xlsx')
BAK = os.path.join(HERE, 'backup_fix_retro_misfile_1014')

STRAY_SID = '261014000272'      # 错位行进「每日油耗数据」
PRICE_SID = '261013000267'      # 单价要修的记录
NEW_PRICE = 8.28

CHECK = '--check' in sys.argv


def disp(v):
    s = str(v or '').strip()
    return s[:-2] if s.endswith('.0') else s


def main():
    os.makedirs(BAK, exist_ok=True)
    if not CHECK:
        for p in (KAOHE, C3, EXPORT):
            shutil.copy2(p, BAK)
        print('已备份 考核表 + C3模板 + 导出 ->', os.path.basename(BAK))

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
    ok = True
    try:
        # ---- A. 删错位行 ----
        wb = app.Workbooks.Open(KAOHE, 0, CHECK)
        print('考核表 ReadOnly =', wb.ReadOnly)
        ws = wb.Worksheets('每日油耗数据')
        ur = ws.UsedRange
        last = ur.Row + ur.Rows.Count - 1
        hit = None
        for r in range(last, 1, -1):
            if disp(ws.Cells(r, 1).Value) == STRAY_SID:
                plate = str(ws.Cells(r, 4).Value or '').strip()
                print('  命中 行%d 车牌=[%s] 类型=[%s] 金额=%s 升=%s'
                      % (r, plate, ws.Cells(r, 5).Value, ws.Cells(r, 9).Value, ws.Cells(r, 10).Value))
                if not plate:          # 只删"缺车牌"的错位行，稳妥
                    hit = r
                break
        if hit is None:
            print('  未找到可删的错位行（或车牌非空，跳过）')
        else:
            if not CHECK:
                ws.Rows(hit).Delete()
                wb.Save()
            print('  -> 已删除行 %d' % hit)
        wb.Close(False)

        # ---- B. 改 C3 模板 单价 ----
        wb2 = app.Workbooks.Open(C3, 0, CHECK)
        print('C3模板 ReadOnly =', wb2.ReadOnly)
        wsd = wb2.Worksheets('加油数据')
        ur2 = wsd.UsedRange
        last2 = ur2.Row + ur2.Rows.Count - 1
        h2 = None
        for r in range(2, last2 + 1):
            if disp(wsd.Cells(r, 1).Value) == PRICE_SID:
                h2 = r
                print('  行%d %s 单价 %s -> %s' % (r, PRICE_SID, wsd.Cells(r, 8).Value, NEW_PRICE))
                if not CHECK:
                    wsd.Cells(r, 8).Value = NEW_PRICE
                break
        if h2 is None:
            print('  !! 模板未找到该编号'); ok = False
        if not CHECK and h2 is not None:
            wb2.Save(); app.Calculate()
        wb2.Close(False)

        # ---- B2. 改导出文件 单价 ----
        wb3 = app.Workbooks.Open(EXPORT, 0, CHECK)
        print('导出 ReadOnly =', wb3.ReadOnly)
        ws3 = wb3.Worksheets(1)
        ur3 = ws3.UsedRange
        last3 = ur3.Row + ur3.Rows.Count - 1
        h3 = None
        for r in range(2, last3 + 1):
            if disp(ws3.Cells(r, 1).Value) == PRICE_SID:
                h3 = r
                print('  行%d %s 单价 %s -> %s' % (r, PRICE_SID, ws3.Cells(r, 8).Value, NEW_PRICE))
                if not CHECK:
                    ws3.Cells(r, 8).Value = NEW_PRICE
                break
        if h3 is None:
            print('  !! 导出未找到该编号'); ok = False
        if not CHECK and h3 is not None:
            wb3.Save()
        wb3.Close(False)
    finally:
        try: app.Quit()
        except Exception: pass
        pythoncom.CoUninitialize()

    # ---- 只读复核 ----
    from python_calamine import CalamineWorkbook
    wb = CalamineWorkbook.from_path(KAOHE)
    rows = wb.get_sheet_by_name('每日油耗数据').to_python(skip_empty_area=False)
    left = [i for i, r in enumerate(rows, 1) if i > 1 and r and disp(r[0]) == STRAY_SID]
    print('复核 每日油耗数据: 错位行残留=%s' % (left or '无（正常）'))
    if not CHECK and left:
        ok = False; print('  !! 未落盘')
    wb2 = CalamineWorkbook.from_path(C3)
    d = wb2.get_sheet_by_name('加油数据').to_python(skip_empty_area=False)
    for r in d:
        if r and disp(r[0]) == PRICE_SID:
            print('复核 C3模板 单价 = %s (期望 %s)' % (r[7], NEW_PRICE))
            if not CHECK and abs(float(r[7]) - NEW_PRICE) > 1e-6:
                ok = False; print('  !! 未落盘')
    wb3 = CalamineWorkbook.from_path(EXPORT)
    e = wb3.get_sheet_by_name(wb3.sheet_names[0]).to_python(skip_empty_area=False)
    for r in e:
        if r and disp(r[0]) == PRICE_SID:
            print('复核 导出 单价 = %s (期望 %s)' % (r[7], NEW_PRICE))

    if CHECK:
        print('\n[--check] 未做任何写操作')
    return 0 if ok else 2


if __name__ == '__main__':
    sys.exit(main())
