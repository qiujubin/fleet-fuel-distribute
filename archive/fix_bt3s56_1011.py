# -*- coding: utf-8 -*-
"""10-11 修正：粤BT3S56 油量小数点右移两位（6433 -> 64.33）导致油耗 1873 L/100km

依据：
  261011000213 粤BT3S56 10-11 03:55  ¥532.65 / 油量录成 6433 / @8.28
  金额 ÷ 单价 = 532.65 / 8.28 = 64.328  -> 真实油量 64.33 L（小数点右移两位）
  同车 10-10 那条 650 / 78.41 @8.29 自洽；6433 L 对 4.2 米车也不可能

污染：考核表「每日油耗数据」行3517 升数 6433 / 油耗 1873.87（尚未进汇总，T+1）

处置（照既有「误分流/坏值修正」流程）：
  1. 备份导出 / 模板 / 考核表
  2. 把 backup_daily 的模板快照拷到独立目录（防重跑覆盖）
  3. 改导出文件 油量 6433 -> 64.33
  4. 用快照回滚 D13-4.2米-粤BT3S56.xlsx（去掉今日 2 行，交重跑重写）
  5. 删考核表「每日油耗数据」今日 2 行（261010000203 / 261011000213）
  6. 只读复核落盘（铁律 8）

用法：python fix_bt3s56_1011.py [--check]
"""
import os, sys, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
TPL_DIR = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板'
KAOHE = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
EXPORT = r'C:\Users\Jubin\Desktop\数据\车辆加油_管理(2026-10-11).xlsx'
SUMMARY = r'C:\Users\Jubin\Desktop\源数据\油耗分析汇总.xlsm'

PLATE = '粤BT3S56'
TARGET = os.path.join(TPL_DIR, 'D13-4.2米-粤BT3S56.xlsx')
UID_FIX = '261011000213'          # 油量要改
UID_ALSO = '261010000203'         # 同一批次另一行，回滚后一并重写
NEW_VOL = 64.33
BAD_VOL = 6433.0

BAK_ROOT = os.path.join(HERE, 'backup_fix_bt3s56_1011')
BAK_TPL_SRC = os.path.join(HERE, 'backup_daily', '车辆油耗计算模板', 'D13-4.2米-粤BT3S56.xlsx')
BAK_TPL_LOCAL = os.path.join(BAK_ROOT, 'D13_回滚源.xlsx')

CHECK = '--check' in sys.argv


def tpl_src():
    return BAK_TPL_LOCAL if os.path.exists(BAK_TPL_LOCAL) else BAK_TPL_SRC


def disp(v):
    s = str(v or '').strip()
    return s[:-2] if s.endswith('.0') else s


def main():
    if not os.path.exists(TARGET):
        print('模板不存在'); return 1
    os.makedirs(BAK_ROOT, exist_ok=True)
    if not CHECK:
        for p in (EXPORT, TARGET, KAOHE):
            shutil.copy2(p, BAK_ROOT)
        _s = tpl_src()
        if _s != BAK_TPL_LOCAL:
            shutil.copy2(_s, BAK_TPL_LOCAL)
        print('已备份 3 个文件；回滚源 =', _s)

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
        # ---- 1) 导出文件：改油量 ----
        wb = app.Workbooks.Open(EXPORT, 0, CHECK)      # 第三参数=ReadOnly，直接传 CHECK
        print('导出 ReadOnly =', wb.ReadOnly)
        ws = wb.Worksheets(1)
        ur = ws.UsedRange
        last = ur.Row + ur.Rows.Count - 1
        hit = None
        for r in range(2, last + 1):
            if disp(ws.Cells(r, 1).Value) == UID_FIX:
                hit = r
                print('  行%d 车牌=%s 金额=%s 油量=%s 单价=%s'
                      % (r, ws.Cells(r, 2).Value, ws.Cells(r, 6).Value,
                         ws.Cells(r, 7).Value, ws.Cells(r, 8).Value))
                if not CHECK:
                    ws.Cells(r, 7).Value = NEW_VOL
                break
        if hit is None:
            print('  !! 未找到'); ok = False
        else:
            if not CHECK:
                wb.Save()
            print('  -> 油量 %.0f => %.2f' % (BAD_VOL, NEW_VOL))
        wb.Close(False)

        # ---- 2) 回滚源校验 ----
        src = tpl_src()
        from python_calamine import CalamineWorkbook
        wb2 = CalamineWorkbook.from_path(src)
        d = wb2.get_sheet_by_name('加油数据').to_python(skip_empty_area=False)
        ids = {disp(r[0]) for r in d if r and r[0]}
        print('回滚源 %s : 加油数据=%d 行, 含今日两行=%s'
              % (os.path.basename(src), len(d), sorted(ids & {UID_FIX, UID_ALSO}) or '无'))
        if ids & {UID_FIX, UID_ALSO}:
            print('  !! 快照含今日行，不能作为回滚源'); ok = False

        # ---- 3) 考核表：删今日 2 行 ----
        wb4 = app.Workbooks.Open(KAOHE, 0, CHECK)
        print('考核表 ReadOnly =', wb4.ReadOnly)
        ws4 = wb4.Worksheets('每日油耗数据')
        ur4 = ws4.UsedRange
        last4 = ur4.Row + ur4.Rows.Count - 1
        found = []
        for r in range(last4, 1, -1):
            uid = disp(ws4.Cells(r, 1).Value)
            if uid in (UID_FIX, UID_ALSO):
                plate = str(ws4.Cells(r, 4).Value or '').strip()
                print('  命中 行%d 编号=%s 车牌=%s 升=%s 油耗=%s'
                      % (r, uid, plate, ws4.Cells(r, 10).Value, ws4.Cells(r, 18).Value))
                if plate == PLATE:
                    found.append((r, uid))
        for r, uid in sorted(found, reverse=True):
            if not CHECK:
                ws4.Rows(r).Delete()
            print('  -> 删除行 %d (%s)' % (r, uid))
        if found and not CHECK:
            wb4.Save()
        wb4.Close(False)

        # ---- 4) 汇总残留检查 ----
        wb5 = app.Workbooks.Open(SUMMARY, 0, True)
        ws5 = wb5.Worksheets(1)
        ur5 = ws5.UsedRange
        last5 = ur5.Row + ur5.Rows.Count - 1
        vals = ws5.Range(ws5.Cells(2, 1), ws5.Cells(last5, 14)).Value
        if not isinstance(vals, tuple):
            vals = ((vals,),)
        bad = [(i + 2, v) for i, v in enumerate(vals)
               if isinstance(v, tuple) and PLATE in str(v[2]) and str(v[8]).startswith('6433')]
        print('汇总含 6433 的 %s 行: %s' % (PLATE, [b[0] for b in bad] or '无（正常）'))
        wb5.Close(False)
    finally:
        try: app.Quit()
        except Exception: pass
        pythoncom.CoUninitialize()

    if not CHECK and ok:
        print('=== 回滚模板 ===')
        shutil.copy2(tpl_src(), TARGET)
        print('  已用快照覆盖 D13-4.2米-粤BT3S56.xlsx')

    # ---- 只读复核（铁律 8） ----
    from python_calamine import CalamineWorkbook
    wb = CalamineWorkbook.from_path(EXPORT)
    rows = wb.get_sheet_by_name(wb.sheet_names[0]).to_python(skip_empty_area=False)
    for r in rows:
        if r and disp(r[0]) == UID_FIX:
            print('复核 导出: 金额=%s 油量=%s 单价=%s' % (r[5], r[6], r[7]))
            if not CHECK and abs(float(r[6]) - NEW_VOL) > 1e-6:
                ok = False; print('  !! 导出未落盘')
    wb2 = CalamineWorkbook.from_path(KAOHE)
    rows2 = wb2.get_sheet_by_name('每日油耗数据').to_python(skip_empty_area=False)
    left = [i for i, r in enumerate(rows2, 1) if i > 1 and r and disp(r[0]) in (UID_FIX, UID_ALSO)]
    print('复核 考核表: 残留行=%s' % (left or '无（正常）'))
    if not CHECK and left:
        ok = False; print('  !! 考核表未落盘')
    if not CHECK:
        wb3 = CalamineWorkbook.from_path(TARGET)
        d3 = wb3.get_sheet_by_name('加油数据').to_python(skip_empty_area=False)
        ids3 = {disp(r[0]) for r in d3 if r and r[0]}
        print('复核 模板: 加油数据=%d 行, 含今日两行=%s' % (len(d3), sorted(ids3 & {UID_FIX, UID_ALSO}) or '无（正常）'))

    if CHECK:
        print('\n[--check] 未做任何写操作')
    else:
        print('\n完成。下一步：重跑主链路 python fleet_fuel_com.py "%s"' % EXPORT)
    return 0 if ok else 2


if __name__ == '__main__':
    sys.exit(main())
