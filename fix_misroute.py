# -*- coding: utf-8 -*-
"""误分流修正通用工具（单价被录错 -> 被"单价<5 归尿素"规则误分到尿素链路）

用法：
  python fix_misroute.py --sid 261018000360 --plate 粤BT802F --price 8.28 \
      --export "C:\\...\\车辆加油_管理(2026-10-19).xlsx" \
      --urea-file "尿素D14-4.2米-粤BT802F.xlsx" [--check]

做四件事（按顺序，任一步校验失败即停）：
  1. 把 backup_daily 里的尿素模板快照拷到 backup_fix_misroute_<日期>/ 作为回滚源
  2. 改导出文件该行的单价
  3. 用回滚源覆盖尿素模板（去掉误入行）
  4. 删考核表「每日尿素数据」里该编号的行（编号+车牌双校验）
全部写操作后用 calamine 只读回读复核（不用 COM 复核，避免句柄冲突）。
"""
import os, sys, shutil, argparse, datetime
from python_calamine import CalamineWorkbook
import pythoncom, win32com.client as wc

HERE = os.path.dirname(os.path.abspath(__file__))
TPL_DIR = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板'
KAOHE = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
SNAP_DIR = os.path.join(HERE, 'backup_daily', '车辆油耗计算模板')


def disp(v):
    s = str(v or '').strip()
    return s[:-2] if s.endswith('.0') else s


def sheet_rows(path, sheet):
    return CalamineWorkbook.from_path(path).get_sheet_by_name(sheet).to_python(skip_empty_area=False)


def com_app():
    pythoncom.CoInitialize()
    for pid in ('Ket.Application', 'et.Application'):
        try:
            app = wc.DispatchEx(pid)
            app.Visible = False
            app.DisplayAlerts = False
            return app
        except Exception:
            continue
    raise SystemExit('无法启动 WPS COM')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--sid', required=True)
    ap.add_argument('--plate', required=True)
    ap.add_argument('--price', type=float, required=True)
    ap.add_argument('--export', required=True)
    ap.add_argument('--urea-file', required=True)
    ap.add_argument('--check', action='store_true')
    a = ap.parse_args()

    ok = True
    today = datetime.date.today().isoformat()
    bak_dir = os.path.join(HERE, 'backup_fix_misroute_' + today.replace('-', ''))
    urea_live = os.path.join(TPL_DIR, a.urea_file)
    urea_snap = os.path.join(SNAP_DIR, a.urea_file)

    # ---------- 0. 预检 ----------
    print('=== 预检 ===')
    rows = sheet_rows(a.export, CalamineWorkbook.from_path(a.export).sheet_names[0])
    exp_row, exp_price = None, None
    for i, r in enumerate(rows, 1):
        if r and disp(r[0]) == a.sid:
            exp_row, exp_price = i, r[7]
            print('  导出 行%d: 车牌=%s 金额=%s 油量=%s 单价=%s -> %.2f'
                  % (i, r[1], r[5], r[6], r[7], a.price))
            break
    if exp_row is None:
        print('  ✗ 导出文件里没有 %s' % a.sid); ok = False

    if not os.path.exists(urea_snap):
        print('  ✗ 回滚源不存在:', urea_snap); ok = False
    else:
        snap_rows = sheet_rows(urea_snap, '加油数据')
        if any(disp(r[0]) == a.sid for r in snap_rows if r):
            print('  ✗ 回滚源已含该编号（快照被污染，不能用）'); ok = False
        else:
            print('  回滚源 OK: 加油数据 %d 行，不含 %s' % (len(snap_rows), a.sid))

    live_rows = sheet_rows(urea_live, '加油数据')
    live_hit = [i for i, r in enumerate(live_rows, 1) if r and disp(r[0]) == a.sid]
    print('  尿素模板 现网 %d 行，含该编号: %s' % (len(live_rows), live_hit or '无'))

    kao_rows = sheet_rows(KAOHE, '每日尿素数据')
    kao_hit = [i for i, r in enumerate(kao_rows, 1) if i > 1 and r and disp(r[0]) == a.sid]
    print('  每日尿素数据 含该编号: %s %s'
          % (kao_hit or '无', [(i, kao_rows[i - 1][3]) for i in kao_hit]))

    if not ok:
        print('RESULT: FAIL（预检不通过，未做任何修改）')
        return 1
    if a.check:
        print('RESULT: OK（--check 模式，未做修改）')
        return 0

    # ---------- 备份 ----------
    os.makedirs(bak_dir, exist_ok=True)
    shutil.copy2(a.export, os.path.join(bak_dir, os.path.basename(a.export)))
    shutil.copy2(urea_live, os.path.join(bak_dir, '尿素_改动前.xlsx'))
    shutil.copy2(KAOHE, os.path.join(bak_dir, '油耗考核_改动前.xlsm'))
    shutil.copy2(urea_snap, os.path.join(bak_dir, '尿素_回滚源.xlsx'))
    print('=== 备份 ->', bak_dir)

    # ---------- 1. 改导出单价 ----------
    app = com_app()
    try:
        wb = app.Workbooks.Open(a.export, 0, False)
        ws = wb.Worksheets(1)
        ws.Cells(exp_row, 8).Value = a.price
        wb.Save()
        wb.Close(False)
    finally:
        app.Quit(); pythoncom.CoUninitialize()

    chk = sheet_rows(a.export, CalamineWorkbook.from_path(a.export).sheet_names[0])
    v = float(chk[exp_row - 1][7])
    print('  导出 行%d 单价 = %s %s' % (exp_row, v, '✓' if abs(v - a.price) < 1e-9 else '✗ 未落盘'))
    if abs(v - a.price) > 1e-9:
        print('RESULT: FAIL'); return 1

    # ---------- 2. 回滚尿素模板 ----------
    shutil.copy2(urea_snap, urea_live)
    chk2 = sheet_rows(urea_live, '加油数据')
    gone = not any(disp(r[0]) == a.sid for r in chk2 if r)
    print('  尿素模板 已回滚: %d 行，含该编号=%s %s'
          % (len(chk2), '否' if gone else '是', '✓' if gone else '✗'))
    if not gone:
        print('RESULT: FAIL'); return 1

    # ---------- 3. 删考核表尿素误入行 ----------
    if kao_hit:
        row = kao_hit[-1]
        app = com_app()
        try:
            wb = app.Workbooks.Open(KAOHE, 0, False)
            if wb.ReadOnly:
                raise SystemExit('考核表只读（可能被 WPS 打开）')
            wsu = wb.Worksheets('每日尿素数据')
            if str(wsu.Cells(row, 4).Value or '').strip() != a.plate:
                raise SystemExit('车牌校验失败: %r' % wsu.Cells(row, 4).Value)
            if disp(wsu.Cells(row, 1).Value) != a.sid:
                raise SystemExit('编号校验失败（行可能已位移）')
            wsu.Rows(row).Delete()
            wb.Save()
            wb.Close(False)
        finally:
            app.Quit(); pythoncom.CoUninitialize()
        chk3 = sheet_rows(KAOHE, '每日尿素数据')
        still = any(i > 1 and r and disp(r[0]) == a.sid for i, r in enumerate(chk3, 1))
        print('  每日尿素数据 删行%d: %s' % (row, '✗ 仍存在' if still else '✓ 已落盘'))
        if still:
            print('RESULT: FAIL'); return 1
    else:
        print('  每日尿素数据 无该行，跳过')

    print('RESULT: OK  下一步请重跑主链路：')
    print('  python fleet_fuel_com.py "%s"' % a.export)
    return 0


if __name__ == '__main__':
    sys.exit(main())
