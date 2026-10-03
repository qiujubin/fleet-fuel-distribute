# -*- coding: utf-8 -*-
"""修正 2026-10-03 两处司机单价笔误（误分流 + 单价字段错）

证据（金额 ÷ 油量 = 真实单价）：
  粤BNT993 261002000027  金额1928.43 / 230.95L = 8.35  ← 单价录成 0.35（漏了 8）
     → 单价<5 被判为尿素，误分到「尿素B14-9.6米-粤BNT993.xlsx」+「每日尿素数据」
  京LPW138 261002000030  金额588.29 / 71.05L = 8.28   ← 单价录成 588.29（把金额填进单价栏）
     → 未误分流（588>5），但模板/考核行单价列错

处置（标准「误分流修正」流程）：
  ① 导出文件改单价（源头修正，以后重跑不再误分）
  ② 从 backup_daily 回滚两个模板（回到今天写入前的干净状态）
  ③ 考核表删掉误入的 2 行（都在各自 sheet 末行）
  ④ 汇总同步「当日油价」列
  ⑤ 之后重跑 fleet_fuel_com.py，让这两条以正确单价重新落位

写入一律 WPS COM（禁止 openpyxl）。--check 只定位不写。
"""
import os, sys, shutil, glob, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
EXPORT = r'C:\Users\Jubin\Desktop\数据\车辆加油_管理(2026-10-03).xlsx'
TPL = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板'
KAOHE = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
SUMMARY = r'C:\Users\Jubin\Desktop\源数据\油耗分析汇总.xlsm'
BAK = os.path.join(HERE, 'backup_daily', '车辆油耗计算模板')

ROLLBACK = ['尿素B14-9.6米-粤BNT993.xlsx', 'D17-4.2米-京LPW138.xlsx']

FIX_PRICE = {          # 导出文件：系统编号 -> 正确单价
    '261002000027': 8.35,      # 粤BNT993（10-02 11:33）
    '261002000030': 8.28,      # 京LPW138（10-02 14:37）
}
SUMPRICE = 8.28        # 汇总 京LPW138(10-02) 的当日油价
CHECK = '--check' in sys.argv


def com_open(app, path, readonly=False):
    return app.Workbooks.Open(path, 0, bool(readonly))


def main():
    print('=' * 62)
    print('  修正 10-03 两处司机单价笔误' + ('   [--check 不写入]' if CHECK else ''))
    print('=' * 62)

    # ---------- ① 导出文件 ----------
    print('\n① 导出文件改单价')
    for sid, price in FIX_PRICE.items():
        print(f'   {sid} -> 单价 {price}')

    # ---------- ② 模板回滚 ----------
    print('\n② 模板回滚（从 backup_daily）')
    for fn in ROLLBACK:
        src = os.path.join(BAK, fn)
        if not os.path.exists(src):
            print(f'   ✗ 备份缺失: {src}'); return 2
        print(f'   {fn}  <- 备份 ({os.path.getsize(src):,} B)')

    # ---------- ③④ 考核表 / 汇总 ----------
    print('\n③ 考核表删除误入行（定位在各自的末行）')
    print('   每日尿素数据: 删末行（261002000027 粤BNT993）')
    print('   每日油耗数据: 删末行（261002000030 京LPW138）')
    print('\n④ 汇总: 京LPW138(10-02) 行的「当日油价」列 = %.2f' % SUMPRICE)

    if CHECK:
        print('\n(--check 模式，未做任何写入)')
        return 0

    # 备份（已由外部完成，这里再确认一次）
    if not os.path.isdir(os.path.join(HERE, 'backup_fix_nt993_lpw138_1003')):
        print('\n✗ 未找到事前备份目录，中止'); return 3

    # ---- 回滚文件（文件级，不经 COM）----
    if not CHECK:
        for fn in ROLLBACK:
            shutil.copy2(os.path.join(BAK, fn), os.path.join(TPL, fn))
        print('\n② 模板已回滚')

    import pythoncom
    import win32com.client as wc
    pythoncom.CoInitialize()
    app = None
    try:
        for pid in ('Ket.Application', 'et.Application'):
            try:
                app = wc.DispatchEx(pid); break
            except Exception:
                continue
        if app is None:
            raise RuntimeError('无法启动 WPS COM')
        app.Visible = False
        app.DisplayAlerts = False
        try: app.ScreenUpdating = False
        except Exception: pass

        # ---- ① 导出文件 ----
        wb = com_open(app, EXPORT)
        if wb.ReadOnly:
            raise RuntimeError('导出文件被占用（只读）')
        ws = wb.Worksheets(1)
        n = 0
        for r in range(2, ws.UsedRange.Rows.Count + 1):
            sid = str(ws.Cells(r, 1).Value or '').strip()
            sid = sid[:-2] if sid.endswith('.0') else sid
            if sid in FIX_PRICE:
                old = ws.Cells(r, 8).Value
                ws.Cells(r, 8).Value = FIX_PRICE[sid]
                print(f'  ① 导出 行{r} {sid} 单价 {old} -> {FIX_PRICE[sid]}')
                n += 1
        wb.Save(); wb.Close(False)
        print(f'  ① 完成，改动 {n} 处')

        # ---- ③ 考核表 ----
        wb = com_open(app, KAOHE)
        if wb.ReadOnly:
            raise RuntimeError('考核表被占用（只读）')
        for sheet, plate, sid in (('每日尿素数据', '粤BNT993', '261002000027'),
                                  ('每日油耗数据', '京LPW138', '261002000030')):
            ws = wb.Worksheets(sheet)
            last = _true_last_row(ws, 1)
            got = str(ws.Cells(last, 1).Value or '').strip()
            got = got[:-2] if got.endswith('.0') else got
            got_plate = str(ws.Cells(last, 4).Value or '').strip()
            if got != sid or got_plate != plate:
                raise RuntimeError(f'{sheet} 末行校验失败: 期望末行={sid}/{plate}, 实际={got}/{got_plate}')
            ws.Rows(last).Delete()
            print(f'  ③ {sheet} 删除末行 {last}（{got} {got_plate}）')
        wb.Save(); wb.Close(False)

        # ---- ④ 汇总 ----
        wb = com_open(app, SUMMARY)
        if wb.ReadOnly:
            raise RuntimeError('汇总被占用（只读）')
        ws = wb.Worksheets('Sheet1')
        last = _true_last_row(ws, 1)
        hit = None
        for r in range(2, last + 1):
            if str(ws.Cells(r, 3).Value or '').strip() == '京LPW138' and \
               str(ws.Cells(r, 7).Value or '')[:10] == '2026-10-02':
                hit = r
        if hit:
            old = ws.Cells(hit, 16).Value
            ws.Cells(hit, 16).Value = SUMPRICE
            print(f'  ④ 汇总 行{hit} 当日油价 {old} -> {SUMPRICE}')
        else:
            print('  ④ 汇总未找到京LPW138(10-02) 行，跳过')
        wb.Save(); wb.Close(False)

        print('\n✅ 修正完成。下一步：重跑 fleet_fuel_com.py 让记录重新落位。')
        return 0
    finally:
        if app is not None:
            try: app.Quit()
            except Exception: pass
        try: pythoncom.CoUninitialize()
        except Exception: pass


def _true_last_row(ws, col=1):
    ur = ws.UsedRange
    end = min(ur.Row + ur.Rows.Count - 1, ws.Rows.Count)
    while end >= 1:
        start = max(1, end - 1999)
        vals = ws.Range(ws.Cells(start, col), ws.Cells(end, col)).Value2
        if vals is None:
            end = start - 1; continue
        if not isinstance(vals, tuple): vals = ((vals,),)
        for i in range(len(vals) - 1, -1, -1):
            v = vals[i][0] if isinstance(vals[i], tuple) else vals[i]
            if v not in (None, ''):
                return start + i
        end = start - 1
    return 0


if __name__ == '__main__':
    sys.exit(main())
