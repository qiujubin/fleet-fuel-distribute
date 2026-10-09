# -*- coding: utf-8 -*-
"""临时代偿：粤B17K86（261008000165，10-08 22:53）GPS 里程缺失 → 里程为负。

现象：汇总行601 / 考核行 GPS里程 = -342145.2、油耗 -0.0213、误判"达标"。
根因：接口里该条 GPS里程 = None（10-04 之后第二次），模板 N 列取 0 → P = 0 - L。

代偿做法（**推算值，非实测**）：用仪表增量推 GPS。
  上一个"是"：仪表 356260 / GPS 342145.2
  本条仪表 356841 → 增量 581
  → 推 GPS 342145.2 + 581 = 342726.2，于是 P = 581，与仪表口径一致
  → 油耗 = 72.83 / 581 × 100 = 12.54 L/100km（4.2米标准14，合理）

写入：加油数据 GPS 列 → 模板自动重算 → 回写考核行 12~22 → 修正汇总行。
**用户补真实 GPS 后应重跑该条覆盖。**

--check 只看不写。
"""
import os, sys, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
TPL = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板\D21-4.2米-粤B17K86.xlsx'
KAOHE = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
SUMMARY = r'C:\Users\Jubin\Desktop\源数据\油耗分析汇总.xlsm'
EXPORT = r'C:\Users\Jubin\Desktop\数据\车辆加油_管理(2026-10-09).xlsx'
BAKDIR = os.path.join(HERE, 'backup_fix_17k86_gps_1009')
SID = '261008000165'
GPS_FIX = 342726.2
CHECK = '--check' in sys.argv


def disp(v):
    s = str(v or '').strip()
    return s[:-2] if s.endswith('.0') else s


def main():
    print('=' * 64)
    print(f'  粤B17K86 {SID} GPS 缺失代偿  -> GPS {GPS_FIX}' + ('   [--check]' if CHECK else ''))
    print('=' * 64)

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
        app.Visible = False; app.DisplayAlerts = False

        if not CHECK:
            os.makedirs(BAKDIR, exist_ok=True)
            for p in (TPL, KAOHE, SUMMARY, EXPORT):
                shutil.copy2(p, BAKDIR)

        # ---- 1. 加油数据补 GPS ----
        wb = app.Workbooks.Open(TPL, 0, bool(CHECK))
        wsd = wb.Worksheets('加油数据')
        lastd = _last(wsd, 1)
        assert disp(wsd.Cells(lastd, 1).Value) == SID, '末行不是目标: %s' % wsd.Cells(lastd, 1).Value
        print(f'  加油数据 行{lastd}: 当前里程={wsd.Cells(lastd,4).Value} GPS={wsd.Cells(lastd,5).Value}')
        if not CHECK:
            wsd.Cells(lastd, 5).Value = GPS_FIX
            wb.Save(); app.Calculate()
        wst = wb.Worksheets('计算模板')
        lastt = _last(wst, 7)
        vals = [wst.Cells(lastt, c).Value for c in range(12, 23)]
        print(f'  模板 行{lastt}: GPS里程={vals[4]} 油耗={vals[6]}')
        wb.Close(False)
        if CHECK:
            print('\n(--check 模式，未写入)'); return 0

        # ---- 2. 回写考核行 ----
        wb = app.Workbooks.Open(KAOHE, 0, False)
        ws = wb.Worksheets('每日油耗数据')
        hit = None
        for r in range(2, _last(ws, 1) + 1):
            if disp(ws.Cells(r, 1).Value) == SID:
                hit = r; break
        if not hit:
            print('  ✗ 考核表未找到'); return 2
        for i, c in enumerate(range(12, 23)):
            ws.Cells(hit, c).Value = vals[i]
        wb.Save(); wb.Close(False)
        print(f'  考核表 行{hit} 已回写 12~22 列')

        # ---- 3. 修正汇总 ----
        wb = app.Workbooks.Open(SUMMARY, 0, False)
        ws = wb.Worksheets('Sheet1')
        hit = None
        for r in range(2, _last(ws, 1) + 1):
            if str(ws.Cells(r, 3).Value or '').strip() == '粤B17K86' and \
               str(ws.Cells(r, 7).Value or '')[:10] == '2026-10-08':
                hit = r; break
        if hit:
            amt = float(ws.Cells(hit, 8).Value or 0)
            vol = float(ws.Cells(hit, 9).Value or 0)
            km = float(vals[4]); std = float(ws.Cells(hit, 13).Value or 14.0)
            stdv = km * std / 100.0
            rate = 1 - (vol - stdv) / stdv
            ok = rate >= 1.0
            ws.Cells(hit, 10).Value = km
            ws.Cells(hit, 11).Value = vol / km * 100
            ws.Cells(hit, 12).Value = '达标' if ok else '不达标'
            ws.Cells(hit, 15).Value = rate
            ws.Cells(hit, 16).Value = '未超出' if ok else float(vals[10])
            ws.Cells(hit, 17).Value = '未超出' if ok else (vol - stdv)
            ws.Cells(hit, 18).Value = '未超出' if ok else (vol - stdv) * float(vals[10])
            print(f'  汇总 行{hit} → 里程={km:.1f} 油耗={vol/km*100:.4f} 达标率={rate:.6f} '
                  f'{"达标" if ok else "不达标"}')
        wb.Save(); wb.Close(False)
        print('\n✅ 代偿完成（推算值，非实测；补真实 GPS 后请重跑覆盖）')
        return 0
    finally:
        if app is not None:
            try: app.Quit()
            except Exception: pass
        try: pythoncom.CoUninitialize()
        except Exception: pass


def _last(ws, col):
    ur = ws.UsedRange
    for r in range(ur.Row + ur.Rows.Count - 1, 1, -1):
        if ws.Cells(r, col).Value not in (None, ''):
            return r
    return 0


if __name__ == '__main__':
    sys.exit(main())
