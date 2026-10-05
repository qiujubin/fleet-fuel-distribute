# -*- coding: utf-8 -*-
"""修正 粤BT802F 261005000084（10-05 02:18）的 GPS 里程「少一位」错误。

现象：模板 GPS 里程 = -462120.7、油耗 -0.0318（负值）。
根因：加油数据里这条的 GPS里程 录成 **51461.1**，而前一条是 514184.6。
      → 明显少了一位数字。按仪表增量（351958→352386，+428）反推应为
        **514611.1**（与 514184.6+428=514612.6 仅差 1.5 km）。
      → 修正后里程 1029.3 km、油耗 14.28 L/100km（4.2米标准 14，合理）

处置：
  ① 导出文件该条 GPS里程 改 514611.1（源头，保持与修正后一致）
  ② D14 加油数据 末行 GPS里程 改 514611.1（模板公式会自动重算）
  ③ 读模板该行 L~V 列，回写考核表 261005000084 行的 12~22 列

写入一律 WPS COM。--check 只定位不写。
"""
import os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
EXPORT = r'C:\Users\Jubin\Desktop\数据\车辆加油_管理(2026-10-05).xlsx'
TPL = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板\D14-4.2米-粤BT802F.xlsx'
KAOHE = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
SID = '261005000084'
BAD_GPS, GOOD_GPS = 51461.1, 514611.1
CHECK = '--check' in sys.argv


def disp(v):
    s = str(v or '').strip()
    return s[:-2] if s.endswith('.0') else s


def main():
    print('=' * 62)
    print('  修正 粤BT802F %s 的 GPS 里程少一位' % SID)
    print('  %.1f -> %.1f' % (BAD_GPS, GOOD_GPS))
    print('=' * 62)

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
        app.Visible = False; app.DisplayAlerts = False

        # ① 导出文件（GPS里程 = 第5列）
        wb = app.Workbooks.Open(EXPORT, 0, bool(CHECK))
        ws = wb.Worksheets(1)
        hit = None
        for r in range(2, ws.UsedRange.Rows.Count + 1):
            if disp(ws.Cells(r, 1).Value) == SID:
                hit = r; break
        if not hit:
            print('  ✗ 导出文件未找到'); return 2
        old = ws.Cells(hit, 5).Value
        print('  ① 导出 行%d GPS里程 = %r' % (hit, old))
        if not CHECK:
            ws.Cells(hit, 5).Value = GOOD_GPS
            wb.Save()
        wb.Close(False)

        # ② 车辆模板加油数据
        wb = app.Workbooks.Open(TPL, 0, bool(CHECK))
        wsd = wb.Worksheets('加油数据')
        lastd = _last(wsd, 1)
        if disp(wsd.Cells(lastd, 1).Value) != SID:
            print('  ✗ D14 加油数据末行不是目标: %s' % wsd.Cells(lastd, 1).Value); return 3
        print('  ② D14 加油数据 行%d GPS里程 = %r -> %s'
              % (lastd, wsd.Cells(lastd, 5).Value, GOOD_GPS))
        if not CHECK:
            wsd.Cells(lastd, 5).Value = GOOD_GPS
            wb.Save()
            app.Calculate()

        # ③ 读模板该行 L~V（12~22 列），回写考核表
        wst = wb.Worksheets('计算模板')
        lastt = _last(wst, 7)
        vals = [wst.Cells(lastt, c).Value for c in range(12, 23)]
        print('  ③ 模板 行%d 重算后：GPS里程=%r 油耗=%r'
              % (lastt, vals[4], vals[6]))
        wb.Close(False)

        wb = app.Workbooks.Open(KAOHE, 0, bool(CHECK))
        ws = wb.Worksheets('每日油耗数据')
        last = _last(ws, 1)
        hit = None
        for r in range(2, last + 1):
            if disp(ws.Cells(r, 1).Value) == SID:
                hit = r; break
        if not hit:
            print('  ✗ 考核表未找到 %s' % SID); return 4
        print('      考核表 行%d 现值 GPS里程=%r 油耗=%r'
              % (hit, ws.Cells(hit, 16).Value, ws.Cells(hit, 18).Value))
        if not CHECK:
            for i, c in enumerate(range(12, 23)):
                ws.Cells(hit, c).Value = vals[i]
            wb.Save()
            print('      ✅ 已回写 12~22 列')
        print('      → 新值 GPS里程=%r 油耗=%r' % (ws.Cells(hit, 16).Value, ws.Cells(hit, 18).Value))
        wb.Close(False)

        print('\n(--check 模式，未写入)' if CHECK else '\n✅ 修正完成（WPS 原生保存）')
        return 0
    finally:
        if app is not None:
            try: app.Quit()
            except Exception: pass
        try: pythoncom.CoUninitialize()
        except Exception: pass


def _last(ws, col):
    ur = ws.UsedRange
    end = ur.Row + ur.Rows.Count - 1
    for r in range(end, 1, -1):
        if ws.Cells(r, col).Value not in (None, ''):
            return r
    return 0


if __name__ == '__main__':
    sys.exit(main())
