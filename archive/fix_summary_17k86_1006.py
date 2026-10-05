# -*- coding: utf-8 -*-
"""修正汇总里 粤B17K86（261004000072）那条 GPS 缺失时期写入的坏值行。

现象：汇总行565 行驶里程 -338169.3、油耗 -0.024、被误判成"达标"。
根因：写入当时加油数据缺 GPS 里程 → 模板 N 取 0 → P = 0 - L。
现在源数据已补（GPS 339973.3），考核表行3462 已是正确值（1890.53 / 228.16 / 1804.0 / 12.647），
但汇总不会自动回更（宏只写新记录）。

做法：从考核表取该周期正确值 → 按汇总既有算法重算 → 写回。
  达标率 = 1 - (升数 - 里程×标准/100) / (里程×标准/100)

写入一律 WPS COM。--check 只看不写。
"""
import os, sys, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
KAOHE = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
SUMMARY = r'C:\Users\Jubin\Desktop\源数据\油耗分析汇总.xlsm'
BAKDIR = os.path.join(HERE, 'backup_fix_17k86_1006')

# (系统编号, 车牌, 汇总里对应的时间B 前缀)
TARGET = ('261004000072', '粤B17K86', '2026-10-04')
CHECK = '--check' in sys.argv


def disp(v):
    s = str(v or '').strip()
    return s[:-2] if s.endswith('.0') else s


def main():
    sid, plate, dayB = TARGET
    print('=' * 64)
    print(f'  修正汇总 {plate}（{sid}）的负里程行' + ('   [--check]' if CHECK else ''))
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

        # ---- 1. 考核表取正确值 ----
        wb = app.Workbooks.Open(KAOHE, 0, True)
        ws = wb.Worksheets('每日油耗数据')
        last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        g = None
        for r in range(2, last + 1):
            if disp(ws.Cells(r, 1).Value) == sid:
                g = dict(amt=float(ws.Cells(r, 9).Value or 0),
                         vol=float(ws.Cells(r, 10).Value or 0),
                         price=float(ws.Cells(r, 11).Value or 0),
                         km=float(ws.Cells(r, 16).Value or 0),
                         tB=str(ws.Cells(r, 7).Value or '')[:10])
                break
        wb.Close(False)
        if not g:
            print(f'  ✗ 考核表未找到 {sid}'); return 2
        print(f'  考核表: 金额={g["amt"]} 升={g["vol"]} 里程={g["km"]} 单价={g["price"]} 时间B={g["tB"]}')

        # ---- 2. 备份 + 打开汇总 ----
        if not CHECK:
            os.makedirs(BAKDIR, exist_ok=True)
            shutil.copy2(SUMMARY, BAKDIR)
        wb = app.Workbooks.Open(SUMMARY, 0, bool(CHECK))
        if not CHECK and wb.ReadOnly:
            raise RuntimeError('汇总被占用（只读）')
        ws = wb.Worksheets('Sheet1')
        last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1

        hit = None
        for r in range(2, last + 1):
            if str(ws.Cells(r, 3).Value or '').strip() == plate and \
               str(ws.Cells(r, 7).Value or '')[:10] == dayB:
                hit = r; break
        if not hit:
            print(f'  ✗ 汇总未找到 {plate} / {dayB}'); return 3

        km, vol, amt = g['km'], g['vol'], g['amt']
        std = float(ws.Cells(hit, 13).Value or 14.0)
        std_vol = km * std / 100.0
        rate = 1 - (vol - std_vol) / std_vol
        excess = vol - std_vol
        ok = rate >= 1.0
        print(f'  汇总 行{hit} 现值: 金额={ws.Cells(hit,8).Value} 升={ws.Cells(hit,9).Value} '
              f'里程={ws.Cells(hit,10).Value} 油耗={ws.Cells(hit,11).Value}')
        print(f'  → 改为: 金额={amt} 升={vol} 里程={km:.1f} 油耗={vol/km*100:.4f} '
              f'达标率={rate:.6f} {"达标" if ok else "不达标"}')
        if not CHECK:
            ws.Cells(hit, 8).Value = amt
            ws.Cells(hit, 9).Value = vol
            ws.Cells(hit, 10).Value = km
            ws.Cells(hit, 11).Value = vol / km * 100
            ws.Cells(hit, 12).Value = '达标' if ok else '不达标'
            ws.Cells(hit, 15).Value = rate
            ws.Cells(hit, 16).Value = '未超出' if ok else g['price']
            ws.Cells(hit, 17).Value = '未超出' if ok else excess
            ws.Cells(hit, 18).Value = '未超出' if ok else excess * g['price']
            wb.Save()
            print('  ✅ 已保存（WPS 原生）')
        else:
            print('  (--check 模式，未写入)')
        wb.Close(False)
        return 0
    finally:
        if app is not None:
            try: app.Quit()
            except Exception: pass
        try: pythoncom.CoUninitialize()
        except Exception: pass


if __name__ == '__main__':
    sys.exit(main())
