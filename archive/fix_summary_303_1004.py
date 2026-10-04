# -*- coding: utf-8 -*-
"""修正汇总里 京AEP303 两条「GPhS 里程缺失时期」写入的坏值行。

现象：行驶里程 = -683887.7 / -688978.8，油耗为负，判成"达标"。
原因：写入当时加油数据缺 GPS 里程 → 模板 N 列取 0 → P = 0 - L。
现在源数据已补齐，考核表里已是正确值，但汇总不会自动回更（宏只处理新记录）。

做法：从考核表取该周期的正确值 → 按汇总既有公式重算 → 写回汇总。
  达标率 = 1 - (升数 - 里程×标准/100) / (里程×标准/100)   ← 用同车其它行反推并双行验算通过
  超出标准油量 = 升数 - 里程×标准/100
  超出标准价格 = 超出标准油量 × 当日油价
  「当日油价」取该周期时间B那条的单价；达标时按既有约定写文本「未超出」

写入一律 WPS COM。--check 只看不写。
"""
import os, sys, shutil, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
KAOHE = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
SUMMARY = r'C:\Users\Jubin\Desktop\源数据\油耗分析汇总.xlsm'
BAKDIR = os.path.join(HERE, 'backup_fix_303_1004')

# (考核表里的系统编号, 汇总里对应的时间B 前缀)
TARGETS = [
    ('260927000575', '2026-09-27'),   # 09-21 08:50 -> 09-27 14:14
    ('261002000017', '2026-10-02'),   # 09-27 14:14 -> 10-02 06:20
]
CHECK = '--check' in sys.argv


def disp(v):
    s = str(v or '').strip()
    return s[:-2] if s.endswith('.0') else s


def main():
    print('=' * 64)
    print('  修正汇总 京AEP303 两条坏值行' + ('   [--check]' if CHECK else ''))
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
        app.Visible = False
        app.DisplayAlerts = False

        # ---- 1. 从考核表取正确值 ----
        wb = app.Workbooks.Open(KAOHE, 0, True)
        ws = wb.Worksheets('每日油耗数据')
        ur = ws.UsedRange; last = ur.Row + ur.Rows.Count - 1
        got = {}
        for r in range(2, last + 1):
            sid = disp(ws.Cells(r, 1).Value)
            for key, _ in TARGETS:
                if sid == key:
                    got[key] = dict(
                        plate=str(ws.Cells(r, 4).Value or '').strip(),
                        amt=float(ws.Cells(r, 9).Value or 0),
                        vol=float(ws.Cells(r, 10).Value or 0),
                        price=float(ws.Cells(r, 11).Value or 0),
                        km=float(ws.Cells(r, 16).Value or 0),       # GPS里程
                        tB=str(ws.Cells(r, 7).Value or '')[:10],
                    )
        wb.Close(False)
        for key, _ in TARGETS:
            if key not in got:
                print(f'  ✗ 考核表未找到 {key}'); return 2
            g = got[key]
            print(f'  考核表 {key}: {g["plate"]} 金额={g["amt"]} 升={g["vol"]} 里程={g["km"]} 单价={g["price"]} 时间B={g["tB"]}')

        # ---- 2. 备份并打开汇总 ----
        if not CHECK:
            os.makedirs(BAKDIR, exist_ok=True)
            shutil.copy2(SUMMARY, BAKDIR)
        wb = app.Workbooks.Open(SUMMARY, 0, bool(CHECK))
        if not CHECK and wb.ReadOnly:
            raise RuntimeError('汇总被占用（只读）')
        ws = wb.Worksheets('Sheet1')
        ur = ws.UsedRange; last = ur.Row + ur.Rows.Count - 1

        n = 0
        for key, dayB in TARGETS:
            hit = None
            for r in range(2, last + 1):
                if str(ws.Cells(r, 3).Value or '').strip() == '京AEP303' and \
                   str(ws.Cells(r, 7).Value or '')[:10] == dayB:
                    hit = r; break
            if not hit:
                print(f'  ✗ 汇总未找到 京AEP303 / {dayB}'); continue

            g = got[key]
            km, vol, amt = g['km'], g['vol'], g['amt']
            std = float(ws.Cells(hit, 13).Value or 26.0)          # 标准油耗
            std_vol = km * std / 100.0
            rate = 1 - (vol - std_vol) / std_vol                   # 达标率
            excess = vol - std_vol                                 # 超出标准油量
            ok = rate >= 1.0
            print(f'  汇总 行{hit}（{dayB}）→ 金额={amt} 升={vol} 里程={km:.1f} 油耗={vol/km*100:.4f} '
                  f'达标率={rate:.6f} {"达标" if ok else "不达标"} 超出={excess:.3f}L')
            if CHECK:
                continue
            ws.Cells(hit, 8).Value = amt
            ws.Cells(hit, 9).Value = vol
            ws.Cells(hit, 10).Value = km
            ws.Cells(hit, 11).Value = vol / km * 100
            ws.Cells(hit, 12).Value = '达标' if ok else '不达标'
            ws.Cells(hit, 15).Value = rate
            ws.Cells(hit, 16).Value = '未超出' if ok else g['price']
            ws.Cells(hit, 17).Value = '未超出' if ok else excess
            ws.Cells(hit, 18).Value = '未超出' if ok else excess * g['price']
            n += 1

        if CHECK:
            print('\n(--check 模式，未写入)')
        else:
            if n:
                wb.Save()
            print(f'\n✅ 已修正 {n} 行并保存（WPS 原生）')
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
