# -*- coding: utf-8 -*-
"""把考核表里 11-04 的 6 条补进汇总（宏只认区域表，补的行不在区域表里带不过去）"""
import os, sys, shutil
import pythoncom
import win32com.client as wc
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from region_map import REGION_MAP

HERE = os.path.dirname(os.path.abspath(__file__))
KAOHE = r"C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm"
SUM = r"C:\Users\Jubin\Desktop\源数据\油耗分析汇总.xlsm"
CHECK = "--check" in sys.argv

SIDS = ["261104000111", "261104000125", "261104000128",
        "261104000130", "261104000135", "261104000136"]

PLATE2REG = {p: reg for reg, ps in REGION_MAP.items() for p in ps}


def main():
    pythoncom.CoInitialize()
    app = None
    try:
        for pid in ("Ket.Application", "et.Application"):
            try:
                app = wc.DispatchEx(pid)
                app.Visible = False
                app.DisplayAlerts = False
                break
            except Exception:
                continue
        if app is None:
            raise SystemExit("无法启动 WPS COM")

        # 1. 从考核表读这 6 条
        wb = app.Workbooks.Open(KAOHE, 0, True)
        ws = wb.Worksheets("每日油耗数据")
        last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        recs = []
        for r in range(2, last + 1):
            sid = str(ws.Cells(r, 1).Value or "").strip().split(".")[0]
            if sid in SIDS:
                recs.append(dict(
                    sid=sid,
                    plate=str(ws.Cells(r, 4).Value or "").strip(),
                    vtype=str(ws.Cells(r, 5).Value or "").strip(),
                    drv=str(ws.Cells(r, 3).Value or "").strip(),
                    tA=ws.Cells(r, 6).Value2, tB=ws.Cells(r, 7).Value2,
                    amt=float(ws.Cells(r, 9).Value or 0),
                    vol=float(ws.Cells(r, 10).Value or 0),
                    price=float(ws.Cells(r, 11).Value or 0),
                    km=float(ws.Cells(r, 16).Value or 0),
                    oil=float(ws.Cells(r, 18).Value or 0),
                ))
        wb.Close(False)
        print(f"考核表读到 {len(recs)} 条")
        if len(recs) != len(SIDS):
            print("  ! 数量不符")

        # 2. 写汇总
        wb = app.Workbooks.Open(SUM, 0, bool(CHECK))
        if not CHECK and wb.ReadOnly:
            raise SystemExit("汇总被占用")
        ws = wb.Worksheets("Sheet1")
        last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        existing = set()
        std_hist = {}
        for r in range(2, last + 1):
            plate = str(ws.Cells(r, 3).Value or "").strip()
            g = ws.Cells(r, 7).Value2
            try:
                gv = round(float(g), 6)
            except (TypeError, ValueError):
                gv = str(g)[:19]
            existing.add((plate, gv))
            v = ws.Cells(r, 13).Value
            if isinstance(v, (int, float)) and v > 0:
                std_hist[plate] = float(v)      # 取该车最后一次的标准值
        print(f"汇总现有 {last-1} 行")

        recs.sort(key=lambda x: x["tB"] or 0)
        row = last
        n = 0
        for d in recs:
            key = (d["plate"], round(float(d["tB"] or 0), 6))
            if key in existing:
                print(f"  跳过已存在 {d['plate']}")
                continue
            # 标准油耗：沿用该车历史标准，没有则按车型兜底
            vt = d["vtype"]
            std = std_hist.get(d["plate"]) or (14.0 if vt.startswith("4.2") else (22.0 if vt.startswith("7.6") else 26.0))
            std_vol = d["km"] * std / 100.0
            rate = 1 - (d["vol"] - std_vol) / std_vol if std_vol else 0
            ok = rate >= 1.0
            excess = d["vol"] - std_vol
            row += 1
            n += 1
            print(f"  + 行{row}: {d['plate']} {d['vtype']} 升={d['vol']:.2f} 里程={d['km']:.1f} "
                  f"油耗={d['vol']/d['km']*100:.2f} 标准={std:.0f} {'达标' if ok else '不达标'}")
            if CHECK:
                continue
            ws.Cells(row, 1).Value = row - 1
            ws.Cells(row, 2).Value = PLATE2REG.get(d["plate"], "")
            ws.Cells(row, 3).Value = d["plate"]
            ws.Cells(row, 4).Value = vt
            ws.Cells(row, 5).Value = d["drv"]
            ws.Cells(row, 6).Value2 = d["tA"]
            ws.Cells(row, 7).Value2 = d["tB"]
            ws.Cells(row, 8).Value = d["amt"]
            ws.Cells(row, 9).Value = d["vol"]
            ws.Cells(row, 10).Value = d["km"]
            ws.Cells(row, 11).Value = d["vol"] / d["km"] * 100
            ws.Cells(row, 12).Value = "达标" if ok else "不达标"
            ws.Cells(row, 13).Value = std
            ws.Cells(row, 14).Value2 = int(d["tB"])
            ws.Cells(row, 15).Value = rate
            ws.Cells(row, 16).Value = "未超出" if ok else d["price"]
            ws.Cells(row, 17).Value = "未超出" if ok else excess
            ws.Cells(row, 18).Value = "未超出" if ok else excess * d["price"]

        if CHECK:
            print("\n(--check 未写入)")
        else:
            if n:
                wb.Save()
            print(f"\n已写入 {n} 行")
        wb.Close(False)
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


if __name__ == "__main__":
    main()
