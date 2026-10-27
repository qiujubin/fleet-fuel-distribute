# -*- coding: utf-8 -*-
"""10-27 收尾修正
1) AEP303 10-25 那条：用户已把真实 GPS 补成 716766.7（模板/考核表已 1520.0 km / 25.7921 达标）
   -> 汇总行仍停在昨天兜底的 1469 / 26.69，按真实值重算
2) 打冷 3 条 BT802F 超长记录（6240/3060/2220 分钟）昨天被用户手工删除过，今天又被加回
   -> 从「打冷记录」删除（恢复用户意图）
--check 只读
"""
import os, sys, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
KAOHE = r"C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm"
SUM = r"C:\Users\Jubin\Desktop\源数据\油耗分析汇总.xlsm"
BAK = os.path.join(HERE, "backup_fix_1027")
CHECK = "--check" in sys.argv
DEL_SIDS = {"261020000051", "261022000057", "261024000058"}

import pythoncom
import win32com.client as wc

pythoncom.CoInitialize()
app = None
try:
    for pid in ("Ket.Application", "et.Application"):
        try:
            app = wc.DispatchEx(pid)
            break
        except Exception:
            continue
    if app is None:
        raise RuntimeError("无法启动 WPS COM")
    app.Visible = False
    app.DisplayAlerts = False

    if not CHECK:
        os.makedirs(BAK, exist_ok=True)
        for p in (KAOHE, SUM):
            shutil.copy2(p, BAK)
        print("备份 ->", BAK)

    # --- 1. 汇总 AEP303 10-25 ---
    wb = app.Workbooks.Open(SUM, 0, bool(CHECK))
    if not CHECK and wb.ReadOnly:
        raise RuntimeError("汇总被占用")
    ws = wb.Worksheets("Sheet1")
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    row = None
    for r in range(2, last + 1):
        if str(ws.Cells(r, 3).Value or "").strip() == "京AEP303" and str(ws.Cells(r, 7).Value or "")[:10] == "2026-10-25":
            row = r
            break
    if row is None:
        raise RuntimeError("汇总未找到 AEP303 10-25")
    vol = float(ws.Cells(row, 9).Value or 0)
    km = 1520.0
    std = float(ws.Cells(row, 13).Value or 26.0)
    price = float(ws.Cells(row, 16).Value) if isinstance(ws.Cells(row, 16).Value, (int, float)) else 8.29
    std_vol = km * std / 100.0
    rate = 1 - (vol - std_vol) / std_vol
    excess = vol - std_vol
    ok = rate >= 1.0
    print(f"汇总 行{row}: 现值 里程={ws.Cells(row,10).Value} 油耗={ws.Cells(row,11).Value}")
    print(f"  -> 里程={km} 油耗={vol/km*100:.4f} 达标率={rate:.6f} {'达标' if ok else '不达标'} 超出={excess:.2f}L")
    if not CHECK:
        ws.Cells(row, 10).Value = km
        ws.Cells(row, 11).Value = vol / km * 100
        ws.Cells(row, 12).Value = "达标" if ok else "不达标"
        ws.Cells(row, 15).Value = rate
        ws.Cells(row, 16).Value = "未超出" if ok else price
        ws.Cells(row, 17).Value = "未超出" if ok else excess
        ws.Cells(row, 18).Value = "未超出" if ok else excess * price
        wb.Save()
        print("  汇总已保存")
    wb.Close(False)

    # --- 2. 打冷记录：删 3 条超长 ---
    wb = app.Workbooks.Open(KAOHE, 0, bool(CHECK))
    if not CHECK and wb.ReadOnly:
        raise RuntimeError("考核表被占用")
    ws = wb.Worksheets("打冷记录")
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    hits = []
    for r in range(2, last + 1):
        if str(ws.Cells(r, 1).Value or "").strip().split(".")[0] in DEL_SIDS:
            hits.append(r)
    print(f"\n打冷记录 命中 {len(hits)} 行: {hits}")
    for r in reversed(hits):
        print(f"  行{r}: {ws.Cells(r,1).Value} {ws.Cells(r,2).Value} 时长={ws.Cells(r,4).Value} "
              f"{ws.Cells(r,7).Value}")
        if not CHECK:
            ws.Rows(r).Delete()
    if not CHECK and hits:
        wb.Save()
        print("  打冷记录已保存")
    wb.Close(False)
    print("\n完成" if not CHECK else "\n(--check 未写入)")
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
