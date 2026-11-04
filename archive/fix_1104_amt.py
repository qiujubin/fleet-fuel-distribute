# -*- coding: utf-8 -*-
"""修正 粤BFE485 / 261104000136 金额小数点错位
录入 20065.51，实为 233.86L × 8.58 ≈ 2006.52（右移一位）
改：模板「加油数据」F 列 → 重算 I/J → 考核表 I 列
"""
import os, sys, shutil
import pythoncom
import win32com.client as wc
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
from com_formula import rebuild_ij

HERE = os.path.dirname(os.path.abspath(__file__))
TPL = r"C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板\B11-9.6米-粤BFE485.xlsx"
KAOHE = r"C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm"
SID = "261104000136"
NEW_AMT = 2006.52
CHECK = "--check" in sys.argv

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

    if not CHECK:
        bak = os.path.join(HERE, "backup_fix_1104")
        os.makedirs(bak, exist_ok=True)
        shutil.copy2(TPL, bak)
        shutil.copy2(KAOHE, bak)
        print("备份 ->", bak)

    # 1. 模板加油数据
    wb = app.Workbooks.Open(TPL, 0, bool(CHECK))
    if not CHECK and wb.ReadOnly:
        raise SystemExit("模板被占用")
    ws = wb.Worksheets("加油数据")
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    hit = None
    for r in range(2, last + 1):
        if str(ws.Cells(r, 1).Value or "").strip().split(".")[0] == SID:
            hit = r
            break
    print(f"模板[加油数据] 行{hit}")
    if hit is None:
        raise SystemExit("未找到该编号")
    print(f"  金额 {ws.Cells(hit,6).Value} -> {NEW_AMT}；油量={ws.Cells(hit,7).Value} 单价={ws.Cells(hit,8).Value}")
    if not CHECK:
        ws.Cells(hit, 6).Value = NEW_AMT

    # 2. 重算该行 I/J
    br = wb.Worksheets(1)
    lastb = br.UsedRange.Row + br.UsedRange.Rows.Count - 1
    brow = None
    for r in range(2, lastb + 1):
        if str(br.Cells(r, 1).Value or "").strip().split(".")[0] == SID:
            brow = r
            break
    print(f"计算模板 行{brow}")
    if not CHECK and brow:
        rebuild_ij(app, TPL, wb=wb, only_rows=[brow], verbose=False)
        br.Calculate()
        i_val = float(br.Cells(brow, 9).Value or 0)
        j_val = float(br.Cells(brow, 10).Value or 0)
        p_val = float(br.Cells(brow, 16).Value or 0)
        print(f"  重算后 I={i_val} J={j_val} P={p_val} R={j_val/p_val*100:.4f}")
        if abs(i_val - NEW_AMT) > 1:
            raise SystemExit(f"I 异常 {i_val}")
        wb.Save()
        print("  模板已保存")
    wb.Close(False)

    # 3. 考核表 I 列
    wb = app.Workbooks.Open(KAOHE, 0, bool(CHECK))
    if not CHECK and wb.ReadOnly:
        raise SystemExit("考核表被占用")
    ws = wb.Worksheets("每日油耗数据")
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    kr = None
    for r in range(2, last + 1):
        if str(ws.Cells(r, 1).Value or "").strip().split(".")[0] == SID:
            kr = r
            break
    print(f"考核表 行{kr}: I={ws.Cells(kr,9).Value} 升={ws.Cells(kr,10).Value} P={ws.Cells(kr,16).Value}")
    if kr and not CHECK:
        ws.Cells(kr, 9).Value = i_val
        p = float(ws.Cells(kr, 16).Value or 1)
        ws.Cells(kr, 19).Value = i_val / p
        wb.Save()
        print("  考核表已保存")
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
