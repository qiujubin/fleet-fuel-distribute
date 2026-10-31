# -*- coding: utf-8 -*-
"""2026-10-31 源数据单价错误处理

1) 粤BY2J27 261029000639 单价又被源系统给成 3.31（昨天改过、源系统没动）
   -> 今天重新取数时被分进「尿素D15」+ 考核表[每日尿素数据]，与加油侧重复
   -> 删尿素侧两处；加油侧昨天已正确（行184 I=1872.47 J=225.71）
2) 单价显示错误（不影响油耗结论，仅修正 K 列/导出）
   - 粤BFE485 261030000683：5.64 -> 8.64（500/57.88）
   - 粤BT3S56 261031000696：60   -> 8.60（446.94/51.97）
   - 粤BY2J27 261029000639：3.31 -> 8.31（600/72.21）
--check 只读
"""
import os, sys, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
TPL = r"C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板"
KAOHE = r"C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm"
EXP = r"C:\Users\Jubin\Desktop\数据\车辆加油_管理(2026-10-31).xlsx"
UREA = os.path.join(TPL, "尿素D15-4.2米-粤BY2J27.xlsx")
T_B11 = os.path.join(TPL, "B11-9.6米-粤BFE485.xlsx")
T_D13 = os.path.join(TPL, "D13-4.2米-粤BT3S56.xlsx")
T_D15 = os.path.join(TPL, "D15-4.2米-粤BY2J27.xlsx")
BAK = os.path.join(HERE, "backup_fix_1031")
CHECK = "--check" in sys.argv

PRICE_FIX = [
    ("261030000683", 8.64),
    ("261031000696", 8.60),
    ("261029000639", 8.31),
]

import pythoncom
import win32com.client as wc

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
        os.makedirs(BAK, exist_ok=True)
        for p in (KAOHE, EXP, UREA, T_B11, T_D13, T_D15):
            if os.path.exists(p):
                shutil.copy2(p, BAK)
        print("备份 ->", BAK)

    # ---------- 1. 导出单价 ----------
    wb = app.Workbooks.Open(EXP, 0, bool(CHECK))
    ws = wb.Worksheets(1)
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    print("导出:")
    for r in range(2, last + 1):
        sid = str(ws.Cells(r, 1).Value or "").strip().split(".")[0]
        for fsid, price in PRICE_FIX:
            if sid == fsid:
                print(f"  行{r} {sid}: 单价 {ws.Cells(r, 8).Value} -> {price}")
                if not CHECK:
                    ws.Cells(r, 8).Value = price
    if not CHECK:
        wb.Save()
    wb.Close(False)

    # ---------- 2. 加油模板 K 列 ----------
    for fp, items in ((T_B11, [("261030000683", 8.64)]),
                      (T_D13, [("261031000696", 8.60)]),
                      (T_D15, [("261029000639", 8.31)])):
        wb = app.Workbooks.Open(fp, 0, bool(CHECK))
        ws = wb.Worksheets(1)
        last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        for r in range(2, last + 1):
            sid = str(ws.Cells(r, 1).Value or "").strip().split(".")[0]
            for fsid, price in items:
                if sid == fsid:
                    k = ws.Cells(r, 11)
                    print(f"  {os.path.basename(fp)} 行{r} {sid}: K={k.Value} -> {price}"
                          f"（金额={ws.Cells(r,9).Value} 升数={ws.Cells(r,10).Value}）")
                    if not CHECK:
                        k.Value = price
        if not CHECK:
            wb.Save()
        wb.Close(False)

    # ---------- 3. 尿素模板删 639 ----------
    wb = app.Workbooks.Open(UREA, 0, bool(CHECK))
    if not CHECK and wb.ReadOnly:
        raise SystemExit("尿素模板被占用")
    ws = wb.Worksheets("加油数据")
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    hit = None
    for r in range(2, last + 1):
        if str(ws.Cells(r, 1).Value or "").strip().split(".")[0] == "261029000639":
            hit = r
            break
    print(f"尿素D15 加油数据: 639 在行{hit}")
    if hit and not CHECK:
        ws.Rows(hit).Delete()
        wb.Save()
        print("  已删除")
    wb.Close(False)

    # ---------- 4. 考核表 ----------
    wb = app.Workbooks.Open(KAOHE, 0, bool(CHECK))
    if not CHECK and wb.ReadOnly:
        raise SystemExit("考核表被占用")
    # 4a 删尿素数据里的 639
    ws = wb.Worksheets("每日尿素数据")
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    for r in range(2, last + 1):
        if str(ws.Cells(r, 1).Value or "").strip().split(".")[0] == "261029000639":
            print(f"考核表[每日尿素数据] 行{r}: {ws.Cells(r,4).Value} -> 删除")
            if not CHECK:
                ws.Rows(r).Delete()
            break
    # 4b 改油耗数据单价
    ws = wb.Worksheets("每日油耗数据")
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    for r in range(2, last + 1):
        sid = str(ws.Cells(r, 1).Value or "").strip().split(".")[0]
        for fsid, price in PRICE_FIX:
            if sid == fsid:
                print(f"考核表[每日油耗数据] 行{r} {sid}: K={ws.Cells(r,11).Value} -> {price}")
                if not CHECK:
                    ws.Cells(r, 11).Value = price
    if not CHECK:
        wb.Save()
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
