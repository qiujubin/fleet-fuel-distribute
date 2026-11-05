# -*- coding: utf-8 -*-
"""粤BNT993 11-04 重复录入修正
261104000140（21:29:51）与 261104000141（21:33:02）
仪表 238008 / GPS 237335.8 / ¥2268.5 / 263.78L 全同 -> 同一笔录两次
导致周期 J=915.11（多算 263.78）、油耗 28.37 判不达标

做法：模板删 140 → rebuild 141 的 I/J（应为 5605.42 / 651.33）
      → 考核表 141 同步 → 汇总对应行改为 651.33 / 20.19 / 达标
"""
import os, sys, shutil
import pythoncom
import win32com.client as wc
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from com_formula import rebuild_ij

HERE = os.path.dirname(os.path.abspath(__file__))
TPL = r"C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板\B13-9.6米-粤BNT993.xlsx"
KAOHE = r"C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm"
SUM = r"C:\Users\Jubin\Desktop\源数据\油耗分析汇总.xlsm"
DUP = "261104000140"
KEEP = "261104000141"
CHECK = "--check" in sys.argv


def _has(wb, name):
    try:
        wb.Worksheets(name)
        return True
    except Exception:
        return False


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
        bak = os.path.join(HERE, "backup_fix_1105_bnt993")
        os.makedirs(bak, exist_ok=True)
        for p in (TPL, KAOHE, SUM):
            shutil.copy2(p, bak)
        print("备份 ->", bak)

    # 1. 模板
    wb = app.Workbooks.Open(TPL, 0, bool(CHECK))
    if not CHECK and wb.ReadOnly:
        raise SystemExit("模板被占用")
    ws = wb.Worksheets("计算模板") if _has(wb, "计算模板") else wb.Worksheets(1)
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    r_dup = r_keep = None
    for r in range(2, last + 1):
        sid = str(ws.Cells(r, 1).Value or "").strip().split(".")[0]
        if sid == DUP:
            r_dup = r
        elif sid == KEEP:
            r_keep = r
    print(f"模板：重复行{r_dup} 保留行{r_keep}")
    print(f"  删前 保留行: I={ws.Cells(r_keep,9).Value} J={ws.Cells(r_keep,10).Value} "
          f"P={ws.Cells(r_keep,16).Value} R={ws.Cells(r_keep,18).Value}")
    if not CHECK:
        ws.Rows(r_dup).Delete()
        if r_keep > r_dup:
            r_keep -= 1
        rebuild_ij(app, TPL, wb=wb, only_rows=[r_keep], verbose=False)
        ws.Calculate()
        i_v = float(ws.Cells(r_keep, 9).Value or 0)
        j_v = float(ws.Cells(r_keep, 10).Value or 0)
        p_v = float(ws.Cells(r_keep, 16).Value or 0)
        r_v = float(ws.Cells(r_keep, 18).Value or 0)
        print(f"  重算后 行{r_keep}: I={i_v} J={j_v} P={p_v} R={r_v:.4f}")
        if abs(j_v - 651.33) > 0.05:
            raise SystemExit(f"J 异常 {j_v}")
        wb.Save()
        print("  模板已保存")
    else:
        p_v = float(ws.Cells(r_keep, 16).Value or 0)
        i_v, j_v = 5605.42, 651.33
        r_v = j_v / p_v * 100
        print(f"  (预检预期) I={i_v} J={j_v} P={p_v} R={r_v:.4f}")
    wb.Close(False)

    # 2. 考核表
    wb = app.Workbooks.Open(KAOHE, 0, bool(CHECK))
    if not CHECK and wb.ReadOnly:
        raise SystemExit("考核表被占用")
    ws = wb.Worksheets("每日油耗数据")
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    # 删重复那条
    for r in range(2, last + 1):
        if str(ws.Cells(r, 1).Value or "").strip().split(".")[0] == DUP:
            print(f"  考核表 删行{r}（{DUP}）")
            if not CHECK:
                ws.Rows(r).Delete()
            break
    # 更新保留那条
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    for r in range(2, last + 1):
        if str(ws.Cells(r, 1).Value or "").strip().split(".")[0] == KEEP:
            print(f"  考核表 行{r}（{KEEP}）: I={ws.Cells(r,9).Value} J={ws.Cells(r,10).Value} "
                  f"R={ws.Cells(r,18).Value} -> I={i_v} J={j_v} R={r_v:.4f}")
            if not CHECK:
                ws.Cells(r, 9).Value = i_v
                ws.Cells(r, 10).Value = j_v
                ws.Cells(r, 18).Value = r_v
                ws.Cells(r, 19).Value = i_v / p_v
            break
    if not CHECK:
        wb.Save()
        print("  考核表已保存")
    wb.Close(False)

    # 3. 汇总
    wb = app.Workbooks.Open(SUM, 0, bool(CHECK))
    if not CHECK and wb.ReadOnly:
        raise SystemExit("汇总被占用")
    ws = wb.Worksheets("Sheet1")
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    for r in range(2, last + 1):
        if str(ws.Cells(r, 3).Value or "").strip() == "粤BNT993" and \
           str(ws.Cells(r, 7).Value or "")[:10] == "2026-11-04":
            std = float(ws.Cells(r, 13).Value or 22.0)
            std_vol = p_v * std / 100.0
            rate = 1 - (j_v - std_vol) / std_vol
            ok = rate >= 1.0
            print(f"  汇总 行{r}: {ws.Cells(r,9).Value} -> {j_v:.2f}L, "
                  f"油耗 {ws.Cells(r,11).Value:.2f} -> {r_v:.2f}, "
                  f"{'达标' if ok else '不达标'} 达标率={rate:.4f}")
            if not CHECK:
                ws.Cells(r, 8).Value = i_v
                ws.Cells(r, 9).Value = j_v
                ws.Cells(r, 11).Value = r_v
                ws.Cells(r, 12).Value = "达标" if ok else "不达标"
                ws.Cells(r, 15).Value = rate
                ws.Cells(r, 16).Value = "未超出" if ok else float(ws.Cells(r,16).Value or 0)
                ws.Cells(r, 17).Value = "未超出" if ok else (j_v - std_vol)
                ws.Cells(r, 18).Value = "未超出"
            break
    if not CHECK:
        wb.Save()
        print("  汇总已保存")
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
