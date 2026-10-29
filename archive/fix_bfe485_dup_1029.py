# -*- coding: utf-8 -*-
"""粤BFE485 重复录入修正（2026-10-29）

现象：10-28 20:27:57（261028000596，庞真）与 20:33:15（261028000597，周渝）
      金额/油量/里程/单价/加油站完全相同 -> 同一笔被两人各录一次
      597 是最后一条"是"，导致 J=558.17（多算 189.37）、油耗 32.58 判不达标

做法：
  模板 B11：删掉 597 那一行，596 升为有效"是"（F 补 10-27 09:51:03），rebuild I/J
  考核表：行3665 由 597 改为 596 的快照值（I=3083.13 J=368.80 P=1713.2 R=21.5269）
  汇总：对应行 368.80 / 1713.2 / 21.53 / 达标
--check 只读
"""
import os, sys, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

TPL = r"C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板\B11-9.6米-粤BFE485.xlsx"
KAOHE = r"C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm"
SUM = r"C:\Users\Jubin\Desktop\源数据\油耗分析汇总.xlsm"
BAK = os.path.join(HERE, "backup_fix_bfe485_1029")
CHECK = "--check" in sys.argv

SID_KEEP = "261028000596"
SID_DUP = "261028000597"

import pythoncom
import win32com.client as wc
from com_formula import rebuild_ij


def main():
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
            for p in (TPL, KAOHE, SUM):
                shutil.copy2(p, BAK)
            print("备份 ->", BAK)

        # ---------- 1. 模板 ----------
        wb = app.Workbooks.Open(TPL, 0, bool(CHECK))
        if not CHECK and wb.ReadOnly:
            raise RuntimeError("模板被占用")
        ws = wb.Worksheets(1)
        last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        row_keep = row_dup = None
        for r in range(2, last + 1):
            sid = str(ws.Cells(r, 1).Value or "").strip().split(".")[0]
            if sid == SID_KEEP:
                row_keep = r
            elif sid == SID_DUP:
                row_dup = r
        print(f"模板：{SID_KEEP} 行{row_keep}，{SID_DUP} 行{row_dup}")
        if not row_keep or not row_dup:
            raise RuntimeError("模板里没找到这两条")

        f_serial = float(ws.Cells(row_dup, 6).Value2 or 0)   # 597 的 F = 10-27 09:51:03
        print(f"  周期起点 F = {f_serial}（10-27 09:51:03）")
        print(f"  删除前 行{row_dup}: I={ws.Cells(row_dup,9).Value} J={ws.Cells(row_dup,10).Value} "
              f"P={ws.Cells(row_dup,16).Value} R={ws.Cells(row_dup,18).Value}")

        if not CHECK:
            ws.Rows(row_dup).Delete()
            # 596 升为有效"是"
            ws.Cells(row_keep, 6).Value = f_serial
            ws.Cells(row_keep, 7).Value2 = ws.Cells(row_keep, 7).Value2
            ws.Cells(row_keep, 8).Value = "是"
            rebuild_ij(app, TPL, wb=wb, only_rows=[row_keep], verbose=False)
            ws.Calculate()
            i_val = float(ws.Cells(row_keep, 9).Value or 0)
            j_val = float(ws.Cells(row_keep, 10).Value or 0)
            p_val = float(ws.Cells(row_keep, 16).Value or 0)
            r_val = float(ws.Cells(row_keep, 18).Value or 0)
            print(f"  重算后 行{row_keep}: I={i_val} J={j_val} P={p_val} R={r_val:.6f}")
            if abs(j_val - 368.80) > 0.05:
                raise RuntimeError(f"J 异常 {j_val}，期望 368.80")
            wb.Save()
            print("  模板已保存")
        wb.Close(False)

        # ---------- 2. 考核表 ----------
        wb = app.Workbooks.Open(KAOHE, 0, bool(CHECK))
        if not CHECK and wb.ReadOnly:
            raise RuntimeError("考核表被占用")
        ws = wb.Worksheets("每日油耗数据")
        last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        kr = None
        for r in range(2, last + 1):
            if str(ws.Cells(r, 1).Value or "").strip().split(".")[0] == SID_DUP:
                kr = r
                break
        print(f"考核表：{SID_DUP} 行{kr}")
        if kr is None:
            raise RuntimeError("考核表未找到 597")
        if not CHECK:
            km = 1713.2
            vol = 368.80
            amt = 3083.13
            ws.Cells(kr, 1).NumberFormat = "@"
            ws.Cells(kr, 1).Value = SID_KEEP
            ws.Cells(kr, 6).Value = f_serial
            ws.Cells(kr, 7).Value = ws.Cells(kr, 7).Value      # 结束时间不变
            ws.Cells(kr, 8).Value = "是"
            ws.Cells(kr, 9).Value = amt
            ws.Cells(kr, 10).Value = vol
            ws.Cells(kr, 16).Value = km
            ws.Cells(kr, 17).Value = 1576.0
            ws.Cells(kr, 18).Value = vol / km * 100
            ws.Cells(kr, 19).Value = amt / km
            ws.Cells(kr, 20).Value = vol / km * 100
            ws.Cells(kr, 21).Value = amt / km
            ws.Cells(kr, 22).Value = vol / km * 100
            wb.Save()
            print(f"  考核表已更新：{SID_KEEP} 升={vol} 里程={km} 油耗={vol/km*100:.4f}")
        wb.Close(False)

        # ---------- 3. 汇总 ----------
        wb = app.Workbooks.Open(SUM, 0, bool(CHECK))
        if not CHECK and wb.ReadOnly:
            raise RuntimeError("汇总被占用")
        ws = wb.Worksheets("Sheet1")
        last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        srow = None
        for r in range(2, last + 1):
            if str(ws.Cells(r, 3).Value or "").strip() == "粤BFE485" and \
               str(ws.Cells(r, 7).Value or "")[:10] == "2026-10-28":
                srow = r
                break
        print(f"汇总 行{srow}: 现值 升={ws.Cells(srow,9).Value} 里程={ws.Cells(srow,10).Value} "
              f"油耗={ws.Cells(srow,11).Value} {ws.Cells(srow,12).Value}")
        if srow is None:
            raise RuntimeError("汇总未找到 BFE485 10-28")
        km, vol = 1713.2, 368.80
        std = float(ws.Cells(srow, 13).Value or 22.0)
        std_vol = km * std / 100.0
        rate = 1 - (vol - std_vol) / std_vol
        ok = rate >= 1.0
        print(f"  -> 升={vol} 里程={km} 油耗={vol/km*100:.4f} 达标率={rate:.6f} "
              f"{'达标' if ok else '不达标'}")
        if not CHECK:
            ws.Cells(srow, 9).Value = vol
            ws.Cells(srow, 10).Value = km
            ws.Cells(srow, 11).Value = vol / km * 100
            ws.Cells(srow, 12).Value = "达标" if ok else "不达标"
            ws.Cells(srow, 15).Value = rate
            ws.Cells(srow, 16).Value = "未超出" if ok else 8.36
            ws.Cells(srow, 17).Value = "未超出" if ok else (vol - std_vol)
            ws.Cells(srow, 18).Value = "未超出" if ok else (vol - std_vol) * 8.36
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


if __name__ == "__main__":
    main()
