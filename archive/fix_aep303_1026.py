# -*- coding: utf-8 -*-
"""京AEP303 GPS 缺口处理（10-26 运行）

状况：
  10-24 条（261024000494）用户已在源系统补真实 GPS 715246.7 -> 模板/考核表已自动更新
        （1494.6 km / 26.1274），但**昨天写进汇总的 1450/26.93 不会自动回更**
  10-25 条（261025000510）GPS 仍为空 -> N=0 -> P=-715246.7 负里程

做法：
  A. 10-25 模板「加油数据」E 补 = 上一笔 GPS 715246.7 + 本段仪表 1469 = 716715.7（兜底口径）
     -> 模板行272 / 考核表行3641 / 汇总对应行 全部按此重算
  B. 10-24 汇总行按真实 GPS 同步为 1494.6 km / 26.1274
--check 只读
"""
import os, sys, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
TPL = r"C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板\B7-9.6米-京AEP303.xlsx"
KAOHE = r"C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm"
SUM = r"C:\Users\Jubin\Desktop\源数据\油耗分析汇总.xlsm"
BAK = os.path.join(HERE, "backup_fix_aep303_1026")
CHECK = "--check" in sys.argv

PREV_GPS = 715246.7      # 10-24 真实 GPS
CUR_ODO = 709780.0       # 10-25 仪表
PREV_ODO = 708311.0      # 10-24 仪表
GPS_EST = PREV_GPS + (CUR_ODO - PREV_ODO)
T25_SERIAL = 46320.693587963   # 2026-10-25 16:38:13

import pythoncom
import win32com.client as wc


def sync_summary(ws, plate, day, km, vol, price):
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    row = None
    for r in range(2, last + 1):
        if str(ws.Cells(r, 3).Value or "").strip() == plate and str(ws.Cells(r, 7).Value or "")[:10] == day:
            row = r
            break
    if row is None:
        print(f"  ! 汇总未找到 {plate} / {day}")
        return None
    std = float(ws.Cells(row, 13).Value or 26.0)
    std_vol = km * std / 100.0
    rate = 1 - (vol - std_vol) / std_vol
    excess = vol - std_vol
    ok = rate >= 1.0
    print(f"  汇总 行{row}: 里程={km:.1f} 油耗={vol/km*100:.4f} 标准={std} "
          f"达标率={rate:.6f} {'达标' if ok else '不达标'} 超出={excess:.2f}L/{excess*price:.2f}元")
    if not CHECK:
        ws.Cells(row, 10).Value = km
        ws.Cells(row, 11).Value = vol / km * 100
        ws.Cells(row, 12).Value = "达标" if ok else "不达标"
        ws.Cells(row, 15).Value = rate
        ws.Cells(row, 16).Value = "未超出" if ok else price
        ws.Cells(row, 17).Value = "未超出" if ok else excess
        ws.Cells(row, 18).Value = "未超出" if ok else excess * price
    return row


def main():
    print(f"10-25 推算 GPS = {PREV_GPS} + ({CUR_ODO} - {PREV_ODO}) = {GPS_EST}")
    print("模式:", "--check 只读" if CHECK else "写入")

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

        # ---- 1. 模板：补 10-25 的 GPS ----
        wb = app.Workbooks.Open(TPL, 0, bool(CHECK))
        if not CHECK and wb.ReadOnly:
            raise RuntimeError("模板被占用")
        ws = wb.Worksheets("加油数据")
        last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        hit = None
        for r in range(2, last + 1):
            if str(ws.Cells(r, 2).Value or "").strip() == "京AEP303" and \
               abs(float(ws.Cells(r, 13).Value2 or 0) - T25_SERIAL) < 1e-3:
                hit = r
                break
        if hit is None:
            raise RuntimeError("模板加油数据未找到 10-25 那条")
        print(f"模板[加油数据] 行{hit}: D={ws.Cells(hit,4).Value} E={ws.Cells(hit,5).Value!r}")
        if not CHECK:
            ws.Cells(hit, 5).Value = GPS_EST
        br = wb.Worksheets(1)
        br.Calculate()
        n25 = float(br.Cells(272, 14).Value or 0)
        p25 = float(br.Cells(272, 16).Value or 0)
        r25 = float(br.Cells(272, 18).Value or 0)
        s25 = float(br.Cells(272, 19).Value or 0)
        t25 = float(br.Cells(272, 20).Value or 0)
        u25 = float(br.Cells(272, 21).Value or 0)
        v25 = float(br.Cells(272, 22).Value or 0)
        print(f"模板 行272: N={n25} P={p25} R={r25:.6f}")
        # 行271（10-24）当前真实值
        p24 = float(br.Cells(271, 16).Value or 0)
        print(f"模板 行271(10-24 真实 GPS): P={p24:.1f} R={float(br.Cells(271,18).Value or 0):.6f}")
        if not CHECK:
            if abs(p25 - (CUR_ODO - PREV_ODO)) > 0.01:
                raise RuntimeError(f"10-25 里程异常 {p25}")
            wb.Save()
            print("  模板已保存")
        wb.Close(False)

        # ---- 2. 考核表 ----
        wb = app.Workbooks.Open(KAOHE, 0, bool(CHECK))
        if not CHECK and wb.ReadOnly:
            raise RuntimeError("考核表被占用")
        ws = wb.Worksheets("每日油耗数据")
        last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        kr = None
        for r in range(2, last + 1):
            if str(ws.Cells(r, 1).Value or "").strip().startswith("261025000510"):
                kr = r
                break
        if kr is None:
            raise RuntimeError("考核表未找到 261025000510")
        vol = float(ws.Cells(kr, 10).Value or 0)
        price = float(ws.Cells(kr, 11).Value or 0)
        print(f"考核表 行{kr}: 升={vol} 单价={price} P={ws.Cells(kr,16).Value}")
        if not CHECK:
            ws.Cells(kr, 14).Value = n25
            ws.Cells(kr, 15).Value = CUR_ODO
            ws.Cells(kr, 16).Value = p25
            ws.Cells(kr, 17).Value = CUR_ODO - PREV_ODO
            ws.Cells(kr, 18).Value = r25
            ws.Cells(kr, 19).Value = s25
            ws.Cells(kr, 20).Value = t25
            ws.Cells(kr, 21).Value = u25
            ws.Cells(kr, 22).Value = v25
            wb.Save()
            print("  考核表已保存")
        wb.Close(False)

        # ---- 3. 汇总：两行一起对齐 ----
        wb = app.Workbooks.Open(SUM, 0, bool(CHECK))
        if not CHECK and wb.ReadOnly:
            raise RuntimeError("汇总被占用")
        ws = wb.Worksheets("Sheet1")
        print("汇总:")
        sync_summary(ws, "京AEP303", "2026-10-24", p24, 390.5, 8.21)
        sync_summary(ws, "京AEP303", "2026-10-25", p25 if not CHECK else (CUR_ODO - PREV_ODO), vol, price)
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


if __name__ == "__main__":
    main()
