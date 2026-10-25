# -*- coding: utf-8 -*-
"""修正 京AEP303 / 261024000494（10-23 15:53 -> 10-24 16:44）
GPS 里程在源系统为空 -> 模板 N 取到 0 -> P = -713752.1 负里程，汇总被判"达标"。
做法：模板「加油数据」补 GPS 里程 = 上一笔 GPS 713752.1 + 本段仪表里程 1450（兜底口径）
      -> 模板 N/P/R 由公式自动重算
      -> 考核表[每日油耗数据] 常量快照同步
      -> 汇总行同步重算达标率/超出油量/超出金额
--check 只看不写。
"""
import os, sys, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
TPL = r"C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板\B7-9.6米-京AEP303.xlsx"
KAOHE = r"C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm"
SUM = r"C:\Users\Jubin\Desktop\源数据\油耗分析汇总.xlsm"
BAK = os.path.join(HERE, "backup_fix_aep303_1025")

SID = "261024000494"
PLATE = "京AEP303"
PREV_GPS = 713752.1        # 上一笔真实 GPS（261023000473）
CUR_ODO = 708311.0         # 本笔仪表
PREV_ODO = 706861.0        # 上一笔仪表
GPS_EST = PREV_GPS + (CUR_ODO - PREV_ODO)   # 715202.1
CHECK = "--check" in sys.argv

print(f"推算 GPS = {PREV_GPS} + ({CUR_ODO} - {PREV_ODO}) = {GPS_EST}")
print("模式:", "--check 只读" if CHECK else "写入")

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
        for p in (TPL, KAOHE, SUM):
            shutil.copy2(p, BAK)
        print("备份 ->", BAK)

    # ---------- 1. 模板「加油数据」补 GPS ----------
    wb = app.Workbooks.Open(TPL, 0, bool(CHECK))
    if not CHECK and wb.ReadOnly:
        raise RuntimeError("模板被占用（只读）")
    ws = wb.Worksheets("加油数据")
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    hit = None
    for r in range(2, last + 1):
        if str(ws.Cells(r, 2).Value or "").strip() == PLATE and \
           abs(float(ws.Cells(r, 13).Value2 or 0) - 46319.6972453704) < 1e-6:
            hit = r
            break
    if hit is None:
        raise RuntimeError("模板加油数据未找到该条")
    print(f"模板[加油数据] 行{hit}: D(仪表)={ws.Cells(hit,4).Value} "
          f"E(GPS)={ws.Cells(hit,5).Value!r} M(时间)={ws.Cells(hit,13).Value}")
    if not CHECK:
        ws.Cells(hit, 5).Value = GPS_EST
    br = wb.Worksheets(1)
    print(f"模板 行271: N={br.Cells(271,14).Value} P={br.Cells(271,16).Value} "
          f"R={br.Cells(271,18).Value}")
    if not CHECK:
        br.Calculate()
        n_v = float(br.Cells(271, 14).Value or 0)
        p_v = float(br.Cells(271, 16).Value or 0)
        r_v = float(br.Cells(271, 18).Value or 0)
        s_v = float(br.Cells(271, 19).Value or 0)
        t_v = float(br.Cells(271, 20).Value or 0)
        u_v = float(br.Cells(271, 21).Value or 0)
        v_v = float(br.Cells(271, 22).Value or 0)
        print(f"  重算后: N={n_v} P={p_v} R={r_v:.6f} S={s_v:.6f} T={t_v:.6f} V={v_v:.6f}")
        if abs(p_v - (CUR_ODO - PREV_ODO)) > 0.01:
            raise RuntimeError(f"模板里程异常 P={p_v}")
        wb.Save()
        print("  模板已保存")
    wb.Close(False)

    # ---------- 2. 考核表「每日油耗数据」 ----------
    wb = app.Workbooks.Open(KAOHE, 0, bool(CHECK))
    if not CHECK and wb.ReadOnly:
        raise RuntimeError("考核表被占用（只读）")
    ws = wb.Worksheets("每日油耗数据")
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    kr = None
    for r in range(2, last + 1):
        if str(ws.Cells(r, 1).Value or "").strip().startswith(SID):
            kr = r
            break
    if kr is None:
        raise RuntimeError("考核表未找到该编号")
    print(f"考核表 行{kr}: N={ws.Cells(kr,14).Value} P={ws.Cells(kr,16).Value} R={ws.Cells(kr,18).Value}")
    if not CHECK:
        ws.Cells(kr, 14).Value = n_v          # GPS
        ws.Cells(kr, 15).Value = CUR_ODO      # 仪表
        ws.Cells(kr, 16).Value = p_v          # GPS 里程
        ws.Cells(kr, 17).Value = CUR_ODO - PREV_ODO   # 仪表里程
        ws.Cells(kr, 18).Value = r_v
        ws.Cells(kr, 19).Value = s_v
        ws.Cells(kr, 20).Value = t_v
        ws.Cells(kr, 21).Value = u_v
        ws.Cells(kr, 22).Value = v_v
        wb.Save()
        print("  考核表已保存")

    # ---------- 3. 汇总重算 ----------
    vol = float(ws.Cells(kr, 10).Value or 0)
    amt_ = float(ws.Cells(kr, 9).Value or 0)
    price = float(ws.Cells(kr, 11).Value or 0)
    wb.Close(False)

    wb = app.Workbooks.Open(SUM, 0, bool(CHECK))
    if not CHECK and wb.ReadOnly:
        raise RuntimeError("汇总被占用（只读）")
    ws = wb.Worksheets("Sheet1")
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    srow = None
    for r in range(2, last + 1):
        if str(ws.Cells(r, 3).Value or "").strip() == PLATE and \
           str(ws.Cells(r, 7).Value or "")[:10] == "2026-10-24":
            srow = r
            break
    if srow is None:
        raise RuntimeError("汇总未找到该行")
    km = float(p_v if not CHECK else (CUR_ODO - PREV_ODO))
    std = float(ws.Cells(srow, 13).Value or 26.0)
    std_vol = km * std / 100.0
    rate = 1 - (vol - std_vol) / std_vol
    excess = vol - std_vol
    ok = rate >= 1.0
    print(f"汇总 行{srow}: 升={vol} 里程={km} 油耗={vol/km*100:.4f} "
          f"标准={std} 达标率={rate:.6f} {'达标' if ok else '不达标'} 超出={excess:.2f}L/{excess*price:.2f}元")
    if not CHECK:
        ws.Cells(srow, 10).Value = km
        ws.Cells(srow, 11).Value = vol / km * 100
        ws.Cells(srow, 12).Value = "达标" if ok else "不达标"
        ws.Cells(srow, 15).Value = rate
        ws.Cells(srow, 16).Value = "未超出" if ok else price
        ws.Cells(srow, 17).Value = "未超出" if ok else excess
        ws.Cells(srow, 18).Value = "未超出" if ok else excess * price
        wb.Save()
        print("  汇总已保存")
    wb.Close(False)

    if CHECK:
        print("\n(--check 未写入)")
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
