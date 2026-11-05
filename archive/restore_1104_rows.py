# -*- coding: utf-8 -*-
"""补回考核表「每日油耗数据」丢失的 6 条 11-04 结算行

背景：11-04/11-05 期间有过一次恢复操作，考核表里 11-04 只剩 BQN066 / BNT993 两条，
      另外 6 条（BFE485/BNX948/BPV550/LPW138/HPN783/B61HY1）丢了，但模板里数据完好。
做法：从各模板「计算模板」对应行读 A..V 整行 → 追加到考核表末尾
      → 之后跑 fleet_fuel_com 的宏同步补汇总
--check 只读
"""
import os, sys, shutil, glob
import pythoncom
import win32com.client as wc

HERE = os.path.dirname(os.path.abspath(__file__))
TPL = r"C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板"
KAOHE = r"C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm"
BAK = os.path.join(HERE, "backup_fix_1105")
CHECK = "--check" in sys.argv

# 编号 -> 模板文件名关键字
NEED = {
    "261104000136": "粤BFE485",
    "261104000135": "粤BNX948",
    "261104000111": "粤BPV550",
    "261104000125": "京LPW138",
    "261104000130": "京HPN783",
    "261104000128": "粤B61HY1",
}
NCOL = 24   # 必须写满 24 列：第 23 列是「日期」，漏了透视表会归到(空白)


def find_tpl(plate):
    for p in glob.glob(os.path.join(TPL, "*.xlsx")):
        n = os.path.basename(p)
        if plate in n and "尿素" not in n and not n.startswith("~$"):
            return p
    return None


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

        if not CHECK:
            os.makedirs(BAK, exist_ok=True)
            shutil.copy2(KAOHE, BAK)
            print("备份 ->", BAK)

        # 1. 从模板读
        picked = []
        for sid, plate in NEED.items():
            fp = find_tpl(plate)
            if not fp:
                print(f"  ! {plate} 无模板")
                continue
            wb = app.Workbooks.Open(fp, 0, True)
            ws = wb.Worksheets("计算模板") if _has(wb, "计算模板") else wb.Worksheets(1)
            last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
            for r in range(2, last + 1):
                if str(ws.Cells(r, 1).Value or "").strip().split(".")[0] == sid:
                    vals = []
                    for c in range(1, NCOL + 1):
                        cell = ws.Cells(r, c)
                        # 时间列存序列数
                        vals.append(cell.Value2 if c in (6, 7) else cell.Value)
                    g = ws.Cells(r, 7).Value2
                    picked.append((sid, plate, r, g, vals))
                    print(f"  读到 {plate} 行{r} 编号{sid} G={g} "
                          f"J={vals[9]} P={vals[15]} R={vals[17]}")
                    break
            wb.Close(False)

        if len(picked) != len(NEED):
            print(f"  ! 只读到 {len(picked)}/{len(NEED)} 条")

        # 2. 写入考核表
        wb = app.Workbooks.Open(KAOHE, 0, bool(CHECK))
        if not CHECK and wb.ReadOnly:
            raise SystemExit("考核表被占用")
        ws = wb.Worksheets("每日油耗数据")
        last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        exist = set()
        for r in range(2, last + 1):
            exist.add(str(ws.Cells(r, 1).Value or "").strip().split(".")[0])
        print(f"考核表现有 {last-1} 行")

        todo = [p for p in picked if p[0] not in exist]
        # 按 G（结束时间）排序，保持时间顺序
        todo.sort(key=lambda x: x[3] or 0)
        print(f"待补 {len(todo)} 条")
        for sid, plate, _, g, vals in todo:
            print(f"  + {plate} {sid} G={g}")
        if CHECK:
            print("\n(--check 未写入)")
            return

        start = last + 1
        for i, (sid, plate, _, g, vals) in enumerate(todo):
            row = start + i
            ws.Cells(row, 1).NumberFormat = "@"
            for c, v in enumerate(vals, start=1):
                if v is None:
                    continue
                if c in (6, 7):
                    ws.Cells(row, c).Value2 = v
                else:
                    ws.Cells(row, c).Value = v
            # 第 23 列「日期」必须单独补：漏写 -> 透视表归到(空白) -> 区域表显示当日未加油
            ws.Cells(row, 23).NumberFormat = 'yyyy-mm-dd'
            ws.Cells(row, 23).Value2 = int(g)
        wb.Save()
        print(f"已追加 {len(todo)} 行（{start} ~ {start+len(todo)-1}）")
        wb.Close(False)
        print("\n完成，下一步跑 fleet_fuel_com 做宏同步补汇总")
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


def _has(wb, name):
    try:
        wb.Worksheets(name)
        return True
    except Exception:
        return False


if __name__ == "__main__":
    main()
