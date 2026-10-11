# -*- coding: utf-8 -*-
"""10-11 本地镜像同步：把「车辆模板加油数据」里与源系统不一致的格对齐

比对结果（本地 vs 接口）：
  ① 粤BQN066 B12 261008000157 金额  本地 537.38   -> 源 2537.38   【本地跟随源】(10-08 只改了导出, 漏改模板)
  ② 粤BNJ052 B15 261007000147 当前里程 本地 26566 -> 源 265901    【本地跟随源】(用户已在源里改正)
  ③ 粤BQN066 B12 261006000108 当前里程 本地 264894 -> 源 26494    【源里仍差这一笔，本脚本不动】

做法：改「加油数据」相应格 -> app.Calculate() -> 读回模板所有「是」行的
      I/J/P/R 与改前对比 -> 若某「是」行数值变了，列出供人工核对考核表/汇总。

用法：python sync_local_1011.py [--check]
"""
import os, sys, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
TPL_DIR = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板'
KAOHE = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
BAK = os.path.join(HERE, 'backup_sync_local_1011')

JOBS = [
    ('B12-9.6米-粤BQN066.xlsx', '261008000157', 6, 2537.38),   # 第6列 = 加注金额
    ('B15-9.6米-粤BNJ052.xlsx', '261007000147', 4, 265901.0),  # 第4列 = 当前里程
]
CHECK = '--check' in sys.argv


def disp(v):
    s = str(v or '').strip()
    return s[:-2] if s.endswith('.0') else s


def read_shi(com, path):
    """返回 {编号: (I金额, J升数, P里程, R油耗)}，只取模板里 是否=是 的行"""
    wb = com.Workbooks.Open(path, 0, True)
    try:
        ws = wb.Worksheets('计算模板')
        ur = ws.UsedRange
        last = ur.Row + ur.Rows.Count - 1
        vals = ws.Range(ws.Cells(2, 1), ws.Cells(last, 18)).Value
        if not isinstance(vals, tuple):
            vals = ((vals,),)
        out = {}
        for row in vals:
            if not isinstance(row, tuple) or len(row) < 18:
                continue
            sid = disp(row[0])
            if not sid:
                continue
            if str(row[7]).strip() == '是':
                out[sid] = (row[8], row[9], row[15], row[17])
        return out
    finally:
        try: wb.Close(False)
        except Exception: pass


def main():
    os.makedirs(BAK, exist_ok=True)
    if not CHECK:
        for fn, *_ in JOBS:
            shutil.copy2(os.path.join(TPL_DIR, fn), BAK)
        shutil.copy2(KAOHE, BAK)
        print('已备份 2 个模板 + 考核表 ->', os.path.basename(BAK))

    import pythoncom, win32com.client as wc
    pythoncom.CoInitialize()
    app = None
    for pid in ('Ket.Application', 'et.Application'):
        try:
            app = wc.DispatchEx(pid); break
        except Exception:
            continue
    app.Visible = False
    app.DisplayAlerts = False
    ok = True
    try:
        # 基线
        before = {}
        for fn, *_ in JOBS:
            before[fn] = read_shi(app, os.path.join(TPL_DIR, fn))
            print('%s 是行 %d 个（基线）' % (fn, len(before[fn])))

        # 改动
        for fn, uid, col, newval in JOBS:
            path = os.path.join(TPL_DIR, fn)
            wb = app.Workbooks.Open(path, 0, CHECK)
            print('[%s] ReadOnly=%s' % (fn, wb.ReadOnly))
            ws = wb.Worksheets('加油数据')
            ur = ws.UsedRange
            last = ur.Row + ur.Rows.Count - 1
            hit = None
            for r in range(2, last + 1):
                if disp(ws.Cells(r, 1).Value) == uid:
                    hit = r
                    print('   行%d %s 第%d列 现值=%s -> %s'
                          % (r, uid, col, ws.Cells(r, col).Value, newval))
                    if not CHECK:
                        ws.Cells(r, col).Value = newval
                    break
            if hit is None:
                print('   !! 未找到'); ok = False
            if not CHECK and hit is not None:
                wb.Save()
                app.Calculate()
            wb.Close(False)

        # 改后对比
        for fn, uid, col, newval in JOBS:
            after = read_shi(app, os.path.join(TPL_DIR, fn))
            b, a = before[fn], after
            changed = [k for k in set(b) | set(a) if b.get(k) != a.get(k)]
            print('[%s] 改后：是行 %d 个，数值发生变化 %d 个 %s'
                  % (fn, len(a), len(changed), changed or ''))
            for k in changed:
                print('     %s  金额/升数/里程/油耗：%s -> %s' % (k, b.get(k), a.get(k)))

        # 只读复核「加油数据」格是否落盘
        for fn, uid, col, newval in JOBS:
            from python_calamine import CalamineWorkbook
            wb = CalamineWorkbook.from_path(os.path.join(TPL_DIR, fn))
            rows = wb.get_sheet_by_name('加油数据').to_python(skip_empty_area=False)
            for r in rows:
                if r and disp(r[0]) == uid:
                    got = r[col - 1]
                    print('复核 %s %s: %s = %s (期望 %s)' % (fn, uid, ['', '编号', '车牌', '日期', '当前里程', 'GPS里程', '金额', '油量', '单价'][col], got, newval))
                    if not CHECK and abs(float(got) - float(newval)) > 1e-6:
                        ok = False; print('   !! 未落盘')
    finally:
        try: app.Quit()
        except Exception: pass
        pythoncom.CoUninitialize()

    if CHECK:
        print('\n[--check] 未做任何写操作')
    return 0 if ok else 2


if __name__ == '__main__':
    sys.exit(main())
