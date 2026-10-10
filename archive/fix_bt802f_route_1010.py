# -*- coding: utf-8 -*-
"""10-10 修正：粤BT802F 单价笔误(0.26 应为 8.26) 导致误分尿素链路 -> 挪回油车

依据：
  261009000184 粤BT802F 10-09 17:52  ¥445.46 / 53.93L / 单价录成 0.26
  金额 ÷ 油量 = 445.46 / 53.93 = 8.2600  -> 真实单价 8.26（司机漏打"8"）
  该车里程完全延续：仪表 353878->354227、GPS 516127.2->516484.7
  -> 确认是粤BT802F 本车记录，只是单价错了，被"单价<5 归尿素"的规则误分

动作（严格按既有「误分流修正」流程）：
  1. 备份 导出文件 / 尿素文件 / 考核表
  2. 先把 backup_daily 的尿素文件快照拷到独立目录（防后续重跑覆盖）
  3. 改导出文件 单价 0.26 -> 8.26
  4. 用 backup_daily 快照回滚 尿素D14-4.2米-粤BT802F.xlsx（去掉误入行）
  5. 删考核表「每日尿素数据」里的误入行
  6. 校验汇总无该编号
之后需重跑主链路，让这条落到油车文件。

用法：python fix_bt802f_route_1010.py [--check]
"""
import os, sys, shutil, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable

TPL_DIR = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板'
KAOHE = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
EXPORT = r'C:\Users\Jubin\Desktop\数据\车辆加油_管理(2026-10-10).xlsx'
SUMMARY = r'C:\Users\Jubin\Desktop\源数据\油耗分析汇总.xlsm'

UID = '261009000184'
PLATE = '粤BT802F'
NEW_PRICE = 8.26

UREA_FILE = os.path.join(TPL_DIR, '尿素D14-4.2米-粤BT802F.xlsx')
BAK_ROOT = os.path.join(HERE, 'backup_fix_bt802f_route_1010')
BAK_UREA_SRC = os.path.join(HERE, 'backup_daily', '车辆油耗计算模板', '尿素D14-4.2米-粤BT802F.xlsx')
# 独立副本优先（backup_daily 会被每次主链路重跑覆盖）
BAK_UREA_LOCAL = os.path.join(HERE, 'backup_fix_bt802f_route_1010', '尿素D14_回滚源.xlsx')


def urea_src():
    return BAK_UREA_LOCAL if os.path.exists(BAK_UREA_LOCAL) else BAK_UREA_SRC

CHECK = '--check' in sys.argv


def disp(v):
    s = str(v or '').strip()
    return s[:-2] if s.endswith('.0') else s


def cells(vals, n):
    """COM 返回值统一成 list[list]"""
    if vals is None:
        return []
    if not isinstance(vals, tuple):
        return [[vals]]
    if vals and not isinstance(vals[0], tuple):
        return [list(vals)]
    return [list(r) for r in vals]


def open_app():
    import pythoncom, win32com.client as wc
    pythoncom.CoInitialize()
    for pid in ('Ket.Application', 'et.Application'):
        try:
            app = wc.DispatchEx(pid)
            app.Visible = False
            app.DisplayAlerts = False
            return app, pythoncom
        except Exception:
            continue
    raise RuntimeError('无法启动 WPS')


def scan(com):
    """只读扫描各定位点"""
    ok = True
    print('--- 1) 导出文件 定位 ---')
    wb = com.Workbooks.Open(EXPORT, 0, CHECK)
    ws = wb.Worksheets(1)
    ur = ws.UsedRange
    last = ur.Row + ur.Rows.Count - 1
    hit = None
    for r in range(2, last + 1):
        if disp(ws.Cells(r, 1).Value) == UID:
            hit = r
            print('  行%d 编号=%s 车牌=%s 金额=%s 油量=%s 单价=%s'
                  % (r, UID, ws.Cells(r, 2).Value, ws.Cells(r, 6).Value,
                     ws.Cells(r, 7).Value, ws.Cells(r, 8).Value))
            if not CHECK:
                ws.Cells(r, 8).Value = NEW_PRICE
            break
    if hit is None:
        print('  !! 未找到'); ok = False
    else:
        if not CHECK:
            wb.Save()
        print('  -> 单价 0.26 => %.2f' % NEW_PRICE)
    wb.Close(False)

    print('--- 2) 尿素文件 回滚源校验 ---')
    src = urea_src()
    if not os.path.exists(src):
        print('  !! 回滚源不存在:', src); ok = False
    else:
        from python_calamine import CalamineWorkbook
        print('  回滚源:', src)
        wb2 = CalamineWorkbook.from_path(src)
        d = wb2.get_sheet_by_name('加油数据').to_python(skip_empty_area=False)
        t = wb2.get_sheet_by_name('计算模板').to_python(skip_empty_area=False)
        f = [i for i, r in enumerate(t, 1) if i >= 2 and r[6] not in (None, '')]
        bad = [i for i, r in enumerate(d, 1) if r and UID in str(r[0])]
        print('  快照 加油数据=%d 行 / 模板有G行=%d / 含%s行=%s'
              % (len(d), len(f), UID, bad or '无'))
        if bad:
            print('  !! 快照里已含该编号，不能作为回滚源'); ok = False
        # 现网状态
        wb3 = CalamineWorkbook.from_path(UREA_FILE)
        d2 = wb3.get_sheet_by_name('加油数据').to_python(skip_empty_area=False)
        t2 = wb3.get_sheet_by_name('计算模板').to_python(skip_empty_area=False)
        f2 = [i for i, r in enumerate(t2, 1) if i >= 2 and r[6] not in (None, '')]
        print('  现网 加油数据=%d 行 / 模板有G行=%d （应比快照多 1）' % (len(d2), len(f2)))

    print('--- 3) 考核表 每日尿素数据 定位 ---')
    wb4 = com.Workbooks.Open(KAOHE, 0, CHECK)
    ws4 = wb4.Worksheets('每日尿素数据')
    ur4 = ws4.UsedRange
    last4 = ur4.Row + ur4.Rows.Count - 1
    found4 = None
    for r in range(last4, 1, -1):
        if disp(ws4.Cells(r, 1).Value) == UID:
            found4 = r
            print('  行%d 编号=%s 车牌=%s 金额=%s 升=%s 单价=%s'
                  % (r, UID, ws4.Cells(r, 4).Value, ws4.Cells(r, 9).Value,
                     ws4.Cells(r, 10).Value, ws4.Cells(r, 11).Value))
            break
    if found4 is None:
        print('  (未找到，视为已删除，跳过)')
    else:
        if not CHECK:
            ws4.Rows(found4).Delete()
            wb4.Save()
            print('  -> 已删除行 %d 并保存（WPS 原生）' % found4)
    wb4.Close(False)

    print('--- 4) 汇总 残留检查 ---')
    wb5 = com.Workbooks.Open(SUMMARY, 0, True)
    ws5 = wb5.Worksheets(1)
    ur5 = ws5.UsedRange
    last5 = ur5.Row + ur5.Rows.Count - 1
    vals = cells(ws5.Range(ws5.Cells(1, 1), ws5.Cells(last5, 14)).Value, 14)
    h = [i + 1 for i, row in enumerate(vals) if any(UID in str(c) for c in row)]
    print('  汇总共 %d 行，含 %s 的行: %s' % (last5, UID, h or '无（正常）'))
    wb5.Close(False)
    return ok


def main():
    if not os.path.exists(UREA_FILE):
        print('尿素文件不存在'); return 1
    os.makedirs(BAK_ROOT, exist_ok=True)
    if not CHECK:
        print('=== 备份 ===')
        for p in (EXPORT, UREA_FILE, KAOHE):
            shutil.copy2(p, BAK_ROOT)
            print('  ', os.path.basename(p))
        # 把 backup_daily 的尿素快照拷到独立目录（防重跑覆盖）
        _s = urea_src()
        if _s != BAK_UREA_LOCAL:
            shutil.copy2(_s, BAK_UREA_LOCAL)
        print('   尿素D14_回滚源.xlsx  <- ', _s)

    app, pythoncom = open_app()
    try:
        ok = scan(app)
        if not CHECK and ok:
            print('=== 回滚尿素文件 ===')
            app.Quit()          # 关闭 WPS 再替换文件，避免占用
            shutil.copy2(urea_src(), UREA_FILE)
            print('  已用快照覆盖 尿素D14-4.2米-粤BT802F.xlsx')
    finally:
        try:
            app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

    if CHECK:
        print('\n[--check] 未做任何写操作')
    else:
        print('\n完成。下一步：重跑主链路 python fleet_fuel_com.py "%s"' % EXPORT)
    return 0 if ok else 2


if __name__ == '__main__':
    sys.exit(main())
