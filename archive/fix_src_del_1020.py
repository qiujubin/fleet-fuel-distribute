# -*- coding: utf-8 -*-
"""2026-10-20 源系统更正后的本地同步（用户删了 405 / 272 / 261006000108）
1. D21-粤B17K86:   删 261019000405（模板行250+加油行252），403 恢复锚点 F=上一是行G
2. 尿素B16-粤BNV030: 删 261014000272（模板行102+加油行105），
   271 恢复锚点 F=上一是行G，261015000298 的 F 改为 271 的 G
3. B12-粤BQN066:   删 261006000108（模板行238+加油行240）
全部改完 Calculate+Save，然后 calamine 只读回读打出是行前后对比。
"""
import os, sys
import pythoncom, win32com.client as wc

HERE = os.path.dirname(os.path.abspath(__file__))
TPL = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板'
D21 = os.path.join(TPL, 'D21-4.2米-粤B17K86.xlsx')
UB16 = os.path.join(TPL, '尿素B16-9.6米-粤BNV030.xlsx')
B12 = os.path.join(TPL, 'B12-9.6米-粤BQN066.xlsx')


def find_row(ws, sid, last):
    for r in range(2, last + 1):
        v = str(ws.Cells(r, 1).Value or '').strip()
        if v.endswith('.0'):
            v = v[:-2]
        if v == sid:
            return r
    return None


def renum(ws, last):
    n = 0
    for r in range(2, last + 1):
        if ws.Cells(r, 1).Value not in (None, ''):
            n += 1
            ws.Cells(r, 2).Value = n


pythoncom.CoInitialize()
app = None
for pid in ('Ket.Application', 'et.Application'):
    try:
        app = wc.DispatchEx(pid)
        app.Visible = False
        app.DisplayAlerts = False
        break
    except Exception:
        continue

try:
    # ---------- D21 ----------
    wb = app.Workbooks.Open(D21, 0, False)
    if wb.ReadOnly:
        raise SystemExit('D21 只读')
    wst = wb.Worksheets('计算模板')
    wsd = wb.Worksheets('加油数据')
    lt = wst.UsedRange.Row + wst.UsedRange.Rows.Count - 1
    ld = wsd.UsedRange.Row + wsd.UsedRange.Rows.Count - 1
    t405 = find_row(wst, '261019000405', lt)
    d405 = find_row(wsd, '261019000405', ld)
    t403 = find_row(wst, '261019000403', lt)
    print('D21: 模板405行=%d 加油405行=%d 模板403行=%d' % (t405, d405, t403))
    assert t405 and d405 and t403
    anchorF = wst.Cells(t403 - 1, 7).Value   # 上一是行(248) 的 G
    wst.Rows(t405).Delete()
    wsd.Rows(d405).Delete()
    # 删行后 403 行号不变(在其上方删除才需要调整; 250>249 无影响)
    wst.Cells(t403, 6).Value = anchorF       # F = 上一是行 G (10-18 21:57)
    print('  D21 已删 405 两行, 403 F <- %s' % anchorF)
    renum(wst, wst.UsedRange.Row + wst.UsedRange.Rows.Count - 1)
    wb.Save()
    wb.Close(False)

    # ---------- 尿素B16 ----------
    wb = app.Workbooks.Open(UB16, 0, False)
    if wb.ReadOnly:
        raise SystemExit('尿素B16 只读')
    wst = wb.Worksheets('计算模板')
    wsd = wb.Worksheets('加油数据')
    lt = wst.UsedRange.Row + wst.UsedRange.Rows.Count - 1
    ld = wsd.UsedRange.Row + wsd.UsedRange.Rows.Count - 1
    t272 = find_row(wst, '261014000272', lt)
    d272 = find_row(wsd, '261014000272', ld)
    t271 = find_row(wst, '261014000271', lt)
    print('尿素B16: 模板272行=%d 加油272行=%d 模板271行=%d' % (t272, d272, t271))
    assert t272 and d272 and t271
    # 找 271 之前的最新是行（G < 271 的 G）
    g271 = wst.Cells(t271, 7).Value
    prev_shi_row, prev_g = None, None
    for r in range(2, lt + 1):
        if r == t271:
            continue
        if str(wst.Cells(r, 8).Value or '').strip() == '是':
            g = wst.Cells(r, 7).Value
            if g is not None and g < g271 and (prev_g is None or g > prev_g):
                prev_shi_row, prev_g = r, g
    print('  271 的上一是行=%d G=%s' % (prev_shi_row, prev_g))
    wst.Rows(t272).Delete()
    wsd.Rows(d272).Delete()
    lt = wst.UsedRange.Row + wst.UsedRange.Rows.Count - 1
    t271 = find_row(wst, '261014000271', lt)          # 行号不变(272 在下方)
    t298 = find_row(wst, '261015000298', lt)          # 103 -> 102
    wst.Cells(t271, 6).Value = prev_g                 # 271 F <- 上一是行 G
    wst.Cells(t298, 6).Value = wst.Cells(t271, 7).Value  # 298 F <- 271 G
    print('  尿素B16 已删 272 两行, 271 F <- 上一是行G, 298 F <- 271 G(行%d)' % t298)
    renum(wst, lt)
    wb.Save()
    wb.Close(False)

    # ---------- B12 ----------
    wb = app.Workbooks.Open(B12, 0, False)
    if wb.ReadOnly:
        raise SystemExit('B12 只读')
    wst = wb.Worksheets('计算模板')
    wsd = wb.Worksheets('加油数据')
    lt = wst.UsedRange.Row + wst.UsedRange.Rows.Count - 1
    ld = wsd.UsedRange.Row + wsd.UsedRange.Rows.Count - 1
    t108 = find_row(wst, '261006000108', lt)
    d108 = find_row(wsd, '261006000108', ld)
    print('B12: 模板108行=%d 加油108行=%d' % (t108, d108))
    assert t108 and d108
    wst.Rows(t108).Delete()
    wsd.Rows(d108).Delete()
    print('  B12 已删 261006000108 两行')
    renum(wst, wst.UsedRange.Row + wst.UsedRange.Rows.Count - 1)
    wb.Save()
    wb.Close(False)

    app.Calculate()
finally:
    try:
        app.Quit()
    except Exception:
        pass
    pythoncom.CoUninitialize()

print('--- 模板编辑完成, calamine 回读对比 ---')
from python_calamine import CalamineWorkbook
for fp, tag in ((D21, 'D21'), (UB16, '尿素B16'), (B12, 'B12')):
    wb = CalamineWorkbook.from_path(fp)
    t = wb.get_sheet_by_name('计算模板').to_python(skip_empty_area=False)
    d = wb.get_sheet_by_name('加油数据').to_python(skip_empty_area=False)
    print('=' * 16, tag, '加油数据=%d 行' % len(d))
    for i, r in enumerate(t, 1):
        if i >= 2 and r[6] not in (None, '') and str(r[7]).strip() == '是':
            print('   是行%d: %s I=%s J=%s P=%s R=%s' % (i, str(r[0])[:14], r[8], r[9],
                  str(r[15])[:9], str(r[17])[:9]))
