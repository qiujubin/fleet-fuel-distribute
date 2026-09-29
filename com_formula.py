# -*- coding: utf-8 -*-
"""车辆模板「计算模板」I/J 周期累计公式 —— 纯 COM 重建

⚠️ 2026-09-28 用户明令：**任何写入只准 win32/COM，禁止 openpyxl 保存文件**。
本模块取代 `rebuild_formulas.rebuild()`（那版用 openpyxl 保存，会重新序列化
styles.xml 与工作表 XML，属违规路径，已停用）。

I/J 公式形态：
    I = =_xlfn.XLOOKUP(D&G, 加油数据!B:B&加油数据!M:M, 加油数据!F:F, 0) + I{r-1} + ... + I{anchor+1}
    J = 同上，取 加油数据!G:G
    anchor = G 列中值等于本行 F 的那一行（= 上一个有效"是"行）；
             找不到 anchor（首周期）时从 first_filled-1 起算
    非有效"是"的行 → 只有 XLOOKUP 部分，无累加

读取用 python-calamine（只读，不碰文件），写入用 Ket.Application（WPS 原生保存）。
"""
import os, datetime

from python_calamine import CalamineWorkbook

X = '_xlfn.XLOOKUP'
EPOCH = datetime.datetime(1899, 12, 30)


def to_serial(v):
    if v is None: return None
    if isinstance(v, datetime.datetime):
        return (v - EPOCH).total_seconds() / 86400.0
    if isinstance(v, datetime.date):
        return (datetime.datetime(v.year, v.month, v.day) - EPOCH).total_seconds() / 86400.0
    return v


def tkey(v):
    """任意时间表示 -> 整数秒匹配键（抗 datetime / 序列数 / 字符串漂移）"""
    if v is None: return None
    if isinstance(v, bool): return None
    if isinstance(v, (int, float)):
        return round(float(v) * 86400)
    if isinstance(v, (datetime.datetime, datetime.date)):
        return round(to_serial(v) * 86400)
    s = str(v).strip()
    for f in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d'):
        try:
            return round(to_serial(datetime.datetime.strptime(s, f)) * 86400)
        except ValueError:
            pass
    try:
        return round(float(s) * 86400)
    except ValueError:
        return None


def _sheet(wb, names):
    for n in names:
        if n in wb.sheet_names:
            return wb.get_sheet_by_name(n)
    return None


def read_plan(fp):
    """返回 [(row, off_rows)] —— 每行应写入的累加行号列表（降序）"""
    wb = CalamineWorkbook.from_path(fp)
    tpl = _sheet(wb, ('计算模板', '油耗计算'))
    dat = wb.get_sheet_by_name('加油数据')
    if tpl is None or dat is None:
        raise RuntimeError('缺少 计算模板/油耗计算 或 加油数据 工作表: ' + fp)
    drows = dat.to_python(skip_empty_area=False)
    trows = tpl.to_python(skip_empty_area=False)

    full_map = {}
    for row in drows[1:]:
        if len(row) > 12 and row[0] not in (None, '') and row[12] not in (None, ''):
            k = tkey(row[12])
            if k is not None:
                full_map[k] = str(row[8] or '').strip()

    g_rows, filled, Fs = {}, [], {}
    for i, row in enumerate(trows, start=1):
        if i < 2: continue
        f = row[5] if len(row) > 5 else None
        g = row[6] if len(row) > 6 else None
        Fs[i] = f
        if g not in (None, ''):
            k = tkey(g)
            if k is not None:
                g_rows[k] = i
                filled.append(i)
    first_filled = filled[0] if filled else 2

    plan = []
    for i in filled:
        f = Fs.get(i)
        g = trows[i - 1][6] if len(trows[i - 1]) > 6 else None
        is_shi = full_map.get(tkey(g)) == '是'
        off = []
        if is_shi and f not in (None, ''):
            a = g_rows.get(tkey(f))
            off = list(range(i - 1, a, -1)) if a is not None else list(range(i - 1, first_filled - 1, -1))
        plan.append((i, off))
    return plan


def _true_last_row(ws, col=1):
    ur = ws.UsedRange
    end = min(ur.Row + ur.Rows.Count - 1, ws.Rows.Count)
    while end >= 1:
        start = max(1, end - 1999)
        vals = ws.Range(ws.Cells(start, col), ws.Cells(end, col)).Value2
        if vals is None:
            end = start - 1; continue
        if not isinstance(vals, tuple): vals = ((vals,),)
        for i in range(len(vals) - 1, -1, -1):
            v = vals[i][0] if isinstance(vals[i], tuple) else vals[i]
            if v not in (None, ''):
                return start + i
        end = start - 1
    return 0


def read_plan_com(wb):
    """从**已打开**的工作簿（内存态）读 F/G 与加油数据，算出每行 off_rows。

    用于"刚在内存里改过 F、还没落盘"的场景 —— 此时读磁盘会拿到旧状态。
    2026-09-29 实战踩坑：retro_fix 改完 F 后直接调 rebuild_ij，磁盘上
    行154 的 F 还在、行155 的 F 还是空 → 行155 被判为非有效是 → 累加项算成 0，
    金额少了一笔（581.75 而非 581.75+640.53）。
    """
    ws = None
    for sn in ('计算模板', '油耗计算'):
        try:
            ws = wb.Worksheets(sn); break
        except Exception:
            continue
    if ws is None:
        raise RuntimeError('无 计算模板/油耗计算 工作表')
    wsd = wb.Worksheets('加油数据')

    full_map = {}
    last_d = _true_last_row(wsd, 1)
    if last_d >= 2:
        dv = wsd.Range(wsd.Cells(2, 1), wsd.Cells(last_d, 13)).Value2
        if dv is not None:
            if not isinstance(dv, tuple): dv = ((dv,),)
            for row in dv:
                if not isinstance(row, tuple): continue
                a = row[0]
                m = row[12] if len(row) > 12 else None
                flag9 = row[8] if len(row) > 8 else None
                if a not in (None, '') and m not in (None, ''):
                    k = tkey(m)
                    if k is not None:
                        full_map[k] = str(flag9 or '').strip()

    Fs, Gs = {}, {}
    last_g = _true_last_row(ws, 7)
    if last_g >= 2:
        tv = ws.Range(ws.Cells(2, 6), ws.Cells(last_g, 7)).Value2
        if tv is not None:
            if not isinstance(tv, tuple): tv = ((tv,),)
            for i, row in enumerate(tv, start=2):
                if isinstance(row, tuple):
                    Fs[i] = row[0]
                    Gs[i] = row[1] if len(row) > 1 else None
                else:
                    Fs[i] = row

    g_rows, filled = {}, []
    for i, g in Gs.items():
        if g not in (None, ''):
            k = tkey(g)
            if k is not None:
                g_rows[k] = i
                filled.append(i)
    filled.sort()
    first_filled = filled[0] if filled else 2

    plan = []
    for i in filled:
        f, g = Fs.get(i), Gs.get(i)
        is_shi = full_map.get(tkey(g)) == '是'
        off = []
        if is_shi and f not in (None, ''):
            a = g_rows.get(tkey(f))
            off = list(range(i - 1, a, -1)) if a is not None else list(range(i - 1, first_filled - 1, -1))
        plan.append((i, off))
    return plan


def build_ij(r, off_rows):
    dg = f'D{r}&G{r}'
    bi = f'={X}({dg},加油数据!B:B&加油数据!M:M,加油数据!F:F,0)'
    bj = f'={X}({dg},加油数据!B:B&加油数据!M:M,加油数据!G:G,0)'
    oi = ''.join(f'+I{x}' for x in off_rows)
    oj = ''.join(f'+J{x}' for x in off_rows)
    return bi + oi, bj + oj


def _write_cell(cell, text, off_rows, col_letter):
    try:
        cell.FormulaArray = text
        return 'array'
    except Exception:
        pass
    try:
        cell.Formula = text
        return 'formula'
    except Exception:
        pass
    if off_rows:      # 超长公式兜底: 用 SUM 等价替换
        head = text.split('+', 1)[0]
        fs = f'{head}+SUM({col_letter}{off_rows[-1]}:{col_letter}{off_rows[0]})'
        cell.Formula = fs
        return 'sum'
    return 'fail'


def rebuild_ij(app, fp, wb=None, only_rows=None, verbose=True):
    """用 COM 重写「计算模板」I/J 列。返回 (写入行数, 明细)

    wb: 传入已打开的工作簿则复用之（不 Save/Close），并**从工作簿内存态读 F/G**
        —— 若刚在内存里改过 F 还没落盘，读磁盘会算错累加范围。
    only_rows: 只重写这些行（None = 全部有 G 值的行）。
    """
    plan = read_plan_com(wb) if wb is not None else read_plan(fp)
    if only_rows is not None:
        keep = set(only_rows)
        plan = [(r, o) for r, o in plan if r in keep]
    own = wb is None
    if own:
        wb = app.Workbooks.Open(fp, 0, False)
        if wb.ReadOnly:
            raise RuntimeError('文件被占用(只读): ' + fp)
    ws = None
    for sn in ('计算模板', '油耗计算'):
        try:
            ws = wb.Worksheets(sn); break
        except Exception:
            continue
    if ws is None:
        if own:
            wb.Close(False)
        raise RuntimeError('无 计算模板/油耗计算 工作表: ' + fp)
    for r, off in plan:
        fi, fj = build_ij(r, off)
        _write_cell(ws.Cells(r, 9), fi, off, 'I')
        _write_cell(ws.Cells(r, 10), fj, off, 'J')
    if own:
        wb.Save()
        wb.Close(False)
    return len(plan)
