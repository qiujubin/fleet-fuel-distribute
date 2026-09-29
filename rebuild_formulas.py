# -*- coding: utf-8 -*-
"""⛔ 已停用：本模块原用 openpyxl 保存车辆模板。

2026-09-28 用户明令：**任何写入只准 win32/COM，禁止 openpyxl 保存文件**。
- `rebuild()` 已改为直接抛异常 —— 请改用 `com_formula.rebuild_ij(app, fp)`。
- 这里只保留 `tkey()` 与公式拼装函数 `build_formulas()`（纯计算，无副作用），
  供只读诊断/参考；不要再用本模块写入任何文件。

历史说明（重建计算模板公式，修复 data_only=True 保存事故）——公式模式:
  A: =_xlfn.XLOOKUP(D&G,加油数据!B:B&加油数据!M:M,加油数据!A:A,0)
  C: -> L:L   H: -> I:I   K: -> H:H
  I: -> F:F (+周期累计)   J: -> G:G (+周期累计)
  L: =_xlfn.XLOOKUP(D&F,...,E:E,0)   M: -> D:D
  N: =_xlfn.XLOOKUP(D&G,...,E:E,0)   O: -> D:D
  P=N-L  Q=O-M  R=J/P*100  S=I/P  T=J/Q*100  U=I/Q  V=(R+T)/2
  W: =_xlfn.XLOOKUP(G,加油数据!M:M,加油数据!C:C)   X: -> N:N
周期偏移(F锚点规则): 是行 -> 找到 G==F 的锚行, 偏移=锚行+1..本行-1 (降序);
                     若无锚行(首周期) -> 从第一个已填行到本行-1
"""
import os, datetime

TPL = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板'
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
    """任意时间表示(datetime / 序列数 / 字符串) -> 整数秒匹配键。

    2026-09-28 修复: 原实现直接 `str(F)[:19]` 与 `str(G)[:19]` 比对, 当 F 由
    retro_fix 写成**序列数**(如 46291.4215393519) 而 G 是 datetime 时字符串
    形态不同 -> 锚行找不到 -> 退化成"首周期" 从第 2 行全量累加,
    I/J 被堆成整表总和(实测 C3 粤BPV550 行271 变成 586719.09 / 82760.16)。
    """
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


def build_formulas(r, F, G, off_rows, special=None):
    """拼装「计算模板」一行 18 列公式（纯字符串，无副作用）"""
    dg = f'D{r}&G{r}'
    df = f'D{r}&F{r}'
    oi = ''.join(f'+I{x}' for x in off_rows)
    oj = ''.join(f'+J{x}' for x in off_rows)
    return {
        'A': f'={X}({dg},加油数据!B:B&加油数据!M:M,加油数据!A:A,0)',
        'C': f'={X}({dg},加油数据!B:B&加油数据!M:M,加油数据!L:L,0)',
        'H': f'={X}({dg},加油数据!B:B&加油数据!M:M,加油数据!I:I,0)',
        'I': f'={X}({dg},加油数据!B:B&加油数据!M:M,加油数据!F:F,0)' + (special or {}).get('I', oi),
        'J': f'={X}({dg},加油数据!B:B&加油数据!M:M,加油数据!G:G,0)' + (special or {}).get('J', oj),
        'K': f'={X}({dg},加油数据!B:B&加油数据!M:M,加油数据!H:H,0)',
        'L': f'={X}({df},加油数据!B:B&加油数据!M:M,加油数据!E:E,0)',
        'M': f'={X}({df},加油数据!B:B&加油数据!M:M,加油数据!D:D,0)',
        'N': f'={X}({dg},加油数据!B:B&加油数据!M:M,加油数据!E:E,0)',
        'O': f'={X}({dg},加油数据!B:B&加油数据!M:M,加油数据!D:D,0)',
        'P': f'=N{r}-L{r}', 'Q': f'=O{r}-M{r}',
        'R': f'=J{r}/P{r}*100', 'S': f'=I{r}/P{r}',
        'T': f'=J{r}/Q{r}*100', 'U': f'=I{r}/Q{r}', 'V': f'=(R{r}+T{r})/2',
        'W': f'={X}(G{r},加油数据!M:M,加油数据!C:C)',
        'X': f'={X}(G{r},加油数据!M:M,加油数据!N:N)',
    }


def rebuild(fp, specials=None, verbose=False):
    raise RuntimeError(
        '⛔ rebuild() 已停用：禁止用 openpyxl 保存文件（2026-09-28 用户明令）。'
        '请改用 com_formula.rebuild_ij(app, fp) —— WPS COM，原生保存。')
