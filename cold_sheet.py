# -*- coding: utf-8 -*-
"""把接口的「打冷记录」同步进考核表「打冷记录」sheet。

两层去重：
  ① 记录编号(scid) 去重 —— 编号已存在则跳过（历史遗留的重复编号不动）
  ② 内容去重 —— 抓「我在系统/本地写了，司机后补上传了一次」的重复：
       车牌相同 + 司机相同 + 打冷时长相同
       且 登记日期相差 <= DUP_DAY_TOL 天
       且 备注相似度 >= DUP_REMARK_TH
     → 判定为同一次打冷，跳过新增；保留表中已有那条（即"有系统编号那个"）。

接口: Page_201_17.bsp
  scid → 记录编号 | sc_code → 车牌号码 | sc_time → 登记日期 | sc_drvname → 司机姓名
  sc_op_name → 登记人 | sc_op_time → 登记时间 | sc_timelong → 打冷时长 | sc_remark → 备注信息

sheet 结构（实测）:
  行1 = 表头：记录编号|车牌号码|登记日期|司机姓名|登记人|登记时间|打冷时长|备注信息
  数据从行2起；历史上只写前 8 列（第14列有零散遗留值，本脚本不动）

写入用 WPS COM（遵守「禁止 openpyxl 保存文件」铁律）。

用法:
  python cold_sheet.py                 # 用 api_config.json 的 days 取数并同步
  python cold_sheet.py --days 10       # 指定天数
  python cold_sheet.py --check         # 只看会做什么，不写入
  python cold_sheet.py --sim 0.6       # 调备注相似度阈值（默认 0.5）
  python cold_sheet.py --tol 0         # 调日期容忍天数（默认 1）
"""
import os, re, sys, json, datetime, difflib, urllib.request
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
CFG_FP = os.path.join(HERE, 'api_config.json')
KAOHE = os.environ.get('FUEL_KAOHE', r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm')
SHEET = '打冷记录'
NCOL = 8                     # 只写前 8 列

DUP_DAY_TOL = 1              # 内容去重：日期容忍（天）
DUP_REMARK_TH = 0.5          # 内容去重：备注相似度阈值
SYS_ID_RE = re.compile(r'^\d{12}$')   # 系统编号形态（12 位数字）


# ----------------------------------------------------------- 基础工具
def load_cfg():
    with open(CFG_FP, encoding='utf-8') as f:
        return json.load(f)


def fetch_cold(cfg, days):
    if not 1 <= cfg.get('limit', 1000) <= 1000:
        raise ValueError('limit 必须在 1-1000 之间')
    url = (f"{cfg['base'].rstrip('/')}/{cfg['cold_page']}?token={cfg['token']}"
           f"&src={cfg.get('src', '3')}&days={days}&limit={cfg.get('limit', 1000)}")
    if cfg.get('msid'):
        url += f"&msid={cfg['msid']}"
    req = urllib.request.Request(url, headers={'User-Agent': 'fleet-fuel/1.0'})
    with urllib.request.urlopen(req, timeout=30) as r:
        d = json.loads(r.read().decode('utf-8'))
    if d.get('status') != 0:
        raise RuntimeError(f"接口 status={d.get('status')} message={d.get('message')}")
    return d.get('result') or []


def clean_sid(v):
    s = str(v or '').strip()
    return s[:-2] if s.endswith('.0') else s


def is_sys_id(s):
    return bool(SYS_ID_RE.match(clean_sid(s)))


def to_row(x):
    return [clean_sid(x.get('scid')),
            str(x.get('sc_code') or '').strip(),
            str(x.get('sc_time') or '')[:10],
            str(x.get('sc_drvname') or '').strip(),
            str(x.get('sc_op_name') or '').strip(),
            str(x.get('sc_op_time') or '')[:19],
            x.get('sc_timelong'),
            str(x.get('sc_remark') or '').strip()]


def from_serial(v):
    """日期序列数 / 字符串 -> date（不用 openpyxl）"""
    if v is None:
        return None
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, (int, float)):
        try:
            return (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(v))).date()
        except Exception:
            return None
    s = str(v).strip()[:10]
    try:
        return datetime.date(*map(int, s.replace('/', '-').split('-')))
    except Exception:
        return None


def _blank_norm(s):
    """去空白（含全角空格）便于车牌/司机比对"""
    return re.sub(r'\s|\u3000', '', str(s or ''))


def norm_remark(s):
    """备注规范化：只留中文字/数字/字母，去掉标点空白"""
    return re.sub(r'[^\u4e00-\u9fff0-9A-Za-z]', '', str(s or ''))


def remark_sim(a, b):
    """备注相似度：2-gram Jaccard 与序列相似度取较大者"""
    a, b = norm_remark(a), norm_remark(b)
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    def grams(x):
        return {x[i:i + 2] for i in range(len(x) - 1)} or {x}
    ga, gb = grams(a), grams(b)
    jac = len(ga & gb) / len(ga | gb)
    seq = difflib.SequenceMatcher(None, a, b).ratio()
    return max(jac, seq)


def num_eq(a, b):
    try:
        return abs(float(a) - float(b)) < 1e-6
    except (TypeError, ValueError):
        return str(a).strip() == str(b).strip()


def dup_fingerprint(row):
    """内容指纹：车牌 + 司机 + 时长"""
    return (_blank_norm(row[1]), _blank_norm(row[3]), round(float(row[6]), 6)
            if _is_num(row[6]) else str(row[6]).strip())


def _is_num(v):
    try:
        float(v); return True
    except (TypeError, ValueError):
        return False


def find_dups(row, index, tol_days, sim_th):
    """在 index（指纹 -> [(行号, 行数据, 日期)]）里找与 row 内容重复的记录"""
    hits = []
    d_new = from_serial(row[2])
    for line_no, other, d_old in index.get(dup_fingerprint(row), []):
        if not d_new or not d_old:
            continue
        dd = abs((d_new - d_old).days)
        if dd > tol_days:
            continue
        s = remark_sim(row[7], other[7])
        if s >= sim_th:
            hits.append((line_no, other, dd, s))
    hits.sort(key=lambda h: -h[3])
    return hits


def build_index(table):
    """表内记录 -> 内容指纹索引（只用于判重，不改动历史）"""
    index = defaultdict(list)
    for line_no, r in table:
        if _blank_norm(r[1]):
            index[dup_fingerprint(r)].append((line_no, r, from_serial(r[2])))
    return index


def dedup_plan(rows, table, tol_days=DUP_DAY_TOL, sim_th=DUP_REMARK_TH):
    """把接口记录分成四类（纯函数，便于测试）。

    返回 (编号已存在, 与表内重复, 批次内重复, 需新增, 表内那条无系统编号的提醒)
      与表内重复: (新记录, 表内已有记录, 日期差, 相似度, 表内行号)
      批次内重复: (新记录, 保留的记录, 日期差, 相似度)
    """
    exist = {clean_sid(r[0]) for _, r in table if clean_sid(r[0])}
    index = build_index(table)

    n_exist, dup_a, dup_b, accepted, no_id_warn = [], [], [], [], []
    acc_index = defaultdict(list)
    for r in rows:
        if r[0] in exist:
            n_exist.append(r)
            continue
        hits = find_dups(r, index, tol_days, sim_th)
        if hits:
            line_no, other, dd, s = hits[0]
            dup_a.append((r, other, dd, s, line_no))
            if not is_sys_id(other[0]):
                no_id_warn.append((r, other, line_no))
            continue
        hits = find_dups(r, acc_index, tol_days, sim_th)
        if hits:
            _, other, dd, s = hits[0]
            dup_b.append((r, other, dd, s))
            continue
        accepted.append(r)
        acc_index[dup_fingerprint(r)].append((None, r, from_serial(r[2])))
    return n_exist, dup_a, dup_b, accepted, no_id_warn


# ----------------------------------------------------------- 主流程
def main():
    args = sys.argv[1:]
    check = '--check' in args
    def opt(name, cast, default):
        if name in args:
            i = args.index(name)
            if i + 1 < len(args):
                try:
                    return cast(args[i + 1])
                except Exception:
                    pass
        return default
    days = opt('--days', int, None)
    tol = opt('--tol', int, DUP_DAY_TOL)
    sim_th = opt('--sim', float, DUP_REMARK_TH)

    cfg = load_cfg()
    days = days or cfg.get('days', 3)
    recs = fetch_cold(cfg, days)
    rows = [to_row(x) for x in recs if clean_sid(x.get('scid'))]
    rows.sort(key=lambda r: (r[2], r[5]))          # 按日期+登记时间升序
    print(f'接口打冷: {len(recs)} 条（最近 {days} 天）-> 有效 {len(rows)} 条')
    if not rows:
        print('接口无数据，跳过'); return 0

    import pythoncom
    import win32com.client as wc
    pythoncom.CoInitialize()
    app = None
    try:
        for pid in ('Ket.Application', 'et.Application', 'Excel.Application'):
            try:
                app = wc.DispatchEx(pid); break
            except Exception:
                continue
        if app is None:
            raise RuntimeError('无法启动 WPS/Excel COM')
        app.Visible = False
        app.DisplayAlerts = False
        try: app.ScreenUpdating = False
        except Exception: pass

        wb = app.Workbooks.Open(KAOHE, 0, bool(check))
        if not check and wb.ReadOnly:
            raise RuntimeError('考核表被占用（只读）')
        try:
            try:
                ws = wb.Worksheets(SHEET)
            except Exception:
                cand = [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)
                        if '打冷' in wb.Worksheets(i).Name]
                print(f'  !! 找不到「{SHEET}」，候选：{cand}')
                raise

            last = _true_last_row(ws, 1)
            print(f'  「{SHEET}」现有数据行: {max(last - 1, 0)}（末行 {last}）')

            table = _read_table(ws, last, NCOL)
            n_exist, dup_a, dup_b, accepted, no_id_warn = \
                dedup_plan(rows, table, tol, sim_th)
            print(f'  现有编号 {len({clean_sid(r[0]) for _, r in table if clean_sid(r[0])})} 个')

            print(f'  编号已存在跳过: {len(n_exist)} 条')
            print(f'  内容重复跳过 : {len(dup_a) + len(dup_b)} 条'
                  f'（表内已有 {len(dup_a)} / 本次批次内 {len(dup_b)}）')
            print(f'  需新增       : {len(accepted)} 条')

            if dup_a or dup_b:
                print()
                print('  ' + '=' * 66)
                print('  ⚠️  检测到疑似「重复登记」（司机后补上传），已跳过不新增')
                print('      判据：车牌+司机+打冷时长相同，登记日期相差≤%d天，备注相似度≥%.2f'
                      % (tol, sim_th))
                print('  ' + '=' * 66)
                for r, other, dd, s, ln in dup_a:
                    print(f'  ▸ 新来（本次不入库）: {r[0]}  {r[1]}  {r[2]}  时长 {_fmt_len(r[6])} 分钟  登记人={r[4]}')
                    print(f'        备注: {r[7][:60]}')
                    print(f'    已有（保留，第{ln}行）: {clean_sid(other[0])}  {other[1]}  {str(other[2])[:10]}'
                          f'  时长 {_fmt_len(other[6])} 分钟  登记人={other[4]}')
                    print(f'        备注: {str(other[7])[:60]}')
                    print(f'        → 相似度 {s:.2f} ｜ 日期差 {dd} 天'
                          + ('' if is_sys_id(other[0]) else '  ｜ ⚠️ 表内那条编号非系统编号，请人工确认'))
                for r, other, dd, s in dup_b:
                    print(f'  ▸ 本次批次内重复，保留较早编号: {other[0]}  {other[1]}  {other[2]}'
                          f'  时长 {_fmt_len(other[6])} 分钟')
                    print(f'    被丢弃: {r[0]}  {r[1]}  {r[2]}  时长 {_fmt_len(r[6])} 分钟'
                          f'  (相似度 {s:.2f} ｜ 日期差 {dd} 天)')
                print('  ' + '=' * 66)
                print('  → 如确认其中某条确实需要单独保留，请手工添加。')
                print()

            if accepted:
                print('  将新增:')
                for r in accepted[:10]:
                    print(f'     {r[0]} {r[1]} {r[2]} 时长{_fmt_len(r[6])} {r[7][:24]}')
                if len(accepted) > 10:
                    print(f'     ...（共 {len(accepted)} 条）')

            _write_report(check, days, rows, n_exist, dup_a, dup_b, accepted, tol, sim_th)

            if check:
                print('  (--check 模式，不写入)'); return 0
            if not accepted:
                print('  无新增，无需写入'); return 0

            start = last + 1
            ws.Cells(start, 1).NumberFormat = '@'      # 记录编号列：文本，防转数字
            ws.Range(ws.Cells(start, 1), ws.Cells(start + len(accepted) - 1, NCOL)).Value = \
                tuple(tuple(r) for r in accepted)
            wb.Save()
            print(f'  ✅ 已追加 {len(accepted)} 行（{start} ~ {start + len(accepted) - 1}）并保存（WPS 原生保存）')
            if dup_a or dup_b:
                print(f'  ⚠️  另有 {len(dup_a) + len(dup_b)} 条疑似重复已跳过，明细见上方提示')
        finally:
            try: wb.Close(False)
            except Exception: pass
        return 0
    finally:
        if app is not None:
            try: app.Quit()
            except Exception: pass
        try: pythoncom.CoUninitialize()
        except Exception: pass


def _fmt_len(v):
    try:
        f = float(v)
        return str(int(f)) if abs(f - int(f)) < 1e-9 else str(f)
    except (TypeError, ValueError):
        return str(v)


def _read_table(ws, last, ncol):
    """一次读回 A2:H{last}，返回 [(行号, [8 列])]"""
    if last < 2:
        return []
    vals = ws.Range(ws.Cells(2, 1), ws.Cells(last, ncol)).Value2
    if vals is None:
        return []
    if not isinstance(vals, tuple):
        vals = ((vals,),)
    out = []
    for i, row in enumerate(vals):
        if not isinstance(row, tuple):
            row = (row,)
        row = list(row) + [None] * (ncol - len(row))
        out.append((2 + i, row[:ncol]))
    return out


def _write_report(check, days, rows, n_exist, dup_a, dup_b, accepted, tol, sim_th):
    """把本次判重结果留档到 outputs/（不入库）"""
    try:
        out_dir = os.path.join(HERE, 'outputs')
        os.makedirs(out_dir, exist_ok=True)
        fp = os.path.join(out_dir, '打冷同步报告_%s.txt' % datetime.date.today().strftime('%Y-%m-%d'))
        lines = [f"打冷记录同步报告  {datetime.datetime.now():%Y-%m-%d %H:%M:%S}",
                 f"接口 {days} 天: {len(rows)} 条 | 编号已存在: {len(n_exist)} | "
                 f"内容重复跳过: {len(dup_a) + len(dup_b)} | 新增: {len(accepted)}"
                 f"{'  (--check 未写入)' if check else ''}",
                 f"判据: 车牌+司机+时长相同 且 日期差<={tol}天 且 备注相似度>={sim_th}", '']
        if dup_a:
            lines.append('【与表内已有记录重复 → 保留表内那条】')
            for r, other, dd, s, ln in dup_a:
                lines.append(f"  跳过 {r[0]} {r[1]} {r[2]} 时长{_fmt_len(r[6])} 登记人={r[4]} | 备注: {r[7]}")
                lines.append(f"  保留 {clean_sid(other[0])} {other[1]} {str(other[2])[:10]} "
                             f"时长{_fmt_len(other[6])} 登记人={other[4]} (第{ln}行) | 备注: {other[7]}")
                lines.append(f"       相似度 {s:.2f} / 日期差 {dd} 天")
        if dup_b:
            lines.append('【本次批次内重复 → 保留较早编号】')
            for r, other, dd, s in dup_b:
                lines.append(f"  保留 {other[0]} {other[1]} {other[2]} | 丢弃 {r[0]} ({s:.2f})")
        if accepted:
            lines.append('')
            lines.append(f'【新增 {len(accepted)} 条】')
            for r in accepted:
                lines.append(f"  {r[0]} {r[1]} {r[2]} 时长{_fmt_len(r[6])} {r[7]}")
        with open(fp, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines))
        print(f'  报告已留档: {os.path.relpath(fp, HERE)}')
    except Exception as e:
        print(f'  (报告留档失败: {e})')


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


if __name__ == '__main__':
    sys.exit(main())
