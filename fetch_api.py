# -*- coding: utf-8 -*-
"""从数据接口取「加油/尿素」与「打冷」数据，生成与「车辆加油_管理(日期).xlsx」
完全同格式的表格，替代每天手工下载导出。

接口（帆软 .bsp 页面，直接返回 JSON）:
  加油+尿素  Page_201_16.bsp  → oiid / vh_code / oi_time / oi_smeter / oi_gpsmeter
                                / oi_amount / oi_oil / oi_price / oi_yesno / oi_caretype
                                / oi_address / oi_op_name / oi_op_time / oi_remark
  打冷       Page_201_17.bsp  → scid / sc_code / sc_time / sc_drvname / sc_op_name
                                / sc_op_time / sc_timelong / sc_remark

⚠️ **oi_yesno 与「是否加满油」相反**：0 → 是, 1 → 否
   （2026-10-01 用最近导出的 57 条记录逐条核对，13 个字段全部一致）

⚠️ 写入用 WPS COM（遵守「禁止 openpyxl 保存文件」铁律）；读取只走 urllib。

用法:
  python fetch_api.py                     # 取 3 天 → 数据/车辆加油_管理(<今天>).xlsx
  python fetch_api.py --days 7            # 拉 7 天
  python fetch_api.py --date 2026-10-01   # 指定文件名日期
  python fetch_api.py --out <dir>         # 指定输出目录（默认桌面「数据」）
  python fetch_api.py --dry-run           # 写到临时目录，不碰业务目录
  python fetch_api.py --also-cold         # 同时导出打冷表
"""
import os, sys, json, atexit, datetime, urllib.request, urllib.error

HERE = os.path.dirname(os.path.abspath(__file__))
CFG_FP = os.path.join(HERE, 'api_config.json')
DEFAULT_DATA_DIR = r'C:\Users\Jubin\Desktop\数据'

# 导出文件的 14 列（顺序与表头必须和系统导出一致）
OIL_HDR = ['系统编号', '车牌号码', '登记日期', '当前里程', 'GPS里程', '加注金额(元)',
           '加注油量（升）', '加注单价（元）', '是否加满油', '加油类型', '当前位置',
           '登记人', '登记时间', '备注信息']
COLD_HDR = ['系统编号', '车牌号码', '日期', '司机姓名', '登记人', '登记时间', '打冷时长', '备注信息']

YESNO = {0: '是', 1: '否'}      # ⚠️ 与直觉相反，别改


def load_cfg():
    with open(CFG_FP, encoding='utf-8') as f:
        c = json.load(f)
    c['token'] = os.environ.get('FUEL_API_TOKEN', c.get('token', ''))
    return c


def fetch(cfg, page, days=None, limit=None):
    days = days or cfg.get('days', 3)
    limit = limit or cfg.get('limit', 1000)
    if not 1 <= limit <= 1000:
        raise ValueError('limit 必须在 1-1000 之间（接口限制）')
    url = (f"{cfg['base'].rstrip('/')}/{page}?token={cfg['token']}"
           f"&src={cfg.get('src', '3')}&days={days}&limit={limit}")
    if cfg.get('msid'):
        url += f"&msid={cfg['msid']}"
    req = urllib.request.Request(url, headers={'User-Agent': 'fleet-fuel/1.0'})
    with urllib.request.urlopen(req, timeout=30) as r:
        data = json.loads(r.read().decode('utf-8'))
    st = data.get('status')
    if st != 0:
        raise RuntimeError(f"接口返回 status={st} message={data.get('message')} (url={page})")
    return data.get('result') or [], data.get('params')


def oil_table(result):
    """接口记录 -> 导出 14 列"""
    rows = []
    for x in result:
        sid = str(x.get('oiid') or '').strip()
        if not sid:
            continue
        y = x.get('oi_yesno')
        rows.append([
            sid,
            str(x.get('vh_code') or '').strip(),
            str(x.get('oi_time') or '')[:10],
            x.get('oi_smeter'),
            x.get('oi_gpsmeter'),
            x.get('oi_amount'),
            x.get('oi_oil'),
            x.get('oi_price'),
            YESNO.get(y, ''),                       # ⚠️ 取反映射
            str(x.get('oi_caretype') or '').strip(),
            str(x.get('oi_address') or '').strip(),
            str(x.get('oi_op_name') or '').strip(),
            str(x.get('oi_op_time') or '')[:19],
            str(x.get('oi_remark') or '').strip(),
        ])
    # 按登记时间升序（与系统导出习惯一致）
    rows.sort(key=lambda r: (str(r[2]), str(r[12])))
    return rows


def cold_table(result):
    rows = []
    for x in result:
        sid = str(x.get('scid') or '').strip()
        if not sid:
            continue
        rows.append([
            sid,
            str(x.get('sc_code') or '').strip(),
            str(x.get('sc_time') or '')[:10],
            str(x.get('sc_drvname') or '').strip(),
            str(x.get('sc_op_name') or '').strip(),
            str(x.get('sc_op_time') or '')[:19],
            x.get('sc_timelong'),
            str(x.get('sc_remark') or '').strip(),
        ])
    rows.sort(key=lambda r: (str(r[2]), str(r[5])))
    return rows


def write_xlsx_com(path, hdr, rows, col_text=(1,), sheet_name='Sheet1'):
    """用 WPS COM 新建并保存 xlsx（禁止 openpyxl 写文件）"""
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

        wb = app.Workbooks.Add()
        while wb.Worksheets.Count > 1:
            wb.Worksheets(wb.Worksheets.Count).Delete()
        ws = wb.Worksheets(1)
        try: ws.Name = sheet_name
        except Exception: pass

        ws.Range(ws.Cells(1, 1), ws.Cells(1, len(hdr))).Value = tuple([tuple(hdr)])
        if rows:
            ws.Range(ws.Cells(2, 1), ws.Cells(1 + len(rows), len(hdr))).Value = \
                tuple(tuple(r) for r in rows)
        for c in col_text:
            ws.Columns(c).NumberFormat = '@'      # 编号列防自动转数字

        os.makedirs(os.path.dirname(path), exist_ok=True)
        if os.path.exists(path):
            os.remove(path)
        wb.SaveAs(path, FileFormat=51)            # 51 = xlOpenXMLWorkbook (.xlsx)
        wb.Close(False)
        return True
    finally:
        if app is not None:
            try: app.Quit()
            except Exception: pass
        try: pythoncom.CoUninitialize()
        except Exception: pass


def main():
    args = sys.argv[1:]

    def opt(name, default=None):
        if name in args:
            i = args.index(name)
            if i + 1 < len(args):
                return args[i + 1]
        return default

    dry = '--dry-run' in args
    also_cold = '--also-cold' in args
    days = int(opt('--days', 0)) or None
    date_s = opt('--date') or datetime.date.today().isoformat()
    out_dir = opt('--out') or (os.path.join(HERE, 'outputs') if dry else DEFAULT_DATA_DIR)

    cfg = load_cfg()
    print(f"接口 {cfg['base']}  days={days or cfg.get('days')}  limit={cfg.get('limit')}")

    oil, p1 = fetch(cfg, cfg['oil_page'], days)
    print(f"加油/尿素: {len(oil)} 条  params={p1}")
    rows = oil_table(oil)
    print(f"  转成 14 列: {len(rows)} 条")
    from collections import Counter
    print('  是否加满分布:', dict(Counter(r[8] for r in rows)))
    print('  类型分布:', dict(Counter(r[9] for r in rows)))
    print('  日期分布:', dict(sorted(Counter(r[2] for r in rows).items())))
    if len(rows) >= (cfg.get('limit') or 1000):
        print('  ⚠️ 已达 limit 上限，可能有数据被截断，请加大 --days 分段或提高 limit')

    out_oil = os.path.join(out_dir, f'车辆加油_管理({date_s}).xlsx')
    write_xlsx_com(out_oil, OIL_HDR, rows)
    print(f'  ✅ 已生成 {out_oil}')

    if also_cold:
        cold, p2 = fetch(cfg, cfg['cold_page'], days)
        print(f"打冷: {len(cold)} 条  params={p2}")
        crows = cold_table(cold)
        out_cold = os.path.join(out_dir, f'车辆打冷_管理({date_s}).xlsx')
        write_xlsx_com(out_cold, COLD_HDR, crows)
        print(f'  ✅ 已生成 {out_cold}（{len(crows)} 条）')
    return 0


if __name__ == '__main__':
    sys.exit(main())
