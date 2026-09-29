# -*- coding: utf-8 -*-
"""每日加油记录 行计算(纯解析, 不开考核表)
源: 数据文件夹最新 车辆加油_管理 导出; 取 登记日期=n-1 的记录
规则: 排除尿素/错标尿素(单价<5);
      排除词记录(发动机/发电机/小油箱/新车)保留但"是"改"否", 独立占一行(不参与同日压缩);
      其余同车牌同日: 有"是"留最后一个"是", 全"否"留最后一个。
"""
import os, glob, re, datetime
from python_calamine import CalamineWorkbook

DATA_DIR = r'C:\Users\Jubin\Desktop\数据'
EXCLUDE_KW = ['发动机', '发电机', '小油箱', '新车']

def parse_dt(v):
    if isinstance(v, datetime.datetime): return v
    s = str(v).strip()
    for f in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d'):
        try: return datetime.datetime.strptime(s, f)
        except ValueError: pass
    return None

def num(v):
    try: return float(v)
    except (TypeError, ValueError): return v

def compute_rows(data_dir=DATA_DIR, run_date=None, verbose=True):
    """返回 (行列表, 源文件名, n1) 。行 = 14 列值列表。"""
    files = []
    for fp in glob.glob(os.path.join(data_dir, '车辆加油_管理*.xlsx')):
        m = re.search(r'\((\d{4}-\d{2}-\d{2})\)', os.path.basename(fp))
        if m: files.append((m.group(1), fp))
    files.sort(reverse=True)
    n = run_date or datetime.date.today()
    n1 = n - datetime.timedelta(days=1)
    recs, used = [], None
    for d, fp in files:
        wb = CalamineWorkbook.from_path(fp)
        rows = wb.get_sheet_by_name(wb.sheet_names[0]).to_python()
        if not rows: continue
        hdr = [str(x).strip() for x in rows[0]]
        if '系统编号' not in hdr: continue
        used = os.path.basename(fp)
        for row in rows[1:]:
            r = dict(zip(hdr, row))
            if not str(r.get('系统编号') or '').strip(): continue
            dd = parse_dt(r.get('登记日期'))
            if dd is None or dd.date() != n1: continue
            recs.append({
                'plate': str(r.get('车牌号码') or '').strip().upper(),
                'time': parse_dt(r.get('登记时间')),
                'vals': [str(r.get('系统编号') or '').strip(), str(r.get('车牌号码') or '').strip(),
                         dd, num(r.get('当前里程')), num(r.get('GPS里程')),
                         num(r.get('加注金额(元)')), num(r.get('加注油量（升）')),
                         num(r.get('加注单价（元）')), str(r.get('是否加满油') or '').strip(),
                         str(r.get('加油类型') or '').strip(), str(r.get('当前位置') or '').strip(),
                         str(r.get('登记人') or '').strip(), parse_dt(r.get('登记时间')),
                         str(r.get('备注信息') or '').strip()]})
        break
    pool, special = [], []      # special = 排除词记录(发动机/发电机/小油箱/新车): 独立保留
    for x in recs:
        if '尿素' in x['vals'][9]: continue                       # 尿素类型
        p = x['vals'][7]
        if isinstance(p, float) and 0 < p < 5: continue            # 错标成加油的尿素
        if any(k in x['vals'][13] for k in EXCLUDE_KW):
            x['vals'][8] = '否'                                    # "是"强制改"否"
            special.append(x)                                      # 不与正常记录竞争同日唯一名额
        else:
            pool.append(x)
    by_plate = {}
    for x in pool: by_plate.setdefault(x['plate'], []).append(x)
    keep = []
    for plate, items in by_plate.items():
        items.sort(key=lambda z: str(z['time'])[:19])
        shi = [z for z in items if z['vals'][8] == '是']
        keep.append(shi[-1] if shi else items[-1])
    # 排除词记录: 同车同日多条只留最后一条(防重复录入), 但不与正常记录抢名额
    sp_plate = {}
    for x in special: sp_plate.setdefault(x['plate'], []).append(x)
    for plate, items in sp_plate.items():
        items.sort(key=lambda z: str(z['time'])[:19])
        keep.append(items[-1])
    keep.sort(key=lambda z: str(z['time'])[:19])
    if verbose:
        print(f"[每日加油记录] 源={used} n-1={n1} 原始{len(recs)} 候选{len(pool)} 保留{len(keep)}")
    return [z['vals'] for z in keep], used, n1

if __name__ == '__main__':
    rows, src, n1 = compute_rows()
    for i, v in enumerate(rows, 1):
        print(i, v[0], v[1], str(v[12])[:19], v[8], v[5], v[6], v[13])
