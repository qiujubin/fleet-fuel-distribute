# -*- coding: utf-8 -*-
"""每日加油名单分发脚本 v2
导出 -> 加油数据(末尾) -> 计算模板(G/F/是行公式) -> 油耗考核(sheet1/2)
全程按系统编号去重(幂等); 排除词只进加油数据; 混合记录跳过; 问题只在报告里说。
"""
import re, os, sys, glob, shutil, datetime
import openpyxl
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.utils import get_column_letter
from python_calamine import CalamineWorkbook

TPL_DIR  = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板'
KAOHE    = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
BACKUP   = r'E:\WorkSpace\workbuddy\2026-09-18-16-38-15\backup_20260919'
EXPORTS  = [r'C:\Users\Jubin\Desktop\数据\车辆加油_管理(2026-09-20).xlsx']
DT_FMT, D_FMT = 'yyyy-mm-dd hh:mm:ss', 'yyyy-mm-dd'
EXCLUDE_KW = ['发动机', '发电机', '小油箱', '新车']
MIN_FREE_ROWS = 15   # 填完后骨架余量下限(每用一行补一行, 余量恒定)

R = {'added': {}, 'shi': [], 'excluded': [], 'mixed': [], 'flags': [], 'errors': [],
     'kaohe_fuel': 0, 'kaohe_urea': 0, 'kaohe_skip': []}

def parse_dt(v):
    if isinstance(v, datetime.datetime): return v
    s = str(v).strip()
    for f in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d'):
        try: return datetime.datetime.strptime(s, f)
        except ValueError: pass
    return None

def num(v):
    try:
        f = float(v)
        return int(f) if f == int(f) else f
    except (TypeError, ValueError): return v

def shift_formula(text, src_row, dst_row):
    return re.sub(r'([A-Z])%d\b' % src_row, lambda m: m.group(1) + str(dst_row), text)

def last_data_row(ws, ncols=14):
    lr = 1
    for r in range(1, ws.max_row + 1):
        if any(ws.cell(r, c).value not in (None, '') for c in range(1, ncols + 1)):
            lr = r
    return lr

def get_sheet(wb, names):
    for n in names:
        if n in wb.sheetnames: return wb[n]
    return None

# ---------- 1. 读取导出 ----------
records = []
for fp in EXPORTS:
    wb = CalamineWorkbook.from_path(fp)
    rows = wb.get_sheet_by_name(wb.sheet_names[0]).to_python()
    hdr = [str(x).strip() for x in rows[0]]
    if '系统编号' not in hdr:
        R['errors'].append(f'导出表头异常, 跳过: {fp}'); continue
    for row in rows[1:]:
        d = dict(zip(hdr, row))
        sid = str(d.get('系统编号') or '').strip()
        if not sid: continue
        t = parse_dt(d.get('登记时间'))
        if t is None:
            R['errors'].append(f"记录{sid} 登记时间无法解析: {d.get('登记时间')}"); continue
        records.append({
            'sid': sid,
            'plate': str(d.get('车牌号码') or '').strip().upper(),
            'plate_raw': str(d.get('车牌号码') or '').strip(),
            'date': parse_dt(d.get('登记日期')),
            'mile': num(d.get('当前里程')), 'gps': num(d.get('GPS里程')),
            'amt': num(d.get('加注金额(元)')), 'vol': num(d.get('加注油量（升）')),
            'price': num(d.get('加注单价（元）')),
            'full': str(d.get('是否加满油') or '').strip(),
            'type': str(d.get('加油类型') or '').strip(),
            'loc': str(d.get('当前位置') or '').strip(),
            'person': str(d.get('登记人') or '').strip(),
            'time': t,
            'note': str(d.get('备注信息') or '').strip(),
        })
print(f'导出记录: {len(records)} 条')
records.sort(key=lambda x: x['time'])   # 按时间顺序处理, 保证同一文件的行号递增

# ---------- 2. 车牌映射 ----------
fuel_map, urea_map = {}, {}
for fp in glob.glob(os.path.join(TPL_DIR, '*.xlsx')):
    fn = os.path.basename(fp)
    if fn.startswith('~$') or fn.startswith('模板'): continue
    plate = os.path.splitext(fn)[0].split('-')[-1].strip().upper()
    if fn.startswith('尿素'): urea_map[plate] = fp
    else: fuel_map[plate] = fp

# ---------- 3. 备份 ----------
os.makedirs(BACKUP, exist_ok=True)
tpl_bak = os.path.join(BACKUP, '车辆油耗计算模板')
for plate in set(r['plate'] for r in records):
    for m in (fuel_map, urea_map):
        if plate in m and not os.path.exists(os.path.join(tpl_bak, os.path.basename(m[plate]))):
            shutil.copy2(m[plate], tpl_bak)
if not os.path.exists(os.path.join(BACKUP, os.path.basename(KAOHE))):
    shutil.copy2(KAOHE, BACKUP)
print('备份完成 ->', BACKUP)

# ---------- 4. 分发 ----------
wb_cache, dirty = {}, set()
fstate = {}   # path -> {'g_rows':[(row,time,full)], 'last_filled':n, 'prev_shi_G':t, 'cycle_start':n}

def get_fstate(path, wst_v, wsd_v):
    if path not in fstate:
        # 是否加满从加油数据反查(不依赖模板H缓存值——缓存已被清空)
        full_map = {}
        for r in range(2, wsd_v.max_row + 1):
            mt = wsd_v.cell(r, 13).value
            if mt is not None:
                full_map[str(mt)[:19]] = str(wsd_v.cell(r, 9).value or '').strip()
        g_rows = [(r, wst_v.cell(r, 7).value)
                  for r in range(2, wst_v.max_row + 1) if wst_v.cell(r, 7).value is not None]
        prev_shi_G, cycle_start = None, 2
        for rr, g in g_rows:
            if full_map.get(str(g)[:19]) == '是':
                prev_shi_G, cycle_start = g, rr + 1
        fstate[path] = {'g_rows': [(r, g, full_map.get(str(g)[:19])) for r, g in g_rows],
                        'last_filled': g_rows[-1][0] if g_rows else 1,
                        'prev_shi_G': prev_shi_G, 'cycle_start': cycle_start}
    return fstate[path]

def get_wbs(path):
    if path not in wb_cache:
        wb_cache[path] = (openpyxl.load_workbook(path, data_only=False),
                          openpyxl.load_workbook(path, data_only=True))
    return wb_cache[path]

def existing_sids(wsd_v):
    return set(str(wsd_v.cell(r, 1).value).strip() for r in range(2, wsd_v.max_row + 1)
               if wsd_v.cell(r, 1).value not in (None, ''))

def record_by_time(wsd_v, plate, t):
    for r in range(2, wsd_v.max_row + 1):
        mt = wsd_v.cell(r, 13).value
        if mt and str(mt)[:19] == str(t)[:19] and \
           str(wsd_v.cell(r, 2).value or '').strip().upper() == plate:
            return {'mile': num(wsd_v.cell(r, 4).value), 'gps': num(wsd_v.cell(r, 5).value)}
    return None

kaohe_pending = []
for rec in records:
    plate = rec['plate']
    is_urea_type = '尿素' in rec['type']
    # 单价纠偏分流: 2~3块=尿素, 6~8块=油气 (用户规则: 有人会把类型填乱)
    try: price_chk = float(rec['price'] or 0)
    except (TypeError, ValueError): price_chk = 0
    if price_chk:
        if not is_urea_type and price_chk < 5 and plate in urea_map:
            R['flags'].append(f"[类型/单价矛盾-已按尿素分发] {rec['sid']} {rec['plate_raw']}: 类型[{rec['type']}] 单价{price_chk} -> 尿素文件")
            is_urea_type = True
        elif is_urea_type and price_chk >= 5 and plate in fuel_map:
            R['flags'].append(f"[类型/单价矛盾-已按油气分发] {rec['sid']} {rec['plate_raw']}: 类型[{rec['type']}] 单价{price_chk} -> 油文件")
            is_urea_type = False
    if not is_urea_type and '尿素' in rec['note'] and '加油' in rec['note'].replace('油卡加油', ''):
        R['mixed'].append(f"{rec['sid']} {rec['plate_raw']} {rec['time']} 备注[{rec['note']}] -> 留手动处理")
        continue
    target_map = urea_map if is_urea_type else fuel_map
    path = target_map.get(plate)
    if not path:
        R['errors'].append(f"{rec['sid']} 车牌 {rec['plate_raw']} 找不到{'尿素' if is_urea_type else ''}模板文件")
        continue
    wb_f, wb_v = get_wbs(path)
    wsd_f, wsd_v = wb_f['加油数据'], wb_v['加油数据']

    if rec['sid'] in existing_sids(wsd_v):
        continue   # 已存在, 跳过

    # --- 追加 加油数据 ---
    nr = last_data_row(wsd_f) + 1
    vals = [rec['sid'], rec['plate_raw'],
            rec['date'].date() if rec['date'] else rec['time'].date(),
            rec['mile'], rec['gps'], rec['amt'], rec['vol'], rec['price'],
            rec['full'], rec['type'], rec['loc'], rec['person'], rec['time'], rec['note']]
    for c, v in enumerate(vals, 1):
        wsd_f.cell(nr, c, v)
    wsd_f.cell(nr, 3).number_format = D_FMT
    wsd_f.cell(nr, 13).number_format = DT_FMT
    dirty.add(path)
    R['added'].setdefault(os.path.basename(path), []).append(
        f"{rec['sid']} {rec['time']:%m-%d %H:%M} {rec['type']} 加满[{rec['full']}] ¥{rec['amt']} {rec['vol']}L @{rec['price']} 备注[{rec['note'] or '-'}]")

    # --- 校验 ---
    try:
        amt, vol, price = float(rec['amt'] or 0), float(rec['vol'] or 0), float(rec['price'] or 0)
        if amt and vol and price and abs(amt - vol * price) / max(amt, 1) > 0.05:
            R['flags'].append(f"[金额≠油量×单价] {rec['sid']} {rec['plate_raw']}: 金额{amt} 油量{vol} 单价{price} 乘积={round(vol*price,2)}")
        if is_urea_type and price > 4:
            R['flags'].append(f"[尿素单价异常] {rec['sid']} {rec['plate_raw']}: {price}")
        if not is_urea_type and 0 < price < 5 or (not is_urea_type and price > 9):
            R['flags'].append(f"[油单价异常] {rec['sid']} {rec['plate_raw']}: {price}")
    except (TypeError, ValueError):
        R['flags'].append(f"[数值列异常] {rec['sid']} {rec['plate_raw']}: {rec['amt']}/{rec['vol']}/{rec['price']}")
    if rec['full'] not in ('是', '否'):
        R['flags'].append(f"[是否加满异常] {rec['sid']} {rec['plate_raw']}: [{rec['full']}]")
    prev_r = None
    for r in range(2, wsd_v.max_row + 1):
        if str(wsd_v.cell(r, 2).value or '').strip().upper() == plate and wsd_v.cell(r, 13).value is not None:
            prev_r = r
    if prev_r:
        try:
            dd = float(rec['mile']) - float(wsd_v.cell(prev_r, 4).value)
            dg = float(rec['gps']) - float(wsd_v.cell(prev_r, 5).value)
            if dg > 5 and dd > 5 and abs(dd - dg) / dg > 0.05:
                R['flags'].append(f"[里程增量偏差] {rec['sid']} {rec['plate_raw']}: 仪表+{round(dd)} vs GPS+{round(dg)} ({abs(dd-dg)/dg:.0%})")
        except (TypeError, ValueError): pass

    # --- 排除词: 只进加油数据 ---
    if any(k in rec['note'] for k in EXCLUDE_KW):
        R['excluded'].append(f"{rec['sid']} {rec['plate_raw']} 备注[{rec['note']}] -> 只进加油数据")
        continue

    # --- 计算模板 ---
    wst_f = get_sheet(wb_f, ['计算模板', '油耗计算'])
    wst_v = get_sheet(wb_v, ['计算模板', '油耗计算'])
    if wst_f is None:
        R['errors'].append(f"{os.path.basename(path)} 无计算模板sheet"); continue

    st = get_fstate(path, wst_v, wsd_v)
    if str(rec['time'])[:19] in set(str(g)[:19] for _, g, _ in st['g_rows']):
        continue
    last_filled = st['last_filled']

    # 填 G / F / 是行公式
    r = last_filled + 1
    wst_f.cell(r, 7, rec['time']).number_format = DT_FMT
    prev_shi_G, cycle_start = st['prev_shi_G'], st['cycle_start']
    # 有效"是"规则: 同一天只算最后一个"是", 当天前面的"是"逻辑按"否"处理
    effective_shi = (rec['full'] == '是') and (
        prev_shi_G is None or rec['time'].date() != prev_shi_G.date())
    if effective_shi:
        if prev_shi_G is not None:
            wst_f.cell(r, 6, prev_shi_G).number_format = DT_FMT
            cycle_rows = list(range(cycle_start, r))
        else:
            first_t = None
            for rr in range(2, wsd_v.max_row + 1):
                if wsd_v.cell(rr, 13).value is not None:
                    first_t = wsd_v.cell(rr, 13).value; break
            wst_f.cell(r, 6, first_t).number_format = DT_FMT
            cycle_rows = list(range(2, r))
        base_i = f'=_xlfn.XLOOKUP(D{r}&G{r},加油数据!B:B&加油数据!M:M,加油数据!F:F,0)'
        base_j = f'=_xlfn.XLOOKUP(D{r}&G{r},加油数据!B:B&加油数据!M:M,加油数据!G:G,0)'
        off_i = ''.join(f'+I{x}' for x in cycle_rows)
        off_j = ''.join(f'+J{x}' for x in cycle_rows)
        wst_f.cell(r, 9, ArrayFormula(f'I{r}', base_i + off_i))
        wst_f.cell(r, 10, ArrayFormula(f'J{r}', base_j + off_j))
        R['shi'].append(f"{os.path.basename(path)} 模板第{r}行[有效是] F={prev_shi_G or '首周期'} G={rec['time']:%m-%d %H:%M} 偏移I={off_i or '无'}")
        kaohe_pending.append((path, r, rec, cycle_rows))
    elif rec['full'] == '是':
        R['shi'].append(f"{os.path.basename(path)} 模板第{r}行[当天非首个有效是->按否处理] G={rec['time']:%m-%d %H:%M}")
    # 推进该文件的映射状态(防止同文件多条记录写同一行)
    st['g_rows'].append((r, rec['time'], rec['full']))
    st['last_filled'] = r
    if effective_shi:
        st['prev_shi_G'] = rec['time']
        st['cycle_start'] = r + 1
    # 骨架扩充: 每消耗一行就补一行, 保证余量恒不低于 MIN_FREE_ROWS
    plate_v = str(wst_v.cell(2, 4).value or rec['plate_raw'])
    type_v = str(wst_v.cell(2, 5).value or '')
    need = max(1, (r + MIN_FREE_ROWS) - wst_f.max_row)
    if need > 0:
        src = wst_f.max_row
        for i in range(1, need + 1):
            rr2 = src + i
            wst_f.cell(rr2, 2, rr2 - 1); wst_f.cell(rr2, 4, plate_v); wst_f.cell(rr2, 5, type_v)
            for c in range(1, 25):
                if c in (2, 4, 5, 6, 7): continue
                v = wst_f.cell(src, c).value
                wst_f.cell(rr2, c).number_format = wst_f.cell(src, c).number_format
                if v is None: continue
                col = get_column_letter(c)
                if isinstance(v, ArrayFormula):
                    wst_f.cell(rr2, c, ArrayFormula(f'{col}{rr2}', shift_formula(v.text, src, rr2)))
                elif isinstance(v, str) and v.startswith('='):
                    wst_f.cell(rr2, c, shift_formula(v, src, rr2))
    dirty.add(path)

# ---------- 5. 保存模板 ----------
for path in dirty:
    wb_cache[path][0].save(path)
print(f'模板文件保存: {len(dirty)} 个被修改')

# ---------- 6. 油耗考核 ----------
from refuel_sheet import compute_rows, EXCLUDE_KW as _EK
wbk = None
if kaohe_pending:
    wbk = openpyxl.load_workbook(KAOHE, keep_vba=True)
    for path, r, rec, cycle_rows in kaohe_pending:
        wb_f, wb_v = get_wbs(path)
        wst_f = get_sheet(wb_f, ['计算模板', '油耗计算'])
        wst_v = get_sheet(wb_v, ['计算模板', '油耗计算'])
        wsd_v = wb_v['加油数据']
        is_urea = '尿素' in os.path.basename(path)
        wk = wbk['每日尿素数据' if is_urea else '每日油耗数据']
        wkv = wbk['每日尿素数据' if is_urea else '每日油耗数据']  # 新写的都是值, 直接读
        ids = set(str(wk.cell(x, 1).value).strip() for x in range(2, wk.max_row + 1)
                  if wk.cell(x, 1).value not in (None, '', '0'))
        if rec['sid'] in ids:
            R['kaohe_skip'].append(f"{rec['sid']} 已在考核表"); continue
        plate_v = str(wst_v.cell(2, 4).value or rec['plate_raw'])
        type_v = str(wst_v.cell(2, 5).value or '')
        amt, vol = float(rec['amt'] or 0), float(rec['vol'] or 0)
        for x in cycle_rows:
            gt = wst_f.cell(x, 7).value
            info = record_by_time(wsd_v, plate_v, gt) if gt else None
            if info:
                try: amt += float(info.get('amt') or 0)
                except (TypeError, ValueError): pass
                try: vol += float(info.get('vol') or 0)
                except (TypeError, ValueError): pass
        F_time = wst_f.cell(r, 6).value
        recF = record_by_time(wsd_f, plate_v, F_time)   # 用 wsd_f: 含本次新增行
        L = num(recF['gps']) if recF else None
        M = num(recF['mile']) if recF else None
        N, O = rec['gps'], rec['mile']
        try: P = round(float(N) - float(L), 4) if (L is not None and N is not None) else None
        except: P = None
        try: Q = round(float(O) - float(M), 4) if (M is not None and O is not None) else None
        except: Q = None
        try: Rv = round(float(vol) / P * 100, 6) if P else None
        except: Rv = None
        try: Sv = round(float(amt) / P, 6) if P else None
        except: Sv = None
        try: Tv = round(float(vol) / Q * 100, 6) if Q else None
        except: Tv = None
        nr = last_data_row(wk, 24) + 1
        Uv = round(float(amt) / Q, 10) if Q else None
        Vv = round((Rv + Tv) / 2, 10) if (Rv is not None and Tv is not None) else None
        rowvals = [rec['sid'], r - 1, rec['person'], plate_v, type_v,
                   F_time, rec['time'], '是', amt, vol, rec['price'], L, M, N, O, P, Q, Rv, Sv, Tv,
                   Uv, Vv, rec['time'].date(), rec['note'] or 0]
        for c, v in enumerate(rowvals, 1):
            cell = wk.cell(nr, c, v)
            if c in (6, 7): cell.number_format = DT_FMT
            if c == 23: cell.number_format = 'yyyy-mm-dd'
        if is_urea: R['kaohe_urea'] += 1
        else: R['kaohe_fuel'] += 1
# ---- 每日加油记录 sheet 刷新(与考核写入共用同一次打开/保存) ----
try:
    if wbk is None:
        wbk = openpyxl.load_workbook(KAOHE, keep_vba=True)
    rows_rf, src_rf, n1_rf = compute_rows()
    wsr = wbk['每日加油记录']
    last_r = 1
    for r in range(1, wsr.max_row + 1):
        if any(wsr.cell(r, c).value not in (None, '') for c in range(1, 15)): last_r = r
    for r in range(2, last_r + 1):
        for c in range(1, 15): wsr.cell(r, c, None)
    for i, vals in enumerate(rows_rf, start=2):
        for c, v in enumerate(vals, start=1): wsr.cell(i, c, v)
        wsr.cell(i, 3).number_format = D_FMT
        wsr.cell(i, 13).number_format = DT_FMT
    print(f"每日加油记录已刷新: 源={src_rf} n-1={n1_rf} 写入{len(rows_rf)}条")
except Exception as e:
    print('每日加油记录刷新失败:', e)

if wbk is not None:
    wbk.save(KAOHE)
    print('考核表已保存(含每日加油记录)')
    # openpyxl 会剥掉 WPS 专属部件(JS宏工程/按钮控件/图片), 立即从备份移植回来
    import subprocess
    subprocess.run([sys.executable, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'final_repair.py')],
                   check=False)

# ---------- 7. 报告 ----------
print('\n===== 分发报告 =====')
print('\n[新增加油数据]')
for fn, items in R['added'].items():
    print(f'  {fn}:')
    for it in items: print('     ', it)
if not R['added']: print('    (无, 全部已存在)')
print('\n[计算模板-新填"是"行]')
for it in R['shi']: print('   ', it)
if not R['shi']: print('    (无)')
print('\n[排除词-只进加油数据]')
for it in R['excluded']: print('   ', it)
if not R['excluded']: print('    (无)')
print('\n[混合记录-留手动]')
for it in R['mixed']: print('   ', it)
if not R['mixed']: print('    (无)')
print('\n[数据质量提醒]')
for it in R['flags']: print('   ', it)
if not R['flags']: print('    (无)')
print('\n[考核表追加] 每日油耗数据 +%d 行 | 每日尿素数据 +%d 行' % (R['kaohe_fuel'], R['kaohe_urea']))
for it in R['kaohe_skip']: print('    跳过:', it)
print('\n[错误]')
for it in R['errors']: print('   ', it)
if not R['errors']: print('    无')
