# -*- coding: utf-8 -*-
"""1) 全量修正计算模板 W 列(日期)数字格式为 yyyy-mm-dd
   2) 对账: 每辆车的"是"行 vs 油耗考核表, 找出差集"""
import os, glob, datetime
import openpyxl
from collections import defaultdict

TPL = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板'
KAOHE = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
W_FMT = 'yyyy-mm-dd'

def gs(wb):
    for n in ('计算模板', '油耗计算'):
        if n in wb.sheetnames: return wb[n]

# ---------- 1. W列格式修正 + 收集"是"行 ----------
vehicle_shi = defaultdict(set)   # plate_upper -> {(g19, sid_or '')}
fmt_fixed = []
for fp in sorted(glob.glob(os.path.join(TPL, '*.xlsx'))):
    fn = os.path.basename(fp)
    if fn.startswith('~$') or fn.startswith('模板'): continue
    plate = os.path.splitext(fn)[0].split('-')[-1].strip().upper()
    is_urea = fn.startswith('尿素')
    wb = openpyxl.load_workbook(fp, data_only=True)
    wst = gs(wb)
    if wst is None:
        print(f'[跳过] {fn}'); wb.close(); continue
    wsd = wb['加油数据']
    # sid 映射: (车牌, 登记时间)->sid
    rec_sid = {}
    rec_full = {}
    for row in wsd.iter_rows(min_row=2, values_only=True):
        if row and row[0] is not None:
            key = (str(row[1] or '').strip().upper(), str(row[12])[:19])
            rec_sid[key] = str(row[0]).strip()
            rec_full[key] = str(row[8] or '').strip()
    # 是行 + W格式
    changed = False
    for r in range(2, wst.max_row + 1):
        g = wst.cell(r, 7).value
        if wst.cell(r, 23).number_format != W_FMT:
            wst.cell(r, 23).number_format = W_FMT
            changed = True
        if g is not None:
            key = (plate, str(g)[:19])
            if rec_full.get(key) == '是':
                vehicle_shi[plate].add((str(g)[:19], rec_sid.get(key, '')))
    if changed:
        wb.save(fp)
        fmt_fixed.append(fn)
    wb.close()
print(f'W列格式修正: {len(fmt_fixed)} 个文件被修改')
print(f'车辆"是"行总数: {sum(len(v) for v in vehicle_shi.values())} (油{sum(len(v) for k,v in vehicle_shi.items() if not k in urea_plates) if False else 0})')

# ---------- 2. 考核表对账 ----------
urea_plates = set()
for fp in glob.glob(os.path.join(TPL, '尿素*.xlsx')):
    fn = os.path.basename(fp)
    if fn.startswith('~$'): continue
    urea_plates.add(os.path.splitext(fn)[0].split('-')[-1].strip().upper())

wbk = openpyxl.load_workbook(KAOHE, data_only=True)
kao = {}
for sn, tag in (('每日油耗数据', '油'), ('每日尿素数据', '尿素')):
    ws = wbk[sn]
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] in (None, '', '0'): continue
        plate = str(row[3] or '').strip().upper()
        g = row[6]
        g19 = str(g)[:19] if g else ''
        kao.setdefault(tag, {}).setdefault(plate, set()).add((g19, str(row[0]).strip()))

print('\n===== 对账结果 =====')
tot_missing = tot_extra = 0
all_plates = sorted(set(list(vehicle_shi.keys()) + [p for tag in kao.values() for p in tag]))
for plate in all_plates:
    tag = '尿素' if plate in urea_plates else '油'
    v_shi = vehicle_shi.get(plate, set())
    k_rows = kao.get(tag, {}).get(plate, set())
    missing = v_shi - k_rows     # 模板有"是"但考核没有
    extra = k_rows - v_shi       # 考核有但模板找不到对应"是"行
    if missing or extra:
        tot_missing += len(missing); tot_extra += len(extra)
        print(f'{plate} ({tag}): 模板是行{len(v_shi)} 考核{len(k_rows)} | 考核缺 {len(missing)} | 考核多 {len(extra)}')
        for g, sid in sorted(missing)[:6]: print(f'    缺: {g} sid={sid}')
        for g, sid in sorted(extra)[:6]: print(f'    多: {g} sid={sid}')
print(f'\n合计: 考核缺 {tot_missing} 条, 考核多(孤儿) {tot_extra} 条')
no_shi = [p for p in all_plates if not vehicle_shi.get(p)]
print(f'模板中无"是"行的车牌数: {len(no_shi)}')
