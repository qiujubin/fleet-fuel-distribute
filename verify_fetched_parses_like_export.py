# -*- coding: utf-8 -*-
"""影子验证：用「主脚本同款解析逻辑」读【接口生成的 xlsx】与【手工导出 xlsx】，
   比较解析结果是否逐条一致 —— 证明接口文件可以无缝替代导出文件。

用法: python verify_fetched_parses_like_export.py
"""
import os, datetime, json
from python_calamine import CalamineWorkbook

GEN = r'E:\WorkSpace\workbuddy\2026-09-18-16-38-15\outputs\车辆加油_管理(2026-10-01).xlsx'
EXP = r'C:\Users\Jubin\Desktop\数据\车辆加油_管理(2026-09-30).xlsx'


# ---- 与 fleet_fuel_com.py 完全相同的解析实现 ----
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


def sid_key(v):
    s = str(v).strip()
    try:
        f = float(s)
        if f == int(f): return str(int(f))
    except ValueError:
        pass
    return s


def read(fp):
    wb = CalamineWorkbook.from_path(fp)
    rows = wb.get_sheet_by_name(wb.sheet_names[0]).to_python(skip_empty_area=False)
    hdr = [str(x).strip() for x in rows[0]]
    out = {}
    bad = []
    for i, row in enumerate(rows[1:], start=2):
        d = dict(zip(hdr, row))
        sid = sid_key(d.get('系统编号'))
        if not sid or sid == '': continue
        t = parse_dt(d.get('登记时间'))
        if t is None:
            bad.append((i, d.get('系统编号'), d.get('登记时间')))
            continue
        out[sid] = {
            'sid': sid,
            'plate': str(d.get('车牌号码') or '').strip().upper(),
            'date': parse_dt(d.get('登记日期')),
            'time': t,
            'mile': num(d.get('当前里程')),
            'gps': num(d.get('GPS里程')),
            'amt': num(d.get('加注金额(元)')),
            'vol': num(d.get('加注油量（升）')),
            'price': num(d.get('加注单价（元）')),
            'full': str(d.get('是否加满油') or '').strip(),
            'type': str(d.get('加油类型') or '').strip(),
            'loc': str(d.get('当前位置') or '').strip(),
            'person': str(d.get('登记人') or '').strip(),
            'note': str(d.get('备注信息') or '').strip(),
        }
    return hdr, out, bad


def _eq(a, b):
    try:
        return abs(float(a) - float(b)) < 1e-6
    except (TypeError, ValueError):
        return str(a) == str(b)


def main():
    h1, g, bad1 = read(GEN)
    h2, e, bad2 = read(EXP)
    print('生成文件表头:', h1)
    print('导出文件表头:', h2)
    print('表头一致:', h1 == h2)
    print(f'生成解析 {len(g)} 条 (解析失败 {len(bad1)}) | 导出解析 {len(e)} 条 (解析失败 {len(bad2)})')
    if bad1: print('  生成侧解析失败样本:', bad1[:3])
    if bad2: print('  导出侧解析失败样本:', bad2[:3])

    common = sorted(set(g) & set(e))
    print(f'\n公共记录 {len(common)} 条，逐字段比对:')
    fields = ['plate', 'time', 'mile', 'gps', 'amt', 'vol', 'price', 'full', 'type', 'loc', 'person', 'note']
    diff = {}
    for sid in common:
        for f in fields:
            a, b = g[sid][f], e[sid][f]
            if not _eq(a, b):
                diff.setdefault(f, []).append((sid, a, b))
    if not diff:
        print('  ✓ 全部字段一致（接口文件可无缝替代导出）')
    else:
        for f, items in diff.items():
            print(f'  ✗ {f}: {len(items)} 处')
            for it in items[:3]: print('      %s: 生成[%s] vs 导出[%s]' % it)

    print('\n日期列类型检查（date 是否解析为 datetime）:')
    for tag, m in (('生成', g), ('导出', e)):
        types = set(type(v['date']).__name__ for v in m.values())
        print(f'  {tag}: {types}')


if __name__ == '__main__':
    main()
