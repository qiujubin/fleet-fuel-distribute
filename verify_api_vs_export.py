# -*- coding: utf-8 -*-
"""验证：加油接口 JSON 能否 1:1 替代日常导出文件

做法：把接口数据映射成导出文件的 14 列格式，与最近的导出文件在**重叠日期**上逐条比对
      （按系统编号对齐），检查字段是否一致、有无遗漏/多余。

用法: python verify_api_vs_export.py
"""
import json, urllib.request, os, datetime

API_OIL = ('http://159.75.177.27:5180/XBM_RootDir/TPL_Base/Page_201_16.bsp'
           '?token=MY_SECRET_TOKEN_HERE&src=3&days=10&limit=1000')
EXPORT = r'C:\Users\Jubin\Desktop\数据\车辆加油_管理(2026-09-30).xlsx'

YESNO = {0: '是', 1: '否'}          # 接口 0=是(加满) 1=否 —— 与导出表相反!


def fetch(url):
    with urllib.request.urlopen(url, timeout=30) as r:
        return json.loads(r.read().decode('utf-8'))


def clean_sid(v):
    s = str(v or '').strip()
    return s[:-2] if s.endswith('.0') else s


def main():
    d = fetch(API_OIL)
    print('接口 status=%s params=%s 条数=%d' % (d.get('status'), d.get('params'), len(d.get('result', []))))

    # 接口 -> 导出 14 列
    api = {}
    for x in d['result']:
        sid = clean_sid(x.get('oiid'))
        api[sid] = {
            'sid': sid,
            'plate': str(x.get('vh_code') or '').strip(),
            'date': str(x.get('oi_time') or '')[:10],
            'mile': x.get('oi_smeter'),
            'gps': x.get('oi_gpsmeter'),
            'amt': x.get('oi_amount'),
            'vol': x.get('oi_oil'),
            'price': x.get('oi_price'),
            'full': YESNO.get(x.get('oi_yesno'), '?'),
            'type': str(x.get('oi_caretype') or '').strip(),
            'addr': str(x.get('oi_address') or '').strip(),
            'op': str(x.get('oi_op_name') or '').strip(),
            'op_time': str(x.get('oi_op_time') or '')[:19],
            'remark': str(x.get('oi_remark') or '').strip(),
        }
    days = sorted(set(v['date'] for v in api.values()))
    print('接口覆盖日期:', days)

    # 导出
    from python_calamine import CalamineWorkbook
    wb = CalamineWorkbook.from_path(EXPORT)
    rows = wb.get_sheet_by_name(wb.sheet_names[0]).to_python()
    exp = {}
    for row in rows[1:]:
        sid = clean_sid(row[0])
        if not sid: continue
        exp[sid] = row
    exp_days = sorted(set(str(r[2])[:10] for r in exp.values() if r[2]))
    print('导出文件日期:', exp_days, '条数', len(exp))

    only_exp = sorted(set(exp) - set(api))
    only_api = sorted(set(api) - set(exp))
    print()
    print('=== 集合差异 ===')
    print('  仅导出有:', len(only_exp), only_exp[:8])
    print('  仅接口有:', len(only_api), only_api[:8])

    # 逐字段比对（公共键）
    common = sorted(set(exp) & set(api))
    print()
    print('=== 公共记录 %d 条，字段比对 ===' % len(common))
    mismatch = {}
    for sid in common:
        a, e = api[sid], exp[sid]
        pairs = [
            ('车牌', a['plate'], str(e[1]).strip()),
            ('日期', a['date'], str(e[2])[:10]),
            ('当前里程', _f(a['mile']), _f(e[3])),
            ('GPS里程', _f(a['gps']), _f(e[4])),
            ('金额', _f(a['amt']), _f(e[5])),
            ('油量', _f(a['vol']), _f(e[6])),
            ('单价', _f(a['price']), _f(e[7])),
            ('是否加满', a['full'], str(e[8]).strip()),
            ('类型', a['type'], str(e[9]).strip()),
            ('登记人', a['op'], str(e[11]).strip()),
            ('登记时间', a['op_time'], str(e[12])[:19]),
            ('地址', a['addr'], str(e[10]).strip()),
        ]
        for name, x, y in pairs:
            if str(x) != str(y):
                mismatch.setdefault(name, []).append((sid, x, y))
    if not mismatch:
        print('  ✓ 全部字段一致')
    else:
        for name, items in mismatch.items():
            print('  ✗ %s: %d 处不一致' % (name, len(items)))
            for it in items[:3]:
                print('      %s: 接口[%s] vs 导出[%s]' % it)


def _f(v):
    if v is None: return ''
    try:
        f = float(v)
        return str(int(f)) if f == int(f) else repr(round(f, 6))
    except (TypeError, ValueError):
        return str(v).strip()


if __name__ == '__main__':
    main()
