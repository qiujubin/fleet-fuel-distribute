# -*- coding: utf-8 -*-
"""验证 cold_sheet 的「内容重复」判定逻辑（只读，不碰业务文件）。

用「打冷记录」sheet 的真实数据当基准，构造 8 个用例：
  应判重复（≤1天 & 备注相近） / 不应判重复（日期远、备注不同、时长不同、司机不同）
"""
import os, sys, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from collections import defaultdict
from python_calamine import CalamineWorkbook
from cold_sheet import (dup_fingerprint, find_dups, remark_sim, from_serial,
                        clean_sid, DUP_DAY_TOL, DUP_REMARK_TH)

KAOHE = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'

wb = CalamineWorkbook.from_path(KAOHE)
rows = wb.get_sheet_by_name('打冷记录').to_python(skip_empty_area=False)
table = [(i, r + [None] * (8 - len(r))) for i, r in enumerate(rows, 1)
         if i >= 2 and r and r[0] not in (None, '')]
print(f'表内数据行 {len(table)}')

index = defaultdict(list)
for ln, r in table:
    if str(r[1] or '').strip():
        index[dup_fingerprint(r)].append((ln, r, from_serial(r[2])))

# 拿一条真实记录当基准（选一条备注较长的，便于构造"备注改写"）
base_ln, base = None, None
for ln, r in table:
    if len(str(r[7] or '')) >= 18 and from_serial(r[2]):
        base_ln, base = ln, r
        break
print(f'基准记录: 第{base_ln}行 {clean_sid(base[0])} {base[1]} {str(base[2])[:10]} '
      f'司机={base[3]} 时长={base[6]}')
print(f'         备注: {base[7]}')
print()

d0 = from_serial(base[2])


def mk(sid, *, plate=None, date=None, driver=None, timelong=None, remark=None):
    r = list(base)
    r[0] = sid
    if plate is not None:     r[1] = plate
    if date is not None:      r[2] = date
    if driver is not None:    r[3] = driver
    if timelong is not None:  r[6] = timelong
    if remark is not None:    r[7] = remark
    return r


def day(offset):
    return (d0 + datetime.timedelta(days=offset)).isoformat()


cases = [
    ('① 同日 + 备注完全相同（司机原样重传）', mk('269901000001'), True),
    ('② 同日 + 备注多几个标点/空格', mk('269901000002', remark='，'.join(str(base[7])) + '  '), True),
    ('③ 次日 + 备注一模一样', mk('269901000003', date=day(1)), True),
    ('④ 同日 + 备注尾部多一句（"已核实"）', mk('269901000004', remark=str(base[7]) + '已核实'), True),
    ('⑤ 日期差 5 天（同一行程不同次）', mk('269901000005', date=day(5)), False),
    ('⑥ 同日 + 备注完全不同（换了地点）', mk('269901000006',
        remark='广州江南市场排队卸货原地打冷3小时，返程东莞仓库打冷2小时'), False),
    ('⑦ 同日 + 时长不同（450 vs 600）', mk('269901000007', timelong=450), False),
    ('⑧ 同日 + 司机不同（换了司机）', mk('269901000008', driver='张三丰'), False),
]

ok = 0
for name, row, expect in cases:
    hits = find_dups(row, index, DUP_DAY_TOL, DUP_REMARK_TH)
    got = bool(hits)
    flag = '✅' if got == expect else '❌'
    if got == expect:
        ok += 1
    detail = ''
    if hits:
        ln, other, dd, s = hits[0]
        detail = f'  命中第{ln}行 {clean_sid(other[0])} 相似度{s:.2f} 日期差{dd}天'
    print(f'{flag} {name}')
    print(f'     期望={"重复" if expect else "不重复"}  实际={"重复" if got else "不重复"}{detail}')

print()
print(f'通过 {ok}/{len(cases)}')

# 额外：确认同类判据不会误伤"同车同线每天跑"的真实记录
print()
print('--- 反向检查：备注相似但日期相隔 >1 天的，不应判重复 ---')
from collections import Counter
fp_count = Counter(dup_fingerprint(r) for _, r in table)
multi = [k for k, v in fp_count.items() if v >= 3]
tested = 0
for ln, r in table:
    if dup_fingerprint(r) not in multi:
        continue
    others = index[dup_fingerprint(r)]
    if len(others) < 2:
        continue
    far = None
    for ln2, r2, d2 in others:
        if ln2 == ln or not d2 or not from_serial(r[2]):
            continue
        if abs((d2 - from_serial(r[2])).days) > 3 and remark_sim(r2[7], r[7]) >= 0.5:
            far = (ln2, r2, abs((d2 - from_serial(r[2])).days))
            break
    if far:
        hits = find_dups(r, index, DUP_DAY_TOL, DUP_REMARK_TH)
        bad = [h for h in hits if h[0] == far[0]]
        print(f'  {"✅" if not bad else "❌"} 第{ln}行 {r[1]} {str(r[2])[:10]} 与 第{far[0]}行 相差{far[2]}天'
              f'（备注相似{remark_sim(far[1][7], r[7]):.2f}）→ 未判重复: {not bad}')
        tested += 1
    if tested >= 5:
        break
if not tested:
    print('  （无合适样本）')

# ---------------------------------------------------------------- 端到端模拟
print()
print('=' * 62)
print('端到端模拟：司机后补上传 → 应被剔除，保留表内已有那条')
print('=' * 62)
from cold_sheet import dedup_plan

# 基准 A：表内真实记录（同一指纹）
sA = list(base)
sA_dup1 = mk('269999000001')                                    # 同日原样重传
sA_dup2 = mk('269999000002', date=day(1), remark=str(base[7]) + ' 已核实')  # 次日+备注加尾
sA_far = mk('269999000003', date=day(20))                       # 同指纹但相隔 20 天 → 真记录
sA_diff = mk('269999000004', remark='广州江南市场排队卸货打冷3小时，东莞仓库再次打冷2小时')

# 基准 B：表内不存在的新指纹，用于验证"批次内重复"
def mknew(sid, plate, date, driver, tl, rmk):
    return [sid, plate, date, driver, '邱钜斌', date + ' 12:00:00', tl, rmk]

b1 = mknew('269999000005', '测试001', '2026-10-01', '测试司机', 123, '测试重复备注甲乙丙')
b2 = mknew('269999000006', '测试001', '2026-10-02', '测试司机', 123, '测试重复备注甲乙丙 ')

rows_sim = [sA_dup1, sA_dup2, sA_far, sA_diff, b1, b2]
n_exist, dup_a, dup_b, accepted, warn = dedup_plan(rows_sim, table)

print(f'输入 {len(rows_sim)} 条 → 编号已存在 {len(n_exist)} | 表内重复 {len(dup_a)} | '
      f'批次内重复 {len(dup_b)} | 需新增 {len(accepted)}')
print()
print('【与表内重复（应剔除，保留表内那条）】')
for r, other, dd, s, ln in dup_a:
    print(f'  ✂ {r[0]} {r[1]} {r[2]} 时长{int(float(r[6]))} 相似度{s:.2f} 日期差{dd}天'
          f'  → 保留 第{ln}行 {clean_sid(other[0])}')
print('【批次内重复（应保留较早编号）】')
for r, other, dd, s in dup_b:
    print(f'  ✂ {r[0]} → 保留 {other[0]}（相似度{s:.2f} 日期差{dd}天）')
print('【真正新增】')
for r in accepted:
    print(f'  ＋ {r[0]} {r[1]} {r[2]} 时长{r[6]} {str(r[7])[:28]}')

exp_dup = {'269999000001', '269999000002'}
exp_acc = {'269999000003', '269999000004', '269999000005'}
got_dup = {d[0][0] for d in dup_a}
got_b = {d[0][0] for d in dup_b}
got_acc = {a[0] for a in accepted}
print()
print(f'  表内重复集合正确: {got_dup == exp_dup}  (期望 {sorted(exp_dup)})')
print(f'  批次内重复 = {sorted(got_b)}  (期望 [269999000006])')
print(f'  新增集合正确: {got_acc == exp_acc}')
print(f'  ★ 总判定: {"PASS" if got_dup == exp_dup and got_b == {"269999000006"} and got_acc == exp_acc else "FAIL"}')
