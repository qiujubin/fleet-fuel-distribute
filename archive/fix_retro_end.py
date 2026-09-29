# -*- coding: utf-8 -*-
"""修 retro_fix.py: 3 处 WPS 不可靠的 End(xlUp) -> 确定性倒扫 true_last_row"""
p = 'retro_fix.py'
s = open(p, encoding='utf-8').read()

# 1) 插入工具函数(放在 import 之后)
if 'def true_last_row' not in s:
    anchor = 'SUM = r' + "'" + r'C:\Users\Jubin\Desktop\源数据\油耗分析汇总.xlsm' + "'"
    helper = '''

def true_last_row(ws, col=1):
    """WPS 的 End(xlUp) 会返回错误行号 -> 用 UsedRange 尾部倒扫(与主脚本一致)"""
    ur = ws.UsedRange
    last_used = ur.Row + ur.Rows.Count - 1
    if last_used < 1:
        return 1
    vals = ws.Range(ws.Cells(1, col), ws.Cells(last_used, col)).Value2
    if vals is None:
        return 1
    if not isinstance(vals, tuple):
        vals = ((vals,),)
    for i in range(len(vals) - 1, -1, -1):
        row = vals[i]
        v = row[0] if isinstance(row, tuple) else row
        if v not in (None, ''):
            return 1 + i
    return 1
'''
    assert anchor in s
    s = s.replace(anchor, anchor + helper, 1)
    print('已插入 true_last_row')

# 2) 替换 3 处 End(xlUp)
subs = [
    ('            last = wsd2.Cells(wsd2.Rows.Count, 1).End(-4162).Row',
     '            last = true_last_row(wsd2, 1)'),
    ('        last = wss.Cells(wss.Rows.Count, 3).End(-4162).Row',
     '        last = true_last_row(wss, 3)'),
    ('            last2 = wss.Cells(wss.Rows.Count, 3).End(-4162).Row',
     '            last2 = true_last_row(wss, 3)'),
]
n = 0
for old, new in subs:
    if old in s:
        s = s.replace(old, new); n += 1
    else:
        print('未匹配:', old.strip()[:50])
open(p, 'w', encoding='utf-8').write(s)
print(f'替换 {n}/3 处 End(xlUp)')
print('剩余 End(-4162):', s.count('End(-4162)'))
