# -*- coding: utf-8 -*-
p = r'C:\Users\Jubin\.workbuddy\skills\fleet-fuel-daily-distribute\SKILL.md'
s = open(p, encoding='utf-8').read()
old = """## 辅助脚本

- ：同日多"是"跨天补录的自动回溯修正（模板降级/升级 + 考核行替换 + 汇总删旧行）。 只检测， 限最近 N 天（默认3，历史约220处按约定不动）。主脚本跑完后建议执行一次 。
- （主脚本内置）：每次处理车辆文件时把 M/C/F/G 列字符串时间转序列数。"""
new = """## 辅助脚本

- `retro_fix.py`：同日多"是"跨天补录的自动回溯修正（模板降级/升级 + 考核行替换 + 汇总删旧行）。`--check` 只检测，`--days N` 限制最近 N 天（默认 3；历史自 2025-08 起约 220 处同类情况，按用户约定不动）。主脚本跑完后建议执行一次 `retro_fix.py --days 3`。
- `normalize_com_times()`（主脚本内置）：每次处理车辆文件时把 M/C/F/G 列字符串时间转序列数。
- `fix_d14_448.py` / `fix_d17_466.py` / `del_summary_453.py`：历史回溯修正的一次性脚本（模式参考）。"""
assert old in s, 'anchor not found'
s = s.replace(old, new)
open(p, 'w', encoding='utf-8').write(s)
print('SKILL.md 修正完成')
