# -*- coding: utf-8 -*-
import os, shutil
p = 'README.md'
s = open(p, encoding='utf-8').read()
fixes = [
    ('把每天从系统导出的加油名单，自动分发到 50+ 辆车的油耗计算模板，再汇总到油耗考核表与油耗分析汇总。\n原来手工做一遍要个把小时，现在一次运行 **10~30 秒**。',
     '把每天从系统导出的加油名单，自动分发到 49 辆车（+15 个尿素文件）的油耗计算模板，再一路汇总到油耗考核表与油耗分析汇总。\n原来全手工逐车复制粘贴，现在一次运行 **10~30 秒**。'),
    ('| 车辆油耗计算模板/                      ← 每辆车一个文件（约 60 个）',
     '| 车辆油耗计算模板/                      ← 共 64 个文件（油车 49 + 尿素 15）'),
    ('  └─ 尿素xx-车牌.xlsx                  ← 尿素文件（只有 16 辆车）',
     '  └─ 尿素xx-车牌.xlsx                  ← 尿素文件（15 辆车）'),
    ('| **region_map.py** | 45 辆车的区域映射（干线/华北/华东/华南）。后台调宏跳过区域填充时，由主脚本用它补做 | 被 import |',
     '| **region_map.py** | 46 辆车的区域映射（干线 20 / 华北 2 / 华南 18 / 华东 6）。后台调宏跳过区域填充时，由主脚本用它补做 | 被 import |'),
    ('**你要看的**：报告里的 4 个部分——新增明细、有效"是"行、排除词记录、数据质量提醒（异常只在报告里说，不写进文件）。发现异常去系统里核对原始数据。',
     '**你要看的**：报告分节——新增明细、有效"是"行、排除词记录、数据质量提醒（异常只在报告里说，从不写进文件）。发现异常去系统里核对原始数据。'),
]
n = 0
for old, new in fixes:
    if old in s:
        s = s.replace(old, new); n += 1
    else:
        print('未匹配:', old[:36])

skill_path = 'C:' + chr(92) + r'Users\Jubin\.workbuddy\skills\fleet-fuel-daily-distribute\SKILL.md'
tail = '| 09-22 | 排除词记录改为独立进每日加油记录；报告改用车牌号；导出误删恢复；脚本归档 + 本文档 |'
note = (tail
        + '\n\n---\n\n> **维护约定**：新增/改动脚本时，请同步更新本文档第 4 节的脚本清单；一次性脚本用完请移入 `archive/`。'
        + '\n> 详细业务规则与踩坑记录见 AI 操作手册：`' + skill_path + '`（与本文件互补：这份给人看，那份给 AI 执行）')
if tail in s and '维护约定' not in s:
    s = s.replace(tail, note); n += 1
open(p, 'w', encoding='utf-8').write(s)
print(f'README 修正 {n} 处')

# region_LIST_*.txt 是提取区域映射的中间产物, 归档
os.makedirs('archive', exist_ok=True)
for f in ['region_LIST_TRUNK.txt', 'region_LIST_NORTH.txt', 'region_LIST_SOUTH.txt', 'region_LIST_EAST.txt']:
    if os.path.exists(f):
        shutil.move(f, os.path.join('archive', f))
print('region_LIST_*.txt 已归档')
