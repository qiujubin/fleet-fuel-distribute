# -*- coding: utf-8 -*-
"""更新 README.md 与 SKILL.md：新增 run_daily.py / del_summary_row.py，记录 09-24 变更"""
import os

# ---------- README ----------
p = 'README.md'
s = open(p, encoding='utf-8').read()
n = 0

old = """**程序会做**：分发 → 写考核 → 刷每日加油记录 → 刷新透视表 → 调宏同步汇总 → 出报告。"""
new = """**一键入口（推荐）**：`python run_daily.py` —— 内部自动串起三段：主链路（分发→考核→每日加油记录→透视表→宏同步）→ 回溯修正检测 → 必要时二次同步。不用再手工分步跑。
**程序会做**：分发 → 写考核 → 刷每日加油记录 → 刷新透视表 → 调宏同步汇总 → 回溯检查 → 出报告。"""
if old in s: s = s.replace(old, new); n += 1

old = "| **fleet_fuel_com.py** | 主脚本。"
new = ("| **run_daily.py** | **每日一键入口**：串起主链路 → 回溯修正 → 二次同步 | `python run_daily.py`（可带导出路径，省略则取最新） |\n"
       "| **fleet_fuel_com.py** | 主脚本。")
if old in s: s = s.replace(old, new, 1); n += 1

old = "| **test_macro_v2.py** | 宏的隔离测试（跑两次验证去重） | 改宏之后验证 |"
new = (old + "\n| **del_summary_row.py** | 删除汇总里指定车牌+时间B 的行（回溯修正后的清理用） | `python del_summary_row.py`（改脚本内 PLATE/DAY_B） |")
if old in s: s = s.replace(old, new); n += 1

tail = '| 09-23 | 修复报告噪音（类型/单价矛盾提醒只对新增记录报）；首次处理「现金加油」类型 |'
if tail in s and '| 09-24 |' not in s:
    s = s.replace(tail, tail + '\n| 09-24 | 新增 `run_daily.py` 一键入口（把回溯修正纳入流程）；修复 retro_fix 的 `End(xlUp)` 不可靠导致汇总旧行漏删 |')
    n += 1
open(p, 'w', encoding='utf-8').write(s)
print(f'README 更新 {n} 处')

# ---------- SKILL ----------
p2 = r'C:\Users\Jubin\.workbuddy\skills\fleet-fuel-daily-distribute\SKILL.md'
s2 = open(p2, encoding='utf-8').read()
n2 = 0
anchor = '## 文档维护约定（重要）'
if 'run_daily.py' not in s2 and anchor in s2:
    add = """## 每日执行入口（推荐）

**直接跑 `python run_daily.py`**（工作区根目录），它会自动完成：
1. `fleet_fuel_com.py` —— 主链路；
2. `retro_fix.py --days 3` —— 回溯修正检测；
3. 若第 2 步检测到修正 → 再跑一次主链路做**二次同步**（刷新透视表 + 宏补新行）。

**为什么需要二次同步**：回溯修正会改写「每日油耗数据」里的考核行，但透视表与 4 张油耗 sheet 的公式缓存、以及汇总里的行都还是修正前的 → 必须再刷一次并重跑宏，否则汇总会留着旧行、缺新行（2026-09-24 实战暴露）。

## 文档维护约定（重要）"""
    s2 = s2.replace(anchor, add, 1); n2 += 1

old2 = '- `normalize_com_times()`（主脚本内置）：每次处理车辆文件时把 M/C/F/G 列字符串时间转序列数。'
new2 = (old2 + '\n- `del_summary_row.py`：删除汇总里指定「车牌 + 时间B」的行（回溯修正后的清理；改脚本内 PLATE/DAY_B 后运行，支持 `--check` 先看）。')
if old2 in s2: s2 = s2.replace(old2, new2); n2 += 1

# 铁律: End(xlUp) 不可靠 -> 已扩到所有脚本
old3 = '3. **WPS 的 `End(xlUp)` 返回错误行号**'
if old3 in s2:
    pass
if '`End(xlUp)` 返回错误行号' in s2 and '所有脚本' not in s2:
    s2 = s2.replace('（A12 有值返回 11）', '（A12 有值返回 11；**retro_fix.py 2026-09-24 因此漏删汇总旧行，已全部改用 true_last_row 倒扫**）', 1)
    n2 += 1
open(p2, 'w', encoding='utf-8').write(s2)
print(f'SKILL 更新 {n2} 处')
