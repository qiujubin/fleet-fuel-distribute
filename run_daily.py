# -*- coding: utf-8 -*-
"""每日一键总跑（推荐入口）

流程：
  ① fleet_fuel_com.py   分发 → 考核写入 → 每日加油记录 → 透视表 → 宏同步
  ② retro_fix.py        回溯修正检测（同日多"是"跨天补录：模板降级/升级 + 考核行替换 + 汇总删旧行）
  ③ 若②有修正 → 再跑一次 fleet_fuel_com.py 做二次同步（刷新透视表 + 宏补新行）
  ④ sort_fuel_sheets.py 收尾：4 张油耗页面「大排序」；当天第一次会先「速度图片清零」

用法：python run_daily.py [导出文件路径]     # 省略路径则自动取「数据」文件夹里最新导出
"""
import os, re, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable
ENV = dict(os.environ, PYTHONIOENCODING='utf-8')


def run(script, args=()):
    cmd = [PY, '-u', os.path.join(HERE, script), *args]
    print(f'\n>>> {script} {" ".join(args)}', flush=True)
    r = subprocess.run(cmd, capture_output=True, text=True,
                       encoding='utf-8', errors='replace', env=ENV)
    out = (r.stdout or '') + (r.stderr or '')
    print(out, flush=True)
    if r.returncode != 0:
        print(f'!!! {script} 退出码 {r.returncode}', flush=True)
    return out, r.returncode


def main():
    export = sys.argv[1] if len(sys.argv) > 1 else None

    print('=' * 24, '① 主链路', '=' * 24)
    out1, rc1 = run('fleet_fuel_com.py', ([export] if export else []))

    print('=' * 24, '② 回溯修正检测', '=' * 24)
    out2, _ = run('retro_fix.py', ['--days', '3'])
    m = re.search(r'检测结果[^:]*:\s*(\d+)\s*处需要回溯修正', out2)
    n_fix = int(m.group(1)) if m else 0

    if n_fix > 0:
        print('=' * 24, f'③ 检测到 {n_fix} 处回溯修正，二次同步', '=' * 24)
        run('fleet_fuel_com.py', ([export] if export else []))
        print(f'主链路 + {n_fix} 处回溯修正 + 二次同步 完成')
    else:
        print('无需回溯修正')

    # ④ 收尾：4 张油耗页面「大排序」（当天第一次会先「速度图片清零」）
    print('=' * 24, '④ 油耗页面排序', '=' * 24)
    run('sort_fuel_sheets.py')

    print('\n完成：全流程结束')
    return rc1


if __name__ == '__main__':
    sys.exit(main())
