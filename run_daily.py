# -*- coding: utf-8 -*-
"""每日一键总跑（推荐入口）

流程：
  ⓪ fetch_api.py        从数据接口取数，生成当天的「车辆加油_管理(日期).xlsx」
  ① fleet_fuel_com.py   分发 → 考核写入 → 每日加油记录 → 透视表 → 宏同步
  ② retro_fix.py        回溯修正检测（同日多"是"跨天补录：模板降级/升级 + 考核行替换 + 汇总删旧行）
  ③ 若②有修正 → 再跑一次 fleet_fuel_com.py 做二次同步（刷新透视表 + 宏补新行）
  ④ cold_sheet.py       把接口新增的「打冷记录」追加进考核表「打冷记录」sheet
                        两层去重：① 记录编号已存在 → 跳过
                                  ② 内容重复（车牌+司机+时长相同、日期相差≤1天、备注相似≥0.5）
                                     → 判定为「司机后补上传」，跳过并保留表内已有那条
  ⑤ sort_fuel_sheets.py 收尾：4 张油耗页面「大排序」；当天第一次会先「速度图片清零」
  ⑥ cleanup_backups.py  清理过期备份：各类 backup_* 只保留最近 3 天

用法：python run_daily.py                 # ⓪ 自动取数（推荐）
      python run_daily.py <导出文件路径>   # 跳过 ⓪，用指定/手里已有的导出文件
      python run_daily.py --no-fetch      # 跳过 ⓪，用「数据」目录里最新的一份
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


def kill_orphan_wps():
    """清理无主窗口的 WPS/ET 残留实例。

    每次 COM 跑完多少会留下几个后台实例，累积几天后会让「考核表被占用」，
    导致整条流程在预检阶段就被拦下。这里只杀 MainWindowHandle=0 的，
    用户正在编辑的窗口（有标题/HWND）一律不动。返回清理个数。
    """
    ps = ("$p = @(Get-Process wps,et -ErrorAction SilentlyContinue | "
          "Where-Object { $_.MainWindowHandle -eq 0 }); "
          "if ($p.Count -gt 0) { $p | Stop-Process -Force -ErrorAction SilentlyContinue }; "
          "$p.Count")
    try:
        r = subprocess.run(['powershell.exe', '-NoProfile', '-NonInteractive', '-Command', ps],
                           capture_output=True, text=True, timeout=30)
        tail = (r.stdout or '').strip().splitlines()
        return int(tail[-1]) if tail and tail[-1].strip().isdigit() else 0
    except Exception:
        return 0


def main():
    argv = [a for a in sys.argv[1:] if not a.startswith('--')]
    export = argv[0] if argv else None
    no_fetch = '--no-fetch' in sys.argv

    n_orphan = kill_orphan_wps()
    if n_orphan:
        print(f'[预清理] 关闭 {n_orphan} 个无窗口的 WPS 残留实例（不影响正在编辑的窗口）')

    # ⓪ 取数：从接口生成当天的加油表格（未指定导出文件时才跑）
    if export or no_fetch:
        print('=' * 24, '⓪ 跳过取数（使用已有导出文件）', '=' * 24)
    else:
        print('=' * 24, '⓪ 从接口取数', '=' * 24)
        out0, rc0 = run('fetch_api.py')
        if rc0 != 0 or '已生成' not in out0:
            print('\n!!! 取数失败，为安全起见中止（未做任何写入）')
            return rc0 or 1

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

    # ④ 打冷记录：编号去重 + 内容去重后追加进考核表「打冷记录」sheet
    print('=' * 24, '④ 打冷记录同步', '=' * 24)
    run('cold_sheet.py', ['--days', '10'])

    # ⑤ 收尾：4 张油耗页面「大排序」（当天第一次会先「速度图片清零」）
    print('=' * 24, '⑤ 油耗页面排序', '=' * 24)
    run('sort_fuel_sheets.py')

    # ⑥ 备份清理：各类 backup_* 只保留最近 3 天（backup_daily 的 latest 不动）
    print('=' * 24, '⑥ 清理过期备份', '=' * 24)
    run('cleanup_backups.py', ['--apply'])

    print('\n完成：全流程结束')
    return rc1


if __name__ == '__main__':
    sys.exit(main())
