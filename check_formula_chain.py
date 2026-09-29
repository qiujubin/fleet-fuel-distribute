# -*- coding: utf-8 -*-
"""全车队「计算模板」公式链体检（只读，全部走 win32/COM）

对每个"有效是"行 r 检查 I 列公式的累加范围是否等于应有的周期范围:
    应为  =XLOOKUP(D&G,...) + I{r-1} + ... + I{anchor+1}
    anchor = G 列中值等于本行 F 的那一行（上一个有效是行）

分类报警:
  【恶性】公式里出现了周期外的行 —— rebuild 锚点匹配失败的特征
          （I/J 会变成整表总和，数值看着像数字但荒谬大，并向下游考核/汇总传播）
          修法: com_formula.rebuild_ij(app, fp)
  【历史】累加项不足 —— 同日多"是"的历史欠账（2025-08 起约 220 处，按用户约定不动）

⚠️ 2026-09-28 用户明令禁止 openpyxl：本脚本读取用 python-calamine，
   公式文本用 WPS COM 批量读（`Range.Formula`），不打开 openpyxl。

用法: python check_formula_chain.py          # 全车队
      python check_formula_chain.py 文件名    # 单个文件
"""
import os, re, sys, glob, atexit

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from com_formula import read_plan

TPL = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板'


def tpl_sheet(wb):
    for sn in ('计算模板', '油耗计算'):
        try:
            return wb.Worksheets(sn)
        except Exception:
            continue
    return None


def check_file(app, fp):
    """返回 (major, minor) —— major=引用了周期外的行, minor=累加项不足"""
    plan = read_plan(fp)                     # [(row, expected_off_rows)]
    if not plan:
        return [], []
    wb = app.Workbooks.Open(fp, 0, True)     # 只读
    try:
        ws = tpl_sheet(wb)
        if ws is None:
            return [], []
        last = max(r for r, _ in plan)
        fml = ws.Range(ws.Cells(2, 9), ws.Cells(last, 9)).Formula
        if not isinstance(fml, tuple):
            fml = ((fml,),)
        major, minor = [], []
        for r, exp in plan:
            cell = fml[r - 2] if r - 2 < len(fml) else None
            t = cell[0] if isinstance(cell, tuple) else cell
            if not isinstance(t, str): continue
            nums = [int(x) for x in re.findall(r'\+I(\d+)', t)]
            extra = sorted(set(nums) - set(exp))
            if extra:
                major.append((os.path.basename(fp), r, len(nums), len(exp),
                              f'多出周期外行 {extra[:5]}' + ('...' if len(extra) > 5 else '')))
            elif len(nums) < len(exp):
                minor.append((os.path.basename(fp), r, len(nums), len(exp)))
        return major, minor
    finally:
        try: wb.Close(False)
        except Exception: pass


def main():
    import pythoncom
    import win32com.client as wc
    pythoncom.CoInitialize()
    app = None
    try:
        for pid in ('Ket.Application', 'et.Application', 'Excel.Application'):
            try:
                app = wc.DispatchEx(pid); print('COM 应用:', pid); break
            except Exception:
                continue
        if app is None:
            raise SystemExit('无法启动 COM')
        app.Visible = False
        app.DisplayAlerts = False
        try: app.ScreenUpdating = False
        except Exception: pass

        if len(sys.argv) > 1:
            files = [os.path.join(TPL, sys.argv[1])]
        else:
            files = sorted(f for f in glob.glob(os.path.join(TPL, '*.xlsx'))
                           if not os.path.basename(f).startswith('~$'))
        total, nmajor, nminor = 0, 0, 0
        minor_files = {}
        for fp in files:
            try:
                major, minor = check_file(app, fp)
            except Exception as e:
                print(f'!! {os.path.basename(fp)}: {e}'); continue
            total += 1
            for b in major:
                nmajor += 1
                print('  [恶性] %s 行%d  累加%d项(期望%d)  %s' % b)
            for b in minor:
                nminor += 1
                minor_files.setdefault(b[0], []).append(b[1])
        print(f'\n扫描 {total} 个文件')
        print(f'  【恶性】引用周期外行: {nmajor} 处' + ('  -> 用 com_formula.rebuild_ij 修复' if nmajor else '  ✓'))
        print(f'  【历史】累加项不足: {nminor} 处' + (f'  涉及 {len(minor_files)} 个文件(按约定不动)' if nminor else ''))
        for fn, rows in minor_files.items():
            print(f'      {fn}: {len(rows)} 行 (如 {rows[:6]})')
        return 1 if nmajor else 0
    finally:
        if app is not None:
            try: app.Quit()
            except Exception: pass
        pythoncom.CoUninitialize()


if __name__ == '__main__':
    sys.exit(main())
