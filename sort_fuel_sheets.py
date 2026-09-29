# -*- coding: utf-8 -*-
"""每日收尾：4 张油耗页面「大排序」（可选先「速度图片清零」）

考核表 JSA 宏（按钮 → 函数名已核实）:
  大排序(全部)   -> 包括速度根据达标率排序   区域 B6:T, N列(达标率)升序 + E列降序
  速度图片清零   -> 速度清零                 R7:S{lastRow} 置 0, T 列清空
  小排序(大部分) -> 根据达标率排序           区域 B6:Q（本脚本不用）

规则（用户 2026-09-28 指定）:
  - 当天第一次执行  -> 先「速度图片清零」，再「大排序」
  - 当天已执行过    -> 只「大排序」

「当天第一次」的判定：state/last_sort.json 里记录的日期是否等于今天。
中途失败不会写状态，下次仍按"第一次"处理（安全侧）。

用法:
  python sort_fuel_sheets.py            # 正常执行
  python sort_fuel_sheets.py --check    # 只看今天是否第一次、会执行哪些宏
  python sort_fuel_sheets.py --force-first  # 强制按"当天第一次"跑（含清零）
"""
import os, sys, glob, json, shutil, atexit, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
KAOHE = os.environ.get('FUEL_KAOHE', r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm')
SHEETS = ['干线油耗', '华北油耗', '华东油耗', '华南油耗']
MACRO_CLEAR = '速度清零'                  # 按钮「速度图片清零」
MACRO_SORT = '包括速度根据达标率排序'      # 按钮「大排序(全部)」
STATE_DIR = os.path.join(HERE, 'state')
STATE_FP = os.path.join(STATE_DIR, 'last_sort.json')

CHECK = '--check' in sys.argv
FORCE_FIRST = '--force-first' in sys.argv
RECORD = '--no-record' not in sys.argv


def load_state():
    try:
        with open(STATE_FP, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def save_state(d):
    os.makedirs(STATE_DIR, exist_ok=True)
    with open(STATE_FP, 'w', encoding='utf-8') as f:
        json.dump(d, f, ensure_ascii=False, indent=2)


def main():
    today = datetime.date.today().isoformat()
    st = load_state()
    first = FORCE_FIRST or st.get('last_sort_date') != today

    print(f'日期={today}  上次排序={st.get("last_sort_date") or "(无记录)"}')
    print(f'当天第一次: {"是 -> 速度清零 + 大排序" if first else "否 -> 只大排序"}')
    print(f'目标页面: {" / ".join(SHEETS)}')
    if CHECK:
        print('(体检模式, 不执行)')
        return 0

    import pythoncom
    import win32com.client as wc

    pythoncom.CoInitialize()
    app = None
    opened = []

    def cleanup():
        global app
        try:
            if app is None: return
            for wb in list(opened):
                try: wb.Close(False)
                except Exception: pass
            opened.clear()
            try: app.Quit()
            except Exception: pass
            app = None
        except Exception:
            pass

    atexit.register(cleanup)

    try:
        for pid in ('Ket.Application', 'et.Application', 'Excel.Application'):
            try:
                app = wc.DispatchEx(pid); print('COM 应用:', pid); break
            except Exception:
                continue
        if app is None:
            raise RuntimeError('无法启动 WPS/Excel COM')
        app.Visible = False
        app.DisplayAlerts = False
        try: app.ScreenUpdating = False
        except Exception: pass

        wbk = app.Workbooks.Open(KAOHE, 0, False)
        if wbk.ReadOnly:
            raise RuntimeError('考核表被占用（只读打开），请先关闭 WPS 里的考核表')
        opened.append(wbk)

        # 清零会物理删除 T 列的嵌入图片（WPS cellimages + media/*.png），
        # 文件内无法恢复 —— 所以第一次执行前先整包备份（保留最近 3 份）
        if first and RECORD:
            stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            bak_dir = os.path.join(HERE, f'backup_sort_{stamp}')
            os.makedirs(bak_dir, exist_ok=True)
            shutil.copy2(KAOHE, os.path.join(bak_dir, os.path.basename(KAOHE)))
            print(f'清零前整包备份 -> {bak_dir}')
            olds = sorted(d for d in glob.glob(os.path.join(HERE, 'backup_sort_*'))
                          if os.path.isdir(d))
            for d in olds[:-3]:
                for root, dirs, files in os.walk(d, topdown=False):
                    for f in files:
                        try: os.remove(os.path.join(root, f))
                        except Exception: pass
                    for dd in dirs:
                        try: os.rmdir(os.path.join(root, dd))
                        except Exception: pass
                try: os.rmdir(d)
                except Exception: pass

        done = []
        for sn in SHEETS:
            try:
                ws = wbk.Worksheets(sn)
            except Exception as e:
                print(f'  !! 找不到工作表 {sn}: {e}')
                continue
            ws.Activate()
            try:
                cur = wbk.ActiveSheet.Name
            except Exception:
                cur = '?'
            if cur != sn:
                print(f'  !! 激活后 ActiveSheet={cur}（期望 {sn}），跳过以免排错表')
                continue
            if first:
                try:
                    app.Run(MACRO_CLEAR)
                    cl = '清零+'
                except Exception as e:
                    cl = f'清零失败({e})+'
                    print(f'  {sn}: 速度清零失败: {e}')
            else:
                cl = ''
            try:
                app.Run(MACRO_SORT)
                print(f'  {sn}: {cl}大排序 完成')
                done.append(sn)
            except Exception as e:
                print(f'  {sn}: 大排序失败: {e}')

        if done:
            wbk.Save()
            print(f'考核表已保存（{len(done)}/{len(SHEETS)} 张完成）')
        else:
            print('无页面执行成功，不保存')

        wbk.Close(False)
        opened.clear()

        if done and RECORD:
            st['last_sort_date'] = today
            st['last_sort_at'] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            st['last_sort_sheets'] = done
            save_state(st)
            print(f'状态已记录 -> {STATE_FP}')
        return 0
    finally:
        cleanup()
        try: pythoncom.CoUninitialize()
        except Exception: pass


if __name__ == '__main__':
    sys.exit(main())
