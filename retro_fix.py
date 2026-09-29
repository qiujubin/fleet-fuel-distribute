# -*- coding: utf-8 -*-
"""同日多"是"跨天补录的自动回溯修正

检测: 模板中某天存在多个"是"，但当前标记为有效是(F非空)的不是当天最后一个"是"
修正: ①旧有效是行降级(F清空) ②当天最后"是"行升级(F=旧锚点, I/J 重写周期累计)
      ③考核表: 就地替换该行  ④汇总: 删掉旧记录对应的行
用法: python retro_fix.py [--check] [--days N]

⚠️ 2026-09-28 用户明令：**禁止用 openpyxl 保存任何文件，只准 win32/COM**。
   本脚本读取改用 python-calamine（只读，不碰文件），F 锚点调整与 I/J 公式重建
   全部改走 WPS COM（com_formula.rebuild_ij）。考核表/汇总本来就是 COM。
"""
import sys, os, glob, datetime
from collections import defaultdict

sys.path.insert(0, r'E:\WorkSpace\workbuddy\2026-09-18-16-38-15')
from python_calamine import CalamineWorkbook
from com_formula import rebuild_ij, tkey, to_serial

TPL = r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板'
KAOHE = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
SUM = r'C:\Users\Jubin\Desktop\源数据\油耗分析汇总.xlsm'
DT_FMT = 'yyyy-mm-dd hh:mm:ss'


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


def ctime(v):
    return str(v)[:19] if v is not None else None


def day_str(v):
    """任意时间表示 -> 'YYYY-MM-DD'（用于日期比较，抗 datetime/序列数漂移）"""
    if v is None: return ''
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(v))).date().isoformat()
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.date().isoformat() if isinstance(v, datetime.datetime) else v.isoformat()
    s = str(v).strip()
    for f in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d'):
        try: return datetime.datetime.strptime(s, f).date().isoformat()
        except ValueError: pass
    try:
        return (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(s))).date().isoformat()
    except ValueError:
        return s[:10]


CHECK_ONLY = '--check' in sys.argv
# 只处理最近 N 天(默认3天)内的记录——历史遗留(2025-08 起约220处)按用户约定不动
DAYS = 3
for i, a in enumerate(sys.argv):
    if a == '--days' and i + 1 < len(sys.argv):
        DAYS = int(sys.argv[i + 1])
CUTOFF = (datetime.date.today() - datetime.timedelta(days=DAYS)).isoformat()

# ---------- 1) 检测（python-calamine 只读，不碰文件） ----------
targets = []   # [(path, old_row, new_row, old_sid, new_sid, fn, F_serial, old_G_time, new_G_time)]
for fp in sorted(glob.glob(os.path.join(TPL, '*.xlsx'))):
    fn = os.path.basename(fp)
    if fn.startswith(('~$', '模板')): continue
    try:
        cwb = CalamineWorkbook.from_path(fp)
        names = cwb.sheet_names
        if '加油数据' not in names: continue
        tname = '计算模板' if '计算模板' in names else ('油耗计算' if '油耗计算' in names else None)
        if tname is None: continue
        drs = cwb.get_sheet_by_name('加油数据').to_python(skip_empty_area=False)
        trs = cwb.get_sheet_by_name(tname).to_python(skip_empty_area=False)
        full_map, sid_of = {}, {}
        for row in drs[1:]:
            if len(row) > 12 and row[0] not in (None, '') and row[12] not in (None, ''):
                k = tkey(row[12])
                if k is None: continue
                full_map[k] = str(row[8] or '').strip()
                sid_of[k] = str(row[0]).strip().replace('.0', '')
        rows = []
        for i, row in enumerate(trs, start=1):
            if i < 2: continue
            g = row[6] if len(row) > 6 else None
            f = row[5] if len(row) > 5 else None
            if g not in (None, ''):
                rows.append((i, g, f, full_map.get(tkey(g), '')))
    except Exception as e:
        print(f'  跳过 {fn}: {e}'); continue

    shi_by_day = defaultdict(list)
    for r, g, f, full in rows:
        if full == '是': shi_by_day[day_str(g)].append((r, g))
    valid = [(r, g, f) for r, g, f, full in rows if f not in (None, '')]
    for r, g, f in valid:
        day = day_str(g)
        if day < CUTOFF: continue        # 历史欠账不处理
        items = shi_by_day.get(day, [])
        if len(items) < 2: continue
        last_shi = max(items, key=lambda x: tkey(x[1]) or 0)
        if last_shi[0] != r:
            old_k, new_k = tkey(g), tkey(last_shi[1])
            targets.append((fp, r, last_shi[0], sid_of.get(old_k, '?'), sid_of.get(new_k, '?'),
                            fn, float(to_serial(f)), float(to_serial(g)), float(to_serial(last_shi[1]))))

print(f'=== 检测结果(仅最近 {DAYS} 天, 自 {CUTOFF}): {len(targets)} 处需要回溯修正')
for fp, r, r2, s1, s2, fn, f, g, g2 in targets:
    print(f'  {fn}: 行{r}({s1} G={ctime(g)}) 应降级 -> 行{r2}({s2} G={ctime(g2)}) 升为有效是')

if not targets:
    print('无需修正')
    sys.exit(0)
if CHECK_ONLY:
    sys.exit(0)

# ---------- 2) 修正（全部走 COM） ----------
import pythoncom
import win32com.client as wc
pythoncom.CoInitialize()
app = None
try:
    app = wc.DispatchEx('Ket.Application')
    app.Visible = False; app.DisplayAlerts = False
    try: app.ScreenUpdating = False
    except Exception: pass

    def _num(x):
        try: return float(x)
        except (TypeError, ValueError): return None

    for fp, old_r, new_r, old_sid, new_sid, fn, F_serial, old_G, new_G in targets:
        w1 = app.Workbooks.Open(fp, 0, False)
        if w1.ReadOnly:
            print(f'  !! {fn} 被占用(只读), 跳过'); continue
        ws1 = None
        for sn in ('计算模板', '油耗计算'):
            try: ws1 = w1.Worksheets(sn); break
            except Exception: continue
        if ws1 is None:
            print(f'  !! {fn} 无计算模板 sheet'); w1.Close(False); continue

        # 2.1 改 F 锚点（COM，写序列数，避免 -8h 时区坑）
        ws1.Cells(old_r, 6).ClearContents()
        ws1.Cells(new_r, 6).Value2 = F_serial
        ws1.Cells(new_r, 6).NumberFormat = DT_FMT
        # 级联: 其他有效是行的 F 若指向"被降级的旧行", 改指"新的有效是行"
        cnt = 0
        lastf = true_last_row(ws1, 7)
        if lastf >= 2:
            vals = ws1.Range(ws1.Cells(2, 6), ws1.Cells(lastf, 6)).Value2
            if vals is not None:
                if not isinstance(vals, tuple): vals = ((vals,),)
                for i, cv in enumerate(vals):
                    v = cv[0] if isinstance(cv, tuple) else cv
                    r2 = 2 + i
                    if r2 == new_r: continue
                    if isinstance(v, (int, float)) and abs(float(v) - old_G) < 1e-6:
                        ws1.Cells(r2, 6).Value2 = new_G
                        ws1.Cells(r2, 6).NumberFormat = DT_FMT
                        cnt += 1
        print(f'  级联: {"行 " + str(cnt) if cnt else "无其他行引用该锚点"}')

        # 2.2 用 COM 重写 I/J 周期累计公式（不用 openpyxl）
        n = rebuild_ij(app, fp, wb=w1)
        print(f'{fn}: F 调整完成, 重写 {n} 行 I/J 公式')

        # 2.3 重算并读回新行值
        try: app.Calculate()
        except Exception: pass
        rng = ws1.Range(ws1.Cells(new_r, 6), ws1.Cells(new_r, 22)).Value2
        v = rng[0] if isinstance(rng[0], tuple) else rng
        w1.Save(); w1.Close(False)
        print(f'  新行 {new_r} ({new_sid}): I={_num(v[3])} J={_num(v[4])} P={_num(v[10])}')

        # 2.4 考核表: 就地替换（old_sid 行 -> new_sid）
        wbk = app.Workbooks.Open(KAOHE, 0, False)
        wsd2 = wbk.Worksheets('每日油耗数据')
        old_row_kao = None
        for i in range(2, 6000):
            x = wsd2.Cells(i, 1).Value2
            if x is not None and str(x).strip().rstrip('.0') == old_sid:
                old_row_kao = i; break
        new_rowvals = [new_sid, None, None, None, None, v[0], v[1], '是', _num(v[3]), _num(v[4]), None,
                       _num(v[6]), _num(v[7]), _num(v[8]), _num(v[9]), _num(v[10]), _num(v[11]),
                       _num(v[12]), _num(v[13]), _num(v[14]), _num(v[15]), _num(v[16]),
                       float(int(v[1])), 0]
        if old_row_kao:
            for c in (3, 4, 5, 11):
                new_rowvals[c - 1] = wsd2.Cells(old_row_kao, c).Value2
            new_rowvals[1] = wsd2.Cells(old_row_kao, 2).Value2
            wsd2.Cells(old_row_kao, 1).NumberFormat = '@'
            wsd2.Range(wsd2.Cells(old_row_kao, 1), wsd2.Cells(old_row_kao, 24)).Value = tuple([tuple(new_rowvals)])
            for c in (6, 7): wsd2.Cells(old_row_kao, c).NumberFormat = DT_FMT
            wsd2.Cells(old_row_kao, 23).NumberFormat = 'yyyy-mm-dd'
            print(f'  考核表行{old_row_kao}: {old_sid} -> {new_sid}')
        else:
            last = true_last_row(wsd2, 1)
            nr = max(last, 1) + 1
            new_rowvals[1] = nr - 1
            wsd2.Cells(nr, 1).NumberFormat = '@'
            wsd2.Range(wsd2.Cells(nr, 1), wsd2.Cells(nr, 24)).Value = tuple([tuple(new_rowvals)])
            print(f'  考核表追加行{nr}: {new_sid}')
        wbk.Save(); wbk.Close(False)

        # 2.5 汇总: 删旧行（车牌 + 时间A + 时间B 匹配）
        wbs = app.Workbooks.Open(SUM, 0, False)
        wss = wbs.Worksheets('Sheet1')
        last = true_last_row(wss, 3)
        plate = os.path.splitext(fn)[0].split('-')[-1].strip()
        fA, fB = _num(F_serial), _num(old_G)   # 匹配被降级的旧行(时间B=旧有效是的时间)
        print(f'  汇总匹配键: 车牌={plate} 时间A={fA} 时间B={fB}')
        del_row = None
        for i in range(2, last + 1):
            if str(wss.Cells(i, 3).Value2).strip() != plate: continue
            ta, tb = wss.Cells(i, 6).Value2, wss.Cells(i, 7).Value2
            try: ta_f, tb_f = float(ta), float(tb)
            except (TypeError, ValueError): continue
            if fA is not None and abs(ta_f - fA) < 0.0001 and abs(tb_f - fB) < 0.0001:
                del_row = i; break
        if del_row:
            wss.Rows(del_row).Delete()
            last2 = true_last_row(wss, 3)
            for i in range(del_row, last2 + 1):
                wss.Cells(i, 1).Value2 = i - 1
            print(f'  汇总: 删除旧行{del_row}并重排序号')
        wbs.Save(); wbs.Close(False)
    print(f'\n回溯修正完成: {len(targets)} 处')
finally:
    if app is not None:
        try: app.Quit()
        except Exception: pass
    pythoncom.CoUninitialize()
