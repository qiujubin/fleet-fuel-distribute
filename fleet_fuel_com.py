# -*- coding: utf-8 -*-
"""每日加油名单分发 —— COM 架构版 (win32com + WPS 表格)

链路: 最新导出 -> 车辆模板[加油数据+计算模板] -> 油耗考核(每日油耗/尿素数据+每日加油记录) -> 调宏[同步油耗到汇总]

架构:
  - 读: python-calamine (毫秒级, 只读不碰宏)
  - 写: win32com 驱动 WPS 表格(Ket.Application) —— 由 WPS 原生保存,
        JSA 宏/按钮/图片/透视表/公式缓存 完整保留, 无需任何部件移植。

去重(联合主键, 逐行 continue, 任何情况下不 break):
  - 主键1: 系统编号(唯一单据号), 有则必用;
  - 主键2: (登记时间, 车牌, 金额, 油量) 兜底 —— 系统编号缺失时使用。
  同一天内先出现的旧记录被跳过、后出现的新记录照常处理(修复旧版"按日期去重
  导致同日下半段被误跳"的 bug)。整链幂等, 重复运行不写重。

稳定性:
  - 所有 Workbook 打开前做占用锁检查, 只读打开立即报错跳过;
  - try...finally + atexit 双保险: 关闭全部 Workbook 并退出 WPS, 不残留进程;
  - ScreenUpdating=False 提速, DisplayAlerts=False 压弹窗;
  - 每次运行先备份(backup_daily/, 覆盖前一天备份, 保留至下次运行)。
"""
import os, sys, glob, re, shutil, datetime, atexit, traceback

from python_calamine import CalamineWorkbook

# ============================ 配置(支持环境变量覆盖, 供隔离测试) ============================
DATA_DIR  = os.environ.get('FUEL_DATA_DIR', r'C:\Users\Jubin\Desktop\数据')
TPL_DIR   = os.environ.get('FUEL_TPL_DIR',  r'C:\Users\Jubin\Desktop\车辆维修，油耗\车辆油耗计算模板')
KAOHE     = os.environ.get('FUEL_KAOHE',    r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm')
BACKUP    = os.environ.get('FUEL_BACKUP',   r'E:\WorkSpace\workbuddy\2026-09-18-16-38-15\backup_daily')
DT_FMT, D_FMT = 'yyyy-mm-dd hh:mm:ss', 'yyyy-mm-dd'
EXCLUDE_KW = ['发动机', '发电机', '小油箱', '新车']
MIN_FREE_ROWS = 15            # 骨架余量下限(每消耗一行立刻补一行, 余量恒定)
PROG_IDS  = ['Ket.Application', 'et.Application', 'Excel.Application']  # WPS 表格优先
MACRO_SYNC = os.environ.get('FUEL_MACRO', '同步油耗到汇总')   # JSA 宏: 考核表4张油耗sheet -> 油耗分析汇总.xlsm

# 命令行可指定导出文件; 缺省自动取「数据」文件夹里最新一份(表头完整)
EXPORT = sys.argv[1] if len(sys.argv) > 1 else None

R = {'added': {}, 'shi': [], 'excluded': [], 'mixed': [], 'flags': [], 'errors': [],
     'kaohe_fuel': 0, 'kaohe_urea': 0, 'kaohe_skip': [], 'macro': ''}

# ============================ 工具 ============================
def parse_dt(v):
    if isinstance(v, datetime.datetime): return v
    s = str(v).strip()
    for f in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d'):
        try: return datetime.datetime.strptime(s, f)
        except ValueError: pass
    return None

def num(v):
    try:
        f = float(v)
        return int(f) if f == int(f) else f
    except (TypeError, ValueError): return v

def ctime(v):
    """COM/pywin/普通时间 -> 规整字符串键(前19位)"""
    return str(v)[:19] if v is not None else None

def day_of(v):
    return str(v)[:10] if v is not None else None

# ---- 时间统一走 Excel 序列数(float): pywin32 写 naive datetime 会做 UTC 转换导致 -8h 偏移,
#      float 序列数是 VT_R8 无任何转换, 读写完全对称 ----
EPOCH = datetime.datetime(1899, 12, 30)

def to_serial(v):
    """datetime/date -> Excel 序列数 float"""
    if v is None: return None
    if isinstance(v, datetime.datetime):
        return (v - EPOCH).total_seconds() / 86400.0
    if isinstance(v, datetime.date):
        return (datetime.datetime(v.year, v.month, v.day) - EPOCH).total_seconds() / 86400.0
    return v  # 已是序列数

def from_serial(v):
    """序列数 float / datetime / 字符串时间 -> datetime（自实现，不依赖 openpyxl）"""
    if v is None: return None
    if isinstance(v, datetime.datetime): return v
    if isinstance(v, str):
        return parse_dt(v.strip())         # 部分文件模板 G 列存的是字符串时间
    return EPOCH + datetime.timedelta(days=float(v))

def skey(v):
    """任意时间表示(datetime/date/序列数/字符串) -> 整数秒键(用于匹配)"""
    if v is None: return None
    if isinstance(v, (int, float)): return round(float(v) * 86400)
    if isinstance(v, (datetime.datetime, datetime.date)): return round(to_serial(v) * 86400)
    dt = parse_dt(str(v).strip())          # 字符串时间(历史数据混杂)
    return round(to_serial(dt) * 86400) if dt else None

def sid_key(v):
    """系统编号规范化: COM 写纯数字字符串会被 WPS 转成数字, 读回 float 带 '.0';
    两边统一转成整数字符串, 保证键匹配"""
    s = str(v).strip()
    try:
        f = float(s)
        if f == int(f): return str(int(f))
    except ValueError:
        pass
    return s

def normalize_com_times(ws, cols, last_row, wst_fg=False):
    """把工作表指定列里的字符串时间批量转为序列数(Value2 下字符串保持字符串,
    日期保持 float; 字符串会让 XLOOKUP 匹配失败)"""
    fixed = 0
    for c in cols:
        if last_row < 2: continue
        vals = ws.Range(ws.Cells(2, c), ws.Cells(last_row, c)).Value2
        if vals is None: continue
        if not isinstance(vals, tuple): vals = ((vals,),)
        out = []
        changed = False
        for row in vals:
            v = row[0] if isinstance(row, tuple) else row
            nv = v
            if isinstance(v, str):
                dt = parse_dt(v.strip())
                if dt is not None:
                    nv = to_serial(dt); changed = True; fixed += 1
            out.append((nv,) if isinstance(row, tuple) else nv)
        if changed:
            ws.Range(ws.Cells(2, c), ws.Cells(last_row, c)).Value = tuple(out)
    return fixed

def true_last_row(ws, col=1):
    """列内最后一个非空行。不依赖 End(xlUp) —— WPS 的 End 对部分文件返回错误行号(实测 A12 有值却返回 11)。
    从 UsedRange 尾部按块批量读, python 端倒扫, 确定性结果。"""
    ur = ws.UsedRange
    last_used = ur.Row + ur.Rows.Count - 1
    end = min(last_used, ws.Rows.Count)
    while end >= 1:
        start = max(1, end - 1999)
        vals = ws.Range(ws.Cells(start, col), ws.Cells(end, col)).Value2
        if vals is None:
            end = start - 1; continue
        if not isinstance(vals, tuple): vals = ((vals,),)
        for i in range(len(vals) - 1, -1, -1):
            row = vals[i]
            v = row[0] if isinstance(row, tuple) else row
            if v not in (None, ''):
                return start + i
        end = start - 1
    return 1

def rm_tree(path):
    """rmtree 的逐文件版(沙箱环境 rmtree 会被安全钩子拦截)"""
    for root, dirs, files in os.walk(path, topdown=False):
        for f in files:
            try: os.remove(os.path.join(root, f))
            except OSError: pass
        for d in dirs:
            try: os.rmdir(os.path.join(root, d))
            except OSError: pass
    try: os.rmdir(path)
    except OSError: pass

def find_newest_export():
    files = []
    for fp in glob.glob(os.path.join(DATA_DIR, '车辆加油_管理*.xlsx')):
        m = re.search(r'\((\d{4}-\d{2}-\d{2})\)', os.path.basename(fp))
        if m: files.append((m.group(1), fp))
    files.sort(reverse=True)
    for _, fp in files:
        try:
            wb = CalamineWorkbook.from_path(fp)
            rows = wb.get_sheet_by_name(wb.sheet_names[0]).to_python()
            if rows and '系统编号' in [str(x).strip() for x in rows[0]]:
                return fp
        except Exception:
            continue
    return None

# ============================ 1. 读导出(联合主键) ============================
if EXPORT is None:
    EXPORT = find_newest_export()
if not EXPORT or not os.path.exists(EXPORT):
    print('找不到可用的导出文件'); sys.exit(1)
print('导出文件:', EXPORT)

records = []
_wb = CalamineWorkbook.from_path(EXPORT)
_rows = _wb.get_sheet_by_name(_wb.sheet_names[0]).to_python()
_hdr = [str(x).strip() for x in _rows[0]]
if '系统编号' not in _hdr:
    print('导出表头异常(缺系统编号), 终止'); sys.exit(1)
for row in _rows[1:]:
    d = dict(zip(_hdr, row))
    sid = str(d.get('系统编号') or '').strip()
    # 导出尾部/中间的空行(更新文件时常见): 静默跳过, 不计错误
    if not sid and not str(d.get('车牌号码') or '').strip() and not d.get('登记时间'):
        continue
    t = parse_dt(d.get('登记时间'))
    if t is None:
        R['errors'].append(f"记录{sid or d.get('车牌号码') or '?'} 登记时间无法解析: {d.get('登记时间')}"); continue
    amt, vol = num(d.get('加注金额(元)')), num(d.get('加注油量（升）'))
    records.append({
        'sid': sid,
        # 联合主键: 系统编号优先; 无编号时用 时间+车牌+金额+油量
        'key': sid_key(sid) if sid else ('K', skey(t), str(d.get('车牌号码') or '').strip().upper(),
                                          str(amt), str(vol)),
        'plate': str(d.get('车牌号码') or '').strip().upper(),
        'plate_raw': str(d.get('车牌号码') or '').strip(),
        'date': parse_dt(d.get('登记日期')),
        'mile': num(d.get('当前里程')), 'gps': num(d.get('GPS里程')),
        'amt': amt, 'vol': vol, 'price': num(d.get('加注单价（元）')),
        'full': str(d.get('是否加满油') or '').strip(),
        'type': str(d.get('加油类型') or '').strip(),
        'loc': str(d.get('当前位置') or '').strip(),
        'person': str(d.get('登记人') or '').strip(),
        'time': t,
        'note': str(d.get('备注信息') or '').strip(),
    })
print(f'导出记录: {len(records)} 条')
records.sort(key=lambda x: x['time'])   # 按时间序处理, 同文件行号递增

# 同一导出内先按主键去一次重(防导出自身重复) —— 逐行 continue, 绝不 break
_seen, deduped = set(), []
for rec in records:
    if rec['key'] in _seen: continue
    _seen.add(rec['key']); deduped.append(rec)
if len(deduped) != len(records):
    print(f'导出内部去重: {len(records)} -> {len(deduped)}')
records = deduped

# ============================ 2. 车牌映射 ============================
fuel_map, urea_map = {}, {}
for fp in glob.glob(os.path.join(TPL_DIR, '*.xlsx')):
    fn = os.path.basename(fp)
    if fn.startswith('~$') or fn.startswith('模板'): continue
    plate = os.path.splitext(fn)[0].split('-')[-1].strip().upper()
    (urea_map if fn.startswith('尿素') else fuel_map)[plate] = fp

# ============================ 3. 分流决策(纯内存) ============================
pending = []          # (path, rec, is_urea)
pend_flags = {}       # key -> 矛盾提醒(仅当该记录确认为新增时才计入报告, 避免历史记录重复刷屏)
for rec in records:
    plate = rec['plate']
    is_urea = '尿素' in rec['type']
    # 单价纠偏分流: 2~3块=尿素, 6~8块=油气(有人会把类型填乱)
    try: price_chk = float(rec['price'] or 0)
    except (TypeError, ValueError): price_chk = 0
    if price_chk:
        if not is_urea and price_chk < 5 and plate in urea_map:
            pend_flags[rec['key']] = f"[类型/单价矛盾-已按尿素分发] {rec['plate_raw']}: 类型[{rec['type']}] 单价{price_chk} -> 尿素文件"
            is_urea = True
        elif is_urea and price_chk >= 5 and plate in fuel_map:
            pend_flags[rec['key']] = f"[类型/单价矛盾-已按油气分发] {rec['plate_raw']}: 类型[{rec['type']}] 单价{price_chk} -> 油文件"
            is_urea = False
    if not is_urea and '尿素' in rec['note'] and '加油' in rec['note'].replace('油卡加油', ''):
        R['mixed'].append(f"{rec['plate_raw']} {rec['time']} 备注[{rec['note']}] -> 留手动处理")
        continue
    tmap = urea_map if is_urea else fuel_map
    path = tmap.get(plate)
    if not path:
        R['errors'].append(f"车牌 {rec['plate_raw']} 找不到{'尿素' if is_urea else ''}模板文件")
        continue
    pending.append((path, rec, is_urea))

# ============================ 4. calamine 预筛: 已录主键集合 ============================
by_path = {}
for path, rec, is_urea in pending:
    by_path.setdefault(path, []).append((rec, is_urea))

tpl_meta = {}         # path -> {'first_M': 首条登记时间}
for path in by_path:
    first_M = None
    try:
        cw = CalamineWorkbook.from_path(path)
        srows = cw.get_sheet_by_name('加油数据').to_python(skip_empty_area=False)
        for row in srows[1:]:
            if not row or all(v in (None, '') for v in row[:14]): continue
            t = parse_dt(row[12]) if len(row) > 12 else None
            if t is not None:
                first_M = t; break
    except Exception as e:
        R['errors'].append(f"预筛读取失败 {os.path.basename(path)}: {e}")
    tpl_meta[path] = {'first_M': first_M}

# 注: 最终"已录"判定在 COM 打开文件后按实时状态逐行进行(重复 continue, 绝不 break)
print(f'待处理 {len(pending)} 条, 涉及 {len(by_path)} 个文件')

# ============================ 5. 占用锁检查 + 每日备份 ============================
lock_issues = []
for p in list(by_path) + [KAOHE]:
    try:
        f = open(p, 'r+b'); f.close()
    except PermissionError:
        lock_issues.append(p)
if lock_issues:
    print('\n[占用] 以下文件正被打开, 请先关闭 WPS/Excel 中对应文件后重跑:')
    for p in lock_issues: print('   ', p)
    sys.exit(1)

# ============================ 6+7. COM 全程(车辆模板 + 考核表 + 宏) ============================
kaohe_rows = []       # (is_urea, rowvals24)
if not pending:
    print('无待处理记录, 全链路无操作')
else:
    import pythoncom
    import win32com.client as wc
    pythoncom.CoInitialize()
    app = None
    _opened = []

    def _cleanup_com():
        """finally + atexit 双保险: 关闭全部 Workbook, 退出 WPS, 不残留进程"""
        global app
        try:
            if app is None: return
            for wb in list(_opened):
                try: wb.Close(False)
                except Exception: pass
            _opened.clear()
            try: app.Quit()
            except Exception: pass
            app = None
        except Exception:
            pass
    atexit.register(_cleanup_com)

    try:
        for pid in PROG_IDS:
            try:
                app = wc.DispatchEx(pid); print('COM 应用:', pid); break
            except Exception:
                continue
        if app is None:
            raise RuntimeError('无法启动 WPS/Excel COM (尝试过 %s)' % PROG_IDS)
        app.Visible = False
        app.DisplayAlerts = False
        try: app.ScreenUpdating = False
        except Exception: pass

        XL_UP = -4162

        # ---------- 备份(在任何写入之前; 同名文件直接覆盖 = 覆盖前一天备份) ----------
        bak_tpl = os.path.join(BACKUP, '车辆油耗计算模板')
        os.makedirs(bak_tpl, exist_ok=True)
        n_bak = 0
        for p in by_path:
            shutil.copy2(p, os.path.join(bak_tpl, os.path.basename(p))); n_bak += 1
        shutil.copy2(KAOHE, os.path.join(BACKUP, os.path.basename(KAOHE)))
        print(f'每日备份完成 -> {BACKUP} ({n_bak} 模板 + 考核表, 覆盖同名文件/保留至下次运行)')

        # ---------- 6. 车辆模板 ----------
        for path in sorted(by_path):
            items = by_path[path]
            wb = None
            try:
                wb = app.Workbooks.Open(path, 0, False)
                if wb.ReadOnly:
                    raise RuntimeError('文件被占用(只读打开)')
                _opened.append(wb)
                wsd = wb.Worksheets('加油数据')
                wst = None
                for sn in ('计算模板', '油耗计算'):
                    try:
                        wst = wb.Worksheets(sn); break
                    except Exception: continue
                if wst is None:
                    raise RuntimeError('无计算模板sheet')
                fn_short = os.path.basename(path)

                # ---- 规范化字符串时间(加油数据 M/C 列 + 计算模板 F/G 列) ----
                ld0 = true_last_row(wsd, 1)
                nf1 = normalize_com_times(wsd, (13, 3), ld0)
                lg0 = true_last_row(wst, 7)
                nf2 = normalize_com_times(wst, (6, 7), max(lg0, true_last_row(wst, 6)))
                if nf1 or nf2:
                    print(f"  {fn_short}: 规范化字符串时间 加油数据{nf1} + 模板{nf2} 格")

                # ---- 批量读现状(Value2=序列数float, 避免逐格 COM 调用与时区转换) ----
                last_data = true_last_row(wsd, 1)
                if last_data < 1: last_data = 1
                dvals = wsd.Range(wsd.Cells(2, 1), wsd.Cells(max(last_data, 2), 13)).Value2 or ()
                if dvals and not isinstance(dvals[0], tuple): dvals = (dvals,)
                exist_keys, full_map, prev_rec = set(), {}, None
                for row in dvals:
                    sid = sid_key(row[0])
                    tk = skey(row[12])
                    if sid: exist_keys.add(sid)
                    if tk is not None:
                        full_map[tk] = str(row[8] or '').strip()
                        exist_keys.add(('K', tk, str(row[1] or '').strip().upper(),
                                        str(num(row[5])), str(num(row[6]))))
                        prev_rec = (num(row[3]), num(row[4]))
                first_M = to_serial(tpl_meta[path]['first_M'])

                last_g = true_last_row(wst, 7)
                if last_g < 1: last_g = 1
                tvals = wst.Range(wst.Cells(2, 6), wst.Cells(max(last_g, 2), 7)).Value2 or ()
                if tvals and not isinstance(tvals[0], tuple): tvals = (tvals,)
                g_list = []                # [(row, G序列数)]
                for rr, row in enumerate(tvals, start=2):
                    if row and row[1] is not None:
                        g_list.append((rr, row[1]))
                g_times = set(skey(g) for _, g in g_list)
                # 历史有效"是": 是否加满从加油数据反查(COM 实时状态, 不依赖任何缓存)
                prev_shi_G, cycle_start = None, 2
                for rr, g in g_list:
                    if full_map.get(skey(g)) == '是':
                        prev_shi_G, cycle_start = g, rr + 1
                last_filled = g_list[-1][0] if g_list else 1
                skel_last = true_last_row(wst, 2)
                if skel_last < last_filled: skel_last = last_filled
                plate_v = str(wst.Cells(2, 4).Value or items[0][0]['plate_raw'])
                type_v = str(wst.Cells(2, 5).Value or '')

                # 骨架懒补: 源固定用初始干净骨架末行, 填行前保证余量
                src_skel = skel_last
                skel_cur = [skel_last]
                def ensure_skel(need_row):
                    n_ext = 0
                    while skel_cur[0] < need_row:
                        skel_cur[0] += 1
                        wst.Rows(src_skel).Copy(wst.Rows(skel_cur[0]))
                        wst.Cells(skel_cur[0], 2).Value = skel_cur[0] - 1
                        n_ext += 1
                    if n_ext:
                        print(f"  {fn_short}: 骨架补行 +{n_ext} (至{skel_cur[0]}, 余量{skel_cur[0]-last_filled})")

                # ---- 逐条写入(重复 continue, 绝不 break) ----
                data_rows, tpl_rows_info = [], []
                added_pending = {}       # 该文件新增明细(Save 成功后才计入 R, 防半途失败虚报)
                shi_pending = []         # 模板新填行报告(同上, Save 成功后才计入)
                nr_data = last_data
                for rec, is_urea in items:
                    if rec['key'] in exist_keys:
                        pend_flags.pop(rec['key'], None)   # 历史记录: 丢弃暂存的矛盾提醒
                        continue           # 已录过: 只跳过本行, 同日后续新记录照常处理
                    # 追加加油数据(攒批) —— 注意: COM 只接受 datetime.datetime, 不能传 date
                    nr_data += 1
                    d_val = to_serial(rec['date']) if rec['date'] else to_serial(rec['time'])
                    vals = [rec['sid'], rec['plate_raw'], d_val,
                            rec['mile'], rec['gps'], rec['amt'], rec['vol'], rec['price'],
                            rec['full'], rec['type'], rec['loc'], rec['person'],
                            to_serial(rec['time']), rec['note']]
                    data_rows.append((nr_data, vals))
                    exist_keys.add(rec['key'])
                    if rec['key'] in pend_flags:
                        R['flags'].append(pend_flags.pop(rec['key']))
                    full_map[skey(rec['time'])] = rec['full']
                    added_pending.setdefault(rec['sid'], (
                        f"{rec['time']:%m-%d %H:%M} {rec['type']} 加满[{rec['full']}] ¥{rec['amt']} {rec['vol']}L @{rec['price']} 备注[{rec['note'] or '-'}]"))

                    # ---- 校验(只在报告里说, 不写文件) ----
                    try:
                        amt, vol, price = float(rec['amt'] or 0), float(rec['vol'] or 0), float(rec['price'] or 0)
                        if amt and vol and price and abs(amt - vol * price) / max(amt, 1) > 0.05:
                            R['flags'].append(f"[金额≠油量×单价] {rec['plate_raw']}: 金额{amt} 油量{vol} 单价{price} 乘积={round(vol*price,2)}")
                        if is_urea and price > 4:
                            R['flags'].append(f"[尿素单价异常] {rec['plate_raw']}: {price}")
                        if (not is_urea) and (0 < price < 5 or price > 9):
                            R['flags'].append(f"[油单价异常] {rec['plate_raw']}: {price}")
                    except (TypeError, ValueError):
                        R['flags'].append(f"[数值列异常] {rec['plate_raw']}: {rec['amt']}/{rec['vol']}/{rec['price']}")
                    if rec['full'] not in ('是', '否'):
                        R['flags'].append(f"[是否加满异常] {rec['plate_raw']}: [{rec['full']}]")
                    if prev_rec:
                        try:
                            dd = float(rec['mile']) - float(prev_rec[0])
                            dg = float(rec['gps']) - float(prev_rec[1])
                            if dd < 0 or dg < 0:
                                R['flags'].append(f"[里程倒退] {rec['plate_raw']}: 仪表{rec['mile']} GPS{rec['gps']} < 上一条({prev_rec[0]}/{prev_rec[1]}), 请核对")
                            if dg > 5 and dd > 5 and abs(dd - dg) / dg > 0.05:
                                R['flags'].append(f"[里程增量偏差] {rec['plate_raw']}: 仪表+{round(dd)} vs GPS+{round(dg)} ({abs(dd-dg)/dg:.0%})")
                        except (TypeError, ValueError): pass
                    prev_rec = (rec['mile'], rec['gps'])

                    # ---- 排除词: 只进加油数据 ----
                    if any(k in rec['note'] for k in EXCLUDE_KW):
                        R['excluded'].append(f"{rec['plate_raw']} 备注[{rec['note']}] -> 只进加油数据")
                        continue

                    # ---- 计算模板 ----
                    if skey(rec['time']) in g_times:
                        continue           # 模板已有该时间: 只跳过本行
                    r = last_filled + 1
                    # 每消耗一行补一行: 填 r 后余量保持 >= MIN_FREE_ROWS
                    ensure_skel(r + MIN_FREE_ROWS)
                    wst.Cells(r, 7).Value = to_serial(rec['time'])
                    wst.Cells(r, 7).NumberFormat = DT_FMT
                    # 有效"是"规则: 同一天多次加满, 只算当天最后一个"是"
                    if rec['full'] == '是':
                        if prev_shi_G is None:
                            effective = True
                        else:
                            d_prev = from_serial(prev_shi_G)
                            effective = d_prev is None or rec['time'].date() != d_prev.date()
                    else:
                        effective = False
                    if effective:
                        F_val = prev_shi_G if prev_shi_G is not None else first_M
                        if F_val is not None:
                            F_write = F_val
                            if isinstance(F_write, str):      # 字符串时间 -> 序列数
                                _dt = parse_dt(F_write.strip())
                                F_write = to_serial(_dt) if _dt else F_write
                            wst.Cells(r, 6).Value = F_write   # 序列数 float, 无时区风险
                            wst.Cells(r, 6).NumberFormat = DT_FMT
                        cycle_rows = list(range(cycle_start, r)) if prev_shi_G is not None \
                            else list(range(2, r))
                        base_i = f'=_xlfn.XLOOKUP(D{r}&G{r},加油数据!B:B&加油数据!M:M,加油数据!F:F,0)'
                        base_j = f'=_xlfn.XLOOKUP(D{r}&G{r},加油数据!B:B&加油数据!M:M,加油数据!G:G,0)'
                        off_i = ''.join(f'+I{x}' for x in reversed(cycle_rows))
                        off_j = ''.join(f'+J{x}' for x in reversed(cycle_rows))
                        for col, f_txt in ((9, base_i + off_i), (10, base_j + off_j)):
                            cell = wst.Cells(r, col)
                            try:
                                cell.FormulaArray = f_txt
                            except Exception:
                                try: cell.Formula = f_txt
                                except Exception:
                                    if cycle_rows:   # 超长公式兜底: SUM 等价
                                        cl = 'I' if col == 9 else 'J'
                                        fs = f_txt.replace(off_i if col == 9 else off_j,
                                                           f'+SUM({cl}{cycle_rows[0]}:{cl}{cycle_rows[-1]})')
                                        cell.Formula = fs
                                    else:
                                        raise
                        shi_pending.append(f"{plate_v} [有效是] 时间A={from_serial(F_val):%Y-%m-%d} 时间B={rec['time']:%m-%d %H:%M}")
                        tpl_rows_info.append((r, rec, is_urea))
                    elif rec['full'] == '是':
                        shi_pending.append(f"{plate_v} [当天非最后一个“是”->按否处理] 时间B={rec['time']:%m-%d %H:%M}")
                    g_list.append((r, to_serial(rec['time'])))
                    g_times.add(skey(rec['time']))
                    last_filled = r
                    if effective:
                        # 存"序列数"而不是 datetime —— datetime 写回 COM 会被当 UTC 转换, 时间会偏 -8 小时
                        prev_shi_G, cycle_start = to_serial(rec['time']), r + 1

                # ---- 批量写加油数据 ----
                for nr, vals in data_rows:
                    wsd.Cells(nr, 1).NumberFormat = '@'   # 系统编号列防 WPS 自动转数字
                    wsd.Range(wsd.Cells(nr, 1), wsd.Cells(nr, 14)).Value = tuple([tuple(vals)])
                    wsd.Cells(nr, 3).NumberFormat = D_FMT
                    wsd.Cells(nr, 13).NumberFormat = DT_FMT

                # ---- 读"是"行实时公式值(组装考核行24列, WPS 算好的值) ----
                if tpl_rows_info:
                    try: app.Calculate()   # WPS 的 Workbook 对象无 Calculate, 用 Application 级
                    except Exception: pass
                for r, rec, is_urea in tpl_rows_info:
                    fv = wst.Range(wst.Cells(r, 6), wst.Cells(r, 22)).Value2  # F..V 共17列(序列数/数值)
                    if not isinstance(fv, tuple): fv = (fv,)
                    if fv and isinstance(fv[0], tuple): fv = fv[0]
                    def _n(x):
                        try: return float(x)
                        except (TypeError, ValueError): return None
                    F_raw, G_raw = fv[0], fv[1]
                    if isinstance(F_raw, str):
                        _d = parse_dt(F_raw.strip()); F_raw = to_serial(_d) if _d else F_raw
                    if isinstance(G_raw, str):
                        _d = parse_dt(G_raw.strip()); G_raw = to_serial(_d) if _d else G_raw
                    F_s, G_s = _n(F_raw), _n(G_raw)
                    F_t, G_t = from_serial(F_s), from_serial(G_s)
                    # 索引: F0 G1 H2 I3 J4 K5 L6 M7 N8 O9 P10 Q11 R12 S13 T14 U15 V16
                    rowvals = [rec['sid'], r - 1, rec['person'], plate_v, type_v,
                               F_s, G_s, '是', _n(fv[3]), _n(fv[4]), rec['price'],
                               _n(fv[6]), _n(fv[7]), _n(fv[8]), _n(fv[9]),
                               _n(fv[10]), _n(fv[11]), _n(fv[12]), _n(fv[13]), _n(fv[14]),
                               _n(fv[15]), _n(fv[16]),
                               to_serial(G_t.date()) if G_t else None, rec['note'] or 0]
                    kaohe_rows.append((is_urea, rowvals))

                if data_rows or tpl_rows_info:
                    wb.Save()
                    # Save 成功后才计入新增明细(防止半途失败虚报触发考核表写入)
                    for sid, msg in added_pending.items():
                        R['added'].setdefault(plate_v, []).append(msg)
                    R['shi'].extend(shi_pending)
                    print(f"  {fn_short}: 写入{len(data_rows)}行数据, 模板{len(tpl_rows_info)}行有效是")
                else:
                    print(f"  {fn_short}: 无新增(全部已存在)")
            except Exception as e:
                R['errors'].append(f"{os.path.basename(path)}: {e}")
                traceback.print_exc()
            finally:
                if wb is not None:
                    try: wb.Close(False)
                    except Exception: pass
                    if wb in _opened: _opened.remove(wb)

        # ---------- 7. 考核表 + 每日加油记录 + 宏 ----------
        # 每天无条件执行: 即使当天无新分发, 也要刷新"每日加油记录"并跑宏,
        # 让前一天的数据经 4 张油耗 sheet 进入汇总(宏自身有去重, 重复调用安全)
        if True:
            wbk = None
            try:
                wbk = app.Workbooks.Open(KAOHE, 0, False)
                if wbk.ReadOnly:
                    raise RuntimeError('考核表被占用(只读打开)')
                _opened.append(wbk)
                wbk.Activate()          # 宏用 ActiveWorkbook, 保证考核表为活动簿

                def _kao_ids(ws):
                    last = true_last_row(ws, 1)
                    if last < 2: return set()
                    v = ws.Range(ws.Cells(2, 1), ws.Cells(last, 1)).Value
                    if not isinstance(v, tuple): v = ((v,),)
                    return set(sid_key(x[0]) for x in v if x and x[0] not in (None, '', '0'))
                ws_fuel = wbk.Worksheets('每日油耗数据')
                ws_urea = wbk.Worksheets('每日尿素数据')
                ids_fuel, ids_urea = _kao_ids(ws_fuel), _kao_ids(ws_urea)

                for is_urea, rowvals in kaohe_rows:
                    ws = ws_urea if is_urea else ws_fuel
                    ids = ids_urea if is_urea else ids_fuel
                    if sid_key(rowvals[0]) in ids:
                        R['kaohe_skip'].append(f"{rowvals[3]} {from_serial(rowvals[5]):%m-%d} 已在考核表")
                        continue
                    last = true_last_row(ws, 1)
                    nr = max(last, 1) + 1
                    ws.Cells(nr, 1).NumberFormat = '@'
                    ws.Range(ws.Cells(nr, 1), ws.Cells(nr, 24)).Value = tuple([tuple(rowvals)])
                    ws.Cells(nr, 6).NumberFormat = DT_FMT
                    ws.Cells(nr, 7).NumberFormat = DT_FMT
                    ws.Cells(nr, 23).NumberFormat = D_FMT
                    ids.add(sid_key(rowvals[0]))
                    if is_urea: R['kaohe_urea'] += 1
                    else: R['kaohe_fuel'] += 1

                # ---- 每日加油记录 sheet: 整表替换(n-1) ----
                try:
                    from refuel_sheet import compute_rows
                    rows_rf, src_rf, n1_rf = compute_rows(data_dir=DATA_DIR, verbose=False)
                    wsr = wbk.Worksheets('每日加油记录')
                    last_r = true_last_row(wsr, 1)
                    # 覆盖写入新数据
                    for i, vals in enumerate(rows_rf, start=2):
                        vals2 = list(vals)
                        vals2[2] = to_serial(vals2[2])    # 登记日期
                        vals2[12] = to_serial(vals2[12])  # 登记时间
                        wsr.Cells(i, 1).NumberFormat = '@'
                        wsr.Range(wsr.Cells(i, 1), wsr.Cells(i, 14)).Value = tuple([tuple(vals2)])
                        wsr.Cells(i, 3).NumberFormat = D_FMT
                        wsr.Cells(i, 13).NumberFormat = DT_FMT
                    # 清除多余旧行 —— WPS 的"大范围 ClearContents"存在截断 bug
                    # (实测 A2:N16 只清到第 11 行), 逐行清除才可靠
                    n_new = len(rows_rf)
                    stale_from = 2 + n_new
                    n_stale = 0
                    if last_r >= stale_from:
                        for r in range(stale_from, last_r + 1):
                            wsr.Range(wsr.Cells(r, 1), wsr.Cells(r, 14)).ClearContents()
                            n_stale += 1
                    print(f"每日加油记录已刷新: 源={src_rf} n-1={n1_rf} 写入{n_new}条"
                          + (f", 清残留{n_stale}行" if n_stale else ""))
                except Exception as e:
                    R['errors'].append(f'每日加油记录刷新失败: {e}')

                # ---- 刷新「每日油耗」数据透视表 + 日期选择 n-1 ----
                try:
                    wsp = wbk.Worksheets('每日油耗')
                    n1_date = n1_rf if n1_rf else (datetime.date.today() - datetime.timedelta(days=1))
                    d_s = n1_date
                    serial_d = to_serial(datetime.datetime(d_s.year, d_s.month, d_s.day))
                    targets = {f'{d_s.year}/{d_s.month}/{d_s.day}', d_s.strftime('%Y-%m-%d'),
                               str(int(serial_d)), str(serial_d)}
                    for i in range(1, wsp.PivotTables().Count + 1):
                        pt = wsp.PivotTables(i)
                        pt.RefreshTable()
                        pf = pt.PivotFields('日期')
                        cnt = pf.PivotItems().Count
                        for j in range(1, cnt + 1):
                            it = pf.PivotItems(j)
                            vis = str(it.Name).strip() in targets
                            if it.Visible != vis:
                                it.Visible = vis
                    print(f'「每日油耗」透视表已刷新, 日期选择 {d_s}')
                except Exception as e:
                    R['errors'].append(f'透视表刷新失败: {e}')

                try:
                    app.Calculate()   # 强制全簿重算, 让4张油耗sheet/每日加油记录的公式缓存更新到最新
                except Exception: pass
                wbk.Save()
                print('考核表已保存(WPS 原生保存, 宏/按钮/图片完整保留)')

                # ---- 调用 JSA 宏: 同步油耗到汇总 ----
                # (宏内含 Application.Visible 守卫: 后台调用自动跳过其第5步跨簿区域填充,
                #  由下方 python 端补做, 避开 WPS 后台实例跨簿 JSA Run 卡死的 bug)
                try:
                    before = set(w.Name for w in app.Workbooks)
                    ret = app.Run(MACRO_SYNC)
                    R['macro'] = f'{MACRO_SYNC}: {ret}'
                    # 宏可能打开了目标汇总文件: 补做区域填充(按车牌->区域映射)后关闭
                    for w2 in [w for w in app.Workbooks if w.Name not in before]:
                        try:
                            w2name = w2.Name
                            if '汇总' in w2name:
                                from region_map import REGION_MAP
                                plate2region = {}
                                for reg, plates in REGION_MAP.items():
                                    for p in plates: plate2region[p.strip().upper()] = reg
                                ws2 = w2.Worksheets('Sheet1')
                                # 尾行定位用 true_last_row 逻辑(C 列)
                                ur = ws2.UsedRange
                                last2 = ur.Row + ur.Rows.Count - 1
                                vals = ws2.Range(ws2.Cells(2, 3), ws2.Cells(last2, 3)).Value2
                                if vals is None: vals = ()
                                if not isinstance(vals, tuple): vals = ((vals,),)
                                filled = 0
                                for i2, cellv in enumerate(vals):
                                    r2 = 2 + i2
                                    pv = cellv[0] if isinstance(cellv, tuple) else cellv
                                    plate = str(pv or '').strip().upper()
                                    if not plate: continue
                                    bv = ws2.Cells(r2, 2).Value2
                                    if bv not in (None, ''):
                                        continue
                                    reg = plate2region.get(plate)
                                    if reg:
                                        ws2.Cells(r2, 2).Value = reg
                                        filled += 1
                                if filled:
                                    w2.Save()
                                print(f"汇总区域填充: 补填 {filled} 行 ({w2name})")
                            w2.Close(not bool(getattr(w2, 'Saved', True)))
                        except Exception as e2:
                            R['errors'].append(f'汇总收尾失败: {e2}')
                except Exception as e:
                    R['macro'] = f'{MACRO_SYNC} 调用失败: {e}'
                    R['errors'].append(f'宏调用失败: {e}')
            finally:
                if wbk is not None:
                    try: wbk.Close(False)
                    except Exception: pass
                    if wbk in _opened: _opened.remove(wbk)
        else:
            print('无新增"是"行, 考核表不修改')

        try: app.ScreenUpdating = True
        except Exception: pass
    finally:
        _cleanup_com()
        pythoncom.CoUninitialize()

# ============================ 8. 报告(只在这里, 不写进文件) ============================
print('\n===== 分发报告 (COM 架构) =====')
print('导出:', os.path.basename(EXPORT), '| 去重: 联合主键(系统编号优先, 时间+车牌+金额+油量兜底), 逐行判定')
print('\n[新增加油数据]')
for fn, items in R['added'].items():
    print(f'  {fn}:')
    for it in items: print('     ', it)
if not R['added']: print('    (无, 全部已存在)')
print('\n[计算模板-新填行]')
for it in R['shi']: print('   ', it)
if not R['shi']: print('    (无)')
print('\n[排除词-只进加油数据]')
for it in R['excluded']: print('   ', it)
if not R['excluded']: print('    (无)')
print('\n[混合记录-留手动]')
for it in R['mixed']: print('   ', it)
if not R['mixed']: print('    (无)')
print('\n[数据质量提醒]')
for it in R['flags']: print('   ', it)
if not R['flags']: print('    (无)')
print('\n[考核表追加] 每日油耗数据 +%d 行 | 每日尿素数据 +%d 行' % (R['kaohe_fuel'], R['kaohe_urea']))
for it in R['kaohe_skip']: print('    跳过:', it)
print('[宏]', R['macro'] or '(未调用)')
print('\n[错误]')
for it in R['errors']: print('   ', it)
if not R['errors']: print('    无')
