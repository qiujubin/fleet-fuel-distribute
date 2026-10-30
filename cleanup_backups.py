# -*- coding: utf-8 -*-
"""备份目录保留策略：只留最近 N 天（默认 3 天）

清理范围（全部按目录名里的日期判断，解析不出才用 mtime）：
  1. 工作区根目录  backup_*        （backup_daily 除外 —— 它是覆盖式，只留最新一份）
  2. archive/       backup_*       （一次性修正脚本的备份）
  3. 桌面「车辆维修，油耗」 backup_track_*  （轨迹截图工具的备份）

用法：
  python cleanup_backups.py            # 只看（dry-run），不删
  python cleanup_backups.py --apply    # 真删
  python cleanup_backups.py --keep 5   # 改保留天数
"""
import os, re, sys, shutil, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
TRACK_DIR = r"C:\Users\Jubin\Desktop\车辆维修，油耗"
KEEP_DAYS = 3
TODAY = datetime.date.today()


def dir_date(name, path):
    """从目录名解析日期，失败则退回 mtime"""
    base = os.path.basename(name)
    m = re.match(r'^(\d{4})-(\d{2})-(\d{2})$', base)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass
    m = re.search(r'(20\d{6})', name)
    if m:
        s = m.group(1)
        try:
            return datetime.date(int(s[:4]), int(s[4:6]), int(s[6:8]))
        except ValueError:
            pass
    m = re.search(r'[_-](\d{2})(\d{2})(?:_|$)', name)
    if m:
        try:
            y = TODAY.year
            d = datetime.date(y, int(m.group(1)), int(m.group(2)))
            if d > TODAY:
                d = datetime.date(y - 1, d.month, d.day)
            return d
        except ValueError:
            pass
    try:
        return datetime.date.fromtimestamp(os.path.getmtime(path))
    except Exception:
        return None


def dir_size(path):
    total = 0
    for root, _, files in os.walk(path):
        for f in files:
            try:
                total += os.path.getsize(os.path.join(root, f))
            except OSError:
                pass
    return total


def human(n):
    for unit in ("B", "K", "M", "G"):
        if n < 1024 or unit == "G":
            return f"{n:.0f}{unit}" if unit == "B" else f"{n:.1f}{unit}"
        n /= 1024


def collect():
    items = []
    # 1. 工作区根目录
    for n in os.listdir(HERE):
        p = os.path.join(HERE, n)
        if os.path.isdir(p) and n.startswith("backup_") and n != "backup_daily" and not n.startswith("backup_daily"):
            items.append((n, p, "根目录"))
    # 2. archive
    arc = os.path.join(HERE, "archive")
    if os.path.isdir(arc):
        for n in os.listdir(arc):
            p = os.path.join(arc, n)
            if os.path.isdir(p) and n.startswith("backup_"):
                items.append((n, p, "archive"))
    # 3. 桌面轨迹备份
    if os.path.isdir(TRACK_DIR):
        for n in os.listdir(TRACK_DIR):
            p = os.path.join(TRACK_DIR, n)
            if os.path.isdir(p) and n.startswith("backup_track_"):
                items.append((n, p, "backup_track"))
    # 4. backup_daily/history/YYYY-MM-DD（每日历史快照，latest 那份不在此列）
    hist_root = os.path.join(HERE, "backup_daily", "history")
    if os.path.isdir(hist_root):
        for n in os.listdir(hist_root):
            p = os.path.join(hist_root, n)
            if os.path.isdir(p):
                items.append((f"backup_daily/history/{n}", p, "history"))
    return items


def main():
    apply_ = "--apply" in sys.argv
    keep = KEEP_DAYS
    if "--keep" in sys.argv:
        i = sys.argv.index("--keep")
        if i + 1 < len(sys.argv):
            keep = int(sys.argv[i + 1])
    cutoff = TODAY - datetime.timedelta(days=keep - 1)
    print(f"保留 {keep} 天（{cutoff} 及以后保留），今天 {TODAY}，模式：{'删除' if apply_ else 'dry-run 仅预览'}")
    print("=" * 74)

    items = collect()
    dels, keeps = [], []
    for name, path, kind in items:
        d = dir_date(name, path)
        size = dir_size(path)
        (dels if (d and d < cutoff) else keeps).append((name, path, kind, d, size))

    print(f"\n【保留】{len(keeps)} 个")
    for name, _, kind, d, size in sorted(keeps, key=lambda x: str(x[3])):
        print(f"  ✓ {name:42s} {kind:10s} {str(d):10s} {human(size)}")

    print(f"\n【删除】{len(dels)} 个")
    freed = 0
    for name, _, kind, d, size in sorted(dels, key=lambda x: str(x[3])):
        print(f"  ✗ {name:42s} {kind:10s} {str(d):10s} {human(size)}")
        freed += size
    print(f"\n合计可释放: {human(freed)}")

    if not apply_:
        print("\n(--dry-run，未删除；加 --apply 执行)")
        return 0

    n = 0
    for name, path, _, _, _ in dels:
        try:
            shutil.rmtree(path)
            n += 1
        except Exception as e:
            print(f"  ! 删除失败 {name}: {e}")
    print(f"\n已删除 {n}/{len(dels)} 个目录，释放 {human(freed)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
