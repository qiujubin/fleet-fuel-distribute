# -*- coding: utf-8 -*-
"""补回真实考核表缺失的 2 个绘图 rels(drawing1/drawing3)"""
import zipfile, re, shutil, os

bak = r'E:\WorkSpace\workbuddy\2026-09-18-16-38-15\backup_20260919\油耗考核10月_宏.xlsm'
cur = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'

try:
    f = open(cur, 'r+b'); f.close(); print('考核表未占用 ✓')
except PermissionError:
    print('考核表被占用!! 先关闭 WPS'); raise SystemExit(1)

zb, zc = zipfile.ZipFile(bak), zipfile.ZipFile(cur)
nc = set(zc.namelist())
ok = True
for d in ('drawing1', 'drawing3'):
    rels = zb.read(f'xl/drawings/_rels/{d}.xml.rels').decode('utf-8')
    for rid, target in re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels):
        p = os.path.normpath(os.path.join('xl/drawings', target)).replace(os.sep, '/')
        if p not in nc:
            print(f'  {d} {rid} -> {target} 缺失: {p}')
            ok = False
print('图片依赖完整:', ok)
if ok:
    zin = zipfile.ZipFile(cur)
    entries = {i.filename: zin.read(i.filename) for i in zin.infolist()}
    zin.close()
    for d in ('drawing1', 'drawing3'):
        entries[f'xl/drawings/_rels/{d}.xml.rels'] = zb.read(f'xl/drawings/_rels/{d}.xml.rels')
    out = cur + '.staging.xlsm'
    zo = zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED)
    for name, data in entries.items():
        zo.writestr(name, data)
    zo.close()
    os.replace(out, cur)
    print('2 个 rels 已补回 ✓')
zc2 = zipfile.ZipFile(cur)
nc2 = set(zc2.namelist())
missing = [n for n in (set(zb.namelist()) - nc2) if not n.endswith('/')]
print('剩余缺失:', missing if missing else '无')
zc2.close()
zb.close(); zc.close()
