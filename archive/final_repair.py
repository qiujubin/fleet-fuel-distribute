# -*- coding: utf-8 -*-
"""一站式修复: 恢复全部 WPS 部件 + 可视化 sheet 按钮接线
- 从备份补齐: jdecontrols(29)/media(24)/metadata/drawings5,7,9(+rels)
- 可视化按钮改名挂载: 备份drawing3->drawing21(sheet11), drawing5->drawing23(sheet13),
  drawing7->drawing24(sheet15), drawing9->drawing25(sheet17), drawing1->drawing26(sheet6)
- 避开 openpyxl 已占用的 drawing1-4 (四个油耗sheet的活按钮)
"""
import zipfile, re, shutil, sys

BAK = r'E:\WorkSpace\workbuddy\2026-09-18-16-38-15\backup_20260919\油耗考核10月_宏.xlsm'
CUR = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
SAFETY = r'E:\WorkSpace\workbuddy\2026-09-18-16-38-15\kaohe_before_final_fix.xlsm'

try:
    f = open(CUR, 'r+b'); f.close()
except PermissionError:
    print('考核表仍被占用, 请先关闭 WPS!'); sys.exit(1)
shutil.copy2(CUR, SAFETY)

zb, zc = zipfile.ZipFile(BAK), zipfile.ZipFile(CUR)
nb, nc = set(zb.namelist()), set(zc.namelist())
entries = {i.filename: zc.read(i.filename) for i in zc.infolist()}

# ---- 1) 原样补齐缺失部件 (跳过 drawing1/3 的rels, 它们对应内容后面改名挂载) ----
add = [n for n in sorted(nb - nc) if not n.endswith('/')
       and n not in ('xl/drawings/_rels/drawing1.xml.rels', 'xl/drawings/_rels/drawing3.xml.rels',
                     'xl/worksheets/_rels/sheet6.xml.rels')]
for n in add:
    entries[n] = zb.read(n)
print(f'原样补齐 {len(add)} 个部件')

# ---- 2) 可视化按钮: 改名挂载 ----
plan = [  # (sheet, 备份drawing, 新drawing名)
    ('sheet6',  'drawing1', 'drawing26'),
    ('sheet11', 'drawing3', 'drawing21'),
    ('sheet13', 'drawing5', 'drawing23'),
    ('sheet15', 'drawing7', 'drawing24'),
    ('sheet17', 'drawing9', 'drawing25'),
]
CT_DRAW = 'application/vnd.openxmlformats-officedocument.drawing+xml'
CT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
for sheet, src, dst in plan:
    entries[f'xl/drawings/{dst}.xml'] = zb.read(f'xl/drawings/{src}.xml')
    r = f'xl/drawings/_rels/{src}.xml.rels'
    entries[f'xl/drawings/_rels/{dst}.xml.rels'] = (
        zb.read(r) if r in zb.namelist() else
        ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>').encode())
    # sheet rels: 移除旧 drawing 条目, 加新
    rn = f'xl/worksheets/_rels/{sheet}.xml.rels'
    if rn in entries:
        rx = entries[rn].decode('utf-8')
        rx = re.sub(r'<Relationship [^>]*relationships/drawing"[^>]*/>', '', rx)
    else:
        rx = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
    used = [int(i) for i in re.findall(r'Id="rId(\d+)"', rx)] or [0]
    rid = f'rId{max(used) + 1}'
    rx = rx.replace('</Relationships>',
                    f'<Relationship Id="{rid}" Type="{CT_REL}" Target="../drawings/{dst}.xml"/></Relationships>')
    entries[rn] = rx.encode('utf-8')
    # sheet xml: 替换/插入 <drawing>
    sx = entries[f'xl/worksheets/{sheet}.xml'].decode('utf-8')
    root_m = re.search(r'<worksheet[^>]*>', sx)
    if root_m and 'xmlns:r=' not in root_m.group(0):
        sx = sx.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ', 1)
    if '<drawing r:id=' in sx:
        sx = re.sub(r'<drawing r:id="rId\d+"/>', f'<drawing r:id="{rid}"/>', sx)
    elif '<extLst' in sx:
        sx = sx.replace('<extLst', f'<drawing r:id="{rid}"/><extLst', 1)
    else:
        sx = sx.replace('</worksheet>', f'<drawing r:id="{rid}"/></worksheet>')
    entries[f'xl/worksheets/{sheet}.xml'] = sx.encode('utf-8')
    print(f'{sheet}: {src} -> {dst}.xml (rId={rid})')

# ---- 3) ContentTypes 合并 ----
ct = entries['[Content_Types].xml'].decode('utf-8')
ctb = zb.read('[Content_Types].xml').decode('utf-8')
def decls(xml):
    d = dict(re.findall(r'<Default Extension="([^"]+)"[^>]*ContentType="([^"]+)"[^>]*/>', xml))
    o = dict(re.findall(r'<Override PartName="([^"]+)"[^>]*ContentType="([^"]+)"[^>]*/>', xml))
    return d, o
db, ob = decls(ctb)
dc, oc = decls(ct)
# 新名字的绘图: 类型声明套用备份里源绘图的声明
for sheet, src, dst in plan:
    spn, dpn = f'/xl/drawings/{src}.xml', f'/xl/drawings/{dst}.xml'
    ctype = ob.get(spn, CT_DRAW)
    if dpn not in oc:
        ct = ct.replace('</Types>', f'<Override PartName="{dpn}" ContentType="{ctype}"/></Types>')
        oc[dpn] = ctype
need_parts = [f'xl/drawings/{d}.xml' for _, _, d in plan] + \
             [f'xl/jdecontrols/jdecontrol{i}.xml' for i in range(1, 30)] + \
             ['xl/metadata.xml', 'xl/JDEData.bin'] + \
             [n for n in entries if n.startswith('xl/media/')]
for ext, ctype in db.items():
    if ext not in dc and any(n.endswith('.' + ext) for n in need_parts):
        ct = ct.replace('</Types>', f'<Default Extension="{ext}" ContentType="{ctype}"/></Types>')
        dc[ext] = ctype
for n in need_parts:
    pn = '/' + n
    if pn not in oc and pn in ob:
        ct = ct.replace('</Types>', f'<Override PartName="{pn}" ContentType="{ob[pn]}"/></Types>')
        oc[pn] = ob[pn]
    elif pn not in oc and pn.endswith('.xml') and pn not in ob:
        # 备份也没有显式声明的 xml 部件(如 jdecontrol 可能走 Default), 跳过
        pass
entries['[Content_Types].xml'] = ct.encode('utf-8')

# ---- 3.5) workbook.xml.rels: 恢复 WPS 宏工程/单元格图片/元数据 的引用 ----
rels_name = 'xl/_rels/workbook.xml.rels'
rx = entries[rels_name].decode('utf-8')
WPS = 'http://www.wps.cn/officeDocument'
wps_rels = [
    (f'{WPS}/2018/jdeExtension', 'JDEData.bin'),
    (f'{WPS}/2020/cellImage', 'cellimages.xml'),
    ('http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata', 'metadata.xml'),
]
used_ids = [int(i) for i in re.findall(r'Id="rId(\d+)"', rx)]
nid = max(used_ids) + 1
for rtype, target in wps_rels:
    if f'Target="{target}"' not in rx:
        rx = rx.replace('</Relationships>',
            f'<Relationship Id="rId{nid}" Type="{rtype}" Target="{target}"/></Relationships>')
        print(f'workbook.rels 新增: rId{nid} -> {target}')
        nid += 1
entries[rels_name] = rx.encode('utf-8')

# ---- 4) 写出并校验 ----
out = CUR + '.staging.xlsm'
zo = zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED)
for name, data in entries.items():
    zo.writestr(name, data)
zo.close()
zv = zipfile.ZipFile(out)
zn = set(zv.namelist())
checks = {
    '宏工程': 'xl/JDEData.bin' in zn,
    '按钮控件': all(f'xl/jdecontrols/jdecontrol{i}.xml' in zn for i in range(1, 30)),
    '图片': all(f'xl/media/image{i}.png' in zn for i in range(1, 33)),
    '可视化绘图': all(f'xl/drawings/{d}.xml' in zn for _, _, d in plan),
    'sheet接线': all(f'xl/worksheets/_rels/{s}.xml.rels' in zn for s, _, _ in plan),
}
zv.close()
print('校验:', checks)
assert all(checks.values()), checks
shutil.move(out, CUR)
print('一站式修复完成! 快照:', SAFETY)
zb.close(); zc.close()
