# -*- coding: utf-8 -*-
"""把 WPS 专属部件(JS宏工程/按钮控件/图片/绘图)从备份移植回当前考核表
openpyxl 不认识这些部件, 保存时会剥掉; 本脚本做 zip 级移植, 不碰任何数据。"""
import zipfile, shutil, re, os, sys

BAK = r'E:\WorkSpace\workbuddy\2026-09-18-16-38-15\backup_20260919\油耗考核10月_宏.xlsm'
CUR = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
SAFETY = r'E:\WorkSpace\workbuddy\2026-09-18-16-38-15\kaohe_before_surgery.xlsm'

# 0) 锁检查
try:
    f = open(CUR, 'r+b'); f.close()
except PermissionError:
    print('文件仍被占用, 请先关闭 WPS/Excel 中的考核表!'); sys.exit(1)

shutil.copy2(CUR, SAFETY)
zb, zc = zipfile.ZipFile(BAK), zipfile.ZipFile(CUR)
nb, nc = set(zb.namelist()), set(zc.namelist())
add_parts = [n for n in sorted(nb - nc) if not n.endswith('/')]
print('待移植部件:', len(add_parts))

# 1) Content_Types 合并: 从备份 CT 提取覆盖缺失部件的声明
ctb = zb.read('[Content_Types].xml').decode('utf-8')
ctc = zc.read('[Content_Types].xml').decode('utf-8')
defaults_b = dict(re.findall(r'<Default Extension="([^"]+)"[^>]*ContentType="([^"]+)"[^>]*/>', ctb))
defaults_c = dict(re.findall(r'<Default Extension="([^"]+)"[^>]*ContentType="([^"]+)"[^>]*/>', ctc))
overrides_b = dict(re.findall(r'<Override PartName="([^"]+)"[^>]*ContentType="([^"]+)"[^>]*/>', ctb))
overrides_c = dict(re.findall(r'<Override PartName="([^"]+)"[^>]*ContentType="([^"]+)"[^>]*/>', ctc))
new_defaults, new_overrides = [], []
for part in add_parts:
    ext = part.rsplit('.', 1)[-1].lower()
    pn = '/' + part
    if pn in overrides_b and pn not in overrides_c:
        new_overrides.append((pn, overrides_b[pn]))
    elif ext in defaults_b and ext not in defaults_c:
        new_defaults.append((ext, defaults_b[ext]))
    elif ext not in defaults_c and pn not in overrides_b:
        # 备份里也没有显式声明, 跳过(可能由默认覆盖)
        pass
print('新增 Default:', new_defaults)
print('新增 Override:', [(p, ct[:50]) for p, ct in new_overrides])
for ext, ct in new_defaults:
    if f'Extension="{ext}"' not in ctc:
        ctc = ctc.replace('</Types>', f'<Default Extension="{ext}" ContentType="{ct}"/></Types>')
for pn, ct in new_overrides:
    if f'PartName="{pn}"' not in ctc:
        ctc = ctc.replace('</Types>', f'<Override PartName="{pn}" ContentType="{ct}"/></Types>')

# 2) workbook.xml.rels: 补 jdeExtension/cellImage/sheetMetadata 三条
rels_c = zc.read('xl/_rels/workbook.xml.rels').decode('utf-8')
used_ids = set(re.findall(r'Id="rId(\d+)"', rels_c))
next_id = max(int(i) for i in used_ids) + 1
def add_rel(xml, rtype, target):
    global next_id, used_ids
    rid = f'rId{next_id}'
    while rid in used_ids:
        next_id += 1; rid = f'rId{next_id}'
    used_ids.add(rid)
    rel = f'<Relationship Id="{rid}" Type="{rtype}" Target="{target}"/>'
    xml = xml.replace('</Relationships>', rel + '</Relationships>')
    next_id += 1
    return xml, rid
WPS = 'http://www.wps.cn/officeDocument'
rels_c, _ = add_rel(rels_c, f'{WPS}/2018/jdeExtension', 'JDEData.bin')
rels_c, _ = add_rel(rels_c, f'{WPS}/2020/cellImage', 'cellimages.xml')
rels_c, _ = add_rel(rels_c, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata', 'metadata.xml')

# 3) 组装新 zip: 当前全部条目(CT/rels替换) + 缺失部件
out = CUR + '.staging.xlsm'
zo = zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED)
for item in zc.infolist():
    data = zc.read(item.filename)
    if item.filename == '[Content_Types].xml':
        data = ctc.encode('utf-8')
    elif item.filename == 'xl/_rels/workbook.xml.rels':
        data = rels_c.encode('utf-8')
    zo.writestr(item, data)
for part in add_parts:
    zo.writestr(part, zb.read(part))
zo.close(); zb.close(); zc.close()

# 4) 校验: 轻量检查(zip 级, 不加载整个工作簿)
zv = zipfile.ZipFile(out)
zn = set(zv.namelist())
ok = ('xl/JDEData.bin' in zn and 'xl/jdecontrols/jdecontrol1.xml' in zn
      and 'xl/media/image1.png' in zn and 'xl/workbook.xml' in zn
      and 'xl/worksheets/sheet1.xml' in zn)
ct_ok = 'vbaProject' in zv.read('[Content_Types].xml').decode('utf-8', 'ignore') or 'JDEData' in zv.read('[Content_Types].xml').decode('utf-8', 'ignore')
zv.close()
assert ok and ct_ok, (ok, ct_ok)
print('校验通过: 宏工程/控件/图片/内容类型齐全')

# 5) 替换原文件(先留手术前快照)
shutil.move(out, CUR)
print('移植完成! 快照: ', SAFETY)
