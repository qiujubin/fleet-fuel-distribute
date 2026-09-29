# -*- coding: utf-8 -*-
"""恢复四个可视化sheet的图片按钮:
备份 drawing1/3/5/7/9 -> 当前 drawing20~24(改名避免与openpyxl的1~4冲突),
为 sheet6/11/13/15/17 重建 rels 并插入 <drawing> 引用。数据零改动。"""
import zipfile, re, shutil, os, sys

BAK = r'E:\WorkSpace\workbuddy\2026-09-18-16-38-15\backup_20260919\油耗考核10月_宏.xlsm'
CUR = r'C:\Users\Jubin\Desktop\车辆维修，油耗\油耗考核10月_宏.xlsm'
SAFETY = r'E:\WorkSpace\workbuddy\2026-09-18-16-38-15\kaohe_before_btn_fix.xlsm'

# 目标sheet -> 备份drawing
plan = {'sheet6': 'drawing1', 'sheet11': 'drawing3', 'sheet13': 'drawing5',
        'sheet15': 'drawing7', 'sheet17': 'drawing9'}

try:
    f = open(CUR, 'r+b'); f.close()
except PermissionError:
    print('考核表仍被占用, 请先关闭 WPS!'); sys.exit(1)
shutil.copy2(CUR, SAFETY)

zb = zipfile.ZipFile(BAK)
zc = zipfile.ZipFile(CUR)
entries = {i.filename: zc.read(i.filename) for i in zc.infolist()}
zb_names = set(zb.namelist())

CT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'

# 1) 移植绘图内容(改名)
new_parts = {}
sheet_map = {}
for sheet, src in plan.items():
    dst = f'drawing{20 + list(plan).index(sheet)}'
    new_parts[f'xl/drawings/{dst}.xml'] = zb.read(f'xl/drawings/{src}.xml')
    rels_name = f'xl/drawings/_rels/{src}.xml.rels'
    if rels_name in zb_names:
        new_parts[f'xl/drawings/_rels/{dst}.xml.rels'] = zb.read(rels_name)
    else:
        # 无rels也生成空rels, 防悬空
        new_parts[f'xl/drawings/_rels/{dst}.xml.rels'] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
    sheet_map[sheet] = dst
print('绘图挂载:', sheet_map)

# 2) 每个目标 sheet: 重建 rels(保留非drawing项) + 插入 <drawing>
for sheet, dst in sheet_map.items():
    rels_name = f'xl/worksheets/_rels/{sheet}.xml.rels'
    rel_xml = entries.get(rels_name)
    if rel_xml is None:
        rel_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
    rel_xml = rel_xml.decode('utf-8') if isinstance(rel_xml, bytes) else rel_xml
    # 删掉旧的 drawing 类型条目
    rel_xml = re.sub(r'<Relationship [^>]*relationships/drawing"[^>]*/>', '', rel_xml)
    # 新 rId
    used = set(re.findall(r'Id="rId(\d+)"', rel_xml))
    nid = max([int(i) for i in used] or [0]) + 1
    rid = f'rId{nid}'
    rel_xml = rel_xml.replace('</Relationships>',
        f'<Relationship Id="{rid}" Type="{CT}" Target="../drawings/{dst}.xml"/></Relationships>')
    entries[rels_name] = rel_xml.encode('utf-8')
    # sheet XML: 替换或插入 <drawing>
    sx = entries[f'xl/worksheets/{sheet}.xml'].decode('utf-8')
    if '<drawing r:id=' in sx:
        sx = re.sub(r'<drawing r:id="rId\d+"/>', f'<drawing r:id="{rid}"/>', sx)
    else:
        anchor = '<extLst'
        if anchor in sx:
            sx = sx.replace(anchor, f'<drawing r:id="{rid}"/>' + anchor, 1)
        else:
            sx = sx.replace('</worksheet>', f'<drawing r:id="{rid}"/></worksheet>')
    entries[f'xl/worksheets/{sheet}.xml'] = sx.encode('utf-8')
    print(f'{sheet}: drawing -> {dst}.xml (rId={rid})')

# 3) ContentTypes: 为 drawing20~24 补 Override
ct = entries['[Content_Types].xml'].decode('utf-8')
for dst in sheet_map.values():
    pn = f'/xl/drawings/{dst}.xml'
    if pn not in ct:
        ct = ct.replace('</Types>', f'<Override PartName="{pn}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
entries['[Content_Types].xml'] = ct.encode('utf-8')

# 4) 写出
out = CUR + '.staging.xlsm'
zo = zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED)
for name, data in entries.items():
    zo.writestr(name, data)
for sheet, dst in sheet_map.items():
    pass
for name, data in new_parts.items():
    zo.writestr(name, data)
zo.close()

# 5) 校验
zv = zipfile.ZipFile(out)
zn = set(zv.namelist())
ok = all(f'xl/drawings/{d}.xml' in zn for d in sheet_map.values()) \
     and 'xl/JDEData.bin' in zn and 'xl/media/image1.png' in zn
# 每个 sheet 的 rels 都能解析出 drawing
for sheet, dst in sheet_map.items():
    rx = zv.read(f'xl/worksheets/_rels/{sheet}.xml.rels').decode('utf-8')
    assert f'{dst}.xml' in rx, (sheet, dst)
zv.close()
assert ok, '部件缺失'
shutil.move(out, CUR)
print('手术完成! 快照:', SAFETY)
zb.close(); zc.close()
