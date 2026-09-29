# -*- coding: utf-8 -*-
"""修正 SKILL.md 中被八进制转义破坏的路径"""
p = r'C:\Users\Jubin\.workbuddy\skills\fleet-fuel-daily-distribute\SKILL.md'
s = open(p, encoding='utf-8').read()
bad = chr(0x82) + '6-09-18-16-38-15'
good = '2026-09-18-16-38-15'
if bad in s:
    s = s.replace(bad, good)
    open(p, 'w', encoding='utf-8').write(s)
    print('已修正损坏的八进制转义')
else:
    print('未发现损坏')

# 全文件体检: 是否还有非预期控制字符
bad_chars = [(i, hex(ord(ch))) for i, ch in enumerate(s) if ord(ch) < 32 and ch not in '\n\r\t']
print('控制字符:', bad_chars if bad_chars else '无')
i = s.find('文档维护约定')
print()
print('修正后内容:')
print(s[i - 5:i + 260])
