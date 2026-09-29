# -*- coding: utf-8 -*-
p = r'C:\Users\Jubin\.workbuddy\skills\fleet-fuel-daily-distribute\SKILL.md'
s = open(p, encoding='utf-8').read()
BS = chr(92)
bad = 'workbuddy2026-09-18-16-38-15'
good = 'workbuddy' + BS + '2026-09-18-16-38-15'
if bad in s:
    s = s.replace(bad, good)
    open(p, 'w', encoding='utf-8').write(s)
    print('已补回路径分隔符')
else:
    print('未找到待修文本')

i = s.find('文档维护约定')
print()
print(s[i - 5:i + 240])
# 顺带核对 README.md 里引用的 skill 路径是否正确
r = open('README.md', encoding='utf-8').read()
j = r.find('AI 操作手册')
print('README 引用行:', r[j - 20:j + 130].strip())
