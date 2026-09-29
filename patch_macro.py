# -*- coding: utf-8 -*-
"""把 macro_sync_v2.js 替换进考核表的 JSA 宏工程(JDEData.bin)
用法:
  python patch_macro.py <考核表路径> [dstPath覆盖]
  不带 dstPath 参数 = 部署到真实考核表(保持真实 dstPath)
"""
import sys, os, zipfile, re, html, shutil

BIN = 'xl/JDEData.bin'

def esc(s):
    """按原文件转义规则: & < > " 和换行(-> &#x0A;)"""
    return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
             .replace('"', '&quot;')
             .replace('\r\n', '&#x0A;').replace('\n', '&#x0A;').replace('\r', '&#x0A;'))

def replace_macro(bin_data: bytes, new_js: str, dst_override=None):
    xml = bin_data.decode('utf-8')
    if dst_override:
        new_js = new_js.replace('C:/Users/Jubin/Desktop/源数据/油耗分析汇总.xlsm', dst_override)
    start = xml.find('function 同步油耗到汇总')
    assert start >= 0, '找不到旧宏函数头'
    # 在转义文本上配对大括号(代码字符串里无大括号)
    i = xml.find('{', start)
    depth = 0
    end = None
    for j in range(i, len(xml)):
        if xml[j] == '{': depth += 1
        elif xml[j] == '}':
            depth -= 1
            if depth == 0: end = j + 1; break
    assert end, '大括号配对失败'
    old_len = end - start
    xml = xml[:start] + esc(new_js) + xml[end:]
    print(f'旧函数体 {old_len} 字符 -> 新 {len(esc(new_js))} 字符')
    return xml.encode('utf-8')

def main():
    xlsx = sys.argv[1]
    dst_override = sys.argv[2] if len(sys.argv) > 2 else None
    new_js = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'macro_sync_v2.js'),
                  encoding='utf-8').read()
    if os.environ.get('TEST_SKIP_STEP5'):
        # 测试模式: 区域填充步骤置空(该步骤曾疑似在后台实例中卡死)
        old5 = """    try {
        Application.Run("'" + dstPath + "'!根据车牌号填充区域");
        runMsg += "OK";
    } catch (e1) {
        try {
            // 容错:如果宏名少了"号"
            Application.Run("'" + dstPath + "'!根据车牌填充区域");
            runMsg += "OK(根据车牌填充区域)";
        } catch (e2) {
            runErr = (e1.message || e1) + " / " + (e2.message || e2);
            runMsg += "失败 — 目标文件没有根据车牌号填充区域宏,请先把上一个 JSA 注入";
        }
    }"""
        assert old5 in new_js, '第5步代码块未匹配'
        new_js = new_js.replace(old5, '    runMsg = "区域填充: (测试模式跳过)";')
        print('测试模式: 第5步(区域填充)已置空')
    # 快照
    shutil.copy2(xlsx, xlsx + '.pre_macro_patch.bak')
    zin = zipfile.ZipFile(xlsx)
    entries = {i.filename: zin.read(i.filename) for i in zin.infolist()}
    zin.close()
    entries[BIN] = replace_macro(entries[BIN], new_js, dst_override)
    out = xlsx + '.staging.xlsm'
    zo = zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED)
    for name, data in entries.items():
        zo.writestr(name, data)
    zo.close()
    os.replace(out, xlsx)
    # 校验: 新宏文本确实在 bin 里
    z = zipfile.ZipFile(xlsx)
    txt = z.read(BIN).decode('utf-8')
    z.close()
    marker = '跳过重复' if dst_override is None else html.escape('跳过重复')
    txt_un = html.unescape(txt)
    assert '跳过重复' in txt_un and 'lastDataRow' in txt_un, '宏替换校验失败'
    print(f'宏已替换进 {os.path.basename(xlsx)} (dstPath={"测试" if dst_override else "真实"})')

if __name__ == '__main__':
    main()
