# -*- coding: utf-8 -*-
"""
docx_inspect.py —— 单个 docx 的检查工具（合并了 _outline / _check_sample / _audit 三个脚本）。

用法:
    python docx_inspect.py 文件.docx             # 综合检查: 结构 + 元素统计 + 关系审计
    python docx_inspect.py 文件.docx --outline    # 只打印标题大纲(标题1~7缩进)
    python docx_inspect.py 文件.docx --stats      # 只统计元素(段落/表格/图片)
    python docx_inspect.py 文件.docx --audit      # 只审计 zip 关系/内容类型一致性
"""
import sys, os, zipfile
import xml.etree.ElementTree as ET

# 强制 UTF-8 输出，避免 Windows 控制台 GBK 乱码
try:
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
REL = '{http://schemas.openxmlformats.org/package/2006/relationships}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def resolve_doc(arg):
    """docx 路径: 原样返回(相对路径按当前工作目录解析)"""
    return arg


def load(fn):
    z = zipfile.ZipFile(fn)
    root = ET.fromstring(z.read('word/document.xml'))
    names = set(z.namelist())
    return z, root, names


def xml_wellformed(z, names):
    for n in names:
        if n.endswith('.xml') or n.endswith('.rels'):
            ET.fromstring(z.read(n))
    return True


def cmd_outline(root):
    IND = {'2': '', '3': '  ', '4': '    ', '5': '      ', '6': '        ',
           '7': '          ', '8': '            '}
    print('--- 大纲 (标题1~7) ---')
    for p in root.iter(W + 'p'):
        style = ''
        pPr = p.find(W + 'pPr')
        if pPr is not None:
            ps = pPr.find(W + 'pStyle')
            if ps is not None:
                style = ps.get(W + 'val', '')
        if style in IND:
            txt = ''.join((t.text or '') for t in p.iter(W + 't'))
            if txt.strip():
                print(IND[style] + txt[:70])


def cmd_stats(z, root, names):
    print('parts:', sorted(names))
    try:
        xml_wellformed(z, names)
        print('XML 全部良构: OK')
    except Exception as e:
        print('XML 良构校验失败:', e)
    print('段落数:', len(list(root.iter(W + 'p'))),
          ' 表格数:', len(list(root.iter(W + 'tbl'))),
          ' 图片:', [n for n in names if n.startswith('word/media/')])


def cmd_audit(z, root, names):
    ok = True

    def check(cond, msg):
        nonlocal ok
        print(('PASS ' if cond else 'FAIL ') + msg)
        ok = ok and cond

    for part in [n for n in names if n.endswith('.rels')]:
        r = ET.fromstring(z.read(part))
        base = part.split('/_rels/')[0]
        if base == part:
            base = ''
        for rel in r:
            target = rel.get('Target')
            if rel.get('TargetMode') == 'External':
                continue
            full = target.lstrip('/') if target.startswith('/') else (
                base + '/' + target if base else target)
            check(full in names, f'{part} -> {rel.get("Id")}: {target} 存在')

    ct = z.read('[Content_Types].xml').decode('utf-8')
    for name in names:
        if name.startswith('word/media/'):
            ext = name.rsplit('.', 1)[1]
            check('Extension="%s"' % ext in ct, f'[Content_Types] 声明 {ext}')
        if name in ('word/document.xml', 'word/styles.xml', 'word/settings.xml'):
            check('PartName="/%s"' % name in ct, f'[Content_Types] Override {name}')

    rels = ET.fromstring(z.read('word/_rels/document.xml.rels'))
    rid_map = {r.get('Id'): r.get('Target') for r in rels}
    for el in root.iter():
        for attr in ('{%s}embed' % R, '{%s}id' % R):
            v = el.get(attr)
            if v:
                tgt = rid_map.get(v)
                check(tgt is not None and ('word/' + tgt) in names,
                      f'body {attr}={v} -> {tgt} 存在')
    print('=== AUDIT', 'PASS' if ok else 'FAIL', '===')


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    flags = set(a for a in sys.argv[1:] if a.startswith('--'))
    if not args:
        print(__doc__)
        return 1
    fn = resolve_doc(args[0])
    if not os.path.exists(fn):
        print('文件不存在:', fn)
        return 1
    z, root, names = load(fn)
    if flags:
        if '--outline' in flags:
            cmd_outline(root)
        if '--stats' in flags:
            cmd_stats(z, root, names)
        if '--audit' in flags:
            cmd_audit(z, root, names)
    else:
        cmd_stats(z, root, names)
        print()
        cmd_outline(root)
        print()
        cmd_audit(z, root, names)
    return 0


if __name__ == '__main__':
    sys.exit(main())
