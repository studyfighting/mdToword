# -*- coding: utf-8 -*-
"""
docx_diff.py —— 两个 docx 的比对工具（合并了 _verify_docx 和 _strict_check）。

用法:
    python docx_diff.py 旧文件.docx 新文件.docx
        # 完整比对: 按第一个章标题对齐(跳过目录区), 输出带标题级的差异到 本目录/_diff.txt
    python docx_diff.py 旧文件.docx 新文件.docx --strict
        # 严格比对: 只比正文句子(忽略标题/目录/样式), 确认无正文丢失或改动
"""
import sys, os, zipfile, difflib
import xml.etree.ElementTree as ET

# 强制 UTF-8 输出，避免 Windows 控制台 GBK 乱码
try:
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
HEAD = {'2', '3', '4', '5', '6', '7', '8'}


def resolve_doc(arg):
    """docx 路径: 原样返回(相对路径按当前工作目录解析)"""
    return arg


def parse(fn):
    """返回 [(style, text), ...]，含所有段落"""
    z = zipfile.ZipFile(fn)
    root = ET.fromstring(z.read('word/document.xml'))
    out = []
    for p in root.iter(W + 'p'):
        style = ''
        pPr = p.find(W + 'pPr')
        if pPr is not None:
            ps = pPr.find(W + 'pStyle')
            if ps is not None:
                style = ps.get(W + 'val', '')
        txt = ''.join((t.text or '') for t in p.iter(W + 't'))
        out.append((style, txt))
    return out


def first_chapter_idx(seq):
    for i, (s, t) in enumerate(seq):
        if s == '2' and t.strip():
            return i
    return 0


def cmd_full(old, new):
    """完整比对(按第一个章标题对齐, 跳过目录区)"""
    a = parse(old)
    b = parse(new)
    la = ['[%s] %s' % (s, t) for s, t in a[first_chapter_idx(a):]]
    lb = ['[%s] %s' % (s, t) for s, t in b[first_chapter_idx(b):]]
    sm = difflib.SequenceMatcher(None, la, lb)
    added = removed = changed = 0
    lines = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal':
            continue
        if tag == 'replace':
            for k in range(max(i2 - i1, j2 - j1)):
                oa = la[i1 + k] if i1 + k < i2 else ''
                ob = lb[j1 + k] if j1 + k < j2 else ''
                lines.append('CHG\n - %s\n + %s' % (oa[:120], ob[:120]))
                changed += 1
        elif tag == 'delete':
            for k in range(i1, i2):
                lines.append('DEL - %s' % la[k][:120])
                removed += 1
        else:
            for k in range(j1, j2):
                lines.append('ADD + %s' % lb[k][:120])
                added += 1
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), '_diff.txt')
    with open(out, 'w', encoding='utf-8') as f:
        f.write('adds:%d dels:%d chgs:%d\n' % (added, removed, changed))
        f.write('\n'.join(lines))
    print('adds:%d dels:%d chgs:%d -> %s' % (added, removed, changed, out))


def _body_lines(fn):
    """正文句子列表: 跳过目录区与标题, 只留正文"""
    z = zipfile.ZipFile(fn)
    root = ET.fromstring(z.read('word/document.xml'))
    paras = list(root.iter(W + 'p'))
    start = 0
    for i, p in enumerate(paras):
        st, txt = '', ''
        pPr = p.find(W + 'pPr')
        if pPr is not None:
            ps = pPr.find(W + 'pStyle')
            if ps is not None:
                st = ps.get(W + 'val', '')
        txt = ''.join((t.text or '') for t in p.iter(W + 't')).strip()
        if st == '2' and txt:
            start = i
            break
    lines = []
    for p in paras[start:]:
        st = ''
        pPr = p.find(W + 'pPr')
        if pPr is not None:
            ps = pPr.find(W + 'pStyle')
            if ps is not None:
                st = ps.get(W + 'val', '')
        if st in HEAD:
            continue
        txt = ''.join((t.text or '') for t in p.iter(W + 't')).strip()
        if txt:
            lines.append(txt)
    return lines


def cmd_strict(old, new):
    a = _body_lines(old)
    b = _body_lines(new)
    print('旧文档正文段落:', len(a), ' 新文档正文段落:', len(b))
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    diffs = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal':
            continue
        for k in range(max(i2 - i1, j2 - j1)):
            oa = a[i1 + k] if i1 + k < i2 else '<<缺>>'
            ob = b[j1 + k] if j1 + k < j2 else '<<缺>>'
            diffs.append((oa[:60], ob[:60]))
    print('正文差异条数:', len(diffs))
    for d in diffs[:10]:
        print(' -', d[0])
        print(' +', d[1])


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    strict = '--strict' in sys.argv
    if len(args) < 2:
        print(__doc__)
        return 1
    old = resolve_doc(args[0])
    new = resolve_doc(args[1])
    if not os.path.exists(old) or not os.path.exists(new):
        print('文件不存在:', old, '|', new)
        return 1
    cmd_strict(old, new) if strict else cmd_full(old, new)
    return 0


if __name__ == '__main__':
    sys.exit(main())
