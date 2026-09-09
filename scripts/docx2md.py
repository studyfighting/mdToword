# -*- coding: utf-8 -*-
"""
docx2md.py —— docx → Markdown 逆转换器（用于把已有 Word 文档转成可持续维护的 .md 源）。
纯标准库实现。

用途: 把 .docx 转成 md2docx.py 能读的 Markdown 源，从而进入 "md 源 + 一键导 Word" 工作流。

用法:
    python docx2md.py 输入.docx [输出.md]
    # 默认: 读 输入.docx, 写 同名 .md(同目录)

处理规则:
    - 标题样式 2/3/4/5/6/7/8 -> #/##/###/####/#####/######/######
    - 标题文字里已有的编号(一、/ 1. / 1.1)原样保留
    - 正文段 -> 普通段落(自动跳过文档开头的目录占位区)
    - 含 drawing 的段 -> 提取图片到输出目录 + 写 ![](图片名)
    - 表格 -> markdown 表格
    - 代码段(以 #define/#if/#include/**/ 等开头) -> ```c 代码块
    - 不生成目录；md 里也不写目录。
"""
import sys, os, re, zipfile, shutil
import xml.etree.ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'


# 标题样式 -> md 级别
STYLE_HEAD = {'2': '#', '3': '##', '4': '###', '5': '####',
              '6': '#####', '7': '######', '8': '######'}

# 代码块行首特征(C 风格)
CODE_PAT = re.compile(
    r'^\s*(?:#\s*(?:include|if|ifdef|ifndef|define|endif|pragma)|'
    r'/\*|(?:\*/|/\*\*)|uint(?:8|16|32|64)_t|int(?:8|16|32|64)_t|'
    r'void\s+\w+\s*\(|struct\s+\w+|typedef\s+|const\s+\w+\s+\w+\s*=|'
    r'\w+\s+\w+\s*=\s*[^;]+;|return\s+[^;]+;|case\s+\w+\s*:)')


def ptxt(p):
    """段落文本(拼接所有 w:t)"""
    return ''.join((t.text or '') for t in p.iter(W + 't'))


def pstyle(p):
    pPr = p.find(W + 'pPr')
    if pPr is None:
        return ''
    ps = pPr.find(W + 'pStyle')
    return ps.get(W + 'val', '') if ps is not None else ''


def cell_text(tc):
    """表格单元格文本"""
    return ''.join((t.text or '') for t in tc.iter(W + 't')).strip()


def tbl_to_md(tbl):
    rows = []
    for tr in tbl.iter(W + 'tr'):
        cells = [cell_text(tc) for tc in tr.findall(W + 'tc')]
        rows.append(cells)
    if not rows:
        return []
    ncols = max(len(r) for r in rows)
    # 第一行作表头, 第二行作分隔(若原本第一行就是表头)
    lines = []
    header = rows[0]
    lines.append('| ' + ' | '.join(header) + ' |')
    lines.append('|' + '|'.join([' --- '] * ncols) + '|')
    for r in rows[1:]:
        padded = list(r) + [''] * (ncols - len(r))
        lines.append('| ' + ' | '.join(padded) + ' |')
    return lines


def is_drawing_para(p):
    return p.find('.//' + W + 'drawing') is not None


def is_code_line(txt):
    return bool(CODE_PAT.match(txt)) and len(txt) < 200


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    src = sys.argv[1]
    if not os.path.exists(src):
        print('输入 docx 不存在:', src)
        return 1
    out = sys.argv[2] if len(sys.argv) > 2 else os.path.splitext(src)[0] + '.md'
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)

    z = zipfile.ZipFile(src)
    root = ET.fromstring(z.read('word/document.xml'))
    body = root.find(W + 'body')

    # 建立 rId -> 图片文件 映射(从 rels 读), 并记录已提取的图片
    rels = ET.fromstring(z.read('word/_rels/document.xml.rels'))
    rid_img = {}
    for rel in rels:
        t = rel.get('Target')
        if t.startswith('media/'):
            rid_img[rel.get('Id')] = t

    img_counter = 0
    used_img = {}          # relId -> 本地文件名
    lines = []
    # body 顶层顺序遍历: p / tbl
    items = [ch for ch in body if ch.tag.split('}')[-1] in ('p', 'tbl')]

    # 跳过开头目录占位区: 直到遇到第一个标题(样式2)才真正开始
    started = False
    code_buf = []
    img_final_dir = os.path.dirname(os.path.abspath(out))

    def flush_code():
        nonlocal code_buf
        if code_buf:
            # 去掉代码块末尾多余空行
            while code_buf and code_buf[-1] == '':
                code_buf.pop()
            lines.append('```c')
            lines.extend(code_buf)
            lines.append('```')
            code_buf = []

    for ch in items:
        tag = ch.tag.split('}')[-1]
        if tag == 'tbl':
            flush_code()
            lines.append('')
            lines.extend(tbl_to_md(ch))
            lines.append('')
            continue

        # 图片段: 提取图片 -> output 目录
        if is_drawing_para(ch):
            flush_code()
            # 找该段所有 embed rId
            rids = [b.get(R + 'embed') for b in ch.iter()
                    if b.tag.split('}')[-1] == 'blip']
            rids = [x for x in rids if x]
            for rid in rids:
                media = rid_img.get(rid)
                if not media:
                    continue
                ext = os.path.splitext(media)[1]
                fname = 'img_%d%s' % (img_counter, ext)
                img_counter += 1
                data = z.read('word/' + media)
                with open(os.path.join(img_final_dir, fname), 'wb') as f:
                    f.write(data)
                lines.append('![](%s)' % fname)
                lines.append('')
            continue

        style = pstyle(ch)
        txt = ptxt(ch)

        # 标题
        if style in STYLE_HEAD:
            flush_code()
            lines.append('')
            lines.append('%s %s' % (STYLE_HEAD[style], txt.strip()))
            lines.append('')
            if not started:
                started = True
            continue

        # 未开始(目录区)直接跳过
        if not started:
            continue

        # 代码块行/空段: 处于代码块中时空段保留为空行; 非代码正文段才 flush
        if is_code_line(txt):
            code_buf.append(txt.rstrip())
            continue
        if not txt.strip():
            if code_buf:
                code_buf.append('')          # 代码块内部空行
            continue

        flush_code()
        # 普通正文(每段后加空行, 保证 md 段落边界 -> 回流时各成一段)
        lines.append(txt.rstrip())
        lines.append('')

    flush_code()

    with open(out, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines).strip() + '\n')
    print('OK ->', os.path.abspath(out))
    print('图片数:', img_counter)
    return 0


if __name__ == '__main__':
    sys.exit(main())
