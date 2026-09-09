# -*- coding: utf-8 -*-
"""
md2docx.py —— 把 Markdown 源文件转成带章节/表格/图片/代码块的 Word 文档。
纯 Python 标准库实现，不依赖 pandoc / python-docx。

用法:
    python md2docx.py 输入.md [输出.docx]
    # 输出缺省为输入同名 .docx（与输入同一目录）

支持的 Markdown 语法:
    # ~ ######          标题(对应 Word 标题1~6，样式建好，不自动生成目录)
    **加粗** `行内代码`  行内格式
    | a | b | ---       表格(第二行须为 |---| 分隔行)
    ![](图片.png)       插入本地图片(自动缩放到合适宽度、居中, 相对 .md 所在目录解析)
    ``` ... ```         代码块(等宽字体、灰底)
    <!-- pagebreak -->  手动分页符
    普通段落            正文；- / 1. 列表会导出为带缩进的列表
说明: 章节编号请直接写在标题文字里(如 "## 1. SoC和MCU架构")，
     目录不自动生成；如需目录在 Word 里用「引用 → 目录」自建。
"""
import sys, os, re, zipfile, struct
import xml.etree.ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'

for prefix, uri in (('w', W), ('r', R), ('wp', WP), ('a', A), ('pic', PIC)):
    ET.register_namespace(prefix, uri)

XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'


def wn(tag):
    return '{%s}%s' % (W, tag)


def mk_par(style=None, jc=None, keep=False):
    p_ = ET.Element(wn('p'))
    if style or jc or keep:
        pPr = ET.SubElement(p_, wn('pPr'))
        if style:
            e = ET.SubElement(pPr, wn('pStyle'))
            e.set(wn('val'), style)
        if jc:
            ET.SubElement(pPr, wn('jc')).set(wn('val'), jc)
        if keep:
            ET.SubElement(pPr, wn('keepNext'))
    return p_


def add_text(par, text):
    """行内: **加粗** `代码` !!红色提示!!; 行尾两空格(\\n) 处插入 <w:br/> 硬换行(md 标记触发)"""
    # 按换行符切行, 每行内处理加粗/代码/红色提示; 行与行之间是 md 硬换行 -> <w:br/>
    for li, line in enumerate(text.split('\n')):
        if li > 0:
            br_r = ET.SubElement(par, wn('r'))
            ET.SubElement(br_r, wn('br'))
        token = re.compile(r'(\*\*.+?\*\*|`[^`]+`|!![^!]+!!)')
        for part in token.split(line):
            if not part:
                continue
            r = ET.SubElement(par, wn('r'))
            bold = part.startswith('**') and part.endswith('**')
            code = (not bold) and part.startswith('`') and part.endswith('`')
            red = (not bold and not code) and part.startswith('!!') and part.endswith('!!')
            if bold or code or red:
                rPr = ET.SubElement(r, wn('rPr'))
                if bold:
                    ET.SubElement(rPr, wn('b'))
                    ET.SubElement(rPr, wn('bCs'))
                if code:
                    rf = ET.SubElement(rPr, wn('rFonts'))
                    rf.set(wn('ascii'), 'Consolas')
                    rf.set(wn('hAnsi'), 'Consolas')
                    ET.SubElement(rPr, wn('sz')).set(wn('val'), '18')
                    ET.SubElement(rPr, wn('szCs')).set(wn('val'), '18')
                if red:
                    ET.SubElement(rPr, wn('b'))
                    ET.SubElement(rPr, wn('bCs'))
                    col = ET.SubElement(rPr, wn('color'))
                    col.set(wn('val'), 'C00000')   # 醒目红
            body = part[2:-2] if bold else (part[1:-1] if code else (part[2:-2] if red else part))
            t = ET.SubElement(r, wn('t'))
            t.set(XML_SPACE, 'preserve')
            t.text = body


def text_par(text, jc=None, style=None, keep=False, num=None, bullet=None):
    """构造普通段落。num 给有序列表编号, bullet 给无序列表项目符号。
    段落内保留 md 换行标记(行尾两空格)为 <w:br/>。"""
    par = mk_par(style=style, jc=jc, keep=keep)
    pPr = par.find(wn('pPr'))
    if pPr is None:
        pPr = ET.Element(wn('pPr'))
        par.insert(0, pPr)
    # 列表：缩进 + 手动序号/符号(不依赖 numbering.xml, 兼容性好且视觉即列表)
    if num is not None or bullet:
        ind = ET.SubElement(pPr, wn('ind'))
        ind.set(wn('left'), '560')
        ind.set(wn('hanging'), '360')
        prefix = ('%s. ' % num) if num is not None else (bullet + ' ')
        text = prefix + text
    add_text(par, text)
    return par


def make_table(rows):
    ncols = max(len(r) for r in rows)
    tbl = ET.Element(wn('tbl'))
    tblPr = ET.SubElement(tbl, wn('tblPr'))
    w = ET.SubElement(tblPr, wn('tblW'))
    w.set(wn('type'), 'pct')
    w.set(wn('w'), '5000')
    borders = ET.SubElement(tblPr, wn('tblBorders'))
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        b = ET.SubElement(borders, wn(side))
        b.set(wn('val'), 'single')
        b.set(wn('sz'), '4')
        b.set(wn('color'), '808080')
    grid = ET.SubElement(tbl, wn('tblGrid'))
    for _ in range(ncols):
        gc = ET.SubElement(grid, wn('gridCol'))
        gc.set(wn('w'), str(int(5000 / ncols)))
    for ri, row in enumerate(rows):
        tr = ET.SubElement(tbl, wn('tr'))
        for ci in range(ncols):
            cell = row[ci] if ci < len(row) else ''
            tc = ET.SubElement(tr, wn('tc'))
            tcPr = ET.SubElement(tc, wn('tcPr'))
            tcw = ET.SubElement(tcPr, wn('tcW'))
            tcw.set(wn('type'), 'pct')
            tcw.set(wn('w'), str(int(5000 / ncols)))
            if ri == 0:
                shd = ET.SubElement(tcPr, wn('shd'))
                shd.set(wn('val'), 'clear')
                shd.set(wn('fill'), 'D9E2F3')
            if ri == 0:
                par = ET.SubElement(tc, wn('p'))
                r = ET.SubElement(par, wn('r'))
                rPr = ET.SubElement(r, wn('rPr'))
                ET.SubElement(rPr, wn('b'))
                t = ET.SubElement(r, wn('t'))
                t.set(XML_SPACE, 'preserve')
                t.text = cell
            else:
                par = ET.SubElement(tc, wn('p'))
                add_text(par, cell)
    return tbl


def image_wh(data):
    if data[:8] == b'\x89PNG\r\n\x1a\n':
        return struct.unpack('>II', data[16:24])
    if data[:2] == b'\xff\xd8':
        i, n = 2, len(data)
        while i < n:
            if data[i] != 0xFF:
                i += 1
                continue
            m = data[i + 1]
            if m in (0xC0, 0xC1, 0xC2, 0xC3):
                h, w = struct.unpack('>HH', data[i + 5:i + 9])
                return w, h
            i += 2 + struct.unpack('>H', data[i + 2:i + 4])[0]
    return 800, 600


def split_row(line):
    s = line.strip()
    if not s.startswith('|'):
        s = '|' + s + '|'
    return [c.strip() for c in s.strip('|').split('|')]


SEP = re.compile(r':?-{2,}:?')
PAGEBREAK = ('<!-- pagebreak -->', '\\newpage')


class Builder:
    """收集块，末尾统一打包"""

    def __init__(self, toc_depth=3):
        self.blocks = []
        self.media = []
        self.toc_depth = toc_depth
    def heading(self, level, text):
        # styleId: 2=标题1 .. 7=标题6 (1=Normal)
        self.blocks.append(text_par(text, style=str(min(level + 1, 7))))

    def para(self, text):
        self.blocks.append(text_par(text))

    def list_item(self, text, num=None, bullet=None):
        """列表项(有序 num=1/2.., 无序 bullet='•'/'·'), 带缩进"""
        self.blocks.append(text_par(text, num=num, bullet=bullet))

    def gap(self):
        """视觉缓冲：一个空行段落，用于块与块之间留白"""
        self.blocks.append(text_par(''))

    def pagebreak_before(self):
        """分页符段落：用于章标题独占一页开头"""
        par = ET.Element(wn('p'))
        r = ET.SubElement(par, wn('r'))
        br = ET.SubElement(r, wn('br'))
        br.set(wn('type'), 'page')
        self.blocks.append(par)

    def code(self, lines):
        par = mk_par()
        pPr = ET.Element(wn('pPr'))
        par.insert(0, pPr)
        shd = ET.SubElement(pPr, wn('shd'))
        shd.set(wn('val'), 'clear')
        shd.set(wn('fill'), 'F2F2F2')
        ind = ET.SubElement(pPr, wn('ind'))
        ind.set(wn('left'), '240')
        for i, line in enumerate(lines):
            r = ET.SubElement(par, wn('r'))
            rPr = ET.SubElement(r, wn('rPr'))
            rf = ET.SubElement(rPr, wn('rFonts'))
            rf.set(wn('ascii'), 'Consolas')
            rf.set(wn('hAnsi'), 'Consolas')
            ET.SubElement(rPr, wn('sz')).set(wn('val'), '18')
            ET.SubElement(rPr, wn('szCs')).set(wn('val'), '18')
            t = ET.SubElement(r, wn('t'))
            t.set(XML_SPACE, 'preserve')
            t.text = line
            if i < len(lines) - 1:
                ET.SubElement(r, wn('br'))
        self.blocks.append(par)

    def table(self, rows):
        self.blocks.append(make_table(rows))

    def image(self, path, base_dir=None):
        if base_dir and not os.path.isabs(path):
            p = os.path.join(base_dir, path)
            if os.path.exists(p):
                path = p
        if not os.path.exists(path):
            self.blocks.append(text_par('[图片缺失: %s]' % path))
            return
        data = open(path, 'rb').read()
        ext = 'png' if data[:8] == b'\x89PNG\r\n\x1a\n' else \
              ('jpg' if data[:2] == b'\xff\xd8' else 'bin')
        w, h = image_wh(data)
        target = 13 * 360000          # 最长边 ~13cm
        if w >= h:
            cx, cy = target, max(1, int(target * h / w))
        else:
            cy, cx = target, max(1, int(target * w / h))
        self.media.append((data, ext))
        idx = len(self.media)
        rid = 'rIdImg%d' % idx

        par = mk_par(jc='center')
        r = ET.SubElement(par, wn('r'))
        drawing = ET.SubElement(r, wn('drawing'))
        inline = ET.SubElement(drawing, '{%s}inline' % WP)
        ext_el = ET.SubElement(inline, '{%s}extent' % WP)
        ext_el.set('cx', str(cx))
        ext_el.set('cy', str(cy))
        docPr = ET.SubElement(inline, '{%s}docPr' % WP)
        docPr.set('id', str(idx))
        docPr.set('name', 'Picture %d' % idx)
        graphic = ET.SubElement(inline, '{%s}graphic' % A)
        gdata = ET.SubElement(graphic, '{%s}graphicData' % A)
        gdata.set('uri', PIC)
        pic = ET.SubElement(gdata, '{%s}pic' % PIC)
        nv = ET.SubElement(pic, '{%s}nvPicPr' % PIC)
        cNvPr = ET.SubElement(nv, '{%s}cNvPr' % PIC)
        cNvPr.set('id', str(idx))
        cNvPr.set('name', 'Picture %d' % idx)
        ET.SubElement(nv, '{%s}cNvPicPr' % PIC)
        blipFill = ET.SubElement(pic, '{%s}blipFill' % PIC)
        blip = ET.SubElement(blipFill, '{%s}blip' % A)
        blip.set('{%s}embed' % R, rid)
        stretch = ET.SubElement(blipFill, '{%s}stretch' % A)
        ET.SubElement(stretch, '{%s}fillRect' % A)
        spPr = ET.SubElement(pic, '{%s}spPr' % PIC)
        xfrm = ET.SubElement(spPr, '{%s}xfrm' % A)
        off = ET.SubElement(xfrm, '{%s}off' % A)
        off.set('x', '0')
        off.set('y', '0')
        ext2 = ET.SubElement(xfrm, '{%s}ext' % A)
        ext2.set('cx', str(cx))
        ext2.set('cy', str(cy))
        geom = ET.SubElement(spPr, '{%s}prstGeom' % A)
        geom.set('prst', 'rect')
        ET.SubElement(geom, '{%s}avLst' % A)
        self.blocks.append(par)

    def pagebreak(self):
        par = ET.Element(wn('p'))
        r = ET.SubElement(par, wn('r'))
        br = ET.SubElement(r, wn('br'))
        br.set(wn('type'), 'page')
        self.blocks.append(par)

    # ---------- 打包 ----------
    def build(self, out_path):
        doc = ET.Element(wn('document'))
        body = ET.SubElement(doc, wn('body'))

        # 不生成目录；文档开头即第一个标题。若需要目录，在 Word 里用“引用→目录”自建。

        for blk in self.blocks:
            body.append(blk)

        sectPr = ET.SubElement(body, wn('sectPr'))
        pgSz = ET.SubElement(sectPr, wn('pgSz'))
        pgSz.set(wn('w'), '11906')
        pgSz.set(wn('h'), '16838')
        pgMar = ET.SubElement(sectPr, wn('pgMar'))
        for k, v in (('top', '1418'), ('right', '1418'), ('bottom', '1418'),
                     ('left', '1418'), ('header', '708'), ('footer', '708'),
                     ('gutter', '0')):
            pgMar.set(wn(k), v)

        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', self._content_types())
            z.writestr('_rels/.rels', self._root_rels())
            z.writestr('word/document.xml',
                       ET.tostring(doc, encoding='UTF-8', xml_declaration=True))
            z.writestr('word/styles.xml', self._styles_xml())
            z.writestr('word/settings.xml', self._settings_xml())
            z.writestr('word/_rels/document.xml.rels', self._doc_rels())
            for i, (data, ext) in enumerate(self.media, 1):
                z.writestr('word/media/image%d.%s' % (i, ext), data)

    def _content_types(self):
        parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                 '<Types xmlns="%s">' % CT,
                 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
                 '<Default Extension="xml" ContentType="application/xml"/>',
                 '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
                 '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
                 '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>']
        seen = set()
        for _, ext in self.media:
            if ext in seen:
                continue
            seen.add(ext)
            ctype = {'png': 'image/png', 'jpg': 'image/jpeg'}.get(ext,
                    'application/octet-stream')
            parts.append('<Default Extension="%s" ContentType="%s"/>' % (ext, ctype))
        parts.append('</Types>')
        return ''.join(parts).encode('utf-8')

    @staticmethod
    def _root_rels():
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="%s">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                '</Relationships>' % REL).encode('utf-8')

    def _doc_rels(self):
        parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                 '<Relationships xmlns="%s">' % REL,
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
                 '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>']
        for i, (_, ext) in enumerate(self.media, 1):
            parts.append('<Relationship Id="rIdImg%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image%d.%s"/>'
                         % (i, i, ext))
        parts.append('</Relationships>')
        return ''.join(parts).encode('utf-8')

    @staticmethod
    def _settings_xml():
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:settings xmlns:w="%s"></w:settings>' % W
                ).encode('utf-8')

    @staticmethod
    def _styles_xml():
        parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                 '<w:styles xmlns:w="%s">' % W,
                 '<w:style w:type="paragraph" w:styleId="1">'
                 '<w:name w:val="Normal"/><w:qFormat/>'
                 '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>'
                 '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style>']
        for sid in range(2, 8):  # 2..7 = 标题1..6
            sz = (32, 28, 24, 22, 22, 22)[sid - 2]
            parts.append(
                '<w:style w:type="paragraph" w:styleId="%d">'
                '<w:name w:val="heading %d"/><w:basedOn w:val="1"/><w:qFormat/>'
                '<w:pPr><w:keepNext/><w:keepLines/>'
                '<w:spacing w:before="240" w:after="120" w:line="360" w:lineRule="auto"/>'
                '<w:outlineLvl w:val="%d"/></w:pPr>'
                '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="黑体"/>'
                '<w:b/><w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr></w:style>'
                % (sid, sid - 1, sid - 2, sz, sz))
        parts.append('</w:styles>')
        return ''.join(parts).encode('utf-8')


def parse_md(text, b, base_dir=None):
    lines = text.splitlines()
    i, n = 0, len(lines)
    para = []
    code = None
    table = None  # 列表[行文本]
    listbuf = None  # 连续列表项缓冲: (type, [(num_or_bullet, text), ...])
    last_kind = None      # 上一个已输出块的类型
    heading_count = 0     # 已输出的一级标题(章)个数

    ORDER_RE = re.compile(r'^\s*\d+\.\s+(.*)$')     # 1. xxx
    BULLET_RE = re.compile(r'^\s*[-*+]\s+(.*)$')    # - xxx / * xxx

    def gap_between(kind):
        """块类型改变时插入视觉缓冲空行。heading/代码块间不插(自带间距)。
        用一个惰性标记确保相邻块之间最多一个空行, 不因图片空段等重复叠加。"""
        nonlocal last_kind
        if last_kind and last_kind != kind:
            if kind not in ('heading',) and last_kind not in ('heading',):
                b.gap()
        last_kind = kind

    def flush_para():
        nonlocal para
        if para:
            txt = ' '.join(x.strip() for x in para if x.strip())
            para.clear()
            if txt:
                gap_between('para')
                b.para(txt)

    def flush_list():
        nonlocal listbuf
        if listbuf:
            typ, items = listbuf
            if typ == 'ordered':
                gap_between('list')
                for k, (_, txt) in enumerate(items, 1):
                    b.list_item(txt, num=k)
            else:
                gap_between('list')
                for _, txt in items:
                    b.list_item(txt, bullet='•')
            listbuf = None
        else:
            pass

    def flush_table():
        nonlocal table
        if table is None:
            return
        rows = [split_row(r) for r in table]
        header = rows[0]
        rest = [r for r in rows[1:] if not all(SEP.fullmatch(c.replace(' ', '')) for c in r)]
        gap_between('table')
        b.table([header] + rest)
        table = None

    def resolve(path):
        if base_dir and not os.path.isabs(path):
            p = os.path.join(base_dir, path)
            if os.path.exists(p):
                return p
        return path

    while i < n:
        line = lines[i]
        s = line.strip()
        if code is not None:
            if s.startswith('```'):
                gap_between('code')
                b.code(code)
                code = None
            else:
                code.append(line)
            i += 1
            continue
        if s.startswith('```'):
            flush_para(); flush_list(); flush_table()
            code = []
            i += 1
            continue
        if s in PAGEBREAK:
            flush_para(); flush_list(); flush_table()
            b.pagebreak()
            i += 1
            continue
        m = re.match(r'^(#{1,6})\s+(.*)$', line)
        if m:
            flush_para(); flush_list(); flush_table()
            lvl = len(m.group(1))
            # 每个大节(章, #)标题前分页, 独占一页开头(第一个章标题除外)
            if lvl == 1 and heading_count > 0:
                b.pagebreak_before()
            if lvl == 1:
                heading_count += 1
            gap_between('heading')
            b.heading(lvl, m.group(2).strip())
            i += 1
            continue
        img = re.match(r'^!\[([^\]]*)\]\(([^)]+)\)\s*$', s)
        if img:
            flush_para(); flush_list(); flush_table()
            gap_between('image')
            b.image(resolve(img.group(2)))
            i += 1
            continue
        if s.startswith('|'):
            flush_para(); flush_list()
            if table is None:
                table = []
            table.append(line)
            i += 1
            continue
        if table is not None:
            flush_table()
        if not s:
            flush_para(); flush_list()
            i += 1
            continue
        # 列表识别
        om = ORDER_RE.match(line)
        bm = BULLET_RE.match(line)
        if om:
            flush_para()
            if listbuf is None or listbuf[0] != 'ordered':
                listbuf = ('ordered', [])
            listbuf[1].append((None, om.group(1).strip()))
            i += 1
            continue
        if bm:
            flush_para()
            if listbuf is None or listbuf[0] != 'unordered':
                listbuf = ('unordered', [])
            listbuf[1].append((None, bm.group(1).strip()))
            i += 1
            continue
        # 普通正文段
        flush_list()
        para.append(line)
        i += 1

    flush_para(); flush_list(); flush_table()
    if code:
        b.code(code)


def resolve_in(arg):
    """输入路径: 原样返回(相对路径按当前工作目录解析)"""
    return arg


def resolve_out(arg, src):
    """输出路径约定:
    - 若显式指定 -> 直接用。
    - 若输入位于某个名为 input/ 的目录 -> 输出到同级 output/ 目录(同名 .docx)。
    - 其余情况 -> 与输入同目录同名 .docx。
    """
    if arg:
        return arg
    src_abs = os.path.abspath(src)
    src_dir = os.path.dirname(src_abs)
    # 输入在 input/ 下 => 输出到同级 output/ 下
    if os.path.basename(src_dir) == 'input':
        out_dir = os.path.join(os.path.dirname(src_dir), 'output')
        return os.path.join(out_dir, os.path.splitext(os.path.basename(src))[0] + '.docx')
    return os.path.splitext(src)[0] + '.docx'


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    src = resolve_in(sys.argv[1])
    if not os.path.exists(src):
        print('输入文件不存在:', src)
        return 1
    dst = resolve_out(sys.argv[2] if len(sys.argv) > 2 else None, src)
    os.makedirs(os.path.dirname(os.path.abspath(dst)), exist_ok=True)
    with open(src, encoding='utf-8') as f:
        text = f.read()
    base_dir = os.path.dirname(os.path.abspath(src))
    b = Builder()
    parse_md(text, b, base_dir=base_dir)
    b.build(dst)
    print('OK ->', os.path.abspath(dst))
    return 0


if __name__ == '__main__':
    sys.exit(main())
