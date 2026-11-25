"""Write the captions test document (a plain .docx built by hand, no dependencies).

What it has, to try Insert Caption, Insert Table of Figures, Update Table and Cross-reference:
- three pictures (small PNGs made here) and three tables over several pages;
- one picture and one table without a caption (to caption with Insert Caption);
- two Figure captions numbered WRONG on purpose (7 and 3: Update Table / F9 renumbers them 1, 2);
- one Table caption above its table;
- a caption with a label the document alone uses ("Map"), which Insert Caption and Cross-reference
  should offer;
- a place near the top for a table of figures, and "See ..." paragraphs for cross-references.

    python make_captions_docx.py captions-test.docx
"""
import struct
import sys
import zipfile
import zlib
from xml.sax.saxutils import escape

out = sys.argv[1] if len(sys.argv) > 1 else 'captions-test.docx'


def png(width, height, rgb):
    """A solid-colour PNG with a darker border."""
    dark = tuple(c // 2 for c in rgb)
    rows = b''
    for y in range(height):
        row = b'\x00'
        for x in range(width):
            edge = x < 3 or y < 3 or x >= width - 3 or y >= height - 3
            row += bytes(dark if edge else rgb)
        rows += row

    def chunk(kind, data):
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data) & 0xffffffff)

    return (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', width, height, 8, 2, 0, 0, 0))
            + chunk(b'IDAT', zlib.compress(rows, 9)) + chunk(b'IEND', b''))


def run(text):
    return f'<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def p(text, style=None, extra_ppr=''):
    style_xml = f'<w:pStyle w:val="{style}"/>' if style else ''
    ppr = f'<w:pPr>{style_xml}{extra_ppr}</w:pPr>' if style or extra_ppr else ''
    return f'<w:p>{ppr}{run(text) if text else ""}</w:p>'


def page_break():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def seq(label, shown):
    return ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r><w:instrText xml:space="preserve"> SEQ {label} \\* ARABIC </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            f'<w:r><w:rPr><w:noProof/></w:rPr><w:t>{shown}</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>')


def caption(label, shown, text, keep_next=False):
    ppr = '<w:pPr><w:pStyle w:val="Caption"/>' + ('<w:keepNext/>' if keep_next else '') + '</w:pPr>'
    return f'<w:p>{ppr}{run(label.replace("_", " ") + " ")}{seq(label, shown)}{run(text)}</w:p>'


pictures = []  # (rId, file name, bytes)


def picture(name, rgb, width_in=2.5, height_in=1.5):
    n = len(pictures) + 1
    rid = f'rIdImg{n}'
    pictures.append((rid, f'image{n}.png', png(120, 72, rgb)))
    cx, cy = int(width_in * 914400), int(height_in * 914400)
    return ('<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing>'
            f'<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
            f'<wp:docPr id="{n}" name="{escape(name)}" descr="{escape(name)}"/>'
            '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:nvPicPr><pic:cNvPr id="{n}" name="image{n}.png"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>')


def table(rows):
    cols = len(rows[0])
    cells = ''.join(
        '<w:tr>' + ''.join(f'<w:tc><w:tcPr><w:tcW w:w="2800" w:type="dxa"/></w:tcPr>{p(c)}</w:tc>' for c in row) + '</w:tr>'
        for row in rows)
    return ('<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:jc w:val="center"/>'
            '<w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/>'
            '<w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>'
            '<w:tblGrid>' + '<w:gridCol w:w="2800"/>' * cols + '</w:tblGrid>' + cells + '</w:tbl>')


filler = ("Body text to push the next item down the page, so the captions land on different pages and the "
          "table of figures has page numbers worth checking. It says nothing in particular, at some length, "
          "and then it stops.")

body = ''.join([
    p('Captions test document', 'Title'),
    p('Put a table of figures here: References > Insert Table of Figures (try Figure, then Table, then Map).'),
    p(''),
    p('1 Pictures', 'Heading1'),
    p('A picture that already has a caption, numbered 7 on purpose: Update Table or F9 makes it 1.'),
    picture('Blue box', (70, 120, 200)),
    caption('Figure', '7', ': A blue box'),
    p(filler),
    p('A picture with NO caption: put the caret on it (or the paragraph it is in) and use Insert Caption. '
      'The new caption should take the number 2 and renumber the green one below to 3.'),
    picture('Orange box', (230, 140, 50)),
    p(filler),
    page_break(),
    p('A second captioned picture, numbered 3 on purpose.'),
    picture('Green box', (80, 170, 90)),
    caption('Figure', '3', ': A green box'),
    p('1.1 A map', 'Heading2'),
    p('A caption with a label only this document uses ("Map"): Insert Caption and Cross-reference should offer it.'),
    caption('Map', '1', ': The way to the station'),
    p(filler),
    page_break(),
    p('2 Tables', 'Heading1'),
    p('A table with its caption above it, as Word puts one.'),
    caption('Table', '1', ': Prices', keep_next=True),
    table([['Item', 'Price'], ['Apples', '1.20'], ['Pears', '1.50']]),
    p(filler),
    p('A table with NO caption: put the caret in a cell and use Insert Caption; it should go ABOVE the table as Table 2.'),
    table([['Name', 'Role'], ['Ana', 'Editor'], ['Ben', 'Reviewer']]),
    p(filler),
    page_break(),
    p('3 Cross-references', 'Heading1'),
    p('Put the caret at the end of each line and use References > Cross-reference, reference type Figure, Table or Map:'),
    p('As shown in '),
    p('The prices are in '),
    p('The caption text alone: '),
    p('On page '),
    p('Then add a caption above Figure 1 (or delete one) and press F9 with everything selected (Ctrl+A): '
      'the numbers, the references and the table of figures should all follow.'),
])

NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')
document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}'
            '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def heading(level, size, color):
    return (f'<w:style w:type="paragraph" w:styleId="Heading{level}"><w:name w:val="heading {level}"/><w:basedOn w:val="Normal"/>'
            f'<w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="60"/>'
            f'<w:outlineLvl w:val="{level - 1}"/></w:pPr><w:rPr><w:b/><w:color w:val="{color}"/><w:sz w:val="{size}"/></w:rPr></w:style>')


styles = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}>'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
          '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>'
          '<w:qFormat/><w:rPr><w:sz w:val="48"/></w:rPr></w:style>'
          + heading(1, 32, '2F5496') + heading(2, 26, '2F5496') +
          '<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>'
          '<w:uiPriority w:val="35"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:spacing w:after="200" w:line="240" w:lineRule="auto"/></w:pPr>'
          '<w:rPr><w:i/><w:iCs/><w:color w:val="44546A"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style>'
          '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style>'
          '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/></w:style>'
          '</w:styles>')

content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                 '<Default Extension="xml" ContentType="application/xml"/>'
                 '<Default Extension="png" ContentType="image/png"/>'
                 '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                 '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
                 '</Types>')
rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')
doc_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            + ''.join(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                      f'Target="media/{name}"/>' for rid, name, _ in pictures)
            + '</Relationships>')

with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', content_types)
    z.writestr('_rels/.rels', rels)
    z.writestr('word/document.xml', document)
    z.writestr('word/_rels/document.xml.rels', doc_rels)
    z.writestr('word/styles.xml', styles)
    for _, name, data in pictures:
        z.writestr(f'word/media/{name}', data)
print('wrote', out)
