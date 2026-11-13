"""Builds verify-mix.docx: tables (merged + nested + shaded), an image, comments, tracked changes, plain paragraphs."""
import sys, zipfile, zlib, struct

out = sys.argv[1]

def png(w, h):
    raw = b"".join(b"\0" + b"".join(bytes((x * 255 // w, y * 255 // h, 160)) for x in range(w)) for y in range(h))
    def chunk(t, d):
        c = struct.pack(">I", len(d)) + t + d
        return c + struct.pack(">I", zlib.crc32(t + d) & 0xffffffff)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'

def p(text, style=None):
    ps = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    return f'<w:p>{ps}<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'

def tc(inner, w=2400, extra=""):
    return f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{extra}</w:tcPr>{inner}</w:tc>'

borders = '<w:tblBorders>' + "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="808080"/>'
                                     for s in ("top", "left", "bottom", "right", "insideH", "insideV")) + '</w:tblBorders>'
def tbl(rows, cols):
    grid = "".join(f'<w:gridCol w:w="{c}"/>' for c in cols)
    return f'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>{borders}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{rows}</w:tbl>'

VMR = '<w:vMerge w:val="restart"/>'
VMC = '<w:vMerge/>'
GS2 = '<w:gridSpan w:val="2"/>'
shade = '<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/>'
nested = tbl("".join(f'<w:tr>{tc(p(f"Nested {r}.1"), 1100)}{tc(p(f"Nested {r}.2"), 1100)}</w:tr>' for r in (1, 2)), [1100, 1100])
table = tbl(
    f'<w:tr>{tc(p("Item"), 2400, shade)}{tc(p("Quantity"), 2400, shade)}{tc(p("Notes"), 2400, shade)}</w:tr>'
    f'<w:tr>{tc(p("Apples"), 2400, VMR)}{tc(p("12"))}{tc(p("Merged across two rows on the left"))}</w:tr>'
    f'<w:tr>{tc(p(""), 2400, VMC)}{tc(p("7"))}{tc(nested + p("After the nested table"))}</w:tr>'
    f'<w:tr>{tc(p("Total spans two columns"), 4800, GS2)}{tc(p("19"))}</w:tr>',
    [2400, 2400, 2400])

image = ('<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><wp:docPr id="1" name="Picture 1"/>'
         '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
         '<pic:nvPicPr><pic:cNvPr id="1" name="gradient.png"/><pic:cNvPicPr/></pic:nvPicPr>'
         '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
         '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
         '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>')

D = 'w:author="Reviewer" w:date="2026-11-13T09:00:00Z"'
commented = ('<w:p><w:r><w:t xml:space="preserve">This sentence has </w:t></w:r><w:commentRangeStart w:id="0"/>'
             '<w:r><w:t>a comment on these words</w:t></w:r><w:commentRangeEnd w:id="0"/>'
             '<w:r><w:commentReference w:id="0"/></w:r><w:r><w:t xml:space="preserve"> and ends here.</w:t></w:r></w:p>')
tracked = (f'<w:p><w:r><w:t xml:space="preserve">Tracked changes: the meeting is on </w:t></w:r>'
           f'<w:del w:id="10" {D}><w:r><w:delText>Monday</w:delText></w:r></w:del>'
           f'<w:ins w:id="11" {D}><w:r><w:t>Tuesday</w:t></w:r></w:ins>'
           f'<w:r><w:t xml:space="preserve"> at ten</w:t></w:r><w:ins w:id="12" {D}><w:r><w:t xml:space="preserve"> in room 4</w:t></w:r></w:ins>'
           f'<w:r><w:t>.</w:t></w:r></w:p>')
commented2 = ('<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>A whole paragraph with a second comment.</w:t></w:r>'
              '<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>')

body = (p("Verify mix", "Heading1") + p("A plain paragraph to type into. Add a word here.")
        + p("Table", "Heading2") + table + p("")
        + p("Image", "Heading2") + image + p("The picture above is a gradient.")
        + p("Comments", "Heading2") + commented + commented2
        + p("Tracked changes", "Heading2") + tracked
        + p("Undo", "Heading2") + p("Type here, then undo and redo.") + p("Last paragraph."))
document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
            '</w:body></w:document>')

def cm(i, text):
    return (f'<w:comment w:id="{i}" w:author="Reviewer" w:date="2026-11-13T09:00:00Z" w:initials="R">'
            f'<w:p><w:r><w:annotationRef/></w:r><w:r><w:t>{text}</w:t></w:r></w:p></w:comment>')
comments = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments {W}>{cm(0, "First comment.")}{cm(1, "Second comment.")}</w:comments>'

def style(sid, name, sz, extra=""):
    return (f'<w:style w:type="paragraph" w:styleId="{sid}"><w:name w:val="{name}"/><w:basedOn w:val="Normal"/>'
            f'<w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="80"/>{extra}</w:pPr><w:rPr><w:b/><w:sz w:val="{sz}"/></w:rPr></w:style>')
styles = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}>'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
          + style("Heading1", "heading 1", 32, '<w:outlineLvl w:val="0"/>') + style("Heading2", "heading 2", 26, '<w:outlineLvl w:val="1"/>')
          + '</w:styles>')

ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>')
rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
docrels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           f'<Relationship Id="rIdStyles" Type="{R}styles" Target="styles.xml"/>'
           f'<Relationship Id="rIdComments" Type="{R}comments" Target="comments.xml"/>'
           f'<Relationship Id="rIdImg" Type="{R}image" Target="media/image1.png"/></Relationships>')

with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    for name, data in [("[Content_Types].xml", ct), ("_rels/.rels", rels), ("word/document.xml", document),
                       ("word/_rels/document.xml.rels", docrels), ("word/styles.xml", styles),
                       ("word/comments.xml", comments), ("word/media/image1.png", png(192, 96))]:
        z.writestr(name, data)
print("wrote", out)
