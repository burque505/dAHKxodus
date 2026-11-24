"""Write the Outline View test document (a plain .docx built by hand, no dependencies)."""
import sys
import zipfile
from xml.sax.saxutils import escape

out = sys.argv[1]

def p(text, style=None):
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ''
    return f'<w:p>{ppr}<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'

def table(rows):
    cells = ''.join(
        '<w:tr>' + ''.join(f'<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>{p(c)}</w:tc>' for c in row) + '</w:tr>'
        for row in rows)
    return ('<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>'
            '<w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/>'
            '<w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' + cells + '</w:tbl>')

long = ("This is a long body paragraph, long enough to wrap onto several lines, so that Show First Line Only "
        "has something to cut: it goes on about nothing in particular for a while, then a while longer, "
        "and then it stops.")

body = ''.join([
    p('Outline View test document', 'Title'),
    p('Body text before the first heading: it belongs to no heading, so it is never hidden by a collapse.'),
    p('1 Introduction', 'Heading1'),
    p('Introduction, first body paragraph.'),
    p(long),
    p('1.1 Background', 'Heading2'),
    p('Background body text.'),
    p('1.1.1 Earlier work', 'Heading3'),
    p('Earlier work body text.'),
    p('1.2 Scope', 'Heading2'),
    p('Scope body text.'),
    p('2 Method', 'Heading1'),
    p('Method body text, followed by a table that belongs to this heading.'),
    table([['Step', 'What'], ['One', 'Collect'], ['Two', 'Count']]),
    p('After the table.'),
    p('2.1 Materials', 'Heading2'),
    p('Materials body text.'),
    p('2.2 Procedure', 'Heading2'),
    p('Procedure body text.'),
    p('3 Results', 'Heading1'),
    p('Results body text.'),
    p('3.1 First result', 'Heading2'),
    p('First result body text.'),
    p('3.2 Second result', 'Heading2'),
    p('Second result body text.'),
    p('4 Conclusion', 'Heading1'),
    p('Conclusion body text, the last paragraph of the document.'),
])

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}'
            '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')

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
          + heading(1, 32, '2F5496') + heading(2, 26, '2F5496') + heading(3, 24, '1F3763') +
          '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style>'
          '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/></w:style>'
          '</w:styles>')

content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                 '<Default Extension="xml" ContentType="application/xml"/>'
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
            '</Relationships>')

with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', content_types)
    z.writestr('_rels/.rels', rels)
    z.writestr('word/document.xml', document)
    z.writestr('word/_rels/document.xml.rels', doc_rels)
    z.writestr('word/styles.xml', styles)
print('wrote', out)
