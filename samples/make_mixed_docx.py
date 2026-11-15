"""Builds mixed-sizes.docx: a Letter portrait section, a Letter landscape section, an A5 portrait section."""
import sys, zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

def p(text):
    return f'<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'

def sect(w, h, orient=None):
    o = f' w:orient="{orient}"' if orient else ""
    return (f'<w:sectPr><w:pgSz w:w="{w}" w:h="{h}"{o}/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')

def last_p(text, sectpr):
    return f'<w:p><w:pPr>{sectpr}</w:pPr><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'

body = (
    p("Section one: Letter portrait.") + p("More text on the portrait page.")
    + last_p("End of section one.", sect(12240, 15840))
    + p("Section two: Letter landscape.") + p("A wide page for a wide table.")
    + last_p("End of section two.", sect(15840, 12240, "landscape"))
    + p("Section three: A5 portrait.") + p("A smaller page.")
    + sect(8391, 11906)
)
doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}</w:body></w:document>'
ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '</Types>')
rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')
with zipfile.ZipFile(sys.argv[1], "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", ct)
    z.writestr("_rels/.rels", rels)
    z.writestr("word/document.xml", doc)
