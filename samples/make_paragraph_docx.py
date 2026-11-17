"""Builds paragraph-dialog-tests.docx: the steps for an on-screen test of the Paragraph dialog, each beside the
paragraphs it is tried on (some set up to differ, one whose style keeps it with the next).
    python make_paragraph_docx.py paragraph-dialog-tests.docx"""
import sys, zipfile
from xml.sax.saxutils import escape

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
LOREM = ("This is a test paragraph with enough words to wrap across several lines, so that indents, "
         "line spacing and alignment show clearly on the page when they change.")


def p(text, ppr="", style=None, bold=False):
    st = f'<w:pStyle w:val="{style}"/>' if style else ""
    rpr = "<w:rPr><w:b/></w:rPr>" if bold else ""
    return f'<w:p><w:pPr>{st}{ppr}</w:pPr><w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'


def h(text):
    return p(text, style="Heading1")


def step(text):
    return p(text, ppr='<w:ind w:left="360"/>')


body = "".join([
    p("Paragraph dialog: on-screen test", style="Title"),
    p("Open the dialog with the small arrow beside PARAGRAPH (Home or Layout), or Ctrl+Shift+P > Paragraph settings. "
      "Each test lists its steps, then the paragraphs to try them on (marked TARGET)."),

    h("1. One paragraph"),
    step("Click in the TARGET. Open the dialog: Left 0\", Special (none), Line spacing Single, Widow/Orphan control ticked."),
    step("Set Left to 1 in, Special Hanging (By fills in 0.5\"), After 6, and on Line and Page Breaks tick Keep with next. OK."),
    step("Expect: the paragraph indents 1\" with a 0.5\" hang, 6 pt after. One Ctrl+Z puts everything back."),
    p("TARGET 1. " + LOREM),

    h("2. Several paragraphs that differ"),
    step("Select from the start of TARGET 2a to the end of TARGET 2c. Open the dialog."),
    step("Expect: Left blank (they differ), Line spacing blank, Keep with next mixed (a dash)."),
    step("Set only Right to 2 cm. OK. Expect: all three get a right indent; each keeps its own left indent and spacing."),
    step("One Ctrl+Z: all three come back at once."),
    p("TARGET 2a. Left 0.5\", keep with next. " + LOREM, ppr='<w:keepNext/><w:ind w:left="720"/>'),
    p("TARGET 2b. No indent, 1.5 line spacing. " + LOREM, ppr='<w:spacing w:line="360" w:lineRule="auto"/>'),
    p("TARGET 2c. Left 1\", first line 0.5\". " + LOREM, ppr='<w:ind w:left="1440" w:firstLine="720"/>'),

    h("3. Line spacing: Exactly and Multiple"),
    step("In TARGET 3, set Line spacing Exactly (At fills in 12 pt), change At to 24 pt. OK. Expect: wide, fixed lines."),
    step("Open again: Exactly, 24 pt shown. Set Multiple, At 1.15. OK. Expect: slightly open lines."),
    p("TARGET 3. " + LOREM),

    h("4. A bad entry"),
    step("In TARGET 4, type abc in Left and press OK. Expect: the dialog stays open, Left marked red, a message under it."),
    step("Type 1.5 cm instead and press Enter. Expect: the dialog closes and the paragraph indents 1.5 cm. Esc instead of OK changes nothing."),
    p("TARGET 4. " + LOREM),

    h("5. A style's own setting"),
    step("Click in the heading TARGET 5 (Heading 1 keeps with next by its style). Line and Page Breaks: Keep with next is ticked."),
    step("Untick it, OK, open again: unticked. Tick it again, OK: the heading follows its style again (nothing of its own)."),
    p("TARGET 5 (a Heading 1)", style="Heading1"),
    p("The paragraph after the heading."),

    h("6. Page breaks (page view)"),
    step("Switch to page view. In TARGET 6 tick Page break before. Expect: it starts a new page. Untick it: it moves back."),
    p("TARGET 6. " + LOREM),

    h("7. Spanish and centimetres"),
    step("Settings > Language Español, then restart. Open the dialog on TARGET 7: the labels are Spanish and Left shows 0 cm."),
    step("Type 1,5 in Left (a decimal comma). Aceptar. Expect: 1.5 cm. Set the language back afterwards."),
    p("TARGET 7. " + LOREM),

    h("8. Macros"),
    step("Record a macro (Ctrl+Shift+R), use the dialog once on TARGET 8, stop. The macro has one Doc.ParagraphFormat(...) line."),
    step("Undo, then run the macro: the same format comes back."),
    p("TARGET 8. " + LOREM),
])

sect = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'

styles = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}>'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
          '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/>'
          '<w:rPr><w:b/><w:sz w:val="36"/></w:rPr></w:style>'
          '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>'
          '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="60"/><w:outlineLvl w:val="0"/></w:pPr>'
          '<w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>'
          '</w:styles>')
ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
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
with zipfile.ZipFile(sys.argv[1], "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", ct)
    z.writestr("_rels/.rels", rels)
    z.writestr("word/document.xml", doc)
    z.writestr("word/_rels/document.xml.rels", doc_rels)
    z.writestr("word/styles.xml", styles)
