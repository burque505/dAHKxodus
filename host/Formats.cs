// Save a Copy As (File tab, Doc.SaveCopy): the document in another format, made by Docxodus's own converters from the
// .docx bytes the page gives (the window goes on editing the .docx). Formats by extension:
//   .html  a web page in one file: Docxodus's HTML converter, pictures inside it (data URIs), footnotes at the end
//   .md    Markdown: Docxodus's Markdown converter (headings, lists, tables, bold and italic; no pictures)
//   .txt   plain text, UTF-8: the paragraphs' text, list numbers included, a table's cells separated by tabs
// Tracked changes are accepted in every copy (the copy shows the document as it would read with them accepted);
// comments are left out.
using System.Text;
using System.Xml.Linq;
using Docxodus;
using DocumentFormat.OpenXml.Packaging;

public static class Formats {
    public static readonly string[] Extensions = [".html", ".md", ".txt"];

    public static bool Supported(string path) =>
        Extensions.Contains(Path.GetExtension(path).ToLowerInvariant()) || Path.GetExtension(path).ToLowerInvariant() == ".htm";

    /// The copy's text for path's extension (.html/.htm, .md, .txt). Throws on an extension it can't write.
    public static string Convert(byte[] docx, string path, string title) {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        return ext switch {
            ".html" or ".htm" => Html(docx, title),
            ".md" => Markdown(docx),
            ".txt" => PlainText(docx),
            _ => throw new NotSupportedException("No copy as " + ext),
        };
    }

    /// Writes the copy (UTF-8; a .txt with Windows line ends) beside the old file and then in its place, as Save does.
    public static void Write(string path, string text) {
        var full = Path.GetFullPath(path);
        var tmp = Path.Combine(Path.GetDirectoryName(full)!, $"~dax{Guid.NewGuid():N}.tmp");
        File.WriteAllText(tmp, text, new UTF8Encoding(false));
        try {
            if (File.Exists(full)) File.Replace(tmp, full, null);
            else File.Move(tmp, full);
        } catch {
            try { File.Delete(tmp); } catch { }
            throw;
        }
    }

    // ---- HTML

    static readonly HashSet<string> VoidElements = ["area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "source", "track", "wbr"];

    static string Html(byte[] docx, string title) {
        using var stream = Writable(docx);
        using var doc = WordprocessingDocument.Open(stream, true);
        var settings = new WmlToHtmlConverterSettings {
            PageTitle = title,
            CssClassPrefix = "docx-",
            FabricateCssClasses = true,
            GeneralCss = "body { font-family: Arial, sans-serif; margin: 20px auto; max-width: 50em; }",
            RenderComments = false,
            RenderFootnotesAndEndnotes = true,
            RenderHeadersAndFooters = false,    // (a web page has no pages to put them on)
            RenderTrackedChanges = false,       // (accepted)
            ImageHandler = image => image.ImageBytes is { Length: > 0 } bytes
                ? image.CreateHtmlImage($"data:{image.ContentType ?? "image/png"};base64,{System.Convert.ToBase64String(bytes)}")
                : null!,
        };
        var html = WmlToHtmlConverter.ConvertToHtml(doc, settings);
        // Written as XML, an empty <span/> or <p/> would open an element a browser never closes: give each one that
        // isn't an HTML void element an end tag (an empty text), and drop the XHTML namespace the converter puts on.
        foreach (var e in html.DescendantsAndSelf().ToList()) {
            e.Name = e.Name.LocalName;
            foreach (var a in e.Attributes().Where(a => a.IsNamespaceDeclaration).ToList()) a.Remove();
            if (e.IsEmpty && !VoidElements.Contains(e.Name.LocalName)) e.Add("");
        }
        return "<!DOCTYPE html>\n" + html.ToString(SaveOptions.DisableFormatting);
    }

    // ---- Markdown

    static string Markdown(byte[] docx) {
        var settings = new WmlToMarkdownConverterSettings {
            Scopes = ProjectionScopes.Body | ProjectionScopes.Footnotes | ProjectionScopes.Endnotes,
            AnchorMode = AnchorRenderMode.None,
            TableMode = TableRenderMode.AlwaysGfm,
            TrackedChanges = TrackedChangeMode.Accept,
            ResolveNumbering = true,
            EmptyParagraphs = EmptyParagraphMode.Suppress,
        };
        var md = WmlToMarkdownConverter.Convert(new WmlDocument("document.docx", docx), settings).Markdown;
        // The converter heads each part it projects (# Document, # Footnotes, # Endnotes, with --- between them) for
        // programs that split it; a reader's copy keeps the notes' definitions without those headings.
        var lines = md.Replace("\r\n", "\n").Split('\n').ToList();
        if (lines.Count > 0 && lines[0] == "# Document") { lines.RemoveAt(0); if (lines.Count > 0 && lines[0] == "") lines.RemoveAt(0); }
        for (var i = lines.Count - 1; i >= 0; i--)
            if (lines[i] is "# Footnotes" or "# Endnotes") {
                lines.RemoveAt(i);
                if (i < lines.Count && lines[i] == "") lines.RemoveAt(i);
                if (i >= 2 && lines[i - 2] == "---" && lines[i - 1] == "") { lines.RemoveRange(i - 2, 2); }
            }
        return string.Join("\n", lines).TrimEnd() + "\n";
    }

    // ---- Plain text

    static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    static readonly XNamespace MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    static string PlainText(byte[] docx) {
        using var stream = Writable(docx);
        using var doc = WordprocessingDocument.Open(stream, true);
        RevisionAccepter.AcceptRevisions(doc);
        var body = doc.MainDocumentPart?.GetXDocument().Root?.Element(W + "body");
        var lines = new List<string>();
        if (body != null) Blocks(doc, body, lines);
        // Notes at the end, after a rule, as they would print
        var notes = new List<string>();
        foreach (var (part, name) in new (OpenXmlPart? part, string name)[] {
                     (doc.MainDocumentPart?.FootnotesPart, "footnote"), (doc.MainDocumentPart?.EndnotesPart, "endnote") }) {
            if (part?.GetXDocument().Root is not { } root) continue;
            foreach (var note in root.Elements(W + name).Where(n => (string?)n.Attribute(W + "type") is null or "normal"))
                Blocks(doc, note, notes);
        }
        if (notes.Any(n => n.Trim() != "")) { lines.Add(""); lines.Add("________"); lines.AddRange(notes); }
        return string.Join("\r\n", lines).TrimEnd() + "\r\n";
    }

    // A container's paragraphs and tables, in order (content controls and custom XML are looked into).
    static void Blocks(WordprocessingDocument doc, XElement container, List<string> lines) {
        foreach (var e in container.Elements()) {
            if (e.Name == W + "p") lines.Add(Paragraph(doc, e));
            else if (e.Name == W + "tbl") {
                foreach (var row in e.Elements(W + "tr"))
                    lines.Add(string.Join("\t", row.Elements(W + "tc").Select(cell => {
                        var inner = new List<string>();
                        Blocks(doc, cell, inner);
                        return string.Join(" ", inner.Where(t => t != "")).Replace("\t", " ").Replace("\r\n", " ");
                    })));
                lines.Add("");
            } else if (e.Name == W + "sdt") { if (e.Element(W + "sdtContent") is { } c) Blocks(doc, c, lines); }
            else if (e.Name == W + "customXml" || e.Name == MC + "AlternateContent") Blocks(doc, e, lines);
            else if (e.Name == MC + "Choice") Blocks(doc, e, lines);
        }
    }

    static string Paragraph(WordprocessingDocument doc, XElement p) {
        var text = new StringBuilder();
        string? marker = null;
        try { marker = ListItemRetriever.RetrieveListItem(doc, p); } catch { }
        if (!string.IsNullOrEmpty(marker)) text.Append(marker).Append('\t');
        Runs(p, text);
        return text.ToString().Replace("\n", "\r\n");
    }

    // The text of a paragraph's runs, in order; not what a drawing's text box or an alternative's fallback holds (the
    // fallback repeats the choice; text boxes are not part of the running text).
    static void Runs(XElement e, StringBuilder text) {
        foreach (var c in e.Elements()) {
            var n = c.Name;
            if (n == W + "t") text.Append(c.Value);
            else if (n == W + "tab" && c.Parent?.Name == W + "r") text.Append('\t');
            else if (n == W + "br" || n == W + "cr") text.Append('\n');
            else if (n == W + "noBreakHyphen") text.Append('-');
            else if (n == W + "pPr" || n == W + "rPr" || n == W + "drawing" || n == W + "pict" || n == W + "object"
                     || n == MC + "Fallback" || n == W + "delText" || n == W + "instrText") { }
            else Runs(c, text);
        }
    }

    static MemoryStream Writable(byte[] docx) {
        var stream = new MemoryStream();
        stream.Write(docx, 0, docx.Length);
        stream.Position = 0;
        return stream;
    }
}
