// Mail merge checks (host\Merge.cs): templates made here, one per kind of field, scanned and merged, the result's text
// compared with what Word gives. Merge.cs needs Settings.Get and MainForm.Log from the program: small stand-ins below.
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

static class Settings { public static string Get(string section, string key, string fallback = "") => fallback; }
static class MainForm { public static void Log(string text) => Console.WriteLine("  log: " + text); }

static class Program {
    static int failed, passed;
    static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    static int Main(string[] args) {
        if (args.Length > 0) return ScanFiles(args);

        // ---- MERGEFIELD and its switches
        {
            var doc = Make(P(R("Dear "), Field(" MERGEFIELD Name \\* Upper "), R(".")),
                P(Field(" MERGEFIELD Amount \\# $#,##0.00 ")),
                P(Field(" MERGEFIELD Date \\@ \"MMMM d, yyyy\" ")),
                P(Field(" MERGEFIELD Title \\b \"Title: \" \\f \"!\" ")),
                P(Field(" MERGEFIELD Empty \\b \"Never: \" ")),
                P(Field(" MERGEFIELD \"First Name\" ")));
            var fields = Merge.Scan(doc);
            Check("mergefield: scan finds 6 names", fields.Count == 6, string.Join(", ", fields.Select(f => f.Name)));
            var text = Text(Merge.Fill(doc, V(("Name", "ann smith"), ("Amount", "1234.5"), ("Date", "2026-11-25"), ("Title", "Ms"), ("Empty", ""), ("First Name", "Ann"))));
            Check("mergefield: \\* Upper", text.Contains("Dear ANN SMITH."), text);
            Check("mergefield: \\# number picture", text.Contains("$1,234.50"), text);
            Check("mergefield: \\@ date picture", text.Contains("November 25, 2026"), text);
            Check("mergefield: \\b \\f around a value", text.Contains("Title: Ms!"), text);
            Check("mergefield: \\b left out for an empty value", !text.Contains("Never"), text);
            Check("mergefield: quoted name with a space", text.Contains("Ann"), text);
        }

        // ---- IF: yes/no over two paragraphs, a comparison with text, wildcards, nested in a table cell
        {
            byte[] Doc() => Make(
                P(R("Before "), Begin(), Instr(" IF "), Field(" MERGEFIELD IsGuar "), Instr(" = True \"First line")),
                P(Instr("Second line\" \"No guaranty\" "), Separate(), R("(old result)"), End(), R(" after")),
                P(Begin(), Instr(" IF "), Field(" MERGEFIELD Status "), Instr(" <> \"LLC\" \"Not an LLC\" \"An LLC\" "), Separate(), End()),
                P(Begin(), Instr(" IF "), Field(" MERGEFIELD Last "), Instr(" = \"Sm*\" \"Smith-like\" \"Other\" "), Separate(), End()),
                Table(P(Begin(), Instr(" IF "), Field(" MERGEFIELD IsGuar "), Instr(" = \"True\" \"Cell yes\" \"Cell no\" "), Separate(), End())));
            var fields = Merge.Scan(Doc());
            Check("if: IsGuar is a yes/no field", fields.Any(f => f.Name == "IsGuar" && f.InCondition), Dump(fields));
            Check("if: Status (compared with LLC) is not yes/no", fields.Any(f => f.Name == "Status" && !f.InCondition), Dump(fields));
            var yes = Text(Merge.Fill(Doc(), V(("IsGuar", "True"), ("Status", "LLC"), ("Last", "Smithers"))));
            Check("if true: text over two paragraphs kept", yes.Contains("Before First line\nSecond line after"), yes);
            Check("if: <> comparison", yes.Contains("An LLC"), yes);
            Check("if: wildcard", yes.Contains("Smith-like"), yes);
            Check("if: in a table cell", yes.Contains("Cell yes"), yes);
            var no = Text(Merge.Fill(Doc(), V(("IsGuar", "False"), ("Status", "Inc."), ("Last", "Jones"))));
            Check("if false: one paragraph, other text", no.Contains("Before No guaranty after") && !no.Contains("First line"), no);
            Check("if false: comparisons", no.Contains("Not an LLC") && no.Contains("Other") && no.Contains("Cell no"), no);
            Check("if: no field left", FieldCount(Merge.Fill(Doc(), V(("IsGuar", "True")))) == 0, "");
        }

        // ---- = formulas around merge fields
        {
            var doc = Make(P(Begin(), Instr(" = "), Field(" MERGEFIELD A "), Instr(" + "), Field(" MERGEFIELD B "), Instr(" * 2 \\# 0.00 "), Separate(), R("!Syntax Error"), End()),
                P(Begin(), Instr(" =("), Field(" MERGEFIELD A "), Instr("-1)/4 \\# 0.000 "), Separate(), End()),
                P(Field(" = SUM(ABOVE) ")));
            var text = Text(Merge.Fill(doc, V(("A", "1,000"), ("B", "2.5"))));
            Check("formula: + and * with a number picture", text.Contains("1005.00"), text);
            Check("formula: parentheses and /", text.Contains("249.750"), text);
            Check("formula: a function is left as it is", FieldCount(Merge.Fill(doc, V(("A", "1"), ("B", "1")))) == 1, text);
        }

        // ---- FILLIN, ASK and the REF that shows the answer
        {
            var doc = Make(P(R("Client: "), Field(" FILLIN \"Client name\" \\d \"Nobody\" ")),
                P(Field(" ASK Judge \"Which judge?\" \\d \"Brown\" ")),
                P(R("Judge "), Field(" REF Judge "), R(" presiding")));
            var fields = Merge.Scan(doc);
            Check("fillin/ask: scan", fields.Any(f => f.Name == "Client name" && f.Kind == MergeFieldKind.FillIn && f.Default == "Nobody")
                && fields.Any(f => f.Name == "Judge" && f.Kind == MergeFieldKind.Ask && f.Label == "Which judge?"), Dump(fields));
            var text = Text(Merge.Fill(doc, V(("Client name", "Acme"), ("Judge", "Grey"))));
            Check("fillin: value", text.Contains("Client: Acme"), text);
            Check("ask+ref: the answer where REF is", text.Contains("Judge Grey presiding"), text);
            var defaults = Text(Merge.Fill(doc, V()));
            Check("fillin/ask: defaults", defaults.Contains("Client: Nobody") && defaults.Contains("Judge Brown presiding"), defaults);
        }

        // ---- Typed placeholders; «Name» text and expressions stay
        {
            var doc = Make(P(R("Hello {{"), R("Name", bold: true), R("}}, from <<City>>.")),
                P(R("{{Missing|Somewhere}} and {{Unknown}}")),
                P(R("«Plain» text and {{ a.b }} stay")));
            var fields = Merge.Scan(doc);
            Check("placeholders: scan", fields.Select(f => f.Name).OrderBy(n => n).SequenceEqual(new[] { "City", "Missing", "Name", "Unknown" }), Dump(fields));
            var text = Text(Merge.Fill(doc, V(("Name", "Ann"), ("City", "Taos"))));
            Check("placeholders: across runs, << >>", text.Contains("Hello Ann, from Taos."), text);
            Check("placeholders: default; a name with no value left", text.Contains("Somewhere and {{Unknown}}"), text);
            Check("placeholders: «Plain» and {{ a.b }} stay", text.Contains("«Plain» text and {{ a.b }} stay"), text);
        }

        // ---- A Jinja template's {{ }} are left for dAHKxodusPro
        {
            var doc = Make(P(R("{% if x %}{{Name}}{% endif %}")));
            Check("jinja: no placeholders", Merge.Scan(doc).Count == 0, Dump(Merge.Scan(doc)));
        }

        // ---- Content controls: by tag, by title, bound to custom XML; block and inline
        {
            var doc = Make(P(R("Borrower: "), Sdt(Pr(tag: "Borrower"), R("Click or tap here to enter text."))),
                P(R("Day "), Sdt(Pr(alias: "Due day"), R("[day]")), R(" and again "), Sdt(Pr(xpath: "/ns0:Loan[1]/ns0:FillDay[1]"), R("x"))),
                Sdt(Pr(tag: "Notice"), P(R("Old notice"))),
                P(R("Untagged: "), Sdt(Pr(), R("stays"))));
            var fields = Merge.Scan(doc);
            Check("controls: scan", fields.Select(f => f.Name).OrderBy(n => n).SequenceEqual(new[] { "Borrower", "Due day", "FillDay", "Notice" }), Dump(fields));
            var merged = Merge.Fill(doc, V(("Borrower", "Ann"), ("Due day", "5th"), ("FillDay", "5th"), ("Notice", "Line one\nLine two")));
            var text = Text(merged);
            Check("controls: inline values", text.Contains("Borrower: Ann") && text.Contains("Day 5th and again 5th"), text);
            Check("controls: block value with a line break", text.Contains("Line one") && text.Contains("Line two") && !text.Contains("Old notice"), text);
            Check("controls: filled ones removed, others kept", Count(merged, "sdt") == 1 && text.Contains("Untagged: stays"), text);
        }

        // ---- Headers, fldSimple, w:mailMerge
        {
            var doc = Make(new[] { P(new XElement(W + "fldSimple", new XAttribute(W + "instr", " MERGEFIELD Name "), R("«Name»"))) },
                header: P(R("Re: "), Field(" MERGEFIELD Name ")), mailMerge: true);
            var merged = Merge.Fill(doc, V(("Name", "Ann")));
            Check("fldSimple: filled", Text(merged) == "Ann", Text(merged));
            Check("header: filled", Text(merged, header: true).Contains("Re: Ann"), Text(merged, header: true));
            using var ms = new MemoryStream(merged);
            using var w = WordprocessingDocument.Open(ms, false);
            Check("settings: w:mailMerge removed", w.MainDocumentPart!.DocumentSettingsPart!.Settings!.OuterXml.Contains("mailMerge") == false, "");
        }

        Console.WriteLine($"\n{passed} passed, {failed} failed");
        return failed == 0 ? 0 : 1;
    }

    // ---- Other templates: dotnet run -c Release -- a.dotx b.docx ...
    static int ScanFiles(string[] paths) {
        foreach (var path in paths) {
            try {
                var bytes = File.ReadAllBytes(path);
                var fields = Merge.Scan(bytes);
                Console.WriteLine($"{Path.GetFileName(path)}: {fields.Count} field(s)");
                foreach (var f in fields) Console.WriteLine($"  {f.Kind,-11} {f.Name}{(f.InCondition ? " (yes/no)" : "")}{(f.Default != "" ? " = " + f.Default : "")}");
                var merged = Merge.Fill(bytes, fields.ToDictionary(f => f.Name, f => f.InCondition ? "True" : "<" + f.Name + ">"));
                Console.WriteLine($"  merged: {Merge.Scan(merged).Count} field(s) left, {FieldCount(merged)} field code(s) left");
                foreach (var code in Part(merged, false).Descendants(W + "instrText").Select(t => t.Value.Trim()).Where(t => t != "").Take(5))
                    Console.WriteLine("    left: " + code);
            } catch (Exception ex) {
                Console.WriteLine($"{Path.GetFileName(path)}: FAILED {ex.Message}");
                failed++;
            }
        }
        return failed == 0 ? 0 : 1;
    }

    static void Check(string name, bool ok, string detail) {
        if (ok) { passed++; Console.WriteLine("ok     " + name); }
        else { failed++; Console.WriteLine("FAILED " + name + (detail != "" ? "\n       " + detail.Replace("\n", "\n       ") : "")); }
    }

    static Dictionary<string, string> V(params (string k, string v)[] values) => values.ToDictionary(x => x.k, x => x.v);
    static string Dump(List<MergeFieldInfo> fields) => string.Join("; ", fields.Select(f => $"{f.Kind} {f.Name} [{f.Label}] ={f.Default}{(f.InCondition ? " yes/no" : "")}"));

    // ---- Making documents

    static XElement R(string text, bool bold = false) => new(W + "r",
        bold ? new XElement(W + "rPr", new XElement(W + "b")) : null,
        new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text));
    static XElement P(params object[] content) => new(W + "p", content);
    static XElement Begin() => new(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "begin")));
    static XElement Separate() => new(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "separate")));
    static XElement End() => new(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "end")));
    static XElement Instr(string code) => new(W + "r", new XElement(W + "instrText", new XAttribute(XNamespace.Xml + "space", "preserve"), code));
    static object[] Field(string code, string result = "«x»") => [Begin(), Instr(code), Separate(), R(result), End()];
    static XElement Table(XElement cellParagraph) => new(W + "tbl", new XElement(W + "tblGrid", new XElement(W + "gridCol", new XAttribute(W + "w", "4000"))),
        new XElement(W + "tr", new XElement(W + "tc", cellParagraph)));
    static XElement Pr(string? tag = null, string? alias = null, string? xpath = null) => new(W + "sdtPr",
        alias != null ? new XElement(W + "alias", new XAttribute(W + "val", alias)) : null,
        tag != null ? new XElement(W + "tag", new XAttribute(W + "val", tag)) : null,
        xpath != null ? new XElement(W + "dataBinding", new XAttribute(W + "xpath", xpath), new XAttribute(W + "storeItemID", "{00000000-0000-0000-0000-000000000001}")) : null);
    static XElement Sdt(XElement pr, params XElement[] content) => new(W + "sdt", pr, new XElement(W + "sdtContent", content));

    static byte[] Make(params XElement[] body) => Make(body, null, false);

    static byte[] Make(XElement[] body, XElement? header, bool mailMerge) {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Template)) {
            var main = doc.AddMainDocumentPart();
            var sect = new XElement(W + "sectPr");
            if (header != null) {
                var h = main.AddNewPart<HeaderPart>();
                Write(h, new XElement(W + "hdr", new XAttribute(XNamespace.Xmlns + "w", W), header));
                sect.Add(new XElement(W + "headerReference", new XAttribute(W + "type", "default"),
                    new XAttribute(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"), main.GetIdOfPart(h))));
            }
            var settings = main.AddNewPart<DocumentSettingsPart>();
            Write(settings, new XElement(W + "settings", new XAttribute(XNamespace.Xmlns + "w", W),
                mailMerge ? new XElement(W + "mailMerge", new XElement(W + "mainDocumentType", new XAttribute(W + "val", "formLetters"))) : null));
            Write(main, new XElement(W + "document", new XAttribute(XNamespace.Xmlns + "w", W),
                new XAttribute(XNamespace.Xmlns + "r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
                new XElement(W + "body", body, sect)));
        }
        return ms.ToArray();
    }

    static void Write(OpenXmlPart part, XElement root) {
        using var s = part.GetStream(FileMode.Create, FileAccess.Write);
        new XDocument(root).Save(s, SaveOptions.DisableFormatting);
    }

    // ---- Reading results

    static XDocument Part(byte[] bytes, bool header) {
        using var ms = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        OpenXmlPart part = header ? doc.MainDocumentPart!.HeaderParts.First() : doc.MainDocumentPart!;
        using var s = part.GetStream();
        return XDocument.Load(s, LoadOptions.PreserveWhitespace);
    }

    // The paragraphs' text (w:t only, not field codes), one line each.
    static string Text(byte[] bytes, bool header = false) =>
        string.Join("\n", Part(bytes, header).Descendants(W + "p").Where(p => !p.Descendants(W + "p").Any())
            .Select(p => string.Concat(p.Descendants().Where(e => e.Name == W + "t" || e.Name == W + "br").Select(e => e.Name == W + "br" ? "\n" : e.Value))));

    static int Count(byte[] bytes, string local) => Part(bytes, false).Descendants(W + local).Count();
    // Complex fields (their begin markers; a field needn't have a separate) and simple ones.
    static int FieldCount(byte[] bytes) =>
        Part(bytes, false).Descendants(W + "fldChar").Count(f => (string?)f.Attribute(W + "fldCharType") == "begin") + Count(bytes, "fldSimple");
}
