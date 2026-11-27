// Mail merge checks (host\Merge.cs): templates made here, one per kind of field, scanned and merged, the result's text
// compared with what Word gives. Merge.cs needs Settings.Get and MainForm.Log from the program: small stand-ins below.
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

static class Settings { public static string Get(string section, string key, string fallback = "") => fallback; }
static class MainForm { public static void Log(string text) => Console.WriteLine("  log: " + text); }
static class Lang { public static string T(string key, params object[] args) => key; }

static class Program {
    static int failed, passed;
    static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    static int Main(string[] args) {
        if (args.Length > 0 && args[0] == "--data") {
            // dotnet run -c Release -- --data file [table]: a data source's tables, columns and first rows
            Console.WriteLine("tables: " + string.Join(" | ", MergeData.Tables(args[1])));
            var t = MergeData.Load(args[1], args.Length > 2 ? args[2] : "");
            Console.WriteLine($"{t.Name}: {t.Columns.Count} columns, {t.Rows.Count} rows\n  " + string.Join(" | ", t.Columns));
            foreach (var r in t.Rows.Take(3)) Console.WriteLine("  " + string.Join(" | ", r.Select(v => v.Replace("\n", "\\n"))));
            return 0;
        }
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

        // ---- Blank lines: a paragraph of merge fields only, all empty, is dropped (Word's default); one with other text stays
        {
            var doc = Make(P(Field(" MERGEFIELD Name ")), P(Field(" MERGEFIELD Company ")), P(Field(" MERGEFIELD Street ")),
                P(Field(" MERGEFIELD Unit "), R(" "), Field(" MERGEFIELD Floor ")), P(R("Attn: "), Field(" MERGEFIELD Contact ")),
                P(Field(" MERGEFIELD City "), R(", "), Field(" MERGEFIELD State ")));
            var text = Text(Merge.Fill(doc, V(("Name", "Ann"), ("Company", ""), ("Street", "1 Main"), ("Unit", ""), ("Floor", ""),
                ("Contact", ""), ("City", "Taos"), ("State", "NM"))));
            Check("blank lines: empty field-only lines dropped", text == "Ann\n1 Main\nAttn: \nTaos, NM", text.Replace("\n", "|"));
        }

        // ---- Records: NEXT takes the next row into the same copy (labels); NEXTIF when its condition holds; SKIPIF
        // leaves the record out; MERGEREC the row's number, MERGESEQ the copy's
        {
            IReadOnlyDictionary<string, string> Row(string name, string status = "") => V(("Name", name), ("Status", status));
            var labels = Make(P(Field(" MERGEFIELD Name "), R(" #"), Field(" MERGEREC ")), P(Field(" NEXT ")),
                P(Field(" MERGEFIELD Name "), R(" #"), Field(" MERGEREC ")), P(R("copy "), Field(" MERGESEQ ")));
            var rows = new[] { Row("Ann"), Row("Bob"), Row("Cy") };
            var copies = Merge.MergeAll(labels, rows);
            Check("next: two rows a copy, three rows two copies", copies.Count == 2, copies.Count.ToString());
            var t1 = Text(copies[0].Bytes);
            var t2 = Text(copies[1].Bytes);
            Check("next: the first copy has rows 1 and 2", t1.Contains("Ann #1") && t1.Contains("Bob #2") && t1.Contains("copy 1"), t1.Replace("\n", "|"));
            Check("next: the second copy row 3, then nothing", t2.Contains("Cy #3") && t2.Contains("copy 2") && !t2.Contains("Ann"), t2.Replace("\n", "|"));
            Check("next: the copy's first row", copies[1].First["Name"] == "Cy", copies[1].First["Name"]);

            var skip = Make(P(Begin(), Instr(" SKIPIF "), Field(" MERGEFIELD Status "), Instr(" = \"void\" "), Separate(), End(),
                R("To "), Field(" MERGEFIELD Name "), R(" ("), Field(" MERGESEQ "), R(")")));
            var kept = Merge.MergeAll(skip, new[] { Row("Ann"), Row("Bob", "void"), Row("Cy") });
            var skipQuoted = Make(P(Begin(), Instr(" SKIPIF \""), Field(" MERGEFIELD Status "), Instr("\" = \"to be void\" "), Separate(), End(), Field(" MERGEFIELD Name ")));
            var keptQuoted = Merge.MergeAll(skipQuoted, new[] { Row("Ann", "to be void"), Row("Bob", "kept") });
            Check("skipif: a quoted field, a value with spaces", keptQuoted.Count == 1 && Text(keptQuoted[0].Bytes) == "Bob", keptQuoted.Count.ToString());
            Check("skipif: the record left out", kept.Count == 2, kept.Count.ToString());
            Check("skipif: the others numbered in order", Text(kept[1].Bytes).Contains("To Cy (2)"), Text(kept[1].Bytes));

            var nextif = Make(P(Field(" MERGEFIELD Name ")),
                P(Begin(), Instr(" NEXTIF "), Field(" MERGEFIELD Status "), Instr(" = \"pair\" "), Separate(), End(), Field(" MERGEFIELD Name ")));
            var pairs = Merge.MergeAll(nextif, new[] { Row("Ann", "pair"), Row("Bob"), Row("Cy") });
            Check("nextif: true takes the next row", pairs.Count == 2 && Text(pairs[0].Bytes) == "Ann\nBob", pairs.Count + " " + Text(pairs[0].Bytes).Replace("\n", "|"));
            Check("nextif: false stays on the row", Text(pairs[1].Bytes) == "Cy\nCy", Text(pairs[1].Bytes).Replace("\n", "|"));
            Check("fill: one row as before", Text(Merge.Fill(labels, Row("Zed"))).StartsWith("Zed #1"), Text(Merge.Fill(labels, Row("Zed"))));
        }

        // ---- Contains (Rules' Contains: Word's wildcard comparison = "*text*")
        {
            byte[] Doc() => Make(P(Begin(), Instr(" IF \""), Field(" MERGEFIELD LeShort "), Instr("\" = \"*Flanagan*\" \"Is a Flanagan\" \"Is not a Flanagan\" "), Separate(), End()));
            var yes = Text(Merge.Fill(Doc(), V(("LeShort", "Mary flanagan Smith"))));
            var no = Text(Merge.Fill(Doc(), V(("LeShort", "Mary Smith"))));
            Check("contains: within the text, any case", yes == "Is a Flanagan", yes);
            Check("contains: not there", no == "Is not a Flanagan", no);
        }

        // ---- A field list (Use a Field List...): names one a line, or every value of a CSV, once each
        {
            var dir = Path.Combine(Path.GetTempPath(), "dax-fieldlist-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            File.WriteAllText(Path.Combine(dir, "names.txt"), "First\r\n\r\n  Last Name \nfirst\nAmount\n");
            File.WriteAllText(Path.Combine(dir, "names.csv"), "First,Last,Amount\nCity,\"State, Code\"\n");
            var txt = MergeData.FieldList(Path.Combine(dir, "names.txt"));
            var csv = MergeData.FieldList(Path.Combine(dir, "names.csv"));
            Directory.Delete(dir, true);
            Check("field list: .txt one a line, trimmed, once each", string.Join("|", txt) == "First|Last Name|Amount", string.Join("|", txt));
            Check("field list: .csv every value", string.Join("|", csv) == "First|Last|Amount|City|State, Code", string.Join("|", csv));
        }

        // ---- Number switches: CardText, OrdText, Ordinal, DollarText, Roman; a percent from a formula
        {
            var doc = Make(P(Field(" MERGEFIELD N \\* CardText ")), P(Field(" MERGEFIELD N \\* OrdText ")),
                P(Field(" MERGEFIELD N \\* Ordinal ")), P(Field(" MERGEFIELD Pay \\* DollarText ")),
                P(Field(" MERGEFIELD N \\* Roman ")), P(Field(" MERGEFIELD N \\* roman ")), P(Field(" MERGEFIELD Twelve \\* Ordinal ")),
                P(Begin(), Instr(" = "), Field(" MERGEFIELD Rate "), Instr(" * 100 \\# \"0.00%\" "), Separate(), End()),
                P(Begin(), Instr(" = "), Field(" MERGEFIELD Rate "), Instr(" * 100 \\* CardText "), Separate(), End()));
            var text = Text(Merge.Fill(doc, V(("N", "42"), ("Pay", "1234.56"), ("Twelve", "12"), ("Rate", "0.255"))));
            Check("switch: \\* CardText", text.Contains("forty-two\n"), text);
            Check("switch: \\* OrdText", text.Contains("forty-second"), text);
            Check("switch: \\* Ordinal", text.Contains("42nd") && text.Contains("12th"), text);
            Check("switch: \\* DollarText", text.Contains("one thousand two hundred thirty-four and 56/100"), text);
            Check("switch: \\* Roman / roman", text.Contains("XLII") && text.Contains("xlii"), text);
            Check("formula: percent", text.Contains("25.50%"), text);
            Check("formula: percent as text", text.Contains("twenty-six") || text.Contains("twenty-five"), text);
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

        DataChecks();

        Console.WriteLine($"\n{passed} passed, {failed} failed");
        return failed == 0 ? 0 : 1;
    }

    // ---- Data sources (host\MergeData.cs)
    static void DataChecks() {
        var dir = Path.Combine(Path.GetTempPath(), "dAHKxodus-merge-checks");
        Directory.CreateDirectory(dir);

        // CSV: quotes with separators, "" and a line break; a blank line; Excel's Windows-1252
        var csv = Path.Combine(dir, "people.csv");
        File.WriteAllBytes(csv, System.Text.Encoding.Latin1.GetBytes("Name,Street,Note\r\n\"Peña, Ann\",\"1 Main St\nSuite 2\",\"said \"\"hi\"\"\"\r\n\r\nBob,2 Elm,\r\n"));
        var t = MergeData.Load(csv);
        Check("csv: columns and rows", t.Columns.SequenceEqual(new[] { "Name", "Street", "Note" }) && t.Rows.Count == 2, string.Join("|", t.Columns) + " / " + t.Rows.Count);
        Check("csv: quoted separator, line break, quote, 1252", t.Rows.Count > 0 && t.Rows[0][0] == "Peña, Ann" && t.Rows[0][1] == "1 Main St\nSuite 2" && t.Rows[0][2] == "said \"hi\"",
            t.Rows.Count > 0 ? string.Join("|", t.Rows[0]) : "");
        var semi = MergeData.Text("A;B\n1,5;x\n", "s");
        Check("csv: semicolons", semi.Columns.Count == 2 && semi.Rows[0][0] == "1,5", string.Join("|", semi.Columns));
        var pasted = MergeData.Text("First Name\tLast\tFirst Name\nAnn\tSmith\tX", "pasted");
        Check("pasted: tabs, a repeated column renamed", pasted.Columns.SequenceEqual(new[] { "First Name", "Last", "First Name 2" }) && pasted.Row(0)["Last"] == "Smith", string.Join("|", pasted.Columns));

        // JSON: wrapped array, nested object, booleans, numbers, null
        var json = MergeData.Json("{\"count\":2,\"loans\":[{\"Bo\":{\"Full\":\"Ann\"},\"IsGuar\":true,\"Rate\":6.5},{\"Bo\":{\"Full\":\"Bob\"},\"IsGuar\":false,\"Extra\":null}]}", "j");
        Check("json: columns from all objects, nested a.b", json.Columns.SequenceEqual(new[] { "Bo.Full", "IsGuar", "Rate", "Extra" }), string.Join("|", json.Columns));
        Check("json: values", json.Row(0)["Bo.Full"] == "Ann" && json.Row(0)["IsGuar"] == "True" && json.Row(0)["Rate"] == "6.5" && json.Row(1)["Rate"] == "", string.Join("|", json.Rows[0]));

        // Excel: shared strings, inline string, booleans, a date, a gap in the columns, two sheets
        var xlsx = Path.Combine(dir, "book.xlsx");
        MakeXlsx(xlsx);
        Check("xlsx: sheets", MergeData.Tables(xlsx).SequenceEqual(new[] { "Loans", "Other" }), string.Join("|", MergeData.Tables(xlsx)));
        var x = MergeData.Load(xlsx, "Loans");
        Check("xlsx: columns (a gap: Column 3)", x.Columns.SequenceEqual(new[] { "Name", "Amount", "Column 3", "Due", "Guar" }), string.Join("|", x.Columns));
        var due = new DateTime(2026, 11, 25).ToString("d", System.Globalization.CultureInfo.CurrentCulture);
        Check("xlsx: values", x.Rows.Count == 1 && x.Row(0)["Name"] == "Ann" && x.Row(0)["Amount"] == "1234.5" && x.Row(0)["Due"] == due && x.Row(0)["Guar"] == "True",
            x.Rows.Count > 0 ? string.Join("|", x.Rows[0]) : "no rows");
        Check("xlsx: the other sheet", MergeData.Load(xlsx, "Other").Row(0)["Z"] == "inline", "");

        // A merge from a row: names matched ignoring case and spaces
        var doc = Make(P(Field(" MERGEFIELD BoFull "), R(" owes "), Field(" MERGEFIELD Amount \\# $#,##0.00 ")));
        var row = x.Row(0);
        var values = Merge.Scan(doc).ToDictionary(f => f.Name, f => row.FirstOrDefault(kv => MergeData.Key(kv.Key) == MergeData.Key(f.Name == "BoFull" ? "Name" : f.Name)).Value ?? "");
        Check("row merge: values from a row", Text(Merge.Fill(doc, values)) == "Ann owes $1,234.50", Text(Merge.Fill(doc, values)));
        // All rows in one document (as the host does it: DocumentBuilder, each row's sections kept)
        {
            var letter = Make(P(R("Dear "), Field(" MERGEFIELD Name ")), P(R("Body")));
            var rows = new[] { "Ann", "Bob", "Cy" }.Select((n, i) => new Docxodus.Source(new Docxodus.WmlDocument($"row{i}.docx", Merge.Fill(letter, V(("Name", n)))), true)).ToList();
            var all = Docxodus.DocumentBuilder.BuildDocument(rows).DocumentByteArray;
            var text = Text(all);
            Check("one document: every row, in order", text.Contains("Dear Ann") && text.IndexOf("Dear Bob") > text.IndexOf("Dear Ann") && text.IndexOf("Dear Cy") > text.IndexOf("Dear Bob"), text);
            Check("one document: a section per row (each starts a page)", Count(all, "sectPr") == 3 && FieldCount(all) == 0, Count(all, "sectPr").ToString());
        }
        Check("key: case, spaces, _ - . ignored", MergeData.Key("Bo Full") == MergeData.Key("bo_full") && MergeData.Key("Bo.Full") == "bofull", "");

        // Access: a database made here through ADOX (only where the Access Database Engine is)
        if (!MergeData.AccessAvailable) { Console.WriteLine("skip   access: no Access Database Engine on this computer"); return; }
        var accdb = Path.Combine(dir, "loans.accdb");
        try {
            File.Delete(accdb);
            MakeAccdb(accdb);
            var tables = MergeData.Tables(accdb);
            Check("access: tables and queries", tables.SequenceEqual(new[] { "Loans", "qGuaranteed" }), string.Join("|", tables));
            var a = MergeData.Load(accdb, "Loans");
            Check("access: columns", a.Columns.SequenceEqual(new[] { "ID", "BoFull", "Principal", "IsGuar", "Signed" }), string.Join("|", a.Columns));
            var signed = new DateTime(2026, 11, 25).ToString("d", System.Globalization.CultureInfo.CurrentCulture);
            Check("access: values (Yes/No as True/False, a date, a currency)", a.Rows.Count == 2 && a.Row(0)["BoFull"] == "Ann" && a.Row(0)["Principal"] == "125000" && a.Row(0)["IsGuar"] == "True" && a.Row(0)["Signed"] == signed && a.Row(1)["IsGuar"] == "False",
                a.Rows.Count > 0 ? string.Join("|", a.Rows[0]) : "no rows");
            var q = MergeData.Load(accdb, "qGuaranteed");
            Check("access: a saved query", q.Rows.Count == 1 && q.Row(0)["BoFull"] == "Ann", q.Rows.Count.ToString());
        } catch (Exception ex) {
            Check("access", false, ex.ToString());
        }
    }

    static void MakeXlsx(string path) {
        using var doc = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var book = doc.AddWorkbookPart();
        void Xml(OpenXmlPart part, string xml) { using var s = part.GetStream(FileMode.Create); using var w = new StreamWriter(s); w.Write(xml); }
        const string ns = "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";
        var s1 = book.AddNewPart<WorksheetPart>();
        var s2 = book.AddNewPart<WorksheetPart>();
        Xml(book, $"<workbook {ns}><sheets><sheet name=\"Loans\" sheetId=\"1\" r:id=\"{book.GetIdOfPart(s1)}\"/><sheet name=\"Other\" sheetId=\"2\" r:id=\"{book.GetIdOfPart(s2)}\"/></sheets></workbook>");
        Xml(book.AddNewPart<SharedStringTablePart>(), $"<sst {ns}><si><t>Name</t></si><si><t>Amount</t></si><si><t>Due</t></si><si><t>Guar</t></si><si><t>Ann</t></si></sst>");
        Xml(book.AddNewPart<WorkbookStylesPart>(), $"<styleSheet {ns}><cellXfs count=\"2\"><xf numFmtId=\"0\"/><xf numFmtId=\"14\" applyNumberFormat=\"1\"/></cellXfs></styleSheet>");
        Xml(s1, $"<worksheet {ns}><sheetData>" +
            "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\" t=\"s\"><v>1</v></c><c r=\"D1\" t=\"s\"><v>2</v></c><c r=\"E1\" t=\"s\"><v>3</v></c></row>" +
            "<row r=\"2\"><c r=\"A2\" t=\"s\"><v>4</v></c><c r=\"B2\"><v>1234.5</v></c><c r=\"D2\" s=\"1\"><v>46351</v></c><c r=\"E2\" t=\"b\"><v>1</v></c></row>" +
            "<row r=\"4\"></row></sheetData></worksheet>");
        Xml(s2, $"<worksheet {ns}><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Z</t></is></c></row><row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>inline</t></is></c></row></sheetData></worksheet>");
    }

    static void MakeAccdb(string path) {
        var catalog = Activator.CreateInstance(Type.GetTypeFromProgID("ADOX.Catalog")!)!;
        var provider = Type.GetTypeFromProgID("Microsoft.ACE.OLEDB.16.0") != null ? "Microsoft.ACE.OLEDB.16.0" : "Microsoft.ACE.OLEDB.12.0";
        var conn = catalog.GetType().InvokeMember("Create", System.Reflection.BindingFlags.InvokeMethod, null, catalog, [$"Provider={provider};Data Source={path};"])!;
        void Exec(string sql) => conn.GetType().InvokeMember("Execute", System.Reflection.BindingFlags.InvokeMethod, null, conn, [sql]);
        Exec("CREATE TABLE Loans (ID COUNTER PRIMARY KEY, BoFull TEXT(100), Principal CURRENCY, IsGuar YESNO, Signed DATETIME)");
        Exec("INSERT INTO Loans (BoFull, Principal, IsGuar, Signed) VALUES ('Ann', 125000, True, #2026-11-25#)");
        Exec("INSERT INTO Loans (BoFull, Principal, IsGuar, Signed) VALUES ('Bob', 5000.5, False, NULL)");
        Exec("CREATE VIEW qGuaranteed AS SELECT * FROM Loans WHERE IsGuar = True");
        conn.GetType().InvokeMember("Close", System.Reflection.BindingFlags.InvokeMethod, null, conn, null);
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(conn);
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(catalog);
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
