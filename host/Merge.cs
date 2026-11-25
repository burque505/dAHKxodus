// Mail merge (File > Mail Merge..., New from Template on a template with fields, Doc.Merge): a document's fields
// filled from values, in the package's XML, before the editor sees it. Loaded by dAHKxodus.cs (#load).
//   - Word's fields: MERGEFIELD (\* Upper/Lower/Caps/FirstCap, \b, \f, \# numbers, \@ dates), IF (comparisons, wildcards,
//     merge fields nested in its condition and texts, texts over several paragraphs: the chosen one kept with its
//     formatting), FILLIN (its prompt is its name), ASK (sets a bookmark's value) and the REF fields showing it.
//   - Typed placeholders: {{Name}} ({{Name|default}}), <<Name>>; the markers from [Merge] Markers (never «Name»: plain
//     text that looks like a merge field stays text). Placeholders
//     that are expressions ({{ a.b }}, {% if %}, filters) are left for dAHKxodusPro's Jinja.
//   - Content controls whose tag or title is a field's name: the control goes, its text becomes the value.
// Word's own data-source setting (w:mailMerge) is removed, so Word doesn't ask for the data source. Other fields
// (PAGE, TOC, REF to ordinary bookmarks) are left as they are.
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

public enum MergeFieldKind { Merge, FillIn, Ask, Placeholder, Control }

/// A field the user fills in: its name (the value's key), the label to show, its default, and whether an IF tests it as yes/no
/// (the dialog offers a check box for a True/False one).
public sealed record MergeFieldInfo(string Name, MergeFieldKind Kind, string Label, string Default, bool InCondition);

public static class Merge {
    static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    static readonly XName R = W + "r", P = W + "p", T = W + "t", RPr = W + "rPr", PPr = W + "pPr",
        FldChar = W + "fldChar", InstrText = W + "instrText", FldSimple = W + "fldSimple", Sdt = W + "sdt";
    static readonly XName FldCharType = W + "fldCharType", Instr = W + "instr", Val = W + "val";

    /// The placeholder markers: [Merge] Markers, pairs "open close" separated by "|"; {{ }} and << >> at first.
    /// « » never: «Name» is how Word shows a MERGEFIELD, and plain text that looks like one (a footer typed by hand, a
    /// field unlinked) must stay plain text, as it does in Word.
    public static List<(string open, string close)> Markers {
        get {
            var text = Settings.Get("Merge", "Markers");
            if (text == "") text = "{{ }}|<< >>";
            return text.Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(p => p.Split(' ', 2, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                .Where(p => p.Length == 2 && p[0] != "«").Select(p => (p[0], p[1])).ToList();
        }
    }

    // ---- The parts a merge looks at

    static IEnumerable<OpenXmlPart> Parts(WordprocessingDocument doc) {
        var main = doc.MainDocumentPart;
        if (main == null) yield break;
        yield return main;
        foreach (var h in main.HeaderParts) yield return h;
        foreach (var f in main.FooterParts) yield return f;
        if (main.FootnotesPart != null) yield return main.FootnotesPart;
        if (main.EndnotesPart != null) yield return main.EndnotesPart;
    }

    static XDocument Load(OpenXmlPart part) {
        using var s = part.GetStream(FileMode.Open, FileAccess.Read);
        return XDocument.Load(s, LoadOptions.PreserveWhitespace);
    }

    static void Store(OpenXmlPart part, XDocument xml) {
        using var s = part.GetStream(FileMode.Create, FileAccess.Write);
        xml.Save(s, SaveOptions.DisableFormatting);
    }

    // ---- Finding the fields

    /// The fields to fill in the package, in document order, each name once. Empty for no fields (or no package).
    public static List<MergeFieldInfo> Scan(byte[] bytes) {
        var found = new List<MergeFieldInfo>();
        var conditions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var notYesNo = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        void Add(MergeFieldInfo f) {
            if (f.Name == "") return;
            var i = found.FindIndex(x => string.Equals(x.Name, f.Name, StringComparison.OrdinalIgnoreCase));
            if (i < 0) found.Add(f);
            else if (found[i].Default == "" && f.Default != "") found[i] = found[i] with { Default = f.Default };
        }
        try {
            using var ms = new MemoryStream(bytes, false);
            using var doc = WordprocessingDocument.Open(ms, false);
            var markers = Markers;
            if (IsJinja(doc)) markers.RemoveAll(m => m.open == "{{");
            foreach (var part in Parts(doc)) {
                var xml = Load(part);
                foreach (var field in FieldTree.Read(xml.Root!)) Visit(field, false);
                foreach (var simple in xml.Descendants(FldSimple)) AddInstruction((string?)simple.Attribute(Instr) ?? "", false);
                foreach (var p in xml.Descendants(P))
                    foreach (var (name, def) in PlaceholdersIn(ParagraphText(p).text, markers))
                        Add(new(name, MergeFieldKind.Placeholder, name, def, false));
                foreach (var sdt in xml.Descendants(Sdt)) {
                    var pr = sdt.Element(W + "sdtPr");
                    var name = ControlName(pr);
                    if (name != "") Add(new(name, MergeFieldKind.Control, (string?)pr?.Element(W + "alias")?.Attribute(Val) ?? name, "", false));
                }
            }
        } catch (Exception ex) {
            MainForm.Log("merge scan: " + ex);
        }
        return found.Select(f => conditions.Contains(f.Name) && !notYesNo.Contains(f.Name) ? f with { InCondition = true } : f).ToList();

        void Visit(Field field, bool inIf) {
            var code = field.PlainCode();
            var isIf = FieldWord(code) == "IF";
            foreach (var child in field.Children) Visit(child, inIf || isIf);
            AddInstruction(code, inIf);
            if (isIf) {
                // (the merge fields in its condition, before its first quoted text: not those in its texts)
                var quote = code.IndexOfAny(['"', '“', '”', '‘']);
                var condition = quote < 0 ? code : code[..quote];
                foreach (var child in field.Children.Where(c => FieldWord(c.PlainCode()) == "MERGEFIELD")) {
                    var marker = "{" + child.PlainCode() + "}";
                    if (!condition.Contains(marker)) continue;
                    // (a check box only for a yes/no field: tested alone, or compared with True/False, Yes/No, 1/0)
                    var name = MergeName(child.PlainCode());
                    conditions.Add(name);
                    var rest = code[(code.IndexOf(marker) + marker.Length)..];
                    var cmp = Regex.Match(rest, "^\\s*(?:=|<>|<=|>=|<|>)\\s*[\"“”]?([^\\s\"“”]*)");
                    if (cmp.Success && !Regex.IsMatch(cmp.Groups[1].Value, "^(true|false|yes|no|y|n|t|f|1|0|-1)$", RegexOptions.IgnoreCase))
                        notYesNo.Add(name);
                }
            }
        }

        void AddInstruction(string code, bool inIf) {
            switch (FieldWord(code)) {
                case "MERGEFIELD":
                    Add(new(MergeName(code), MergeFieldKind.Merge, MergeName(code), "", false));
                    break;
                case "FILLIN": {
                    var args = Arguments(code);
                    var prompt = args.Count > 1 ? args[1] : "";
                    Add(new(prompt, MergeFieldKind.FillIn, prompt, Switch(code, 'd') ?? "", false));
                    break;
                }
                case "ASK": {
                    var args = Arguments(code);
                    if (args.Count > 1) Add(new(args[1], MergeFieldKind.Ask, args.Count > 2 ? args[2] : args[1], Switch(code, 'd') ?? "", false));
                    break;
                }
            }
        }
    }

    // ---- Filling

    /// The package with its fields filled from values (keys as Scan names them; a missing one is empty).
    public static byte[] Fill(byte[] bytes, IReadOnlyDictionary<string, string> values) {
        var map = new Dictionary<string, string>(values, StringComparer.OrdinalIgnoreCase);
        using var ms = new MemoryStream();
        ms.Write(bytes, 0, bytes.Length);
        using (var doc = WordprocessingDocument.Open(ms, true)) {
            var markers = Markers;
            if (IsJinja(doc)) markers.RemoveAll(m => m.open == "{{");
            // ASK fields first, anywhere: their bookmarks' values for the REF fields
            var asked = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var part in Parts(doc))
                foreach (var f in FieldTree.Read(Load(part).Root!).SelectMany(Flatten))
                    if (FieldWord(f.PlainCode()) == "ASK" && Arguments(f.PlainCode()) is { Count: > 1 } a)
                        asked[a[1]] = Get(map, a[1]) is { Length: > 0 } v ? v : Switch(f.PlainCode(), 'd') ?? "";
            foreach (var part in Parts(doc)) {
                var xml = Load(part);
                var root = xml.Root!;
                // fields, innermost first, again after each change (an IF's texts change paragraphs)
                for (var guard = 0; guard < 10000; guard++) {
                    var field = FieldTree.Read(root).SelectMany(Flatten).FirstOrDefault(f => f.Children.Count == 0 && Handles(f, asked));
                    if (field == null) break;
                    Evaluate(field, map, asked);
                }
                foreach (var simple in root.Descendants(FldSimple).ToList()) EvaluateSimple(simple, map, asked);
                FillControls(root, map);
                foreach (var p in root.Descendants(P).ToList()) FillPlaceholders(p, map, markers);
                Store(part, xml);
            }
            // (Word's data source: the merged document has none)
            if (doc.MainDocumentPart?.DocumentSettingsPart is { } settings) {
                var xml = Load(settings);
                var mm = xml.Root?.Element(W + "mailMerge");
                if (mm != null) { mm.Remove(); Store(settings, xml); }
            }
        }
        return ms.ToArray();
    }

    static string Get(Dictionary<string, string> map, string name) => map.TryGetValue(name, out var v) ? v ?? "" : "";

    static IEnumerable<Field> Flatten(Field f) => f.Children.SelectMany(Flatten).Append(f);

    // The fields a merge replaces; others (PAGE, TOC, a REF to an ordinary bookmark) stay. An IF whose texts hold a
    // field that stays (a PAGE) is replaced too: that field goes through in the chosen text.
    static bool Handles(Field f, Dictionary<string, string> asked) {
        var code = f.PlainCode();
        return FieldWord(code) switch {
            "MERGEFIELD" or "FILLIN" or "ASK" or "IF" or "MERGEREC" or "MERGESEQ" => true,
            "=" => f.Children.Count == 0 && Formula(code) != null,
            "REF" => Arguments(code) is { Count: > 1 } a && asked.ContainsKey(a[1]),
            _ => false,
        };
    }

    // Leaf fields only (FieldTree is read again after each): a field whose children were replaced has none left.
    static void Evaluate(Field f, Dictionary<string, string> map, Dictionary<string, string> asked) {
        var code = f.PlainCode();
        switch (FieldWord(code)) {
            case "MERGEFIELD": f.ReplaceWithText(Format(Get(map, MergeName(code)), code)); break;
            case "FILLIN": {
                var a = Arguments(code);
                var v = a.Count > 1 ? Get(map, a[1]) : "";
                f.ReplaceWithText(Format(v != "" ? v : Switch(code, 'd') ?? "", code));
                break;
            }
            case "ASK": f.ReplaceWithText(""); break;
            case "REF": f.ReplaceWithText(Format(asked.TryGetValue(Arguments(code)[1], out var r) ? r : "", code)); break;
            case "MERGEREC" or "MERGESEQ": f.ReplaceWithText("1"); break;
            case "IF": f.ReplaceWithAtoms(EvaluateIf(f.CodeAtoms())); break;
            case "=": f.ReplaceWithText(Format(Formula(code)!.Value.ToString(CultureInfo.InvariantCulture), code)); break;
        }
    }

    // An = field's arithmetic (numbers, + - * / ^ %, parentheses; $ and , in numbers ignored): null when it is
    // something else (functions, bookmarks), which is left as it is.
    static double? Formula(string code) {
        var text = code.TrimStart();
        if (!text.StartsWith('=')) return null;
        text = text[1..];
        var cut = text.IndexOf('\\');
        if (cut >= 0) text = text[..cut];
        text = text.Replace("$", "").Replace(",", "").Replace(" ", "");
        var i = 0;
        double? Sum() {
            var v = Product();
            while (v != null && i < text.Length && (text[i] == '+' || text[i] == '-')) {
                var op = text[i++];
                var r = Product();
                if (r == null) return null;
                v = op == '+' ? v + r : v - r;
            }
            return v;
        }
        double? Product() {
            var v = Power();
            while (v != null && i < text.Length && (text[i] == '*' || text[i] == '/')) {
                var op = text[i++];
                var r = Power();
                if (r == null || (op == '/' && r == 0)) return null;
                v = op == '*' ? v * r : v / r;
            }
            return v;
        }
        double? Power() {
            var v = Unary();
            if (v != null && i < text.Length && text[i] == '^') { i++; var r = Power(); return r == null ? null : Math.Pow(v.Value, r.Value); }
            return v;
        }
        double? Unary() {
            if (i < text.Length && text[i] == '-') { i++; return -Unary(); }
            if (i < text.Length && text[i] == '+') { i++; return Unary(); }
            double? v;
            if (i < text.Length && text[i] == '(') {
                i++;
                v = Sum();
                if (i >= text.Length || text[i] != ')') return null;
                i++;
            } else {
                var start = i;
                while (i < text.Length && (char.IsDigit(text[i]) || text[i] == '.')) i++;
                if (!double.TryParse(text[start..i], NumberStyles.Float, CultureInfo.InvariantCulture, out var n)) return null;
                v = n;
            }
            if (v != null && i < text.Length && text[i] == '%') { i++; v /= 100; }
            return v;
        }
        if (text == "") return null;
        var result = Sum();
        return i == text.Length ? result : null;
    }

    static void EvaluateSimple(XElement simple, Dictionary<string, string> map, Dictionary<string, string> asked) {
        var code = (string?)simple.Attribute(Instr) ?? "";
        string? value = FieldWord(code) switch {
            "MERGEFIELD" => Format(Get(map, MergeName(code)), code),
            "FILLIN" => Arguments(code) is { Count: > 1 } a ? (Get(map, a[1]) is { Length: > 0 } v ? v : Switch(code, 'd') ?? "") : "",
            "ASK" => "",
            "=" when Formula(code) is double d => Format(d.ToString(CultureInfo.InvariantCulture), code),
            "REF" when Arguments(code) is { Count: > 1 } a && asked.TryGetValue(a[1], out var r) => Format(r, code),
            _ => null,
        };
        if (value == null) return;
        var rPr = simple.Descendants(RPr).FirstOrDefault();
        simple.ReplaceWith(TextRuns(value, rPr));
    }

    // ---- IF

    // IF expression1 operator expression2 "true text" "false text" (or IF expression "true text": true unless empty,
    // 0 or False). Atoms are the field code's characters with their run formatting, and paragraph ends.
    static List<Atom> EvaluateIf(List<Atom> atoms) {
        var pos = 0;
        List<Atom>? Next() {
            while (pos < atoms.Count && atoms[pos].IsSpace) pos++;
            if (pos >= atoms.Count) return null;
            var c = atoms[pos].Char;
            // (a quoted text: straight or curly quotes, which Word's AutoFormat mixes; single curly ones too)
            var closers = c switch { '"' or '\u201C' or '\u201D' => "\"\u201D", '\u2018' => "\u2019", _ => "" };
            var token = new List<Atom>();
            if (closers != "") {
                pos++;
                while (pos < atoms.Count && (atoms[pos].IsBreak || atoms[pos].Whole != null || !closers.Contains(atoms[pos].Char))) token.Add(atoms[pos++]);
                pos++;
                return token;
            }
            if ("=<>".Contains(c)) {
                while (pos < atoms.Count && "=<>".Contains(atoms[pos].Char)) token.Add(atoms[pos++]);
                return token;
            }
            while (pos < atoms.Count && !atoms[pos].IsSpace && !"=<>".Contains(atoms[pos].Char)) token.Add(atoms[pos++]);
            return token;
        }
        string Text(List<Atom>? a) => a == null ? "" : new string(a.Where(x => !x.IsBreak).Select(x => x.Char).ToArray());

        Next();    // IF
        var left = Next();
        var op = Next();
        bool truth;
        List<Atom>? yes, no;
        if (op != null && Text(op) is "=" or "<>" or "<" or ">" or "<=" or ">=") {
            var right = Next();
            yes = Next();
            no = Next();
            truth = Compare(Text(left), Text(op), Text(right));
        } else {
            // (no operator: op was the true text)
            yes = op;
            no = Next();
            var v = Text(left).Trim();
            truth = v != "" && v != "0" && !v.Equals("False", StringComparison.OrdinalIgnoreCase);
        }
        return (truth ? yes : no) ?? new List<Atom>();
    }

    static bool Compare(string a, string op, string b) {
        int cmp;
        if (decimal.TryParse(Number(a), NumberStyles.Any, CultureInfo.CurrentCulture, out var x) && decimal.TryParse(Number(b), NumberStyles.Any, CultureInfo.CurrentCulture, out var y))
            cmp = x.CompareTo(y);
        else if (op is "=" or "<>" && (b.Contains('*') || b.Contains('?'))) {
            var like = Regex.IsMatch(a, "^" + Regex.Escape(b).Replace(@"\*", ".*").Replace(@"\?", ".") + "$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            return op == "=" ? like : !like;
        } else cmp = string.Compare(a, b, StringComparison.OrdinalIgnoreCase);
        return op switch { "=" => cmp == 0, "<>" => cmp != 0, "<" => cmp < 0, ">" => cmp > 0, "<=" => cmp <= 0, _ => cmp >= 0 };
    }

    static string Number(string s) => s.Replace("$", "").Replace("%", "").Trim();

    // ---- Field codes

    static string FieldWord(string code) {
        var t = code.TrimStart();
        if (t.StartsWith('=')) return "=";
        var end = 0;
        while (end < t.Length && !char.IsWhiteSpace(t[end]) && t[end] != '\\') end++;
        return t[..end].ToUpperInvariant();
    }

    /// The code's words and quoted texts, switches left out (MERGEFIELD Name \* Upper -> [MERGEFIELD, Name]).
    static List<string> Arguments(string code) {
        var list = new List<string>();
        foreach (Match m in Regex.Matches(code, "\\\\\\S\\s*(?:\"[^\"]*\"|[\u201C][^\u201D]*[\u201D]|\\S+)?|\"([^\"]*)\"|[\u201C]([^\u201D]*)[\u201D]|(\\S+)")) {
            if (m.Value.StartsWith('\\')) continue;
            list.Add(m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Success ? m.Groups[2].Value : m.Groups[3].Value);
        }
        return list;
    }

    static string MergeName(string code) => Arguments(code) is { Count: > 1 } a ? a[1] : "";

    /// A switch's argument (\d "default" -> default); null when the code has no such switch.
    static string? Switch(string code, char name) {
        var m = Regex.Match(code, "\\\\" + Regex.Escape(name.ToString()) + "\\s*(?:\"([^\"]*)\"|[\u201C]([^\u201D]*)[\u201D]|([^\\s\\\\]+))", RegexOptions.IgnoreCase);
        return !m.Success ? null : m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Success ? m.Groups[2].Value : m.Groups[3].Value;
    }

    /// A value as the field's switches show it: \# number, \@ date, \* case, \b before / \f after (only when not empty).
    static string Format(string value, string code) {
        if (Switch(code, '#') is { } picture && decimal.TryParse(Number(value), NumberStyles.Any, CultureInfo.CurrentCulture, out var n))
            value = NumberPicture(n, picture);
        else if (Switch(code, '@') is { } datePicture && DateTime.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.None, out var d))
            value = d.ToString(DatePicture(datePicture), CultureInfo.CurrentCulture);
        foreach (Match m in Regex.Matches(code, @"\\\*\s*(\w+)")) {
            value = m.Groups[1].Value.ToLowerInvariant() switch {
                "upper" => value.ToUpper(CultureInfo.CurrentCulture),
                "lower" => value.ToLower(CultureInfo.CurrentCulture),
                "caps" => CultureInfo.CurrentCulture.TextInfo.ToTitleCase(value.ToLower(CultureInfo.CurrentCulture)),
                "firstcap" => value.Length > 0 ? char.ToUpper(value[0], CultureInfo.CurrentCulture) + value[1..] : value,
                _ => value,
            };
        }
        if (value != "") {
            if (Switch(code, 'b') is { } before) value = before + value;
            if (Switch(code, 'f') is { } after) value += after;
        }
        return value;
    }

    // Word's number picture: literal text before and after, 0 and # digits, "," grouping, "." decimals ($,0.00).
    static string NumberPicture(decimal n, string picture) {
        var first = picture.IndexOfAny("0#".ToCharArray());
        if (first < 0) return n.ToString(CultureInfo.CurrentCulture);
        var last = picture.LastIndexOfAny("0#".ToCharArray());
        string prefix = picture[..first], body = picture[first..(last + 1)], suffix = picture[(last + 1)..];
        if (prefix.EndsWith(',')) { prefix = prefix[..^1]; body = "," + body; }
        var dot = body.IndexOf('.');
        var decimals = dot < 0 ? 0 : body[(dot + 1)..].Count(c => c is '0' or '#');
        var group = picture.Contains(',');
        return prefix + n.ToString((group ? "#,0" : "0") + (decimals > 0 ? "." + new string('0', decimals) : ""), CultureInfo.CurrentCulture) + suffix;
    }

    // Word's date picture in .NET's terms: the same letters, AM/PM -> tt, quoted text kept.
    static string DatePicture(string picture) =>
        Regex.Replace(picture, "AM/PM|am/pm", "tt").Replace("'", "\\'").Replace("\"", "");

    // ---- Typed placeholders

    // {{Name}}, {{Name|default}}, <<Name>>: names of letters, digits, spaces and _ - ' (a.b, an expression, is left).
    static IEnumerable<(string name, string def)> PlaceholdersIn(string text, List<(string open, string close)> markers) {
        foreach (var (open, close) in markers)
            foreach (Match m in Regex.Matches(text, Regex.Escape(open) + @"\s*([\p{L}\p{N}_][\p{L}\p{N}_ '\-]*?)\s*(?:\|\s*([^{}|]*?)\s*)?" + Regex.Escape(close)))
                yield return (m.Groups[1].Value, m.Groups[2].Value);
    }

    // A Jinja template ({% ... %} anywhere): its {{ }} are expressions for dAHKxodusPro, not placeholders.
    static bool IsJinja(WordprocessingDocument doc) =>
        Parts(doc).Any(part => Load(part).Descendants(P).Any(p => Regex.IsMatch(ParagraphText(p).text, @"\{%.*?%\}")));

    // A paragraph's text from its own runs (not a field's code), with each character's run.
    static (string text, List<(XElement t, int index)> map) ParagraphText(XElement p) {
        var sb = new StringBuilder();
        var map = new List<(XElement, int)>();
        foreach (var r in p.Elements().SelectMany(e => e.Name == R ? new[] { e }
                : e.Name == W + "hyperlink" || e.Name == W + "ins" || e.Name == W + "smartTag" ? e.Elements(R).ToArray() : Array.Empty<XElement>()))
            foreach (var t in r.Elements(T)) {
                var s = t.Value;
                for (var i = 0; i < s.Length; i++) { sb.Append(s[i]); map.Add((t, i)); }
            }
        return (sb.ToString(), map);
    }

    static void FillPlaceholders(XElement p, Dictionary<string, string> map, List<(string open, string close)> markers) {
        for (var guard = 0; guard < 500; guard++) {
            var (text, chars) = ParagraphText(p);
            Match? hit = null;
            string value = "";
            foreach (var (open, close) in markers) {
                var m = Regex.Match(text, Regex.Escape(open) + @"\s*([\p{L}\p{N}_][\p{L}\p{N}_ '\-]*?)\s*(?:\|\s*([^{}|]*?)\s*)?" + Regex.Escape(close));
                if (!m.Success) continue;
                // (a name not given and no default: left as it is, a marker that isn't a field)
                var v = Get(map, m.Groups[1].Value);
                if (v == "" && !map.ContainsKey(m.Groups[1].Value) && !m.Groups[2].Success) continue;
                if (hit == null || m.Index < hit.Index) { hit = m; value = v != "" ? v : m.Groups[2].Value; }
            }
            if (hit == null) return;
            // The value goes into the first character's text element; the rest of the placeholder is taken out.
            var (firstT, firstI) = chars[hit.Index];
            var edits = chars.Skip(hit.Index).Take(hit.Length).GroupBy(c => c.t).ToList();
            foreach (var g in edits) {
                var s = g.Key.Value;
                var from = g.Min(c => c.index);
                var to = g.Max(c => c.index) + 1;
                g.Key.Value = s[..from] + (g.Key == firstT ? value : "") + s[to..];
                PreserveSpace(g.Key);
            }
            if (value.Contains('\n')) SplitLines(firstT);
        }
    }

    static void PreserveSpace(XElement t) {
        if (t.Value.Length > 0 && (char.IsWhiteSpace(t.Value[0]) || char.IsWhiteSpace(t.Value[^1])))
            t.SetAttributeValue(XNamespace.Xml + "space", "preserve");
    }

    // A value of several lines in one w:t: line breaks (w:br) between them, in the same run.
    static void SplitLines(XElement t) {
        var lines = t.Value.Replace("\r\n", "\n").Split('\n');
        var items = new List<XElement>();
        for (var i = 0; i < lines.Length; i++) {
            if (i > 0) items.Add(new XElement(W + "br"));
            var nt = new XElement(T, lines[i]);
            PreserveSpace(nt);
            items.Add(nt);
        }
        t.ReplaceWith(items);
    }

    // ---- Content controls

    // A content control's field name: its tag, its title, or the custom XML node it is bound to (FillDay for
    // /ns0:Loan[1]/ns0:FillDay[1]); "" for none.
    static string ControlName(XElement? pr) {
        if ((string?)pr?.Element(W + "tag")?.Attribute(Val) is { Length: > 0 } tag) return tag;
        if ((string?)pr?.Element(W + "alias")?.Attribute(Val) is { Length: > 0 } alias) return alias;
        var xpath = (string?)pr?.Element(W + "dataBinding")?.Attribute(W + "xpath") ?? "";
        var m = Regex.Match(xpath, @"(?:^|/)(?:[\w.\-]+:)?([\w.\-]+)(?:\[\d+\])?$");
        return m.Success ? m.Groups[1].Value : "";
    }

    static void FillControls(XElement root, Dictionary<string, string> map) {
        foreach (var sdt in root.Descendants(Sdt).Reverse().ToList()) {
            var pr = sdt.Element(W + "sdtPr");
            var name = ControlName(pr);
            if (name == "" || !map.TryGetValue(name, out var value)) continue;
            var content = sdt.Element(W + "sdtContent");
            if (content == null) { sdt.Remove(); continue; }
            var rPr = content.Descendants(RPr).FirstOrDefault(x => x.Parent?.Name == R);
            var paragraphs = content.Elements(P).ToList();
            if (paragraphs.Count > 0) {
                // block: the first paragraph keeps its properties, its runs become the value; the others go
                var first = paragraphs[0];
                first.Elements().Where(e => e.Name != PPr).Remove();
                first.Add(TextRuns(value, rPr));
                sdt.ReplaceWith(first);
            } else sdt.ReplaceWith(TextRuns(value, rPr));
        }
    }

    // ---- Runs

    // A value as runs with the given formatting: lines with w:br between them.
    static List<XElement> TextRuns(string value, XElement? rPr) {
        var run = new XElement(R);
        if (rPr != null) run.Add(new XElement(rPr));
        var lines = value.Replace("\r\n", "\n").Split('\n');
        for (var i = 0; i < lines.Length; i++) {
            if (i > 0) run.Add(new XElement(W + "br"));
            var t = new XElement(T, lines[i]);
            PreserveSpace(t);
            run.Add(t);
        }
        return [run];
    }

    /// A character of a field's code with its run's formatting, or a paragraph's end (Break, with that paragraph's
    /// properties), or an element passed through whole (a field that stays, a picture).
    public sealed record Atom(char Char, XElement? RPr, XElement? Break = null, XElement? Whole = null) {
        public bool IsBreak => Break != null;
        public bool IsSpace => Break == null && Whole == null && char.IsWhiteSpace(Char);
    }

    // ---- Complex fields as a tree

    /// A complex field: its begin, separate (or null) and end runs, and the fields inside its code.
    public sealed class Field {
        public XElement Begin = null!, End = null!;
        public XElement? Separate;
        public readonly List<Field> Children = new();
        /// Every run and paragraph-end of the document in order (FieldTree's list), and this field's place in it.
        public List<XElement> Order = null!;
        public int BeginIndex, SeparateIndex = -1, EndIndex;

        /// The code's text (instrText and text left by fields replaced inside it), children's codes as {code}.
        public string PlainCode() {
            var sb = new StringBuilder();
            var stop = SeparateIndex >= 0 ? SeparateIndex : EndIndex;
            for (var i = BeginIndex + 1; i < stop; i++) {
                var child = Children.FirstOrDefault(c => c.BeginIndex == i);
                if (child != null) { sb.Append('{').Append(child.PlainCode()).Append('}'); i = child.EndIndex; continue; }
                var e = Order[i];
                if (e.Name == R) foreach (var x in e.Elements()) {
                        if (x.Name == InstrText || x.Name == T) sb.Append(x.Value);
                        else if (x.Name == W + "tab") sb.Append('\t');
                    }
                else sb.Append('\n');
            }
            return sb.ToString();
        }

        /// The code as atoms (only for fields with no children left).
        public List<Atom> CodeAtoms() {
            var atoms = new List<Atom>();
            var stop = SeparateIndex >= 0 ? SeparateIndex : EndIndex;
            for (var i = BeginIndex + 1; i < stop; i++) {
                var e = Order[i];
                if (e.Name != R) { atoms.Add(new('\n', null, Break: e)); continue; }
                var rPr = e.Element(RPr);
                foreach (var x in e.Elements()) {
                    if (x.Name == InstrText || x.Name == T) foreach (var c in x.Value) atoms.Add(new(c, rPr));
                    else if (x.Name == W + "tab") atoms.Add(new('\t', rPr));
                    else if (x.Name == W + "br") atoms.Add(new('\v', rPr, Whole: new XElement(R, rPr == null ? null : new XElement(rPr), new XElement(x))));
                    else if (x.Name != RPr && x.Name != FldChar) atoms.Add(new(' ', rPr, Whole: new XElement(R, rPr == null ? null : new XElement(rPr), new XElement(x))));
                }
            }
            return atoms;
        }

        /// The field (begin to end) replaced by text in the formatting of its result's first run (else its begin's).
        public void ReplaceWithText(string value) {
            XElement? rPr = null;
            if (SeparateIndex >= 0)
                for (var i = SeparateIndex + 1; i < EndIndex && rPr == null; i++)
                    if (Order[i].Name == R && Order[i].Element(T) != null) rPr = Order[i].Element(RPr);
            rPr ??= Begin.Element(RPr);
            ReplaceWith(TextRuns(value, rPr).Select(r => (object)r).ToList(), keepBreaks: false);
        }

        /// The field replaced by atoms (an IF's chosen text): characters grouped into runs by formatting, paragraph ends
        /// as paragraph ends.
        public void ReplaceWithAtoms(List<Atom> atoms) {
            var items = new List<object>();
            XElement? run = null;
            XElement? runPr = null;
            var sb = new StringBuilder();
            void Flush() {
                if (run != null && sb.Length > 0) { var t = new XElement(T, sb.ToString()); PreserveSpace(t); run.Add(t); }
                sb.Clear();
                if (run != null && run.Elements().Any(x => x.Name != RPr)) items.Add(run);
                run = null;
            }
            foreach (var a in atoms) {
                if (a.IsBreak) { Flush(); items.Add(a.Break!); continue; }
                if (a.Whole != null) { Flush(); items.Add(a.Whole); continue; }
                if (run == null || !XNode.DeepEquals(runPr, a.RPr)) {
                    Flush();
                    run = new XElement(R);
                    runPr = a.RPr;
                    if (a.RPr != null) run.Add(new XElement(a.RPr));
                }
                if (a.Char == '\t') { if (sb.Length > 0) { var t = new XElement(T, sb.ToString()); PreserveSpace(t); run.Add(t); sb.Clear(); } run.Add(new XElement(W + "tab")); }
                else sb.Append(a.Char);
            }
            Flush();
            ReplaceWith(items, keepBreaks: true);
        }

        // items: runs, and (keepBreaks) paragraph ends (the paragraph elements whose ends they are) between them.
        void ReplaceWith(List<object> items, bool keepBreaks) {
            var firstP = Begin.Parent!;
            var lastP = End.Parent!;
            if (firstP == lastP && !items.OfType<XElement>().Any(e => e.Name == P)) {
                // one paragraph: the runs from begin to end out, the items in
                var inside = Begin.ElementsAfterSelf().TakeWhile(e => e != End).ToList();
                Begin.AddBeforeSelf(items);
                foreach (var e in inside) e.Remove();
                Begin.Remove();
                End.Remove();
                return;
            }
            // Over paragraphs: the first paragraph's content before the field, the items (a paragraph end closes a
            // paragraph with the properties of the paragraph it ended), the last paragraph's content after the field
            // in a paragraph with the last paragraph's properties.
            var before = Begin.ElementsBeforeSelf().Where(e => e.Name != PPr).ToList();
            var after = End.ElementsAfterSelf().ToList();
            var paragraphs = new List<XElement>();
            var current = new XElement(P, firstP.Element(PPr) is { } fp ? new XElement(fp) : null, before.Select(e => new XElement(e)));
            foreach (var item in items) {
                if (item is XElement e && e.Name == P) {
                    // (a paragraph end: the paragraph being built ends with that paragraph's properties)
                    current.Element(PPr)?.Remove();
                    if (e.Element(PPr) is { } pp) current.AddFirst(new XElement(pp));
                    paragraphs.Add(current);
                    current = new XElement(P);
                } else current.Add(item);
            }
            current.Element(PPr)?.Remove();
            if (lastP.Element(PPr) is { } lp) current.AddFirst(new XElement(lp));
            current.Add(after.Select(e => new XElement(e)));
            paragraphs.Add(current);
            // the paragraphs from the first to the last out, the new ones in
            var sibling = firstP.ElementsAfterSelf().TakeWhile(e => e != lastP).ToList();
            firstP.AddBeforeSelf(paragraphs);
            foreach (var e in sibling) e.Remove();
            firstP.Remove();
            lastP.Remove();
        }
    }

    /// The complex fields of a part: its runs and paragraph ends in order, fields nested as Word nests them. A field whose
    /// runs aren't a paragraph's own (inside a hyperlink, a tracked insertion) or that crosses a table is left out.
    public static class FieldTree {
        public static List<Field> Read(XElement root) {
            var order = new List<XElement>();
            foreach (var p in root.Descendants(P)) {
                foreach (var r in p.Elements(R)) order.Add(r);
                order.Add(p);    // (the paragraph's end)
            }
            var top = new List<Field>();
            var stack = new Stack<Field>();
            for (var i = 0; i < order.Count; i++) {
                var type = (string?)order[i].Element(FldChar)?.Attribute(FldCharType);
                if (order[i].Name != R || type == null) continue;
                if (type == "begin") {
                    var f = new Field { Begin = order[i], BeginIndex = i, Order = order };
                    if (stack.Count > 0) stack.Peek().Children.Add(f); else top.Add(f);
                    stack.Push(f);
                } else if (type == "separate" && stack.Count > 0) {
                    stack.Peek().Separate = order[i];
                    stack.Peek().SeparateIndex = i;
                } else if (type == "end" && stack.Count > 0) {
                    var f = stack.Pop();
                    f.End = order[i];
                    f.EndIndex = i;
                }
            }
            // (fields not closed, or whose paragraphs aren't siblings, are not touched)
            return top.Where(Whole).ToList();

            static bool Whole(Field f) =>
                f.End != null && f.Begin.Parent!.Parent == f.End.Parent!.Parent && f.Children.All(Whole);
        }
    }
}
