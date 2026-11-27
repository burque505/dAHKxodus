// Mail merge data (the Mail Merge dialog's Data source): a table of rows from a file, the first row (or the keys) its
// column names. CSV/TSV/TXT (the separator found from the first line), Excel (.xlsx, .xlsm: a sheet), JSON (an array
// of objects; nested objects' keys joined with "."), Access (.accdb, .mdb: a table or a saved select query, through
// ADODB and the Access Database Engine, when the computer has it), or text pasted (Excel's copy: tabs).
// Loaded by dAHKxodus.cs (#load); tests\merge checks it.
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

public sealed class MergeTable {
    public string Name { get; init; } = "";
    public List<string> Columns { get; init; } = new();
    public List<string[]> Rows { get; init; } = new();

    /// A row's values by column name (the template's field names match them through MergeData.Key).
    public Dictionary<string, string> Row(int index) {
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var row = Rows[index];
        for (var c = 0; c < Columns.Count; c++) values.TryAdd(Columns[c], c < row.Length ? row[c] : "");
        return values;
    }
}

public static class MergeData {
    /// The file types the Open dialog offers.
    public const string Extensions = "*.csv;*.tsv;*.txt;*.xlsx;*.xlsm;*.json;*.accdb;*.mdb";

    /// A name compared with others: case, spaces, _ - and . left out ("Bo Full" = "bo_full" = "BoFull").
    public static string Key(string name) =>
        new string(name.Where(ch => !char.IsWhiteSpace(ch) && ch is not '_' and not '-' and not '.').ToArray()).ToLowerInvariant();

    static string Ext(string path) => Path.GetExtension(path).ToLowerInvariant();

    /// The tables a file holds: Excel's sheets, Access's tables and select queries; one ("") for the others.
    public static List<string> Tables(string path) => Ext(path) switch {
        ".xlsx" or ".xlsm" => ExcelSheets(path),
        ".accdb" or ".mdb" => AccessTables(path),
        _ => new() { "" },
    };

    public static MergeTable Load(string path, string table = "") => Ext(path) switch {
        ".xlsx" or ".xlsm" => Excel(path, table),
        ".accdb" or ".mdb" => Access(path, table),
        ".json" => Json(File.ReadAllText(path), Path.GetFileNameWithoutExtension(path)),
        _ => Text(ReadText(path), Path.GetFileNameWithoutExtension(path)),
    };

    // ---- Text (CSV, TSV, pasted)

    /// A file's text: its BOM's encoding, else UTF-8 when it is valid UTF-8, else Windows-1252 (Excel's "CSV" on
    /// Western Windows).
    static string ReadText(string path) {
        var bytes = File.ReadAllBytes(path);
        if (bytes.Length >= 2 && (bytes[0] == 0xFF && bytes[1] == 0xFE || bytes[0] == 0xFE && bytes[1] == 0xFF)
            || bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
            return new StreamReader(new MemoryStream(bytes), detectEncodingFromByteOrderMarks: true).ReadToEnd();
        try { return new UTF8Encoding(false, throwOnInvalidBytes: true).GetString(bytes); }
        catch (DecoderFallbackException) {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            return Encoding.GetEncoding(1252).GetString(bytes);
        }
    }

    /// A field list (Mailings > Use a Field List...): a text file of names, one a line, or a CSV whose every value is
    /// a name (as MergeFielder reads them), in order, once each.
    public static List<string> FieldList(string path) {
        var text = ReadText(path);
        IEnumerable<string> names = Ext(path) == ".csv"
            ? Text(text, Path.GetFileNameWithoutExtension(path)) is var t ? t.Columns.Concat(t.Rows.SelectMany(r => r)) : []
            : text.Split('\n');
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        return names.Select(n => n.Trim()).Where(n => n != "" && seen.Add(n)).ToList();
    }

    /// Delimited text: tabs if the first line has one, else ; or , (whichever the first line has more of outside
    /// quotes); "quoted" values may hold separators, line breaks and "" for a quote. Blank lines are left out.
    public static MergeTable Text(string text, string name) {
        var first = text.Split('\n')[0];
        char sep;
        if (first.Contains('\t')) sep = '\t';
        else {
            int commas = 0, semis = 0;
            var quoted = false;
            foreach (var ch in first) {
                if (ch == '"') quoted = !quoted;
                else if (!quoted && ch == ',') commas++;
                else if (!quoted && ch == ';') semis++;
            }
            sep = semis > commas ? ';' : ',';
        }
        var rows = new List<string[]>();
        var row = new List<string>();
        var field = new StringBuilder();
        var inQuotes = false;
        void EndRow() {
            row.Add(field.ToString());
            field.Clear();
            if (row.Any(v => v.Trim() != "")) rows.Add(row.ToArray());
            row = new();
        }
        for (var i = 0; i < text.Length; i++) {
            var ch = text[i];
            if (inQuotes) {
                if (ch == '"') {
                    if (i + 1 < text.Length && text[i + 1] == '"') { field.Append('"'); i++; }
                    else inQuotes = false;
                } else field.Append(ch);
            } else if (ch == '"' && field.Length == 0) inQuotes = true;
            else if (ch == sep) { row.Add(field.ToString()); field.Clear(); }
            else if (ch == '\r') { }
            else if (ch == '\n') EndRow();
            else field.Append(ch);
        }
        EndRow();
        return FromRows(rows, name);
    }

    // The first row as column names (blank ones: Column 1, ...; a repeat gets " 2"), the others as rows.
    static MergeTable FromRows(List<string[]> rows, string name) {
        var table = new MergeTable { Name = name };
        if (rows.Count == 0) return table;
        foreach (var (h, i) in rows[0].Select((h, i) => (h.Trim(), i))) {
            var col = h != "" ? h : $"Column {i + 1}";
            var unique = col;
            for (var n = 2; table.Columns.Contains(unique, StringComparer.OrdinalIgnoreCase); n++) unique = $"{col} {n}";
            table.Columns.Add(unique);
        }
        foreach (var r in rows.Skip(1))
            table.Rows.Add(Enumerable.Range(0, table.Columns.Count).Select(c => c < r.Length ? r[c].Trim() : "").ToArray());
        return table;
    }

    // ---- JSON

    /// An array of objects (or an object holding one, its first such property; or one object: one row). Nested
    /// objects give "a.b" columns; arrays and other values are kept as their JSON text, strings as they are.
    public static MergeTable Json(string text, string name) {
        using var doc = JsonDocument.Parse(text, new JsonDocumentOptions { AllowTrailingCommas = true, CommentHandling = JsonCommentHandling.Skip });
        var root = doc.RootElement;
        IEnumerable<JsonElement> items;
        if (root.ValueKind == JsonValueKind.Array) items = root.EnumerateArray();
        else if (root.ValueKind == JsonValueKind.Object && root.EnumerateObject()
            .FirstOrDefault(p => p.Value.ValueKind == JsonValueKind.Array && p.Value.EnumerateArray().Any(e => e.ValueKind == JsonValueKind.Object)) is { Value.ValueKind: JsonValueKind.Array } list)
            items = list.Value.EnumerateArray();
        else if (root.ValueKind == JsonValueKind.Object) items = new[] { root };
        else throw new InvalidDataException("the JSON is not an object or an array of objects");
        var table = new MergeTable { Name = name };
        var records = new List<Dictionary<string, string>>();
        foreach (var item in items.Where(e => e.ValueKind == JsonValueKind.Object)) {
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            Flatten(item, "", values);
            foreach (var k in values.Keys) if (!table.Columns.Contains(k, StringComparer.OrdinalIgnoreCase)) table.Columns.Add(k);
            records.Add(values);
        }
        foreach (var r in records) table.Rows.Add(table.Columns.Select(c => r.GetValueOrDefault(c) ?? "").ToArray());
        return table;
    }

    static void Flatten(JsonElement e, string prefix, Dictionary<string, string> into) {
        foreach (var p in e.EnumerateObject()) {
            var key = prefix + p.Name;
            switch (p.Value.ValueKind) {
                case JsonValueKind.Object: Flatten(p.Value, key + ".", into); break;
                case JsonValueKind.String: into[key] = p.Value.GetString() ?? ""; break;
                case JsonValueKind.True: into[key] = "True"; break;
                case JsonValueKind.False: into[key] = "False"; break;
                case JsonValueKind.Null: into[key] = ""; break;
                default: into[key] = p.Value.GetRawText(); break;
            }
        }
    }

    // ---- Excel

    static List<string> ExcelSheets(string path) {
        using var doc = SpreadsheetDocument.Open(path, false);
        return doc.WorkbookPart?.Workbook?.Sheets?.Elements<Sheet>().Select(s => s.Name?.Value ?? "").Where(n => n != "").ToList() ?? new();
    }

    /// A sheet (the first if none is named): its cells as Excel shows them, roughly: shared and inline strings,
    /// TRUE/FALSE as True/False, numbers as typed (a formula's last value), dates (a date number format) as the
    /// computer's short date.
    static MergeTable Excel(string path, string sheetName) {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);    // (open in Excel too)
        using var doc = SpreadsheetDocument.Open(fs, false);
        var book = doc.WorkbookPart ?? throw new InvalidDataException("the workbook has no sheets");
        var sheets = book.Workbook?.Sheets?.Elements<Sheet>().ToList() ?? new();
        var sheet = sheets.FirstOrDefault(s => sheetName == "" || s.Name?.Value == sheetName)
            ?? throw new InvalidDataException($"there is no sheet \"{sheetName}\"");
        var part = (WorksheetPart)book.GetPartById(sheet.Id!.Value!);
        var shared = book.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().Select(i => i.InnerText).ToList() ?? new();
        var formats = book.WorkbookStylesPart?.Stylesheet;
        var cellFormats = formats?.CellFormats?.Elements<CellFormat>().ToList() ?? new();
        var custom = formats?.NumberingFormats?.Elements<NumberingFormat>()
            .ToDictionary(f => f.NumberFormatId?.Value ?? 0, f => f.FormatCode?.Value ?? "") ?? new();
        bool IsDate(uint style) {
            if (style >= cellFormats.Count) return false;
            var id = cellFormats[(int)style].NumberFormatId?.Value ?? 0;
            if (id is >= 14 and <= 22 or >= 45 and <= 47) return true;
            if (!custom.TryGetValue(id, out var code)) return false;
            code = System.Text.RegularExpressions.Regex.Replace(code, "\"[^\"]*\"|\\[[^\\]]*\\]|\\\\.", "");
            return code.IndexOfAny(['d', 'y', 'D', 'Y']) >= 0 || code.Contains('m') && !code.Contains("0");
        }
        var rows = new List<string[]>();
        foreach (var r in part.Worksheet?.Descendants<Row>() ?? []) {
            var values = new List<string>();
            foreach (var c in r.Elements<Cell>()) {
                var col = Column(c.CellReference?.Value ?? "");
                while (col >= 0 && values.Count < col) values.Add("");
                var raw = c.CellValue?.Text ?? "";
                var type = c.DataType?.Value;
                string text;
                if (type == CellValues.SharedString) text = int.TryParse(raw, out var si) && si < shared.Count ? shared[si] : "";
                else if (type == CellValues.InlineString) text = c.InlineString?.InnerText ?? "";
                else if (type == CellValues.Boolean) text = raw == "1" ? "True" : "False";
                else if (type == CellValues.String || type == CellValues.Error) text = raw;
                else if (raw != "" && IsDate(c.StyleIndex?.Value ?? 0) && double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var oa)) {
                    var d = DateTime.FromOADate(oa);
                    text = d.TimeOfDay == TimeSpan.Zero ? d.ToString("d", CultureInfo.CurrentCulture) : d.ToString("g", CultureInfo.CurrentCulture);
                } else text = raw;
                values.Add(text);
            }
            rows.Add(values.ToArray());
        }
        return FromRows(rows.Where(r => r.Any(v => v.Trim() != "")).ToList(), sheet.Name?.Value ?? "");
    }

    // "AB12" -> 27 (0-based); -1 for none.
    static int Column(string reference) {
        var n = 0;
        foreach (var ch in reference) {
            if (!char.IsLetter(ch)) break;
            n = n * 26 + (char.ToUpperInvariant(ch) - 'A' + 1);
        }
        return n - 1;
    }

    // ---- Access (ADODB through COM: Windows has it; the Access Database Engine provider comes with Office or as
    // Microsoft's free download)

    /// Whether this computer can read Access files.
    public static bool AccessAvailable => Type.GetTypeFromProgID("Microsoft.ACE.OLEDB.16.0") != null || Type.GetTypeFromProgID("Microsoft.ACE.OLEDB.12.0") != null;

    static object OpenAccess(string path) {
        if (!AccessAvailable) throw new InvalidOperationException(Lang.T("mergeData.noAce"));
        var provider = Type.GetTypeFromProgID("Microsoft.ACE.OLEDB.16.0") != null ? "Microsoft.ACE.OLEDB.16.0" : "Microsoft.ACE.OLEDB.12.0";
        var conn = Activator.CreateInstance(Type.GetTypeFromProgID("ADODB.Connection")!)!;
        Call(conn, "Open", $"Provider={provider};Data Source={path};Mode=Read;");
        return conn;
    }

    static object? Call(object target, string name, params object[] args) =>
        target.GetType().InvokeMember(name, BindingFlags.InvokeMethod, null, target, args);
    static object? Get(object target, string name, params object[] args) =>
        target.GetType().InvokeMember(name, BindingFlags.GetProperty, null, target, args);

    static List<string> AccessTables(string path) {
        var conn = OpenAccess(path);
        try {
            var rs = Call(conn, "OpenSchema", 20)!;    // (adSchemaTables)
            var names = new List<string>();
            while (!(bool)Get(rs, "EOF")!) {
                var fields = Get(rs, "Fields")!;
                var type = Get(Get(fields, "Item", "TABLE_TYPE")!, "Value") as string;
                var name = Get(Get(fields, "Item", "TABLE_NAME")!, "Value") as string ?? "";
                if (type is "TABLE" or "VIEW" && !name.StartsWith("~")) names.Add(name);    // (VIEW: a saved select query)
                Call(rs, "MoveNext");
            }
            Call(rs, "Close");
            return names.OrderBy(n => n, StringComparer.CurrentCultureIgnoreCase).ToList();
        } finally { Call(conn, "Close"); }
    }

    static MergeTable Access(string path, string tableName) {
        if (tableName == "") tableName = AccessTables(path).FirstOrDefault() ?? throw new InvalidDataException("the database has no tables");
        var conn = OpenAccess(path);
        try {
            var rs = Call(conn, "Execute", $"SELECT * FROM [{tableName.Replace("]", "]]")}]")!;
            var fields = Get(rs, "Fields")!;
            var count = (int)Get(fields, "Count")!;
            var table = new MergeTable { Name = tableName };
            for (var i = 0; i < count; i++) table.Columns.Add((string)Get(Get(fields, "Item", i)!, "Name")!);
            while (!(bool)Get(rs, "EOF")!) {
                var row = new string[count];
                for (var i = 0; i < count; i++) row[i] = AccessValue(Get(Get(fields, "Item", i)!, "Value"));
                table.Rows.Add(row);
                Call(rs, "MoveNext");
            }
            Call(rs, "Close");
            return table;
        } finally { Call(conn, "Close"); }
    }

    // Yes/No as True/False (what the IF fields compare with), dates as the computer's short date (with the time when
    // there is one), numbers as stored (the fields' \# formats them).
    static string AccessValue(object? v) => v switch {
        null or DBNull => "",
        bool b => b ? "True" : "False",
        DateTime d => d.TimeOfDay == TimeSpan.Zero ? d.ToString("d", CultureInfo.CurrentCulture) : d.ToString("g", CultureInfo.CurrentCulture),
        decimal m => m.ToString(CultureInfo.InvariantCulture),
        double x => x.ToString(CultureInfo.InvariantCulture),
        float f => f.ToString(CultureInfo.InvariantCulture),
        IFormattable n => n.ToString(null, CultureInfo.InvariantCulture),
        _ => v.ToString() ?? "",
    };
}
