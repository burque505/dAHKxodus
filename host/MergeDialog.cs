// The Mail Merge dialog (New from Template on a template with fields, File > Mail Merge...): a box per field
// (host\Merge.cs finds them), a check box for a yes/no one an IF tests, saved records (the values as a JSON file in
// the Records folder beside settings.ini), and a data source (host\MergeData.cs): a list of rows whose columns fill
// the fields they match, shown a row at a time in the boxes, merged as the row shown, all rows in one document, a file
// per row or a PDF per row. A template's data source and its fields' columns are remembered ([Merge Sources],
// [Merge Match]). Loaded by dAHKxodus.cs (#load).
using System.Drawing;
using System.Text.Json;
using System.Windows.Forms;

public static class MergeRecords {
    public static string Folder => Path.Combine(AppFolders.Roaming, "Records");

    /// The saved records' names, sorted; none when there is no folder yet.
    public static List<string> List() {
        try {
            if (!Directory.Exists(Folder)) return new();
            return Directory.GetFiles(Folder, "*.json").Select(f => Path.GetFileNameWithoutExtension(f))
                .OrderBy(n => n, StringComparer.CurrentCultureIgnoreCase).ToList();
        } catch (Exception ex) {
            MainForm.Log("records folder: " + ex.Message);
            return new();
        }
    }

    static string PathOf(string name) => Path.Combine(Folder, name + ".json");

    /// A record's values (field name: text); null when there is none of that name or it can't be read.
    public static Dictionary<string, string>? Load(string name) {
        try {
            if (!File.Exists(PathOf(name))) return null;
            using var doc = JsonDocument.Parse(File.ReadAllText(PathOf(name)));
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var p in doc.RootElement.EnumerateObject())
                values[p.Name] = p.Value.ValueKind switch {
                    JsonValueKind.String => p.Value.GetString() ?? "",
                    JsonValueKind.True => "True",
                    JsonValueKind.False => "False",
                    JsonValueKind.Null => "",
                    _ => p.Value.GetRawText(),
                };
            return values;
        } catch (Exception ex) {
            MainForm.Log($"record {name}: {ex.Message}");
            return null;
        }
    }

    /// Writes the values under the name (the record's earlier values for fields not given are kept: one record can
    /// serve several templates).
    public static void Save(string name, Dictionary<string, string> values) {
        Directory.CreateDirectory(Folder);
        var all = Load(name) ?? new(StringComparer.OrdinalIgnoreCase);
        foreach (var (k, v) in values) all[k] = v;
        File.WriteAllText(PathOf(name), JsonSerializer.Serialize(all, new JsonSerializerOptions { WriteIndented = true }));
    }

    public static void Delete(string name) => File.Delete(PathOf(name));

    /// A name usable as a file name.
    public static bool ValidName(string name) =>
        name.Trim() != "" && name.IndexOfAny(Path.GetInvalidFileNameChars()) < 0;

    /// A yes/no value as a check box reads it.
    public static bool IsYes(string value) =>
        value.Trim().ToLowerInvariant() is "true" or "yes" or "y" or "t" or "1" or "-1";
}

/// What a merge makes from a data source.
public enum MergeOutput { CurrentRow, OneDocument, FilePerRow, PdfPerRow }

public sealed class MergeDialog : Form {
    readonly List<MergeFieldInfo> fields;
    readonly string templateKey;
    readonly Dictionary<string, Control> boxes = new(StringComparer.OrdinalIgnoreCase);
    readonly ComboBox record = new() { Width = 260, DropDownStyle = ComboBoxStyle.DropDown };

    // The data source: the table, where it came from ("" for pasted), the rows left out, the row shown, and each
    // field's column ("" none).
    MergeTable? data;
    string dataPath = "", dataTable = "";
    readonly HashSet<int> skipped = new();
    int row;
    readonly Dictionary<string, string> match = new(StringComparer.OrdinalIgnoreCase);
    readonly Label source = new() { AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 0) };
    readonly FlowLayoutPanel nav = new() { AutoSize = true, WrapContents = false, Visible = false, Margin = new Padding(0, 4, 0, 0) };
    readonly TextBox rowBox = new() { Width = 56, TextAlign = HorizontalAlignment.Right };
    readonly Label rowCount = new() { AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(4, 6, 4, 0) };
    readonly Button recipients = new() { AutoSize = true, Enabled = false }, matchButton = new() { AutoSize = true, Enabled = false },
        clear = new() { AutoSize = true, Enabled = false };
    readonly FlowLayoutPanel outputRow = new() { AutoSize = true, WrapContents = false, Visible = false, Margin = new Padding(0, 6, 0, 0) };
    readonly ComboBox output = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 230 };
    readonly TextBox pattern = new() { Width = 220 };
    readonly Label patternLabel = new() { AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(12, 6, 4, 0) };

    /// The values typed (or the row shown), by field name (a check box: True or False); set when the dialog closes
    /// with OK.
    public Dictionary<string, string> Values { get; } = new(StringComparer.OrdinalIgnoreCase);
    /// With a data source: what to make, and (for files) the file name pattern ({Field}, {#} the row's number).
    public MergeOutput Output { get; private set; } = MergeOutput.CurrentRow;
    public string FilePattern => pattern.Text.Trim();
    /// The rows to merge (data source, not CurrentRow): each one's values.
    public List<Dictionary<string, string>> Rows { get; } = new();

    /// templateKey: the template's (or document's) name, under which its data source and matches are remembered.
    public MergeDialog(string documentName, List<MergeFieldInfo> fields, bool fromTemplate, string templateKey = "") {
        this.fields = fields;
        this.templateKey = templateKey != "" ? templateKey : documentName;
        Text = Lang.T("merge.title", documentName);
        FormBorderStyle = FormBorderStyle.Sizable;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        AutoScaleMode = AutoScaleMode.Dpi;
        Padding = new Padding(12);
        Font = SystemFonts.MessageBoxFont ?? Font;

        // Data source: name  Open... Paste... Recipients... Match Fields... Clear;  |< < [n] of N > >|
        var dataRow = new FlowLayoutPanel { AutoSize = true, WrapContents = true, Margin = new Padding(0) };
        var open = new Button { Text = Lang.T("mergeData.open"), AutoSize = true };
        var paste = new Button { Text = Lang.T("mergeData.paste"), AutoSize = true };
        recipients.Text = Lang.T("mergeData.recipients");
        matchButton.Text = Lang.T("mergeData.match");
        clear.Text = Lang.T("mergeData.clear");
        source.Text = Lang.T("mergeData.none");
        dataRow.Controls.AddRange([new Label { Text = Lang.T("mergeData.source"), AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 6, 0) },
            source, open, paste, recipients, matchButton, clear]);
        open.Click += (s, e) => OpenSource();
        paste.Click += (s, e) => PasteSource();
        recipients.Click += (s, e) => EditRecipients();
        matchButton.Click += (s, e) => MatchFields();
        clear.Click += (s, e) => { SetData(null, "", ""); Remember(); };
        Button NavButton(string text, Func<int> target) {
            var b = new Button { Text = text, Width = 34, AutoSize = false, Height = rowBox.PreferredHeight + 2 };
            b.Click += (s, e) => ShowRow(target());
            return b;
        }
        nav.Controls.AddRange([NavButton("|◀", () => Step(-1, from: -1)), NavButton("◀", () => Step(-1, from: row)), rowBox, rowCount,
            NavButton("▶", () => Step(1, from: row)), NavButton("▶|", () => Step(1, from: data?.Rows.Count ?? 0))]);
        rowBox.KeyDown += (s, e) => {
            if (e.KeyCode != Keys.Enter) return;
            e.SuppressKeyPress = true;
            if (int.TryParse(rowBox.Text, out var n) && data != null) ShowRow(Math.Clamp(n - 1, 0, data.Rows.Count - 1));
        };
        var dataBox = new GroupBox { Text = Lang.T("mergeData.group"), AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Top, Padding = new Padding(8, 4, 8, 6) };
        var dataInner = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false, Dock = DockStyle.Fill };
        dataInner.Controls.Add(dataRow);
        dataInner.Controls.Add(nav);
        dataBox.Controls.Add(dataInner);

        // Record: [name ▾] Save  Delete
        var top = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Top, WrapContents = false, Margin = new Padding(0, 0, 0, 8), Padding = new Padding(0, 8, 0, 4) };
        var recordLabel = new Label { Text = Lang.T("merge.record"), AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 6, 0) };
        var save = new Button { Text = Lang.T("merge.save"), AutoSize = true };
        var delete = new Button { Text = Lang.T("merge.delete"), AutoSize = true };
        top.Controls.AddRange([recordLabel, record, save, delete]);
        foreach (var name in MergeRecords.List()) record.Items.Add(name);
        record.SelectedIndexChanged += (s, e) => { if (record.SelectedItem is string name) Fill(MergeRecords.Load(name)); };
        save.Click += (s, e) => SaveRecord();
        delete.Click += (s, e) => DeleteRecord();

        // The fields: label, box
        var grid = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Top };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        foreach (var f in fields) {
            Control box;
            if (f.InCondition) {
                box = new CheckBox { Text = f.Label, AutoSize = true, Checked = MergeRecords.IsYes(f.Default), Margin = new Padding(3, 4, 3, 4) };
                grid.Controls.Add(new Label { AutoSize = true });
            } else {
                var multi = f.Name.Contains("address", StringComparison.OrdinalIgnoreCase) || f.Name.Contains("legal", StringComparison.OrdinalIgnoreCase);
                var text = new TextBox { Text = f.Default, Width = 340, Anchor = AnchorStyles.Left | AnchorStyles.Right };
                if (multi) { text.Multiline = true; text.AcceptsReturn = true; text.ScrollBars = ScrollBars.Vertical; text.Height = text.Font.Height * 3 + 8; }
                box = text;
                grid.Controls.Add(new Label { Text = f.Label + ":", AutoSize = true, MaximumSize = new Size(260, 0), Anchor = AnchorStyles.Left | AnchorStyles.Top, Margin = new Padding(0, 6, 8, 0) });
            }
            boxes[f.Name] = box;
            grid.Controls.Add(box);
        }
        var scroll = new Panel { AutoScroll = true, Dock = DockStyle.Fill };
        scroll.Controls.Add(grid);

        // Output (with a data source), hint, buttons
        output.Items.AddRange([Lang.T("mergeData.outCurrent"), Lang.T("mergeData.outOne"), Lang.T("mergeData.outFiles"), Lang.T("mergeData.outPdfs")]);
        output.SelectedIndex = 0;
        patternLabel.Text = Lang.T("mergeData.pattern");
        pattern.Text = Settings.Get("Merge Patterns", this.templateKey, "");
        output.SelectedIndexChanged += (s, e) => patternLabel.Visible = pattern.Visible = output.SelectedIndex >= 2;
        patternLabel.Visible = pattern.Visible = false;
        outputRow.Controls.AddRange([new Label { Text = Lang.T("mergeData.output"), AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 6, 0) }, output, patternLabel, pattern]);
        var note = new Label {
            Text = Lang.T(fromTemplate ? "merge.hintTemplate" : "merge.hint"), AutoSize = true, MaximumSize = new Size(620, 0),
            ForeColor = SystemColors.GrayText, Dock = DockStyle.Top, Margin = new Padding(0, 8, 0, 0),
        };
        var ok = new Button { Text = Lang.T("merge.ok"), DialogResult = DialogResult.OK, AutoSize = true, MinimumSize = new Size(88, 0) };
        var cancel = new Button { Text = Lang.T(fromTemplate ? "merge.skip" : "settings.cancel"), DialogResult = DialogResult.Cancel, AutoSize = true, MinimumSize = new Size(88, 0) };
        var folder = new LinkLabel { Text = Lang.T("merge.folder"), AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 8, 0, 0) };
        folder.LinkClicked += (s, e) => {
            try {
                Directory.CreateDirectory(MergeRecords.Folder);
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("explorer.exe", $"\"{MergeRecords.Folder}\"") { UseShellExecute = true });
            } catch (Exception ex) { MainForm.Log("records folder: " + ex.Message); }
        };
        var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 8, 0, 0) };
        buttons.Controls.Add(cancel);
        buttons.Controls.Add(ok);
        var bottom = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Dock = DockStyle.Bottom };
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        bottom.Controls.Add(outputRow);
        bottom.SetColumnSpan(outputRow, 2);
        bottom.Controls.Add(folder);
        bottom.Controls.Add(buttons);
        bottom.Controls.Add(note);
        bottom.SetColumnSpan(note, 2);

        Controls.Add(scroll);
        Controls.Add(bottom);
        Controls.Add(top);
        Controls.Add(dataBox);
        AcceptButton = fields.Any(f => !f.InCondition && boxes[f.Name] is TextBox { Multiline: true }) ? null : ok;
        CancelButton = cancel;

        // Size: the grid's, up to most of the screen; scrolls beyond
        Load += (s, e) => {
            var area = Screen.FromControl(Owner ?? this).WorkingArea;
            var w = Math.Min(area.Width - 80, Math.Max(Math.Max(grid.PreferredSize.Width + SystemInformation.VerticalScrollBarWidth + 30, top.PreferredSize.Width + 30), dataBox.PreferredSize.Width + 30));
            var h = Math.Min(area.Height - 80, grid.PreferredSize.Height + top.PreferredSize.Height + dataBox.PreferredSize.Height + bottom.PreferredSize.Height + Padding.Vertical + 90);
            ClientSize = new Size(Math.Max(w, 560), h);
            MinimumSize = new Size(480, 360);
            CenterToParent();
        };
        // The data source this template had last time (loaded once the dialog shows: Access can take a moment)
        Shown += (s, e) => {
            var remembered = Settings.Get("Merge Sources", this.templateKey, "");
            if (remembered == "") return;
            var parts = remembered.Split('|', 2);
            if (File.Exists(parts[0])) LoadSource(parts[0], parts.Length > 1 ? parts[1] : "", quiet: true);
        };
        ok.Click += (s, e) => Finish();
    }

    // ---- The values

    Dictionary<string, string> Current() {
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (name, box) in boxes)
            values[name] = box is CheckBox c ? (c.Checked ? "True" : "False") : box.Text.Replace("\r\n", "\n");
        return values;
    }

    void Fill(Dictionary<string, string>? values) {
        if (values == null) return;
        foreach (var (name, box) in boxes) {
            if (!values.TryGetValue(name, out var v)) continue;
            if (box is CheckBox c) c.Checked = MergeRecords.IsYes(v);
            else box.Text = box is TextBox { Multiline: true } ? v.Replace("\r\n", "\n").Replace("\n", "\r\n") : v;
        }
    }

    // A row's values: the boxes (fields with no column: typed once, the same in every row), the matched columns over them.
    Dictionary<string, string> ValuesOf(int index, Dictionary<string, string> typed) {
        var values = new Dictionary<string, string>(typed, StringComparer.OrdinalIgnoreCase);
        if (data == null) return values;
        var r = data.Row(index);
        foreach (var (field, column) in match)
            if (column != "" && r.TryGetValue(column, out var v)) values[field] = v;
        return values;
    }

    void Finish() {
        foreach (var (k, v) in Current()) Values[k] = v;
        Output = data == null ? MergeOutput.CurrentRow : (MergeOutput)output.SelectedIndex;
        if (Output == MergeOutput.CurrentRow || data == null) return;
        var typed = Current();
        // (the row shown keeps what was typed over it; the others take their columns)
        for (var i = 0; i < data.Rows.Count; i++)
            if (!skipped.Contains(i)) Rows.Add(i == row ? new Dictionary<string, string>(typed, StringComparer.OrdinalIgnoreCase) : ValuesOf(i, typed));
        if (Output is MergeOutput.FilePerRow or MergeOutput.PdfPerRow) Settings.Set("Merge Patterns", templateKey, FilePattern);
        if (Rows.Count == 0) {
            MessageBox.Show(this, Lang.T("mergeData.noRows"), Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            DialogResult = DialogResult.None;
        }
    }

    // ---- The data source

    void OpenSource() {
        using var dlg = new OpenFileDialog { Filter = Lang.T("mergeData.filter", MergeData.Extensions), InitialDirectory = dataPath != "" ? Path.GetDirectoryName(dataPath) : "" };
        if (dlg.ShowDialog(this) != DialogResult.OK) return;
        List<string> tables;
        try {
            UseWaitCursor = true;
            tables = MergeData.Tables(dlg.FileName);
        } catch (Exception ex) {
            MessageBox.Show(this, Lang.T("mergeData.failed", dlg.FileName, ex.Message), Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        } finally { UseWaitCursor = false; }
        var table = tables.Count == 1 ? tables[0] : ChooseTable(tables);
        if (table == null) return;
        if (LoadSource(dlg.FileName, table, quiet: false)) Remember();
    }

    bool LoadSource(string path, string table, bool quiet) {
        try {
            UseWaitCursor = true;
            var loaded = MergeData.Load(path, table);
            SetData(loaded, path, table);
            return true;
        } catch (Exception ex) {
            MainForm.Log($"merge data {path} {table}: {ex}");
            if (!quiet) MessageBox.Show(this, Lang.T("mergeData.failed", path, ex.Message), Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        } finally { UseWaitCursor = false; }
    }

    void PasteSource() {
        using var dlg = new Form {
            Text = Lang.T("mergeData.pasteTitle"), FormBorderStyle = FormBorderStyle.Sizable, MinimizeBox = false, MaximizeBox = false, ShowInTaskbar = false,
            StartPosition = FormStartPosition.CenterParent, Font = Font, ClientSize = new Size(620, 340), Padding = new Padding(10),
        };
        var note = new Label { Text = Lang.T("mergeData.pasteHint"), AutoSize = true, Dock = DockStyle.Top, MaximumSize = new Size(600, 0) };
        var text = new TextBox { Multiline = true, AcceptsReturn = true, AcceptsTab = true, ScrollBars = ScrollBars.Both, WordWrap = false, Dock = DockStyle.Fill, Font = new Font(FontFamily.GenericMonospace, Font.Size) };
        if (data != null && dataPath == "") text.Text = string.Join("\r\n", new[] { string.Join("\t", data.Columns) }.Concat(data.Rows.Select(r => string.Join("\t", r))));
        var ok = new Button { Text = Lang.T("settings.ok"), DialogResult = DialogResult.OK, AutoSize = true };
        var cancel = new Button { Text = Lang.T("settings.cancel"), DialogResult = DialogResult.Cancel, AutoSize = true };
        var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Bottom };
        buttons.Controls.AddRange([cancel, ok]);
        dlg.Controls.Add(text);
        dlg.Controls.Add(note);
        dlg.Controls.Add(buttons);
        dlg.CancelButton = cancel;
        if (dlg.ShowDialog(this) != DialogResult.OK || text.Text.Trim() == "") return;
        SetData(MergeData.Text(text.Text, Lang.T("mergeData.pasted")), "", "");
    }

    string? ChooseTable(List<string> tables) {
        using var dlg = new Form {
            Text = Lang.T("mergeData.tableTitle"), FormBorderStyle = FormBorderStyle.FixedDialog, MinimizeBox = false, MaximizeBox = false, ShowInTaskbar = false,
            StartPosition = FormStartPosition.CenterParent, Font = Font, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(10),
        };
        var list = new ListBox { Width = 320, Height = 240, IntegralHeight = false };
        list.Items.AddRange(tables.Cast<object>().ToArray());
        list.SelectedIndex = 0;
        var ok = new Button { Text = Lang.T("settings.ok"), DialogResult = DialogResult.OK, AutoSize = true };
        var cancel = new Button { Text = Lang.T("settings.cancel"), DialogResult = DialogResult.Cancel, AutoSize = true };
        list.DoubleClick += (s, e) => dlg.DialogResult = DialogResult.OK;
        var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill };
        buttons.Controls.AddRange([cancel, ok]);
        var outer = new TableLayoutPanel { ColumnCount = 1, AutoSize = true };
        outer.Controls.Add(list);
        outer.Controls.Add(buttons);
        dlg.Controls.Add(outer);
        dlg.AcceptButton = ok;
        dlg.CancelButton = cancel;
        return dlg.ShowDialog(this) == DialogResult.OK ? list.SelectedItem as string : null;
    }

    void SetData(MergeTable? table, string path, string tableName) {
        data = table;
        dataPath = path;
        dataTable = tableName;
        skipped.Clear();
        row = 0;
        match.Clear();
        if (table != null) {
            // remembered matches first (when the column is still there), then same names
            foreach (var pair in Settings.Get("Merge Match", templateKey, "").Split(';', StringSplitOptions.RemoveEmptyEntries)) {
                var kv = pair.Split('=', 2);
                if (kv.Length == 2 && boxes.ContainsKey(kv[0]) && (kv[1] == "" || table.Columns.Contains(kv[1], StringComparer.OrdinalIgnoreCase))) match[kv[0]] = kv[1];
            }
            foreach (var f in fields)
                if (!match.ContainsKey(f.Name))
                    match[f.Name] = table.Columns.FirstOrDefault(c => MergeData.Key(c) == MergeData.Key(f.Name)) ?? "";
        }
        var name = table == null ? Lang.T("mergeData.none")
            : (path != "" ? Path.GetFileName(path) : table.Name) + (tableName != "" ? " › " + tableName : "");
        source.Text = table == null ? name : Lang.T("mergeData.loaded", name, table.Rows.Count, match.Values.Count(c => c != ""), fields.Count);
        recipients.Enabled = matchButton.Enabled = clear.Enabled = table != null;
        nav.Visible = outputRow.Visible = table != null && table.Rows.Count > 0;
        if (table != null && table.Rows.Count > 0) ShowRow(Step(1, from: -1));
    }

    // The row shown in the boxes (the fields with a column; the others keep what was typed).
    void ShowRow(int index) {
        if (data == null || data.Rows.Count == 0 || index < 0 || index >= data.Rows.Count) return;
        row = index;
        var r = data.Row(index);
        Fill(match.Where(m => m.Value != "" && r.ContainsKey(m.Value)).ToDictionary(m => m.Key, m => r[m.Value], StringComparer.OrdinalIgnoreCase));
        rowBox.Text = (index + 1).ToString();
        var included = data.Rows.Count - skipped.Count;
        rowCount.Text = Lang.T("mergeData.of", data.Rows.Count) + (skipped.Count > 0 ? " " + Lang.T("mergeData.included", included) : "")
            + (skipped.Contains(index) ? " " + Lang.T("mergeData.leftOut") : "");
    }

    // The next row included, from a row (-1: before the first; Count: after the last), in a direction; the row shown
    // when there is none.
    int Step(int dir, int from) {
        if (data == null) return 0;
        for (var i = from + dir; i >= 0 && i < data.Rows.Count; i += dir)
            if (!skipped.Contains(i)) return i;
        return row;
    }

    void Remember() {
        Settings.Set("Merge Sources", templateKey, data == null || dataPath == "" ? "" : dataPath + (dataTable != "" ? "|" + dataTable : ""));
    }

    // ---- Recipients: the rows, with a check box each, a filter, and the row to show

    void EditRecipients() {
        if (data == null) return;
        using var dlg = new Form {
            Text = Lang.T("mergeData.recipientsTitle"), FormBorderStyle = FormBorderStyle.Sizable, MinimizeBox = false, ShowInTaskbar = false,
            StartPosition = FormStartPosition.CenterParent, Font = Font, ClientSize = new Size(900, 520), Padding = new Padding(10),
        };
        var view = new DataGridView {
            Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false, RowHeadersVisible = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect, MultiSelect = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
        };
        view.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "", Width = 30, AutoSizeMode = DataGridViewAutoSizeColumnMode.None });
        foreach (var c in data.Columns) view.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = c, ReadOnly = true, SortMode = DataGridViewColumnSortMode.Automatic });
        for (var i = 0; i < data.Rows.Count; i++) {
            var cells = new object[data.Columns.Count + 1];
            cells[0] = !skipped.Contains(i);
            for (var c = 0; c < data.Columns.Count; c++) cells[c + 1] = c < data.Rows[i].Length ? data.Rows[i][c].Replace("\n", " ⏎ ") : "";
            var n = view.Rows.Add(cells);
            view.Rows[n].Tag = i;
        }
        if (row < view.Rows.Count) view.CurrentCell = view.Rows[row].Cells[0];
        var filter = new TextBox { Width = 260, PlaceholderText = Lang.T("mergeData.filter.hint") };
        filter.TextChanged += (s, e) => {
            var words = filter.Text.Trim();
            view.CurrentCell = null;
            foreach (DataGridViewRow r in view.Rows)
                r.Visible = words == "" || data.Rows[(int)r.Tag!].Any(v => v.Contains(words, StringComparison.CurrentCultureIgnoreCase));
        };
        void CheckVisible(bool on) { view.EndEdit(); foreach (DataGridViewRow r in view.Rows) if (r.Visible) r.Cells[0].Value = on; }
        var all = new Button { Text = Lang.T("mergeData.all"), AutoSize = true };
        var none = new Button { Text = Lang.T("mergeData.noneButton"), AutoSize = true };
        all.Click += (s, e) => CheckVisible(true);
        none.Click += (s, e) => CheckVisible(false);
        var topRow = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Top, WrapContents = false };
        topRow.Controls.AddRange([new Label { Text = Lang.T("mergeData.find"), AutoSize = true, Margin = new Padding(0, 6, 6, 0) }, filter, all, none]);
        var ok = new Button { Text = Lang.T("settings.ok"), DialogResult = DialogResult.OK, AutoSize = true, MinimumSize = new Size(88, 0) };
        var cancel = new Button { Text = Lang.T("settings.cancel"), DialogResult = DialogResult.Cancel, AutoSize = true, MinimumSize = new Size(88, 0) };
        var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Bottom, Padding = new Padding(0, 8, 0, 0) };
        buttons.Controls.AddRange([cancel, ok, new Label { Text = Lang.T("mergeData.recipientsHint"), AutoSize = true, ForeColor = SystemColors.GrayText, Margin = new Padding(0, 6, 12, 0) }]);
        dlg.Controls.Add(view);
        dlg.Controls.Add(topRow);
        dlg.Controls.Add(buttons);
        dlg.CancelButton = cancel;
        var show = -1;
        view.CellDoubleClick += (s, e) => { if (e.RowIndex >= 0) { show = (int)view.Rows[e.RowIndex].Tag!; dlg.DialogResult = DialogResult.OK; } };
        if (dlg.ShowDialog(this) != DialogResult.OK) return;
        view.EndEdit();
        skipped.Clear();
        foreach (DataGridViewRow r in view.Rows) if (r.Cells[0].Value is false) skipped.Add((int)r.Tag!);
        if (show < 0 && view.CurrentRow?.Tag is int current) show = current;
        if (show >= 0) skipped.Remove(show);
        ShowRow(show >= 0 ? show : skipped.Contains(row) ? Step(1, from: -1) : row);
    }

    // ---- Match Fields: each field's column

    void MatchFields() {
        if (data == null) return;
        using var dlg = new Form {
            Text = Lang.T("mergeData.matchTitle"), FormBorderStyle = FormBorderStyle.Sizable, MinimizeBox = false, MaximizeBox = false, ShowInTaskbar = false,
            StartPosition = FormStartPosition.CenterParent, Font = Font, Padding = new Padding(10),
        };
        var grid = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
        var choices = new Dictionary<string, ComboBox>(StringComparer.OrdinalIgnoreCase);
        var noColumn = Lang.T("mergeData.noColumn");
        foreach (var f in fields) {
            var combo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 260 };
            combo.Items.Add(noColumn);
            combo.Items.AddRange(data.Columns.Cast<object>().ToArray());
            var col = match.GetValueOrDefault(f.Name) ?? "";
            combo.SelectedIndex = col == "" ? 0 : Math.Max(0, data.Columns.FindIndex(c => string.Equals(c, col, StringComparison.OrdinalIgnoreCase)) + 1);
            grid.Controls.Add(new Label { Text = f.Label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 0) });
            grid.Controls.Add(combo);
            choices[f.Name] = combo;
        }
        var scroll = new Panel { AutoScroll = true, Dock = DockStyle.Fill };
        scroll.Controls.Add(grid);
        var ok = new Button { Text = Lang.T("settings.ok"), DialogResult = DialogResult.OK, AutoSize = true, MinimumSize = new Size(88, 0) };
        var cancel = new Button { Text = Lang.T("settings.cancel"), DialogResult = DialogResult.Cancel, AutoSize = true, MinimumSize = new Size(88, 0) };
        var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Bottom, Padding = new Padding(0, 8, 0, 0) };
        buttons.Controls.AddRange([cancel, ok]);
        dlg.Controls.Add(scroll);
        dlg.Controls.Add(buttons);
        dlg.AcceptButton = ok;
        dlg.CancelButton = cancel;
        var area = Screen.FromControl(this).WorkingArea;
        dlg.ClientSize = new Size(grid.PreferredSize.Width + SystemInformation.VerticalScrollBarWidth + 30, Math.Min(area.Height - 120, grid.PreferredSize.Height + 80));
        if (dlg.ShowDialog(this) != DialogResult.OK) return;
        foreach (var (field, combo) in choices) match[field] = combo.SelectedIndex <= 0 ? "" : (string)combo.SelectedItem!;
        // (remembered for the template: a pair for each field whose column isn't found by its name)
        Settings.Set("Merge Match", templateKey, string.Join(";", match
            .Where(m => m.Value != (data.Columns.FirstOrDefault(c => MergeData.Key(c) == MergeData.Key(m.Key)) ?? ""))
            .Select(m => m.Key + "=" + m.Value)));
        source.Text = Lang.T("mergeData.loaded", (dataPath != "" ? Path.GetFileName(dataPath) : data.Name) + (dataTable != "" ? " › " + dataTable : ""),
            data.Rows.Count, match.Values.Count(c => c != ""), fields.Count);
        ShowRow(row);
    }

    // ---- Records

    void SaveRecord() {
        var name = record.Text.Trim();
        if (!MergeRecords.ValidName(name)) {
            MessageBox.Show(this, Lang.T("merge.badName"), Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            record.Focus();
            return;
        }
        try {
            MergeRecords.Save(name, Current());
            if (!record.Items.Contains(name)) record.Items.Add(name);
            MainForm.Log($"merge: record \"{name}\" saved");
        } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) {
            MessageBox.Show(this, Lang.T("merge.notSaved", ex.Message), Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    void DeleteRecord() {
        var name = record.Text.Trim();
        if (!record.Items.Contains(name)) return;
        if (MessageBox.Show(this, Lang.T("merge.deleteAsk", name), Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK) return;
        try {
            MergeRecords.Delete(name);
            record.Items.Remove(name);
            record.Text = "";
        } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) {
            MessageBox.Show(this, Lang.T("merge.notSaved", ex.Message), Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
