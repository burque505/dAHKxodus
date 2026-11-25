// The Mail Merge dialog (New from Template on a template with fields, File > Mail Merge...): a box per field
// (host\Merge.cs finds them), a check box for a yes/no one an IF tests, and saved records: the values as a JSON file in
// the Records folder beside settings.ini, to fill the boxes again next time. Loaded by dAHKxodus.cs (#load).
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

public sealed class MergeDialog : Form {
    readonly List<MergeFieldInfo> fields;
    readonly Dictionary<string, Control> boxes = new(StringComparer.OrdinalIgnoreCase);
    readonly ComboBox record = new() { Width = 260, DropDownStyle = ComboBoxStyle.DropDown };

    /// The values typed, by field name (a check box: True or False); set when the dialog closes with OK.
    public Dictionary<string, string> Values { get; } = new(StringComparer.OrdinalIgnoreCase);

    public MergeDialog(string documentName, List<MergeFieldInfo> fields, bool fromTemplate) {
        this.fields = fields;
        Text = Lang.T("merge.title", documentName);
        FormBorderStyle = FormBorderStyle.Sizable;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        AutoScaleMode = AutoScaleMode.Dpi;
        Padding = new Padding(12);
        Font = SystemFonts.MessageBoxFont ?? Font;

        // Record: [name ▾] Save  Delete
        var top = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Top, WrapContents = false, Margin = new Padding(0, 0, 0, 8) };
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

        // Hint, buttons
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
        bottom.Controls.Add(folder);
        bottom.Controls.Add(buttons);
        bottom.Controls.Add(note);
        bottom.SetColumnSpan(note, 2);

        Controls.Add(scroll);
        Controls.Add(bottom);
        Controls.Add(top);
        AcceptButton = fields.Any(f => !f.InCondition && boxes[f.Name] is TextBox { Multiline: true }) ? null : ok;
        CancelButton = cancel;

        // Size: the grid's, up to most of the screen; scrolls beyond
        Load += (s, e) => {
            var area = Screen.FromControl(Owner ?? this).WorkingArea;
            var w = Math.Min(area.Width - 80, Math.Max(grid.PreferredSize.Width + SystemInformation.VerticalScrollBarWidth + 30, top.PreferredSize.Width + 30));
            var h = Math.Min(area.Height - 80, grid.PreferredSize.Height + top.PreferredSize.Height + bottom.PreferredSize.Height + Padding.Vertical + 60);
            ClientSize = new Size(Math.Max(w, 480), h);
            MinimumSize = new Size(420, 300);
            CenterToParent();
        };
        ok.Click += (s, e) => { foreach (var (k, v) in Current()) Values[k] = v; };
    }

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
