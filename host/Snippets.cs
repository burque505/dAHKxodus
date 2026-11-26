// Snippets (Insert > Snippets): pieces of documents kept to insert again, in one folder, Snippets beside settings.ini.
// A snippet is an .html file (saved from the editor's own copy of the selection: its formatting inline; inserted as
// Keep Source Formatting pastes it) or a .txt file (inserted as plain text, taking the formatting where it lands).
// Fields, footnotes, comments and tracked changes don't come along: a snippet is what a paste would bring.
// Each may have a hotkey: settings.ini's [SnippetHotkeys] (the file name = "Ctrl+Alt+1"), given in the Macros tab's
// Hotkeys dialog with the macros'.
//   - Snippets: the folder, its snippets, saving, renaming, deleting; their hotkeys.
//   - SnippetNameDialog (Save Selection as Snippet…): the new snippet's name.
//   - SnippetsDialog (Manage Snippets…): insert, rename, delete, the folder.
// Loaded by dAHKxodus.cs (#load).
using System.Drawing;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

public static class Snippets {
    public record Entry(string File, string Name, bool Html, string Preview) {
        public string Path => System.IO.Path.Combine(Folder, File);
    }

    /// (Measure runs) a folder of their own, never the user's.
    public static string? OverrideFolder { get; set; }
    public static Dictionary<string, string>? OverrideHotkeys { get; set; }

    public static string Folder => OverrideFolder ?? System.IO.Path.Combine(AppFolders.Roaming, "Snippets");

    static bool IsSnippet(string file) => System.IO.Path.GetExtension(file).ToLowerInvariant() is ".html" or ".htm" or ".txt";

    /// The snippets in the folder, by name; none when there is no folder yet.
    public static List<Entry> List() {
        try {
            if (!Directory.Exists(Folder)) return new();
            return Directory.GetFiles(Folder).Where(IsSnippet)
                .Select(f => new Entry(System.IO.Path.GetFileName(f), System.IO.Path.GetFileNameWithoutExtension(f),
                    !f.EndsWith(".txt", StringComparison.OrdinalIgnoreCase), PreviewOf(f)))
                .OrderBy(e => e.Name, StringComparer.CurrentCultureIgnoreCase).ThenBy(e => e.Html ? 0 : 1)
                .ToList();
        } catch (Exception ex) {
            MainForm.Log("snippets folder: " + ex.Message);
            return new();
        }
    }

    /// The snippet called name: its file name ("Letterhead.html"), or its name alone (.html first, then .htm, .txt).
    public static Entry? Find(string name) {
        var all = List();
        return all.FirstOrDefault(e => e.File.Equals(name, StringComparison.OrdinalIgnoreCase))
            ?? all.FirstOrDefault(e => e.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
    }

    // The start of a snippet's text, for its tooltip.
    static string PreviewOf(string path) {
        try {
            var text = File.ReadAllText(path);
            if (!path.EndsWith(".txt", StringComparison.OrdinalIgnoreCase)) text = TextOf(Fragment(text));
            text = Regex.Replace(text, @"\s+", " ").Trim();
            return text.Length > 120 ? text[..117] + "…" : text;
        } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { return ""; }
    }

    // An HTML snippet's own part: between <!--StartFragment--> and <!--EndFragment--> (as the clipboard marks it),
    // else the body's contents, else all of it.
    public static string Fragment(string html) {
        int start = html.IndexOf("<!--StartFragment-->", StringComparison.OrdinalIgnoreCase),
            end = html.IndexOf("<!--EndFragment-->", StringComparison.OrdinalIgnoreCase);
        if (start >= 0 && end > start) return html[(start + "<!--StartFragment-->".Length)..end];
        var body = Regex.Match(html, @"<body[^>]*>(.*)</body>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
        return body.Success ? body.Groups[1].Value : html;
    }

    static string TextOf(string html) =>
        WebUtility.HtmlDecode(Regex.Replace(Regex.Replace(html, @"<(style|script)[^>]*>.*?</\1>", "", RegexOptions.IgnoreCase | RegexOptions.Singleline), "<[^>]+>", " "));

    /// What to insert: (true, the HTML fragment) for an .html snippet, (false, the text) for a .txt one.
    public static (bool Html, string Content) Read(Entry entry) {
        var text = File.ReadAllText(entry.Path);
        return entry.Html ? (true, Fragment(text)) : (false, text.Replace("\r\n", "\n"));
    }

    public static bool ValidName(string name) =>
        name.Trim() != "" && name.Trim() == name && name.IndexOfAny(System.IO.Path.GetInvalidFileNameChars()) < 0 && !name.EndsWith('.');

    /// The selection kept as name.html: a page of its own (it opens in a browser too), the selection between the
    /// clipboard's fragment marks. Returns the file name.
    public static string Save(string name, string html) {
        Directory.CreateDirectory(Folder);
        var file = name + ".html";
        var page = new StringBuilder()
            .Append("<!DOCTYPE html>\r\n<html>\r\n<head>\r\n<meta charset=\"utf-8\">\r\n<title>")
            .Append(WebUtility.HtmlEncode(name))
            .Append("</title>\r\n</head>\r\n<body>\r\n<!--StartFragment-->")
            .Append(html)
            .Append("<!--EndFragment-->\r\n</body>\r\n</html>\r\n");
        File.WriteAllText(System.IO.Path.Combine(Folder, file), page.ToString(), new UTF8Encoding(false));
        return file;
    }

    /// A snippet under a new name (its extension kept), its hotkey with it. Returns the new file name.
    public static string Rename(Entry entry, string newName) {
        var file = newName + System.IO.Path.GetExtension(entry.File);
        File.Move(entry.Path, System.IO.Path.Combine(Folder, file));
        var keys = HotkeyOf(entry.File);
        if (keys != "") { SetHotkey(entry.File, ""); SetHotkey(file, keys); }
        return file;
    }

    /// A snippet to the Recycle Bin, its hotkey gone.
    public static void Delete(Entry entry) {
        Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile(entry.Path, Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs,
            Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin);
        SetHotkey(entry.File, "");
    }

    public static void OpenFolder() {
        try {
            Directory.CreateDirectory(Folder);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("explorer.exe", $"\"{Folder}\"") { UseShellExecute = true });
        } catch (Exception ex) { MainForm.Log("snippets folder: " + ex.Message); }
    }

    // ---- Hotkeys ([SnippetHotkeys]: file name = keys)

    public static Dictionary<string, string> Hotkeys() {
        if (OverrideHotkeys != null) return new(OverrideHotkeys, StringComparer.OrdinalIgnoreCase);
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (key, value) in Settings.Pairs("SnippetHotkeys"))
            if (MacroHotkeys.Normalize(value) is { Length: > 0 } keys) map[key] = keys;
        return map;
    }

    public static string HotkeyOf(string file) => Hotkeys().GetValueOrDefault(file) ?? "";

    public static void SaveHotkeys(Dictionary<string, string> map) {
        var kept = map.Where(p => p.Value != "").OrderBy(p => p.Key, StringComparer.OrdinalIgnoreCase).ToList();
        if (OverrideHotkeys != null) { OverrideHotkeys = kept.ToDictionary(p => p.Key, p => p.Value, StringComparer.OrdinalIgnoreCase); return; }
        Settings.SetPairs("SnippetHotkeys", kept.Select(p => (p.Key, p.Value)));
    }

    public static void SetHotkey(string file, string keys) {
        var map = Hotkeys();
        if (keys == "") map.Remove(file); else map[file] = keys;
        SaveHotkeys(map);
    }

    /// In the Hotkeys dialog's one list, a snippet's key ("snippet:" + its file name: never a macro's file name).
    public const string HotkeyPrefix = "snippet:";
}

/// Save Selection as Snippet…: the new snippet's name (asking before one of that name is replaced).
public sealed class SnippetNameDialog : Form {
    readonly TextBox name = new() { Width = 320 };
    public string SnippetName => name.Text.Trim();

    public SnippetNameDialog(string preview) {
        Text = Lang.T("snippets.saveTitle");
        MinimizeBox = MaximizeBox = false;
        ShowInTaskbar = false;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        AutoScaleMode = AutoScaleMode.Dpi;
        Font = SystemFonts.MessageBoxFont ?? Font;
        ClientSize = new Size(500, 190);
        Padding = new Padding(12);
        name.Width = 400;

        var table = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Dock = DockStyle.Top };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        name.Anchor = AnchorStyles.Left | AnchorStyles.Right;
        table.Controls.Add(new Label { Text = Lang.T("snippets.name"), AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) });
        table.Controls.Add(name);
        var shown = new Label { Text = preview, AutoSize = true, MaximumSize = new Size(470, 66), ForeColor = SystemColors.GrayText, Margin = new Padding(0, 6, 0, 0) };
        table.Controls.Add(shown);
        table.SetColumnSpan(shown, 2);
        var ok = new Button { Text = Lang.T("settings.ok"), DialogResult = DialogResult.OK, AutoSize = true, MinimumSize = new Size(88, 0) };
        var cancel = new Button { Text = Lang.T("settings.cancel"), DialogResult = DialogResult.Cancel, AutoSize = true, MinimumSize = new Size(88, 0) };
        var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Bottom, Padding = new Padding(0, 10, 0, 0) };
        buttons.Controls.Add(cancel);
        buttons.Controls.Add(ok);
        Controls.Add(table);
        Controls.Add(buttons);
        AcceptButton = ok;
        CancelButton = cancel;

        // (a name from the first words of the selection)
        var words = Regex.Replace(preview, @"[^\p{L}\p{N} ]+", " ").Split(' ', StringSplitOptions.RemoveEmptyEntries).Take(4);
        name.Text = string.Join(" ", words);
        ok.Click += (s, e) => {
            if (!Snippets.ValidName(SnippetName)) {
                MessageBox.Show(this, Lang.T("snippets.badName"), Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                name.Focus();
            } else if (File.Exists(Path.Combine(Snippets.Folder, SnippetName + ".html"))
                && MessageBox.Show(this, Lang.T("snippets.replace", SnippetName), Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) {
                DialogResult = DialogResult.None;
                name.Focus();
            }
        };
        Shown += (s, e) => { name.Focus(); name.SelectAll(); };
    }
}

/// Manage Snippets…: every snippet, to insert, rename (F2) or delete (Del); the folder.
public sealed class SnippetsDialog : Form {
    readonly ListView list = new() {
        View = View.Details, FullRowSelect = true, MultiSelect = false, HideSelection = false, Dock = DockStyle.Fill,
        HeaderStyle = ColumnHeaderStyle.Nonclickable, LabelEdit = true,
    };
    readonly Label preview = new() { AutoSize = false, Height = 54, Dock = DockStyle.Bottom, ForeColor = SystemColors.GrayText, Padding = new Padding(0, 6, 0, 0) };

    /// The snippet to insert once the dialog closes (Insert); null when none.
    public Snippets.Entry? ToInsert { get; private set; }
    /// Whether a snippet was renamed or deleted (the page's list is drawn again).
    public bool Changed { get; private set; }

    public SnippetsDialog() {
        Text = Lang.T("snippets.manageTitle");
        MinimizeBox = false;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        AutoScaleMode = AutoScaleMode.Dpi;
        Font = SystemFonts.MessageBoxFont ?? Font;
        ClientSize = new Size(600, 420);
        MinimumSize = new Size(440, 300);
        Padding = new Padding(12);

        list.Columns.Add(Lang.T("snippets.snippet"), 300);
        list.Columns.Add(Lang.T("snippets.kind"), 120);
        list.Columns.Add(Lang.T("hotkeys.hotkey"), 140);

        var insert = new Button { Text = Lang.T("snippets.insert"), AutoSize = true, MinimumSize = new Size(88, 0) };
        var rename = new Button { Text = Lang.T("snippets.rename"), AutoSize = true, MinimumSize = new Size(88, 0) };
        var delete = new Button { Text = Lang.T("snippets.delete"), AutoSize = true, MinimumSize = new Size(88, 0) };
        var folder = new Button { Text = Lang.T("snippets.folder"), AutoSize = true };
        var close = new Button { Text = Lang.T("snippets.close"), DialogResult = DialogResult.Cancel, AutoSize = true, MinimumSize = new Size(88, 0) };
        var left = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Left, WrapContents = false };
        left.Controls.AddRange([insert, rename, delete, folder]);
        var row = new Panel { Dock = DockStyle.Bottom, Height = 40, Padding = new Padding(0, 8, 0, 0) };
        close.Dock = DockStyle.Right;
        row.Controls.Add(left);
        row.Controls.Add(close);

        Controls.Add(list);
        Controls.Add(preview);
        Controls.Add(row);
        CancelButton = close;

        Fill(null);
        Snippets.Entry? Chosen() => list.SelectedItems.Count == 1 ? (Snippets.Entry)list.SelectedItems[0].Tag! : null;
        void Enable() {
            var on = Chosen() != null;
            insert.Enabled = rename.Enabled = delete.Enabled = on;
            preview.Text = Chosen()?.Preview ?? "";
        }
        list.SelectedIndexChanged += (s, e) => Enable();
        void Insert() { if (Chosen() is { } c) { ToInsert = c; DialogResult = DialogResult.OK; } }
        insert.Click += (s, e) => Insert();
        list.ItemActivate += (s, e) => Insert();
        rename.Click += (s, e) => { if (list.SelectedItems.Count == 1) list.SelectedItems[0].BeginEdit(); };
        list.AfterLabelEdit += (s, e) => {
            e.CancelEdit = true;    // (the list is drawn again from the folder)
            var entry = (Snippets.Entry)list.Items[e.Item].Tag!;
            var name = e.Label?.Trim();
            if (name == null || name == entry.Name) return;
            if (!Snippets.ValidName(name)) { MessageBox.Show(this, Lang.T("snippets.badName"), Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if (File.Exists(Path.Combine(Snippets.Folder, name + Path.GetExtension(entry.File)))) {
                MessageBox.Show(this, Lang.T("snippets.exists", name), Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try {
                var file = Snippets.Rename(entry, name);
                MainForm.Log($"snippet renamed: {entry.File} -> {file}");
                Changed = true;
                BeginInvoke(() => Fill(file));
            } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) {
                MessageBox.Show(this, ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        };
        void Delete() {
            if (Chosen() is not { } c) return;
            if (MessageBox.Show(this, Lang.T("snippets.deleteAsk", c.Name), Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            try {
                Snippets.Delete(c);
                MainForm.Log("snippet deleted: " + c.File);
                Changed = true;
                Fill(null);
            } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or OperationCanceledException) {
                MainForm.Log("snippet not deleted: " + ex.Message);
            }
        }
        delete.Click += (s, e) => Delete();
        list.KeyDown += (s, e) => {
            if (e.KeyCode == Keys.F2 && list.SelectedItems.Count == 1) { list.SelectedItems[0].BeginEdit(); e.Handled = true; }
            else if (e.KeyCode == Keys.Delete) { Delete(); e.Handled = true; }
        };
        folder.Click += (s, e) => Snippets.OpenFolder();
        Shown += (s, e) => list.Focus();

        void Fill(string? select) {
            list.BeginUpdate();
            list.Items.Clear();
            var hotkeys = Snippets.Hotkeys();
            foreach (var entry in Snippets.List()) {
                var item = new ListViewItem([entry.Name, Lang.T(entry.Html ? "snippets.kindHtml" : "snippets.kindText"), hotkeys.GetValueOrDefault(entry.File) ?? ""]) { Tag = entry };
                list.Items.Add(item);
                if (entry.File.Equals(select, StringComparison.OrdinalIgnoreCase)) { item.Selected = item.Focused = true; }
            }
            if (list.Items.Count > 0 && list.SelectedItems.Count == 0) { list.Items[0].Selected = list.Items[0].Focused = true; }
            list.EndUpdate();
            Enable();
            if (list.Items.Count == 0) preview.Text = Lang.T("snippets.none", Snippets.Folder);
        }
    }
}
