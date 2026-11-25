// Templates (File > New from Template, Save as Template, Settings' "New documents start from"): Word templates
// (.dotx, .dotm) and documents kept in one folder, Templates beside settings.ini. A new document from one is untitled
// and saved under a new name, as in Word; File > Open on a template edits the template itself.
// The editor only ever gets a document: a template's package is turned into a document's on the way in (ToDocument),
// and back into a template's when it is written under a .dotx/.dotm name (ForPath). Loaded by dAHKxodus.cs (#load).
using System.Drawing;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

public static class Templates {
    public static string Folder => Path.Combine(AppFolders.Roaming, "Templates");

    public static bool IsTemplate(string path) =>
        Path.GetExtension(path).ToLowerInvariant() is ".dotx" or ".dotm";

    /// The templates in the folder (.dotx, .dotm, .docx), by name; none when there is no folder yet.
    public static List<string> List() {
        try {
            if (!Directory.Exists(Folder)) return new();
            return Directory.GetFiles(Folder)
                .Where(f => Path.GetExtension(f).ToLowerInvariant() is ".dotx" or ".dotm" or ".docx" && !Path.GetFileName(f).StartsWith("~$"))
                .OrderBy(f => Path.GetFileNameWithoutExtension(f), StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        } catch (Exception ex) {
            MainForm.Log("templates folder: " + ex.Message);
            return new();
        }
    }

    /// The template New starts from ([General] NewFrom, a file name in the folder); null for a blank document, or
    /// when that file is gone.
    public static string? Default {
        get {
            var name = Settings.Get("General", "NewFrom");
            if (name == "") return null;
            var path = Path.Combine(Folder, name);
            return File.Exists(path) ? path : null;
        }
    }

    public static void OpenFolder() {
        try {
            Directory.CreateDirectory(Folder);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("explorer.exe", $"\"{Folder}\"") { UseShellExecute = true });
        } catch (Exception ex) { MainForm.Log("templates folder: " + ex.Message); }
    }

    /// The package as a document (a template's own type changed; macros kept: .dotm -> .docm). Bytes that are no
    /// Word package are given back as they came (the editor then says what is wrong with them).
    public static byte[] ToDocument(byte[] bytes) => Retype(bytes, template: false);

    /// The package typed for the file it is written to: a template under .dotx/.dotm, else a document.
    public static byte[] ForPath(byte[] bytes, string path) => Retype(bytes, template: IsTemplate(path));

    static byte[] Retype(byte[] bytes, bool template) {
        try {
            using var ms = new MemoryStream();
            ms.Write(bytes, 0, bytes.Length);
            using (var doc = WordprocessingDocument.Open(ms, true)) {
                var type = doc.DocumentType;
                var macros = type is WordprocessingDocumentType.MacroEnabledDocument or WordprocessingDocumentType.MacroEnabledTemplate;
                var isTemplate = type is WordprocessingDocumentType.Template or WordprocessingDocumentType.MacroEnabledTemplate;
                if (isTemplate == template) return bytes;
                doc.ChangeDocumentType(template
                    ? (macros ? WordprocessingDocumentType.MacroEnabledTemplate : WordprocessingDocumentType.Template)
                    : (macros ? WordprocessingDocumentType.MacroEnabledDocument : WordprocessingDocumentType.Document));
            }
            return ms.ToArray();
        } catch (Exception ex) {
            MainForm.Log("template type: " + ex.Message);
            return bytes;
        }
    }
}

// File > New from Template...: the templates in the folder; OK (or a double-click) gives the chosen one's path.
public sealed class TemplateDialog : Form {
    readonly ListBox list = new() { Width = 360, Height = 220, IntegralHeight = false };
    readonly List<string> paths;
    public string? Chosen { get; private set; }

    public TemplateDialog() {
        Text = Lang.T("templates.title");
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = MinimizeBox = false;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        AutoScaleMode = AutoScaleMode.Dpi;
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        Padding = new Padding(12);
        Font = SystemFonts.MessageBoxFont ?? Font;

        paths = Templates.List();
        foreach (var p in paths) list.Items.Add(Path.GetFileNameWithoutExtension(p) + (Path.GetExtension(p).ToLowerInvariant() == ".dotx" ? "" : $"  ({Path.GetExtension(p).ToLowerInvariant()})"));
        var def = Templates.Default;
        list.SelectedIndex = paths.Count == 0 ? -1 : Math.Max(0, def == null ? 0 : paths.FindIndex(p => string.Equals(p, def, StringComparison.OrdinalIgnoreCase)));

        var note = new Label {
            Text = paths.Count == 0 ? Lang.T("templates.none", Templates.Folder) : Lang.T("templates.hint"),
            AutoSize = true, MaximumSize = new Size(360, 0), ForeColor = SystemColors.GrayText, Margin = new Padding(0, 6, 0, 0),
        };
        var folder = new LinkLabel { Text = Lang.T("templates.folder"), AutoSize = true, Anchor = AnchorStyles.Left };
        folder.LinkClicked += (s, e) => Templates.OpenFolder();
        var ok = new Button { Text = Lang.T("settings.ok"), DialogResult = DialogResult.OK, AutoSize = true, MinimumSize = new Size(88, 0), Enabled = paths.Count > 0 };
        var cancel = new Button { Text = Lang.T("settings.cancel"), DialogResult = DialogResult.Cancel, AutoSize = true, MinimumSize = new Size(88, 0) };
        var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 14, 0, 0) };
        buttons.Controls.Add(cancel);
        buttons.Controls.Add(ok);
        var bottom = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Dock = DockStyle.Fill };
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        bottom.Controls.Add(folder);
        bottom.Controls.Add(buttons);
        var outer = new TableLayoutPanel { ColumnCount = 1, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Location = new Point(Padding.Left, Padding.Top) };
        outer.Controls.Add(list);
        outer.Controls.Add(note);
        outer.Controls.Add(bottom);
        Controls.Add(outer);
        AcceptButton = ok;
        CancelButton = cancel;
        ok.Click += (s, e) => Chosen = list.SelectedIndex >= 0 ? paths[list.SelectedIndex] : null;
        list.DoubleClick += (s, e) => {
            if (list.SelectedIndex < 0) return;
            Chosen = paths[list.SelectedIndex];
            DialogResult = DialogResult.OK;
        };
    }
}
