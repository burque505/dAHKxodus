// The Settings dialog (the ribbon's Settings button, Ctrl+,): the user's name and initials, the language, the recent
// list's length, opening the last document at the start, the template new documents start from, and how often a recovery copy is saved. OK writes them to
// settings.ini in one go; MainForm then applies them to the open window. Loaded by dAHKxodus.cs (#load).
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

public sealed class SettingsDialog : Form {
    readonly TextBox name = new() { Width = 260 }, initials = new() { Width = 80 };
    readonly ComboBox language = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 260 };
    readonly NumericUpDown recent = new() { Minimum = 0, Maximum = 30, Width = 70 };
    readonly NumericUpDown autosave = new() { Minimum = 0, Maximum = 120, Width = 70 };
    readonly CheckBox openLast = new() { AutoSize = true };
    readonly ComboBox newFrom = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 260 };
    readonly List<string> newFromNames = new();    // ("" = a blank document; else a file name in Templates.Folder)
    readonly List<string> codes = new();

    public SettingsDialog() {
        Text = Lang.T("settings.title");
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = MinimizeBox = false;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        AutoScaleMode = AutoScaleMode.Dpi;
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        Padding = new Padding(12);
        Font = SystemFonts.MessageBoxFont ?? Font;

        var (fallbackName, fallbackInitials, from) = UserIdentity.ReadFallback();
        name.Text = Settings.Get("User", "Name");
        initials.Text = Settings.Get("User", "Initials");
        name.PlaceholderText = fallbackName;
        initials.PlaceholderText = fallbackInitials;

        // Automatic: Windows's display language when there is a file for it, else English (Lang.Load).
        var available = Lang.Available();
        var windows = System.Globalization.CultureInfo.CurrentUICulture.TwoLetterISOLanguageName;
        var automatic = available.FirstOrDefault(a => a.Code == windows).Name ?? available.FirstOrDefault(a => a.Code == "en").Name ?? "English";
        language.Items.Add(Lang.T("settings.languageAuto", automatic));
        codes.Add("");
        foreach (var (code, langName) in available) { language.Items.Add(langName); codes.Add(code); }
        language.SelectedIndex = Math.Max(0, codes.IndexOf(Settings.Get("General", "Language")));
        recent.Value = RecentFiles.Max;
        autosave.Value = Math.Clamp(Settings.GetInt("General", "AutosaveMinutes", 10), 0, 120);
        openLast.Text = Lang.T("settings.openLast");
        openLast.Checked = Settings.GetBool("General", "OpenLastAtStart", false);
        newFrom.Items.Add(Lang.T("settings.newBlank"));
        newFromNames.Add("");
        foreach (var t in Templates.List()) { newFrom.Items.Add(Path.GetFileNameWithoutExtension(t)); newFromNames.Add(Path.GetFileName(t)); }
        newFrom.SelectedIndex = Math.Max(0, newFromNames.FindIndex(n => string.Equals(n, Settings.Get("General", "NewFrom"), StringComparison.OrdinalIgnoreCase)));

        var grid = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
        void Heading(string text) {
            var l = new Label { Text = text, AutoSize = true, Font = new Font(Font, FontStyle.Bold), Margin = new Padding(0, 10, 0, 4) };
            grid.Controls.Add(l);
            grid.SetColumnSpan(l, 2);
        }
        void Row(string label, Control control) {
            var l = new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(12, 6, 8, 6) };
            grid.Controls.Add(l);
            grid.Controls.Add(control);
            control.Anchor = AnchorStyles.Left;
        }
        void Whole(Control control, int left = 12) {
            control.Margin = new Padding(left, 4, 0, 4);
            grid.Controls.Add(control);
            grid.SetColumnSpan(control, 2);
        }
        Heading(Lang.T("settings.you"));
        Row(Lang.T("settings.name"), name);
        Row(Lang.T("settings.initials"), initials);
        Whole(new Label { Text = Lang.T("settings.empty", fallbackName, from), AutoSize = true, ForeColor = SystemColors.GrayText });
        Heading(Lang.T("settings.general"));
        Row(Lang.T("settings.language"), language);
        Row(Lang.T("settings.recent"), recent);
        Row(Lang.T("settings.autosave"), autosave);
        Row(Lang.T("settings.newFrom"), newFrom);
        Whole(openLast);

        var folder = new LinkLabel { Text = Lang.T("settings.folder"), AutoSize = true, Anchor = AnchorStyles.Left };
        folder.LinkClicked += (s, e) => {
            try {
                Directory.CreateDirectory(AppFolders.Roaming);
                Process.Start(new ProcessStartInfo("explorer.exe", $"\"{AppFolders.Roaming}\"") { UseShellExecute = true });
            } catch (Exception ex) { MainForm.Log("settings folder: " + ex.Message); }
        };
        var ok = new Button { Text = Lang.T("settings.ok"), DialogResult = DialogResult.OK, AutoSize = true, MinimumSize = new Size(88, 0) };
        var cancel = new Button { Text = Lang.T("settings.cancel"), DialogResult = DialogResult.Cancel, AutoSize = true, MinimumSize = new Size(88, 0) };
        var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 14, 0, 0) };
        buttons.Controls.Add(cancel);
        buttons.Controls.Add(ok);
        var bottom = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Dock = DockStyle.Fill };
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        bottom.Controls.Add(folder);
        bottom.Controls.Add(buttons);
        // (one panel that sizes itself, and the form to it: docked panels would take the form's size instead)
        var outer = new TableLayoutPanel { ColumnCount = 1, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Location = new Point(Padding.Left, Padding.Top) };
        outer.Controls.Add(grid);
        outer.Controls.Add(bottom);
        Controls.Add(outer);
        AcceptButton = ok;
        CancelButton = cancel;
        ok.Click += (s, e) => Store();
    }

    void Store() {
        Settings.Update(() => {
            Settings.Set("User", "Name", name.Text.Trim());
            Settings.Set("User", "Initials", initials.Text.Trim());
            Settings.Set("General", "Language", codes[Math.Max(0, language.SelectedIndex)]);
            Settings.Set("General", "RecentMax", ((int)recent.Value).ToString());
            Settings.Set("General", "AutosaveMinutes", ((int)autosave.Value).ToString());
            Settings.Set("General", "OpenLastAtStart", openLast.Checked ? "1" : "0");
            Settings.Set("General", "NewFrom", newFromNames[Math.Max(0, newFrom.SelectedIndex)]);
        });
        RecentFiles.Trim();
    }
}
