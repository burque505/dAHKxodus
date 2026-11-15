// The Save Macro dialog (after a recording stops: Ctrl+Shift+R, the Macros tab's Stop): the new macro's name, a
// hotkey for it if the user wants one (host\MacroHotkeys.cs), and its script, which can be changed before it is saved.
// MainForm saves it in the user's Macros folder (Macros.Save). Loaded by dAHKxodus.cs (#load).
using System.Drawing;
using System.Windows.Forms;

public sealed class MacroDialog : Form {
    readonly TextBox name = new() { Width = 320 };
    readonly HotkeyBox hotkey = new() { Width = 220 };
    readonly Dictionary<string, string> hotkeys;
    readonly TextBox script = new() {
        Multiline = true, ScrollBars = ScrollBars.Both, WordWrap = false, AcceptsTab = true, AcceptsReturn = true,
        Dock = DockStyle.Fill, Font = new Font("Consolas", 10f),
    };

    public string MacroName => name.Text.Trim();
    public string Script => script.Text;
    /// The hotkey chosen ("": none).
    public string Hotkey => hotkey.Hotkey;
    /// Every macro's hotkey once this one's is given (another macro may have lost it).
    public Dictionary<string, string> Hotkeys => hotkeys;

    public MacroDialog(string suggestedName, string recorded, Dictionary<string, string> hotkeys, IReadOnlyDictionary<string, string> taken) {
        this.hotkeys = new(hotkeys, StringComparer.OrdinalIgnoreCase);
        Text = Lang.T("macro.saveTitle");
        MinimizeBox = false;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        AutoScaleMode = AutoScaleMode.Dpi;
        Font = SystemFonts.MessageBoxFont ?? Font;
        ClientSize = new Size(720, 460);
        MinimumSize = new Size(480, 320);
        Padding = new Padding(12);

        name.Text = suggestedName;
        script.Text = recorded.Replace("\r\n", "\n").Replace("\n", "\r\n");

        var top = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Dock = DockStyle.Top };
        top.Controls.Add(new Label { Text = Lang.T("macro.name"), AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) });
        top.Controls.Add(name);
        top.Controls.Add(new Label { Text = Lang.T("macro.hotkey"), AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) });
        top.Controls.Add(hotkey);
        var scriptLabel = new Label { Text = Lang.T("macro.script"), AutoSize = true, Margin = new Padding(0, 8, 0, 4) };
        top.Controls.Add(scriptLabel);
        top.SetColumnSpan(scriptLabel, 2);

        var ok = new Button { Text = Lang.T("settings.ok"), DialogResult = DialogResult.OK, AutoSize = true, MinimumSize = new Size(88, 0) };
        var cancel = new Button { Text = Lang.T("settings.cancel"), DialogResult = DialogResult.Cancel, AutoSize = true, MinimumSize = new Size(88, 0) };
        var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Bottom, Padding = new Padding(0, 10, 0, 0) };
        buttons.Controls.Add(cancel);
        buttons.Controls.Add(ok);

        // (the script fills what is left: added first, as docking lays out the last-added first)
        Controls.Add(script);
        Controls.Add(top);
        Controls.Add(buttons);
        AcceptButton = ok;
        CancelButton = cancel;
        ok.Click += (s, e) => {
            if (MacroName == "" || MacroName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) {
                MessageBox.Show(this, Lang.T("macro.badName"), Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                name.Focus();
            } else if (Hotkey != "" && !MacroHotkeys.Confirm(this, Hotkey, MacroName + ".ahk", this.hotkeys, taken)) {
                DialogResult = DialogResult.None;
                hotkey.Focus();
            }
        };
        Shown += (s, e) => { name.Focus(); name.SelectAll(); };
    }
}
