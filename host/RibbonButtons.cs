// Your own buttons on the ribbon: macros as buttons in a "My buttons" group on the Macros tab (web\index.html draws
// them from the host's "macros" message). Kept in settings.ini's [RibbonButtons], in order:
//   Button1=<macro file name>|<label>
// (the macro found as the Macros list finds it: the user's own folder first, then the program's).
//   - RibbonButtons: the buttons kept.
//   - RibbonButtonsDialog (the group's Ribbon buttons...): add, rename, remove and move them.
// Loaded by dAHKxodus.cs (#load).
using System.Drawing;
using System.Windows.Forms;

public static class RibbonButtons {
    public record Button(string Macro, string Label);

    /// (Measure runs) buttons of their own, never the user's: kept here, not written.
    public static List<Button>? Override { get; set; }

    public static List<Button> All() {
        if (Override != null) return new(Override);
        var list = new List<Button>();
        foreach (var value in Settings.Values("RibbonButtons")) {
            var bar = value.IndexOf('|');
            var macro = (bar < 0 ? value : value[..bar]).Trim();
            var label = bar < 0 ? "" : value[(bar + 1)..].Trim();
            if (macro != "") list.Add(new Button(macro, label != "" ? label : Path.GetFileNameWithoutExtension(macro)));
        }
        return list;
    }

    public static void Save(IEnumerable<Button> buttons) {
        var kept = buttons.Where(b => b.Macro != "").ToList();
        if (Override != null) { Override = kept; return; }
        Settings.SetValues("RibbonButtons", "Button", kept.Select(b => $"{b.Macro}|{b.Label.Replace("|", "/")}"));
    }
}

/// The Macros tab's Ribbon buttons...: the buttons in order, each a macro and the label it shows.
public sealed class RibbonButtonsDialog : Form {
    readonly ListView list = new() {
        View = View.Details, FullRowSelect = true, MultiSelect = false, HideSelection = false, Dock = DockStyle.Fill,
        HeaderStyle = ColumnHeaderStyle.Nonclickable,
    };
    readonly ComboBox macro = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 240 };
    readonly TextBox label = new() { Width = 200 };
    readonly Button add, change, remove, up, down;

    /// The buttons as the dialog leaves them, in order.
    public List<RibbonButtons.Button> Result =>
        list.Items.Cast<ListViewItem>().Select(i => new RibbonButtons.Button((string)i.Tag!, i.Text)).ToList();

    public RibbonButtonsDialog(IEnumerable<string> macroFiles, List<RibbonButtons.Button> buttons) {
        Text = Lang.T("rbuttons.title");
        MinimizeBox = false;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        AutoScaleMode = AutoScaleMode.Dpi;
        Font = SystemFonts.MessageBoxFont ?? Font;
        ClientSize = new Size(640, 440);
        MinimumSize = new Size(520, 340);
        Padding = new Padding(12);

        var intro = new Label { Text = Lang.T("rbuttons.intro"), AutoSize = true, MaximumSize = new Size(610, 0), Dock = DockStyle.Top, Padding = new Padding(0, 0, 0, 10) };
        list.Columns.Add(Lang.T("rbuttons.label"), 220);
        list.Columns.Add(Lang.T("rbuttons.macro"), 300);
        foreach (var b in buttons) list.Items.Add(Item(b.Macro, b.Label));
        foreach (var m in macroFiles) macro.Items.Add(m);
        if (macro.Items.Count > 0) macro.SelectedIndex = 0;

        Button Small(string text) => new() { Text = text, AutoSize = true, Margin = new Padding(6, 2, 0, 0) };
        add = Small(Lang.T("rbuttons.add"));
        change = Small(Lang.T("rbuttons.change"));
        remove = Small(Lang.T("rbuttons.remove"));
        up = Small(Lang.T("rbuttons.up"));
        down = Small(Lang.T("rbuttons.down"));
        var row = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Bottom, Padding = new Padding(0, 10, 0, 0), WrapContents = true };
        row.Controls.AddRange([
            new Label { Text = Lang.T("rbuttons.macro") + ":", AutoSize = true, Margin = new Padding(0, 6, 4, 0) }, macro,
            new Label { Text = Lang.T("rbuttons.label") + ":", AutoSize = true, Margin = new Padding(10, 6, 4, 0) }, label,
        ]);
        var actions = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Bottom, Padding = new Padding(0, 6, 0, 0) };
        actions.Controls.AddRange([add, change, remove, up, down]);
        var ok = new Button { Text = Lang.T("settings.ok"), DialogResult = DialogResult.OK, AutoSize = true, MinimumSize = new Size(88, 0) };
        var cancel = new Button { Text = Lang.T("settings.cancel"), DialogResult = DialogResult.Cancel, AutoSize = true, MinimumSize = new Size(88, 0) };
        var buttonsRow = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Bottom, Padding = new Padding(0, 10, 0, 0) };
        buttonsRow.Controls.Add(cancel);
        buttonsRow.Controls.Add(ok);

        Controls.Add(list);
        Controls.Add(intro);
        Controls.Add(row);
        Controls.Add(actions);
        Controls.Add(buttonsRow);
        AcceptButton = ok;
        CancelButton = cancel;

        // (a macro chosen: its name is the label, unless the user has typed one of their own)
        var labelTyped = false;
        label.TextChanged += (s, e) => { if (label.Focused) labelTyped = label.Text != ""; };
        macro.SelectedIndexChanged += (s, e) => {
            if (!labelTyped && macro.SelectedItem is string m) label.Text = Path.GetFileNameWithoutExtension(m);
            Enable();
        };
        if (macro.SelectedItem is string first) label.Text = Path.GetFileNameWithoutExtension(first);
        list.SelectedIndexChanged += (s, e) => {
            if (Selected is { } item) {
                var at = macro.Items.Cast<string>().ToList().FindIndex(m => m.Equals((string)item.Tag!, StringComparison.OrdinalIgnoreCase));
                if (at >= 0) macro.SelectedIndex = at;
                label.Text = item.Text;
                labelTyped = true;
            }
            Enable();
        };
        add.Click += (s, e) => {
            if (macro.SelectedItem is not string m) return;
            var item = Item(m, LabelText(m));
            list.Items.Add(item);
            item.Selected = true;
            labelTyped = false;
        };
        change.Click += (s, e) => {
            if (Selected is not { } item || macro.SelectedItem is not string m) return;
            item.Tag = m;
            item.Text = LabelText(m);
            item.SubItems[1].Text = m;
        };
        remove.Click += (s, e) => { if (Selected is { } item) list.Items.Remove(item); Enable(); };
        up.Click += (s, e) => Move(-1);
        down.Click += (s, e) => Move(1);
        Enable();
    }

    ListViewItem? Selected => list.SelectedItems.Count == 1 ? list.SelectedItems[0] : null;

    string LabelText(string m) => label.Text.Trim().Replace("|", "/") is { Length: > 0 } t ? t : Path.GetFileNameWithoutExtension(m);

    static ListViewItem Item(string macro, string label) => new([label, macro]) { Tag = macro };

    void Move(int by) {
        if (Selected is not { } item) return;
        var at = item.Index + by;
        if (at < 0 || at >= list.Items.Count) return;
        list.Items.Remove(item);
        list.Items.Insert(at, item);
        item.Selected = true;
        Enable();
    }

    void Enable() {
        var item = Selected;
        add.Enabled = macro.SelectedItem != null;
        change.Enabled = remove.Enabled = item != null && macro.SelectedItem != null;
        remove.Enabled = item != null;
        up.Enabled = item != null && item.Index > 0;
        down.Enabled = item != null && item.Index < list.Items.Count - 1;
    }
}
