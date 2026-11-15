// Macro hotkeys, as AhkWord's: a key combination that runs a macro while dAHKxodus's window is active. Kept in
// settings.ini's [MacroHotkeys] (the macro's file name = "Ctrl+Alt+D"); the page takes the keys before the ribbon and
// the editor see them (web\index.html, macroHotkey) and asks the host to run the macro, as the Macros list does.
//   - MacroHotkeys: the hotkeys kept, their text ("Ctrl+Alt+Shift+D": Ctrl, Alt, Shift, then the key, named as the
//     page names it from the key's virtual-key code), and the questions asked before one is given (Confirm).
//   - HotkeyBox: a box that shows the keys pressed in it.
//   - MacroHotkeysDialog (the Macros tab's Hotkeys…): every macro with its hotkey.
// Loaded by dAHKxodus.cs (#load).
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

public static class MacroHotkeys {
    /// (Measure runs) hotkeys of their own, never the user's: kept here, not written.
    public static Dictionary<string, string>? Override { get; set; }

    /// Every hotkey: the macro's file name ("Insert date.ahk") = its keys.
    public static Dictionary<string, string> All() {
        if (Override != null) return new(Override, StringComparer.OrdinalIgnoreCase);
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (key, value) in Settings.Pairs("MacroHotkeys"))
            if (Normalize(value) is { Length: > 0 } keys) map[key] = keys;
        return map;
    }

    public static string Of(string fileName) => All().GetValueOrDefault(fileName) ?? "";

    public static void Save(Dictionary<string, string> map) {
        var kept = map.Where(p => p.Value != "").OrderBy(p => p.Key, StringComparer.OrdinalIgnoreCase).ToList();
        if (Override != null) { Override = kept.ToDictionary(p => p.Key, p => p.Value, StringComparer.OrdinalIgnoreCase); return; }
        Settings.SetPairs("MacroHotkeys", kept.Select(p => (p.Key, p.Value)));
    }

    /// One macro's hotkey ("" takes it away).
    public static void Set(string fileName, string keys) {
        var map = All();
        if (keys == "") map.Remove(fileName); else map[fileName] = keys;
        Save(map);
    }

    // ---- The keys' text

    /// The key's name in a hotkey (null: not one a hotkey may use).
    public static string? KeyName(Keys key) => key switch {
        >= Keys.A and <= Keys.Z => ((char)key).ToString(),
        >= Keys.D0 and <= Keys.D9 => ((char)key).ToString(),
        >= Keys.F1 and <= Keys.F24 => "F" + (key - Keys.F1 + 1),
        >= Keys.NumPad0 and <= Keys.NumPad9 => "Num" + (key - Keys.NumPad0),
        Keys.Insert => "Insert", Keys.Delete => "Delete", Keys.Home => "Home", Keys.End => "End",
        Keys.PageUp => "PageUp", Keys.PageDown => "PageDown", Keys.Space => "Space", Keys.Pause => "Pause",
        Keys.Up => "Up", Keys.Down => "Down", Keys.Left => "Left", Keys.Right => "Right",
        _ => null,
    };

    /// "Ctrl+Alt+D" for the keys pressed ("" when they can't be a hotkey: no Ctrl or Alt, and not a function key).
    public static string Text(Keys keyData) {
        var name = KeyName(keyData & Keys.KeyCode);
        if (name == null) return "";
        bool ctrl = keyData.HasFlag(Keys.Control), alt = keyData.HasFlag(Keys.Alt), shift = keyData.HasFlag(Keys.Shift);
        if (!ctrl && !alt && !(name.Length > 1 && name[0] == 'F' && char.IsDigit(name[1]))) return "";
        return (ctrl ? "Ctrl+" : "") + (alt ? "Alt+" : "") + (shift ? "Shift+" : "") + name;
    }

    /// A hotkey's text in the usual form: the modifiers in the order Ctrl, Alt, Shift, the key's name as KeyName
    /// gives it ("ctrl+alt+d" -> "Ctrl+Alt+D"); a key it doesn't know keeps its own name ("Ctrl+Shift+>").
    public static string Normalize(string text) {
        var parts = text.Split('+', StringSplitOptions.TrimEntries);
        if (parts.Length == 0 || parts[^1] == "") {    // ("Ctrl++": the plus key)
            if (!text.TrimEnd().EndsWith("++")) return "";
            parts = [.. text.TrimEnd()[..^2].Split('+', StringSplitOptions.TrimEntries), "+"];
        }
        bool ctrl = false, alt = false, shift = false;
        foreach (var p in parts[..^1])
            switch (p.ToLowerInvariant()) {
                case "ctrl" or "control" or "strg": ctrl = true; break;
                case "alt": alt = true; break;
                case "shift" or "mayús" or "mayus": shift = true; break;
                default: return "";
            }
        var key = parts[^1];
        var known = new[] { "Insert", "Delete", "Home", "End", "PageUp", "PageDown", "Space", "Pause", "Up", "Down", "Left", "Right" }
            .FirstOrDefault(k => k.Equals(key, StringComparison.OrdinalIgnoreCase));
        if (known != null) key = known;
        else if (key.Length == 1) key = key.ToUpperInvariant();
        else if (System.Text.RegularExpressions.Regex.IsMatch(key, @"^[fF]\d{1,2}$")) key = key.ToUpperInvariant();
        else if (System.Text.RegularExpressions.Regex.IsMatch(key, @"^num\d$", System.Text.RegularExpressions.RegexOptions.IgnoreCase)) key = "Num" + key[^1];
        return (ctrl ? "Ctrl+" : "") + (alt ? "Alt+" : "") + (shift ? "Shift+" : "") + key;
    }

    [DllImport("user32.dll")] static extern IntPtr GetKeyboardLayout(uint thread);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    static extern int ToUnicodeEx(uint vk, uint scan, byte[] state, StringBuilder buffer, int size, uint flags, IntPtr layout);

    /// What Ctrl+Alt (AltGr) with this key types on the keyboard in use ("@" for AltGr+2 on a Spanish one), or null:
    /// a macro on those keys takes the character away.
    public static string? AltGrCharacter(Keys keyData) {
        if ((keyData & (Keys.Control | Keys.Alt)) != (Keys.Control | Keys.Alt)) return null;
        var state = new byte[256];
        state[0x11] = state[0x12] = 0x80;    // VK_CONTROL, VK_MENU
        state[0xA2] = state[0xA5] = 0x80;    // VK_LCONTROL, VK_RMENU (AltGr)
        if (keyData.HasFlag(Keys.Shift)) state[0x10] = state[0xA0] = 0x80;
        var buffer = new StringBuilder(8);
        // (flag 4: the keyboard's state, a dead key's included, stays as it was)
        var n = ToUnicodeEx((uint)(keyData & Keys.KeyCode), 0, state, buffer, buffer.Capacity, 4, GetKeyboardLayout(0));
        if (n == 0) return null;
        var text = buffer.ToString(0, Math.Max(1, Math.Abs(n)));
        return text.Length > 0 && !char.IsControl(text[0]) ? text : null;
    }

    /// The keys as WinForms names them, from a hotkey's text (Keys.None when it isn't one this box can press).
    public static Keys Parse(string text) {
        var normal = Normalize(text);
        if (normal == "") return Keys.None;
        var parts = normal.Split('+');
        Keys mods = (normal.StartsWith("Ctrl+") ? Keys.Control : 0) | (normal.Contains("Alt+") ? Keys.Alt : 0) | (normal.Contains("Shift+") ? Keys.Shift : 0);
        for (var k = 0; k < 256; k++)
            if (KeyName((Keys)k) == parts[^1]) return (Keys)k | mods;
        return Keys.None;
    }

    /// Before keys become a macro's hotkey: true when they may (the user said yes to each question). taken: the keys
    /// dAHKxodus itself uses (the page's list: "Ctrl+B" = "Bold"); map: every macro's hotkey, from which the keys are
    /// taken off another macro once the user agrees.
    public static bool Confirm(IWin32Window owner, string keys, string fileName, Dictionary<string, string> map,
            IReadOnlyDictionary<string, string> taken) {
        bool Ask(string text) => MessageBox.Show(owner, text, Lang.T("hotkeys.title"), MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
        if (taken.TryGetValue(keys, out var command) && !Ask(Lang.T("hotkeys.taken", keys, command))) return false;
        if (AltGrCharacter(Parse(keys)) is { } character && !Ask(Lang.T("hotkeys.altGr", keys, character))) return false;
        var other = map.FirstOrDefault(p => p.Value == keys && !p.Key.Equals(fileName, StringComparison.OrdinalIgnoreCase)).Key;
        if (other != null) {
            if (!Ask(Lang.T("hotkeys.other", keys, Path.GetFileNameWithoutExtension(other), Path.GetFileNameWithoutExtension(fileName)))) return false;
            map.Remove(other);
        }
        return true;
    }

    /// The page's list of the keys dAHKxodus uses ([{keys, name}]), by their usual text.
    public static Dictionary<string, string> Taken(System.Text.Json.JsonElement list) {
        var taken = new Dictionary<string, string>();
        if (list.ValueKind != System.Text.Json.JsonValueKind.Array) return taken;
        foreach (var item in list.EnumerateArray()) {
            var keys = item.TryGetProperty("keys", out var k) ? Normalize(k.GetString() ?? "") : "";
            if (keys != "" && item.TryGetProperty("name", out var n)) taken.TryAdd(keys, n.GetString() ?? "");
        }
        return taken;
    }
}

/// A box that shows the keys pressed in it ("Ctrl+Alt+D"); Backspace or Delete alone empties it. Tab, Enter and
/// Escape keep their usual work in a dialog.
public sealed class HotkeyBox : TextBox {
    string hotkey = "";
    public event EventHandler? HotkeyChanged;

    public HotkeyBox() {
        ShortcutsEnabled = false;
        PlaceholderText = Lang.T("hotkeys.press");
    }

    public string Hotkey {
        get => hotkey;
        set { hotkey = value; Text = value; }
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData) {
        var key = keyData & Keys.KeyCode;
        var mods = keyData & Keys.Modifiers;
        if (mods == 0 && key is Keys.Tab or Keys.Enter or Keys.Escape) return base.ProcessCmdKey(ref msg, keyData);
        if (mods == Keys.Shift && key == Keys.Tab) return base.ProcessCmdKey(ref msg, keyData);
        if (key is Keys.ControlKey or Keys.ShiftKey or Keys.Menu or Keys.LWin or Keys.RWin) return true;
        if (mods == 0 && key is Keys.Back or Keys.Delete) { Set(""); return true; }
        var text = MacroHotkeys.Text(keyData);
        if (text == "") System.Media.SystemSounds.Beep.Play();
        else Set(text);
        return true;
    }

    void Set(string text) {
        if (text == hotkey) return;
        Hotkey = text;
        HotkeyChanged?.Invoke(this, EventArgs.Empty);
    }

    protected override void OnKeyPress(KeyPressEventArgs e) { e.Handled = true; }
}

/// The Macros tab's Hotkeys…: each macro with its hotkey. Choose one, press the keys in the box (Backspace: none).
public sealed class MacroHotkeysDialog : Form {
    readonly ListView list = new() {
        View = View.Details, FullRowSelect = true, MultiSelect = false, HideSelection = false, Dock = DockStyle.Fill,
        HeaderStyle = ColumnHeaderStyle.Nonclickable,
    };
    readonly HotkeyBox box = new() { Width = 220 };
    readonly Dictionary<string, string> map;
    readonly IReadOnlyDictionary<string, string> taken;

    /// Every macro's hotkey as the dialog leaves it (file name = keys).
    public Dictionary<string, string> Result => map;

    public MacroHotkeysDialog(IEnumerable<(string fileName, string label, bool user)> macros, Dictionary<string, string> hotkeys,
            IReadOnlyDictionary<string, string> taken) {
        map = new(hotkeys, StringComparer.OrdinalIgnoreCase);
        this.taken = taken;
        Text = Lang.T("hotkeys.title");
        MinimizeBox = false;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        AutoScaleMode = AutoScaleMode.Dpi;
        Font = SystemFonts.MessageBoxFont ?? Font;
        ClientSize = new Size(620, 460);
        MinimumSize = new Size(460, 340);
        Padding = new Padding(12);

        var intro = new Label { Text = Lang.T("hotkeys.intro"), AutoSize = true, MaximumSize = new Size(590, 0), Dock = DockStyle.Top, Padding = new Padding(0, 0, 0, 10) };
        list.Columns.Add(Lang.T("hotkeys.macro"), 280);
        list.Columns.Add(Lang.T("hotkeys.hotkey"), 150);
        list.Columns.Add(Lang.T("hotkeys.where"), 130);
        foreach (var (fileName, name, user) in macros) {
            var item = new ListViewItem([name, map.GetValueOrDefault(fileName) ?? "", Lang.T(user ? "hotkeys.mine" : "hotkeys.program")]) { Tag = fileName };
            list.Items.Add(item);
        }

        var row = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Bottom, Padding = new Padding(0, 10, 0, 0), WrapContents = false };
        var label = new Label { Text = Lang.T("hotkeys.box"), AutoSize = true, Margin = new Padding(0, 6, 8, 0) };
        var remove = new Button { Text = Lang.T("hotkeys.remove"), AutoSize = true, Margin = new Padding(8, 2, 0, 0) };
        row.Controls.AddRange([label, box, remove]);
        var ok = new Button { Text = Lang.T("settings.ok"), DialogResult = DialogResult.OK, AutoSize = true, MinimumSize = new Size(88, 0) };
        var cancel = new Button { Text = Lang.T("settings.cancel"), DialogResult = DialogResult.Cancel, AutoSize = true, MinimumSize = new Size(88, 0) };
        var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Bottom, Padding = new Padding(0, 10, 0, 0) };
        buttons.Controls.Add(cancel);
        buttons.Controls.Add(ok);

        Controls.Add(list);
        Controls.Add(intro);
        Controls.Add(row);
        Controls.Add(buttons);
        AcceptButton = ok;
        CancelButton = cancel;

        list.SelectedIndexChanged += (s, e) => {
            var on = list.SelectedItems.Count == 1;
            box.Enabled = remove.Enabled = on;
            box.Hotkey = on ? map.GetValueOrDefault((string)list.SelectedItems[0].Tag!) ?? "" : "";
        };
        box.HotkeyChanged += (s, e) => Assign(box.Hotkey);
        remove.Click += (s, e) => { box.Hotkey = ""; Assign(""); box.Focus(); };
        // (a double click on a macro: the box, to press its keys)
        list.DoubleClick += (s, e) => box.Focus();
        box.Enabled = remove.Enabled = false;
        Shown += (s, e) => {
            if (list.Items.Count > 0) { list.Items[0].Selected = true; list.Items[0].Focused = true; }
            list.Focus();
        };
    }

    // The keys pressed become the chosen macro's hotkey, once the user agrees (another macro's, dAHKxodus's own, an
    // AltGr character); refused, the box shows the macro's hotkey again.
    void Assign(string keys) {
        if (list.SelectedItems.Count != 1) return;
        var fileName = (string)list.SelectedItems[0].Tag!;
        if (keys != "" && !MacroHotkeys.Confirm(this, keys, fileName, map, taken)) {
            box.Hotkey = map.GetValueOrDefault(fileName) ?? "";
            return;
        }
        if (keys == "") map.Remove(fileName); else map[fileName] = keys;
        foreach (ListViewItem item in list.Items)
            item.SubItems[1].Text = map.GetValueOrDefault((string)item.Tag!) ?? "";
    }
}
