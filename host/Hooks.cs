// Event hooks, as AhkWord's (its rte_hooks.cs): a macro that dAHKxodus runs by itself when something happens. The
// program starts, a document is opened, a new one made, a document is about to be saved or has been saved, a PDF
// was exported, the window is about to close. No script has to keep running (as one with Doc.On must). The macro
// is started as the Macros tab starts one (dAHKxodus.ahk runs it, told this window), and is told what happened:
//   A_Args[1]   the event's name ("Saved")
//   A_Args[2]   what it was about: the document's file (Started, Opened, BeforeSave, Saved, Closing; "" when it has
//               none yet), the PDF (PdfExported); "" for NewDocument
//   EnvGet("DAHKXODUS_EVENT"), EnvGet("DAHKXODUS_EVENT_DETAIL")   the same
// BeforeSave and Closing wait for their macro to end (at most WaitSeconds), so that it can change the document before
// it is saved, or save it before the window closes; the others don't wait. While an event's macro runs, that event
// doesn't start it again (a Saved hook that saves). A hook can be switched off: its macro stays chosen, not run.
//   settings.ini [Hooks]: <Event>=<macro> (a file name in the macros folders, the user's first, or a full path),
//   Off=<Event>,<Event> (the ones switched off).
//   - EventHooks: the events, the hooks kept, a new hook macro's starting script.
//   - EventHooksDialog (the Macros tab's Event hooks...): a macro for each event; New... makes one, Edit opens it,
//     Run now tries it.
// Loaded by dAHKxodus.cs (#load).
using System.Diagnostics;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

/// The events a hook can run on.
public enum HookEvent { Started, Opened, NewDocument, BeforeSave, Saved, PdfExported, Closing }

public static class EventHooks {
    /// How long BeforeSave and Closing wait for their macro.
    public const int WaitSeconds = 15;

    public static bool Waits(HookEvent e) => e is HookEvent.BeforeSave or HookEvent.Closing;

    /// (Measure runs) hooks of their own, never the user's: kept here, not written. Off: the ones switched off.
    public static Dictionary<string, string>? Override { get; set; }
    static HashSet<string> overrideOff = new(StringComparer.OrdinalIgnoreCase);

    public static bool TryParse(string name, out HookEvent e) =>
        Enum.TryParse(name.Trim(), true, out e) && Enum.IsDefined(e) && !int.TryParse(name, out _);

    /// Every hook: event name = macro.
    public static Dictionary<string, string> All() {
        if (Override != null) return new(Override, StringComparer.OrdinalIgnoreCase);
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (key, value) in Settings.Pairs("Hooks"))
            if (TryParse(key, out var e) && value != "") map[e.ToString()] = value;
        return map;
    }

    /// The events whose hooks are switched off.
    public static HashSet<string> Off() {
        if (Override != null) return new(overrideOff, StringComparer.OrdinalIgnoreCase);
        return Settings.Get("Hooks", "Off").Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
            .Where(n => TryParse(n, out _)).ToHashSet(StringComparer.OrdinalIgnoreCase);
    }

    public static string MacroFor(HookEvent e) => All().GetValueOrDefault(e.ToString()) ?? "";
    public static bool IsOn(HookEvent e) => !Off().Contains(e.ToString());

    /// Writes every hook at once (the dialog's OK).
    public static void Save(Dictionary<string, string> hooks, IEnumerable<string> off) {
        var kept = Enum.GetValues<HookEvent>().Select(e => e.ToString())
            .Where(n => hooks.GetValueOrDefault(n) is { Length: > 0 }).Select(n => (n, hooks[n])).ToList();
        var offKept = off.Where(n => kept.Any(k => k.n.Equals(n, StringComparison.OrdinalIgnoreCase))).ToList();
        if (Override != null) {
            Override = kept.ToDictionary(k => k.n, k => k.Item2, StringComparer.OrdinalIgnoreCase);
            overrideOff = offKept.ToHashSet(StringComparer.OrdinalIgnoreCase);
            return;
        }
        var pairs = kept.Select(k => (k.n, k.Item2)).ToList();
        if (offKept.Count > 0) pairs.Add(("Off", string.Join(",", offKept)));
        Settings.SetPairs("Hooks", pairs);
    }

    /// One event's macro ("": none), as a script sets it (Doc.SetHook).
    public static void Set(HookEvent e, string macro) {
        var hooks = All();
        if (macro == "") hooks.Remove(e.ToString()); else hooks[e.ToString()] = macro;
        Save(hooks, Off());
    }

    /// Switches an event's hook on or off, its macro kept (Doc.SetHookOn).
    public static void SetOn(HookEvent e, bool on) {
        var off = Off();
        if (on) off.Remove(e.ToString()); else off.Add(e.ToString());
        Save(All(), off);
    }

    /// A hook's macro file: a full path as it is; a name, from the user's macros folder, else the program's.
    public static string PathOf(string macro, string programFolder) {
        if (Path.IsPathRooted(macro)) return macro;
        var mine = Path.Combine(Macros.UserFolder, macro);
        return File.Exists(mine) ? mine : Path.Combine(programFolder, "macros", macro);
    }

    /// What the event means, in the user's language.
    public static string Describe(HookEvent e) => Lang.T("hooks.event" + e);

    /// A new hook macro's starting script: what it is told, and an example to start from.
    public static string Starter(HookEvent e) {
        var sb = new StringBuilder();
        sb.AppendLine("#Requires AutoHotkey v2.0");
        sb.AppendLine("#Include <Doc>");
        sb.AppendLine($"; A hook for dAHKxodus's {e} event (Macros > Event hooks): dAHKxodus runs it by itself each time.");
        sb.AppendLine("; Doc is the dAHKxodus window it happened in (Doc.Help() lists what it can do).");
        sb.AppendLine($"Event := A_Args.Length >= 1 ? A_Args[1] : \"\"     ; \"{e}\"");
        sb.AppendLine("Detail := A_Args.Length >= 2 ? A_Args[2] : \"\"    ; " + e switch {
            HookEvent.PdfExported => "the PDF's file",
            HookEvent.NewDocument => "(empty)",
            _ => "the document's file (empty: not saved yet)",
        });
        sb.AppendLine();
        sb.AppendLine(e switch {
            HookEvent.Saved => "; for example, a copy of the saved file in a Backups folder beside it, with the date and time in its name:\n; if Detail != \"\" {\n;     SplitPath(Detail, &Name, &Dir, &Ext, &Stem)\n;     DirCreate(Dir \"\\Backups\")\n;     FileCopy(Detail, Dir \"\\Backups\\\" Stem \" \" FormatTime(, \"yyyy-MM-dd HHmm\") \".\" Ext, true)\n; }",
            HookEvent.BeforeSave => "; for example, the date of the save at the end of the document (dAHKxodus waits for this macro, then saves):\n; Doc.MoveCaret(\"documentEnd\")\n; Doc.InsertText(\"`nSaved \" FormatTime(, \"LongDate\"))",
            HookEvent.Opened => "; for example, each document opened in page view, the page as wide as the window:\n; Doc.PageView(true)\n; Doc.PageWidth()",
            HookEvent.NewDocument => "; for example, every new document with today's date at the top:\n; Doc.InsertText(FormatTime(, \"LongDate\") \"`n\")",
            HookEvent.Closing => "; for example, a document with changes saved before the window closes (dAHKxodus waits for this macro):\n; if Detail != \"\" && Doc.Modified\n;     Doc.Save()",
            _ => "; what to do:\n; TrayTip(Event \" \" Detail, \"dAHKxodus\")",
        });
        return sb.ToString().Replace("\r\n", "\n");
    }
}

/// The Macros tab's Event hooks...: a macro for each event (none: nothing runs), and whether it is on.
public sealed class EventHooksDialog : Form {
    readonly Dictionary<HookEvent, ComboBox> boxes = new();
    readonly Dictionary<HookEvent, CheckBox> ons = new();
    readonly string programFolder;
    readonly Action<HookEvent, string> tryHook;

    /// The hooks chosen (event name = macro).
    public Dictionary<string, string> Result => boxes.Where(b => b.Value.SelectedIndex > 0)
        .ToDictionary(b => b.Key.ToString(), b => (string)b.Value.SelectedItem!, StringComparer.OrdinalIgnoreCase);
    /// The hooks switched off (their macros kept): event names.
    public List<string> OffResult => boxes.Where(b => b.Value.SelectedIndex > 0 && !ons[b.Key].Checked).Select(b => b.Key.ToString()).ToList();

    /// tryHook: runs a macro as its event would (Run now).
    public EventHooksDialog(string programFolder, Action<HookEvent, string> tryHook) {
        this.programFolder = programFolder;
        this.tryHook = tryHook;
        Text = Lang.T("hooks.title");
        MinimizeBox = MaximizeBox = false;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        AutoScaleMode = AutoScaleMode.Dpi;
        Font = SystemFonts.MessageBoxFont ?? Font;
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        Padding = new Padding(12);

        var hooks = EventHooks.All();
        var off = EventHooks.Off();
        var layout = new TableLayoutPanel { AutoSize = true, ColumnCount = 6, Dock = DockStyle.Fill };
        var intro = new Label { AutoSize = true, MaximumSize = new Size(760, 0), Margin = new Padding(3, 0, 3, 10), Text = Lang.T("hooks.intro") };
        layout.Controls.Add(intro);
        layout.SetColumnSpan(intro, 6);
        var tip = new ToolTip();
        foreach (var e in Enum.GetValues<HookEvent>()) {
            var name = new Label { Text = e.ToString(), AutoSize = true, Anchor = AnchorStyles.Left, Font = new Font(Font, FontStyle.Bold), Margin = new Padding(3, 8, 8, 0) };
            var box = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 260, AccessibleName = Lang.T("hooks.macroName", e), Margin = new Padding(3, 4, 3, 0) };
            boxes[e] = box;
            var on = new CheckBox { Text = Lang.T("hooks.on"), AutoSize = true, Anchor = AnchorStyles.Left, Checked = !off.Contains(e.ToString()), AccessibleName = Lang.T("hooks.onName", e), Margin = new Padding(8, 6, 3, 0) };
            ons[e] = on;
            Fill(box, hooks.GetValueOrDefault(e.ToString()));
            var make = Small(Lang.T("hooks.new"), () => NewMacro(e), Lang.T("hooks.newName", e));
            var edit = Small(Lang.T("hooks.edit"), () => { if (box.SelectedIndex > 0) EditMacro((string)box.SelectedItem!); }, Lang.T("hooks.editName", e));
            var run = Small(Lang.T("hooks.run"), () => { if (box.SelectedIndex > 0) this.tryHook(e, (string)box.SelectedItem!); }, Lang.T("hooks.runName", e));
            void Enable() => edit.Enabled = run.Enabled = on.Enabled = box.SelectedIndex > 0;
            box.SelectedIndexChanged += (s, a) => Enable();
            Enable();
            tip.SetToolTip(name, EventHooks.Describe(e));
            tip.SetToolTip(on, Lang.T("hooks.onTip"));
            layout.Controls.AddRange([name, box, on, make, edit, run]);
            var what = new Label { Text = EventHooks.Describe(e), AutoSize = true, ForeColor = SystemColors.GrayText, Margin = new Padding(3, 0, 3, 2), MaximumSize = new Size(600, 0) };
            layout.Controls.Add(new Label { AutoSize = true });
            layout.Controls.Add(what);
            layout.SetColumnSpan(what, 5);
        }
        var ok = new Button { Text = Lang.T("settings.ok"), DialogResult = DialogResult.OK, AutoSize = true, MinimumSize = new Size(88, 0) };
        var cancel = new Button { Text = Lang.T("settings.cancel"), DialogResult = DialogResult.Cancel, AutoSize = true, MinimumSize = new Size(88, 0) };
        var folder = new Button { Text = Lang.T("hooks.folder"), AutoSize = true };
        folder.Click += (s, a) => {
            try {
                Directory.CreateDirectory(Macros.UserFolder);
                Macros.UpdateLib(Macros.UserFolder, programFolder);
                Process.Start(new ProcessStartInfo("explorer.exe", $"\"{Macros.UserFolder}\"") { UseShellExecute = true });
            } catch (Exception ex) { MainForm.Log("macros folder: " + ex.Message); }
        };
        var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Anchor = AnchorStyles.Right, Margin = new Padding(0, 12, 0, 0) };
        buttons.Controls.AddRange([cancel, ok, folder]);
        layout.Controls.Add(buttons);
        layout.SetColumnSpan(buttons, 6);
        Controls.Add(layout);
        AcceptButton = ok;
        CancelButton = cancel;
    }

    Button Small(string text, Action action, string name) {
        var button = new Button { Text = text, AutoSize = true, AccessibleName = name, Margin = new Padding(3, 3, 0, 0) };
        button.Click += (s, e) => action();
        return button;
    }

    // The macros there are (the program's and the user's, by file name), and a macro chosen elsewhere; "(none)" first.
    void Fill(ComboBox box, string? chosen) {
        var macros = Macros.List(programFolder).Select(m => Path.GetFileName(m.path))
            .Distinct(StringComparer.OrdinalIgnoreCase).Order(StringComparer.OrdinalIgnoreCase).ToList();
        if (chosen != null && !macros.Contains(chosen, StringComparer.OrdinalIgnoreCase)) macros.Insert(0, chosen);
        box.Items.Clear();
        box.Items.Add(Lang.T("hooks.none"));
        foreach (var m in macros) box.Items.Add(m);
        var at = chosen == null ? 0 : box.Items.Cast<string>().ToList().FindIndex(i => string.Equals(i, chosen, StringComparison.OrdinalIgnoreCase));
        box.SelectedIndex = Math.Max(0, at);
    }

    // New...: a macro for the event from a starting script ("On Saved.ahk", or the next free name) in the user's macros
    // folder, chosen, and opened for editing.
    void NewMacro(HookEvent e) {
        try {
            var name = $"On {e}.ahk";
            for (var n = 2; File.Exists(Path.Combine(Macros.UserFolder, name)); n++) name = $"On {e} {n}.ahk";
            Macros.Save(Path.Combine(Macros.UserFolder, name), EventHooks.Starter(e), programFolder);
            foreach (var (other, box) in boxes) {
                var kept = box.SelectedIndex > 0 ? (string)box.SelectedItem! : null;
                Fill(box, other == e ? name : kept);
            }
            EditMacro(name);
        } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) {
            MessageBox.Show(this, Lang.T("macro.saveFailed", ex.Message), Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    // Edit: the macro in the editor Windows has for scripts (AutoHotkey's Edit), or Notepad.
    void EditMacro(string macro) {
        var path = EventHooks.PathOf(macro, programFolder);
        try { Process.Start(new ProcessStartInfo(path) { UseShellExecute = true, Verb = "edit" }); }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or InvalidOperationException) {
            Process.Start(new ProcessStartInfo("notepad.exe") { ArgumentList = { path }, UseShellExecute = false });
        }
    }
}
