#r "lib\webview2\Microsoft.Web.WebView2.Core.dll"
#r "lib\webview2\Microsoft.Web.WebView2.WinForms.dll"
// (from the NuGet package Microsoft.Web.WebView2 1.0.4314.35: lib_manual\netcoreapp3.0, which a plain package
// restore doesn't reference, and runtimes\win-x64\native\WebView2Loader.dll, found beside Core.dll)
#r "lib\docxodus\DocxodusNativeBridge.dll"
#r "lib\docxodus\Docxodus.dll"
#r "lib\docxodus\DocumentFormat.OpenXml.dll"
#r "lib\docxodus\DocumentFormat.OpenXml.Framework.dll"
#r "lib\docxodus\System.IO.Packaging.dll"
// (the Docxodus engine, run natively: build-engine.ps1 builds it from the Docxodus fork and copies it here)
#load "host\Settings.cs"
#load "host\Lang.cs"
#load "host\SettingsDialog.cs"
#load "host\Templates.cs"
#load "host\Merge.cs"
#load "host\MergeData.cs"
#load "host\MergeDialog.cs"
#load "host\Recovery.cs"
#load "host\Formats.cs"
#load "host\MacroHotkeys.cs"
#load "host\MacroDialog.cs"
#load "host\Hooks.cs"
#load "host\RibbonButtons.cs"
#load "host\Snippets.cs"
#load "host\PrintPages.cs"
#load "dax_extensions.cs"
// dAHKxodus: a WebView2 window showing Docxodus's ribbon editor (web\index.html), hosted by dAHKxodus.ahk.
// args: [0] the AHK script's window (bridge calls go there as WM_COPYDATA), [1] the program folder,
//       [2] a .docx to open (or ""), [3] measure mode: where to save the edited copy (or "").
// The window runs on its own thread so that this script returns and AutoHotkey stays free to answer bridge calls;
// when the window closes, the AHK script is told to exit (WM_APP + 1).
// Built as dAHKxodus.exe (vs\dAHKxodus.csproj), the roles swap: the exe is the program, and starts dAHKxodus.ahk
// /companion beside it, which tells the window its own (WM_COPYDATA "DAX0"); args[0] is then "0" until it does.
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

args = AppFolders.CompleteArgs?.Invoke(args) ?? args;
var ahk = (IntPtr)long.Parse(args[0]);
var dir = args[1];
var start = args.Length > 2 ? args[2] : "";
var measureOut = args.Length > 3 ? args[3] : "";
AppFolders.ProgramFolder = dir;
MainForm.LogPath = Path.Combine(dir, "logs", "dahkxodus.log");
MainForm.Log("---- start " + (start == "" ? "(no document)" : Path.GetFileName(start)) + (measureOut != "" ? " [measure]" : "")
    + (AppFolders.Portable ? " [portable]" : ""));
Settings.Init();
Lang.Load();

// dAHKxodus is running already (from this same folder): that one opens the file, or a new document, in a window of its
// own, and this start ends here. (Not for a measure run: it runs on a desktop of its own, by itself.)
if (measureOut == "" && Environment.GetEnvironmentVariable("DAHKXODUS_NEW_PROCESS") != "1" && MainForm.HandOff(dir, start)) {
    MainForm.Log("handed to the dAHKxodus already running: " + (start == "" ? "(a new document)" : Path.GetFileName(start)));
    if (ahk != IntPtr.Zero) MainForm.PostMessageW(ahk, 0x8001, IntPtr.Zero, IntPtr.Zero);    // (dAHKxodus.ahk: done)
    return 0;
}

var ui = new Thread(() => {
    try {
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        Application.ThreadException += (s, e) => MainForm.Log("UI exception: " + e.Exception);
        try { Application.SetHighDpiMode(HighDpiMode.PerMonitorV2); } catch { }
        Application.EnableVisualStyles();
        // Every window runs on this thread; the program ends with the last of them (MainForm.OnFormClosed).
        new MainForm(ahk, dir, start, measureOut).Show();
        Application.Run(new ApplicationContext());
    } catch (Exception ex) {
        MainForm.Log("fatal: " + ex);
    }
    MainForm.PostMessageW(MainForm.Ahk, 0x8001, IntPtr.Zero, IntPtr.Zero);
});
ui.SetApartmentState(ApartmentState.STA);
ui.Start();
if (AppFolders.ProgramExe != null) ui.Join();    // (the exe ends with its window; under Csharp-in-AHK this returns)
return 0;

class MainForm : Form {
    /// dAHKxodus's version (four parts, the fourth always 0 for the Microsoft Store): the AutoHotkey side, the installer and the About box read it here.
    public const string Version = "0.0.7.0";
    public static string LogPath = "";
    static readonly DateTime ProcessStart = Process.GetCurrentProcess().StartTime;

    public static void Log(string text) {
        try {
            Directory.CreateDirectory(Path.GetDirectoryName(LogPath)!);
            var since = (DateTime.Now - ProcessStart).TotalMilliseconds;
            File.AppendAllText(LogPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} +{since,7:0} ms  {text}\r\n");
        } catch { }
    }

    const int WM_COPYDATA = 0x004A, SMTO_ABORTIFHUNG = 2;
    // With dAHKxodus.ahk: "DAX0" (the companion's hello), "DAX1" (a call to a Bridge function: "id\tfn\targ"),
    // "DAX2", "DAX3" (its answer, its error: "id\ttext"), "DAX7" (run a macro: "id\tpath", or for an event hook
    // "id\tpath\tevent\tdetail"; answered as DAX1 is, with the macro's process id).
    // From another dAHKxodus starting: "DAX8" (HandOff).
    // With any script (Doc.ahk): "DAX4" (a call, JSON {id, name, args}), "DAX5" (its answer, JSON {id, result} or
    // {id, error, code}), "DAX6" (an event, JSON {event, ...}).
    const long Hello = 0x44415830, CallTag = 0x44415831, ReplyOk = 0x44415832, ReplyError = 0x44415833,
        DocCall = 0x44415834, DocReply = 0x44415835, DocEvent = 0x44415836, RunMacroTag = 0x44415837,
        // From a second start of dAHKxodus: "DAX8", "folder\tfile" (file "": a new document), to open in a window here.
        HandOffTag = 0x44415838;
    static string Untitled => Lang.T("file.untitled");

    [StructLayout(LayoutKind.Sequential)]
    struct COPYDATASTRUCT { public IntPtr dwData; public int cbData; public IntPtr lpData; }

    [DllImport("user32.dll")] public static extern bool PostMessageW(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] static extern bool IsWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    static extern IntPtr SendMessageTimeoutW(IntPtr hWnd, int msg, IntPtr wParam, ref COPYDATASTRUCT lParam,
        int flags, int timeout, out IntPtr result);

    /// The AHK script's window: bridge calls go there, and it is told when this window has closed. (The exe: 0 until
    /// its companion says hello.)
    public static IntPtr Ahk { get; private set; }
    // The open windows, oldest first, all on the one UI thread. The first is the one the program started with: only it
    // offers recovery copies and the last document. Later ones are new windows (New, Open, a file handed over).
    static readonly List<MainForm> Windows = new();
    static Task<CoreWebView2Environment>? browser;    // (one browser for every window: one profile, one process group)
    static bool measureProcess;                        // (a measure run's windows leave the user's settings alone)
    readonly bool firstWindow;
    readonly string dir, measureOut;
    readonly WebView2 web = new() { Dock = DockStyle.Fill };
    readonly EngineHost Engine = new();
    readonly Stopwatch clock = Stopwatch.StartNew();

    // The document: its file ("" until it is saved), and whether it has changed since (the page's ribbon says).
    string current;
    bool dirty;
    bool closeConfirmed;    // the user has answered the ask-to-save question; the next Close goes through
    bool asking;            // a question or a file dialog is up: another request waits its turn
    // Templates (host\Templates.cs): a template this window starts a new document from once its page is ready; a start
    // file that is a template to edit (File > Open, Recent) rather than to start a new document from; the next Save As
    // being Save as Template; a new document from a macro-enabled template (untitled.docm).
    string startTemplate = "";
    bool editTemplate, templateSave, newMacro;
    // Saves the host asked the page for (the page answers "saved" with the id): Save or Save As, and who waits.
    readonly Dictionary<int, (bool saveAs, string? target, TaskCompletionSource<bool> done)> pendingSaves = new();
    int nextSave;
    // The page's view as it last reported it ("view"): settings.ini's [View] when the window closes.
    Dictionary<string, string>? view;

    // Autosave (host\Recovery.cs): this window's recovery copy, written every AutosaveMinutes and when the window loses
    // focus while the document has changes not saved.
    readonly Recovery recovery;
    readonly System.Windows.Forms.Timer autosaveTimer = new();
    readonly Dictionary<int, TaskCompletionSource<byte[]?>> pendingSnapshots = new();
    int nextSnapshot;
    bool pageReady, autosaving, autosaveFailed;
    DateTime lastAutosave = DateTime.MinValue;
    // A document recovered from a copy left behind: changed (the page counts it unchanged once open) until saved or
    // replaced; the copy it came from goes once this window has a copy of its own.
    bool recovered;
    Recovery.Orphan? recovering;
    CloseReason closeReason = CloseReason.None;

    // Changes not saved: the page's, or a recovered document's.
    bool Changed => dirty || recovered;

    public MainForm(IntPtr ahk, string dir, string start, string measureOut, MainForm? from = null) {
        Ahk = ahk; this.dir = dir; current = start; this.measureOut = measureOut;
        firstWindow = Windows.Count == 0;
        if (measureOut != "") measureProcess = true;
        Windows.Add(this);
        // (a measure run's copies go to a folder of its own, never among the user's)
        recovery = new Recovery(measureProcess ? Path.Combine(dir, "logs", "measure-recovery") : Recovery.DefaultFolder);
        // (a measure run's macro hotkeys are its own: one for the page's check, the user's untouched)
        if (measureOut != "") MacroHotkeys.Override = new(StringComparer.OrdinalIgnoreCase) { ["Insert date.ahk"] = "Ctrl+Alt+J" };
        // (and its event hooks: none, until its check sets some)
        if (measureOut != "") EventHooks.Override = new(StringComparer.OrdinalIgnoreCase);
        // (and its ribbon button: one for the page's check)
        if (measureOut != "") RibbonButtons.Override = [new("Insert date.ahk", "Date")];
        // (and its snippets: a folder of its own, emptied, and one hotkey for the page's check)
        if (measureOut != "") {
            Snippets.OverrideFolder = Path.Combine(dir, "logs", "measure-snippets");
            if (firstWindow) try { if (Directory.Exists(Snippets.OverrideFolder)) Directory.Delete(Snippets.OverrideFolder, true); } catch (IOException) { }
            Snippets.OverrideHotkeys = new(StringComparer.OrdinalIgnoreCase);
        }
        autosaveTimer.Tick += (s, e) => _ = AutosaveAsync();
        // Leaving the window (another program, a dialog of its own excepted) is a moment to keep the copy current; not
        // more than twice a minute.
        Deactivate += (s, e) => { if (!asking && (DateTime.Now - lastAutosave).TotalSeconds >= 30) _ = AutosaveAsync(); };
        Size = new System.Drawing.Size(1400, 950);
        StartPosition = FormStartPosition.CenterScreen;
        Controls.Add(web);
        if (!measureProcess) WindowPlacement.Restore(this);    // (a measure run keeps the defaults: its runs compare)
        // A new window: a step down and to the right of the one it came from, not on top of it.
        if (from != null && from.WindowState == FormWindowState.Normal) {
            StartPosition = FormStartPosition.Manual;
            var step = SystemInformation.CaptionHeight + LogicalToDeviceUnits(8);
            Bounds = Cascaded(from.Bounds, Screen.FromControl(from).WorkingArea, step);
        } else if (from != null && from.WindowState == FormWindowState.Maximized) WindowState = FormWindowState.Maximized;
        UpdateTitle();
        if (measureOut != "") {    // a measure run that goes wrong still closes
            var watchdog = new System.Windows.Forms.Timer { Interval = 120000 };
            watchdog.Tick += (s, e) => { Log("measure: gave up after 120 s"); closeConfirmed = true; Close(); };
            watchdog.Start();
        }
        // (the exe: one companion for every window, started with the first)
        if (AppFolders.StartCompanion != null && firstWindow) HandleCreated += (s, e) => AppFolders.StartCompanion(Handle);
        Load += async (s, e) => {
            try { await Init(); } catch (Exception ex) { Log("init failed: " + ex); if (measureOut != "") Close(); }
        };
        // (back in the window: Headline Case's word lists again, if the file was edited meanwhile)
        // (and Find's options, which another window may have changed)
        Activated += (s, e) => { if (pageReady) { SendHeadlineLists(); SendFindOptions(); } };
    }

    async Task Init() {
        var data = Path.Combine(AppFolders.Local, "WebView2");
        browser ??= CoreWebView2Environment.CreateAsync(null, data);
        var env = await browser;
        await web.EnsureCoreWebView2Async(env);
        Log($"dAHKxodus {Version}, webview2 ready {clock.ElapsedMilliseconds} ms (runtime {env.BrowserVersionString})");
        DaxExtensions.Started(this);
        var cw = web.CoreWebView2;
        // The browser's own keys (Ctrl+S saves the page, Ctrl+P prints it, F5 reloads it, Ctrl+N opens a browser window,
        // ...) are off: the page gets those keys and acts as a word processor does. Editing keys are not affected.
        cw.Settings.AreBrowserAcceleratorKeysEnabled = false;
        cw.Settings.IsStatusBarEnabled = false;
        // (.example, not .local: a .local host name cost a fixed 2 s before the page's first script ran)
        cw.SetVirtualHostNameToFolderMapping("app.example", Path.Combine(dir, "web"), CoreWebView2HostResourceAccessKind.Allow);
        cw.WebMessageReceived += OnMessage;
        // A file dropped on the window that the page did not take (it takes .docx and .docm) must not replace the page.
        cw.NavigationStarting += (s, e) => {
            if (!e.Uri.StartsWith("https://app.example/", StringComparison.OrdinalIgnoreCase)) { e.Cancel = true; Log("navigation refused: " + e.Uri); }
        };
        cw.NewWindowRequested += (s, e) => { e.Handled = true; };
        // The ribbon's Paste (and Doc.Paste) reads the clipboard with the page's Clipboard API, which asks first.
        cw.PermissionRequested += (s, e) => {
            if (e.PermissionKind == CoreWebView2PermissionKind.ClipboardRead) e.State = CoreWebView2PermissionState.Allow;
        };
        // The right-click menu keeps the editing items (cut, copy, paste, spelling, emoji) and loses the browser's:
        // Refresh and Back would reload the editor and lose the document's changes without asking; Save as, Print,
        // Inspect and the rest are the page's, not the document's.
        cw.ContextMenuRequested += (s, e) => {
            var items = e.MenuItems;
            for (var i = items.Count - 1; i >= 0; i--)
                if (BrowserMenuItems.Contains(items[i].Name)) items.RemoveAt(i);
            // (no separator first, last, or twice running)
            for (var i = items.Count - 1; i >= 0; i--)
                if (items[i].Kind == CoreWebView2ContextMenuItemKind.Separator
                    && (i == 0 || i == items.Count - 1 || items[i - 1].Kind == CoreWebView2ContextMenuItemKind.Separator))
                    items.RemoveAt(i);
            if (items.Count == 0) e.Handled = true;    // (nothing left: no empty menu)
        };
        // The editing session runs here, in .NET, not in the page's WASM runtime: the page calls the engine through
        // this object (chrome.webview.hostObjects.sync.engine). DAHKXODUS_ENGINE=wasm keeps it in the page, to compare.
        cw.AddHostObjectToScript("engine", Engine);
        var query = new List<string>();
        var window = Environment.GetEnvironmentVariable("DAHKXODUS_WINDOW");    // (paragraphs per mount window, for tests)
        if (!string.IsNullOrEmpty(window)) query.Add("window=" + Uri.EscapeDataString(window));
        var engine = Environment.GetEnvironmentVariable("DAHKXODUS_ENGINE");
        if (!string.IsNullOrEmpty(engine)) query.Add("engine=" + Uri.EscapeDataString(engine));
        if (Environment.GetEnvironmentVariable("DAHKXODUS_VERIFY") == "1") query.Add("verify=1");    // (both engines, compared)
        var (user, initials, from) = UserIdentity.Read();
        Log($"user \"{user}\"{(initials == "" ? "" : $" ({initials})")} from {from}; language {Lang.Code}");
        query.Add("user=" + Uri.EscapeDataString(user));
        if (initials != "") query.Add("initials=" + Uri.EscapeDataString(initials));
        // The view the last window closed with: page or continuous, zoom, comments, the rail, Show/Hide ¶, the ruler, the
        // Navigation Pane, the Styles pane with its filter and Show Preview (not for a measure run).
        if (measureOut == "")
            foreach (var key in new[] { "Paginated", "Zoom", "Comments", "Rail", "Marks", "Ruler", "Nav", "Styles", "StylesFilter", "StylesPreview" }) {
                var value = Settings.Get("View", key);
                if (value != "") query.Add(key.ToLowerInvariant() + "=" + Uri.EscapeDataString(value));
            }
        cw.Navigate("https://app.example/index.html" + (query.Count == 0 ? "" : "?" + string.Join("&", query)));
    }

    // The browser's own right-click items (WebView2's names), taken off the menu.
    static readonly HashSet<string> BrowserMenuItems = new(StringComparer.OrdinalIgnoreCase) {
        "back", "forward", "reload", "saveAs", "print", "createQrCode", "share", "webCapture", "inspectElement",
        "openLinkInNewWindow", "saveLinkAs", "copyLinkLocation", "copyLinkToHighlight", "saveImageAs", "copyImageLocation",
        "openImageInNewWindow", "saveMediaAs", "loop", "showAllControls", "pictureInPicture", "readAloud", "translate",
        "sendTabToSelf",
    };

    void Send(object message) => web.CoreWebView2.PostWebMessageAsJson(JsonSerializer.Serialize(message));

    string DocName => current == "" ? (newMacro ? Path.ChangeExtension(Untitled, ".docm") : Untitled) : Path.GetFileName(current);

    // "*report.docx - dAHKxodus": the star while the document has changes not saved.
    void UpdateTitle() => Text = (Changed ? "*" : "") + DocName + " - dAHKxodus";

    // A line on the ribbon's status bar.
    void SendStatus(string text) => Send(new { type = "status", text });

    // ---- Autosave and recovery

    /// Writes the recovery copy every AutosaveMinutes ([General], 10 unless set; 0: never) while there are changes.
    void StartAutosave() {
        autosaveTimer.Stop();
        if (measureOut != "") return;    // (a measure run writes its copy when it asks: measureAutosave)
        var minutes = Math.Clamp(Settings.GetInt("General", "AutosaveMinutes", 10), 0, 120);
        if (minutes == 0) { ClearRecovery(); return; }
        autosaveTimer.Interval = minutes * 60_000;
        autosaveTimer.Start();
    }

    // The recovery copy: written while there are changes not saved (unless they are the ones written last), removed
    // once there are none. Returns whether a copy is there now.
    async Task<bool> AutosaveAsync() {
        if (autosaving || asking || !pageReady || web.CoreWebView2 == null) return recovery.HasCopy;
        autosaving = true;
        try {
            if (!recovered && !await IsDirtyAsync()) { ClearRecovery(); return false; }
            var bytes = await SnapshotAsync();
            if (bytes == null) return recovery.HasCopy;
            if (recovery.Write(bytes, current, DocName)) Log($"autosave: recovery copy written ({bytes.Length:N0} bytes)");
            lastAutosave = DateTime.Now;
            autosaveFailed = false;
            return true;
        } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) {
            Log("autosave failed: " + ex.Message);
            if (!autosaveFailed) SendStatus("⚠ " + Lang.T("recovery.notWritten", ex.Message));    // (once, not every time)
            autosaveFailed = true;
            return recovery.HasCopy;
        } finally { autosaving = false; }
    }

    // The document's bytes as Save would write them, without saving: the page answers "snapshot" with the id (null:
    // it couldn't, or took more than 20 s).
    async Task<byte[]?> SnapshotAsync() {
        var id = ++nextSnapshot;
        var done = new TaskCompletionSource<byte[]?>();
        pendingSnapshots[id] = done;
        Send(new { type = "snapshot", id });
        var first = await Task.WhenAny(done.Task, Task.Delay(20_000));
        pendingSnapshots.Remove(id);
        return first == done.Task ? done.Task.Result : null;
    }

    void ClearRecovery() {
        try { recovery.Clear(); } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
    }

    // At the start: a copy that a dAHKxodus which didn't close normally left behind is offered, the newest first (one
    // a start; the others at the next ones). Started with a file, it waits for a start without one.
    // Returns whether the copy is being recovered (its open is under way).
    bool OfferRecovery() {
        List<Recovery.Orphan> orphans;
        try { orphans = Recovery.Orphans(recovery.Folder); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { return false; }
        if (orphans.Count == 0) return false;
        var newest = orphans[0];
        var when = Lang.T("recovery.when", newest.Info.Saved.ToShortTimeString(), newest.Info.Saved.ToShortDateString());
        Log($"recovery: {orphans.Count} copy(ies) left behind, the newest \"{newest.Info.Name}\" of {newest.Info.Saved:yyyy-MM-dd HH:mm}");
        if (current != "") {
            SendStatus(Lang.T("recovery.later", newest.Info.Name, when));
            return false;
        }
        // (DAHKXODUS_RECOVER=yes or no answers for a test run: its window is on a desktop of its own)
        var auto = Environment.GetEnvironmentVariable("DAHKXODUS_RECOVER");
        DialogResult answer;
        if (auto is "yes" or "no") answer = auto == "yes" ? DialogResult.Yes : DialogResult.No;
        else {
            var more = orphans.Count > 1 ? "\n\n" + Lang.T("recovery.more", orphans.Count - 1) : "";
            asking = true;
            try {
                answer = MessageBox.Show(this, Lang.T("recovery.offer", newest.Info.Name, when) + more, "dAHKxodus",
                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            } finally { asking = false; }
        }
        if (answer == DialogResult.No) { Recovery.Discard(newest); Log("recovery: discarded"); }
        if (answer != DialogResult.Yes) return false;

        byte[] bytes;
        try { bytes = Recovery.Read(newest); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) {
            Log("recovery: unreadable: " + ex.Message);
            MessageBox.Show(this, Lang.T("recovery.unreadable", ex.Message), "dAHKxodus", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }
        // As changes not saved to the document it was: Save replaces that file, Save As keeps both.
        current = newest.Info.FileName;
        recovered = true;
        recovering = newest;
        UpdateTitle();
        Log($"recovery: recovering \"{newest.Info.Name}\" ({bytes.Length:N0} bytes)");
        Send(new { type = "open", name = newest.Info.Name, b64 = Convert.ToBase64String(bytes) });
        return true;
    }

    // The recovered document is open: this window writes a copy of its own, then the one left behind goes.
    async void RecoveredOpen() {
        var orphan = recovering!;
        recovering = null;
        if (await AutosaveAsync()) Recovery.Discard(orphan);
        var file = orphan.Info.FileName != "" ? Lang.T("recovery.hasFile", Path.GetFileName(orphan.Info.FileName)) : Lang.T("recovery.noFile");
        SendStatus(Lang.T("recovery.recovered", orphan.Info.Name, orphan.Info.Saved.ToShortTimeString()) + " " + file);
        Log($"recovery: recovered \"{orphan.Info.Name}\", title \"{Text}\", copy left behind {(File.Exists(orphan.DocxPath) ? "STILL THERE" : "removed")}");
    }

    // The fonts installed on this computer, for the ribbon's Font box and the Font and style dialogs (a page cannot
    // list them itself), read once, by the first window; and the recently used ones ([Fonts] Recent, newest first,
    // kept by "recentFonts"), listed first. A measure run starts with none. The caption labels the user made
    // ([Captions] Labels, kept by "captionLabels") ride along.
    static string[]? installedFonts;
    void SendFonts() {
        installedFonts ??= new System.Drawing.Text.InstalledFontCollection().Families.Select(f => f.Name).ToArray();
        var recent = measureOut != "" ? Array.Empty<string>()
            : Settings.Get("Fonts", "Recent").Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var captionLabels = measureOut != "" ? Array.Empty<string>()
            : Settings.Get("Captions", "Labels").Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        // (Mailings > Envelopes' return address: [Envelopes] ReturnAddress, its lines as a JSON string)
        var returnAddress = "";
        if (measureOut == "")
            try { returnAddress = JsonSerializer.Deserialize<string>(Settings.Get("Envelopes", "ReturnAddress", "\"\"")) ?? ""; } catch (JsonException) { }
        Send(new { type = "fonts", names = installedFonts, recent, captionLabels, returnAddress });
    }

    // The labels the page adds to the ribbon, in the user's language.
    void SendUi() => Send(new {
        type = "ui", units = Lang.T("page.units"), settings = Lang.T("page.settings"), settingsTip = Lang.T("page.settingsTip"),
        ahk = Lang.T("page.ahk"), ahkTip = Lang.T("page.ahkTip"), macros = Lang.T("page.macros"), macrosTip = Lang.T("page.macrosTip"),
        close = Lang.T("page.close"), closeTip = Lang.T("page.closeTip"), hint = Lang.T("page.hint"),
        print = Lang.T("page.print"), printTip = Lang.T("page.printTip"), pdf = Lang.T("page.pdf"), pdfTip = Lang.T("page.pdfTip"),
        copy = Lang.T("page.copy"), copyTip = Lang.T("page.copyTip"),
        newTemplate = Lang.T("page.newTemplate"), newTemplateTip = Lang.T("page.newTemplateTip"),
        merge = Lang.T("page.merge"), mergeTip = Lang.T("page.mergeTip"),
        saveTemplate = Lang.T("page.saveTemplate"), saveTemplateTip = Lang.T("page.saveTemplateTip"),
        palette = Lang.T("page.palette"), paletteTip = Lang.T("page.paletteTip"), palettePlaceholder = Lang.T("page.palettePlaceholder"),
        paletteOn = Lang.T("page.paletteOn"), paletteNotNow = Lang.T("page.paletteNotNow"), paletteFile = Lang.T("page.paletteFile"),
        paletteMacro = Lang.T("page.paletteMacro"), paletteTitleBar = Lang.T("page.paletteTitleBar"),
        paletteStatusBar = Lang.T("page.paletteStatusBar"), paletteNone = Lang.T("page.paletteNone"),
        snippets = Lang.T("page.snippets"), snippetsTip = Lang.T("page.snippetsTip"), snippetSave = Lang.T("page.snippetSave"),
        snippetSaveTip = Lang.T("page.snippetSaveTip"), snippetManage = Lang.T("page.snippetManage"),
        snippetManageTip = Lang.T("page.snippetManageTip"), snippetNone = Lang.T("page.snippetNone"), paletteSnippet = Lang.T("page.paletteSnippet"),
        record = Lang.T("page.record"), recordTip = Lang.T("page.recordTip"), recordStop = Lang.T("page.recordStop"),
        recording = Lang.T("page.recording"), nothingRecorded = Lang.T("page.nothingRecorded"),
        recordHeader = Lang.T("page.recordHeader"), recordMoved = Lang.T("page.recordMoved"), recordSkipped = Lang.T("page.recordSkipped"),
        macrosFolder = Lang.T("page.macrosFolder"), macroDeleteTip = Lang.T("page.macroDeleteTip"),
        hotkeys = Lang.T("page.hotkeys"), hotkeysTip = Lang.T("page.hotkeysTip"),
        hooks = Lang.T("page.hooks"), hooksTip = Lang.T("page.hooksTip"),
        myButtons = Lang.T("page.myButtons"), ribbonButtons = Lang.T("page.ribbonButtons"), ribbonButtonsTip = Lang.T("page.ribbonButtonsTip"),
        buttonMissing = Lang.T("page.buttonMissing"),
        ribbon = Lang.Ribbon(),
    });

    // ---- Settings

    void ShowSettings() {
        if (asking || measureOut != "") return;
        DialogResult answer;
        asking = true;
        try {
            using var dlg = new SettingsDialog();
            answer = dlg.ShowDialog(this);
        } finally { asking = false; }
        if (answer != DialogResult.OK) return;
        Lang.Load();
        Everywhere(w => w.SettingsChanged());
    }

    // What changed applies now, in every window: the host's language, the recent list, the name on new comments and
    // changes, and the autosave interval.
    void SettingsChanged() {
        UpdateTitle();
        SendUi();
        SendRecent();
        var (user, initials, from) = UserIdentity.Read();
        Log($"settings changed: user \"{user}\" from {from}, language {Lang.Code}, recent {RecentFiles.Max}, autosave {Settings.GetInt("General", "AutosaveMinutes", 10)} min");
        Send(new { type = "user", name = user, initials });
        StartAutosave();
    }

    protected override void OnFormClosed(FormClosedEventArgs e) {
        autosaveTimer.Stop();
        RaiseEvent("Closed");
        // Closed as the user chose (saved, or Don't Save): the copy goes. Closed by Windows shutting down, nobody was
        // asked: a document with changes keeps its copy, which the next start offers.
        if (closeReason != CloseReason.WindowsShutDown || !Changed) ClearRecovery();
        else Log("closed by Windows with changes not saved: the recovery copy stays");
        Windows.Remove(this);
        if (!measureProcess) {
            try {
                WindowPlacement.Save(this);
                if (view != null)
                    Settings.Update(() => { foreach (var (k, v) in view) Settings.Set("View", k, v); });
            } catch (Exception ex) { Log("settings at close: " + ex.Message); }
        }
        base.OnFormClosed(e);
        if (Windows.Count == 0) Application.ExitThread();    // (the last window: the program ends)
    }

    void SendRecent() {
        if (web.CoreWebView2 == null) return;
        var entries = RecentFiles.List().Select(p => new { id = p, name = Path.GetFileName(p), title = p }).ToArray();
        Send(new { type = "recent", entries });
    }

    // ---- Opening

    // False when the file can't be read (the user is told, unless quiet); else the page is opening it. asNew: the file
    // is a template (or a document used as one) to start a new, untitled document from; the file itself stays as it is.
    // A template opened to edit comes to the page as a document too, and goes back to a template when saved (WriteDocument).
    // A new document from a template with fields to fill (host\Merge.cs) asks for their values first, unless quiet (the
    // dialog: "Without merging" keeps the fields); values given (a script's): filled with those, no dialog.
    bool OpenFile(string path, bool quiet = false, bool asNew = false, Dictionary<string, string>? values = null, bool askMerge = true) {
        byte[] bytes;
        try {
            bytes = File.ReadAllBytes(path);
            if (MergedFiles.Remove(path)) { askMerge = false; try { File.Delete(path); } catch (IOException) { } }
        } catch (Exception ex) {
            Log($"open {path} failed: {ex.Message}");
            if (measureOut == "" && !quiet) MessageBox.Show(this, Lang.T("file.openFailed", path, ex.Message), "dAHKxodus", MessageBoxButtons.OK, MessageBoxIcon.Error);
            if (!asNew && ex is FileNotFoundException or DirectoryNotFoundException) { RecentFiles.Remove(path); Everywhere(w => w.SendRecent()); }
            return false;
        }
        if (asNew || Templates.IsTemplate(path)) bytes = Templates.ToDocument(bytes);
        if (asNew && values != null) bytes = Merge.Fill(bytes, values);
        else if (asNew && askMerge && !quiet && measureOut == "" && Merge.Scan(bytes) is { Count: > 0 } fields) {
            // (the dialog runs its own loop: not inside the page's message)
            BeginInvoke(() => {
                var filled = AskMerge(bytes, fields, Path.GetFileNameWithoutExtension(path), fromTemplate: true, out var written);
                if (written) StartBlank();    // (files written: this window starts blank)
                else Deliver(filled ?? bytes);
            });
            return true;
        }
        Deliver(bytes);
        return true;

        void Deliver(byte[] bytes) {
            current = asNew ? "" : Path.GetFullPath(path);
            newMacro = asNew && Path.GetExtension(path).ToLowerInvariant() is ".dotm" or ".docm";
            if (asNew) newRequested = true;    // (the page's "opened" then runs the NewDocument hook, as for a blank one)
            dirty = recovered = false;
            ClearRecovery();    // (the document before it went as the user chose)
            UpdateTitle();
            if (measureOut == "" && !asNew) { RecentFiles.Add(current); Everywhere(w => w.SendRecent()); }
            Log($"open {Path.GetFileName(path)}{(asNew ? " as a new document" : "")} ({bytes.Length:N0} bytes) at {clock.ElapsedMilliseconds} ms");
            Send(new { type = "open", name = asNew ? DocName : Path.GetFileName(path), b64 = Convert.ToBase64String(bytes) });
        }
    }

    // ---- Mail merge (host\Merge.cs, host\MergeDialog.cs)

    // Merged documents on their way to a window (File > Mail Merge...): a file in the temp folder that the window
    // opens as a new document, without asking again, and then deletes.
    static readonly HashSet<string> MergedFiles = new(StringComparer.OrdinalIgnoreCase);

    // The dialog: the document to open (the values typed or the row shown; all rows in one document), or null when
    // the user chose not to merge or the merge went to files (written: true; PDFs go on being written after).
    byte[]? AskMerge(byte[] bytes, List<MergeFieldInfo> fields, string name, bool fromTemplate, out bool written) {
        written = false;
        asking = true;
        try {
            using var dlg = new MergeDialog(name, fields, fromTemplate, Path.GetFileNameWithoutExtension(name));
            if (dlg.ShowDialog(this) != DialogResult.OK) { Log($"merge: {name} not merged"); return null; }
            switch (dlg.Output) {
                case MergeOutput.CurrentRow:
                    Log($"merge: {name}, {fields.Count} field(s) filled");
                    return Merge.Fill(bytes, dlg.Values);
                case MergeOutput.OneDocument: {
                    var clock = Stopwatch.StartNew();
                    // (a copy for each record: NEXT takes the next row into the same copy, SKIPIF leaves a record out)
                    var copies = Merge.MergeAll(bytes, dlg.Rows);
                    if (copies.Count == 0) { MessageBox.Show(this, Lang.T("mergeData.noRows"), "dAHKxodus", MessageBoxButtons.OK, MessageBoxIcon.Information); return null; }
                    var sources = copies.Select(c => new Docxodus.Source(new Docxodus.WmlDocument($"row{c.Number}.docx", c.Bytes), true)).ToList();
                    var all = Docxodus.DocumentBuilder.BuildDocument(sources).DocumentByteArray;
                    Log($"merge: {name}, {dlg.Rows.Count} rows, {copies.Count} record(s) in one document ({all.Length:N0} bytes, {clock.ElapsedMilliseconds} ms)");
                    return all;
                }
                default: {
                    var pdf = dlg.Output == MergeOutput.PdfPerRow;
                    using var pick = new FolderBrowserDialog {
                        Description = Lang.T("mergeData.folderTitle"), UseDescriptionForTitle = true,
                        InitialDirectory = Settings.Get("Merge", "OutputFolder", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)),
                    };
                    if (pick.ShowDialog(this) != DialogResult.OK) return null;
                    Settings.Set("Merge", "OutputFolder", pick.SelectedPath);
                    var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    var items = Merge.MergeAll(bytes, dlg.Rows).Select(c => (bytes: c.Bytes, path: MergedFileName(pick.SelectedPath, dlg.FilePattern, Path.GetFileNameWithoutExtension(name), c.First, c.Number, pdf ? ".pdf" : ".docx", names))).ToList();
                    written = true;
                    if (pdf) { _ = WriteMergedPdfsAsync(items, pick.SelectedPath); return null; }
                    foreach (var (merged, path) in items) File.WriteAllBytes(path, merged);
                    Log($"merge: {name}, {items.Count} file(s) written to {pick.SelectedPath}");
                    SendStatus(Lang.T("mergeData.written", items.Count, pick.SelectedPath));
                    OpenFolder(pick.SelectedPath);
                    return null;
                }
            }
        } catch (Exception ex) {
            Log("merge failed: " + ex);
            MessageBox.Show(this, Lang.T("merge.failed", ex.Message), "dAHKxodus", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return null;
        } finally { asking = false; }
    }

    // A merged file's path: the pattern's {Field} as the row's value (names compared as MergeData.Key does), {#} the
    // row's number; the template's name and the number when there is no pattern. Characters a file name can't have
    // become _, and a name used already (in this merge or in the folder) gets " (2)", " (3)"...
    static string MergedFileName(string folder, string pattern, string template, IReadOnlyDictionary<string, string> values, int number, string ext, HashSet<string> used) {
        if (pattern == "") pattern = template + " {#}";
        var name = System.Text.RegularExpressions.Regex.Replace(pattern, @"\{([^{}]+)\}", m => {
            var key = m.Groups[1].Value.Trim();
            if (key == "#") return number.ToString();
            return values.FirstOrDefault(kv => MergeData.Key(kv.Key) == MergeData.Key(key)).Value ?? "";
        });
        name = string.Concat(name.Select(ch => Path.GetInvalidFileNameChars().Contains(ch) ? '_' : ch)).Trim().TrimEnd('.');
        if (name == "") name = $"{template} {number}";
        if (name.EndsWith(ext, StringComparison.OrdinalIgnoreCase)) name = name[..^ext.Length];
        var path = Path.Combine(folder, name + ext);
        for (var n = 2; used.Contains(path) || File.Exists(path); n++) path = Path.Combine(folder, $"{name} ({n}){ext}");
        used.Add(path);
        return path;
    }

    static void OpenFolder(string folder) {
        try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("explorer.exe", "\"" + folder + "\"") { UseShellExecute = true }); }
        catch (Exception ex) { Log("merge folder: " + ex.Message); }
    }

    // A PDF per row: a window of its own opens each merged document and writes its pages (the page's PDF export),
    // then closes; this window's status bar counts.
    async Task WriteMergedPdfsAsync(List<(byte[] bytes, string path)> items, string folder) {
        var w = OpenWindow("", this);
        var done = 0;
        try {
            await w.firstDocument.Task;
            foreach (var (merged, path) in items) {
                var temp = Path.Combine(Path.GetTempPath(), $"dAHKxodus-merge-{Guid.NewGuid():N}.docx");
                File.WriteAllBytes(temp, merged);
                MergedFiles.Add(temp);
                if (!await w.WaitForDocument(() => w.OpenFile(temp, quiet: true, asNew: true))) { Log("merge pdf: could not open row " + (done + 1)); break; }
                if (!await w.ExportPdfAsync(path, quiet: true)) { Log("merge pdf: FAILED " + path); break; }
                done++;
                SendStatus(Lang.T("mergeData.pdfProgress", done, items.Count));
            }
        } catch (Exception ex) {
            Log("merge pdf failed: " + ex);
        } finally {
            if (!w.IsDisposed) { w.closeConfirmed = true; w.Close(); }
        }
        Log($"merge: {done} of {items.Count} PDF(s) written to {folder}");
        SendStatus(Lang.T("mergeData.written", done, folder));
        if (done < items.Count) MessageBox.Show(this, Lang.T("mergeData.pdfStopped", done, items.Count), "dAHKxodus", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        if (done > 0) OpenFolder(folder);
    }

    // File > Mail Merge...: this document's fields filled, as a new untitled document (this one stays as it is).
    async void MailMerge() {
        if (asking) return;
        var bytes = await SnapshotAsync();
        if (bytes == null) return;
        var fields = Merge.Scan(bytes);
        if (fields.Count == 0) {
            MessageBox.Show(this, Lang.T("merge.none"), "dAHKxodus", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        var filled = AskMerge(bytes, fields, DocName, fromTemplate: false, out _);
        if (filled == null) return;
        OpenMerged(filled);
    }

    // The key the merge dialog keeps this document's data source under ([Merge Sources]).
    string MergeKey => Path.GetFileNameWithoutExtension(DocName);

    // Mailings > Select Recipients > Use an Existing List... (and the Merge Fields pane's Browse...): a data source for
    // this document (the merge dialog starts from it), its columns listed by the Merge Fields pane. It replaces a field
    // list.
    void SelectRecipients() {
        if (asking) return;
        var remembered = Settings.Get("Merge Sources", MergeKey, "").Split('|', 2)[0];
        string path;
        List<string> tables;
        asking = true;
        try {
            using var dlg = new OpenFileDialog { Filter = Lang.T("mergeData.filter", MergeData.Extensions), InitialDirectory = remembered != "" ? Path.GetDirectoryName(remembered) : "" };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;
            path = dlg.FileName;
            try { tables = MergeData.Tables(path); }
            catch (Exception ex) { MessageBox.Show(this, Lang.T("mergeData.failed", path, ex.Message), "dAHKxodus", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
            var table = tables.Count == 1 ? tables[0] : MergeDialog.ChooseTable(this, Font, tables);
            if (table == null) return;
            MergeTable loaded;
            try { loaded = MergeData.Load(path, table); }
            catch (Exception ex) { MessageBox.Show(this, Lang.T("mergeData.failed", path, ex.Message), "dAHKxodus", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
            Settings.Set("Merge Sources", MergeKey, path + (table != "" ? "|" + table : ""));
            Settings.Set("Merge Skipped", MergeKey, "");
            Settings.Set("Merge Field Lists", MergeKey, "");
            Log($"merge: recipients {path} {table}, {loaded.Columns.Count} column(s), {loaded.Rows.Count} row(s)");
            Send(new { type = "mergeColumns", columns = loaded.Columns, source = SourceName(path, table), rows = PreviewRows(loaded) });
        } finally { asking = false; }
    }

    // The first rows of a data source, for Address Block's and Greeting Line's preview.
    static List<string[]> PreviewRows(MergeTable table) => table.Rows.Take(100).ToList();

    static string SourceName(string path, string table) => Path.GetFileName(path) + (table != "" ? " › " + table : "");

    // Mailings > Select Recipients > Use a Field List...: a list of field names (one a line, or a CSV of names) for the
    // Merge Fields pane, as MergeFielder reads one; no recipients. It replaces a data source.
    void SelectFieldList() {
        if (asking) return;
        asking = true;
        try {
            var remembered = Settings.Get("Merge Field Lists", MergeKey, "");
            using var dlg = new OpenFileDialog { Filter = Lang.T("mergeData.fieldListFilter"), InitialDirectory = remembered != "" ? Path.GetDirectoryName(remembered) : "" };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;
            List<string> names;
            try { names = MergeData.FieldList(dlg.FileName); }
            catch (Exception ex) { MessageBox.Show(this, Lang.T("mergeData.failed", dlg.FileName, ex.Message), "dAHKxodus", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
            Settings.Set("Merge Field Lists", MergeKey, dlg.FileName);
            Settings.Set("Merge Sources", MergeKey, "");
            Settings.Set("Merge Skipped", MergeKey, "");
            Log($"merge: field list {dlg.FileName}, {names.Count} name(s)");
            Send(new { type = "mergeColumns", columns = names, source = Lang.T("mergeData.fieldListName", Path.GetFileName(dlg.FileName)) });
        } finally { asking = false; }
    }

    // Mailings > Select Recipients > Edit Recipient List...: the data source's rows, ticked or not; the merge leaves out
    // the unticked ones.
    void EditRecipientList() {
        if (asking) return;
        var parts = Settings.Get("Merge Sources", MergeKey, "").Split('|', 2);
        if (parts[0] == "" || !File.Exists(parts[0])) return;
        asking = true;
        try {
            MergeTable data;
            try { data = MergeData.Load(parts[0], parts.Length > 1 ? parts[1] : ""); }
            catch (Exception ex) { MessageBox.Show(this, Lang.T("mergeData.failed", parts[0], ex.Message), "dAHKxodus", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
            var skipped = MergeDialog.LoadSkipped(MergeKey, data.Rows.Count);
            if (!MergeDialog.EditRows(this, Font, data, skipped, 0, out _)) return;
            MergeDialog.SaveSkipped(MergeKey, skipped, data.Rows.Count);
            Log($"merge: {data.Rows.Count - skipped.Count} of {data.Rows.Count} recipient(s) included");
        } finally { asking = false; }
    }

    // Mailings > Select Recipients > Clear Recipients (and the Merge Fields pane's Clear): this document has no data
    // source or field list any more.
    void ClearRecipients() {
        foreach (var section in new[] { "Merge Sources", "Merge Skipped", "Merge Field Lists" }) Settings.Set(section, MergeKey, "");
        Log("merge: recipients cleared");
        Send(new { type = "mergeColumns", columns = Array.Empty<string>(), source = "" });
    }

    // The columns of this document's remembered data source (or the names of its field list), for the Merge Fields
    // pane; none: an empty list.
    async void SendMergeColumns() {
        var remembered = Settings.Get("Merge Sources", MergeKey, "");
        var parts = remembered.Split('|', 2);
        var list = Settings.Get("Merge Field Lists", MergeKey, "");
        List<string> columns = new();
        List<string[]> rows = new();
        var source = "";
        try {
            if (remembered != "" && File.Exists(parts[0])) {
                var table = parts.Length > 1 ? parts[1] : "";
                var loaded = await Task.Run(() => MergeData.Load(parts[0], table));
                columns = loaded.Columns;
                rows = PreviewRows(loaded);
                source = SourceName(parts[0], table);
            } else if (list != "" && File.Exists(list)) {
                columns = await Task.Run(() => MergeData.FieldList(list));
                source = Lang.T("mergeData.fieldListName", Path.GetFileName(list));
            }
        } catch (Exception ex) { Log($"merge data {remembered}{list}: {ex.Message}"); }
        if (!web.IsDisposed) Send(new { type = "mergeColumns", columns, source, rows });
    }

    // A merged document in a new window (or this one, if it holds only an untouched new document).
    void OpenMerged(byte[] filled) {
        var temp = Path.Combine(Path.GetTempPath(), $"dAHKxodus-merge-{Guid.NewGuid():N}.docx");
        File.WriteAllBytes(temp, filled);
        MergedFiles.Add(temp);
        NewFromTemplate(temp);
    }

    // New (Ctrl+N, File > New): a new window, as in Word, unless this one holds only an untouched new document. It
    // starts from the template in Settings ("New documents start from"), if one is set and still there.
    void NewDocument() {
        if (Templates.Default is { } template) NewFromTemplate(template);
        else if (Untouched) StartBlank();
        else OpenWindow("", this);
    }

    // File > New from Template..., a template opened from Explorer or dropped, New with a template in Settings: a new
    // untitled document from it, here if this window holds only an untouched new document, else in a new window.
    void NewFromTemplate(string path) {
        if (Untouched) { if (!OpenFile(path, asNew: true)) StartBlank(); }
        else OpenWindow("", this, template: path);
    }

    void ChooseTemplate() {
        if (asking) return;
        string? chosen;
        asking = true;
        try {
            using var dlg = new TemplateDialog();
            if (dlg.ShowDialog(this) != DialogResult.OK) return;
            chosen = dlg.Chosen;
        } finally { asking = false; }
        if (chosen != null) NewFromTemplate(chosen);
    }

    void StartBlank() {
        newRequested = true;    // (the page's blankOpened then runs the NewDocument hook)
        current = "";
        newMacro = false;
        dirty = recovered = false;
        ClearRecovery();
        UpdateTitle();
        Send(new { type = "blank", name = Untitled });
    }

    void OpenWithDialog() {
        if (asking) return;
        string path;
        asking = true;
        try {
            using var dlg = new OpenFileDialog {
                Filter = Lang.T("file.filterOpen"),
                InitialDirectory = current != "" ? Path.GetDirectoryName(current) : "",
            };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;
            path = dlg.FileName;
        } finally { asking = false; }
        OpenPath(path, editTemplate: true);
    }

    // Open (the dialog, Recent, a file dropped or handed over): the window that has it already comes to the front;
    // else it opens here if this window holds only an untouched new document, and in a new window if not. A template
    // (.dotx, .dotm) starts a new document, as a double-click on one does in Word, unless it is opened to be edited
    // (editTemplate: File > Open, Recent).
    void OpenPath(string path, bool editTemplate = false) {
        var full = Path.GetFullPath(path);
        if (Templates.IsTemplate(full) && !editTemplate) { NewFromTemplate(full); return; }
        var open = Windows.FirstOrDefault(w => string.Equals(w.current, full, StringComparison.OrdinalIgnoreCase));
        if (open != null) { open.BringForward(); return; }
        if (Untouched) OpenFile(full);
        else OpenWindow(full, this, editTemplate: true);
    }

    // Only a new document nobody has typed in yet: New and Open use this window instead of opening another.
    bool Untouched => current == "" && !Changed && !asking && pageReady && recovering == null;

    // ---- Several windows

    /// A new window with the file (or a new document: "", from a template if one is given), cascaded from the one it
    /// came from. editTemplate: a template file is opened to be edited, not to start a new document from.
    static MainForm OpenWindow(string file, MainForm? from, string template = "", bool editTemplate = false) {
        var window = new MainForm(Ahk, from?.dir ?? AppFolders.ProgramFolder, file, "", from) { startTemplate = template, editTemplate = editTemplate };
        window.Show();
        Log($"new window ({Windows.Count} open): " + (file != "" ? Path.GetFileName(file) : template != "" ? $"(a new document from {Path.GetFileName(template)})" : "(a new document)"));
        return window;
    }

    void BringForward() {
        if (WindowState == FormWindowState.Minimized) WindowState = FormWindowState.Normal;
        Activate();
    }

    /// Calls f for each open window whose page is ready (what the user changed applies to all of them).
    static void Everywhere(Action<MainForm> f) {
        foreach (var w in Windows.ToList())
            if (w.pageReady && !w.IsDisposed) f(w);
    }

    /// A window's place a step down and to the right of bounds, back at the top left of the working area once it would
    /// run off it. (AhkWord's.)
    static System.Drawing.Rectangle Cascaded(System.Drawing.Rectangle bounds, System.Drawing.Rectangle area, int step) {
        var moved = new System.Drawing.Rectangle(bounds.X + step, bounds.Y + step, bounds.Width, bounds.Height);
        if (moved.Right > area.Right || moved.Bottom > area.Bottom) moved.Location = new System.Drawing.Point(area.Left + step, area.Top + step);
        return moved;
    }

    // A second start of dAHKxodus (a .docx double-clicked, the Start menu) hands its file to the one running from the
    // same folder, which opens it in a window here. True when a running dAHKxodus took it.
    delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowTextW(IntPtr hwnd, System.Text.StringBuilder text, int max);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
    [DllImport("user32.dll")] static extern bool AllowSetForegroundWindow(uint processId);

    public static bool HandOff(string folder, string file) {
        var candidates = new List<(IntPtr hwnd, uint pid)>();
        var self = (uint)Environment.ProcessId;
        EnumWindows((hwnd, _) => {
            if (!IsWindowVisible(hwnd)) return true;
            var title = new System.Text.StringBuilder(512);
            GetWindowTextW(hwnd, title, title.Capacity);
            if (!title.ToString().EndsWith(" - dAHKxodus", StringComparison.Ordinal)) return true;
            GetWindowThreadProcessId(hwnd, out var pid);
            if (pid != self) candidates.Add((hwnd, pid));
            return true;
        }, IntPtr.Zero);
        foreach (var (hwnd, pid) in candidates) {
            AllowSetForegroundWindow(pid);    // (so the window it opens can come to the front)
            if (SendCopyData(hwnd, HandOffTag, Path.GetFullPath(folder) + "\t" + file, 10000, IntPtr.Zero)) return true;
        }
        return false;
    }

    // A file handed over by a second start: here, if it came from the same folder (another copy of dAHKxodus, such as
    // a portable one, keeps its own).
    bool TakeHandOff(string text) {
        var parts = text.Split('\t', 2);
        if (parts.Length != 2 || !string.Equals(Path.GetFullPath(parts[0]).TrimEnd('\\'), Path.GetFullPath(dir).TrimEnd('\\'), StringComparison.OrdinalIgnoreCase))
            return false;
        var file = parts[1];
        Log("handed over: " + (file == "" ? "(a new document)" : Path.GetFileName(file)));
        BeginInvoke(() => {
            if (file == "") OpenWindow("", this).BringForward();
            else OpenPath(file);
        });
        return true;
    }

    // ---- Saving

    // Asks the page for the document's bytes (it answers "saved" with the id); true once they are written. A target
    // (a script's Doc.Save(path)) is written without asking where. The BeforeSave hook runs first, and is waited for
    // (not for a script's save: the script that saves is the one to change the document first, and the hook's own Doc
    // calls would wait for that script's call to end).
    async Task<bool> SaveAsync(bool saveAs, string? target = null, bool hooks = true) {
        if (hooks) await RunHookAsync(HookEvent.BeforeSave, current);
        var id = ++nextSave;
        var done = new TaskCompletionSource<bool>();
        pendingSaves[id] = (saveAs, target, done);
        Send(new { type = "save", id });
        return await done.Task;
    }

    // The page's bytes for Save (Ctrl+S, the ribbon's Save), or for a save the host asked for (id).
    void OnSaved(byte[] bytes, int id) {
        var saveAs = false;
        string? target = null;
        TaskCompletionSource<bool>? done = null;
        if (id != 0 && pendingSaves.Remove(id, out var pending)) (saveAs, target, done) = pending;
        var ok = WriteDocument(bytes, saveAs, target);
        done?.TrySetResult(ok);
    }

    bool WriteDocument(byte[] bytes, bool saveAs, string? target = null) {
        var path = current;
        // A measure run never writes over the document it opened, nor asks where: its saves go beside its copy.
        if (target != null) path = target;
        else if (measureOut != "") path = Path.ChangeExtension(measureOut, ".host.docx");
        else if (saveAs || path == "") {
            // Save as Template: the template types only, in the templates folder, the macros kept as they are.
            var asTemplate = templateSave;
            templateSave = false;
            var macros = DocName.EndsWith(".docm", StringComparison.OrdinalIgnoreCase) || DocName.EndsWith(".dotm", StringComparison.OrdinalIgnoreCase);
            asking = true;
            try {
                using var dlg = new SaveFileDialog {
                    Filter = Lang.T(asTemplate ? "file.filterTemplate" : "file.filterSave"), OverwritePrompt = true, AddExtension = true,
                    FileName = asTemplate ? Path.GetFileNameWithoutExtension(DocName) : DocName,
                    FilterIndex = asTemplate ? (macros ? 2 : 1)
                        : Path.GetExtension(DocName).ToLowerInvariant() switch { ".docm" => 2, ".dotx" => 3, ".dotm" => 4, _ => 1 },
                    InitialDirectory = asTemplate ? Templates.Folder : current != "" ? Path.GetDirectoryName(current) : "",
                };
                if (asTemplate) Directory.CreateDirectory(Templates.Folder);
                if (dlg.ShowDialog(this) != DialogResult.OK) return false;
                path = dlg.FileName;
            } finally { asking = false; }
        }
        // (a template under .dotx/.dotm, a document under any other name: the page's bytes are always a document's)
        bytes = Templates.ForPath(bytes, path);
        try {
            WriteAtomically(path, bytes);
        } catch (Exception ex) {
            Log($"save {path} failed: {ex.Message}");
            MessageBox.Show(this, Lang.T("file.saveFailed", path, ex.Message), "dAHKxodus", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
        current = Path.GetFullPath(path);
        Log($"saved {Path.GetFileName(current)} ({bytes.Length:N0} bytes)");
        if (measureOut == "") { RecentFiles.Add(current); Everywhere(w => w.SendRecent()); }
        // The page marks the document saved (its ribbon's markSaved) and answers "dirty" false; the title loses its
        // star now, not a message later.
        Send(new { type = "savedOk", name = Path.GetFileName(current) });
        dirty = recovered = false;
        ClearRecovery();    // (saved: nothing left to recover)
        UpdateTitle();
        RaiseEvent("Saved", new JsonObject { ["path"] = current });
        ReportModified();
        _ = RunHookAsync(HookEvent.Saved, current);
        return true;
    }

    // (measure runs) The host's saved copy as a template and back: each package's type, and the text kept.
    static string MeasureTemplates(string docx) {
        try {
            var bytes = File.ReadAllBytes(docx);
            string TypeOf(byte[] b) {
                using var ms = new MemoryStream(b);
                using var d = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(ms, false);
                return d.DocumentType + ":" + (d.MainDocumentPart?.Document?.Body?.InnerText.Length ?? -1);
            }
            var asTemplate = Templates.ForPath(bytes, "x.dotx");
            var back = Templates.ToDocument(asTemplate);
            var again = Templates.ForPath(bytes, "x.docx");
            string a = TypeOf(bytes), t = TypeOf(asTemplate), d = TypeOf(back);
            var ok = a.StartsWith("Document:") && t.StartsWith("Template:") && d.StartsWith("Document:")
                && a.Split(':')[1] == t.Split(':')[1] && t.Split(':')[1] == d.Split(':')[1] && ReferenceEquals(again, bytes);
            return $"{(ok ? "ok" : "FAILED")} ({a} -> {t} -> {d})";
        } catch (Exception ex) { return "FAILED: " + ex.Message; }
    }

    // The new file is written beside the old one and then takes its place, so a failure part-way (a full disk, a
    // lost network drive) leaves the old file as it was. File.Replace keeps the old file's attributes and permissions.
    static void WriteAtomically(string path, byte[] bytes) {
        var full = Path.GetFullPath(path);
        var tmp = Path.Combine(Path.GetDirectoryName(full)!, $"~dax{Guid.NewGuid():N}.tmp");
        File.WriteAllBytes(tmp, bytes);
        try {
            if (File.Exists(full)) File.Replace(tmp, full, null);
            else File.Move(tmp, full);
        } catch {
            try { File.Delete(tmp); } catch { }
            throw;
        }
    }

    // ---- Asking to save

    // The page's own answer (it counts typing not yet committed); the last "dirty" message if the page can't answer.
    async Task<bool> IsDirtyAsync() {
        try {
            var r = await web.CoreWebView2.ExecuteScriptAsync("window.dax ? window.dax.isDirty() : false");
            return r == "true";
        } catch (Exception ex) {
            Log("isDirty: " + ex.Message);
            return dirty;
        }
    }

    // Before the document is closed or replaced: true when it may go (unchanged, saved now, or the user said Don't Save).
    async Task<bool> CanDiscard(bool hooks = true) {
        if (measureOut != "" || web.CoreWebView2 == null) return true;
        if (asking) return false;
        if (!recovered && !await IsDirtyAsync()) return true;
        asking = true;
        DialogResult answer;
        try {
            answer = MessageBox.Show(this, Lang.T("file.askSave", DocName), "dAHKxodus",
                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
        } finally { asking = false; }
        if (answer == DialogResult.Cancel) return false;
        if (answer == DialogResult.No) return true;
        return await SaveAsync(saveAs: false, hooks: hooks);
    }

    protected override async void OnFormClosing(FormClosingEventArgs e) {
        closeReason = e.CloseReason;
        if (closeConfirmed || measureOut != "" || web.CoreWebView2 == null || e.CloseReason == CloseReason.WindowsShutDown) {
            base.OnFormClosing(e);
            return;
        }
        e.Cancel = true;    // (decided below, after asking the page; then closed again)
        if (closingHook) return;    // (the Closing hook is running: this close waits for it)
        // The Closing hook first, waited for: it may save the document (then nobody is asked) or change it.
        closingHook = true;
        try { await RunHookAsync(HookEvent.Closing, current); } finally { closingHook = false; }
        if (!await CanDiscard()) return;
        closeConfirmed = true;
        BeginInvoke(Close);
    }

    // ---- Messages from the page

    void OnMessage(object? sender, CoreWebView2WebMessageReceivedEventArgs e) {
        try {
            using var doc = JsonDocument.Parse(e.WebMessageAsJson);
            var m = doc.RootElement;
            string Str(string name) => m.TryGetProperty(name, out var v) ? v.ToString() : "";
            switch (Str("type")) {
                case "ready":
                    Log($"editor ready {clock.ElapsedMilliseconds} ms ({Str("engine")} engine; page: {Str("sinceNavigation")} ms since navigation, scripts {Str("scripts")} ms, boot {Str("boot")} ms)");
                    SendUi();
                    SendFonts();
                    if (measureOut == "") SendRecent();
                    SendMacros();    // (their hotkeys work from the start)
                    SendSnippets();
                    SendHooks();
                    SendListDefaults();
                    SendHeadlineLists();
                    SendFindOptions();
                    pageReady = true;
                    StartAutosave();
                    // A copy left behind by a dAHKxodus that didn't close normally, offered before anything opens.
                    if (firstWindow && OfferRecovery()) break;
                    // Started without a document: the last one, if the user wants that (and it is still there).
                    if (firstWindow && current == "" && measureOut == "" && Settings.GetBool("General", "OpenLastAtStart", false))
                        current = RecentFiles.List().FirstOrDefault(File.Exists) ?? "";
                    // A template handed over (the command line, a second start) starts a new document from it.
                    if (current != "" && Templates.IsTemplate(current) && !editTemplate) { startTemplate = current; current = ""; }
                    if (current != "") OpenFile(current);
                    else if (startTemplate != "" && OpenFile(startTemplate, asNew: true)) { }
                    // (a start without a document: the template in Settings, if one is set; no NewDocument hook, as
                    // for the blank one)
                    else if (measureOut == "" && Templates.Default is { } template && OpenFile(template, quiet: true, asNew: true)) newRequested = false;
                    else Send(new { type = "blank", name = Untitled });
                    break;
                case "opened":
                    if (newRequested) { newRequested = false; _ = RunHookAsync(HookEvent.NewDocument, ""); }
                    Log($"opened {clock.ElapsedMilliseconds} ms (window {Environment.GetEnvironmentVariable("DAHKXODUS_WINDOW") ?? "default"}; open {Str("open")} ms: first window after {Str("firstWindow")} ms, longest frozen {Str("longestGap")} ms; first paint {Str("paint")} ms, {Str("blocks")} blocks)");
                    if (recovering != null) RecoveredOpen();
                    if (measureOut != "" && !measureStarted) { measureStarted = true; Send(new { type = "measure" }); }
                    DocumentReady(true);
                    SendMergeColumns();
                    break;
                case "mailings":    // Mailings > Start Mail Merge..., Select Recipients..., Finish & Merge...
                    switch (Str("command")) {
                        case "recipients": BeginInvoke(SelectRecipients); break;
                        case "fieldlist": BeginInvoke(SelectFieldList); break;
                        case "editrecipients": BeginInvoke(EditRecipientList); break;
                        case "clearrecipients": ClearRecipients(); break;
                        default: BeginInvoke(MailMerge); break;
                    }
                    break;
                case "blankOpened":    // a new document (New, or a start without one) is in the editor
                    DocumentReady(true);
                    if (newRequested) { newRequested = false; _ = RunHookAsync(HookEvent.NewDocument, ""); }
                    break;
                case "openFailed":
                    Log("open failed in the page: " + Str("error"));
                    if (recovering != null) { recovering = null; recovered = false; }    // (its copy stays for next time)
                    var scriptOpen = documentWait != null;
                    documentWait?.TrySetResult(false);
                    documentWait = null;
                    if (!scriptOpen) MessageBox.Show(this, Lang.T("file.openFailed", DocName, Str("error")), "dAHKxodus", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    current = "";
                    UpdateTitle();
                    Send(new { type = "blank", name = Untitled });
                    break;
                case "measured":
                    var bytes = Convert.FromBase64String(Str("b64"));
                    File.WriteAllBytes(measureOut, bytes);
                    var ahkInfo = m.GetProperty("ahk");
                    Log($"measured: AHK returned \"{ahkInfo.GetProperty("text")}\" in {ahkInfo.GetProperty("roundTrip")} ms; save {Str("save")} ms, {bytes.Length:N0} bytes");
                    closeConfirmed = true;
                    BeginInvoke(Close);    // (not inside the page's message: scripts are told Closed)
                    break;
                case "saved":
                    OnSaved(Convert.FromBase64String(Str("b64")), m.TryGetProperty("id", out var sid) && sid.ValueKind == JsonValueKind.Number ? sid.GetInt32() : 0);
                    break;
                case "saveFailed":
                    var failedId = m.TryGetProperty("id", out var fid) && fid.ValueKind == JsonValueKind.Number ? fid.GetInt32() : 0;
                    Log("save failed in the page: " + Str("error"));
                    if (pendingSaves.Remove(failedId, out var failed)) failed.done.TrySetResult(false);
                    MessageBox.Show(this, Lang.T("file.saveFailed", DocName, Str("error")), "dAHKxodus", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                case "snapshot":    // the document's bytes for the recovery copy (SnapshotAsync)
                    var snapId = m.TryGetProperty("id", out var snap) && snap.ValueKind == JsonValueKind.Number ? snap.GetInt32() : 0;
                    if (Str("error") != "") Log("autosave: the page could not give the document: " + Str("error"));
                    if (pendingSnapshots.Remove(snapId, out var waiting))
                        waiting.TrySetResult(Str("b64") is { Length: > 0 } b64 ? Convert.FromBase64String(b64) : null);
                    break;
                case "measureAutosave":    // (measure runs) autosave now: is the copy there, as a .docx of the document?
                    _ = MeasureAutosave();
                    break;
                case "measureRecoveryState":    // (measure runs) after the save: the copy must be gone
                    var still = File.Exists(recovery.DocxPath);
                    Log($"measure: recovery copy after the save {(still ? "STILL THERE" : "removed")}");
                    Send(new { type = "measureReply", what = "recoveryState", ok = !still });
                    break;
                case "printRequest":    // File > Print, Ctrl+P
                    BeginInvoke(Print);
                    break;
                case "pdfRequest":    // File > Export to PDF
                    BeginInvoke(() => ExportPdf());
                    break;
                case "copyRequest":    // File > Save a Copy As (HTML, Markdown, plain text)
                    BeginInvoke(() => _ = SaveCopyAsync(null));
                    break;
                case "measurePdf":    // (measure runs) a PDF beside the measure copy, no dialog
                    if (measureOut != "") ExportPdf(Path.ChangeExtension(measureOut, ".pdf"));
                    break;
                case "printReady":
                    var printId = m.TryGetProperty("id", out var pid) && pid.ValueKind == JsonValueKind.Number ? pid.GetInt32() : 0;
                    if (pendingPrints.Remove(printId, out var printWaiter)) printWaiter.TrySetResult(m.Clone());
                    break;
                case "closeRequest":    // File > Close: as the window's own close (asks to save a changed document)
                    BeginInvoke(Close);
                    break;
                case "settingsRequest":
                    BeginInvoke(ShowSettings);    // (not inside the page's message: the dialog runs its own loop)
                    break;
                case "view":    // the page's view changed: kept for settings.ini's [View] at close
                    view = new() {
                        ["Paginated"] = Str("paginated") is "True" or "true" ? "1" : "0",
                        // (a number, 1 = 100%; "width": Page width; "page": One page, a whole page in view)
                        ["Zoom"] = m.TryGetProperty("zoom", out var z) && z.ValueKind == JsonValueKind.Number
                            ? z.GetDouble().ToString(System.Globalization.CultureInfo.InvariantCulture)
                            : Str("zoom") is "width" or "page" ? Str("zoom") : "",
                        ["Comments"] = Str("comments") is "True" or "true" ? "1" : "0",
                        ["Rail"] = Str("rail") is "True" or "true" ? "1" : "0",
                        ["Marks"] = Str("marks") is "True" or "true" ? "1" : "0",
                        ["Ruler"] = Str("ruler") is "True" or "true" ? "1" : "0",
                        ["Nav"] = Str("nav") is "True" or "true" ? "1" : "0",
                        ["Styles"] = Str("styles") is "True" or "true" ? "1" : "0",
                        ["StylesFilter"] = Str("stylesFilter") is "recommended" or "inuse" or "all" ? Str("stylesFilter") : "",
                        ["StylesPreview"] = Str("stylesPreview") is "True" or "true" ? "1" : "0",
                    };
                    if (view["Zoom"] == "") view.Remove("Zoom");
                    if (view["StylesFilter"] == "") view.Remove("StylesFilter");
                    break;
                case "dirty":
                    dirty = Str("dirty") == "True" || Str("dirty") == "true";
                    UpdateTitle();
                    if (measureOut != "") Log($"measure: page says {(dirty ? "changed" : "unchanged")}, title \"{Text}\"");
                    ReportModified();
                    break;
                case "newRequest":
                    NewDocument();
                    break;
                case "openRequest":
                    OpenWithDialog();
                    break;
                case "openRecent":
                    OpenPath(Str("path"));
                    break;
                case "saveAsRequest":
                    _ = SaveAsync(saveAs: true);
                    break;
                case "fileMenu":    // the host's own File menu commands that have no message of their own
                    switch (Str("id")) {
                        case "newTemplate": BeginInvoke(ChooseTemplate); break;    // (the dialog runs its own loop)
                        case "saveTemplate": templateSave = true; _ = SaveAsync(saveAs: true); break;
                        case "merge": BeginInvoke(MailMerge); break;
                        default: Log("file menu: unknown command " + Str("id")); break;
                    }
                    break;
                case "saveRequest":    // Save (Ctrl+S, the ribbon's) while a BeforeSave hook is set: the hook, then the save
                    _ = SaveAsync(saveAs: false);
                    break;
                case "ribbonButtonsRequest":    // the Macros tab's Ribbon buttons...
                    BeginInvoke(ShowRibbonButtons);
                    break;
                case "hooksRequest":    // the Macros tab's Event hooks...
                    BeginInvoke(ShowHooks);
                    break;
                case "headlineDefaults":    // the editor's own Headline Case lists, for a headline-case.txt not made yet
                    try {
                        if (!File.Exists(HeadlineListsPath)) {
                            Directory.CreateDirectory(Path.GetDirectoryName(HeadlineListsPath)!);
                            File.WriteAllText(HeadlineListsPath, Str("text").Replace("\r\n", "\n").Replace("\n", "\r\n"), new System.Text.UTF8Encoding(false));
                            Log("headline case: made " + HeadlineListsPath);
                        }
                    } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) {
                        Log("headline case: " + ex.Message);
                    }
                    break;
                case "recentFonts":    // a font was applied: [Fonts] Recent, for the next window and the next start
                    if (measureOut != "") { Log("recent fonts (measure run, not kept): " + Str("names")); break; }
                    if (m.TryGetProperty("names", out var names) && names.ValueKind == JsonValueKind.Array)
                        Settings.Set("Fonts", "Recent", string.Join("|", names.EnumerateArray().Select(n => n.GetString() ?? "")
                            .Where(n => n != "" && !n.Contains('|'))));
                    break;
                case "returnAddress":  // Mailings > Envelopes' return address, for the next envelope
                    if (measureOut != "") break;
                    Settings.Set("Envelopes", "ReturnAddress", JsonSerializer.Serialize(Str("text")));
                    break;
                case "captionLabels":  // Insert Caption's New Label / Delete Label: [Captions] Labels, for the next window and start
                    if (measureOut != "") { Log("caption labels (measure run, not kept): " + Str("labels")); break; }
                    if (m.TryGetProperty("labels", out var labels) && labels.ValueKind == JsonValueKind.Array)
                        Settings.Set("Captions", "Labels", string.Join("|", labels.EnumerateArray().Select(n => n.GetString() ?? "")
                            .Where(n => n != "" && !n.Contains('|'))));
                    break;
                case "findOptions":    // the find bar's Options: [Find] Options, for every window
                    if (measureOut != "") { Log("find options (measure run, not kept): " + Str("options")); break; }
                    sentFindOptions = Str("options");
                    Settings.Set("Find", "Options", sentFindOptions);
                    break;
                case "listDefaults":    // List levels > Also for new documents: [Lists] Defaults, for every window
                    if (measureOut != "") { Log("list defaults (measure run, not kept): " + Str("levels")); break; }
                    Settings.Set("Lists", "Defaults", Str("levels"));
                    break;
                case "measureResize":    // (measure runs) the window narrower or back, as a snap to one side does
                    if (measureOut != "") {
                        var width = m.TryGetProperty("width", out var w) && w.ValueKind == JsonValueKind.Number ? w.GetInt32() : 1400;
                        Size = new System.Drawing.Size(width, Height);
                        Log($"measure: window {Width} x {Height}");
                        Send(new { type = "measureReply", what = "resized", ok = true });
                    }
                    break;
                case "measureHooks":    // (measure runs) hooks of its own, then a save as Ctrl+S makes one
                    if (measureOut != "") _ = MeasureHooks();
                    break;
                case "measureSave":    // the measure run's check of the host's own save (SaveAsync)
                    _ = SaveAsync(saveAs: false).ContinueWith(t => {
                        Log($"measure: host save {(t.Result ? "written" : "FAILED")}, title \"{Text}\"");
                        Log("measure: templates " + MeasureTemplates(Path.ChangeExtension(measureOut, ".host.docx")));
                    }, TaskScheduler.FromCurrentSynchronizationContext());
                    break;
                case "drop":    // files dropped on the page (postMessageWithAdditionalObjects): open the first document
                    var dropped = e.AdditionalObjects?.OfType<CoreWebView2File>().Select(f => f.Path)
                        .FirstOrDefault(p => Path.GetExtension(p).ToLowerInvariant() is ".docx" or ".docm" or ".dotx" or ".dotm");
                    if (dropped != null) { Activate(); OpenPath(dropped); }
                    break;
                case "ahk":    // a Bridge function in dAHKxodus.ahk; its answer comes later (WndProc: ahkResult)
                    CallAhk(CallTag, m.GetProperty("id").GetInt32(), Str("fn") + "\t" + Str("arg"));
                    break;
                case "listMacros":    // the Macros tab's list is opening
                    try { Macros.UpdateLib(Macros.UserFolder, dir); } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
                    SendMacros();
                    break;
                case "hotkeysRequest":    // the Macros tab's Hotkeys...: with the keys dAHKxodus itself uses (taken)
                    var takenNow = MacroHotkeys.Taken(m.TryGetProperty("taken", out var tk) ? tk.Clone() : default);
                    BeginInvoke(() => ShowHotkeys(takenNow));
                    break;
                case "runMacro":    // a macro from that list: dAHKxodus.ahk starts it, told this window (DAHKXODUS_HWND)
                    var macroId = m.GetProperty("id").GetInt32();
                    if (Macros.List(dir).Any(x => string.Equals(x.path, Str("path"), StringComparison.OrdinalIgnoreCase))) {
                        Log("macro: " + Path.GetFileName(Str("path")) + (Str("hotkey") != "" ? $" (hotkey {Str("hotkey")})" : ""));
                        CallAhk(RunMacroTag, macroId, Str("path"));
                    } else Send(new { type = "ahkResult", id = macroId, result = "", error = Lang.T("macro.notFound", Str("path")) });
                    break;
                case "deleteMacro":    // the Macros list, right-click: one of the user's own macros, after asking
                    var doomed = Str("path");
                    BeginInvoke(() => DeleteMacro(doomed));
                    break;
                case "listSnippets":    // Insert > Snippets is opening: the folder may have changed
                    SendSnippets();
                    break;
                case "insertSnippet":    // a snippet from that list, its hotkey or the palette
                    var snippetFile = Str("file");
                    _ = InsertSnippetAsync(snippetFile, Str("hotkey"));
                    break;
                case "saveSnippetRequest":    // Save Selection as Snippet...: the selection's HTML (the editor's copy of it)
                    var snippetHtml = Str("html");
                    var snippetText = Str("text");
                    BeginInvoke(() => SaveSnippet(snippetHtml, snippetText));    // (the dialog runs its own loop)
                    break;
                case "manageSnippets":
                    BeginInvoke(ManageSnippets);
                    break;
                case "measureSnippets":    // (measure runs) the page's selection kept as a snippet, and inserted back
                    if (measureOut != "") _ = MeasureSnippets(Str("html"));
                    break;
                case "openMacrosFolder":
                    try {
                        Directory.CreateDirectory(Macros.UserFolder);
                        Macros.UpdateLib(Macros.UserFolder, dir);
                        Process.Start(new ProcessStartInfo("explorer.exe", $"\"{Macros.UserFolder}\"") { UseShellExecute = true });
                    } catch (Exception ex) { Log("macros folder: " + ex.Message); }
                    break;
                case "measureDocScript":    // (measure runs) tests\doc-check.ahk drives this window through Doc.ahk
                    if (measureOut != "") _ = MeasureDocScript();
                    break;
                case "recorded":    // a macro recording stopped: its script, to name and save
                    var recordedScript = Str("script");
                    var takenKeys = MacroHotkeys.Taken(m.TryGetProperty("taken", out var rt) ? rt.Clone() : default);
                    if (measureOut != "") MeasureRecorded(recordedScript);
                    else BeginInvoke(() => SaveRecordedMacro(recordedScript, takenKeys));    // (the dialog runs its own loop)
                    break;
                case "measureRunRecorded":    // (measure runs) the recorded macro, replayed
                    if (measureOut != "") CallAhk(RunMacroTag, 0, Path.Combine(dir, "logs", "measure-recorded.ahk"));
                    break;
                case "command":    // a ribbon command ran (onCommand): scripts' Command event
                    if (Subscribed("Command"))
                        RaiseEvent("Command", new JsonObject {
                            ["command"] = Str("name"), ["label"] = Str("label"),
                            ["args"] = m.TryGetProperty("args", out var ca) && ca.ValueKind == JsonValueKind.Object ? JsonNode.Parse(ca.GetRawText()) : new JsonObject(),
                        });
                    break;
                case "textChanged":    // typing paused, or a command changed the document: scripts' TextChanged event
                    RaiseEvent("TextChanged");
                    break;
                case "log":
                    Log("page: " + Str("text"));
                    break;
                case "capture":    // (measure runs) a picture of the editor as it shows: logs\<name>.png
                    if (measureOut != "") _ = CaptureAsync(Str("name"));
                    break;
                case "engineStats":    // time spent in the native engine itself (the page's own timing adds the crossing)
                    Log("host: " + Engine.TakeStats());
                    break;
            }
        } catch (Exception ex) {
            Log("message failed: " + ex);
        }
    }

    // ---- Printing and PDF: the page view's pages, printed by WebView2 (the page's print CSS shows the pages only)

    readonly Dictionary<int, TaskCompletionSource<JsonElement>> pendingPrints = new();
    int nextPrint;

    // The page switches to page view at 100% and says the first page's size (inches), the count and every page's size
    // (points, "612x792,792x612"); null when it can't.
    async Task<(double width, double height, int pages, string sizes)?> PreparePrintAsync() {
        var id = ++nextPrint;
        var done = new TaskCompletionSource<JsonElement>();
        pendingPrints[id] = done;
        Send(new { type = "preparePrint", id });
        var first = await Task.WhenAny(done.Task, Task.Delay(15_000));
        pendingPrints.Remove(id);
        if (first != done.Task) { Log("print: the page did not get ready"); return null; }
        var r = done.Task.Result;
        if (r.TryGetProperty("error", out var err) && err.GetString() is { Length: > 0 } message) { Log("print: " + message); return null; }
        return (r.GetProperty("widthIn").GetDouble(), r.GetProperty("heightIn").GetDouble(), r.GetProperty("pages").GetInt32(),
            r.TryGetProperty("sizes", out var sizes) ? sizes.GetString() ?? "" : "");
    }

    // File > Print (Ctrl+P): the browser's print dialog, with its preview, printer and pages; the page puts its view
    // back when the dialog closes (afterprint).
    async void Print() {
        if (asking || web.CoreWebView2 == null) return;
        var size = await PreparePrintAsync();
        if (size == null) { MessageBox.Show(this, Lang.T("print.failed", DocName), "dAHKxodus", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        Log($"print: {size.Value.pages} pages, {size.Value.width} x {size.Value.height} in");
        // (pages that are not the printer's default paper, an envelope: each on paper of its size, as Word sends them)
        var pages = PrintPages.Parse(size.Value.sizes);
        if (pages.Count > 0 && !PrintPages.FitDefaultPaper(pages)) { await PrintBySizeAsync(pages); return; }
        web.CoreWebView2.ShowPrintUI(CoreWebView2PrintDialogKind.Browser);
    }

    // Print pages of their own paper sizes: the printer, copies and pages asked here, then a job for each run of pages
    // of one size, with that paper size (the driver picks its form of that size: Envelope #10 for a #10 envelope), no
    // margins and none of the browser's headers and footers. The page's view is put back after.
    async Task PrintBySizeAsync(List<(double w, double h)> pages) {
        string printer;
        int copies, first, last;
        asking = true;
        try {
            using var dlg = new PrintPagesDialog(DocName, pages);
            if (dlg.ShowDialog(this) != DialogResult.OK) { Send(new { type = "printDone" }); return; }
            (printer, copies, first, last) = (dlg.Printer, dlg.Copies, dlg.First, dlg.Last);
        } finally { asking = false; }
        var failed = new List<string>();
        foreach (var run in PrintPages.Runs(pages, first, last)) {
            var settings = web.CoreWebView2.Environment.CreatePrintSettings();
            settings.PrinterName = printer;
            settings.Copies = copies;
            settings.PageRanges = $"{run.first}-{run.last}";
            settings.ShouldPrintHeaderAndFooter = false;
            settings.ShouldPrintBackgrounds = true;
            settings.MarginTop = settings.MarginBottom = settings.MarginLeft = settings.MarginRight = 0;
            settings.ScaleFactor = 1;
            double wIn = run.w / 72, hIn = run.h / 72;
            var landscape = wIn > hIn;
            settings.Orientation = landscape ? CoreWebView2PrintOrientation.Landscape : CoreWebView2PrintOrientation.Portrait;
            // The printer's own form of that size (COM#10 for a #10 envelope), made this user's paper on it for the job
            // and put back after, as choosing it by hand does: a driver takes its own forms, where many ignore a
            // custom size. With no such form, a custom size (the short side across, a wide page turned on it).
            var form = PrintPages.FormOf(printer, (run.w, run.h));
            IDisposable? paper = form is { } f ? PrintPages.UsePaper(printer, f.kind, landscape) : null;
            if (paper != null) settings.MediaSize = CoreWebView2PrintMediaSize.Default;
            else {
                settings.MediaSize = CoreWebView2PrintMediaSize.Custom;
                settings.PageWidth = Math.Min(wIn, hIn);
                settings.PageHeight = Math.Max(wIn, hIn);
            }
            CoreWebView2PrintStatus status;
            try { status = await web.CoreWebView2.PrintAsync(settings); }
            catch (Exception ex) { Log("print: " + ex.Message); status = CoreWebView2PrintStatus.OtherError; }
            finally { paper?.Dispose(); }
            Log($"print: pages {run.first}-{run.last}, {PrintPages.Name((run.w, run.h))}, on {printer} " +
                (paper != null ? $"as its form {form!.Value.name} ({form.Value.kind})" : form != null ? $"(its form {form.Value.name} could not be set: a custom size)" : "(no form of that size: a custom size)") + $": {status}");
            if (status != CoreWebView2PrintStatus.Succeeded) failed.Add($"{run.first}-{run.last} ({status})");
        }
        Send(new { type = "printDone" });
        if (failed.Count > 0)
            MessageBox.Show(this, Lang.T("printPages.failed", printer, string.Join(", ", failed)), "dAHKxodus", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        else SendStatus(Lang.T("printPages.sent", printer));
    }

    // File > Export to PDF: asks where (the document's name and folder to start), writes the PDF, puts the view back.
    async void ExportPdf(string? path = null) => await ExportPdfAsync(path);

    // ---- Save a Copy As

    // The document in another format (Formats: .html, .md, .txt), from the bytes Save would write; the window goes on
    // with the .docx, which stays as it is (changed or not). Quiet (a script's Doc.SaveCopy): no dialog, no message.
    async Task<bool> SaveCopyAsync(string? path, bool quiet = false) {
        if (asking || web.CoreWebView2 == null || !pageReady) return false;
        if (path == null) {
            asking = true;
            try {
                using var dlg = new SaveFileDialog {
                    Filter = Lang.T("file.filterCopy"), FileName = Path.GetFileNameWithoutExtension(DocName), OverwritePrompt = true, AddExtension = true,
                    FilterIndex = Math.Clamp(Settings.GetInt("General", "CopyFormat", 1), 1, Formats.Extensions.Length),
                    InitialDirectory = current != "" ? Path.GetDirectoryName(current) : "",
                };
                if (dlg.ShowDialog(this) != DialogResult.OK) return false;
                path = dlg.FileName;
                // (a name typed with another extension than the type chosen: the extension decides)
                if (!Formats.Supported(path)) path += Formats.Extensions[dlg.FilterIndex - 1];
                Settings.Set("General", "CopyFormat", dlg.FilterIndex.ToString());
            } finally { asking = false; }
        }
        var clock = Stopwatch.StartNew();
        try {
            var bytes = await SnapshotAsync() ?? throw new InvalidOperationException(Lang.T("file.copyNoDocument"));
            var title = Path.GetFileNameWithoutExtension(DocName);
            var target = path;
            var text = await Task.Run(() => Formats.Convert(bytes, target, title));
            await Task.Run(() => Formats.Write(target, text));
            Log($"copy: {Path.GetFileName(path)} written ({text.Length:N0} chars, {clock.ElapsedMilliseconds} ms)");
            if (!quiet) SendStatus(Lang.T("file.copySaved", path));
            return true;
        } catch (Exception ex) {
            Log($"copy {path} failed: {ex}");
            if (!quiet) MessageBox.Show(this, Lang.T("file.copyFailed", path, ex.Message), "dAHKxodus", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
    }

    // Quiet (a script's Doc.ExportPdf): no message, no measure check; the result says whether the PDF was written.
    async Task<bool> ExportPdfAsync(string? path, bool quiet = false) {
        if (asking || web.CoreWebView2 == null) return false;
        if (path == null) {
            asking = true;
            try {
                using var dlg = new SaveFileDialog {
                    Filter = Lang.T("file.filterPdf"), FileName = Path.ChangeExtension(DocName, ".pdf"), OverwritePrompt = true, AddExtension = true,
                    InitialDirectory = current != "" ? Path.GetDirectoryName(current) : "",
                };
                if (dlg.ShowDialog(this) != DialogResult.OK) return false;
                path = dlg.FileName;
            } finally { asking = false; }
        }
        var size = await PreparePrintAsync();
        var ok = false;
        try {
            if (size != null) {
                var clock = Stopwatch.StartNew();
                var how = "devtools";
                ok = await PrintPdfDevToolsAsync(path, size.Value.width, size.Value.height);
                if (!ok) {
                    // (WebView2's own: one paper size for every page; a landscape page is turned, other sizes are not)
                    how = "PrintToPdfAsync";
                    var settings = web.CoreWebView2.Environment.CreatePrintSettings();
                    settings.PageWidth = size.Value.width;
                    settings.PageHeight = size.Value.height;
                    settings.MarginTop = settings.MarginBottom = settings.MarginLeft = settings.MarginRight = 0;
                    settings.ShouldPrintBackgrounds = true;
                    settings.ShouldPrintHeaderAndFooter = false;
                    ok = await web.CoreWebView2.PrintToPdfAsync(path, settings);
                }
                Log($"pdf: {(ok ? "written" : "FAILED")} {Path.GetFileName(path)} by {how} ({(File.Exists(path) ? new FileInfo(path).Length : 0):N0} bytes, {size.Value.pages} pages, {size.Value.sizes} pt, {clock.ElapsedMilliseconds} ms)");
            }
        } catch (Exception ex) {
            Log("pdf failed: " + ex.Message);
        } finally {
            Send(new { type = "printDone" });
        }
        if (ok) _ = RunHookAsync(HookEvent.PdfExported, path);
        if (quiet) return ok;
        if (measureOut != "") { MeasurePdf(path, ok, size?.pages ?? 0, size?.sizes ?? ""); return ok; }
        if (ok) SendStatus(Lang.T("print.pdfSaved", path));
        else MessageBox.Show(this, Lang.T("print.pdfFailed", path), "dAHKxodus", MessageBoxButtons.OK, MessageBoxIcon.Error);
        return ok;
    }

    // The PDF through the DevTools protocol (Page.printToPDF), which gives each page the size of its @page box
    // (preferCSSPageSize): a document with landscape or other-sized sections keeps them. Paper and margins are only
    // the default, for anything without a size. False (and a log line) when it fails: the caller falls back.
    async Task<bool> PrintPdfDevToolsAsync(string path, double widthIn, double heightIn) {
        try {
            var args = JsonSerializer.Serialize(new {
                printBackground = true, preferCSSPageSize = true, displayHeaderFooter = false,
                paperWidth = widthIn, paperHeight = heightIn, marginTop = 0, marginBottom = 0, marginLeft = 0, marginRight = 0,
            });
            var json = await web.CoreWebView2.CallDevToolsProtocolMethodAsync("Page.printToPDF", args);
            using var result = JsonDocument.Parse(json);
            var bytes = Convert.FromBase64String(result.RootElement.GetProperty("data").GetString() ?? "");
            if (bytes.Length == 0) return false;
            var temp = path + ".tmp";
            File.WriteAllBytes(temp, bytes);
            File.Move(temp, path, overwrite: true);
            return true;
        } catch (Exception ex) {
            Log("pdf by devtools failed: " + ex.Message);
            return false;
        }
    }

    // (measure runs) The PDF is a PDF, with as many pages as the page view has, each the size of its page there
    // (points, rounded: the PDF's MediaBoxes in page order).
    void MeasurePdf(string path, bool ok, int pages, string sizes) {
        var text = File.Exists(path) ? System.Text.Encoding.Latin1.GetString(File.ReadAllBytes(path)) : "";
        var inPdf = System.Text.RegularExpressions.Regex.Matches(text, @"/Type\s*/Page[^s]").Count;
        var boxes = string.Join(",", System.Text.RegularExpressions.Regex.Matches(text, @"/MediaBox\s*\[\s*[-\d.]+\s+[-\d.]+\s+([-\d.]+)\s+([-\d.]+)\s*\]")
            .Select(b => $"{Math.Round(double.Parse(b.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture))}x{Math.Round(double.Parse(b.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture))}"));
        var good = ok && text.StartsWith("%PDF-") && inPdf == pages && pages > 0 && boxes == sizes;
        Log($"measure: pdf {(good ? "ok" : "FAILED")} ({inPdf} pages in the PDF, {pages} in the page view; sizes in the PDF {boxes}, in the page view {sizes})");
        Send(new { type = "measureReply", what = "pdf", ok = good });
    }

    async Task MeasureAutosave() {
        var wrote = await AutosaveAsync();
        var path = recovery.DocxPath;
        // (a .docx: a zip whose first entry Word wrote, "PK" first)
        var docx = File.Exists(path) && File.ReadAllBytes(path) is { Length: > 4 } b && b[0] == (byte)'P' && b[1] == (byte)'K';
        var info = Path.ChangeExtension(path, ".json");
        // (kept for tests\check-saved.ps1: the copy must be the document, its comment and picture included)
        if (docx) File.Copy(path, Path.ChangeExtension(measureOut, ".autosave.docx"), overwrite: true);
        Log($"measure: autosave {(wrote && docx && File.Exists(info) ? "copy written" : "FAILED")} ({(File.Exists(path) ? new FileInfo(path).Length : 0):N0} bytes, info {(File.Exists(info) ? "there" : "MISSING")})");
        Send(new { type = "measureReply", what = "autosave", ok = wrote && docx && File.Exists(info) });
    }

    async Task CaptureAsync(string name) {
        try {
            var path = Path.Combine(dir, "logs", Path.GetFileName(name) + ".png");
            using (var file = File.Create(path))
                await web.CoreWebView2.CapturePreviewAsync(CoreWebView2CapturePreviewImageFormat.Png, file);
            Log("captured " + path);
            Send(new { type = "captured", name });
        } catch (Exception ex) { Log("capture failed: " + ex.Message); Send(new { type = "captured", name }); }
    }

    // Asks dAHKxodus.ahk to call Bridge.<fn>(arg) (CallTag: "fn\targ") or to start a macro (RunMacroTag: its path).
    // AutoHotkey takes the call at once and runs it in a thread of its own, so the function may use Doc.ahk (which
    // talks to this window, now free); its answer comes back later as "DAX2"/"DAX3" with the id (WndProc), and goes
    // to the page as ahkResult.
    void CallAhk(long tag, int id, string text) {
        if (Ahk != IntPtr.Zero && SendCopyData(Ahk, tag, id + "\t" + text, 10000)) return;
        Send(new { type = "ahkResult", id, result = "", error = Lang.T(Ahk == IntPtr.Zero ? "bridge.notRunning" : "bridge.noAnswer") });
    }

    // WM_COPYDATA to another window (wParam: this one); true when it took it (answered 1).
    bool SendCopyData(IntPtr to, long tag, string text, int timeout) => SendCopyData(to, tag, text, timeout, Handle);

    static bool SendCopyData(IntPtr to, long tag, string text, int timeout, IntPtr from) {
        var ptr = Marshal.StringToHGlobalUni(text);
        try {
            var cds = new COPYDATASTRUCT { dwData = (IntPtr)tag, cbData = (text.Length + 1) * 2, lpData = ptr };
            return SendMessageTimeoutW(to, WM_COPYDATA, from, ref cds, SMTO_ABORTIFHUNG, timeout, out var answer) != IntPtr.Zero
                && answer == (IntPtr)1;
        } finally {
            Marshal.FreeHGlobal(ptr);
        }
    }

    // (Inside a message another process sent, a call to WebView2 fails: what these messages start runs afterwards,
    // through BeginInvoke.)
    protected override void WndProc(ref Message m) {
        if (m.Msg == WM_COPYDATA) {
            var cds = Marshal.PtrToStructure<COPYDATASTRUCT>(m.LParam);
            var tag = (long)cds.dwData;
            if (tag == HandOffTag) {
                m.Result = (IntPtr)(TakeHandOff(Marshal.PtrToStringUni(cds.lpData, cds.cbData / 2).TrimEnd('\0')) ? 1 : 0);
                return;
            }
            if (tag == Hello && AppFolders.ProgramExe != null) {
                Ahk = m.WParam;
                Log($"companion ready {clock.ElapsedMilliseconds} ms");
                m.Result = (IntPtr)1;
                return;
            }
            if (tag == ReplyOk || tag == ReplyError) {
                var text = Marshal.PtrToStringUni(cds.lpData, cds.cbData / 2).TrimEnd('\0');
                var tab = text.IndexOf('\t');
                var id = tab > 0 && int.TryParse(text.AsSpan(0, tab), out var n) ? n : 0;
                var result = tab >= 0 ? text[(tab + 1)..] : text;
                var ok = tag == ReplyOk;
                BeginInvoke(() => {
                    if (pendingAhk.Remove(id, out var waiter)) waiter.TrySetResult((ok, result));
                    else Send(new { type = "ahkResult", id, result = ok ? result : "", error = ok ? "" : result });
                });
                m.Result = (IntPtr)1;
                return;
            }
            if (tag == DocCall) {
                var caller = m.WParam;
                var text = Marshal.PtrToStringUni(cds.lpData, cds.cbData / 2).TrimEnd('\0');
                BeginInvoke(() => _ = RunScriptCallAsync(caller, text));
                m.Result = (IntPtr)1;
                return;
            }
        }
        base.WndProc(ref m);
    }

    // ---- Scripts: the Doc object (Doc.ahk). A call names an editor command (the page's ribbon.exec runs it) or one
    // of this window's own (HostCommands); calls run one at a time, in the order they came, from any script.

    sealed class ScriptError : Exception {
        public string Code { get; }
        public ScriptError(string code, string message) : base(message) { Code = code; }
    }

    readonly SemaphoreSlim scriptTurn = new(1, 1);
    // The scripts' windows that asked for events, and which.
    readonly Dictionary<IntPtr, HashSet<string>> subscribers = new();
    // A script's Open or New: done when the page has the document (true) or could not open it (false).
    TaskCompletionSource<bool>? documentWait;
    bool measureStarted;
    bool? reportedModified;

    // The program's own commands, as the page lists the editor's (name, group, description, arguments in order).
    static readonly (string name, string group, string description, string[] args, string? returns)[] HostCommands = [
        ("new", "file", "Start a new document (asks to save the one open first, if it has changes): blank, or from template (a .dotx, .dotm or .docx; a name alone is looked for in the templates folder). Settings' template is not used: a script's new document is blank unless it names one. The template's merge fields are filled from values (an object, field name: text, True/False for yes/no fields; or a saved record's name); without it they stay, and no dialog is shown.", ["template?", "values?"], "true, or false when the template could not be opened or the user cancelled"),
        ("merge", "file", "Mail merge: the document's fields filled, and the filled document becomes this window's, untitled. From values (an object, field name: text, True/False for yes/no fields; or a saved record's name); without it, the Mail Merge dialog asks.", ["values?"], "true, or false when the document has no fields or the user cancelled"),
        ("mergeFields", "query", "The document's merge fields: an array of {name, kind (merge, fillin, ask, placeholder, control), label, default, yesNo}.", [], "array"),
        ("open", "file", "Open a .docx (asks to save the one open first, if it has changes). A template (.dotx, .dotm) opens to be edited.", ["path"], "true, or false when it could not be opened or the user cancelled"),
        ("save", "file", "Save the document: to its file (asking where if it has none), or to path.", ["path?"], "true once written, false when not"),
        ("exportPdf", "file", "Write the document's pages as a PDF.", ["path"], "true once written"),
        ("saveCopy", "file", "Write a copy in another format, by path's extension: .html (a web page), .md (Markdown) or .txt (plain text). The document stays as it is.", ["path"], "true once written"),
        ("print", "file", "Show the print dialog.", [], null),
        ("close", "file", "Close this window (asks to save a changed document); dAHKxodus ends with its last window.", [], null),
        ("newWindow", "file", "Open a new window, with path or a new document. Set Doc.Window to the answer to send Doc calls there.", ["path?"], "the new window (its handle)"),
        ("setView", "view", "Page view (paginated true) or continuous (false), and the zoom: \"width\" (page width), \"page\" (one page) or a scale (1 = 100%). Either may be left out.", ["paginated?", "zoom?"], null),
        ("view", "query", "The view: {paginated, zoom} (zoom: a scale, \"width\" or \"page\").", [], "object"),
        ("setStatus", "program", "Show a line on the status bar.", ["text"], null),
        ("log", "program", "Write a line to dAHKxodus's log (logs\\dahkxodus.log).", ["text"], null),
        ("path", "query", "The document's file; \"\" until it is saved.", [], "string"),
        ("name", "query", "The document's name, as the title bar shows it.", [], "string"),
        ("modified", "query", "Whether the document has changes not saved.", [], "true or false"),
        ("appVersion", "query", "dAHKxodus's version.", [], "string"),
        ("snippets", "query", "The snippets (Insert > Snippets): an array of {name, file, kind (html or text), hotkey}.", [], "array"),
        ("insertSnippet", "edit", "Insert a snippet at the caret, replacing the selection: its name (or file name) in the snippets folder. An HTML snippet keeps its own formatting; a text one takes the formatting where it lands.", ["name"], "true once inserted"),
        ("saveSnippet", "edit", "Keep the selection as a snippet called name (one of that name is replaced).", ["name"], "true once saved, false with nothing selected"),
        ("hooks", "hooks", "The event hooks: an array of {event, macro, on}, one per event (macro \"\": none).", [], "array"),
        ("setHook", "hooks", "Give an event (Started, Opened, NewDocument, BeforeSave, Saved, PdfExported, Closing) a macro to run: a file name in the macros folders, or a full path; \"\" takes it away.", ["event", "macro"], null),
        ("setHookOn", "hooks", "Switch an event's hook on or off; its macro stays chosen.", ["event", "on"], null),
    ];

    async Task RunScriptCallAsync(IntPtr caller, string text) {
        JsonObject reply;
        var id = 0;
        string name = "";
        await scriptTurn.WaitAsync();
        try {
            var call = JsonNode.Parse(text) as JsonObject ?? throw new ScriptError("bad_arguments", "the call is not a JSON object");
            id = call["id"]?.GetValue<int>() ?? 0;
            name = call["name"]?.GetValue<string>() ?? "";
            reply = await ScriptCommandAsync(caller, name, call["args"] as JsonObject ?? new JsonObject());
        } catch (ScriptError ex) {
            reply = new JsonObject { ["error"] = ex.Message, ["code"] = ex.Code };
        } catch (Exception ex) {
            Log($"script: {name} failed: {ex}");
            reply = new JsonObject { ["error"] = $"{name}: {ex.Message}", ["code"] = "error" };
        } finally {
            scriptTurn.Release();
        }
        reply["id"] = id;
        if (IsWindow(caller) && !SendCopyData(caller, DocReply, reply.ToJsonString(), 5000)) Log($"script: the answer to {name} was not taken");
    }

    static JsonObject Ok(JsonNode? value) => new() { ["result"] = value };

    // A script's merge values: "values", an object (field name: text) or a saved record's name. Null when not given.
    static Dictionary<string, string>? MergeValues(string name, JsonObject args) {
        Dictionary<string, string>? values = null;
        var given = args["values"] switch {
            JsonObject o => o,
            JsonValue v when v.TryGetValue<string>(out var s) && s.Trim().StartsWith('{') =>
                JsonNode.Parse(s) as JsonObject ?? throw new ScriptError("bad_arguments", $"{name}: \"values\" must be an object or a record's name"),
            JsonValue v when v.TryGetValue<string>(out var record) && record.Trim() != "" =>
                (values = MergeRecords.Load(record.Trim())) != null ? null
                    : throw new ScriptError("refused", $"{name}: there is no record \"{record}\" in {MergeRecords.Folder}"),
            null => null,
            JsonValue v when v.TryGetValue<string>(out _) => null,
            _ => throw new ScriptError("bad_arguments", $"{name}: \"values\" must be an object or a record's name"),
        };
        if (given != null) {
            values ??= new(StringComparer.OrdinalIgnoreCase);
            foreach (var (k, v) in given)
                values[k] = v is JsonValue jv && jv.TryGetValue<string>(out var s) ? s
                    : v is JsonValue b && b.TryGetValue<bool>(out var flag) ? (flag ? "True" : "False")
                    : v?.ToJsonString() ?? "";
        }
        return values;
    }

    async Task<JsonObject> ScriptCommandAsync(IntPtr caller, string name, JsonObject args) {
        string Arg(string key) => args[key] is JsonValue v && v.TryGetValue<string>(out var s) ? s
            : throw new ScriptError("bad_arguments", $"{name}: argument \"{key}\" must be a string");
        string? OptArg(string key) => args[key] is null ? null : Arg(key);
        void NotAsking() { if (asking) throw new ScriptError("busy", $"{name}: dAHKxodus is waiting for the user (a dialog is open)"); }
        if (!pageReady || web.CoreWebView2 == null) throw new ScriptError("closed", $"{name}: dAHKxodus is still starting");

        switch (name) {
            case "appVersion": return Ok(Version);
            case "path": return Ok(current);
            case "name": return Ok(DocName);
            case "modified": return Ok(recovered || await IsDirtyAsync());
            case "setStatus": SendStatus(Arg("text")); return Ok(null);
            case "log": Log("script: " + Arg("text")); return Ok(null);
            case "subscribe":
            case "unsubscribe":
                var ev = Arg("event");
                if (!subscribers.TryGetValue(caller, out var set)) subscribers[caller] = set = new(StringComparer.OrdinalIgnoreCase);
                if (name == "subscribe") set.Add(ev); else set.Remove(ev);
                if (set.Count == 0) subscribers.Remove(caller);
                return Ok(null);
            case "commands":
                var list = await PageCallAsync("window.dax.commands()") as JsonArray ?? new JsonArray();
                foreach (var c in HostCommands) {
                    var info = new JsonObject {
                        ["name"] = c.name, ["group"] = c.group, ["description"] = c.description,
                        ["args"] = new JsonObject(c.args.Select(a => KeyValuePair.Create(a, (JsonNode?)(a.StartsWith("values") ? "object or string" : "string")))),
                        ["order"] = new JsonArray(c.args.Select(a => (JsonNode?)a).ToArray()),
                    };
                    if (c.returns != null) info["returns"] = c.returns;
                    list.Add(info);
                }
                return Ok(list);
            case "new":
                NotAsking();
                var template = OptArg("template") ?? "";
                if (template != "" && !Path.IsPathRooted(template)) {
                    // (a name alone: the templates folder's file of that name, .dotx first)
                    template = new[] { "", ".dotx", ".dotm", ".docx" }.Select(ext => Path.Combine(Templates.Folder, template + ext))
                        .FirstOrDefault(File.Exists) ?? Path.Combine(Templates.Folder, template);
                }
                if (template != "" && !File.Exists(template)) throw new ScriptError("refused", $"new: there is no template {template}");
                var newValues = MergeValues(name, args);
                if (!await CanDiscard(hooks: false)) return Ok(false);
                return Ok(await WaitForDocument(() => { if (template == "") { StartBlank(); return true; } return OpenFile(template, quiet: true, asNew: true, newValues, askMerge: false); }));
            case "mergeFields": {
                var bytes = await SnapshotAsync() ?? throw new ScriptError("error", "mergeFields: the document could not be read");
                return Ok(new JsonArray(Merge.Scan(bytes).Select(f => (JsonNode?)new JsonObject {
                    ["name"] = f.Name, ["kind"] = f.Kind.ToString().ToLowerInvariant(), ["label"] = f.Label,
                    ["default"] = f.Default, ["yesNo"] = f.InCondition,
                }).ToArray()));
            }
            case "merge": {
                NotAsking();
                var mergeValues = MergeValues(name, args);
                var bytes = await SnapshotAsync() ?? throw new ScriptError("error", "merge: the document could not be read");
                var fields = Merge.Scan(bytes);
                if (fields.Count == 0) return Ok(false);
                byte[]? filled;
                if (mergeValues != null) filled = Merge.Fill(bytes, mergeValues);
                else {
                    var done = new TaskCompletionSource<byte[]?>();
                    var written = false;
                    BeginInvoke(() => done.SetResult(AskMerge(bytes, fields, DocName, fromTemplate: false, out written)));
                    filled = await done.Task;
                    if (written) return Ok(true);    // (to files: this window's document stays)
                    if (filled == null) return Ok(false);
                }
                // (this window's document becomes the merged one, untitled: the script goes on with it)
                if (!await CanDiscard(hooks: false)) return Ok(false);
                var temp = Path.Combine(Path.GetTempPath(), $"dAHKxodus-merge-{Guid.NewGuid():N}.docx");
                File.WriteAllBytes(temp, filled);
                MergedFiles.Add(temp);
                return Ok(await WaitForDocument(() => OpenFile(temp, quiet: true, asNew: true)));
            }
            case "open":
                NotAsking();
                var path = Arg("path");
                if (!File.Exists(path)) throw new ScriptError("refused", $"open: there is no file {path}");
                if (!await CanDiscard(hooks: false)) return Ok(false);
                return Ok(await WaitForDocument(() => OpenFile(path, quiet: true)));
            case "save":
                NotAsking();
                var target = OptArg("path");
                return Ok(await SaveAsync(saveAs: false, target is null or "" ? null : Path.GetFullPath(target), hooks: false));
            case "exportPdf":
                NotAsking();
                return Ok(await ExportPdfAsync(Path.GetFullPath(Arg("path")), quiet: true));
            case "saveCopy":
                NotAsking();
                var copyPath = Path.GetFullPath(Arg("path"));
                if (!Formats.Supported(copyPath)) throw new ScriptError("bad_arguments", "saveCopy: a copy is .html, .md or .txt, not " + Path.GetExtension(copyPath));
                return Ok(await SaveCopyAsync(copyPath, quiet: true));
            case "print":
                NotAsking();
                BeginInvoke(Print);
                return Ok(null);
            case "close":
                BeginInvoke(Close);    // (after this answer)
                return Ok(null);
            case "newWindow": {
                var file = OptArg("path") is { Length: > 0 } np ? Path.GetFullPath(np) : "";
                if (file != "" && !File.Exists(file)) throw new ScriptError("refused", $"newWindow: there is no file {file}");
                var window = OpenWindow(file, this);
                var ready = await Task.WhenAny(window.firstDocument.Task, Task.Delay(60000));
                if (ready != window.firstDocument.Task || !window.firstDocument.Task.Result)
                    throw new ScriptError("error", "newWindow: the new window did not get its document");
                return Ok((long)window.Handle);
            }
            case "setView": {
                var v = new JsonObject();
                if (args["paginated"] is JsonValue pv)
                    v["paginated"] = pv.TryGetValue<bool>(out var pb) ? pb
                        : pv.TryGetValue<double>(out var pn) ? pn != 0
                        : pv.TryGetValue<string>(out var ps) && ps.Trim() is "1" or "true" or "True" or "on";
                if (args["zoom"] is JsonValue zv) {
                    if (zv.TryGetValue<double>(out var zn)) v["zoom"] = zn;
                    else if (zv.TryGetValue<string>(out var zs) && zs.Trim().ToLowerInvariant() is var zt) {
                        if (zt is "width" or "page") v["zoom"] = zt;
                        else if (double.TryParse(zt.TrimEnd('%'), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var zp))
                            v["zoom"] = zt.EndsWith('%') ? zp / 100 : zp;
                        else throw new ScriptError("bad_arguments", "setView: zoom is \"width\", \"page\" or a scale (1 = 100%)");
                    }
                }
                web.CoreWebView2.PostWebMessageAsJson(new JsonObject { ["type"] = "applyView", ["view"] = v }.ToJsonString());
                return Ok(null);
            }
            case "view": {
                var paginated = view != null && view.TryGetValue("Paginated", out var vp) && vp == "1";
                JsonNode? zoom = view != null && view.TryGetValue("Zoom", out var vz)
                    ? (double.TryParse(vz, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var zd) ? zd : vz)
                    : null;
                return Ok(new JsonObject { ["paginated"] = paginated, ["zoom"] = zoom });
            }
            case "snippets":
                var snippetKeys = Snippets.Hotkeys();
                return Ok(new JsonArray(Snippets.List().Select(sn => (JsonNode?)new JsonObject {
                    ["name"] = sn.Name, ["file"] = sn.File, ["kind"] = sn.Html ? "html" : "text", ["hotkey"] = snippetKeys.GetValueOrDefault(sn.File) ?? "",
                }).ToArray()));
            case "insertSnippet": {
                NotAsking();
                var sn = Snippets.Find(Arg("name")) ?? throw new ScriptError("refused", $"insertSnippet: there is no snippet \"{Arg("name")}\" in {Snippets.Folder}");
                return Ok(await InsertSnippetAsync(sn.File));
            }
            case "saveSnippet": {
                var snippetName = Arg("name").Trim();
                if (!Snippets.ValidName(snippetName)) throw new ScriptError("bad_arguments", $"saveSnippet: \"{snippetName}\" can't be a file name");
                var html = (await PageAwaitAsync("window.dax.exec(\"selectedHtml\")") as JsonObject)?["result"]?.ToString() ?? "";
                if (html == "") return Ok(false);
                var saved = Snippets.Save(snippetName, html);
                Log("snippet saved: " + saved);
                Everywhere(w => w.SendSnippets());
                return Ok(true);
            }
            case "hooks":
                var hookList = EventHooks.All();
                var hooksOff = EventHooks.Off();
                return Ok(new JsonArray(Enum.GetValues<HookEvent>().Select(ev => (JsonNode?)new JsonObject {
                    ["event"] = ev.ToString(), ["macro"] = hookList.GetValueOrDefault(ev.ToString()) ?? "", ["on"] = !hooksOff.Contains(ev.ToString()),
                }).ToArray()));
            case "setHook":
            case "setHookOn":
                if (!EventHooks.TryParse(Arg("event"), out var hookEvent))
                    throw new ScriptError("bad_arguments", $"{name}: there is no event \"{Arg("event")}\" (events: {string.Join(", ", Enum.GetNames<HookEvent>())})");
                if (name == "setHook") EventHooks.Set(hookEvent, Arg("macro"));
                else {
                    var onArg = args["on"] as JsonValue ?? throw new ScriptError("bad_arguments", "setHookOn: argument \"on\" is missing");
                    var on = onArg.TryGetValue<bool>(out var bv) ? bv : onArg.TryGetValue<int>(out var iv) ? iv != 0
                        : onArg.TryGetValue<string>(out var sv) && sv.Trim() is "1" or "true" or "True" or "on";
                    EventHooks.SetOn(hookEvent, on);
                }
                Log($"script: hook {hookEvent} = \"{EventHooks.MacroFor(hookEvent)}\"{(EventHooks.IsOn(hookEvent) ? "" : " (off)")}");
                Everywhere(w => w.SendHooks());
                return Ok(null);
            case "measureDone":
                if (measureOut == "") break;
                var good = args["ok"] is JsonValue okv && (okv.TryGetValue<bool>(out var b) ? b : okv.TryGetValue<int>(out var i) && i == 1);
                Log($"measure: doc api {(good ? "ok" : "FAILED")}: {OptArg("text")}");
                measureDocDone?.TrySetResult(good);
                return Ok(null);
        }
        // An editor command: the page runs it (ribbon.exec) and answers {result} or {error, code}; a command that takes
        // a while (paste) answers when it is done.
        var answer = await PageAwaitAsync($"window.dax.exec({JsonSerializer.Serialize(name)}, {args.ToJsonString()})") as JsonObject
            ?? throw new ScriptError("error", $"{name}: the editor did not answer");
        if (answer["error"] is JsonNode error)
            throw new ScriptError(answer["code"]?.GetValue<string>() ?? "error", error.GetValue<string>());
        return Ok(answer["result"]?.DeepClone());
    }

    // A script in the page; its value as JSON.
    async Task<JsonNode?> PageCallAsync(string script) => JsonNode.Parse(await web.CoreWebView2.ExecuteScriptAsync(script));

    // A script in the page whose value may be a promise: its value, once settled, as JSON (ExecuteScriptAsync answers
    // a promise as {}). Through DevTools' Runtime.evaluate, in the page's own world.
    async Task<JsonNode?> PageAwaitAsync(string script) {
        var json = await web.CoreWebView2.CallDevToolsProtocolMethodAsync("Runtime.evaluate",
            JsonSerializer.Serialize(new { expression = script, awaitPromise = true, returnByValue = true }));
        var reply = JsonNode.Parse(json);
        if (reply?["exceptionDetails"] is JsonNode failure)
            throw new ScriptError("error", failure["exception"]?["description"]?.ToString() ?? failure["text"]?.ToString() ?? "the page failed");
        return reply?["result"]?["value"]?.DeepClone();
    }

    // Starts an Open or a New and waits for the page to have the document.
    async Task<bool> WaitForDocument(Func<bool> start) {
        var wait = new TaskCompletionSource<bool>();
        documentWait = wait;
        if (!start()) { documentWait = null; return false; }
        var first = await Task.WhenAny(wait.Task, Task.Delay(180_000));
        if (documentWait == wait) documentWait = null;
        return first == wait.Task && wait.Task.Result;
    }

    // The page has a document (opened, or a new one): a script's Open or New is done, and scripts hear Opened.
    // Done when this window's first document is in the editor (Doc.NewWindow waits for it).
    readonly TaskCompletionSource<bool> firstDocument = new(TaskCreationOptions.RunContinuationsAsynchronously);

    void DocumentReady(bool ok) {
        documentWait?.TrySetResult(ok);
        documentWait = null;
        RaiseEvent("Opened", new JsonObject { ["path"] = current, ["fileName"] = DocName });
        ReportModified();
        if (!ok) return;
        // The hooks: Started once, when the first document is in; Opened for each document opened from a file.
        if (!startedHookRun && firstWindow) { startedHookRun = true; _ = RunHookAsync(HookEvent.Started, current); }
        firstDocument.TrySetResult(ok);
        if (current != "" && !newRequested) _ = RunHookAsync(HookEvent.Opened, current);
    }

    // Scripts' Modified event: when the document goes from saved to changed, or back.
    void ReportModified() {
        if (reportedModified == Changed) return;
        reportedModified = Changed;
        RaiseEvent("Modified", new JsonObject { ["modified"] = Changed });
    }

    bool Subscribed(string name) => subscribers.Values.Any(s => s.Contains(name));

    // Tells the scripts that asked for this event ("DAX6"); a script whose window has gone is forgotten.
    void RaiseEvent(string name, JsonObject? data = null) {
        if (!Subscribed(name)) return;
        var message = data ?? new JsonObject();
        message["event"] = name;
        var text = message.ToJsonString();
        foreach (var (window, events) in subscribers.ToList()) {
            if (!IsWindow(window)) { subscribers.Remove(window); continue; }
            if (events.Contains(name)) SendCopyData(window, DocEvent, text, 2000);
        }
    }

    // ---- Event hooks (host\Hooks.cs): a macro dAHKxodus runs by itself when something happens

    // Answers from dAHKxodus.ahk that the host itself waits for (a hook's macro started: its process id).
    readonly Dictionary<int, TaskCompletionSource<(bool ok, string text)>> pendingAhk = new();
    int nextAhk = 1_000_000;    // (apart from the page's ids)
    // The events whose macro is running: the event doesn't start it again meanwhile.
    readonly HashSet<HookEvent> hooksRunning = new();
    bool newRequested, startedHookRun, closingHook;
    // (measure runs) the hooks run: event, detail, how it went
    readonly List<string> hookRuns = new();

    // Asks dAHKxodus.ahk for something and waits for its answer (at most 10 s).
    async Task<(bool ok, string text)> AhkAsync(long tag, string text) {
        if (Ahk == IntPtr.Zero) return (false, Lang.T("bridge.notRunning"));
        var id = ++nextAhk;
        var done = new TaskCompletionSource<(bool, string)>();
        pendingAhk[id] = done;
        if (!SendCopyData(Ahk, tag, id + "\t" + text, 10000)) { pendingAhk.Remove(id); return (false, Lang.T("bridge.noAnswer")); }
        var first = await Task.WhenAny(done.Task, Task.Delay(10_000));
        pendingAhk.Remove(id);
        return first == done.Task ? done.Task.Result : (false, Lang.T("bridge.noAnswer"));
    }

    // The event's hook, if it has one switched on and not running already: its macro, told the event and the detail.
    // BeforeSave and Closing wait for it to end (at most EventHooks.WaitSeconds); the others only start it. force: Run
    // now in the dialog (switched off or not).
    async Task RunHookAsync(HookEvent e, string detail, string? macro = null) {
        var force = macro != null;
        macro ??= EventHooks.MacroFor(e);
        if (macro == "" || (!force && !EventHooks.IsOn(e)) || hooksRunning.Contains(e)) return;
        var path = EventHooks.PathOf(macro, dir);
        if (!File.Exists(path)) {
            Log($"hook {e}: no macro {path}");
            SendStatus("⚠ " + Lang.T("hooks.missing", e, macro));
            return;
        }
        hooksRunning.Add(e);
        try {
            Log($"hook {e}: {Path.GetFileName(path)} \"{detail}\"");
            var (ok, answer) = await AhkAsync(RunMacroTag, path + "\t" + e + "\t" + detail);
            if (!ok) {
                Log($"hook {e} not started: {answer}");
                SendStatus("⚠ " + Lang.T("hooks.failed", e, answer));
                hookRuns.Add($"{e} not started");
                return;
            }
            Process? process = null;
            try { if (int.TryParse(answer, out var pid)) process = Process.GetProcessById(pid); } catch (ArgumentException) { }    // (ended already)
            if (EventHooks.Waits(e) && process != null) {
                var ended = await Task.WhenAny(process.WaitForExitAsync(), Task.Delay(EventHooks.WaitSeconds * 1000));
                if (!process.HasExited) { Log($"hook {e}: still running after {EventHooks.WaitSeconds} s, not waited for"); SendStatus(Lang.T("hooks.slow", e, EventHooks.WaitSeconds)); }
            } else if (process != null) await process.WaitForExitAsync();
            hookRuns.Add($"{e} {detail}");
        } finally {
            hooksRunning.Remove(e);
        }
    }

    // The page saves through the host when a BeforeSave hook is on (it runs before the document's bytes are taken).
    // The list levels new lists start from (List levels > Also for new documents): a JSON array of levels, as the
    // editor's setListLevels takes them. A measure run starts from the engine's own.
    void SendListDefaults() {
        var levels = measureOut == "" ? Settings.Get("Lists", "Defaults") : "";
        if (levels != "") Send(new { type = "listDefaults", levels });
    }

    // Find's options (the find bar's Options menu: Match case, Find whole words only, Use wildcards...): [Find] Options,
    // the names of those that are on, comma-separated (matchCase,wholeWord,matchPrefix,matchSuffix,ignorePunctuation,
    // ignoreSpace,wildcards,regex). Sent when the page is ready and again when the window is active and they changed.
    // A measure run starts with none.
    string? sentFindOptions;

    void SendFindOptions() {
        var options = measureOut == "" ? Settings.Get("Find", "Options") : "";
        if (options == sentFindOptions) return;
        sentFindOptions = options;
        Send(new { type = "findOptions", options });
    }

    // Headline Case's word lists (Change Case > Headline Case): headline-case.txt beside the settings, as AhkWord's, to
    // be edited. Made with the editor's own lists the first time (the page sends them); sent to the page when it is
    // ready and when the window is active again after the file changed. A measure run uses the editor's own.
    static string HeadlineListsPath => Path.Combine(AppFolders.Roaming, "headline-case.txt");
    DateTime headlineListsTime;

    void SendHeadlineLists() {
        if (measureOut != "" || web.CoreWebView2 == null) return;
        try {
            if (!File.Exists(HeadlineListsPath)) { Send(new { type = "headlineLists", create = true }); return; }
            var time = File.GetLastWriteTimeUtc(HeadlineListsPath);
            if (time == headlineListsTime) return;
            headlineListsTime = time;
            Send(new { type = "headlineLists", text = File.ReadAllText(HeadlineListsPath) });
        } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) {
            Log("headline case: " + ex.Message);
        }
    }

    void SendHooks() {
        if (web.CoreWebView2 == null) return;
        Send(new { type = "hooks", beforeSave = EventHooks.MacroFor(HookEvent.BeforeSave) != "" && EventHooks.IsOn(HookEvent.BeforeSave) });
    }

    // The Macros tab's Event hooks...: a macro for each event; OK keeps them.
    void ShowHooks() {
        if (asking || measureOut != "") return;
        asking = true;
        try {
            using var dlg = new EventHooksDialog(dir, (ev, macro) => _ = RunHookAsync(ev, ev == HookEvent.PdfExported ? "" : current, macro));
            if (dlg.ShowDialog(this) != DialogResult.OK) return;
            EventHooks.Save(dlg.Result, dlg.OffResult);
            Log("hooks: " + string.Join(", ", EventHooks.All().Select(p => $"{p.Key}={p.Value}{(EventHooks.IsOn(Enum.Parse<HookEvent>(p.Key)) ? "" : " (off)")}")));
        } finally { asking = false; }
        Everywhere(w => w.SendHooks());
    }

    // (Measure runs) tests\hook-check.ahk as the BeforeSave and Saved hooks; the page then saves as Ctrl+S does. The
    // BeforeSave hook types a mark at the end, which the saved file must have; then both must have run, in order.
    async Task MeasureHooks() {
        var script = Path.Combine(dir, "tests", "hook-check.ahk");
        EventHooks.Override = new(StringComparer.OrdinalIgnoreCase) { ["BeforeSave"] = script, ["Saved"] = script };
        hookRuns.Clear();
        SendHooks();
        Send(new { type = "measureReply", what = "hooksSet", ok = true });
        var until = DateTime.Now.AddSeconds(30);
        while (DateTime.Now < until && !hookRuns.Any(r => r.StartsWith("Saved"))) await Task.Delay(200);
        EventHooks.Override = new(StringComparer.OrdinalIgnoreCase);
        SendHooks();
        var file = Path.ChangeExtension(measureOut, ".host.docx");
        var marked = false;
        try {
            using var zip = System.IO.Compression.ZipFile.OpenRead(file);
            using var reader = new StreamReader(zip.GetEntry("word/document.xml")!.Open());
            marked = reader.ReadToEnd().Contains("[hook before save]");
        } catch (Exception ex) { Log("measure: hooks: " + ex.Message); }
        var order = string.Join(" | ", hookRuns);
        var good = marked && hookRuns.Count == 2 && hookRuns[0].StartsWith("BeforeSave") && hookRuns[1] == "Saved " + Path.GetFullPath(file);
        Log($"measure: hooks {(good ? "ok" : "FAILED")} (ran: {order}; BeforeSave's mark {(marked ? "in" : "NOT in")} the saved file)");
        Send(new { type = "measureReply", what = "hooks", ok = good });
    }

    // ---- Recorded macros

    // The recording's script: named and saved in the user's Macros folder (the program's may be read-only), where the
    // Macros tab lists it. The user may change the script first, or discard it.
    void SaveRecordedMacro(string script, Dictionary<string, string> taken) {
        if (asking) return;
        string name, hotkey;
        asking = true;
        try {
            using var dlg = new MacroDialog(Macros.UniqueName(Lang.T("macro.defaultName")), script, MacroHotkeys.All(), taken);
            if (dlg.ShowDialog(this) != DialogResult.OK) { SendStatus(Lang.T("macro.discarded")); return; }
            name = dlg.MacroName;
            hotkey = dlg.Hotkey;
            script = dlg.Script;
            var path = Macros.UserPath(name);
            if (File.Exists(path) && MessageBox.Show(this, Lang.T("macro.replace", name), "dAHKxodus",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) {
                SendStatus(Lang.T("macro.notSaved"));
                return;
            }
            Macros.Save(path, script, dir);
            // (its hotkey, or none: a macro it replaces loses its own; another macro may have given it up)
            var hotkeys = dlg.Hotkeys;
            if (hotkey == "") hotkeys.Remove(Path.GetFileName(path)); else hotkeys[Path.GetFileName(path)] = hotkey;
            MacroHotkeys.Save(hotkeys);
            Log($"macro recorded: {Path.GetFileName(path)} ({script.Split('\n').Length} lines){(hotkey != "" ? ", hotkey " + hotkey : "")}");
        } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) {
            Log("macro not saved: " + ex.Message);
            MessageBox.Show(this, Lang.T("macro.saveFailed", ex.Message), "dAHKxodus", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        } finally { asking = false; }
        SendStatus(hotkey != "" ? Lang.T("macro.savedHotkey", name, hotkey) : Lang.T("macro.saved", name));
        Everywhere(w => w.SendMacros());
    }

    // The macros, with their hotkeys, for the page's Macros list, palette and keys; and the user's ribbon buttons, each
    // with its macro (path "": the macro is not there any more).
    void SendMacros() {
        if (web.CoreWebView2 == null) return;
        var entries = Macros.List(dir);
        var buttons = RibbonButtons.All().Select(b => {
            var m = entries.LastOrDefault(x => string.Equals(Path.GetFileName(x.path), b.Macro, StringComparison.OrdinalIgnoreCase));
            return new { label = b.Label, macro = b.Macro, path = m?.path ?? "", title = m?.title ?? "", hotkey = m?.hotkey ?? "" };
        }).ToArray();
        Send(new { type = "macros", entries, buttons });
    }

    // ---- Snippets (host\Snippets.cs): Insert > Snippets lists them; the page sends the selection's HTML to keep, and
    // the host reads a snippet back and pastes it (exec pasteHtml, Keep Source Formatting; a .txt one exec pasteText).

    void SendSnippets() {
        if (web.CoreWebView2 == null) return;
        var hotkeys = Snippets.Hotkeys();
        var entries = Snippets.List().Select(x => new { file = x.File, name = x.Name, html = x.Html, preview = x.Preview, hotkey = hotkeys.GetValueOrDefault(x.File) ?? "" }).ToArray();
        Send(new { type = "snippets", entries });
    }

    // A snippet at the caret, replacing the selection: true once the editor took it.
    async Task<bool> InsertSnippetAsync(string file, string hotkey = "") {
        var snippet = Snippets.List().FirstOrDefault(x => x.File.Equals(file, StringComparison.OrdinalIgnoreCase));
        if (snippet == null) {
            SendStatus(Lang.T("snippets.notFound", Path.GetFileNameWithoutExtension(file)));
            SendSnippets();
            return false;
        }
        try {
            var (html, content) = Snippets.Read(snippet);
            Log("snippet: " + snippet.File + (hotkey != "" ? $" (hotkey {hotkey})" : ""));
            var call = html
                ? "window.dax.exec(\"pasteHtml\", " + JsonSerializer.Serialize(new { html = content, formatting = "keep" }) + ")"
                : "window.dax.exec(\"pasteText\", " + JsonSerializer.Serialize(new { text = content }) + ")";
            var answer = await PageAwaitAsync(call) as JsonObject;
            if (answer?["result"] is JsonValue v && v.TryGetValue<bool>(out var done) && done) return true;
            Log($"snippet {snippet.File} not inserted: {answer?["error"]?.ToString() ?? "the editor refused it"}");
        } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ScriptError or ArgumentException) {
            Log($"snippet {snippet.File} not inserted: {ex.Message}");
        }
        SendStatus(Lang.T("snippets.notInserted", snippet.Name));
        return false;
    }

    // Save Selection as Snippet...: its name asked, then kept.
    void SaveSnippet(string html, string text) {
        if (asking) return;
        if (html == "") { SendStatus(Lang.T("snippets.noSelection")); return; }
        string file;
        asking = true;
        try {
            using var dlg = new SnippetNameDialog(text.Length > 200 ? text[..200] + "…" : text);
            if (dlg.ShowDialog(this) != DialogResult.OK) return;
            try { file = Snippets.Save(dlg.SnippetName, html); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) {
                MessageBox.Show(this, Lang.T("snippets.saveFailed", ex.Message), "dAHKxodus", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        } finally { asking = false; }
        Log($"snippet saved: {file} ({html.Length:N0} characters of HTML)");
        SendStatus(Lang.T("snippets.saved", Path.GetFileNameWithoutExtension(file)));
        Everywhere(w => w.SendSnippets());
    }

    // Manage Snippets...: insert, rename, delete; the list drawn again in every window after.
    void ManageSnippets() {
        if (asking) return;
        Snippets.Entry? insert;
        bool changed;
        asking = true;
        try {
            using var dlg = new SnippetsDialog();
            dlg.ShowDialog(this);
            insert = dlg.ToInsert;
            changed = dlg.Changed;
        } finally { asking = false; }
        if (changed) Everywhere(w => w.SendSnippets());
        if (insert != null) _ = InsertSnippetAsync(insert.File);
    }

    // (Measure runs) The page's selection (its HTML) kept as a snippet in the run's own folder, a text snippet beside
    // it; both inserted at the caret; then a rename takes the hotkey along. The page checks the document's text.
    async Task MeasureSnippets(string html) {
        var notes = new List<string>();
        var ok = true;
        void Check(string what, bool good) { notes.Add($"{what} {(good ? "ok" : "FAILED")}"); ok &= good; }
        try {
            var file = Snippets.Save("Measure", html);
            File.WriteAllText(Path.Combine(Snippets.Folder, "Measure text.txt"), "plain line");
            var list = Snippets.List();
            Check("listed", list.Count == 2 && list[0].File == "Measure.html" && list[0].Html && !list[1].Html);
            Check("read back", Snippets.Read(list[0]).Content == html);
            Check("found by name", Snippets.Find("measure")?.File == "Measure.html" && Snippets.Find("Measure text")?.Html == false);
            Check("html inserted", await InsertSnippetAsync(file));
            Check("text inserted", await InsertSnippetAsync("Measure text.txt"));
            Snippets.SetHotkey(file, "Ctrl+Alt+9");
            var renamed = Snippets.Rename(list[0], "Measure 2");
            Check("rename keeps the hotkey", Snippets.HotkeyOf(renamed) == "Ctrl+Alt+9" && Snippets.HotkeyOf(file) == "");
            SendSnippets();
        } catch (Exception ex) { notes.Add("FAILED: " + ex.Message); ok = false; }
        Log("measure: snippets host " + string.Join(", ", notes));
        Send(new { type = "measureReply", what = "snippets", ok });
    }

    // The Macros tab's Ribbon buttons...: the user's own buttons, each a macro; OK keeps them.
    void ShowRibbonButtons() {
        if (asking || measureOut != "") return;
        asking = true;
        try {
            var files = Macros.List(dir).Select(x => Path.GetFileName(x.path)).Distinct(StringComparer.OrdinalIgnoreCase)
                .Order(StringComparer.OrdinalIgnoreCase).ToList();
            using var dlg = new RibbonButtonsDialog(files, RibbonButtons.All());
            if (dlg.ShowDialog(this) != DialogResult.OK) return;
            RibbonButtons.Save(dlg.Result);
            Log("ribbon buttons: " + string.Join(", ", RibbonButtons.All().Select(b => $"{b.Label}={b.Macro}")));
        } finally { asking = false; }
        Everywhere(w => w.SendMacros());
    }

    // The Macros tab's Hotkeys...: every macro with its hotkey; OK keeps what the user set.
    void ShowHotkeys(Dictionary<string, string> taken) {
        if (asking || measureOut != "") return;
        asking = true;
        try {
            // (the macros, then the snippets: one list, so a hotkey is never both a macro's and a snippet's)
            var macros = Macros.List(dir).Select(x => (Path.GetFileName(x.path), x.label, Lang.T(x.user ? "hotkeys.mine" : "hotkeys.program"))).ToList();
            var snippets = Snippets.List().Select(x => (Snippets.HotkeyPrefix + x.File, x.Name, Lang.T("hotkeys.snippet"))).ToList();
            var all = MacroHotkeys.All();
            foreach (var (file, keys) in Snippets.Hotkeys()) all[Snippets.HotkeyPrefix + file] = keys;
            using var dlg = new MacroHotkeysDialog(macros.Concat(snippets), all, taken);
            if (dlg.ShowDialog(this) != DialogResult.OK) return;
            // (only the macros and snippets there are: a hotkey of one deleted or renamed meanwhile goes)
            var names = macros.Select(x => x.Item1).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var snippetNames = snippets.Select(x => x.Item1).ToHashSet(StringComparer.OrdinalIgnoreCase);
            MacroHotkeys.Save(dlg.Result.Where(p => names.Contains(p.Key)).ToDictionary(p => p.Key, p => p.Value, StringComparer.OrdinalIgnoreCase));
            Snippets.SaveHotkeys(dlg.Result.Where(p => snippetNames.Contains(p.Key))
                .ToDictionary(p => p.Key[Snippets.HotkeyPrefix.Length..], p => p.Value, StringComparer.OrdinalIgnoreCase));
            Log("macro hotkeys: " + string.Join(", ", MacroHotkeys.All().Select(p => $"{p.Key}={p.Value}")));
            Log("snippet hotkeys: " + string.Join(", ", Snippets.Hotkeys().Select(p => $"{p.Key}={p.Value}")));
        } finally { asking = false; }
        Everywhere(w => { w.SendMacros(); w.SendSnippets(); });
    }

    // A macro of the user's (not one that came with the program), to the Recycle Bin once the user says yes.
    void DeleteMacro(string path) {
        if (asking) return;
        var name = Path.GetFileNameWithoutExtension(path);
        var mine = Macros.List(dir).Any(m => m.user && string.Equals(m.path, path, StringComparison.OrdinalIgnoreCase));
        if (!mine) { SendStatus(Lang.T("macro.notYours", name)); return; }
        asking = true;
        try {
            if (MessageBox.Show(this, Lang.T("macro.deleteAsk", name), "dAHKxodus", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            try {
                Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile(path, Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs,
                    Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin);
            } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or OperationCanceledException) {
                Log("macro not deleted: " + ex.Message);
                return;
            }
        } finally { asking = false; }
        Log("macro deleted: " + Path.GetFileName(path));
        // (its hotkey goes too, unless a macro of the same name is left in the program's folder)
        if (!Macros.List(dir).Any(x => string.Equals(Path.GetFileName(x.path), Path.GetFileName(path), StringComparison.OrdinalIgnoreCase)))
            MacroHotkeys.Set(Path.GetFileName(path), "");
        SendStatus(Lang.T("macro.deleted", name));
        Everywhere(w => w.SendMacros());
    }

    // (Measure runs) The recording goes to logs\measure-recorded.ahk, no dialog; measureRunRecorded replays it.
    void MeasureRecorded(string script) {
        var path = Path.Combine(dir, "logs", "measure-recorded.ahk");
        var ok = false;
        try {
            Macros.Save(path, script, dir);
            ok = true;
        } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { Log("measure: recorded macro not written: " + ex.Message); }
        Log($"measure: recorded macro {(ok ? "written" : "FAILED")}: {script.Replace("\r\n", "\n").Replace("\n", " | ")}");
        Send(new { type = "measureReply", what = "recorded", ok });
    }

    // (Measure runs) tests\doc-check.ahk drives this window through Doc.ahk and says how it went (measureDone).
    TaskCompletionSource<bool>? measureDocDone;
    async Task MeasureDocScript() {
        measureDocDone = new TaskCompletionSource<bool>();
        CallAhk(RunMacroTag, 0, Path.Combine(dir, "tests", "doc-check.ahk"));
        var first = await Task.WhenAny(measureDocDone.Task, Task.Delay(60_000));
        var ok = first == measureDocDone.Task && measureDocDone.Task.Result;
        if (first != measureDocDone.Task) Log("measure: doc api FAILED: tests\\doc-check.ahk did not finish in 60 s");
        Send(new { type = "measureReply", what = "docScript", ok });
    }
}

// The Macros tab's list: the .ahk files in the program's macros folder and in the user's (Macros beside
// settings.ini), each with its first comment line as its tip.
static class Macros {
    /// user: one of the user's own (their Macros folder), which they may delete from the list. hotkey: its keys, or "".
    public record Entry(string path, string label, string title, bool user, string hotkey);

    public static List<Entry> List(string programFolder) {
        var list = new List<Entry>();
        var hotkeys = MacroHotkeys.All();
        foreach (var (folder, user) in new[] { (Path.Combine(programFolder, "macros"), false), (UserFolder, true) }) {
            if (!Directory.Exists(folder)) continue;
            foreach (var file in Directory.EnumerateFiles(folder, "*.ahk").Order(StringComparer.OrdinalIgnoreCase))
                list.Add(new Entry(file, Path.GetFileNameWithoutExtension(file), Tip(file), user, hotkeys.GetValueOrDefault(Path.GetFileName(file)) ?? ""));
        }
        return list;
    }

    /// The user's own macros: recorded ones are saved here (Macros beside settings.ini).
    public static string UserFolder => Path.Combine(AppFolders.Roaming, "Macros");

    public static string UserPath(string name) => Path.Combine(UserFolder, name + ".ahk");

    /// "Recorded macro 3": the first number not taken in the user's folder.
    public static string UniqueName(string stem) {
        for (var i = 1; ; i++)
            if (!File.Exists(UserPath($"{stem} {i}"))) return $"{stem} {i}";
    }

    /// Writes a macro (CRLF, UTF-8 without BOM), and beside it Lib\Doc.ahk, a copy of the program's, so its
    /// `#Include <Doc>` finds Doc.ahk wherever the program is (AutoHotkey looks in the script's Lib folder first).
    public static void Save(string path, string script, string programFolder) {
        var folder = Path.GetDirectoryName(Path.GetFullPath(path))!;
        Directory.CreateDirectory(folder);
        UpdateLib(folder, programFolder);
        File.WriteAllText(path, script.Replace("\r\n", "\n").Replace("\n", "\r\n"), new System.Text.UTF8Encoding(false));
    }

    /// Keeps <folder>\Lib\Doc.ahk the same as the program's Doc.ahk (a newer dAHKxodus may have changed it).
    public static void UpdateLib(string folder, string programFolder) {
        var source = Path.Combine(programFolder, "Doc.ahk");
        if (!File.Exists(source) || !Directory.Exists(folder)) return;
        var lib = Path.Combine(folder, "Lib");
        var copy = Path.Combine(lib, "Doc.ahk");
        var bytes = File.ReadAllBytes(source);
        if (File.Exists(copy) && File.ReadAllBytes(copy).AsSpan().SequenceEqual(bytes)) return;
        Directory.CreateDirectory(lib);
        File.WriteAllBytes(copy, bytes);
    }

    static string Tip(string file) {
        try {
            foreach (var line in File.ReadLines(file).Take(30)) {
                var t = line.Trim();
                if (t.StartsWith(';')) return t.TrimStart(';').Trim();
            }
        } catch (IOException) { }
        return "";
    }
}

// The Docxodus engine, called by the page as chrome.webview.hostObjects.sync.engine (a COM object WebView2 hands to
// the page; each call blocks the page's script until it returns, as a WASM call does). Manifest() lists the bridge
// methods and their types; Call(method, argsJson) runs one (DocxodusNative.NativeSessionBridge has the wire forms).
// A result that is null, or an exception, comes back as text starting with a NUL character: "\0null", "\0error:...".
// WebView2 calls it on the window's thread, one call at a time, which is what the engine's sessions require.
[ComVisible(true)]
[ClassInterface(ClassInterfaceType.AutoDual)]
public class EngineHost {
    readonly Dictionary<string, (int n, double ms)> stats = new();

    public string Manifest() => DocxodusNative.NativeSessionBridge.Manifest();

    public string Call(string method, string argsJson) {
        var t = Stopwatch.GetTimestamp();
        try {
            return DocxodusNative.NativeSessionBridge.Call(method, argsJson) ?? "\0null";
        } catch (Exception ex) {
            MainForm.Log($"engine: {method} failed: {ex}");
            return "\0error:" + ex.GetType().Name + ": " + ex.Message;
        } finally {
            var name = method.Substring(method.IndexOf('.') + 1);
            stats.TryGetValue(name, out var s);
            stats[name] = (s.n + 1, s.ms + Stopwatch.GetElapsedTime(t).TotalMilliseconds);
        }
    }

    // A pasted picture the page cannot read itself (paste-html.ts, resolvePasteImage): Word's clipboard HTML names its
    // pictures as files in the temporary folder (file:///C:/Users/.../AppData/Local/Temp/msohtmlclip1/01/clip_image001.png).
    // Answers a data: URL, or "" for anything else: only pictures in the temporary folder, up to 20 MB, are read.
    public string ReadPastedImage(string src) {
        try {
            if (!Uri.TryCreate(src, UriKind.Absolute, out var uri)) return "";
            if (uri.Scheme == Uri.UriSchemeHttps || uri.Scheme == Uri.UriSchemeHttp) return DownloadPastedImage(uri);
            if (!uri.IsFile) return "";
            var path = Path.GetFullPath(uri.LocalPath);
            var temp = Path.GetFullPath(Path.GetTempPath());
            if (!path.StartsWith(temp, StringComparison.OrdinalIgnoreCase)) return "";
            var type = Path.GetExtension(path).ToLowerInvariant() switch {
                ".png" => "image/png", ".jpg" or ".jpeg" => "image/jpeg", ".gif" => "image/gif", ".bmp" => "image/bmp",
                _ => "",
            };
            var info = new FileInfo(path);
            if (type == "" || !info.Exists || info.Length > 20 * 1024 * 1024) return "";
            return $"data:{type};base64,{Convert.ToBase64String(File.ReadAllBytes(path))}";
        } catch (Exception ex) {
            MainForm.Log($"paste: picture {src} not read: {ex.Message}");
            return "";
        }
    }

    // A picture on the web that the page could not fetch itself (the site does not let other pages read it): the host
    // downloads it, as Word does when it pastes a web page. 6 s and 20 MB at most; only images.
    static readonly System.Net.Http.HttpClient web = new() { Timeout = TimeSpan.FromSeconds(6) };
    static string DownloadPastedImage(Uri uri) {
        try {
            using var response = web.GetAsync(uri, System.Net.Http.HttpCompletionOption.ResponseHeadersRead).GetAwaiter().GetResult();
            var type = response.Content.Headers.ContentType?.MediaType ?? "";
            if (!response.IsSuccessStatusCode || !type.StartsWith("image/") || type == "image/svg+xml") return "";
            if (response.Content.Headers.ContentLength > 20 * 1024 * 1024) return "";
            var bytes = response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();
            return bytes.Length > 20 * 1024 * 1024 ? "" : $"data:{type};base64,{Convert.ToBase64String(bytes)}";
        } catch (Exception ex) {
            MainForm.Log($"paste: picture {uri} not downloaded: {ex.Message}");
            return "";
        }
    }

    // "engine N ms; Method Kx M ms; ..." since the last call (busiest first), then starts counting again.
    public string TakeStats() {
        var total = stats.Values.Sum(s => s.ms);
        var top = stats.OrderByDescending(kv => kv.Value.ms).Take(6).Select(kv => $"{kv.Key} {kv.Value.n}x {kv.Value.ms:0} ms").ToList();
        stats.Clear();
        return $"engine {total:0} ms; " + string.Join("; ", top);
    }
}

// Who is editing: the name and initials the editor stamps on comments and tracked changes. The first that has a name:
// settings.ini ([User] Name=, Initials=; host\Settings.cs says where it is), then Office's own user information (what
// Word stamps on the same machine), then the Windows account name. Initials are "" when nothing names them: the
// editor then takes them from the name, as Word does.
static class UserIdentity {
    public static (string Name, string Initials, string From) Read() {
        var name = Settings.Get("User", "Name");
        if (name != "") return (name, Settings.Get("User", "Initials"), "settings.ini");
        return ReadFallback();
    }

    /// Who the user is when settings.ini doesn't say (the Settings dialog shows it).
    public static (string Name, string Initials, string From) ReadFallback() {
        try {
            using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\Common\UserInfo");
            var name = (key?.GetValue("UserName") as string)?.Trim() ?? "";
            if (name != "") return (name, (key?.GetValue("UserInitials") as string)?.Trim() ?? "", Lang.T("identity.office"));
        } catch { }
        return (Environment.UserName, "", Lang.T("identity.windows"));
    }
}
