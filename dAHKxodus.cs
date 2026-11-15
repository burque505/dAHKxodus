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
#load "host\Recovery.cs"
#load "host\MacroHotkeys.cs"
#load "host\MacroDialog.cs"
#load "host\Hooks.cs"
#load "host\RibbonButtons.cs"
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

var ui = new Thread(() => {
    try {
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        Application.ThreadException += (s, e) => MainForm.Log("UI exception: " + e.Exception);
        try { Application.SetHighDpiMode(HighDpiMode.PerMonitorV2); } catch { }
        Application.EnableVisualStyles();
        Application.Run(new MainForm(ahk, dir, start, measureOut));
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
    public const string Version = "0.0.2.0";
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
    // With any script (Doc.ahk): "DAX4" (a call, JSON {id, name, args}), "DAX5" (its answer, JSON {id, result} or
    // {id, error, code}), "DAX6" (an event, JSON {event, ...}).
    const long Hello = 0x44415830, CallTag = 0x44415831, ReplyOk = 0x44415832, ReplyError = 0x44415833,
        DocCall = 0x44415834, DocReply = 0x44415835, DocEvent = 0x44415836, RunMacroTag = 0x44415837;
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
    readonly string dir, measureOut;
    readonly WebView2 web = new() { Dock = DockStyle.Fill };
    readonly EngineHost Engine = new();
    readonly Stopwatch clock = Stopwatch.StartNew();

    // The document: its file ("" until it is saved), and whether it has changed since (the page's ribbon says).
    string current;
    bool dirty;
    bool closeConfirmed;    // the user has answered the ask-to-save question; the next Close goes through
    bool asking;            // a question or a file dialog is up: another request waits its turn
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

    public MainForm(IntPtr ahk, string dir, string start, string measureOut) {
        Ahk = ahk; this.dir = dir; current = start; this.measureOut = measureOut;
        // (a measure run's copies go to a folder of its own, never among the user's)
        recovery = new Recovery(measureOut != "" ? Path.Combine(dir, "logs", "measure-recovery") : Recovery.DefaultFolder);
        // (a measure run's macro hotkeys are its own: one for the page's check, the user's untouched)
        if (measureOut != "") MacroHotkeys.Override = new(StringComparer.OrdinalIgnoreCase) { ["Insert date.ahk"] = "Ctrl+Alt+J" };
        // (and its event hooks: none, until its check sets some)
        if (measureOut != "") EventHooks.Override = new(StringComparer.OrdinalIgnoreCase);
        // (and its ribbon button: one for the page's check)
        if (measureOut != "") RibbonButtons.Override = [new("Insert date.ahk", "Date")];
        autosaveTimer.Tick += (s, e) => _ = AutosaveAsync();
        // Leaving the window (another program, a dialog of its own excepted) is a moment to keep the copy current; not
        // more than twice a minute.
        Deactivate += (s, e) => { if (!asking && (DateTime.Now - lastAutosave).TotalSeconds >= 30) _ = AutosaveAsync(); };
        Size = new System.Drawing.Size(1400, 950);
        StartPosition = FormStartPosition.CenterScreen;
        Controls.Add(web);
        if (measureOut == "") WindowPlacement.Restore(this);    // (a measure run keeps the defaults: its runs compare)
        UpdateTitle();
        if (measureOut != "") {    // a measure run that goes wrong still closes
            var watchdog = new System.Windows.Forms.Timer { Interval = 120000 };
            watchdog.Tick += (s, e) => { Log("measure: gave up after 120 s"); closeConfirmed = true; Close(); };
            watchdog.Start();
        }
        if (AppFolders.StartCompanion != null) HandleCreated += (s, e) => AppFolders.StartCompanion(Handle);
        Load += async (s, e) => {
            try { await Init(); } catch (Exception ex) { Log("init failed: " + ex); if (measureOut != "") Close(); }
        };
    }

    async Task Init() {
        var data = Path.Combine(AppFolders.Local, "WebView2");
        var env = await CoreWebView2Environment.CreateAsync(null, data);
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
        // The view the last window closed with: page or continuous, zoom, comments, the rail (not for a measure run).
        if (measureOut == "")
            foreach (var key in new[] { "Paginated", "Zoom", "Comments", "Rail" }) {
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

    string DocName => current == "" ? Untitled : Path.GetFileName(current);

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

    // The labels the page adds to the ribbon, in the user's language.
    void SendUi() => Send(new {
        type = "ui", settings = Lang.T("page.settings"), settingsTip = Lang.T("page.settingsTip"),
        ahk = Lang.T("page.ahk"), ahkTip = Lang.T("page.ahkTip"), macros = Lang.T("page.macros"), macrosTip = Lang.T("page.macrosTip"),
        close = Lang.T("page.close"), closeTip = Lang.T("page.closeTip"), hint = Lang.T("page.hint"),
        print = Lang.T("page.print"), printTip = Lang.T("page.printTip"), pdf = Lang.T("page.pdf"), pdfTip = Lang.T("page.pdfTip"),
        palette = Lang.T("page.palette"), paletteTip = Lang.T("page.paletteTip"), palettePlaceholder = Lang.T("page.palettePlaceholder"),
        paletteOn = Lang.T("page.paletteOn"), paletteNotNow = Lang.T("page.paletteNotNow"), paletteFile = Lang.T("page.paletteFile"),
        paletteMacro = Lang.T("page.paletteMacro"), paletteTitleBar = Lang.T("page.paletteTitleBar"),
        paletteStatusBar = Lang.T("page.paletteStatusBar"), paletteNone = Lang.T("page.paletteNone"),
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
        // What changed applies now: the host's language, the recent list, and the name on new comments and changes.
        Lang.Load();
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
        if (measureOut == "") {
            try {
                WindowPlacement.Save(this);
                if (view != null)
                    Settings.Update(() => { foreach (var (k, v) in view) Settings.Set("View", k, v); });
            } catch (Exception ex) { Log("settings at close: " + ex.Message); }
        }
        base.OnFormClosed(e);
    }

    void SendRecent() {
        if (web.CoreWebView2 == null) return;
        var entries = RecentFiles.List().Select(p => new { id = p, name = Path.GetFileName(p), title = p }).ToArray();
        Send(new { type = "recent", entries });
    }

    // ---- Opening

    // False when the file can't be read (the user is told, unless quiet); else the page is opening it.
    bool OpenFile(string path, bool quiet = false) {
        byte[] bytes;
        try {
            bytes = File.ReadAllBytes(path);
        } catch (Exception ex) {
            Log($"open {path} failed: {ex.Message}");
            if (measureOut == "" && !quiet) MessageBox.Show(this, Lang.T("file.openFailed", path, ex.Message), "dAHKxodus", MessageBoxButtons.OK, MessageBoxIcon.Error);
            if (ex is FileNotFoundException or DirectoryNotFoundException) { RecentFiles.Remove(path); SendRecent(); }
            return false;
        }
        current = Path.GetFullPath(path);
        dirty = recovered = false;
        ClearRecovery();    // (the document before it went as the user chose)
        UpdateTitle();
        if (measureOut == "") { RecentFiles.Add(current); SendRecent(); }
        Log($"open {Path.GetFileName(path)} ({bytes.Length:N0} bytes) at {clock.ElapsedMilliseconds} ms");
        Send(new { type = "open", name = Path.GetFileName(path), b64 = Convert.ToBase64String(bytes) });
        return true;
    }

    async void NewDocument() {
        if (!await CanDiscard()) return;
        StartBlank();
    }

    void StartBlank() {
        newRequested = true;    // (the page's blankOpened then runs the NewDocument hook)
        current = "";
        dirty = recovered = false;
        ClearRecovery();
        UpdateTitle();
        Send(new { type = "blank", name = Untitled });
    }

    async void OpenWithDialog() {
        if (!await CanDiscard()) return;
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
        OpenFile(path);
    }

    async void OpenPath(string path) {
        if (current != "" && string.Equals(Path.GetFullPath(path), current, StringComparison.OrdinalIgnoreCase) && !Changed) return;
        if (!await CanDiscard()) return;
        OpenFile(path);
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
            asking = true;
            try {
                using var dlg = new SaveFileDialog {
                    Filter = Lang.T("file.filterSave"), FileName = DocName, OverwritePrompt = true, AddExtension = true,
                    FilterIndex = DocName.EndsWith(".docm", StringComparison.OrdinalIgnoreCase) ? 2 : 1,
                    InitialDirectory = current != "" ? Path.GetDirectoryName(current) : "",
                };
                if (dlg.ShowDialog(this) != DialogResult.OK) return false;
                path = dlg.FileName;
            } finally { asking = false; }
        }
        try {
            WriteAtomically(path, bytes);
        } catch (Exception ex) {
            Log($"save {path} failed: {ex.Message}");
            MessageBox.Show(this, Lang.T("file.saveFailed", path, ex.Message), "dAHKxodus", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
        current = Path.GetFullPath(path);
        Log($"saved {Path.GetFileName(current)} ({bytes.Length:N0} bytes)");
        if (measureOut == "") { RecentFiles.Add(current); SendRecent(); }
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
                    if (measureOut == "") SendRecent();
                    SendMacros();    // (their hotkeys work from the start)
                    SendHooks();
                    pageReady = true;
                    StartAutosave();
                    // A copy left behind by a dAHKxodus that didn't close normally, offered before anything opens.
                    if (OfferRecovery()) break;
                    // Started without a document: the last one, if the user wants that (and it is still there).
                    if (current == "" && measureOut == "" && Settings.GetBool("General", "OpenLastAtStart", false))
                        current = RecentFiles.List().FirstOrDefault(File.Exists) ?? "";
                    if (current != "") OpenFile(current);
                    else Send(new { type = "blank", name = Untitled });
                    break;
                case "opened":
                    newRequested = false;
                    Log($"opened {clock.ElapsedMilliseconds} ms (window {Environment.GetEnvironmentVariable("DAHKXODUS_WINDOW") ?? "default"}; open {Str("open")} ms: first window after {Str("firstWindow")} ms, longest frozen {Str("longestGap")} ms; first paint {Str("paint")} ms, {Str("blocks")} blocks)");
                    if (recovering != null) RecoveredOpen();
                    if (measureOut != "" && !measureStarted) { measureStarted = true; Send(new { type = "measure" }); }
                    DocumentReady(true);
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
                    };
                    if (view["Zoom"] == "") view.Remove("Zoom");
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
                case "saveRequest":    // Save (Ctrl+S, the ribbon's) while a BeforeSave hook is set: the hook, then the save
                    _ = SaveAsync(saveAs: false);
                    break;
                case "ribbonButtonsRequest":    // the Macros tab's Ribbon buttons...
                    BeginInvoke(ShowRibbonButtons);
                    break;
                case "hooksRequest":    // the Macros tab's Event hooks...
                    BeginInvoke(ShowHooks);
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
                    _ = SaveAsync(saveAs: false).ContinueWith(t => Log($"measure: host save {(t.Result ? "written" : "FAILED")}, title \"{Text}\""),
                        TaskScheduler.FromCurrentSynchronizationContext());
                    break;
                case "drop":    // files dropped on the page (postMessageWithAdditionalObjects): open the first document
                    var dropped = e.AdditionalObjects?.OfType<CoreWebView2File>().Select(f => f.Path)
                        .FirstOrDefault(p => p.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) || p.EndsWith(".docm", StringComparison.OrdinalIgnoreCase));
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
        web.CoreWebView2.ShowPrintUI(CoreWebView2PrintDialogKind.Browser);
    }

    // File > Export to PDF: asks where (the document's name and folder to start), writes the PDF, puts the view back.
    async void ExportPdf(string? path = null) => await ExportPdfAsync(path);

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
    bool SendCopyData(IntPtr to, long tag, string text, int timeout) {
        var ptr = Marshal.StringToHGlobalUni(text);
        try {
            var cds = new COPYDATASTRUCT { dwData = (IntPtr)tag, cbData = (text.Length + 1) * 2, lpData = ptr };
            return SendMessageTimeoutW(to, WM_COPYDATA, Handle, ref cds, SMTO_ABORTIFHUNG, timeout, out var answer) != IntPtr.Zero
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
        ("new", "file", "Start a new document (asks to save the one open first, if it has changes).", [], "true, or false when the user cancelled"),
        ("open", "file", "Open a .docx (asks to save the one open first, if it has changes).", ["path"], "true, or false when it could not be opened or the user cancelled"),
        ("save", "file", "Save the document: to its file (asking where if it has none), or to path.", ["path?"], "true once written, false when not"),
        ("exportPdf", "file", "Write the document's pages as a PDF.", ["path"], "true once written"),
        ("print", "file", "Show the print dialog.", [], null),
        ("close", "file", "Close dAHKxodus (asks to save a changed document).", [], null),
        ("setStatus", "program", "Show a line on the status bar.", ["text"], null),
        ("log", "program", "Write a line to dAHKxodus's log (logs\\dahkxodus.log).", ["text"], null),
        ("path", "query", "The document's file; \"\" until it is saved.", [], "string"),
        ("name", "query", "The document's name, as the title bar shows it.", [], "string"),
        ("modified", "query", "Whether the document has changes not saved.", [], "true or false"),
        ("appVersion", "query", "dAHKxodus's version.", [], "string"),
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
                        ["args"] = new JsonObject(c.args.Select(a => KeyValuePair.Create(a, (JsonNode?)"string"))),
                        ["order"] = new JsonArray(c.args.Select(a => (JsonNode?)a).ToArray()),
                    };
                    if (c.returns != null) info["returns"] = c.returns;
                    list.Add(info);
                }
                return Ok(list);
            case "new":
                NotAsking();
                if (!await CanDiscard(hooks: false)) return Ok(false);
                return Ok(await WaitForDocument(() => { StartBlank(); return true; }));
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
            case "print":
                NotAsking();
                BeginInvoke(Print);
                return Ok(null);
            case "close":
                BeginInvoke(Close);    // (after this answer)
                return Ok(null);
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
                SendHooks();
                return Ok(null);
            case "measureDone":
                if (measureOut == "") break;
                var good = args["ok"] is JsonValue okv && (okv.TryGetValue<bool>(out var b) ? b : okv.TryGetValue<int>(out var i) && i == 1);
                Log($"measure: doc api {(good ? "ok" : "FAILED")}: {OptArg("text")}");
                measureDocDone?.TrySetResult(good);
                return Ok(null);
        }
        // An editor command: the page runs it (ribbon.exec) and answers {result} or {error, code}.
        var answer = await PageCallAsync($"window.dax.exec({JsonSerializer.Serialize(name)}, {args.ToJsonString()})") as JsonObject
            ?? throw new ScriptError("error", $"{name}: the editor did not answer");
        if (answer["error"] is JsonNode error)
            throw new ScriptError(answer["code"]?.GetValue<string>() ?? "error", error.GetValue<string>());
        return Ok(answer["result"]?.DeepClone());
    }

    // A script in the page; its value as JSON.
    async Task<JsonNode?> PageCallAsync(string script) => JsonNode.Parse(await web.CoreWebView2.ExecuteScriptAsync(script));

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
    void DocumentReady(bool ok) {
        documentWait?.TrySetResult(ok);
        documentWait = null;
        RaiseEvent("Opened", new JsonObject { ["path"] = current, ["fileName"] = DocName });
        ReportModified();
        if (!ok) return;
        // The hooks: Started once, when the first document is in; Opened for each document opened from a file.
        if (!startedHookRun) { startedHookRun = true; _ = RunHookAsync(HookEvent.Started, current); }
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
        SendHooks();
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
        SendMacros();
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
        SendMacros();
    }

    // The Macros tab's Hotkeys...: every macro with its hotkey; OK keeps what the user set.
    void ShowHotkeys(Dictionary<string, string> taken) {
        if (asking || measureOut != "") return;
        asking = true;
        try {
            var macros = Macros.List(dir).Select(x => (Path.GetFileName(x.path), x.label, x.user)).ToList();
            using var dlg = new MacroHotkeysDialog(macros, MacroHotkeys.All(), taken);
            if (dlg.ShowDialog(this) != DialogResult.OK) return;
            // (only the macros there are: a hotkey of one deleted or renamed meanwhile goes)
            var names = macros.Select(x => x.Item1).ToHashSet(StringComparer.OrdinalIgnoreCase);
            MacroHotkeys.Save(dlg.Result.Where(p => names.Contains(p.Key)).ToDictionary(p => p.Key, p => p.Value, StringComparer.OrdinalIgnoreCase));
            Log("macro hotkeys: " + string.Join(", ", MacroHotkeys.All().Select(p => $"{p.Key}={p.Value}")));
        } finally { asking = false; }
        SendMacros();
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
        SendMacros();
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
