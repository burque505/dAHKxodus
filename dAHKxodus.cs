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
    public const string Version = "0.0.1.0";
    public static string LogPath = "";
    static readonly DateTime ProcessStart = Process.GetCurrentProcess().StartTime;

    public static void Log(string text) {
        try {
            Directory.CreateDirectory(Path.GetDirectoryName(LogPath)!);
            var since = (DateTime.Now - ProcessStart).TotalMilliseconds;
            File.AppendAllText(LogPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} +{since,7:0} ms  {text}\r\n");
        } catch { }
    }

    const int WM_COPYDATA = 0x004A;
    // "DAX0" (the companion's hello), "DAX1" (a call), "DAX2", "DAX3" (its answer, its error)
    const long Hello = 0x44415830, CallTag = 0x44415831, ReplyOk = 0x44415832, ReplyError = 0x44415833;
    static string Untitled => Lang.T("file.untitled");

    [StructLayout(LayoutKind.Sequential)]
    struct COPYDATASTRUCT { public IntPtr dwData; public int cbData; public IntPtr lpData; }

    [DllImport("user32.dll")] public static extern bool PostMessageW(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
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
    string? reply;
    bool replyIsError;

    // The document: its file ("" until it is saved), and whether it has changed since (the page's ribbon says).
    string current;
    bool dirty;
    bool closeConfirmed;    // the user has answered the ask-to-save question; the next Close goes through
    bool asking;            // a question or a file dialog is up: another request waits its turn
    // Saves the host asked the page for (the page answers "saved" with the id): Save or Save As, and who waits.
    readonly Dictionary<int, (bool saveAs, TaskCompletionSource<bool> done)> pendingSaves = new();
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
        ahk = Lang.T("page.ahk"), ahkTip = Lang.T("page.ahkTip"),
        close = Lang.T("page.close"), closeTip = Lang.T("page.closeTip"), hint = Lang.T("page.hint"),
        print = Lang.T("page.print"), printTip = Lang.T("page.printTip"), pdf = Lang.T("page.pdf"), pdfTip = Lang.T("page.pdfTip"),
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

    void OpenFile(string path) {
        byte[] bytes;
        try {
            bytes = File.ReadAllBytes(path);
        } catch (Exception ex) {
            Log($"open {path} failed: {ex.Message}");
            if (measureOut == "") MessageBox.Show(this, Lang.T("file.openFailed", path, ex.Message), "dAHKxodus", MessageBoxButtons.OK, MessageBoxIcon.Error);
            if (ex is FileNotFoundException or DirectoryNotFoundException) { RecentFiles.Remove(path); SendRecent(); }
            return;
        }
        current = Path.GetFullPath(path);
        dirty = recovered = false;
        ClearRecovery();    // (the document before it went as the user chose)
        UpdateTitle();
        if (measureOut == "") { RecentFiles.Add(current); SendRecent(); }
        Log($"open {Path.GetFileName(path)} ({bytes.Length:N0} bytes) at {clock.ElapsedMilliseconds} ms");
        Send(new { type = "open", name = Path.GetFileName(path), b64 = Convert.ToBase64String(bytes) });
    }

    async void NewDocument() {
        if (!await CanDiscard()) return;
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

    // Asks the page for the document's bytes (it answers "saved" with the id); true once they are written.
    Task<bool> SaveAsync(bool saveAs) {
        var id = ++nextSave;
        var done = new TaskCompletionSource<bool>();
        pendingSaves[id] = (saveAs, done);
        Send(new { type = "save", id });
        return done.Task;
    }

    // The page's bytes for Save (Ctrl+S, the ribbon's Save), or for a save the host asked for (id).
    void OnSaved(byte[] bytes, int id) {
        var saveAs = false;
        TaskCompletionSource<bool>? done = null;
        if (id != 0 && pendingSaves.Remove(id, out var pending)) (saveAs, done) = pending;
        var ok = WriteDocument(bytes, saveAs);
        done?.TrySetResult(ok);
    }

    bool WriteDocument(byte[] bytes, bool saveAs) {
        var path = current;
        // A measure run never writes over the document it opened, nor asks where: its saves go beside its copy.
        if (measureOut != "") path = Path.ChangeExtension(measureOut, ".host.docx");
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
    async Task<bool> CanDiscard() {
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
        return await SaveAsync(saveAs: false);
    }

    protected override async void OnFormClosing(FormClosingEventArgs e) {
        closeReason = e.CloseReason;
        if (closeConfirmed || measureOut != "" || web.CoreWebView2 == null || e.CloseReason == CloseReason.WindowsShutDown) {
            base.OnFormClosing(e);
            return;
        }
        e.Cancel = true;    // (decided below, after asking the page; then closed again)
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
                    Log($"opened {clock.ElapsedMilliseconds} ms (window {Environment.GetEnvironmentVariable("DAHKXODUS_WINDOW") ?? "default"}; open {Str("open")} ms: first window after {Str("firstWindow")} ms, longest frozen {Str("longestGap")} ms; first paint {Str("paint")} ms, {Str("blocks")} blocks)");
                    if (recovering != null) RecoveredOpen();
                    if (measureOut != "") Send(new { type = "measure" });
                    break;
                case "openFailed":
                    Log("open failed in the page: " + Str("error"));
                    if (recovering != null) { recovering = null; recovered = false; }    // (its copy stays for next time)
                    MessageBox.Show(this, Lang.T("file.openFailed", DocName, Str("error")), "dAHKxodus", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    Close();
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
                        ["Zoom"] = m.TryGetProperty("zoom", out var z) && z.ValueKind == JsonValueKind.Number
                            ? z.GetDouble().ToString(System.Globalization.CultureInfo.InvariantCulture) : "",
                        ["Comments"] = Str("comments") is "True" or "true" ? "1" : "0",
                        ["Rail"] = Str("rail") is "True" or "true" ? "1" : "0",
                    };
                    if (view["Zoom"] == "") view.Remove("Zoom");
                    break;
                case "dirty":
                    dirty = Str("dirty") == "True" || Str("dirty") == "true";
                    UpdateTitle();
                    if (measureOut != "") Log($"measure: page says {(dirty ? "changed" : "unchanged")}, title \"{Text}\"");
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
                case "measureSave":    // the measure run's check of the host's own save (SaveAsync)
                    _ = SaveAsync(saveAs: false).ContinueWith(t => Log($"measure: host save {(t.Result ? "written" : "FAILED")}, title \"{Text}\""),
                        TaskScheduler.FromCurrentSynchronizationContext());
                    break;
                case "drop":    // files dropped on the page (postMessageWithAdditionalObjects): open the first document
                    var dropped = e.AdditionalObjects?.OfType<CoreWebView2File>().Select(f => f.Path)
                        .FirstOrDefault(p => p.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) || p.EndsWith(".docm", StringComparison.OrdinalIgnoreCase));
                    if (dropped != null) { Activate(); OpenPath(dropped); }
                    break;
                case "ahk":
                    var result = CallAhk(Str("fn"), Str("arg"));
                    Send(replyIsError
                        ? new { type = "ahkResult", id = m.GetProperty("id").GetInt32(), result = "", error = result }
                        : (object)new { type = "ahkResult", id = m.GetProperty("id").GetInt32(), result, error = "" });
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

    // The page switches to page view at 100% and says the pages' size (inches) and count; null when it can't.
    async Task<(double width, double height, int pages)?> PreparePrintAsync() {
        var id = ++nextPrint;
        var done = new TaskCompletionSource<JsonElement>();
        pendingPrints[id] = done;
        Send(new { type = "preparePrint", id });
        var first = await Task.WhenAny(done.Task, Task.Delay(15_000));
        pendingPrints.Remove(id);
        if (first != done.Task) { Log("print: the page did not get ready"); return null; }
        var r = done.Task.Result;
        if (r.TryGetProperty("error", out var err) && err.GetString() is { Length: > 0 } message) { Log("print: " + message); return null; }
        return (r.GetProperty("widthIn").GetDouble(), r.GetProperty("heightIn").GetDouble(), r.GetProperty("pages").GetInt32());
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
    async void ExportPdf(string? path = null) {
        if (asking || web.CoreWebView2 == null) return;
        if (path == null) {
            asking = true;
            try {
                using var dlg = new SaveFileDialog {
                    Filter = Lang.T("file.filterPdf"), FileName = Path.ChangeExtension(DocName, ".pdf"), OverwritePrompt = true, AddExtension = true,
                    InitialDirectory = current != "" ? Path.GetDirectoryName(current) : "",
                };
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                path = dlg.FileName;
            } finally { asking = false; }
        }
        var size = await PreparePrintAsync();
        var ok = false;
        try {
            if (size != null) {
                var settings = web.CoreWebView2.Environment.CreatePrintSettings();
                settings.PageWidth = size.Value.width;
                settings.PageHeight = size.Value.height;
                settings.MarginTop = settings.MarginBottom = settings.MarginLeft = settings.MarginRight = 0;
                settings.ShouldPrintBackgrounds = true;
                settings.ShouldPrintHeaderAndFooter = false;
                var clock = Stopwatch.StartNew();
                ok = await web.CoreWebView2.PrintToPdfAsync(path, settings);
                Log($"pdf: {(ok ? "written" : "FAILED")} {Path.GetFileName(path)} ({(File.Exists(path) ? new FileInfo(path).Length : 0):N0} bytes, {size.Value.pages} pages, {size.Value.width} x {size.Value.height} in, {clock.ElapsedMilliseconds} ms)");
            }
        } catch (Exception ex) {
            Log("pdf failed: " + ex.Message);
        } finally {
            Send(new { type = "printDone" });
        }
        if (measureOut != "") { MeasurePdf(path, ok, size?.pages ?? 0); return; }
        if (ok) SendStatus(Lang.T("print.pdfSaved", path));
        else MessageBox.Show(this, Lang.T("print.pdfFailed", path), "dAHKxodus", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }

    // (measure runs) The PDF is a PDF, with as many pages as the page view has.
    void MeasurePdf(string path, bool ok, int pages) {
        var text = File.Exists(path) ? System.Text.Encoding.Latin1.GetString(File.ReadAllBytes(path)) : "";
        var inPdf = System.Text.RegularExpressions.Regex.Matches(text, @"/Type\s*/Page[^s]").Count;
        var good = ok && text.StartsWith("%PDF-") && inPdf == pages && pages > 0;
        Log($"measure: pdf {(good ? "ok" : "FAILED")} ({inPdf} pages in the PDF, {pages} in the page view)");
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

    // Calls Bridge.<fn>(arg) in dAHKxodus.ahk. AutoHotkey answers with WM_COPYDATA to this window while we wait
    // (a thread waiting in SendMessage still takes messages sent to it), so the answer is there when SendMessage returns.
    string CallAhk(string fn, string arg) {
        var text = fn + "\t" + arg;
        var ptr = Marshal.StringToHGlobalUni(text);
        try {
            reply = null;
            if (Ahk == IntPtr.Zero) { replyIsError = true; return Lang.T("bridge.notRunning"); }
            var cds = new COPYDATASTRUCT { dwData = (IntPtr)CallTag, cbData = (text.Length + 1) * 2, lpData = ptr };
            var sent = SendMessageTimeoutW(Ahk, WM_COPYDATA, Handle, ref cds, 0, 10000, out _);
            if (sent == IntPtr.Zero || reply == null) { replyIsError = true; return Lang.T("bridge.noAnswer"); }
            return reply;
        } finally {
            Marshal.FreeHGlobal(ptr);
        }
    }

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
                reply = Marshal.PtrToStringUni(cds.lpData, cds.cbData / 2).TrimEnd('\0');
                replyIsError = tag == ReplyError;
                m.Result = (IntPtr)1;
                return;
            }
        }
        base.WndProc(ref m);
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
