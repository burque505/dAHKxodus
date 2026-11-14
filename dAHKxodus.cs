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
#load "dax_extensions.cs"
// dAHKxodus: a WebView2 window showing Docxodus's ribbon editor (web\index.html), hosted by dAHKxodus.ahk.
// args: [0] the AHK script's window (bridge calls go there as WM_COPYDATA), [1] the program folder,
//       [2] a .docx to open (or ""), [3] measure mode: where to save the edited copy (or "").
// The window runs on its own thread so that this script returns and AutoHotkey stays free to answer bridge calls;
// when the window closes, the AHK script is told to exit (WM_APP + 1).
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

var ahk = (IntPtr)long.Parse(args[0]);
var dir = args[1];
var start = args.Length > 2 ? args[2] : "";
var measureOut = args.Length > 3 ? args[3] : "";
MainForm.LogPath = Path.Combine(dir, "logs", "dahkxodus.log");
MainForm.Log("---- start " + (start == "" ? "(no document)" : Path.GetFileName(start)) + (measureOut != "" ? " [measure]" : ""));

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
    MainForm.PostMessageW(ahk, 0x8001, IntPtr.Zero, IntPtr.Zero);
});
ui.SetApartmentState(ApartmentState.STA);
ui.Start();
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
    const long CallTag = 0x44415831, ReplyOk = 0x44415832, ReplyError = 0x44415833;    // "DAX1", "DAX2", "DAX3"

    [StructLayout(LayoutKind.Sequential)]
    struct COPYDATASTRUCT { public IntPtr dwData; public int cbData; public IntPtr lpData; }

    [DllImport("user32.dll")] public static extern bool PostMessageW(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")]
    static extern IntPtr SendMessageTimeoutW(IntPtr hWnd, int msg, IntPtr wParam, ref COPYDATASTRUCT lParam,
        int flags, int timeout, out IntPtr result);

    readonly IntPtr ahk;
    readonly string dir, measureOut;
    string current;
    readonly WebView2 web = new() { Dock = DockStyle.Fill };
    readonly EngineHost Engine = new();
    readonly Stopwatch clock = Stopwatch.StartNew();
    string? reply;
    bool replyIsError;

    public MainForm(IntPtr ahk, string dir, string start, string measureOut) {
        this.ahk = ahk; this.dir = dir; current = start; this.measureOut = measureOut;
        Text = "dAHKxodus" + (start == "" ? "" : " - " + Path.GetFileName(start));
        Size = new System.Drawing.Size(1400, 950);
        StartPosition = FormStartPosition.CenterScreen;
        Controls.Add(web);
        if (measureOut != "") {    // a measure run that goes wrong still closes
            var watchdog = new System.Windows.Forms.Timer { Interval = 120000 };
            watchdog.Tick += (s, e) => { Log("measure: gave up after 120 s"); Close(); };
            watchdog.Start();
        }
        Load += async (s, e) => {
            try { await Init(); } catch (Exception ex) { Log("init failed: " + ex); if (measureOut != "") Close(); }
        };
    }

    async Task Init() {
        var data = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "dAHKxodus", "WebView2");
        var env = await CoreWebView2Environment.CreateAsync(null, data);
        await web.EnsureCoreWebView2Async(env);
        Log($"dAHKxodus {Version}, webview2 ready {clock.ElapsedMilliseconds} ms (runtime {env.BrowserVersionString})");
        DaxExtensions.Started(this);
        var cw = web.CoreWebView2;
        // (.example, not .local: a .local host name cost a fixed 2 s before the page's first script ran)
        cw.SetVirtualHostNameToFolderMapping("app.example", Path.Combine(dir, "web"), CoreWebView2HostResourceAccessKind.Allow);
        cw.WebMessageReceived += OnMessage;
        // The editing session runs here, in .NET, not in the page's WASM runtime: the page calls the engine through
        // this object (chrome.webview.hostObjects.sync.engine). DAHKXODUS_ENGINE=wasm keeps it in the page, to compare.
        cw.AddHostObjectToScript("engine", Engine);
        var query = new List<string>();
        var window = Environment.GetEnvironmentVariable("DAHKXODUS_WINDOW");    // (paragraphs per mount window, for tests)
        if (!string.IsNullOrEmpty(window)) query.Add("window=" + Uri.EscapeDataString(window));
        var engine = Environment.GetEnvironmentVariable("DAHKXODUS_ENGINE");
        if (!string.IsNullOrEmpty(engine)) query.Add("engine=" + Uri.EscapeDataString(engine));
        if (Environment.GetEnvironmentVariable("DAHKXODUS_VERIFY") == "1") query.Add("verify=1");    // (both engines, compared)
        var (user, initials, from) = UserIdentity.Read(dir);
        Log($"user \"{user}\"{(initials == "" ? "" : $" ({initials})")} from {from}");
        query.Add("user=" + Uri.EscapeDataString(user));
        if (initials != "") query.Add("initials=" + Uri.EscapeDataString(initials));
        cw.Navigate("https://app.example/index.html" + (query.Count == 0 ? "" : "?" + string.Join("&", query)));
    }

    void Send(object message) => web.CoreWebView2.PostWebMessageAsJson(JsonSerializer.Serialize(message));

    void OpenFile(string path) {
        var bytes = File.ReadAllBytes(path);
        current = path;
        Text = "dAHKxodus - " + Path.GetFileName(path);
        Log($"open {Path.GetFileName(path)} ({bytes.Length:N0} bytes) at {clock.ElapsedMilliseconds} ms");
        Send(new { type = "open", name = Path.GetFileName(path), b64 = Convert.ToBase64String(bytes) });
    }

    void OnMessage(object? sender, CoreWebView2WebMessageReceivedEventArgs e) {
        try {
            using var doc = JsonDocument.Parse(e.WebMessageAsJson);
            var m = doc.RootElement;
            string Str(string name) => m.TryGetProperty(name, out var v) ? v.ToString() : "";
            switch (Str("type")) {
                case "ready":
                    Log($"editor ready {clock.ElapsedMilliseconds} ms ({Str("engine")} engine; page: {Str("sinceNavigation")} ms since navigation, scripts {Str("scripts")} ms, boot {Str("boot")} ms)");
                    if (current != "") OpenFile(current);
                    else Send(new { type = "blank" });
                    break;
                case "opened":
                    Log($"opened {clock.ElapsedMilliseconds} ms (window {Environment.GetEnvironmentVariable("DAHKXODUS_WINDOW") ?? "default"}; open {Str("open")} ms: first window after {Str("firstWindow")} ms, longest frozen {Str("longestGap")} ms; first paint {Str("paint")} ms, {Str("blocks")} blocks)");
                    if (measureOut != "") Send(new { type = "measure" });
                    break;
                case "measured":
                    var bytes = Convert.FromBase64String(Str("b64"));
                    File.WriteAllBytes(measureOut, bytes);
                    var ahkInfo = m.GetProperty("ahk");
                    Log($"measured: AHK returned \"{ahkInfo.GetProperty("text")}\" in {ahkInfo.GetProperty("roundTrip")} ms; save {Str("save")} ms, {bytes.Length:N0} bytes");
                    Close();
                    break;
                case "saved":
                    SaveBytes(Convert.FromBase64String(Str("b64")), Str("name"));
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
                case "engineStats":    // time spent in the native engine itself (the page's own timing adds the crossing)
                    Log("host: " + Engine.TakeStats());
                    break;
            }
        } catch (Exception ex) {
            Log("message failed: " + ex);
        }
    }

    void SaveBytes(byte[] bytes, string name) {
        using var dlg = new SaveFileDialog {
            Filter = "Word document (*.docx)|*.docx", FileName = name,
            InitialDirectory = current != "" ? Path.GetDirectoryName(current) : "",
        };
        if (dlg.ShowDialog(this) != DialogResult.OK) return;
        File.WriteAllBytes(dlg.FileName, bytes);
        current = dlg.FileName;
        Text = "dAHKxodus - " + Path.GetFileName(current);
        Log($"saved {Path.GetFileName(current)} ({bytes.Length:N0} bytes)");
    }

    // Calls Bridge.<fn>(arg) in dAHKxodus.ahk. AutoHotkey answers with WM_COPYDATA to this window while we wait
    // (a thread waiting in SendMessage still takes messages sent to it), so the answer is there when SendMessage returns.
    string CallAhk(string fn, string arg) {
        var text = fn + "\t" + arg;
        var ptr = Marshal.StringToHGlobalUni(text);
        try {
            reply = null;
            var cds = new COPYDATASTRUCT { dwData = (IntPtr)CallTag, cbData = (text.Length + 1) * 2, lpData = ptr };
            var sent = SendMessageTimeoutW(ahk, WM_COPYDATA, Handle, ref cds, 0, 10000, out _);
            if (sent == IntPtr.Zero || reply == null) { replyIsError = true; return "AutoHotkey did not answer"; }
            return reply;
        } finally {
            Marshal.FreeHGlobal(ptr);
        }
    }

    protected override void WndProc(ref Message m) {
        if (m.Msg == WM_COPYDATA) {
            var cds = Marshal.PtrToStructure<COPYDATASTRUCT>(m.LParam);
            var tag = (long)cds.dwData;
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
// settings.ini beside the program ([User] Name=, Initials=), then Office's own user information (what Word stamps on
// the same machine), then the Windows account name. Initials are "" when nothing names them: the editor then takes
// them from the name, as Word does.
static class UserIdentity {
    public static (string Name, string Initials, string From) Read(string dir) {
        var ini = Path.Combine(dir, "settings.ini");
        if (File.Exists(ini)) {
            var name = IniValue(ini, "User", "Name");
            if (name != "") return (name, IniValue(ini, "User", "Initials"), "settings.ini");
        }
        try {
            using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\Common\UserInfo");
            var name = (key?.GetValue("UserName") as string)?.Trim() ?? "";
            if (name != "") return (name, (key?.GetValue("UserInitials") as string)?.Trim() ?? "", "Office user information");
        } catch { }
        return (Environment.UserName, "", "the Windows account");
    }

    static string IniValue(string path, string section, string key) {
        var inSection = false;
        foreach (var raw in File.ReadLines(path)) {
            var line = raw.Trim();
            if (line.StartsWith('[') && line.EndsWith(']')) { inSection = line[1..^1].Trim().Equals(section, StringComparison.OrdinalIgnoreCase); continue; }
            if (!inSection || line.StartsWith(';')) continue;
            var eq = line.IndexOf('=');
            if (eq > 0 && line[..eq].Trim().Equals(key, StringComparison.OrdinalIgnoreCase)) return line[(eq + 1)..].Trim();
        }
        return "";
    }
}
