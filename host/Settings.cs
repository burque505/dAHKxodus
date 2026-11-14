// Where dAHKxodus keeps its files (AppFolders), its settings.ini (Settings) and the recent documents (RecentFiles).
// Loaded by dAHKxodus.cs (#load).
using System.Text;

// dAHKxodus's own files: settings.ini (and the recent documents in it) in %APPDATA%\dAHKxodus, this computer's state
// (WebView2's profile, recovery copies) in %LOCALAPPDATA%\dAHKxodus. In portable mode, a file named "portable" or
// "portable.txt" beside the program, all of them go in a "data" folder beside it instead, so the folder can be carried
// to another computer with them. (As AhkWord's AppFolders.)
public static class AppFolders {
    /// The program's folder (dAHKxodus.ahk's): set before anything else is read.
    public static string ProgramFolder { get; set; } = AppContext.BaseDirectory;

    /// dAHKxodus.exe when the program runs as its own exe (set by vs\VsHost.cs); null when dAHKxodus.ahk runs it
    /// through Csharp-in-AHK.
    public static string? ProgramExe { get; set; }

    /// (The exe) Turns its command line, [file.docx] [/measure out.docx], into the arguments dAHKxodus.ahk passes.
    public static Func<string[], string[]>? CompleteArgs { get; set; }

    /// (The exe) Starts dAHKxodus.ahk beside it as the bridge's AutoHotkey side, given the window to report to.
    public static Action<IntPtr>? StartCompanion { get; set; }

    public static bool Portable =>
        File.Exists(Path.Combine(ProgramFolder, "portable")) || File.Exists(Path.Combine(ProgramFolder, "portable.txt"));

    /// The settings' folder (roaming).
    public static string Roaming => Portable ? Path.Combine(ProgramFolder, "data")
        : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "dAHKxodus");

    /// This computer's state: WebView2's profile, recovery copies.
    public static string Local => Portable ? Path.Combine(ProgramFolder, "data")
        : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "dAHKxodus");
}

// settings.ini in AppFolders.Roaming: sections and key=value lines, UTF-8, comments (;) kept. Read once, written back
// whole (through a temporary file) on each change. Earlier versions kept it beside the program: Init moves it.
public static class Settings {
    static List<string>? lines;

    public static string FilePath => Path.Combine(AppFolders.Roaming, "settings.ini");

    /// Moves a settings.ini left beside the program by an earlier version into the settings folder, once.
    public static void Init() {
        var old = Path.Combine(AppFolders.ProgramFolder, "settings.ini");
        try {
            Directory.CreateDirectory(AppFolders.Roaming);
            if (File.Exists(old) && !File.Exists(FilePath)
                && !string.Equals(Path.GetFullPath(old), Path.GetFullPath(FilePath), StringComparison.OrdinalIgnoreCase)) {
                File.Copy(old, FilePath);
                try { File.Delete(old); } catch (IOException) { } catch (UnauthorizedAccessException) { }
                MainForm.Log($"settings: moved {old} to {FilePath}");
            }
        } catch (Exception ex) {
            MainForm.Log("settings: " + ex.Message);
        }
    }

    static List<string> Lines {
        get {
            if (lines != null) return lines;
            try { lines = File.Exists(FilePath) ? File.ReadAllLines(FilePath, Encoding.UTF8).ToList() : new(); }
            catch (Exception ex) { MainForm.Log("settings: " + ex.Message); lines = new(); }
            return lines;
        }
    }

    static bool IsSection(string line, out string name) {
        var t = line.Trim();
        name = t.Length > 1 && t[0] == '[' && t[^1] == ']' ? t[1..^1].Trim() : "";
        return name != "";
    }

    // The lines [start, end) of a section's body (start = -1: no such section).
    static (int start, int end) Find(string section) {
        var all = Lines;
        for (int i = 0; i < all.Count; i++) {
            if (!IsSection(all[i], out var name) || !name.Equals(section, StringComparison.OrdinalIgnoreCase)) continue;
            int end = i + 1;
            while (end < all.Count && !IsSection(all[end], out _)) end++;
            return (i + 1, end);
        }
        return (-1, -1);
    }

    static bool KeyOf(string line, out string key, out string value) {
        var t = line.Trim();
        var eq = t.IndexOf('=');
        key = eq > 0 && !t.StartsWith(';') ? t[..eq].Trim() : "";
        value = key == "" ? "" : t[(eq + 1)..].Trim();
        return key != "";
    }

    public static string Get(string section, string key, string fallback = "") {
        var (start, end) = Find(section);
        for (int i = start; start >= 0 && i < end; i++)
            if (KeyOf(Lines[i], out var k, out var v) && k.Equals(key, StringComparison.OrdinalIgnoreCase)) return v;
        return fallback;
    }

    public static int GetInt(string section, string key, int fallback) =>
        int.TryParse(Get(section, key), out var n) ? n : fallback;

    public static void Set(string section, string key, string value) {
        var all = Lines;
        var (start, end) = Find(section);
        if (start < 0) {
            if (all.Count > 0 && all[^1].Trim() != "") all.Add("");
            all.Add($"[{section}]");
            all.Add($"{key}={value}");
        } else {
            int at = -1;
            for (int i = start; i < end; i++)
                if (KeyOf(all[i], out var k, out _) && k.Equals(key, StringComparison.OrdinalIgnoreCase)) { at = i; break; }
            if (at >= 0) all[at] = $"{key}={value}";
            else {
                int insert = end;
                while (insert > start && all[insert - 1].Trim() == "") insert--;    // before the blank lines ending it
                all.Insert(insert, $"{key}={value}");
            }
        }
        Save();
    }

    /// The values of a section's keys, in file order.
    public static List<string> Values(string section) {
        var (start, end) = Find(section);
        var result = new List<string>();
        for (int i = start; start >= 0 && i < end; i++)
            if (KeyOf(Lines[i], out _, out var v) && v != "") result.Add(v);
        return result;
    }

    /// Replaces a section's keys with prefix1=..., prefix2=..., keeping its place in the file.
    public static void SetValues(string section, string prefix, IEnumerable<string> values) {
        var all = Lines;
        var body = values.Select((v, i) => $"{prefix}{i + 1}={v}").ToList();
        var (start, end) = Find(section);
        if (start < 0) {
            if (all.Count > 0 && all[^1].Trim() != "") all.Add("");
            all.Add($"[{section}]");
            all.AddRange(body);
        } else {
            int last = end;
            while (last > start && all[last - 1].Trim() == "") last--;
            all.RemoveRange(start, last - start);
            all.InsertRange(start, body);
        }
        Save();
    }

    static void Save() {
        try {
            Directory.CreateDirectory(AppFolders.Roaming);
            var tmp = FilePath + ".tmp";
            File.WriteAllLines(tmp, Lines, new UTF8Encoding(false));
            File.Move(tmp, FilePath, overwrite: true);
        } catch (Exception ex) {
            MainForm.Log("settings: could not write " + FilePath + ": " + ex.Message);
        }
    }
}

// The documents opened or saved lately, most recent first: settings.ini's [Recent] File1=...
public static class RecentFiles {
    const int Max = 10;

    public static List<string> List() => Settings.Values("Recent");

    public static void Add(string path) {
        path = Path.GetFullPath(path);
        var list = List();
        list.RemoveAll(p => p.Equals(path, StringComparison.OrdinalIgnoreCase));
        list.Insert(0, path);
        Settings.SetValues("Recent", "File", list.Take(Max));
    }

    public static void Remove(string path) {
        var list = List();
        if (list.RemoveAll(p => p.Equals(path, StringComparison.OrdinalIgnoreCase)) > 0) Settings.SetValues("Recent", "File", list);
    }
}
