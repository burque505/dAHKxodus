// The host's own text (its dialogs and messages, and the few labels the page adds to the ribbon) in the user's
// language: lang\<code>.json beside the program, flat "key": "text" pairs, {0}-style placeholders. A key missing from
// a language falls back to English, then to the key itself. Loaded by dAHKxodus.cs (#load).
// The ribbon's own labels come from Docxodus and are English for now.
using System.Globalization;
using System.Text.Json;

public static class Lang {
    static Dictionary<string, string> strings = new(), english = new();

    /// The language in use ("en", "es", ...).
    public static string Code { get; private set; } = "en";

    static string Folder => Path.Combine(AppFolders.ProgramFolder, "lang");

    /// The languages there are files for: (code, its own name), English first.
    public static List<(string Code, string Name)> Available() {
        var list = new List<(string, string)>();
        try {
            foreach (var file in Directory.GetFiles(Folder, "*.json").OrderBy(f => Path.GetFileNameWithoutExtension(f) == "en" ? "" : f)) {
                var code = Path.GetFileNameWithoutExtension(file);
                if (code.Contains('.')) continue;    // (es.ribbon.json: the ribbon's own texts, not a language)
                list.Add((code, Read(code).GetValueOrDefault("language.name") ?? code));
            }
        } catch (Exception ex) { MainForm.Log("lang: " + ex.Message); }
        return list;
    }

    /// settings.ini's [General] Language= ("" or missing: Windows's display language, when there is a file for it).
    public static void Load() {
        // (DAHKXODUS_LANGUAGE wins: a test run in another language, the user's settings untouched)
        var setting = Environment.GetEnvironmentVariable("DAHKXODUS_LANGUAGE") is { Length: > 0 } forced ? forced
            : Settings.Get("General", "Language");
        var windows = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName;
        Code = setting != "" && File.Exists(Path.Combine(Folder, setting + ".json")) ? setting
            : File.Exists(Path.Combine(Folder, windows + ".json")) ? windows : "en";
        english = Read("en");
        strings = Code == "en" ? english : Read(Code);
    }

    static Dictionary<string, string> Read(string code) {
        try {
            return JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(Path.Combine(Folder, code + ".json"))) ?? new();
        } catch (Exception ex) {
            MainForm.Log($"lang {code}: {ex.Message}");
            return new();
        }
    }

    /// The ribbon's own text in this language: lang\<code>.ribbon.json, keyed by the ribbon's English ({} for English,
    /// or when there is no such file: the ribbon then stays English).
    public static Dictionary<string, string> Ribbon() {
        var path = Path.Combine(Folder, Code + ".ribbon.json");
        if (Code == "en" || !File.Exists(path)) return new();
        try {
            return JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(path)) ?? new();
        } catch (Exception ex) {
            MainForm.Log($"lang {Code} ribbon: {ex.Message}");
            return new();
        }
    }

    public static string T(string key, params object[] args) {
        var text = strings.GetValueOrDefault(key) ?? english.GetValueOrDefault(key) ?? key;
        if (args.Length == 0) return text;
        try { return string.Format(text, args); } catch (FormatException) { return text; }
    }
}
