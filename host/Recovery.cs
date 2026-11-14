// Recovery copies of the document being edited (as AhkWord's Recovery). Loaded by dAHKxodus.cs (#load).
//
// While the document has changes not saved, the host writes a copy of it (the .docx the page would save) every few
// minutes ([General] AutosaveMinutes, 0: never) and when the window loses focus. The copy is removed when the document
// is saved, replaced (New, Open) or closed as the user chose, so a copy still there at the next start is one whose
// dAHKxodus ended without closing (a crash, a power cut, Windows restarting): that start offers it. Each dAHKxodus
// window has copies of its own, named by its process, so two running at once never take each other's for left behind.
// The document's own file is never written by autosave.
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;

public sealed class Recovery {
    /// What a copy is: the document's file ("" when it was never saved), the name the title showed, when the copy
    /// was written, and the process that wrote it (its id and start time: ids are reused).
    public sealed record Info(string FileName, string Name, DateTime Saved, int ProcessId, long ProcessStart);

    /// A copy left behind: what it was, and its files.
    public sealed record Orphan(Info Info, string DocxPath, string InfoPath);

    readonly string id;
    readonly int processId;
    readonly long processStart;
    byte[]? writtenHash;    // the copy written last (null: there is none now)

    public Recovery(string folder) {
        Folder = folder;
        using var process = Process.GetCurrentProcess();
        processId = process.Id;
        processStart = process.StartTime.ToUniversalTime().Ticks;
        id = $"{processId}-{processStart}";
    }

    /// %LOCALAPPDATA%\dAHKxodus\Recovery (portable: data\Recovery beside the program).
    public static string DefaultFolder => Path.Combine(AppFolders.Local, "Recovery");

    public string Folder { get; }
    public string DocxPath => Path.Combine(Folder, id + ".docx");
    string InfoPath => Path.Combine(Folder, id + ".json");

    /// Whether this window has a copy now.
    public bool HasCopy => writtenHash != null;

    /// Writes the copy (unless it is the one written last); fileName: the document's file ("" if never saved).
    /// Returns whether it wrote. Throws IOException / UnauthorizedAccessException when the folder can't be written.
    public bool Write(byte[] docx, string fileName, string name) {
        var hash = SHA256.HashData(docx);
        if (writtenHash != null && hash.AsSpan().SequenceEqual(writtenHash) && File.Exists(DocxPath)) return false;
        Directory.CreateDirectory(Folder);
        // (written beside, then moved over: a crash while writing leaves the copy before it whole)
        var temporary = DocxPath + ".tmp";
        File.WriteAllBytes(temporary, docx);
        File.Move(temporary, DocxPath, overwrite: true);
        File.WriteAllText(InfoPath, JsonSerializer.Serialize(new Info(fileName, name, DateTime.Now, processId, processStart)));
        writtenHash = hash;
        return true;
    }

    /// Removes this window's copy (the document was saved, replaced or closed as the user chose).
    public void Clear() {
        if (writtenHash == null && !File.Exists(InfoPath) && !File.Exists(DocxPath)) return;
        Delete(DocxPath, InfoPath, DocxPath + ".tmp");
        writtenHash = null;
    }

    static void Delete(params string[] paths) {
        foreach (var path in paths)
            try { File.Delete(path); } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
    }

    /// The copies in the folder whose dAHKxodus isn't running any more, the newest first.
    public static List<Orphan> Orphans(string folder) {
        var orphans = new List<Orphan>();
        if (!Directory.Exists(folder)) return orphans;
        foreach (var infoPath in Directory.GetFiles(folder, "*.json")) {
            var docxPath = Path.ChangeExtension(infoPath, ".docx");
            Info? info;
            try { info = JsonSerializer.Deserialize<Info>(File.ReadAllText(infoPath)); }
            catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException) { continue; }
            if (info == null || !File.Exists(docxPath) || Running(info)) continue;
            orphans.Add(new Orphan(info, docxPath, infoPath));
        }
        return orphans.OrderByDescending(o => o.Info.Saved).ToList();
    }

    // (a process with the copy's id and start time: its dAHKxodus is still running)
    static bool Running(Info info) {
        try {
            using var process = Process.GetProcessById(info.ProcessId);
            return process.StartTime.ToUniversalTime().Ticks == info.ProcessStart;
        } catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or System.ComponentModel.Win32Exception) {
            return false;
        }
    }

    /// The .docx of a copy left behind.
    public static byte[] Read(Orphan orphan) => File.ReadAllBytes(orphan.DocxPath);

    /// Removes a copy left behind (recovered or discarded).
    public static void Discard(Orphan orphan) => Delete(orphan.DocxPath, orphan.InfoPath);
}
