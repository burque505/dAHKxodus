// Headless checks of the Docxodus engine, through the same entry point the editor page uses in dAHKxodus.
// A check marked Known is a fault already in the plan (PLAN.md, section 2a): it is reported but does not fail the run,
// and it is reported loudly once it passes, so the mark gets removed.
using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;
using DocxodusNative;

var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
var sample = args.Length > 0 ? args[0] : Path.Combine(root, "samples", "verify-mix.docx");
XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
int failed = 0, known = 0, fixedKnown = 0;

string C(string method, params object?[] a) =>
    NativeSessionBridge.Call("DocxSessionBridge." + method, JsonSerializer.Serialize(a)) ?? "";

void Check(string name, bool ok, string detail = "", string? knownIssue = null) {
    if (knownIssue != null) {
        if (ok) { fixedKnown++; Console.WriteLine($"FIXED  {name} (remove the known mark: {knownIssue})"); }
        else { known++; Console.WriteLine($"known  {name}: {detail} [{knownIssue}]"); }
        return;
    }
    if (!ok) failed++;
    Console.WriteLine($"{(ok ? "ok    " : "FAIL  ")} {name}{(ok || detail == "" ? "" : ": " + detail)}");
}

int Open(byte[] bytes) => int.Parse(C("OpenSession", Convert.ToBase64String(bytes), ""));
byte[] Save(int h) => Convert.FromBase64String(C("Save", h));
ZipArchive Zip(byte[] b) => new(new MemoryStream(b));
string Entry(ZipArchive z, string name) {
    var e = z.GetEntry(name) ?? throw new Exception(name + " missing");
    using var r = new StreamReader(e.Open());
    return r.ReadToEnd();
}

// The first body paragraph's anchor, from the render plan.
string FirstParagraph(int h) {
    var plan = JsonDocument.Parse(C("ListBlocks", h)).RootElement;
    foreach (var b in plan.GetProperty("body").EnumerateArray()) {
        var id = b.GetProperty("id").GetString();
        if (id != null && id.StartsWith("p:body:")) return id;
    }
    throw new Exception("no body paragraph in ListBlocks: " + plan.GetRawText()[..Math.Min(300, plan.GetRawText().Length)]);
}

Console.WriteLine("sample: " + sample);
var bytes = File.ReadAllBytes(sample);

// 1. Open, save, reopen: the package survives a round trip with nothing changed.
{
    var h = Open(bytes);
    var saved = Save(h);
    C("CloseSession", h);
    var h2 = Open(saved);
    Check("open, save, reopen", saved.Length > 0 && h2 > 0);
    using var a = Zip(bytes); using var b = Zip(saved);
    var lost = a.Entries.Select(e => e.FullName).Except(b.Entries.Select(e => e.FullName)).ToList();
    Check("no part lost on save", lost.Count == 0, string.Join(", ", lost));
    // (Save writes the XML out again, e.g. "utf-8" in the declaration, so compare what the document says.)
    string Text(string xml) => string.Join("\n", XDocument.Parse(xml).Descendants(w + "p").Select(p => string.Concat(p.Descendants(w + "t").Select(t => t.Value))));
    Check("document text unchanged on save", Text(Entry(a, "word/document.xml")) == Text(Entry(b, "word/document.xml")));
    C("CloseSession", h2);
}

// 2. An edit, then Undo, gives the text back; Redo applies it again.
{
    var h = Open(bytes);
    var p = FirstParagraph(h);
    C("ReplaceText", h, p, "Replaced by the engine checks");
    var edited = Entry(Zip(Save(h)), "word/document.xml");
    Check("edit lands", edited.Contains("Replaced by the engine checks"));
    var undone = C("Undo", h) == "true";
    var afterUndo = Entry(Zip(Save(h)), "word/document.xml");
    Check("Undo restores the paragraph", undone && !afterUndo.Contains("Replaced by the engine checks"));
    C("Redo", h);
    Check("Redo applies it again", Entry(Zip(Save(h)), "word/document.xml").Contains("Replaced by the engine checks"));
    C("CloseSession", h);
}

// 3. A comment given an author, initials and a date keeps all three (finding 4 is the editor not passing them).
{
    var h = Open(bytes);
    var p = FirstParagraph(h);
    C("AddComment", h, p, "", "Engine Check", "EC", "2026-11-13T12:00:00Z", "A comment from the engine checks");
    var comments = XDocument.Parse(Entry(Zip(Save(h)), "word/comments.xml"));
    var c = comments.Descendants(w + "comment").FirstOrDefault(e => e.Value.Contains("A comment from the engine checks"));
    Check("comment saved", c != null);
    Check("comment keeps author, initials and date",
        c != null && (string?)c.Attribute(w + "author") == "Engine Check" && (string?)c.Attribute(w + "initials") == "EC"
        && ((string?)c.Attribute(w + "date"))?.StartsWith("2026-11-13") == true,
        c?.ToString().Split('>')[0] ?? "");
    C("CloseSession", h);
}

// 4. An inserted picture is stored as Word stores it: in word/media/, never at the package root (finding 5). (Its
//    relationship target is absolute, "/word/media/imageN.png", as the SDK writes every new relationship; Word reads it.)
{
    var h = Open(bytes);
    var p = FirstParagraph(h);
    const string png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
    var result = C("InsertImage", h, p, 0, png, "{}");
    using var z = Zip(Save(h));
    var before = Zip(bytes).Entries.Select(e => e.FullName).ToHashSet();
    var added = z.Entries.Select(e => e.FullName).Where(n => !before.Contains(n) && n.EndsWith(".png")).ToList();
    Check("picture inserted", added.Count == 1, result.Length > 200 ? result[..200] : result);
    var rels = Entry(z, "word/_rels/document.xml.rels");
    Check("picture stored in word/media, nothing pointing at /media",
        added.Count == 1 && added[0].StartsWith("word/media/") && !rels.Contains("Target=\"/media/"),
        "stored at " + string.Join(", ", added));
    C("CloseSession", h);
}

Console.WriteLine($"\n{(failed == 0 ? "PASSED" : "FAILED")}: {failed} failed, {known} known, {fixedKnown} known now fixed");
return failed == 0 ? 0 : 1;
