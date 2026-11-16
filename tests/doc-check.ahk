#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
#Include "..\Doc.ahk"
; The measure run's check of Doc.ahk (tests\run-measure.ps1): dAHKxodus starts this script beside its window, as it
; starts a macro, near the end of the run. It drives the editor through Doc as a macro would, and reports what it found
; to the log (Doc.Exec("measureDone")), "doc api: ..." lines.

results := [], failed := false
Check(name, ok, detail := "") {
    global failed
    results.Push(name " " (ok ? "ok" : "FAILED") (detail != "" ? " (" detail ")" : ""))
    if !ok
        failed := true
}
ErrorCode(fn) {
    try {
        fn()
        return "no error"
    } catch DocError as e
        return e.Code
    catch as e
        return "not a DocError: " e.Message
}

events := Map()
Seen(e) {
    events[e.Name] := events.Get(e.Name, 0) + 1
    if e.Name = "Command"
        events["Command:" e.Command] := 1
    if e.Name = "Modified"
        events["Modified:" (e.Modified ? 1 : 0)] := 1
}

try {
    Check("running", Doc.IsRunning)
    Check("version", Doc.AppVersion ~= "^\d+\.\d+\.\d+\.\d+$", Doc.AppVersion)
    names := Map()
    for c in Doc.Commands()
        names[c.Name] := c
    Check("commands", names.Has("bold") && names.Has("insertText") && names.Has("save") && names["find"].Args.Length = 3,
        names.Count " commands")
    Check("help", InStr(Doc.Help(), "InsertText(text)"))
    for name in ["Opened", "Saved", "Modified", "TextChanged", "Command"]
        Doc.On(name, Seen)

    Check("new", Doc.New() = true && Doc.Path = "" && Doc.Text = "")
    Doc.InsertText("Hello from Doc.ahk`nSecond line, ñ € 😀")
    Check("insertText and text", Doc.Text == "Hello from Doc.ahk`nSecond line, ñ € 😀", Doc.Text)
    Check("find", Doc.Find("Doc.ahk") = 1 && Doc.SelectedText == "Doc.ahk", Doc.SelectedText)
    Doc.Bold(true)
    Check("bold", Doc.FormatState().bold = true && Doc.SelectedText == "Doc.ahk")
    Doc.Undo()
    Check("undo", Doc.FormatState().bold = false)
    Check("find by object", Doc.Find({text: "second", matchCase: true}) = 0 && Doc.Find({text: "Second"}) = 1)
    Check("replaceAll", Doc.ReplaceAll("Second", "Next") = 1 && InStr(Doc.Text, "Next line"))
    Check("unknown command (here)", ErrorCode(() => Doc.NoSuchCommand()) = "unknown_command")
    Check("unknown command (editor)", ErrorCode(() => Doc.Exec("noSuchCommand")) = "unknown_command")
    Check("bad arguments", ErrorCode(() => Doc.FontSize("big")) = "bad_arguments")
    Check("query as a command refused", ErrorCode(() => Doc.Bold) = "error")
    Check("modified", Doc.Modified = true)

    t := A_TickCount
    Loop 20
        Doc.WordCount
    perCall := Round((A_TickCount - t) / 20, 1)
    Check("call time", perCall < 200, perCall " ms a call")

    path := A_ScriptDir "\..\logs\measure-doc.docx"
    try FileDelete(path)
    Check("save", Doc.Save(path) = true && FileExist(path) && Doc.Path = Doc_FullPath(path) && Doc.Modified = false, Doc.Path)
    Check("open", Doc.Open(path) = true && InStr(Doc.Text, "Hello from Doc.ahk") = 1 && Doc.Name = "measure-doc.docx", Doc.Name)
    pdf := A_ScriptDir "\..\logs\measure-doc.pdf"
    try FileDelete(pdf)
    Check("exportPdf", Doc.ExportPdf(pdf) = true && FileExist(pdf))
    for ext in [".html", ".md", ".txt"] {
        copy := A_ScriptDir "\..\logs\measure-doc" ext
        try FileDelete(copy)
        Check("saveCopy " ext, Doc.SaveCopy(copy) = true && InStr(FileRead(copy, "UTF-8"), "Hello from Doc.ahk")
            && Doc.Path = Doc_FullPath(path), copy)
    }

    Sleep(600)    ; (events come as they happen: TextChanged after a pause)
    Check("events", events.Has("Opened") && events.Has("Saved") && events.Has("TextChanged")
        && events.Has("Command:bold") && events.Has("Modified:1") && events.Has("Modified:0"))
} catch as e {
    Check("script", false, Type(e) ": " e.Message " (" e.What ")")
}

text := ""
for r in results
    text .= (A_Index > 1 ? "; " : "") r
for k in events
    text .= (A_Index = 1 ? "; events: " : ", ") k
try Doc.Exec("measureDone", {ok: !failed, text: text})
ExitApp

Doc_FullPath(path) {
    buf := Buffer(2048)
    DllCall("GetFullPathNameW", "str", path, "uint", 1024, "ptr", buf, "ptr", 0)
    return StrGet(buf)
}
