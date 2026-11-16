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

    ; Paste: HTML (a list), plain text (one undo step), the clipboard as Ctrl+V reads it, and a Word picture (a file in
    ; the temporary folder, which the host reads for the page). The text pasted here is undone again.
    Doc.MoveCaret("documentEnd")
    before := Doc.Text
    Check("pasteHtml", Doc.PasteHtml("<p>Pasted</p><ol><li>one</li><li>two</li></ol>") = true
        && InStr(Doc.Text, "Pasted`none`ntwo"), Doc.Text)
    Check("pasteText", Doc.PasteText("p1`tx`np2") = true && InStr(Doc.Text, "twop1`tx`np2"), Doc.Text)
    Doc.Undo()
    Check("paste is one undo step", !InStr(Doc.Text, "p1") && InStr(Doc.Text, "Pasted"), Doc.Text)
    saved := ClipboardAll()
    A_Clipboard := "from the clipboard"
    try ok := Doc.Paste("text")
    catch as e
        ok := e.Message
    A_Clipboard := saved
    Check("paste (clipboard)", ok = true && InStr(Doc.Text, "from the clipboard"), ok " | " Doc.Text)
    clipDir := A_Temp "\dax-measure-clip"
    DirCreate(clipDir)
    WritePng(clipDir "\clip_image001.png")
    pictureUrl := "file:///" StrReplace(clipDir "\clip_image001.png", "\", "/")
    pasted := Doc.PasteHtml('<p>pic <img src="' pictureUrl '" width="1" height="1"/></p>')
    Doc.SaveCopy(clipDir "\pasted.html")    ; (a web page copy carries the document's pictures)
    Check("paste picture (Word's clip file)", pasted = true && InStr(FileRead(clipDir "\pasted.html", "UTF-8"), "<img"), pasted)
    try DirDelete(clipDir, true)
    Loop 3    ; (the clipboard's paste, the picture's, the HTML's: the text's is undone already)
        Doc.Undo()
    Check("pastes undone", Doc.Text == before, Doc.Text)
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

    ; a second window: Doc.NewWindow opens it and answers with it; Doc.SetView and Doc.View there; then it closes
    first := Doc.Window || EnvGet("DAHKXODUS_HWND") + 0
    second := Doc.NewWindow()
    Check("newWindow", second && second != first && WinExist("ahk_id " second) && WinGetPID("ahk_id " second) = WinGetPID("ahk_id " first), second)
    Doc.Window := second
    Doc.SetView(true, "width")
    Sleep(500)
    v := Doc.View
    Check("setView", v.paginated = true && v.zoom = "width", v.paginated " " v.zoom)
    Doc.Close()
    Doc.Window := first
    Sleep(1000)
    Check("closeWindow", !WinExist("ahk_id " second) && WinExist("ahk_id " first) && Doc.Name = "measure-doc.docx")

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

; A 1x1 PNG, as Word leaves a copied picture in the temporary folder.
WritePng(path) {
    b64 := "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="
    size := 0
    DllCall("crypt32\CryptStringToBinaryW", "str", b64, "uint", 0, "uint", 1, "ptr", 0, "uint*", &size, "ptr", 0, "ptr", 0)
    buf := Buffer(size)
    DllCall("crypt32\CryptStringToBinaryW", "str", b64, "uint", 0, "uint", 1, "ptr", buf, "uint*", &size, "ptr", 0, "ptr", 0)
    FileOpen(path, "w").RawWrite(buf, size)
}
