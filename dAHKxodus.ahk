#Requires AutoHotkey v2.0 64-bit
#SingleInstance Off
#Include "Csharp-in-AHK.ahk"
#Include "Doc.ahk"
; dAHKxodus: Docxodus's browser editor (web\index.html, npm package docxodus) in a WebView2 window
; (dAHKxodus.cs) hosted by this script. The editor calls AutoHotkey functions through a bridge: the methods of
; class Bridge below (the Macros tab's "Insert legal date" calls Bridge.LegalDate). They may use Doc (Doc.ahk) to
; drive the editor; so do the macros in the macros folder, which this script starts for the Macros tab.
;
; Usage:  AutoHotkey64.exe dAHKxodus.ahk [file.docx] [/measure out.docx]
;   /measure: open the file, put the caret at the end of the first paragraph, call LegalDate through the bridge,
;   save the edited copy to out.docx and close (run it on a hidden desktop). Timings go to logs\dahkxodus.log.

class Bridge {
    ; "the 13th day of November, 2026"
    static LegalDate(arg := "") {
        d := A_DD + 0, m := Mod(d, 10)
        suffix := (d >= 11 && d <= 13) || m = 0 || m > 3 ? "th" : ["st", "nd", "rd"][m]
        return "the " d suffix " day of " FormatTime(, "MMMM") ", " A_YYYY
    }
}

Persistent
OnMessage(0x004A, BridgeReceive)        ; WM_COPYDATA: a call from the editor
OnMessage(0x8001, (*) => ExitApp())     ; the editor's window has closed

; From the editor's window (wParam): a call, dwData "DAX1", text "id`tFunction`targument", or a macro to start,
; "DAX7", text "id`tpath". Taken at once (the window goes on), and run in a thread of its own, so a function may use
; Doc, which waits for that window; the answer goes back as WM_COPYDATA "DAX2" ("id`tresult") or "DAX3" ("id`terror").
BridgeReceive(wParam, lParam, msg, hwnd) {
    tag := NumGet(lParam, 0, "uptr")
    if tag != 0x44415831 && tag != 0x44415837
        return
    text := RTrim(StrGet(NumGet(lParam, 16, "ptr"), NumGet(lParam, 8, "uint") // 2, "UTF-16"), Chr(0))
    Doc.Window := wParam    ; (this script's Doc calls go to that window)
    SetTimer(BridgeRun.Bind(wParam, tag, text), -1)
    return 1
}

BridgeRun(editor, tag, text) {
    parts := StrSplit(text, "`t", , 3)
    id := parts[1]
    try {
        if tag = 0x44415837 {
            ; a macro: its own AutoHotkey process, told which window to drive (Doc.ahk reads DAHKXODUS_HWND)
            path := parts[2]
            if !FileExist(path)
                throw Error("No macro " path)
            SplitPath(path, , &folder)
            EnvSet("DAHKXODUS_HWND", editor)
            Run('"' A_AhkPath '" "' path '"', folder)
            result := ""
        } else {
            if !Bridge.HasMethod(parts[2])
                throw Error("No AutoHotkey bridge function " parts[2])
            result := String(Bridge.%parts[2]%(parts.Length > 2 ? parts[3] : ""))
        }
        answer := 0x44415832
    } catch as e {
        result := e.Message, answer := 0x44415833
    }
    reply := id "`t" result
    cds := Buffer(24, 0)
    NumPut("uptr", answer, cds, 0)
    NumPut("uint", (StrLen(reply) + 1) * 2, cds, 8)
    NumPut("ptr", StrPtr(reply), cds, 16)
    DllCall("SendMessageTimeoutW", "ptr", editor, "uint", 0x004A, "ptr", A_ScriptHwnd, "ptr", cds.Ptr,
        "uint", 2, "uint", 10000, "ptr*", 0)    ; (SMTO_ABORTIFHUNG)
}

; dAHKxodus.exe (vs\dAHKxodus.csproj) is the program and starts this script beside it as the bridge's AutoHotkey side:
;   AutoHotkey64.exe dAHKxodus.ahk /companion <window>
; It tells that window its own (WM_COPYDATA "DAX0"), answers its calls, and ends when the window is gone (the window
; posts WM_APP + 1 as it closes; the timer catches a window that ended without saying so).
if A_Args.Length >= 2 && A_Args[1] = "/companion" {
    Editor := A_Args[2] + 0
    Doc.Window := Editor
    cds := Buffer(24, 0)
    NumPut("uptr", 0x44415830, cds, 0)
    DetectHiddenWindows(true)
    try SendMessage(0x004A, A_ScriptHwnd, cds.Ptr, , "ahk_id " Editor, , , , 10000)
    catch
        ExitApp(1)
    SetTimer(() => WinExist("ahk_id " Editor) || ExitApp(), 1000)
    return
}

StartFile := "", MeasureOut := ""
for i, a in A_Args {
    if a = "/measure"
        MeasureOut := A_Args.Has(i + 1) ? A_Args[i + 1] : ""
    else if StartFile = "" && (i = 1 || A_Args[i - 1] != "/measure")
        StartFile := a
}
; .NET reads this when it starts (here, with the first C# call). The Docxodus engine runs in this process, and with
; profile-guided optimisation on, its precompiled code is first recompiled with counters added, and the open of a long
; document ran slower for it: a 9,500-paragraph document opened in 3.4 s with it off, 3.9 s with it on. A value set
; beforehand wins.
if EnvGet("DOTNET_TieredPGO") = ""
    EnvSet("DOTNET_TieredPGO", "0")
try
    RunCSharpFile(A_ScriptDir "\dAHKxodus.cs", A_ScriptHwnd, A_ScriptDir, StartFile, MeasureOut)
catch as e {
    DirCreate(A_ScriptDir "\logs")
    FileAppend(FormatTime(, "yyyy-MM-dd HH:mm:ss") " start failed: " e.Message "`r`n", A_ScriptDir "\logs\dahkxodus.log", "UTF-8")
    if MeasureOut = ""
        MsgBox("dAHKxodus could not start:`n`n" e.Message, "dAHKxodus", "Iconx")
    ExitApp(1)
}
