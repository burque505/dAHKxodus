#Requires AutoHotkey v2.0
#Include "..\Doc.ahk"
; Each time the document is saved, a PDF of it is made beside it (report.docx -> report.pdf). Run it to switch that on or off.
; It is an event hook: dAHKxodus runs it by itself after each save (Macros > Event hooks, Saved), so nothing keeps
; running. Run it yourself (the Macros list, the palette, or a hotkey) to switch PDF on save on or off.

if A_Args.Length >= 1 && A_Args[1] = "Saved" {
    ; run by dAHKxodus after a save: A_Args[2] is the file
    SavedFile := A_Args.Length >= 2 ? A_Args[2] : ""
    if SavedFile = ""
        ExitApp
    SplitPath(SavedFile, , &Folder, , &Name)
    try {
        Doc.ExportPdf(Folder "\" Name ".pdf")
        Doc.SetStatus("Saved, and made " Name ".pdf beside it.")
    } catch as e
        TrayTip(e.Message, "PDF on save failed", "Iconx")
    ExitApp
}

; run from the Macros list: PDF on save on or off
Hooked := "", On := false
for h in Doc.Hooks()
    if h.event = "Saved"
        Hooked := h.macro, On := h.on
if Hooked = "" || Hooked = A_ScriptName {
    On := Hooked = "" ? true : !On
    Doc.SetHook("Saved", A_ScriptName)
    Doc.SetHookOn("Saved", On)
    Doc.SetStatus(On ? "PDF on save is on: each save also makes a PDF beside the document (run PDF on save again to switch it off)."
        : "PDF on save is off (run it again, or tick On in Macros > Event hooks, to switch it on).")
} else if MsgBox("Saving already runs another macro, " Hooked " (Macros > Event hooks, Saved), and an event runs one "
    . "macro.`n`nRun PDF on save on each save instead?", "PDF on save", "YesNo Icon?") = "Yes" {
    Doc.SetHook("Saved", A_ScriptName)
    Doc.SetHookOn("Saved", true)
    Doc.SetStatus("PDF on save is on, instead of " Hooked ".")
}
