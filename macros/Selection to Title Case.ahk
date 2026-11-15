#Requires AutoHotkey v2.0
#Include "..\Doc.ahk"
; Title-cases the selected text (within one paragraph; the selection's formatting becomes that of its start).
Selected := Doc.SelectedText
if Selected = "" {
    Doc.SetStatus("Select some text first.")
    ExitApp
}
if InStr(Selected, "`n") {
    Doc.SetStatus("Select text within one paragraph.")
    ExitApp
}
Doc.InsertText(StrTitle(Selected))
Doc.SetStatus("Selection converted to Title Case.")
