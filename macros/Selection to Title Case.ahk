#Requires AutoHotkey v2.0
#Include "..\Doc.ahk"
; Capitalizes each word of the selection (with nothing selected, the word at the caret), keeping its formatting.
; Doc.ChangeCase also takes "headline", "sentence", "lower", "upper", "toggle" and "cycle" (Shift+F3).
if Doc.ChangeCase("title")
    Doc.SetStatus("Selection converted to Title Case.")
else
    Doc.SetStatus("Nothing to change.")
