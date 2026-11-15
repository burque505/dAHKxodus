#Requires AutoHotkey v2.0
#Include "..\Doc.ahk"
; Types today's date at the caret ("Sunday, 15 November 2026").
Doc.InsertText(FormatTime(, "dddd, d MMMM yyyy"))
