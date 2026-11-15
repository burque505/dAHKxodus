#Requires AutoHotkey v2.0
#Include "..\Doc.ahk"
; (Measure runs) The BeforeSave and Saved event hooks of tests\run-measure.ps1's run: dAHKxodus runs this script by
; itself. As BeforeSave (dAHKxodus waits for it before saving) it types a mark at the end of the document, which the
; saved file must have; told the event the same way in A_Args and in the environment, or it types nothing.
Event := A_Args.Length >= 1 ? A_Args[1] : ""
Detail := A_Args.Length >= 2 ? A_Args[2] : ""
Doc.Log("hook-check: " Event " `"" Detail "`" (environment: " EnvGet("DAHKXODUS_EVENT") " `"" EnvGet("DAHKXODUS_EVENT_DETAIL") "`")")
if Event = "BeforeSave" && EnvGet("DAHKXODUS_EVENT") = Event && EnvGet("DAHKXODUS_EVENT_DETAIL") = Detail {
    Sleep(300)    ; (dAHKxodus must wait for this, not save at once)
    Doc.MoveCaret("documentEnd")
    Doc.InsertText("[hook before save]")
}
