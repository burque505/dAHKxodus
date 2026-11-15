#Requires AutoHotkey v2.0
#Include "..\Doc.ahk"
; A small always-on-top window showing progress towards a word goal, updated as you type.
Answer := InputBox("Word goal for this document:", "Word goal", "w240 h110", 500)
if Answer.Result != "OK" || !IsInteger(Answer.Value) || Answer.Value < 1
    ExitApp
Goal := Integer(Answer.Value)

Tracker := Gui("+AlwaysOnTop +ToolWindow", "Word goal")
Tracker.SetFont("s9", "Segoe UI")
Bar := Tracker.AddProgress("w240 h16 Range0-" Goal)
Label := Tracker.AddText("w240 Center")
Tracker.OnEvent("Close", (*) => (Doc.Off(), ExitApp()))

UpdateGoal(*) {
    Words := Doc.WordCount
    Bar.Value := Min(Words, Goal)
    Label.Text := Words " of " Goal " words" (Words >= Goal ? "  -  goal reached!" : "")
}
Doc.On("TextChanged", UpdateGoal)
Doc.On("Opened", UpdateGoal)
Doc.On("Closed", (*) => ExitApp())
UpdateGoal()
Tracker.Show("NoActivate")
