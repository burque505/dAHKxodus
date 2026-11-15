#Requires AutoHotkey v2.0
#Include "..\Doc.ahk"
; Shows dAHKxodus's events as they happen (what Doc.On can react to). Close this window, or dAHKxodus, to stop.
Monitor := Gui("+AlwaysOnTop +Resize +MinSize300x150", "dAHKxodus events")
Monitor.SetFont("s9", "Segoe UI")
List := Monitor.AddListView("w560 r14", ["Time", "Event", "Details"])
List.ModifyCol(1, 70), List.ModifyCol(2, 100), List.ModifyCol(3, 360)
Monitor.OnEvent("Size", (G, MinMax, W, H) => MinMax != -1 && List.Move(, , W - 20, H - 20))
Monitor.OnEvent("Close", (*) => (Doc.Off(), ExitApp()))

for Name in ["Opened", "Saved", "Modified", "TextChanged", "Command", "Closed"]
    Doc.On(Name, ShowEvent)
Monitor.Show("NoActivate")

ShowEvent(e) {
    switch e.Name {
        case "Opened", "Saved": Details := e.Path != "" ? e.Path : "(new document)"
        case "Modified": Details := e.Modified ? "changes not saved" : "no changes not saved"
        case "Command":
            Details := (e.Command != "" ? e.Command : e.Label)
            for k, v in e.Args.OwnProps()
                Details .= (A_Index = 1 ? ": " : ", ") k " " (IsObject(v) ? "{...}" : v)
        case "Closed": Details := "dAHKxodus has closed; this window closes in 2 s"
        default: Details := ""
    }
    Row := List.Add(, FormatTime(, "HH:mm:ss"), e.Name, Details)
    List.Modify(Row, "Vis")
    if e.Name = "Closed"
        SetTimer(() => ExitApp(), -2000)
}
