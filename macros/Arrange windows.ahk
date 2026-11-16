#Requires AutoHotkey v2.0
#Include "..\Doc.ahk"
; Arranges the open dAHKxodus windows side by side on the monitor of the active one (up to 3 in columns, more in a
; grid), each in page view with its page fitted to its width. (AhkWord's Arrange windows, for dAHKxodus.)

A_TitleMatchMode := "RegEx"
Windows := WinGetList(" - dAHKxodus$ ahk_class ^WindowsForms10\.Window")    ; front to back: the active one first
if Windows.Length = 0
    ExitApp
; the monitor the front window is on, without the taskbar
WinGetPos(&WX, &WY, &WW, &WH, Windows[1])
Mon := MonitorFromPoint(WX + WW // 2, WY + WH // 2)
MonitorGetWorkArea(Mon, &Left, &Top, &Right, &Bottom)
N := Windows.Length
Cols := N <= 3 ? N : Ceil(Sqrt(N)), Rows := Ceil(N / Cols)
W := (Right - Left) // Cols, H := (Bottom - Top) // Rows
for i, Hwnd in Windows {
    ; maximized first (a minimized or off-screen window comes back), then restored, as WinMove needs
    WinMaximize(Hwnd)
    WinRestore(Hwnd)
    Col := Mod(i - 1, Cols), Row := (i - 1) // Cols
    MoveVisible(Hwnd, Left + Col * W, Top + Row * H, W, H)
}
; each window: page view, the page's width fitted to the window (after the move: the fit is for the new size)
for Hwnd in Windows {
    Doc.Window := Hwnd
    try Doc.SetView(true, "width")
}
WinActivate(Windows[1])

; WinMove by the visible frame: Windows 10/11 windows have invisible borders around them.
MoveVisible(Hwnd, X, Y, W, H) {
    WinMove(X, Y, W, H, Hwnd)
    WinGetPos(&OX, &OY, &OW, &OH, Hwnd)
    Rect := Buffer(16)
    if DllCall("dwmapi\DwmGetWindowAttribute", "Ptr", Hwnd, "UInt", 9, "Ptr", Rect, "UInt", 16) != 0
        return                                    ; (9: DWMWA_EXTENDED_FRAME_BOUNDS)
    L := NumGet(Rect, 0, "Int"), T := NumGet(Rect, 4, "Int"), R := NumGet(Rect, 8, "Int"), B := NumGet(Rect, 12, "Int")
    WinMove(X - (L - OX), Y - (T - OY), W + (OW - (R - L)), H + (OH - (B - T)), Hwnd)
}

MonitorFromPoint(X, Y) {
    loop MonitorGetCount() {
        MonitorGet(A_Index, &ML, &MT, &MR, &MB)
        if X >= ML && X < MR && Y >= MT && Y < MB
            return A_Index
    }
    return MonitorGetPrimary()
}
