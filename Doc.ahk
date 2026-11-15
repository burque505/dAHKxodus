#Requires AutoHotkey v2.0

; Doc.ahk - drive a running dAHKxodus from any AutoHotkey v2 script, and react to its events.
;
; Usage:
;   #Include "Doc.ahk"
;   Doc.InsertText("Dear Sir,`n")         ; any editor command by name (Doc.Help() lists them)
;   Doc.Bold()                            ; toggles; Doc.Bold(true) sets
;   Doc.FontSize(14)                      ; arguments in the order Doc.Help() shows them ...
;   Doc.Find({text: "Sir", matchCase: true})    ; ... or as one object, by name
;   MsgBox Doc.Text                       ; queries read as properties (or Doc.Text())
;   Doc.Save("C:\Letters\letter.docx")    ; the program's own: Open, New, Save, ExportPdf, Close, SetStatus, ...
;
;   Doc.On("Saved", (e) => ToolTip("Saved " e.Path))    ; react to an event
;   Doc.Off("Saved")                                    ; stop (Doc.Off() stops all)
;
; Commands act where the user would: on the selection, or at the caret. Each editing command is one undo step.
; Results come back as AutoHotkey values: text, numbers, true/false (1/0), objects (e.g. Doc.FormatState().bold) and
; arrays; JSON null as "". A command that fails throws a DocError, whose Code says why: "unknown_command",
; "bad_arguments", "closed" (no document), "refused" (nothing to act on), "cancelled" (the user cancelled a dialog),
; "busy", "error".
;
; Calls wait for their result (at most Doc.Timeout ms). The editor runs each one when the ones before it, from any
; script, have finished. A script started from dAHKxodus's Macros tab talks to that window (DAHKXODUS_HWND); any
; other script to the dAHKxodus used last (the frontmost; several can be open), or to Doc.Window if it is set.
; Don't call Doc from a Critical thread: the answer comes as a message, which a Critical thread doesn't take.
;
; Events: callbacks run one at a time, in the order the events happened, in a thread of their own; they may call Doc.
; Each gets an object e with e.Name and, depending on the event:
;   Opened       e.Path ("" for a new document), e.FileName a document was opened, or a new one started
;   Saved        e.Path
;   Modified     e.Modified (true: changes not saved)       when that changes
;   TextChanged                                             when typing pauses (about 0.3 s), or after a command
;   Command      e.Command (its name, "" for a ribbon control that has none), e.Label, e.Args (an object)
;   Closed                                                  the window has closed
; Doc.On makes the script persistent: exit with ExitApp (for example on the Closed event).
;
; Event hooks need no script running: dAHKxodus starts a macro by itself when an event happens (Macros > Event hooks,
; or Doc.SetHook("Saved", "PDF on save.ahk")). The macro gets the event in A_Args[1] (Started, Opened, NewDocument,
; BeforeSave, Saved, PdfExported, Closing) and the document's file or the PDF in A_Args[2] (also in the environment:
; DAHKXODUS_EVENT, DAHKXODUS_EVENT_DETAIL). BeforeSave and Closing wait for it (at most 15 s). Doc.Hooks() lists them;
; Doc.SetHookOn("Saved", false) switches one off, its macro kept.

class DocError extends Error {
    Code := "error"
}

class Doc {
    static Timeout := 30000
    ; The dAHKxodus window to talk to; 0: find it (see above).
    static Window := 0

    ; True if a dAHKxodus is running.
    static IsRunning => this._Target(false) != 0

    ; Doc.Exec("fontSize", {pts: 14}) is the same as Doc.FontSize(14).
    static Exec(Name, Args?) => this._Request(Name, IsSet(Args) ? Args : Map())

    ; Every command: an array of objects with Name, Group, Description, Args (names in order, "?" after optional
    ; ones) and Returns.
    static Commands() {
        list := []
        for c in this._Info().list
            list.Push({Name: c.name, Group: c.group, Description: c.description, Args: c.order.Clone(),
                Returns: c.HasOwnProp("returns") ? c.returns : ""})
        return list
    }

    ; The commands as text, one a line: what Doc can do.
    static Help() {
        text := "", group := ""
        for c in this._Info().list {
            if c.group != group
                text .= (text = "" ? "" : "`n") "[" (group := c.group) "]`n"
            args := ""
            for a in c.order
                args .= (args = "" ? "" : ", ") a
            text .= "  " StrUpper(SubStr(c.name, 1, 1)) SubStr(c.name, 2) "(" args ")"
                . (c.HasOwnProp("returns") && c.returns != "" ? " -> " c.returns : "") "  " c.description "`n"
        }
        return text
    }

    ; Calls Callback(e) whenever the event happens (see the list above).
    static On(Event, Callback) {
        Name := this._EventName(Event)
        if !this._Handlers.Has(Name) {
            this._Request("subscribe", Map("event", Name))
            this._Handlers[Name] := []
        }
        this._Handlers[Name].Push(Callback)
        Persistent()
    }

    ; Removes one callback, all callbacks of an event, or (no parameters) all callbacks.
    static Off(Event := "*", Callback?) {
        Names := []
        if Event = "*" {
            for Name in this._Handlers
                Names.Push(Name)
        } else
            Names.Push(this._EventName(Event))
        for Name in Names {
            if !this._Handlers.Has(Name)
                continue
            Callbacks := this._Handlers[Name]
            if IsSet(Callback) {
                Loop Callbacks.Length
                    if Callbacks[i := Callbacks.Length - A_Index + 1] = Callback
                        Callbacks.RemoveAt(i)
                if Callbacks.Length
                    continue
            }
            this._Handlers.Delete(Name)
            try this._Request("unsubscribe", Map("event", Name))    ; (the window may have closed already)
        }
    }

    static __Call(Name, Params) {
        c := this._Command(Name)
        return this._Request(c.name, this._ArgsFor(c, Params))
    }

    ; Queries (and the program's Path, Name, Modified) read as properties: Doc.Text, Doc.WordCount, ...
    static __Get(Name, Params) {
        c := this._Command(Name)
        if c.group != "query" || Params.Length
            throw DocError(Name " is a command, not a property: call it as Doc." Name "()", -1)
        return this._Request(c.name, Map())
    }

    static _NextId := 0
    static _Replies := Map()
    static _Handlers := Map()
    static _Queue := []
    static _InfoCache := 0
    static _InfoWindow := 0
    static _EventNames := ["Opened", "Saved", "Modified", "TextChanged", "Command", "Closed"]

    static __New() {
        if this != Doc
            return
        OnMessage(0x4A, ObjBindMethod(this, "_OnCopyData"))    ; WM_COPYDATA: answers and events
        this._RunHandlersTimer := ObjBindMethod(this, "_RunHandlers")
    }

    static _Request(Name, Args) {
        Window := this._Target()
        Id := ++this._NextId
        Text := DocJson.Dump(Map("id", Id, "name", Name, "args", Args))
        Data := Buffer(3 * A_PtrSize, 0)    ; COPYDATASTRUCT: "DAX4", a call
        NumPut("ptr", 0x44415834, "uint", (StrLen(Text) + 1) * 2, Data)
        NumPut("ptr", StrPtr(Text), Data, 2 * A_PtrSize)
        Hidden := A_DetectHiddenWindows
        A_DetectHiddenWindows := true
        try {
            try Accepted := SendMessage(0x4A, A_ScriptHwnd, Data.Ptr, , "ahk_id " Window, , , , 5000)
            catch TimeoutError
                throw DocError("dAHKxodus did not answer (" Name ").", -2)
            if Accepted != 1
                throw DocError("dAHKxodus did not take the call (" Name ").", -2)
            ; The answer comes as a message of its own ("DAX5"), when the editor has run the command.
            Deadline := A_TickCount + this.Timeout
            while !this._Replies.Has(Id) {
                if A_TickCount > Deadline
                    throw DocError("dAHKxodus did not finish " Name " within " this.Timeout " ms.", -2)
                if !WinExist("ahk_id " Window)
                    throw DocError("dAHKxodus has closed (" Name ").", -2)
                ; (wakes for any message, or after 50 ms; Sleep -1 then handles it)
                DllCall("MsgWaitForMultipleObjects", "uint", 0, "ptr", 0, "int", 0, "uint", 50, "uint", 0x4FF)
                Sleep(-1)
            }
        } finally
            A_DetectHiddenWindows := Hidden
        Reply := this._Replies.Delete(Id)
        if Reply.HasOwnProp("error") {
            e := DocError(Reply.error, -2, Name)
            e.Code := Reply.HasOwnProp("code") && Reply.code != "" ? Reply.code : "error"
            throw e
        }
        return Reply.HasOwnProp("result") ? Reply.result : ""
    }

    static _OnCopyData(wParam, lParam, Msg, Hwnd) {
        Tag := NumGet(lParam, "ptr")
        if Tag != 0x44415835 && Tag != 0x44415836
            return
        Text := StrGet(NumGet(lParam, 2 * A_PtrSize, "ptr"), NumGet(lParam, A_PtrSize, "uint") // 2, "UTF-16")
        try Value := DocJson.Load(RTrim(Text, Chr(0)))
        catch
            return 1
        if Tag = 0x44415835 {    ; an answer
            this._Replies[Value.id] := Value
            return 1
        }
        ; An event: queued; the callbacks run in a timer thread, so this returns at once.
        e := {Name: Value.event}
        for k, v in Value.OwnProps()
            if k != "event"
                e.%StrUpper(SubStr(k, 1, 1)) SubStr(k, 2)% := v
        this._Queue.Push(e)
        SetTimer(this._RunHandlersTimer, -1)
        return 1
    }

    static _RunHandlers() {
        while this._Queue.Length {
            e := this._Queue.RemoveAt(1)
            for Callback in this._Handlers.Get(e.Name, []).Clone()
                Callback(e)
        }
    }

    static _EventName(Event) {
        for Name in this._EventNames
            if Name = Event
                return Name
        throw DocError("Unknown event: " Event ". Events: Opened, Saved, Modified, TextChanged, Command, Closed.", -3)
    }

    ; The command list, fetched once from the window (and again if the window changes).
    static _Info() {
        Window := this._Target()
        if this._InfoCache && this._InfoWindow = Window
            return this._InfoCache
        list := this._Request("commands", Map())
        byName := Map()
        byName.CaseSense := false
        for c in list
            byName[c.name] := c
        this._InfoWindow := Window
        return this._InfoCache := {list: list, byName: byName}
    }

    static _Command(Name) {
        info := this._Info()
        if !info.byName.Has(Name) {
            e := DocError('There is no command "' Name '" (Doc.Help() lists them).', -3, Name)
            e.Code := "unknown_command"
            throw e
        }
        return info.byName[Name]
    }

    ; The arguments as the editor takes them: one object (or Map) as they are, else by position in the listed order.
    static _ArgsFor(c, Params) {
        if Params.Length = 1 && IsObject(Params[1]) && !(Params[1] is Array) && !(c.order.Length = 1 && InStr(c.args.%c.order[1]%, "object"))
            return Params[1]
        if Params.Length > c.order.Length {
            e := DocError(c.name " takes " c.order.Length " argument(s), not " Params.Length ".", -3, c.name)
            e.Code := "bad_arguments"
            throw e
        }
        args := Map()
        for i, p in Params
            if IsSet(p)
                args[RTrim(c.order[i], "?")] := p
        return args
    }

    static _Target(Required := true) {
        Hidden := A_DetectHiddenWindows, Mode := A_TitleMatchMode
        A_DetectHiddenWindows := true, A_TitleMatchMode := "RegEx"
        try {
            if this.Window && WinExist("ahk_id " this.Window)
                return this.Window + 0
            if (Hwnd := EnvGet("DAHKXODUS_HWND")) != "" && WinExist("ahk_id " Hwnd)
                return Hwnd + 0
            ; the dAHKxodus used last: the frontmost of its windows (WinGetList lists them front to back)
            for Window in WinGetList(" - dAHKxodus$ ahk_class ^WindowsForms10\.Window")
                return Window
        } finally
            A_DetectHiddenWindows := Hidden, A_TitleMatchMode := Mode
        if Required
            throw DocError("dAHKxodus is not running.", -3)
        return 0
    }
}

; JSON for the calls: AutoHotkey values to JSON text and back (objects come back as Objects, arrays as Arrays, true
; and false as 1 and 0, null as "").
class DocJson {
    static Dump(v) {
        if v is Map {
            s := ""
            for k, x in v
                s .= (s = "" ? "" : ",") this._Str(String(k)) ":" this.Dump(x)
            return "{" s "}"
        }
        if v is Array {
            s := ""
            Loop v.Length
                s .= (A_Index = 1 ? "" : ",") (v.Has(A_Index) ? this.Dump(v[A_Index]) : "null")
            return "[" s "]"
        }
        if IsObject(v) {
            s := ""
            for k, x in v.OwnProps()
                s .= (s = "" ? "" : ",") this._Str(k) ":" this.Dump(x)
            return "{" s "}"
        }
        if v is Integer || v is Float
            return String(v)
        return this._Str(v)
    }

    static _Str(s) {
        s := StrReplace(StrReplace(s, "\", "\\"), '"', '\"')
        s := StrReplace(StrReplace(StrReplace(s, "`n", "\n"), "`r", "\r"), "`t", "\t")
        while RegExMatch(s, "[\x00-\x1f]", &m)
            s := StrReplace(s, m[0], Format("\u{:04x}", Ord(m[0])))
        return '"' s '"'
    }

    static Load(text) {
        pos := 1
        v := this._Value(text, &pos)
        return v
    }

    static _Value(t, &p) {
        p := RegExMatch(t, "\S", , p)
        if !p
            throw ValueError("JSON: unexpected end")
        c := SubStr(t, p, 1)
        if c = "{" {
            obj := {}, p++
            if RegExMatch(t, "\G\s*\}", &m, p)
                return (p += m.Len, obj)
            Loop {
                p := RegExMatch(t, "\S", , p)
                k := this._String(t, &p)
                if !RegExMatch(t, "\G\s*:", &m, p)
                    throw ValueError("JSON: ':' expected at " p)
                p += m.Len
                obj.%k% := this._Value(t, &p)
                if !RegExMatch(t, "\G\s*([,}])", &m, p)
                    throw ValueError("JSON: ',' or '}' expected at " p)
                p += m.Len
                if m[1] = "}"
                    return obj
            }
        }
        if c = "[" {
            arr := [], p++
            if RegExMatch(t, "\G\s*\]", &m, p)
                return (p += m.Len, arr)
            Loop {
                arr.Push(this._Value(t, &p))
                if !RegExMatch(t, "\G\s*([,\]])", &m, p)
                    throw ValueError("JSON: ',' or ']' expected at " p)
                p += m.Len
                if m[1] = "]"
                    return arr
            }
        }
        if c = '"'
            return this._String(t, &p)
        if RegExMatch(t, "\G-?\d+(\.\d+)?([eE][-+]?\d+)?", &m, p) {
            p += m.Len
            return m[1] = "" && m[2] = "" ? Integer(m[0]) : Float(m[0])
        }
        for word, value in Map("true", true, "false", false, "null", "")
            if SubStr(t, p, StrLen(word)) == word
                return (p += StrLen(word), value)
        throw ValueError("JSON: unexpected '" c "' at " p)
    }

    static _String(t, &p) {
        if SubStr(t, p, 1) != '"'
            throw ValueError("JSON: string expected at " p)
        ; (the closing quote: the first one after an even number of backslashes; found by hand, as a pattern would
        ; recurse once per escape and a long document's text has thousands)
        i := p + 1
        Loop {
            if !q := InStr(t, '"', true, i)
                throw ValueError("JSON: unterminated string at " p)
            n := 0
            while SubStr(t, q - 1 - n, 1) = "\"
                n++
            if !Mod(n, 2)
                break
            i := q + 1
        }
        s := SubStr(t, p + 1, q - p - 1)
        p := q + 1
        if !InStr(s, "\")
            return s
        out := "", i := 1
        while RegExMatch(s, "\\(u([0-9a-fA-F]{4})|.)", &e, i) {
            out .= SubStr(s, i, e.Pos - i)
            i := e.Pos + e.Len
            switch e[1], true {
                case "n": out .= "`n"
                case "r": out .= "`r"
                case "t": out .= "`t"
                case "b": out .= "`b"
                case "f": out .= "`f"
                default:
                    if e[2] != "" {
                        code := Integer("0x" e[2])
                        ; (a surrogate pair: the second half follows as \uXXXX)
                        if code >= 0xD800 && code <= 0xDBFF && RegExMatch(s, "\G\\u([dD][c-fC-F][0-9a-fA-F]{2})", &lo, i) {
                            code := 0x10000 + ((code - 0xD800) << 10) + (Integer("0x" lo[1]) - 0xDC00)
                            i += lo.Len
                        }
                        out .= Chr(code)
                    } else
                        out .= e[1]
            }
        }
        return out SubStr(s, i)
    }
}
