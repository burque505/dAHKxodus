#Requires AutoHotkey v2.0
#Include "..\Doc.ahk"
; Shows the document's word, character and paragraph counts, and its file.
Text := Doc.Text
Paragraphs := 0, Characters := 0
Loop Parse Text, "`n" {
    if Trim(A_LoopField) != ""
        Paragraphs++
    Characters += StrLen(A_LoopField)
}
Path := Doc.Path
MsgBox(Format("Words:`t{}`nCharacters:`t{}`nParagraphs:`t{}`n`nFile:`t{}"
    , Doc.WordCount, Characters, Paragraphs, Path != "" ? Path : "(not saved yet)"), "Document statistics", "Iconi")
