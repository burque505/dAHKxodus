#Requires AutoHotkey v2.0
#Include "..\Doc.ahk"
; Joins the paragraphs the selection touches into one, keeping the last paragraph mark: text copied from a PDF, with a
; paragraph for every printed line, becomes running text again. A space goes where two lines meet without one (none
; after a hyphen). One undo step.
if Doc.JoinParagraphs()
    Doc.SetStatus("Lines joined.")
else
    Doc.SetStatus("Select text across two or more paragraphs (not in a table) first.")
