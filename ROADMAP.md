# Roadmap

Version numbers have four parts, as the Microsoft Store expects, and the fourth is always 0 (the Store reserves it):
0.0.N.0 when milestone N is finished. 0.0.2.0 will be the first installer preview, and 1.0.0.0 comes when dAHKxodus
can replace AhkWord for daily work. Milestone 0 was released as 0.0.1, before this scheme.

Most of what is planned changes Docxodus's editor or engine. That work goes into the fork,
[burque505/Docxodus](https://github.com/burque505/Docxodus), in small branches that can be offered upstream.
dAHKxodus itself adds what only a Windows program can: files, printing, AutoHotkey.

## Milestones

- [x] **0. The repository** (0.0.1): the trial's host and page, the build script and lock, headless engine checks
- [x] **1. Edits you can trust** (0.0.1.0): Undo and Redo by keyboard as by the buttons; Undo in steps like Word's; your own
  name and initials on changes and comments; comments with their date; pictures stored where Word stores them
- [x] **2. A desktop program** (0.0.2.0): Open, Save, Save As, recent files, ask to save on close, autosave and
  recovery, settings, print and PDF; the first installer preview, in a full and a light edition
- [x] **3. AutoHotkey** (0.0.3.0): a `Doc` object for scripts, with events; the command palette; macro recording and hotkeys;
  event hooks; your own buttons on the ribbon
- [x] **4. Other formats** (0.0.4.0): File > Save a Copy As: a web page (.html), Markdown (.md) or plain text (.txt), made by
  Docxodus's own converters; the window goes on with the .docx. `Doc.SaveCopy` for scripts
- [x] **5. Paste and paragraphs** (0.0.5.0): several windows, as in Word (done), with an Arrange windows macro; pasting from Word or a browser keeps paragraphs, lists and tables; the Paragraph
  dialog (exact indents, tab stops, keep with next...); multilevel lists as in AhkWord (define, change and keep each
  level's number format, text such as 1.1.1., indent and start; all nine levels); Format Painter; Change Case; more
  find options (done: whole words, wildcards, regular expressions, ^p across paragraphs)
- [x] **6. Pages** (0.0.6.0): the ruler (done: View > Ruler, indents, tab stops, Tabs dialog, margins, table columns); section breaks (done: Layout > Breaks, Section start, Ctrl+Enter page break); columns (done: Layout > Columns, presets and dialog); line numbers (done: Layout > Line Numbers, options dialog, page view); section breaks removable (Breaks menu, Backspace/Delete, page-view marks with ¶); table properties (done: Table tab Cell Size and Alignment, Table Properties dialog);  picture properties (done: Picture Format tab, size, wrap, position, alt text, drag in page view); Tab types a tab
- [ ] **7. References and styles:** bookmarks (done: Insert > Bookmark, Ctrl+Shift+F5, View > Bookmarks, links to a
  bookmark); the Navigation Pane (done: View > Navigation Pane, go to, drag to move with text, Word's right-click
  menu; the Styles box shows the caret's style); cross-references (done: Insert > Cross-reference to a heading, bookmark
  or numbered item; the Bookmark and Cross-reference dialogs stay open beside the text; Ctrl+Home/End, Ctrl+PgUp/PgDn); updating the table of contents and fields; new and
  changed styles, Copy Styles from Document; perhaps Outline View (collapse and expand under headings, show levels 1 to N)
- [ ] **8. Formats and parts:** RTF, OpenDocument and .doc through a converter found on the computer (LibreOffice,
  Word or Pandoc); opening those formats; templates and snippets
- [ ] **9. Spanish:** the whole program, the ribbon included
- [ ] **10. Text boxes:** editing in them, then inserting and moving them

Alongside every milestone: faster opening of long documents, and the fork kept level with Docxodus itself.

## Smaller downloads

The light installer leaves out .NET and AutoHotkey, but the editor itself is still about 36 MB before compression,
32 MB of it the WebAssembly build of Docxodus's engine. dAHKxodus edits through the native engine, so the page uses the
WebAssembly only for a few rarely used functions, yet loads it at every start. In the fork: make the WebAssembly
optional, loaded only when one of those functions is used (or not shipped at all once the native engine covers them).
That would shrink both installers and speed up the start.
