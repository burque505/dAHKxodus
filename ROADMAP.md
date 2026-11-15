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
- [ ] **3. AutoHotkey:** a `Doc` object for scripts, with events; the command palette; macro recording and hotkeys;
  event hooks; your own buttons on the ribbon
- [ ] **4. Paste and paragraphs:** pasting from Word or a browser keeps paragraphs, lists and tables; the Paragraph
  dialog (exact indents, tab stops, keep with next...); multilevel lists as in AhkWord (define, change and keep each
  level's number format, text such as 1.1.1., indent and start; all nine levels); Format Painter; Change Case; more
  find options
- [ ] **5. Pages:** the ruler; section breaks; columns; line numbers; table and picture properties
- [ ] **6. References and styles:** cross-references, updating the table of contents and fields, bookmarks; new and
  changed styles, Copy Styles from Document
- [ ] **7. Formats and parts:** Markdown, HTML, plain text, RTF, OpenDocument; templates and snippets
- [ ] **8. Spanish:** the whole program, the ribbon included
- [ ] **9. Text boxes:** editing in them, then inserting and moving them

Alongside every milestone: faster opening of long documents, and the fork kept level with Docxodus itself.

## Smaller downloads

The light installer leaves out .NET and AutoHotkey, but the editor itself is still about 36 MB before compression,
32 MB of it the WebAssembly build of Docxodus's engine. dAHKxodus edits through the native engine, so the page uses the
WebAssembly only for a few rarely used functions, yet loads it at every start. In the fork: make the WebAssembly
optional, loaded only when one of those functions is used (or not shipped at all once the native engine covers them).
That would shrink both installers and speed up the start.
