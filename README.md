# dAHKxodus

A word processor for Windows that opens and saves Word documents (.docx) as they are, and that AutoHotkey scripts can drive while you type.

dAHKxodus is the successor of [AhkWord](https://github.com/burque505/AhkWord). It shows [Docxodus](https://github.com/JSv4/Docxodus)'s Word editor in a WebView2 window, and runs Docxodus's editing engine natively in .NET, so a long document opens in a few seconds. AutoHotkey v2 starts it and answers the editor's calls through a bridge. [Csharp-in-AHK](https://github.com/burque505/Csharp-in-AHK) compiles the C# host when it starts.

**Status: 0.0.2.0, a desktop program.** It opens, edits and saves .docx files, with recent files, autosave and recovery, settings (English and Spanish), print and PDF; most of what a word processor needs is still to come. See [ROADMAP.md](ROADMAP.md).

## Installing it

Download from the [0.0.2.0 installer preview](https://github.com/burque505/dAHKxodus/releases/tag/installer-0.0.2.0):

- **Light** (`dAHKxodus-0.0.2.0-light-setup.exe`, 23 MB): uses the AutoHotkey v2 and the .NET 10 Desktop Runtime already on your computer; Setup offers to get them if it doesn't find them.
- **Full** (`dAHKxodus-0.0.2.0-setup.exe`, 56 MB): everything inside, nothing else to install.
- **Portable** (`...-portable.zip`, either edition): unzip and run `dAHKxodus.exe`; nothing is installed.

Both install for your Windows user only, without administrator rights. They aren't signed yet, so Windows warns about them; the release notes say what to do.

## Running it from the source

You need Windows 10 or 11 (with the WebView2 runtime, which Windows 11 has), [AutoHotkey v2](https://www.autohotkey.com/) 64-bit, the .NET 10 SDK with the `wasm-tools` workload, Node.js, and Git for Windows.

1. Clone this repository and, beside it, the Docxodus fork:
   `git clone https://github.com/burque505/Docxodus docxodus-src`
2. In the fork's clone, check out the commit named in [`docxodus.lock`](docxodus.lock).
3. In this repository, build the engine and the editor package (the WebAssembly build takes several minutes):
   `.\build-engine.ps1`
4. Start it: `AutoHotkey64.exe dAHKxodus.ahk [file.docx]`

### As dAHKxodus.exe

`vs\dAHKxodus.csproj` builds the same sources into `dAHKxodus.exe`, which starts without Csharp-in-AHK's compile. `.\build-exe.ps1` publishes it, precompiled, to `vs\publish`. You can also open `vs\dAHKxodus.slnx` in Visual Studio, where errors and the debugger point at the real files. Start it with `vs\publish\dAHKxodus.exe [file.docx]`. The exe finds `web\` and `dAHKxodus.ahk` in the folder above it, and starts `dAHKxodus.ahk /companion` beside it for the AutoHotkey bridge (AutoHotkey v2 from `DAHKXODUS_AUTOHOTKEY`, beside the exe, or where its installer put it). Running `dAHKxodus.ahk` directly stays the quick way to try a change to the sources.

## Checks

- `tests\engine`: headless checks of the editing engine (`dotnet run -c Release` in that folder).
- `AutoHotkey64.exe dAHKxodus.ahk file.docx /measure out.docx`: opens the file, makes one edit through the AutoHotkey bridge, checks Undo and Redo by keyboard, adds a comment and a picture, checks that the document reads as changed, saves it the way the program does before closing (to `out.host.docx`, never over the file opened) and checks that it then reads as unchanged, saves the copy and closes; timings and results go to `logs\dahkxodus.log`. `tests\run-measure.ps1` runs it on a desktop of its own; with `-Exe`, it runs `dAHKxodus.exe` instead.
  The measure run also writes a recovery copy (kept as `out.autosave.docx`) and checks that the save removes it. `-Recovery` starts it as after a crash (a recovery copy left behind, started without a document, the copy recovered); `-RecoveryLater` leaves a copy and starts with the file (the copy must wait); `-Language es` runs it in Spanish.
- `tests\check-lang.ps1`: checks that each `lang\*.json` has every text of `lang\en.json`, with the same placeholders. `node tests\check-ribbon-lang.mjs` checks the ribbon's tables (`lang\*.ribbon.json`).
- `tests\check-saved.ps1 out.docx`: checks that copy: the comment's author, initials and date, and the picture stored in `word/media`.

## Files

The **File** menu (the File tab, or Alt+F) has New, Open, Save, Save As, the recent documents, Settings and Close (Settings is also a button beside Save). New, Open and Save are also in the title bar, with Ctrl+N, Ctrl+O and Ctrl+S; Save As is F12 or Ctrl+Shift+S. The Recent list beside Open has the documents opened or saved lately, and a Word document dropped on the window opens in it. The title shows `*` while the document has changes not saved, and dAHKxodus asks before closing or replacing it. A save writes a new file beside the old one and then puts it in its place, so a failed save leaves the old file as it was.

### Print and PDF

**File > Print** (Ctrl+P) shows the print dialog, with its preview and printer; **File > Export to PDF** saves the document's pages as a PDF. Both print the page view as it shows (dAHKxodus switches to it at 100% for them, and back after), on paper the size of the document's first page, without the ribbon or the comment cards.

### Autosave and recovery

While the document has changes not saved, dAHKxodus keeps a recovery copy of it: every 10 minutes (Settings changes this; 0 turns it off) and whenever you switch to another window. The copy is removed when you save the document, open or start another, or close it (saved, or with Don't Save). The document's own file is never written by autosave.

If dAHKxodus ends without closing (a crash, a power cut, Windows restarting), the next start without a document offers the copy: **Yes** recovers it as changes not saved to the document it was (Save replaces that file; Save As keeps both), **No** discards it, **Cancel** decides next time. Started with a document, dAHKxodus says on the status bar that a copy is waiting. Each dAHKxodus window keeps copies of its own, so two windows never take each other's for left behind. The copies are in `%LOCALAPPDATA%\dAHKxodus\Recovery` (portable mode: `data\Recovery`).

## Settings

**File > Settings** (or Ctrl+,) opens dAHKxodus's settings:

- **Your name on comments and changes**, and your initials (see below).
- **Language**: English or Spanish, or Automatic (Windows's display language), for the ribbon and dAHKxodus's own dialogs and messages. The ribbon's translations are in `lang\<code>.ribbon.json`, keyed by its English (`tests\check-ribbon-lang.mjs` checks them). Still English: the status bar's running text and the comment cards' buttons.
- **Recent documents listed:** how many documents the Recent list keeps (0: no list).
- **Save a recovery copy every … minutes** (see Autosave and recovery; 0: never).
- **Open the last document when dAHKxodus starts without one.**

dAHKxodus also remembers the window's place and size, and the view (page or continuous, zoom, comments shown, the side rail), and starts with them next time.

`settings.ini` is in `%APPDATA%\dAHKxodus`, with the list of recent documents. (Earlier versions kept it beside `dAHKxodus.ahk`; it is moved the first time.) dAHKxodus's other files on this computer, WebView2's profile among them, are in `%LOCALAPPDATA%\dAHKxodus`.

**Portable mode:** with a file named `portable` (or `portable.txt`) beside `dAHKxodus.ahk`, all of these go in a `data` folder beside it instead, so the folder can be carried to another computer.

### Your name on comments and changes

dAHKxodus puts your name and initials on the comments you add and the changes it tracks. It takes them from Office's user information (what Word uses) when Office is installed, otherwise from your Windows account. To choose them yourself, type them in Settings, or add them to `settings.ini`:

```ini
[User]
Name=Ada Lovelace
Initials=AL
```

## License

[MIT](LICENSE). Third-party components: [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).
