# dAHKxodus

A word processor for Windows that opens and saves Word documents (.docx) as they are, and that AutoHotkey scripts can drive while you type.

dAHKxodus is the successor of [AhkWord](https://github.com/burque505/AhkWord). It shows [Docxodus](https://github.com/JSv4/Docxodus)'s Word editor in a WebView2 window, and runs Docxodus's editing engine natively in .NET, so a long document opens in a few seconds. AutoHotkey v2 starts it and answers the editor's calls through a bridge. [Csharp-in-AHK](https://github.com/burque505/Csharp-in-AHK) compiles the C# host when it starts.

**Status: 0.0.1.0, an early start.** It opens, edits and saves .docx files; most of what a word processor needs is still to come. See [ROADMAP.md](ROADMAP.md).

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
- `tests\check-saved.ps1 out.docx`: checks that copy: the comment's author, initials and date, and the picture stored in `word/media`.

## Files

New, Open and Save are in the title bar, with Ctrl+N, Ctrl+O and Ctrl+S; Save As is F12 or Ctrl+Shift+S. The Recent list beside Open has the documents opened or saved lately, and a Word document dropped on the window opens in it. The title shows `*` while the document has changes not saved, and dAHKxodus asks before closing or replacing it. A save writes a new file beside the old one and then puts it in its place, so a failed save leaves the old file as it was.

## Settings

`settings.ini` is in `%APPDATA%\dAHKxodus`, with the list of recent documents. (Earlier versions kept it beside `dAHKxodus.ahk`; it is moved the first time.) dAHKxodus's other files on this computer, WebView2's profile among them, are in `%LOCALAPPDATA%\dAHKxodus`.

**Portable mode:** with a file named `portable` (or `portable.txt`) beside `dAHKxodus.ahk`, all of these go in a `data` folder beside it instead, so the folder can be carried to another computer.

### Your name on comments and changes

dAHKxodus puts your name and initials on the comments you add and the changes it tracks. It takes them from Office's user information (what Word uses) when Office is installed, otherwise from your Windows account. To choose them yourself, add them to `settings.ini`:

```ini
[User]
Name=Ada Lovelace
Initials=AL
```

## License

[MIT](LICENSE). Third-party components: [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).
