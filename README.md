# dAHKxodus

A word processor for Windows that opens and saves Word documents (.docx) as they are, and that AutoHotkey scripts can drive while you type.

dAHKxodus is the successor of [AhkWord](https://github.com/burque505/AhkWord). It shows [Docxodus](https://github.com/JSv4/Docxodus)'s Word editor in a WebView2 window, and runs Docxodus's editing engine natively in .NET, so a long document opens in a few seconds. AutoHotkey v2 starts it and answers the editor's calls through a bridge. [Csharp-in-AHK](https://github.com/burque505/Csharp-in-AHK) compiles the C# host when it starts.

**Status: 0.0.1, an early start.** It opens, edits and saves .docx files; most of what a word processor needs is still to come. See [ROADMAP.md](ROADMAP.md).

## Running it from the source

You need Windows 10 or 11 (with the WebView2 runtime, which Windows 11 has), [AutoHotkey v2](https://www.autohotkey.com/) 64-bit, the .NET 10 SDK with the `wasm-tools` workload, Node.js, and Git for Windows.

1. Clone this repository and, beside it, the Docxodus fork:
   `git clone https://github.com/burque505/Docxodus docxodus-src`
2. In the fork's clone, check out the commit named in [`docxodus.lock`](docxodus.lock).
3. In this repository, build the engine and the editor package (the WebAssembly build takes several minutes):
   `.\build-engine.ps1`
4. Start it: `AutoHotkey64.exe dAHKxodus.ahk [file.docx]`

## Checks

- `tests\engine`: headless checks of the editing engine (`dotnet run -c Release` in that folder).
- `AutoHotkey64.exe dAHKxodus.ahk file.docx /measure out.docx`: opens the file, makes one edit through the AutoHotkey bridge, saves the copy and closes; timings go to `logs\dahkxodus.log`.

## License

[MIT](LICENSE). Third-party components: [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).
