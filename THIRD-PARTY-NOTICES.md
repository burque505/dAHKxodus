# Third-party notices

dAHKxodus itself is under the [MIT license](LICENSE). It builds on the components below, each under its own license; the full texts are in [`licenses/`](licenses).

## Built from source by `build-engine.ps1`

Not included in this repository: `build-engine.ps1` builds them from the Docxodus fork commit named in [`docxodus.lock`](docxodus.lock).

| Component | Used for | License | Text |
|---|---|---|---|
| [Docxodus](https://github.com/JSv4/Docxodus), from the fork [burque505/Docxodus](https://github.com/burque505/Docxodus) | The editor (the npm package's ribbon editor, in `web\`) and its editing engine (run natively, in `lib\docxodus`) | MIT | [Docxodus-MIT.txt](licenses/Docxodus-MIT.txt) |
| [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK), DocumentFormat.OpenXml.Framework | Reading and writing Word files, inside the engine | MIT | [Open-XML-SDK-MIT.txt](licenses/Open-XML-SDK-MIT.txt) |
| [System.IO.Packaging](https://github.com/dotnet/runtime) | Needed by the Open XML SDK | MIT | [dotnet-MIT.txt](licenses/dotnet-MIT.txt) |

The npm package carries further libraries of its own; their notices are in `web\node_modules\docxodus` after a build.

## Included in this repository

| Component | Where | License | Text |
|---|---|---|---|
| [Microsoft.Web.WebView2](https://www.nuget.org/packages/Microsoft.Web.WebView2) 1.0.4314.35 (the SDK; the WebView2 runtime itself comes with Windows) | `lib\webview2` | Microsoft's BSD-style license for the SDK | [WebView2-SDK-LICENSE.txt](licenses/WebView2-SDK-LICENSE.txt) |
| [Csharp-in-AHK](https://github.com/burque505/Csharp-in-AHK) | `Csharp-in-AHK.ahk` | MIT, by the author of dAHKxodus | [LICENSE](LICENSE) |
