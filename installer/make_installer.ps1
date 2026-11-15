# Makes dAHKxodus's installers in ..\dAHKxodus-releases (Inno Setup 6, installer\dAHKxodus.iss), each with a portable
# zip of the same files (a "portable" file beside dAHKxodus.exe: settings, recent files, recovery copies and the WebView2
# profile in its data folder, nothing in %APPDATA% or %LOCALAPPDATA%). (As AhkWord's make_installer.ps1.)
#   full:  dAHKxodus-<version>-setup.exe, -portable.zip: everything inside. dAHKxodus.exe self-contained (.NET 10 in it),
#          and AutoHotkey64.exe beside it.
#   light: dAHKxodus-<version>-light-setup.exe, -light-portable.zip: no .NET and no AutoHotkey. dAHKxodus.exe needs the
#          .NET 10 Desktop Runtime and uses the AutoHotkey v2 already installed; Setup finds both, or offers to download
#          .NET and to locate AutoHotkey64.exe (AutoHotkey.path) or get it from the Microsoft Store.
# With -Publish it puts them on GitHub as the pre-release "dAHKxodus <version> (installer preview)" (tag installer-<version>).
# Both carry dAHKxodus.exe built by vs\dAHKxodus.csproj from the same sources dAHKxodus.ahk runs (ReadyToRun), the editor
# (web\: index.html and the Docxodus package docxodus.lock names), lang\, the licenses, and dAHKxodus.ahk (the bridge).
#   .\installer\make_installer.ps1             both editions, from the last commit (uncommitted changes are left out)
#   -Edition Full | Light                      only one
#   -Working                                   from the working tree as it is (named ...-working; for trying it out)
#   -Publish                                   then create or update the GitHub pre-release (needs a clean, pushed main)
#   -AutoHotkey <path>                         the AutoHotkey64.exe the full edition carries (its license.txt beside it or above)
# The engine: the native part is rebuilt here (build-engine.ps1 -Part native: the fork clone must be at the lock's
# commit); the editor package is the one web\ has installed, and must be the one docxodus.lock names (its hash).
# Neither dAHKxodus.exe nor the installers are signed (yet): the release notes say what Windows will show.
param(
	[string]$AutoHotkey = "$env:ProgramFiles\AutoHotkey\v2\AutoHotkey64.exe",
	[ValidateSet("Both", "Full", "Light")][string]$Edition = "Both",
	[switch]$Working,
	[switch]$Publish
)
$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot
$releases = Join-Path (Split-Path $root) "dAHKxodus-releases"
$repo = "burque505/dAHKxodus"
if ($Working -and $Publish) { throw "-Publish makes the installers from the last commit: leave out -Working." }
if ($Publish -and $Edition -ne "Both") { throw "-Publish makes both editions: leave out -Edition." }
$editions = if ($Edition -eq "Both") { @("Full", "Light") } else { @($Edition) }

# the version (dAHKxodus.cs's Version, as committed, or as it is with -Working)
$cs = if ($Working) { Get-Content (Join-Path $root "dAHKxodus.cs") } else { git -C $root show HEAD:dAHKxodus.cs }
$version = [regex]::Match(($cs -join "`n"), 'public const string Version = "(\d+\.\d+\.\d+\.\d+)";').Groups[1].Value
if (-not $version) { throw "dAHKxodus.cs has no Version." }
$changes = git -C $root status --porcelain
if ($Publish) {
	if ($changes) { throw "-Publish: commit first (there are uncommitted changes)." }
	git -C $root fetch -q origin main
	if ((git -C $root rev-parse HEAD) -ne (git -C $root rev-parse origin/main)) { throw "-Publish: push main first (HEAD isn't origin/main)." }
} elseif ($changes -and -not $Working) { "Note: there are uncommitted changes; the installers are made from the last commit without them." }
$suffix = if ($Working) { "-working" } else { "" }
function NameOf([string]$edition, [string]$kind) {
	$light = if ($edition -eq "Light") { "-light" } else { "" }
	"dAHKxodus-$version$suffix$light-$kind"
}

# the tools: Inno Setup's compiler; AutoHotkey and its license (the full edition)
$iscc = @("$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe", "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe", "$env:ProgramFiles\Inno Setup 6\ISCC.exe") |
	Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) { throw "Inno Setup 6 isn't installed (winget install JRSoftware.InnoSetup)." }
$ahkVersion = ""
if ($editions -contains "Full") {
	if (-not (Test-Path $AutoHotkey)) { throw "AutoHotkey64.exe was not found at $AutoHotkey (pass -AutoHotkey <path>)." }
	$ahkVersion = (Get-Item $AutoHotkey).VersionInfo.FileVersion.Trim()
	$license = @((Join-Path (Split-Path $AutoHotkey) "license.txt"), (Join-Path (Split-Path (Split-Path $AutoHotkey)) "license.txt")) | Where-Object { Test-Path $_ } | Select-Object -First 1
	if (-not $license) { throw "AutoHotkey's license.txt was not found beside $AutoHotkey or in the folder above it." }
}

# the engine: the native part rebuilt from the lock's commit, the editor package checked against the lock
$lockText = if ($Working) { Get-Content (Join-Path $root "docxodus.lock") } else { git -C $root show HEAD:docxodus.lock }
$lock = @{}
foreach ($line in $lockText) { if ($line -match '^\s*([\w-]+)=(.*)$') { $lock[$Matches[1]] = $Matches[2].Trim() } }
"Building the native engine (build-engine.ps1 -Part native, fork commit $($lock.commit.Substring(0, 7))) ..."
& (Join-Path $root "build-engine.ps1") -Part native | Out-Host
$tgz = Join-Path $root "web\vendor\docxodus-0.0.0.tgz"
$package = Join-Path $root "web\node_modules\docxodus"
if (-not (Test-Path $tgz) -or -not (Test-Path (Join-Path $package "dist\embed.bundle.js"))) { throw "The editor package isn't built: run .\build-engine.ps1 -Part npm." }
$tgzHash = (Get-FileHash $tgz -Algorithm SHA256).Hash.ToLowerInvariant()
if ($tgzHash -ne $lock["npm-sha256"]) { throw "web\vendor's editor package ($tgzHash) isn't the one docxodus.lock names ($($lock["npm-sha256"])): run .\build-engine.ps1 -Part npm." }

# the privacy check's pattern (the owner's, kept outside the repositories)
$patternFile = Join-Path (Split-Path $root) "AhkWord-tests\privacy_pattern.txt"
if (-not (Test-Path $patternFile)) { throw "The privacy check's pattern isn't at $patternFile." }
$pattern = (Get-Content $patternFile -Raw).Trim()

New-Item -ItemType Directory -Force $releases | Out-Null
$work = Join-Path ([IO.Path]::GetTempPath()) ("dahkxodus-installer-" + [guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Force $work | Out-Null
$made = @()
$dotnetVersion = ""
try {
	# the sources: the last commit (git archive), or the working tree's files; and the engine's DLLs (not in git)
	$src = Join-Path $work "src"
	if ($Working) {
		foreach ($file in git -C $root ls-files -co --exclude-standard) {
			$from = Join-Path $root $file
			if (-not (Test-Path $from -PathType Leaf)) { continue }
			$to = Join-Path $src $file
			New-Item -ItemType Directory -Force (Split-Path $to) | Out-Null
			Copy-Item $from $to
		}
	} else {
		$archive = Join-Path $work "head.zip"
		git -C $root archive --format=zip -o $archive HEAD
		Expand-Archive $archive $src
	}
	Copy-Item (Join-Path $root "lib\docxodus") (Join-Path $src "lib\docxodus") -Recurse -Force

	# the .NET the full edition carries: the latest .NET 10 (a security fix most months), pinned by
	# RuntimeFrameworkVersion (else the SDK's own, often older). The light edition isn't pinned: any .NET 10 installed
	# runs it (it rolls forward to the newest patch there).
	$latest = $null
	if ($editions -contains "Full") {
		try { $latest = (Invoke-RestMethod "https://builds.dotnet.microsoft.com/dotnet/release-metadata/10.0/releases.json" -TimeoutSec 30)."latest-runtime" }
		catch { "Note: Microsoft's list of .NET releases couldn't be read; the SDK's .NET is used." }
	}

	foreach ($ed in $editions) {
		$full = $ed -eq "Full"
		# dAHKxodus.exe, ReadyToRun (no .pdb: nothing of this computer's paths goes with it)
		"Building dAHKxodus.exe for the $($ed.ToLower()) edition (dotnet publish, $(if ($full) { 'self-contained' } else { 'framework-dependent' })) ..."
		$pubDir = Join-Path $work "publish-$ed"
		$pin = if ($full -and $latest) { "-p:RuntimeFrameworkVersion=$latest" } else { "-p:DaxNoPin=1" }
		dotnet publish (Join-Path $src "vs\dAHKxodus.csproj") -c Release -r win-x64 --self-contained $(if ($full) { "true" } else { "false" }) -o $pubDir $pin `
			-p:PublishReadyToRun=true -p:Version=$version -p:FileVersion=$version -p:AssemblyVersion=$version `
			-p:DebugType=none -p:DebugSymbols=false -p:Deterministic=true -p:ContinuousIntegrationBuild=true -nologo -v:q | Out-Host
		if ($LASTEXITCODE -ne 0 -or -not (Test-Path (Join-Path $pubDir "dAHKxodus.exe"))) { throw "dotnet publish failed." }
		if ($full) {
			$included = (Get-Content (Join-Path $pubDir "dAHKxodus.runtimeconfig.json") -Raw | ConvertFrom-Json).runtimeOptions.includedFrameworks
			$dotnetVersion = ($included | Where-Object name -eq "Microsoft.NETCore.App").version
			$desktopVersion = ($included | Where-Object name -eq "Microsoft.WindowsDesktop.App").version
			if ($latest -and ([version]$dotnetVersion -ne [version]$latest -or [version]$desktopVersion -ne [version]$latest)) {
				throw ".NET $dotnetVersion / Windows Desktop $desktopVersion came out, not $latest."
			}
			"dAHKxodus.exe with .NET $dotnetVersion (Windows Desktop $desktopVersion)"
		}

		# the program's folder: the exe and its files, then dAHKxodus's own files beside them (only what the exe uses:
		# the .cs sources, tests, samples and build scripts stay in the repository)
		$folder = Join-Path $work "dAHKxodus-$ed"
		Copy-Item $pubDir $folder -Recurse
		foreach ($own in "dAHKxodus.ahk", "Csharp-in-AHK.ahk", "LICENSE", "README.md", "ROADMAP.md", "THIRD-PARTY-NOTICES.md") {
			$to = Join-Path $folder $own
			if (Test-Path $to) { throw "dAHKxodus.exe's file $own has the name of one of dAHKxodus's own files." }
			Copy-Item (Join-Path $src $own) $to
		}
		foreach ($dir in "lang", "licenses") { Copy-Item (Join-Path $src $dir) (Join-Path $folder $dir) -Recurse }
		New-Item -ItemType Directory -Force (Join-Path $folder "web") | Out-Null
		foreach ($own in "index.html", "ribbon-open-async.js") { Copy-Item (Join-Path $src "web\$own") (Join-Path $folder "web\$own") }
		# the editor package, without its source maps and type declarations (11 MB the page never loads)
		$toPackage = Join-Path $folder "web\node_modules\docxodus"
		foreach ($item in Get-ChildItem $package -Recurse -File) {
			if ($item.Name -match '\.(map|d\.ts)$') { continue }
			$to = Join-Path $toPackage $item.FullName.Substring($package.Length).TrimStart('\')
			New-Item -ItemType Directory -Force (Split-Path $to) | Out-Null
			Copy-Item $item.FullName $to
		}

		if ($full) {
			# AutoHotkey beside dAHKxodus.exe, its license and what it is there for
			Copy-Item $AutoHotkey (Join-Path $folder "AutoHotkey64.exe")
			Copy-Item $license (Join-Path $folder "licenses\AutoHotkey-GPL-2.0.txt")
			# .NET's license and notices, from the runtime packs the publish used
			$nuget = if ($env:NUGET_PACKAGES) { $env:NUGET_PACKAGES } else { Join-Path $env:USERPROFILE ".nuget\packages" }
			$dotnetLicenses = Join-Path $folder "licenses\dotnet-runtime"
			New-Item -ItemType Directory -Force $dotnetLicenses | Out-Null
			$core = Join-Path $nuget "microsoft.netcore.app.runtime.win-x64\$dotnetVersion"
			$desk = Join-Path $nuget "microsoft.windowsdesktop.app.runtime.win-x64\$desktopVersion"
			Copy-Item (Join-Path $core "LICENSE.TXT") (Join-Path $dotnetLicenses "LICENSE.txt")
			Copy-Item (Join-Path $core "THIRD-PARTY-NOTICES.TXT") (Join-Path $dotnetLicenses "ThirdPartyNotices.txt")
			Copy-Item (Join-Path $desk "LICENSE") (Join-Path $dotnetLicenses "WindowsDesktop-LICENSE.txt")
			Set-Content (Join-Path $folder "AutoHotkey64.exe.txt") -Encoding UTF8 @(
				"AutoHotkey64.exe is AutoHotkey $ahkVersion (64-bit), unchanged. dAHKxodus.exe starts it with dAHKxodus.ahk",
				"(and Csharp-in-AHK.ahk, which it includes) for the AutoHotkey bridge: the Macros tab's macros.",
				"",
				"AutoHotkey is free software under the GNU General Public License, version 2: see licenses\AutoHotkey-GPL-2.0.txt.",
				"Its source code, for this version: https://github.com/AutoHotkey/AutoHotkey/tree/v$ahkVersion",
				"AutoHotkey's site: https://www.autohotkey.com/",
				"",
				"dAHKxodus's own files are under the MIT license (LICENSE); AutoHotkey64.exe is only bundled with them.")
		}

		# every package dAHKxodus.exe carries must be named in THIRD-PARTY-NOTICES.md (its references are DLLs from
		# lib\, named there by hand; this catches NuGet packages, should any come in)
		$named = Get-Content (Join-Path $folder "THIRD-PARTY-NOTICES.md") -Raw
		$unnamed = @()
		$assetsJson = Get-Content (Join-Path $src "vs\obj\project.assets.json") -Raw | ConvertFrom-Json
		foreach ($lib in $assetsJson.libraries.PSObject.Properties) {
			if ($lib.Value.type -ne "package") { continue }
			$id = ($lib.Name -split '/')[0]
			if ($id -like "Microsoft.NETCore.App.*" -or $id -like "Microsoft.WindowsDesktop.App.*" -or $id -like "Microsoft.AspNetCore.App.*") { continue }
			if ($id -like "Microsoft.NET.ILLink.*" -or $id -like "Microsoft.DotNet.ILCompiler*" -or $id -like "Microsoft.NET.Sdk.*") { continue }    # (build tools: not carried)
			if ($named -notmatch [regex]::Escape($id)) { $unnamed += @($lib.Name) }
		}
		if ($unnamed) { throw "THIRD-PARTY-NOTICES.md doesn't name these packages, which dAHKxodus.exe carries:`n$($unnamed -join "`n")" }

		# the privacy check over what is shipped: dAHKxodus's text files, and the built DLLs, the exe and the editor
		# package for this computer's user folder
		$userPart = [regex]::Escape($env:USERNAME)
		$texts = Get-ChildItem $folder -Recurse -File | Where-Object {
			$_.Extension -in ".ahk", ".cs", ".txt", ".md", ".json", ".html", ".css", ".ini", ".js" -and $_.FullName -notlike "*\node_modules\*"
		}
		$hits = @($texts | Select-String -Pattern $pattern -List)
		$built = @(Get-ChildItem $folder -File | Where-Object { $_.Name -in "dAHKxodus.dll", "dAHKxodus.exe", "Docxodus.dll", "DocxodusNativeBridge.dll" }) +
			@(Get-ChildItem (Join-Path $toPackage "dist") -Recurse -File -Include *.js, *.wasm, *.dll)
		foreach ($file in $built) {
			$bytes = [IO.File]::ReadAllBytes($file.FullName)
			foreach ($enc in [Text.Encoding]::Latin1, [Text.Encoding]::Unicode) {
				if ($enc.GetString($bytes) -match ('\\Users\\' + $userPart + '\\')) { $hits += "$($file.Name) has this computer's user folder in it" }
			}
		}
		if ($hits) { $hits | Out-Host; throw "The privacy check found something in what the $($ed.ToLower()) edition would carry." }

		# the installer
		"Making the $($ed.ToLower()) installer (Inno Setup) ..."
		$setupName = NameOf $ed "setup"
		$defines = @("/DVersion=$version", "/DSourceDir=$folder", "/DOutputDir=$releases", "/DRepo=$src")
		if (-not $full) { $defines += "/DLight" }
		& $iscc /Q @defines "/F$setupName" (Join-Path $src "installer\dAHKxodus.iss") | Out-Host
		if ($LASTEXITCODE -ne 0) { throw "Inno Setup failed." }

		# the portable zip: the same files in a folder dAHKxodus-<version>, with the "portable" file (AppFolders.Portable)
		"Making the $($ed.ToLower()) portable zip ..."
		New-Item -ItemType File -Force (Join-Path $folder "portable") | Out-Null
		$needs = if ($full) { @(
			"2. Double-click dAHKxodus.exe. Nothing needs installing: .NET 10 and AutoHotkey are inside. The editor runs in",
			"   Microsoft Edge WebView2, part of Windows 11; on Windows 10 without it, install the ""Evergreen"" WebView2 Runtime",
			"   from https://developer.microsoft.com/microsoft-edge/webview2/")
		} else { @(
			"2. Double-click dAHKxodus.exe. This is the light edition: it needs",
			"   - the .NET 10 Desktop Runtime (x64): https://dotnet.microsoft.com/download/dotnet/10.0",
			"   - AutoHotkey v2 for the Macros tab (everything else works without it): https://www.autohotkey.com/ or the",
			"     Microsoft Store. dAHKxodus finds an installed AutoHotkey by itself; one it doesn't find (another folder):",
			"     put the full path of its AutoHotkey64.exe as the only line of a text file AutoHotkey.path beside dAHKxodus.exe.",
			"   - Microsoft Edge WebView2, part of Windows 11; on Windows 10 without it, install the ""Evergreen"" WebView2",
			"     Runtime from https://developer.microsoft.com/microsoft-edge/webview2/")
		}
		$ownAhk = if ($full) { @(
			"AutoHotkey64.exe here is dAHKxodus's own private copy (see AutoHotkey64.exe.txt); an AutoHotkey you have installed",
			"is not touched. README.md tells about dAHKxodus; https://github.com/burque505/dAHKxodus is its home.")
		} else { @("README.md tells about dAHKxodus; https://github.com/burque505/dAHKxodus is its home.") }
		Set-Content (Join-Path $folder "README-portable.txt") -Encoding UTF8 (@(
			"dAHKxodus $version, portable$(if (-not $full) { ' (light edition)' })",
			"",
			"1. Unzip this folder anywhere you can write to (your Documents, a USB stick; not Program Files).",
			"   Tip: before unzipping, right-click the zip > Properties > Unblock, so Windows asks only once.") + $needs + @(
			"   - Windows may say ""Windows protected your PC"": dAHKxodus isn't signed yet. Click ""More info"", then ""Run anyway"".",
			"   - Smart App Control, if it is on, blocks unsigned programs and offers no way past it.",
			"",
			"The empty file named portable beside dAHKxodus.exe keeps everything dAHKxodus remembers (settings, recent files,",
			"recovery copies, the editor's WebView2 profile) in the data folder here. Nothing goes into %APPDATA% or",
			"%LOCALAPPDATA%, and deleting this folder removes dAHKxodus completely. Delete the portable file to share the",
			"settings in %APPDATA%\dAHKxodus with the installed dAHKxodus instead.",
			"",
			"A newer release: unzip it beside this one and move this folder's data folder into it.",
			"") + $ownAhk)
		$portableFolder = Join-Path $work "zip-$ed\dAHKxodus-$version"
		New-Item -ItemType Directory -Force (Split-Path $portableFolder) | Out-Null
		Move-Item $folder $portableFolder
		$portableZip = Join-Path $releases "$(NameOf $ed 'portable').zip"
		if (Test-Path $portableZip) { Remove-Item $portableZip }
		Add-Type -AssemblyName System.IO.Compression.FileSystem
		[IO.Compression.ZipFile]::CreateFromDirectory($portableFolder, $portableZip, [IO.Compression.CompressionLevel]::Optimal, $true)
		$made += Get-Item (Join-Path $releases "$setupName.exe"), (Get-Item $portableZip)
	}
} finally {
	if (Test-Path $work) { [IO.Directory]::Delete($work, $true) }
}
$hashes = [ordered]@{}
foreach ($file in $made) {
	$hashes[$file.Name] = (Get-FileHash $file.FullName -Algorithm SHA256).Hash
	"{0} ({1:N1} MB)" -f $file.FullName, ($file.Length / 1MB)
	"SHA-256 $($hashes[$file.Name])"
}

if (-not $Publish) { return }

# the GitHub pre-release "dAHKxodus <version> (installer preview)", tag installer-<version> at this commit (moved there
# when the installers of the same version are made again after a later commit)
$tag = "installer-$version"
$title = "dAHKxodus $version (installer preview)"
$commit = git -C $root rev-parse --short HEAD
$fullSetup, $fullZip, $lightSetup, $lightZip = $made | ForEach-Object Name
$notes = Join-Path ([IO.Path]::GetTempPath()) "dahkxodus-installer-notes.md"
Set-Content $notes -Encoding UTF8 (@(
	"**A preview of dAHKxodus with an installer**, built from commit ${commit}: Docxodus's Word-like editor in a Windows",
	"program, with AutoHotkey. Two editions, the same program:",
	"",
	"- **Light** (``$lightSetup``, the small one): uses the AutoHotkey v2 and the .NET 10 Desktop Runtime already on your",
	"  computer. If Setup doesn't find .NET, it offers to download and install it from Microsoft (about 60 MB; Windows",
	"  asks for administrator permission for it). If it doesn't find AutoHotkey, it lets you locate AutoHotkey64.exe or",
	"  opens the Microsoft Store; without AutoHotkey everything but the Macros tab works.",
	"- **Full** (``$fullSetup``): everything inside, nothing else to install. It brings its own private copy of",
	"  AutoHotkey64.exe, used only by dAHKxodus: an AutoHotkey you have installed is not touched.",
	"",
	"Either installs for your Windows user only (no administrator rights), in ``%LOCALAPPDATA%\Programs\dAHKxodus``, with a",
	"Start menu entry; Settings > Apps uninstalls it. Installing one edition over the other replaces it. Your settings,",
	"recent files and recovery copies stay in ``%APPDATA%\dAHKxodus`` and ``%LOCALAPPDATA%\dAHKxodus``, and are kept when it",
	"is uninstalled.",
	"",
	"**Portable instead?** ``$fullZip`` or ``$lightZip``: unzip it anywhere you can write to (a USB stick, your",
	"Documents) and run ``dAHKxodus.exe`` in it: nothing is installed. What it remembers stays in a ``data`` folder beside",
	"it, and deleting the folder removes it completely. Tip: right-click the zip > Properties > **Unblock** before",
	"unzipping. README-portable.txt inside tells more.",
	"",
	"**Microsoft Edge WebView2** runs the editor. Windows 11 has it; on Windows 10 without it, Setup says so and links to",
	"Microsoft's ""Evergreen"" WebView2 Runtime.",
	"",
	"**Not signed yet.** Windows will warn about it:",
	"- *Windows protected your PC* (SmartScreen): click **More info**, then **Run anyway**.",
	"- **Smart App Control**, if it is on, blocks unsigned programs and offers no way past it.",
	"- Some antivirus programs distrust new unsigned installers.",
	"",
	"To check the downloads (PowerShell: ``Get-FileHash <file>``), SHA-256:") +
	@($hashes.Keys | ForEach-Object { "- ``$_``: ``$($hashes[$_])``" }) + @(
	"",
	"Inside: dAHKxodus $version, Docxodus (fork commit $($lock.commit.Substring(0, 7))); the full edition also .NET $dotnetVersion (self-contained) and AutoHotkey $ahkVersion (GPL v2)."))
try {
	git -C $root tag -f $tag HEAD | Out-Null
	git -C $root push -f origin "refs/tags/$tag" 2>&1 | Out-Host
	$exists = $true
	gh release view $tag --repo $repo *> $null
	if ($LASTEXITCODE -ne 0) { $exists = $false }
	$files = $made | ForEach-Object FullName
	if ($exists) {
		gh release upload $tag @files --clobber --repo $repo | Out-Host
		gh release edit $tag --title $title --notes-file $notes --prerelease --repo $repo | Out-Host
	} else {
		gh release create $tag @files --title $title --notes-file $notes --prerelease --latest=false --repo $repo | Out-Host
	}
	if ($LASTEXITCODE -ne 0) { throw "gh failed." }
	"Published: $title (https://github.com/$repo/releases/tag/$tag)"
} finally {
	Remove-Item $notes -ErrorAction SilentlyContinue
}
