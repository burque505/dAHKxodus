# Builds the Docxodus engine dAHKxodus runs, from the fork commit named in docxodus.lock:
#   native: the .NET engine the host runs (DocxodusNativeBridge, in the fork's tools\native-bridge) -> lib\docxodus
#   npm:    the browser editor package (WASM + JS)                                        -> web\vendor\docxodus-0.0.0.tgz,
#           then installed into web\node_modules
# Neither output is kept in git: a clean clone runs this once.
#
#   .\build-engine.ps1 [-Part all|native|npm] [-Source <fork clone>] [-NoReadyToRun] [-UpdateLock]
#
# -Source: a clone of the fork (default ..\docxodus-src), checked out at the lock's commit with no local changes.
# -UpdateLock: build whatever the clone has checked out (a newer commit) and write its commit and the new package's
#   hash into docxodus.lock.
# The npm part needs the .NET wasm-tools workload, Node/npm and Git's bash; its WASM build takes several minutes.
param(
    [ValidateSet("all", "native", "npm")][string]$Part = "all",
    [string]$Source = (Join-Path $PSScriptRoot "..\docxodus-src"),
    [switch]$NoReadyToRun,
    [switch]$UpdateLock
)
$ErrorActionPreference = "Stop"
$lockPath = Join-Path $PSScriptRoot "docxodus.lock"
$lock = [ordered]@{}
foreach ($line in Get-Content $lockPath) {
    if ($line -match '^\s*([\w-]+)=(.*)$') { $lock[$Matches[1]] = $Matches[2].Trim() }
}

# The fork clone must be at the lock's commit, with nothing changed, so the build is the one the lock names.
$Source = (Resolve-Path $Source).Path
$head = (git -C $Source rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0) { throw "$Source is not a git clone of the Docxodus fork" }
$dirty = git -C $Source status --porcelain --untracked-files=no
if ($dirty) { throw "The fork clone $Source has uncommitted changes; commit or stash them first." }
if (-not $UpdateLock -and $head -ne $lock.commit) {
    throw "The fork clone is at $head, but docxodus.lock names $($lock.commit). Check that out (git -C `"$Source`" checkout $($lock.commit)), or build with -UpdateLock to move the lock."
}
Write-Host "Docxodus fork at $head"

if ($Part -in "all", "native") {
    # ReadyToRun: the DLLs carry precompiled code, so the first document opened doesn't wait for .NET to compile
    # Docxodus's code on first use.
    $project = Join-Path $Source "tools\native-bridge\DocxodusNativeBridge.csproj"
    $out = Join-Path $Source "tools\native-bridge\bin\Release\publish-win-x64"
    $r2r = if ($NoReadyToRun) { "false" } else { "true" }
    dotnet publish $project -c Release -r win-x64 --self-contained false -p:PublishReadyToRun=$r2r -o $out --nologo -v quiet
    if ($LASTEXITCODE -ne 0) { throw "native engine build failed" }
    $lib = Join-Path $PSScriptRoot "lib\docxodus"
    New-Item -ItemType Directory -Force $lib | Out-Null
    foreach ($name in "DocxodusNativeBridge", "Docxodus", "DocumentFormat.OpenXml", "DocumentFormat.OpenXml.Framework", "System.IO.Packaging") {
        Copy-Item (Join-Path $out "$name.dll") $lib -Force
    }
    Write-Host "native engine -> lib\docxodus"
}

if ($Part -in "all", "npm") {
    # Git's own bash (not WSL's in System32): npm's build:wasm script fails under cmd on Windows.
    $bash = "C:\Program Files\Git\bin\bash.exe"
    if (-not (Test-Path $bash)) { $bash = "C:\Program Files\Git\usr\bin\bash.exe" }
    if (-not (Test-Path $bash)) { throw "Git's bash.exe not found (needed for the WASM build)" }
    Push-Location (Join-Path $Source "npm")
    try {
        if (-not (Test-Path "node_modules")) { npm ci; if ($LASTEXITCODE -ne 0) { throw "npm ci failed in the fork" } }
        & $bash ../scripts/build-wasm.sh
        if ($LASTEXITCODE -ne 0) { throw "WASM build failed" }
        npm run build:js
        if ($LASTEXITCODE -ne 0) { throw "npm build:js failed" }
        $vendor = Join-Path $PSScriptRoot "web\vendor"
        New-Item -ItemType Directory -Force $vendor | Out-Null
        npm pack --pack-destination $vendor
        if ($LASTEXITCODE -ne 0) { throw "npm pack failed" }
    } finally { Pop-Location }
    $tgz = Join-Path $PSScriptRoot "web\vendor\docxodus-0.0.0.tgz"
    $hash = (Get-FileHash $tgz -Algorithm SHA256).Hash.ToLowerInvariant()
    if (-not $UpdateLock -and $hash -ne $lock["npm-sha256"]) {
        Write-Warning "The package's hash ($hash) differs from docxodus.lock's; npm install will change web\package-lock.json. If that is expected, rebuild with -UpdateLock."
    }
    # Install the package just packed, not whatever node_modules holds: a plain "npm install" finds the same name and
    # version (0.0.0) already there and keeps the old build. Naming the file also brings package-lock.json's integrity
    # up to date when the package changed (and leaves it alone when it did not).
    Push-Location (Join-Path $PSScriptRoot "web")
    try {
        if (Test-Path "node_modules\docxodus") { Remove-Item "node_modules\docxodus" -Recurse -Force }
        npm install ./vendor/docxodus-0.0.0.tgz
        if ($LASTEXITCODE -ne 0) { throw "npm install failed in web" }
    } finally { Pop-Location }
    Write-Host "editor package -> web\vendor, installed in web\node_modules"
    if ($UpdateLock) { $lock["npm-sha256"] = $hash }
}

if ($UpdateLock) {
    $lock.commit = $head
    $text = (Get-Content $lockPath) | ForEach-Object {
        if ($_ -match '^\s*([\w-]+)=') { "$($Matches[1])=$($lock[$Matches[1]])" } else { $_ }
    }
    Set-Content $lockPath $text -Encoding utf8NoBOM
    Write-Host "docxodus.lock now names $head"
}
