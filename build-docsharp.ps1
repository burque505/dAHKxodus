# Builds the converters dAHKxodus opens older formats with, from the DocSharp fork commit named in docsharp.lock:
#   .doc (Word 97-2003): DocSharp.Binary.Doc and DocSharp.Binary.Common -> lib\docsharp
# The output is not kept in git: a clean clone runs this once (as build-engine.ps1 for the engine).
#
#   .\build-docsharp.ps1 [-Source <fork clone>] [-NoReadyToRun] [-UpdateLock]
#
# -Source: a clone of the fork (default ..\DocSharp), checked out at the lock's commit with no local changes.
# -UpdateLock: build whatever the clone has checked out (a newer commit) and write its commit into docsharp.lock.
param(
    [string]$Source = (Join-Path $PSScriptRoot "..\DocSharp"),
    [switch]$NoReadyToRun,
    [switch]$UpdateLock
)
$ErrorActionPreference = "Stop"
$lockPath = Join-Path $PSScriptRoot "docsharp.lock"
$lock = [ordered]@{}
foreach ($line in Get-Content $lockPath) {
    if ($line -match '^\s*([\w-]+)=(.*)$') { $lock[$Matches[1]] = $Matches[2].Trim() }
}

# The fork clone must be at the lock's commit, with nothing changed, so the build is the one the lock names.
$Source = (Resolve-Path $Source).Path
$head = (git -C $Source rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0) { throw "$Source is not a git clone of the DocSharp fork" }
$dirty = git -C $Source status --porcelain --untracked-files=no
if ($dirty) { throw "The fork clone $Source has uncommitted changes; commit or stash them first." }
if (-not $UpdateLock -and $head -ne $lock.commit) {
    throw "The fork clone is at $head, but docsharp.lock names $($lock.commit). Check that out (git -C `"$Source`" checkout $($lock.commit)), or build with -UpdateLock to move the lock."
}
Write-Host "DocSharp fork at $head"

# Only the net10.0 build, for win-x64 as dAHKxodus is; ReadyToRun so the first .doc opened doesn't wait for .NET to
# compile the converter.
$project = Join-Path $Source "src\DocSharp.Binary\DocSharp.Binary.Doc\DocSharp.Binary.Doc.csproj"
$out = Join-Path $Source "src\DocSharp.Binary\DocSharp.Binary.Doc\bin\Release\publish-dahkxodus"
$r2r = if ($NoReadyToRun) { "false" } else { "true" }
dotnet publish $project -c Release -f net10.0 -r win-x64 --self-contained false -p:PublishReadyToRun=$r2r -o $out --nologo -v quiet
if ($LASTEXITCODE -ne 0) { throw "DocSharp build failed" }
$lib = Join-Path $PSScriptRoot "lib\docsharp"
New-Item -ItemType Directory -Force $lib | Out-Null
foreach ($name in "DocSharp.Binary.Doc", "DocSharp.Binary.Common") {
    Copy-Item (Join-Path $out "$name.dll") $lib -Force
}
Write-Host "DocSharp -> lib\docsharp"

if ($UpdateLock) {
    $text = (Get-Content $lockPath) | ForEach-Object {
        if ($_ -match '^\s*commit=') { "commit=$head" } else { $_ }
    }
    Set-Content $lockPath $text -Encoding utf8NoBOM
    Write-Host "docsharp.lock now names $head"
}
