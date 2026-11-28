# Builds dAHKxodus.exe (vs\dAHKxodus.csproj) from the same sources dAHKxodus.ahk runs through Csharp-in-AHK:
#   .\build-exe.ps1 [-Configuration Release] [-NoReadyToRun]
# A publish to vs\publish, ReadyToRun (precompiled, so the window comes up without waiting for .NET to compile it).
# The exe finds web\ and dAHKxodus.ahk in the repository folder above it. Needs lib\ (build-engine.ps1) first.
param(
    [string]$Configuration = "Release",
    [switch]$NoReadyToRun
)
$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
if (-not (Test-Path (Join-Path $root "lib\docxodus\Docxodus.dll"))) { throw "lib\docxodus is missing: run .\build-engine.ps1 first" }
if (-not (Test-Path (Join-Path $root "lib\docsharp\DocSharp.Binary.Doc.dll"))) { throw "lib\docsharp is missing: run .\build-docsharp.ps1 first" }
$out = Join-Path $root "vs\publish"
$r2r = if ($NoReadyToRun) { "false" } else { "true" }
dotnet publish (Join-Path $root "vs\dAHKxodus.csproj") -c $Configuration -r win-x64 --self-contained false `
    -p:PublishReadyToRun=$r2r -o $out --nologo -v quiet
if ($LASTEXITCODE -ne 0) { throw "dotnet publish failed ($LASTEXITCODE)" }
Write-Host "built $out\dAHKxodus.exe"
