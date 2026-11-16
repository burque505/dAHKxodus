# Checks lang\*.json against lang\en.json: every file has the same keys, and each text the same {0}-style
# placeholders as the English one (a translation that drops {1} would lose the error message it shows). The ribbon's
# own tables (lang\*.ribbon.json, keyed by its English) have tests\check-ribbon-lang.mjs.
#   .\tests\check-lang.ps1
$ErrorActionPreference = "Stop"
$folder = Join-Path $PSScriptRoot "..\lang"
$read = { param($path) Get-Content -Raw -Encoding utf8 $path | ConvertFrom-Json -AsHashtable }
$en = & $read (Join-Path $folder "en.json")
$placeholders = { param($text) ([regex]::Matches($text, "\{\d+\}") | ForEach-Object Value | Sort-Object -Unique) -join " " }
$failed = 0
# Every file must parse as the host reads it (System.Text.Json, strict): ConvertFrom-Json also takes what the host
# refuses, such as a line break inside a string, and the host then shows the keys instead of the texts.
foreach ($file in Get-ChildItem $folder -Filter *.json) {
    try { [System.Text.Json.JsonDocument]::Parse([System.IO.File]::ReadAllText($file.FullName)).Dispose() }
    catch { Write-Host "FAILED $($file.Name): not valid JSON: $($_.Exception.InnerException.Message ?? $_.Exception.Message)"; $failed++ }
}
foreach ($file in Get-ChildItem $folder -Filter *.json | Where-Object { $_.Name -ne "en.json" -and $_.Name -notlike "*.ribbon.json" }) {
    $other = & $read $file.FullName
    foreach ($key in $en.Keys) {
        if (-not $other.ContainsKey($key)) { Write-Host "FAILED $($file.Name): missing $key"; $failed++; continue }
        if ((& $placeholders $en[$key]) -ne (& $placeholders $other[$key])) { Write-Host "FAILED $($file.Name): $key has other placeholders"; $failed++ }
    }
    foreach ($key in $other.Keys) { if (-not $en.ContainsKey($key)) { Write-Host "FAILED $($file.Name): $key is not in en.json"; $failed++ } }
    Write-Host "ok     $($file.Name): $($other.Count) texts"
}
if ($failed) { Write-Host "`n$failed FAILED"; exit 1 }
Write-Host "`nPASSED"
