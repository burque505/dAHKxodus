# Checks a copy saved by a measure run (AutoHotkey64.exe dAHKxodus.ahk file.docx /measure out.docx) for what
# milestone 1 promises in the file itself:
#   - the comment the run added carries an author, initials and a date in Word's form (UTC, to the second);
#   - every picture is stored in word/media, and no relationship points at the package root's /media.
#
#   .\tests\check-saved.ps1 out.docx [-Author "name"]
# -Author: the name the comment must carry (the host's user; logs\dahkxodus.log says which it used).
# Exit code 0 when every check passes.
param(
    [Parameter(Mandatory)][string]$Path,
    [string]$Author
)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression.FileSystem
$failed = 0
function Check([string]$name, [bool]$ok, [string]$detail = "") {
    if (-not $ok) { $script:failed++ }
    Write-Host ("{0} {1}{2}" -f ($(if ($ok) { "ok    " } else { "FAIL  " }), $name, $(if ($ok -or $detail -eq "") { "" } else { ": $detail" })))
}

$zip = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path $Path).Path)
try {
    function Entry([string]$name) {
        $e = $zip.GetEntry($name)
        if (-not $e) { return $null }
        $r = [System.IO.StreamReader]::new($e.Open())
        try { return $r.ReadToEnd() } finally { $r.Dispose() }
    }
    $w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    $comments = Entry "word/comments.xml"
    Check "comments part" ($null -ne $comments)
    if ($comments) {
        $xml = [xml]$comments
        $ns = [System.Xml.XmlNamespaceManager]::new($xml.NameTable)
        $ns.AddNamespace("w", $w)
        $mine = $xml.SelectNodes("//w:comment", $ns) | Where-Object { $_.InnerText -like "*A comment from the measure run.*" } | Select-Object -First 1
        Check "the measure run's comment" ($null -ne $mine)
        if ($mine) {
            $a = $mine.GetAttribute("author", $w); $i = $mine.GetAttribute("initials", $w); $d = $mine.GetAttribute("date", $w)
            $expected = if ($Author) { ", expected '$Author'" } else { "" }
            Check "comment author" ($a -ne "" -and (-not $Author -or $a -eq $Author)) "author '$a'$expected"
            Check "comment initials" ($i -ne "") "no w:initials"
            Check "comment date, Word's form" ($d -match '^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ$') "w:date `"$d`""
        }
    }

    $pictures = @($zip.Entries | Where-Object { $_.FullName -match '\.(png|jpe?g|gif|bmp|tiff?|webp|emf|wmf|svg)$' } | ForEach-Object FullName)
    Check "pictures stored in word/media" ($pictures.Count -gt 0 -and -not ($pictures | Where-Object { $_ -notlike "word/media/*" })) ($pictures -join ", ")
    $rootTargets = @($zip.Entries | Where-Object { $_.FullName -like "*.rels" } | Where-Object { (Entry $_.FullName) -match 'Target="/media/' } | ForEach-Object FullName)
    Check "no relationship targets /media/" ($rootTargets.Count -eq 0) ($rootTargets -join ", ")
} finally { $zip.Dispose() }

Write-Host ""
Write-Host $(if ($failed -eq 0) { "PASSED" } else { "FAILED: $failed" })
exit $(if ($failed -eq 0) { 0 } else { 1 })
