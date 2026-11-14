# Runs dAHKxodus's measure mode on a desktop of its own, so its window never appears on yours (or takes the focus):
#   .\tests\run-measure.ps1 [file.docx] [-Out out.docx] [-AutoHotkey path] [-TimeoutSeconds 150] [-Exe [path]]
# -Exe: run dAHKxodus.exe (default vs\publish\dAHKxodus.exe, built by build-exe.ps1) instead of dAHKxodus.ahk; it
# starts dAHKxodus.ahk /companion with -AutoHotkey for the bridge.
# It opens the file (default samples\verify-mix.docx), runs the page's checks, saves the edited copy to -Out (default
# logs\measure-out.docx) and closes; the results are the run's lines in logs\dahkxodus.log, which this prints.
# Follow it with tests\check-saved.ps1 on the copy.
param(
    [string]$File = (Join-Path $PSScriptRoot "..\samples\verify-mix.docx"),
    [string]$Out = (Join-Path $PSScriptRoot "..\logs\measure-out.docx"),
    [string]$AutoHotkey = "C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe",
    [int]$TimeoutSeconds = 150,
    [switch]$Exe,
    [string]$ExePath = (Join-Path $PSScriptRoot "..\vs\publish\dAHKxodus.exe")
)
$ErrorActionPreference = "Stop"
$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$File = (Resolve-Path $File).Path
$Out = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Out)    # (relative to this shell's folder)
New-Item -ItemType Directory -Force (Split-Path $Out) | Out-Null
if (Test-Path $Out) { Remove-Item $Out }
$log = Join-Path $root "logs\dahkxodus.log"
$logLines = if (Test-Path $log) { (Get-Content $log).Count } else { 0 }

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class HiddenDesktop {
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    struct STARTUPINFO {
        public int cb; public string lpReserved; public string lpDesktop; public string lpTitle;
        public int dwX, dwY, dwXSize, dwYSize, dwXCountChars, dwYCountChars, dwFillAttribute, dwFlags;
        public short wShowWindow, cbReserved2; public IntPtr lpReserved2, hStdInput, hStdOutput, hStdError;
    }
    [StructLayout(LayoutKind.Sequential)]
    struct PROCESS_INFORMATION { public IntPtr hProcess, hThread; public int dwProcessId, dwThreadId; }
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern IntPtr CreateDesktopW(string name, IntPtr device, IntPtr devmode, int flags, uint access, IntPtr sa);
    [DllImport("user32.dll")] static extern bool CloseDesktop(IntPtr desktop);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool CreateProcessW(string app, string cmd, IntPtr pa, IntPtr ta, bool inherit, int flags, IntPtr env,
        string dir, ref STARTUPINFO si, out PROCESS_INFORMATION pi);
    [DllImport("kernel32.dll")] static extern uint WaitForSingleObject(IntPtr h, uint ms);
    [DllImport("kernel32.dll")] static extern bool GetExitCodeProcess(IntPtr h, out uint code);
    [DllImport("kernel32.dll")] static extern bool TerminateProcess(IntPtr h, uint code);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);

    // Returns the process's exit code, or -1 when it was stopped after the timeout.
    public static long Run(string commandLine, string dir, int timeoutMs) {
        IntPtr desktop = CreateDesktopW("dAHKxodus-measure", IntPtr.Zero, IntPtr.Zero, 0, 0x10000000 /* GENERIC_ALL */, IntPtr.Zero);
        if (desktop == IntPtr.Zero) throw new Exception("CreateDesktop failed: " + Marshal.GetLastWin32Error());
        try {
            var si = new STARTUPINFO { cb = Marshal.SizeOf<STARTUPINFO>(), lpDesktop = "dAHKxodus-measure" };
            if (!CreateProcessW(null, commandLine, IntPtr.Zero, IntPtr.Zero, false, 0, IntPtr.Zero, dir, ref si, out var pi))
                throw new Exception("CreateProcess failed: " + Marshal.GetLastWin32Error());
            try {
                if (WaitForSingleObject(pi.hProcess, (uint)timeoutMs) != 0) { TerminateProcess(pi.hProcess, 1); return -1; }
                GetExitCodeProcess(pi.hProcess, out var code);
                return code;
            } finally { CloseHandle(pi.hProcess); CloseHandle(pi.hThread); }
        } finally { CloseDesktop(desktop); }
    }
}
"@

if ($Exe) {
    $ExePath = (Resolve-Path $ExePath).Path
    $env:DAHKXODUS_AUTOHOTKEY = $AutoHotkey    # (inherited: the exe starts its companion with it)
    $cmd = "`"$ExePath`" `"$File`" /measure `"$Out`""
} else {
    $cmd = "`"$AutoHotkey`" `"$root\dAHKxodus.ahk`" `"$File`" /measure `"$Out`""
}
$code = [HiddenDesktop]::Run($cmd, $root, $TimeoutSeconds * 1000)
if (Test-Path $log) { Get-Content $log | Select-Object -Skip $logLines }
if ($code -eq -1) { Write-Host "measure: stopped after $TimeoutSeconds s"; exit 1 }
if (-not (Test-Path $Out)) { Write-Host "measure: no copy saved (exit code $code)"; exit 1 }
Write-Host "measure: saved $Out"
