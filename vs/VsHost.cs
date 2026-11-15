// What dAHKxodus.ahk and Csharp-in-AHK give dAHKxodus.cs, and dAHKxodus.exe has to give itself.
// dAHKxodus.exe [file.docx] [/measure out.docx]: the arguments are dAHKxodus.ahk's own; the exe completes them to the
// ones dAHKxodus.ahk passes ("0" for the AHK window, the program folder, the document, the measure copy), and starts
// "AutoHotkey64.exe dAHKxodus.ahk /companion <window>" for the bridge once its window exists.
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

static class VsHost
{
	[ModuleInitializer]
	internal static void Start()
	{
		AppFolders.ProgramExe = Environment.ProcessPath;
		AppFolders.CompleteArgs = Complete;
		AppFolders.StartCompanion = StartCompanion;
		AppDomain.CurrentDomain.UnhandledException += (_, e) => {
			try { MainForm.Log("fatal: " + e.ExceptionObject); } catch { }
			if (measure) return;    // (a measure run is on a desktop of its own: nobody would see the message)
			try { MessageBox.Show("dAHKxodus stopped because of an error:\n\n" + e.ExceptionObject, "dAHKxodus", MessageBoxButtons.OK, MessageBoxIcon.Error); }
			catch { }
		};
	}

	static bool measure;

	/// [file.docx] [/measure out.docx] -> "0", program folder, file or "", measure copy or "" (as dAHKxodus.ahk reads them).
	static string[] Complete(string[] args)
	{
		string file = "", measureOut = "";
		for (int i = 0; i < args.Length; i++) {
			if (args[i].Equals("/measure", StringComparison.OrdinalIgnoreCase)) {
				measureOut = i + 1 < args.Length ? Path.GetFullPath(args[++i]) : "";
				measure = true;
			}
			else if (file == "") file = Path.GetFullPath(args[i]);
		}
		return ["0", ProgramFolder(), file, measureOut];
	}

	/// The folder with dAHKxodus's files (web\index.html, dAHKxodus.ahk): the exe's own, or the nearest one above it
	/// (the development exe is in vs\bin\... or vs\publish), else the current folder.
	static string ProgramFolder()
	{
		for (var dir = new DirectoryInfo(AppContext.BaseDirectory); dir != null; dir = dir.Parent)
			if (File.Exists(Path.Combine(dir.FullName, "web", "index.html")) && File.Exists(Path.Combine(dir.FullName, "dAHKxodus.ahk")))
				return dir.FullName;
		return Environment.CurrentDirectory;
	}

	/// AutoHotkey v2 (64-bit): DAHKXODUS_AUTOHOTKEY, else the one beside dAHKxodus.exe (the full installer's copy), else
	/// the one named in AutoHotkey.path beside it (the light installer's: the AutoHotkey64.exe the user located, one
	/// line, UTF-8), else where its own installer registers it, else its usual place, else on PATH (a Microsoft Store
	/// AutoHotkey's command alias); "" if not found.
	static string FindAutoHotkey()
	{
		var named = Environment.GetEnvironmentVariable("DAHKXODUS_AUTOHOTKEY");
		if (!string.IsNullOrEmpty(named) && File.Exists(named)) return named;
		string beside = Path.Combine(AppContext.BaseDirectory, "AutoHotkey64.exe");
		if (File.Exists(beside)) return beside;
		try {
			string pathFile = Path.Combine(AppContext.BaseDirectory, "AutoHotkey.path");
			if (File.Exists(pathFile)) {
				string located = File.ReadAllText(pathFile, System.Text.Encoding.UTF8).Trim().Trim('"');
				if (located != "" && File.Exists(located)) return located;
			}
		} catch (Exception) { }    // (unreadable: the other places)
		foreach (var hive in new[] { Microsoft.Win32.Registry.CurrentUser, Microsoft.Win32.Registry.LocalMachine }) {
			using var key = hive.OpenSubKey(@"SOFTWARE\AutoHotkey");
			if (key?.GetValue("InstallDir") is string dir && File.Exists(Path.Combine(dir, "v2", "AutoHotkey64.exe")))
				return Path.Combine(dir, "v2", "AutoHotkey64.exe");
		}
		string usual = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "AutoHotkey", "v2", "AutoHotkey64.exe");
		if (File.Exists(usual)) return usual;
		foreach (var dir in (Environment.GetEnvironmentVariable("PATH") ?? "").Split(';', StringSplitOptions.RemoveEmptyEntries)) {
			try {
				string onPath = Path.Combine(dir.Trim().Trim('"'), "AutoHotkey64.exe");
				if (File.Exists(onPath)) return onPath;
			} catch (Exception) { }    // (a malformed PATH entry)
		}
		return "";
	}

	/// The bridge's AutoHotkey side. dAHKxodus works without it (the bridge's calls then answer with an error).
	static void StartCompanion(IntPtr window)
	{
		try {
			string folder = AppFolders.ProgramFolder, ahk = FindAutoHotkey(), script = Path.Combine(folder, "dAHKxodus.ahk");
			if (ahk == "" || !File.Exists(script)) { MainForm.Log("companion: AutoHotkey v2 or dAHKxodus.ahk not found"); return; }
			var start = new ProcessStartInfo(ahk) { UseShellExecute = false, WorkingDirectory = folder };
			start.ArgumentList.Add(script);
			start.ArgumentList.Add("/companion");
			start.ArgumentList.Add(((long)window).ToString());
			using var _ = Process.Start(start);
		}
		catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or IOException or InvalidOperationException) {
			MainForm.Log("companion could not be started: " + ex.Message);
		}
	}
}
