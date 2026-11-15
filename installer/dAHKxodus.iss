; dAHKxodus's installer (Inno Setup 6), made by installer\make_installer.ps1, which passes
;   /DVersion=0.0.N.0 /DSourceDir=<the staged program folder> /DOutputDir=<..\dAHKxodus-releases> /DRepo=<the sources>
; Installed for the current Windows user only (no administrator, no UAC prompt), in %LOCALAPPDATA%\Programs\dAHKxodus:
; dAHKxodus.exe (self-contained: .NET inside), the editor (web\), the languages (lang\), and AutoHotkey64.exe with
; dAHKxodus.ahk for the AutoHotkey bridge. Settings, recent files and recovery copies stay in %APPDATA%\dAHKxodus and
; %LOCALAPPDATA%\dAHKxodus: uninstalling leaves them.
; The editor runs in Microsoft Edge WebView2, part of Windows 11 (and of most Windows 10 computers by now): when it is
; missing, the installer says where to get it.
; /DLight: the light installer (make_installer.ps1 -Light): dAHKxodus.exe needs the .NET 10 Desktop Runtime installed,
; and no AutoHotkey comes with it. Setup looks for both: it offers to download and install .NET (Microsoft's installer,
; which asks for administrator permission itself), and to locate AutoHotkey64.exe (saved in AutoHotkey.path beside
; dAHKxodus.exe) or get it from the Microsoft Store. Same AppId and folder as the full installer: either replaces the other.

#ifndef Version
  #error Run installer\make_installer.ps1 (it passes /DVersion, /DSourceDir, /DOutputDir and /DRepo).
#endif

[Setup]
; (the AppId must never change: Windows knows the installed dAHKxodus by it, and a newer installer replaces it)
AppId={{B5280039-0821-4130-A6FC-60E4F901D8FE}
AppName=dAHKxodus
AppVersion={#Version}
#ifdef Light
AppVerName=dAHKxodus {#Version} (light installer preview)
#else
AppVerName=dAHKxodus {#Version} (installer preview)
#endif
AppPublisher=burque505
AppPublisherURL=https://github.com/burque505/dAHKxodus
AppSupportURL=https://github.com/burque505/dAHKxodus/issues
AppUpdatesURL=https://github.com/burque505/dAHKxodus/releases
VersionInfoVersion={#Version}
PrivilegesRequired=lowest
DefaultDirName={autopf}\dAHKxodus
DefaultGroupName=dAHKxodus
DisableProgramGroupPage=yes
DisableWelcomePage=no
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0.17763
LicenseFile={#Repo}\LICENSE
UninstallDisplayIcon={app}\dAHKxodus.exe
OutputDir={#OutputDir}
#ifdef Light
OutputBaseFilename=dAHKxodus-{#Version}-light-setup
#else
OutputBaseFilename=dAHKxodus-{#Version}-setup
#endif
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
RestartApplications=no
ChangesAssociations=yes

[Languages]
Name: "en"; MessagesFile: "compiler:Default.isl"
Name: "es"; MessagesFile: "compiler:Languages\Spanish.isl"

[Messages]
#ifdef Light
; (the welcome page: what the light installer needs)
en.WelcomeLabel2=This will install [name/ver] on your computer.%n%nThis is the light installer: dAHKxodus uses the AutoHotkey v2 and the .NET 10 Desktop Runtime already on your computer. If Setup can't find them, it will offer to get them for you.%n%nIt is recommended that you close all other applications before continuing.
es.WelcomeLabel2=Este programa instalará [name/ver] en su equipo.%n%nEste es el instalador ligero: dAHKxodus usa el AutoHotkey v2 y el .NET 10 Desktop Runtime que ya tenga su equipo. Si la instalación no los encuentra, le ofrecerá conseguirlos.%n%nSe recomienda cerrar todas las demás aplicaciones antes de continuar.
#else
; (the welcome page: an AutoHotkey already installed stays as it is)
en.WelcomeLabel2=This will install [name/ver] on your computer.%n%nIf you already have AutoHotkey, it is left alone: dAHKxodus brings its own private copy of AutoHotkey, used only by dAHKxodus. Your AutoHotkey installation, its version and your scripts are not changed.%n%nIt is recommended that you close all other applications before continuing.
es.WelcomeLabel2=Este programa instalará [name/ver] en su equipo.%n%nSi ya tiene AutoHotkey, no se modifica: dAHKxodus trae su propia copia privada de AutoHotkey, que solo usa dAHKxodus. Su instalación de AutoHotkey, su versión y sus scripts no cambian.%n%nSe recomienda cerrar todas las demás aplicaciones antes de continuar.
#endif

[CustomMessages]
en.OpenWith=Offer dAHKxodus in "Open with" for .docx files (doesn't change your default program)
es.OpenWith=Ofrecer dAHKxodus en "Abrir con" para archivos .docx (no cambia su programa predeterminado)
en.RunNow=Start dAHKxodus
es.RunNow=Iniciar dAHKxodus
en.NoWebView2=dAHKxodus's editor runs in Microsoft Edge WebView2, which this computer doesn't seem to have.%n%nSetup will go on, but before dAHKxodus can start, install the "Evergreen" WebView2 Runtime from Microsoft:%nhttps://developer.microsoft.com/microsoft-edge/webview2/%n%nOpen that page now?
es.NoWebView2=El editor de dAHKxodus funciona en Microsoft Edge WebView2, que este equipo no parece tener.%n%nLa instalación continuará, pero antes de poder iniciar dAHKxodus instale el WebView2 Runtime "Evergreen" de Microsoft:%nhttps://developer.microsoft.com/microsoft-edge/webview2/%n%n¿Abrir esa página ahora?

#ifdef Light
en.AhkCaption=AutoHotkey
es.AhkCaption=AutoHotkey
en.AhkDescription=dAHKxodus's macros run in AutoHotkey v2
es.AhkDescription=Las macros de dAHKxodus se ejecutan en AutoHotkey v2
en.AhkText=Setup didn't find AutoHotkey v2 (AutoHotkey64.exe) on this computer. dAHKxodus works without it, except for the Macros tab.
es.AhkText=La instalación no encontró AutoHotkey v2 (AutoHotkey64.exe) en este equipo. dAHKxodus funciona sin él, salvo la pestaña Macros.
en.AhkLocate=I have it: let me locate AutoHotkey64.exe
es.AhkLocate=Lo tengo: dejarme buscar AutoHotkey64.exe
en.AhkStore=Get AutoHotkey from the Microsoft Store (opens the Store; dAHKxodus finds it once it is installed)
es.AhkStore=Conseguir AutoHotkey en Microsoft Store (abre la Store; dAHKxodus lo encuentra una vez instalado)
en.AhkSkip=Continue without AutoHotkey
es.AhkSkip=Continuar sin AutoHotkey
en.AhkBrowse=Where is AutoHotkey64.exe?
es.AhkBrowse=¿Dónde está AutoHotkey64.exe?
en.AhkNotFound=That file doesn't exist. Choose AutoHotkey64.exe, or another option.
es.AhkNotFound=Ese archivo no existe. Elija AutoHotkey64.exe u otra opción.
en.AhkFound=AutoHotkey: %1
es.AhkFound=AutoHotkey: %1
en.DotNetCaption=.NET 10 Desktop Runtime
es.DotNetCaption=.NET 10 Desktop Runtime
en.DotNetDescription=dAHKxodus needs Microsoft's .NET 10 Desktop Runtime (x64)
es.DotNetDescription=dAHKxodus necesita el .NET 10 Desktop Runtime (x64) de Microsoft
en.DotNetText=Setup didn't find the .NET 10 Desktop Runtime on this computer. dAHKxodus won't start without it.
es.DotNetText=La instalación no encontró el .NET 10 Desktop Runtime en este equipo. dAHKxodus no se inicia sin él.
en.DotNetGet=Download and install it now (about 60 MB, from Microsoft; Windows will ask for administrator permission)
es.DotNetGet=Descargarlo e instalarlo ahora (unos 60 MB, de Microsoft; Windows pedirá permiso de administrador)
en.DotNetSkip=Skip: I'll install it myself (https://dotnet.microsoft.com/download/dotnet/10.0)
es.DotNetSkip=Omitir: lo instalaré yo (https://dotnet.microsoft.com/download/dotnet/10.0)
en.DotNetFailed=The .NET 10 Desktop Runtime doesn't seem to be installed yet (its installer ended with code %1). Setup will go on; install it from https://dotnet.microsoft.com/download/dotnet/10.0 before starting dAHKxodus.
es.DotNetFailed=El .NET 10 Desktop Runtime no parece estar instalado aún (su instalador terminó con el código %1). La instalación continuará; instálelo desde https://dotnet.microsoft.com/download/dotnet/10.0 antes de iniciar dAHKxodus.
#endif

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "openwith"; Description: "{cm:OpenWith}"; Flags: unchecked

[InstallDelete]
; (an older dAHKxodus's libraries, .NET and editor: whatever this one doesn't carry any more mustn't stay)
Type: files; Name: "{app}\*.dll"
Type: files; Name: "{app}\*.deps.json"
Type: files; Name: "{app}\*.runtimeconfig.json"
Type: filesandordirs; Name: "{app}\runtimes"
Type: filesandordirs; Name: "{app}\web"

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\dAHKxodus"; Filename: "{app}\dAHKxodus.exe"; WorkingDir: "{app}"
Name: "{autodesktop}\dAHKxodus"; Filename: "{app}\dAHKxodus.exe"; WorkingDir: "{app}"; Tasks: desktopicon

[Registry]
; "Open with" only (HKCU, this user): Windows offers dAHKxodus for .docx, the default program stays as it is
Root: HKCU; Subkey: "Software\Classes\Applications\dAHKxodus.exe"; ValueType: string; ValueName: "FriendlyAppName"; ValueData: "dAHKxodus"; Flags: uninsdeletekey; Tasks: openwith
Root: HKCU; Subkey: "Software\Classes\Applications\dAHKxodus.exe\DefaultIcon"; ValueType: string; ValueData: """{app}\dAHKxodus.exe"",0"; Tasks: openwith
Root: HKCU; Subkey: "Software\Classes\Applications\dAHKxodus.exe\shell\open\command"; ValueType: string; ValueData: """{app}\dAHKxodus.exe"" ""%1"""; Tasks: openwith
Root: HKCU; Subkey: "Software\Classes\Applications\dAHKxodus.exe\SupportedTypes"; ValueType: string; ValueName: ".docx"; ValueData: ""; Tasks: openwith
Root: HKCU; Subkey: "Software\Classes\.docx\OpenWithList\dAHKxodus.exe"; Flags: uninsdeletekey; Tasks: openwith

[Run]
#ifdef Light
; (not offered while .NET is missing: the exe would only say so)
Filename: "{app}\dAHKxodus.exe"; Description: "{cm:RunNow}"; WorkingDir: "{app}"; Flags: nowait postinstall skipifsilent; Check: HasDotNet10
#else
Filename: "{app}\dAHKxodus.exe"; Description: "{cm:RunNow}"; WorkingDir: "{app}"; Flags: nowait postinstall skipifsilent
#endif

[UninstallDelete]
; (written while dAHKxodus runs, not installed: its log)
Type: filesandordirs; Name: "{app}\logs"
#ifdef Light
; (written by Setup, not installed: the AutoHotkey64.exe the user located)
Type: files; Name: "{app}\AutoHotkey.path"
#endif

[Code]
// The WebView2 Runtime, as Microsoft's documentation tells installers to look for it: its "pv" (version) under
// EdgeUpdate's client key, for all users (64-bit Windows: the WOW6432Node view) or for this user.
const
  WebView2Key = 'Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}';

function HasWebView2(): Boolean;
var
  Version: String;
begin
  Result :=
    (RegQueryStringValue(HKLM, 'SOFTWARE\WOW6432Node\' + WebView2Key, 'pv', Version) and (Version <> '') and (Version <> '0.0.0.0')) or
    (RegQueryStringValue(HKCU, 'Software\' + WebView2Key, 'pv', Version) and (Version <> '') and (Version <> '0.0.0.0'));
end;

#ifdef Light
// ---- The light installer: AutoHotkey and .NET, found or got

const
  // Microsoft's permanent link to the latest .NET 10 Desktop Runtime installer for x64
  DotNetUrl = 'https://aka.ms/dotnet/10.0/windowsdesktop-runtime-win-x64.exe';

var
  AhkPage: TInputOptionWizardPage;
  DotNetPage: TInputOptionWizardPage;
  DownloadPage: TDownloadWizardPage;
  LocatedAhk: String;
  AhkChosen: Boolean;    // (the user located it: saved in AutoHotkey.path)

// The .NET 10 Desktop Runtime (x64): a 10.x folder of Microsoft.WindowsDesktop.App in the machine-wide dotnet.
function HasDotNet10(): Boolean;
var
  FindRec: TFindRec;
begin
  Result := False;
  if FindFirst(ExpandConstant('{commonpf64}\dotnet\shared\Microsoft.WindowsDesktop.App\10.*'), FindRec) then
    try
      repeat
        if (FindRec.Attributes and FILE_ATTRIBUTE_DIRECTORY) <> 0 then Result := True;
      until Result or not FindNext(FindRec);
    finally
      FindClose(FindRec);
    end;
end;

// AutoHotkey v2 where dAHKxodus.exe looks for it (VsHost.FindAutoHotkey): AutoHotkey.path from an earlier install,
// where AutoHotkey's installer registers it, its usual place, PATH (a Store AutoHotkey's alias); '' if not found.
function FindAutoHotkey(): String;
var
  Dir: String;
  Lines: TArrayOfString;
begin
  Result := '';
  if FileExists(ExpandConstant('{app}\AutoHotkey64.exe')) then begin Result := ExpandConstant('{app}\AutoHotkey64.exe'); exit; end;
  if LoadStringsFromFile(ExpandConstant('{app}\AutoHotkey.path'), Lines) then
    if (GetArrayLength(Lines) > 0) and FileExists(Trim(Lines[0])) then begin Result := Trim(Lines[0]); exit; end;
  if RegQueryStringValue(HKCU, 'SOFTWARE\AutoHotkey', 'InstallDir', Dir) then
    if FileExists(AddBackslash(Dir) + 'v2\AutoHotkey64.exe') then begin Result := AddBackslash(Dir) + 'v2\AutoHotkey64.exe'; exit; end;
  if RegQueryStringValue(HKLM64, 'SOFTWARE\AutoHotkey', 'InstallDir', Dir) then
    if FileExists(AddBackslash(Dir) + 'v2\AutoHotkey64.exe') then begin Result := AddBackslash(Dir) + 'v2\AutoHotkey64.exe'; exit; end;
  if FileExists(ExpandConstant('{commonpf64}\AutoHotkey\v2\AutoHotkey64.exe')) then
    begin Result := ExpandConstant('{commonpf64}\AutoHotkey\v2\AutoHotkey64.exe'); exit; end;
  Result := FileSearch('AutoHotkey64.exe', GetEnv('PATH'));
end;

function OnDownloadProgress(const Url, FileName: String; const Progress, ProgressMax: Int64): Boolean;
begin
  Result := True;
end;

procedure InitializeWizard();
begin
  AhkPage := CreateInputOptionPage(wpSelectTasks, CustomMessage('AhkCaption'), CustomMessage('AhkDescription'),
    CustomMessage('AhkText'), True, False);
  AhkPage.Add(CustomMessage('AhkLocate'));
  AhkPage.Add(CustomMessage('AhkStore'));
  AhkPage.Add(CustomMessage('AhkSkip'));
  AhkPage.SelectedValueIndex := 0;
  DotNetPage := CreateInputOptionPage(AhkPage.ID, CustomMessage('DotNetCaption'), CustomMessage('DotNetDescription'),
    CustomMessage('DotNetText'), True, False);
  DotNetPage.Add(CustomMessage('DotNetGet'));
  DotNetPage.Add(CustomMessage('DotNetSkip'));
  DotNetPage.SelectedValueIndex := 0;
  DownloadPage := CreateDownloadPage(SetupMessage(msgWizardPreparing), SetupMessage(msgPreparingDesc), @OnDownloadProgress);
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  Result := False;
  if PageID = AhkPage.ID then begin
    if LocatedAhk = '' then LocatedAhk := FindAutoHotkey();
    Result := LocatedAhk <> '';
  end else if PageID = DotNetPage.ID then
    Result := HasDotNet10();
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  FileName: String;
  Code: Integer;
begin
  Result := True;
  if CurPageID = AhkPage.ID then begin
    if AhkPage.SelectedValueIndex = 0 then begin
      FileName := 'AutoHotkey64.exe';
      if not GetOpenFileName(CustomMessage('AhkBrowse'), FileName, ExpandConstant('{commonpf64}'),
        'AutoHotkey64.exe|AutoHotkey64.exe|*.exe|*.exe', 'exe') then begin Result := False; exit; end;
      if not FileExists(FileName) then begin
        MsgBox(CustomMessage('AhkNotFound'), mbError, MB_OK);
        Result := False;
        exit;
      end;
      LocatedAhk := FileName;
      AhkChosen := True;
    end else if AhkPage.SelectedValueIndex = 1 then
      ShellExec('open', 'ms-windows-store://search/?query=AutoHotkey', '', '', SW_SHOWNORMAL, ewNoWait, Code);
  end else if (CurPageID = wpReady) and not HasDotNet10() and (DotNetPage.SelectedValueIndex = 0) then begin
    DownloadPage.Clear;
    DownloadPage.Add(DotNetUrl, 'windowsdesktop-runtime-10-win-x64.exe', '');
    DownloadPage.Show;
    try
      try
        DownloadPage.Download;
        // (ShellExec, not Exec: Microsoft's installer asks for administrator permission itself)
        ShellExec('', ExpandConstant('{tmp}\windowsdesktop-runtime-10-win-x64.exe'), '/install /passive /norestart', '',
          SW_SHOWNORMAL, ewWaitUntilTerminated, Code);
        if not HasDotNet10() then MsgBox(FmtMessage(CustomMessage('DotNetFailed'), [IntToStr(Code)]), mbInformation, MB_OK);
      except
        // (a failed or cancelled download: dAHKxodus is still installed, .NET can come later)
        MsgBox(AddPeriod(GetExceptionMessage), mbInformation, MB_OK);
      end;
    finally
      DownloadPage.Hide;
    end;
  end;
end;

function UpdateReadyMemo(Space, NewLine, MemoUserInfoInfo, MemoDirInfo, MemoTypeInfo, MemoComponentsInfo,
  MemoGroupInfo, MemoTasksInfo: String): String;
begin
  Result := MemoDirInfo;
  if MemoGroupInfo <> '' then Result := Result + NewLine + NewLine + MemoGroupInfo;
  if MemoTasksInfo <> '' then Result := Result + NewLine + NewLine + MemoTasksInfo;
  if LocatedAhk <> '' then Result := Result + NewLine + NewLine + FmtMessage(CustomMessage('AhkFound'), [LocatedAhk]);
end;
#endif

procedure CurStepChanged(CurStep: TSetupStep);
var
  Code: Integer;
#ifdef Light
  Lines: TArrayOfString;
#endif
begin
#ifdef Light
  // the AutoHotkey64.exe the user located, for dAHKxodus.exe (one found the usual ways isn't pinned)
  if (CurStep = ssPostInstall) and AhkChosen then begin
    SetArrayLength(Lines, 1);
    Lines[0] := LocatedAhk;
    SaveStringsToUTF8File(ExpandConstant('{app}\AutoHotkey.path'), Lines, False);
  end;
#endif
  if (CurStep = ssPostInstall) and not HasWebView2() and not WizardSilent() then
    if MsgBox(CustomMessage('NoWebView2'), mbInformation, MB_YESNO) = IDYES then
      ShellExec('open', 'https://developer.microsoft.com/microsoft-edge/webview2/', '', '', SW_SHOWNORMAL, ewNoWait, Code);
end;
