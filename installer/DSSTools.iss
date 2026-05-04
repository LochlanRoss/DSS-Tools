; Inno Setup 6 — click-through installer for DSS Tools
;
; Prerequisites: PyInstaller must have built ..\dist\DSSToolsUpdater.exe first, then ..\dist\DSSTools.exe
; (main onefile embeds the updater; the installer also places DSSToolsUpdater.exe beside DSSTools.exe).
; Requires ..\dss_tools.ico at repo root for wizard + shortcut icons (CI runs tools\ensure_dss_tools_ico.py first).
; Compile locally (from repo root), after PyInstaller:
;   "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" /DMyAppVersion=0.1.0 installer\DSSTools.iss
;
#ifndef MyAppVersion
#define MyAppVersion "0.0.0"
#endif
#define MyAppName "DSS Tools"
#define MyAppExeName "DSSTools.exe"
#define MyAppPublisher "DSS Tools"

#if FileExists(SourcePath + "..\dss_tools.ico")
#define HasAppIco
#endif

[Setup]
AppId={{E7B8F9A0-1D2C-4E5F-8A9B-0C1D2E3F4A5B}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
CloseApplications=yes
RestartApplications=no
UsePreviousAppDir=yes
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=..\dist
OutputBaseFilename=DSSToolsSetup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
DisableProgramGroupPage=no
#ifdef HasAppIco
SetupIconFile=..\dss_tools.ico
UninstallDisplayIcon={app}\dss_tools.ico
#else
UninstallDisplayIcon={app}\{#MyAppExeName}
#endif

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "..\dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\dist\DSSToolsUpdater.exe"; DestDir: "{app}"; Flags: ignoreversion
#ifdef HasAppIco
Source: "..\dss_tools.ico"; DestDir: "{app}"; DestName: "dss_tools.ico"; Flags: ignoreversion
#endif

[Icons]
#ifdef HasAppIco
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\dss_tools.ico"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon; IconFilename: "{app}\dss_tools.ico"
#else
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
#endif

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
