; CCTV IP Toolkit - Inno Setup script
; Builds an installer that:
;   - Installs to Program Files (admin) or %LocalAppData%\Programs (non-admin)
;   - Adds Start Menu entry
;   - Optional Desktop shortcut (checkbox in installer)
;   - Detects existing install via AppId and offers in-place upgrade
;   - Generates an uninstaller (Add/Remove Programs entry)
;
; Build: ISCC.exe installer.iss
; Output: Output\CCTVIPToolkit-Setup-vX.Y.Z.exe
;
; The AppId GUID below is the UPGRADE KEY for this product across all versions.
; NEVER change it — Inno uses it to find prior installs and replace them in place.

#define MyAppName "CCTV IP Toolkit"
#define MyAppPublisher "theLostPing"
#define MyAppURL "https://cctv.thelostping.net"
#define MyAppExeName "CCTVIPToolkit.exe"
#define MyAppId "{{EFA229D8-28D4-4122-A173-C9B028181C50}"

; Version is read from the PyInstaller-built EXE so we never have it in two places.
; Override on the command line: ISCC.exe /DMyAppVersion=4.3.0 installer.iss
#ifndef MyAppVersion
  #define MyAppVersion "0.0.0"
#endif

[Setup]
AppId={#MyAppId}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}

; Install location: Program Files when admin, per-user folder when not.
; "lowest" = installer first asks if it can elevate; if user declines, falls back to per-user.
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
DefaultDirName={autopf}\{#MyAppPublisher}\{#MyAppName}
DefaultGroupName={#MyAppName}

; Detect prior install via AppId and upgrade in place
UsePreviousAppDir=yes
UsePreviousGroup=yes
UsePreviousTasks=yes

; Cosmetics
WizardStyle=modern
SetupIconFile=app.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName}
DisableDirPage=auto
DisableProgramGroupPage=yes
ShowLanguageDialog=no

; Compression
Compression=lzma2/ultra64
SolidCompression=yes

; In-app upgrade flow:
;   - The running CCTVIPToolkit.exe launches this installer detached, then exits.
;   - If the app is still alive when Inno scans (race window), CloseApplications=yes
;     triggers the "Setup has detected that this application is currently running"
;     page with a "Close all and continue" button (uses Restart Manager API).
;   - RestartApplications=yes asks the OS to relaunch the app post-install.
;   - The [Run] section's launch (without skipifsilent) is the belt to that suspenders.
CloseApplications=yes
RestartApplications=yes

; Output
OutputDir=Output
OutputBaseFilename=CCTVIPToolkit-Setup-v{#MyAppVersion}
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion isreadme
Source: "CHANGELOG.md"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Auto-launch INTENTIONALLY DROPPED in v4.4.4 after Brian saw a "Failed to load
; Python DLL ... python312.dll" error from PyInstaller's onefile bootloader on
; the in-app updater path. The race: Inno's RestartManager closes the running
; v4.4.x process, replaces files in {app}, then [Run] fires and starts the new
; .exe before the OS / antivirus / Defender finish releasing locks on the freshly
; extracted _MEI<random> temp dir. PyInstaller's bootloader then can't load
; python312.dll from that temp dir.
;
; The user sees one extra click — Start Menu / Desktop shortcut — but the new
; version launches reliably from a fresh process context. The in-app updater
; UI text was updated in v4.4.4 to set this expectation explicitly.
;
; If you ever want to re-enable, the standard mitigation is to launch via a
; small launcher (cmd /c "timeout 3 && start <exe>") so the new process starts
; ~3s after Inno's file replacement settles. Don't use shellexec without that
; delay — it reproduces the DLL race.
;
; (Old line for reference, do not uncomment without delay launcher:)
; Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,...}"; Flags: nowait postinstall shellexec

[UninstallDelete]
; Clean up app data on uninstall (optional — comment out to keep user settings)
; Type: filesandordirs; Name: "{userappdata}\theLostPing\CCTV IP Toolkit"
