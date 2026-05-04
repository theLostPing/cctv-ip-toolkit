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
; v4.4.6: was =yes — Restart Manager would auto-relaunch the closed app after
; install. Setting to =no for the same reason [Run] auto-launch was dropped:
; the relaunched exe could fire before file locks released on the freshly-
; extracted PyInstaller temp dir. User opens from Start Menu after install.
RestartApplications=no

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
; Auto-launch DROPPED AGAIN in v4.4.6.
; v4.4.5 added a 3s delay launcher (`cmd /c "timeout 3 && start <exe>"`) to dodge
; the PyInstaller DLL-load race that bit v4.4.0. Real-world testing showed 3s
; isn't always enough on slower disks / AV-active systems — the in-app update
; download → install → relaunch still occasionally fired the new exe before
; Windows fully released file locks on the freshly-extracted _MEI temp dir.
; Reverted to v4.4.4 behavior: installer ends with the standard Setup Complete
; page, user opens the toolkit from Start Menu / Desktop shortcut.
;
; (No [Run] entries — intentional. Don't re-add without solving the race
; deterministically — e.g. waiting on a process-handle from the running app
; rather than a fixed sleep, or shipping a tiny stub launcher that polls the
; install dir for stability before exec'ing the toolkit.)

[UninstallDelete]
; Clean up app data on uninstall (optional — comment out to keep user settings)
; Type: filesandordirs; Name: "{userappdata}\theLostPing\CCTV IP Toolkit"
