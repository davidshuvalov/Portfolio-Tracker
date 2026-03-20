; Portfolio Tracker v2 — Inno Setup Installation Script
; Requires: Inno Setup 6.x  (https://jrsoftware.org/isinfo.php)
;
; How to build:
;   1. Run build_windows.bat  (creates dist\PortfolioTracker\)
;   2. Open this file in Inno Setup Compiler  — OR —
;      Run from command line:
;        "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer\setup.iss
;
; Output: installer\Output\PortfolioTracker-v2.0.0-Setup.exe

#define MyAppName      "Portfolio Tracker"
#define MyAppVersion   "2.0.0"
#define MyAppPublisher "Portfolio Tracker"
#define MyAppURL       "https://github.com/davidshuvalov/Portfolio-Tracker"
#define MyAppExeName   "PortfolioTracker.exe"
#define MyAppId        "{{A3F2B91C-4E7D-4A2F-8C1B-9D6E5F3A2B10}"

[Setup]
AppId={#MyAppId}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} v{#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
; Compression
Compression=lzma2/ultra64
SolidCompression=yes
; Installer appearance
WizardStyle=modern
WizardSizePercent=130
; Output
OutputDir=installer\Output
OutputBaseFilename=PortfolioTracker-v{#MyAppVersion}-Setup
; Require administrator rights for Program Files install
PrivilegesRequired=admin
; Minimum Windows version: Windows 10
MinVersion=10.0
; Architecture
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
; Uninstall
UninstallDisplayName={#MyAppName} v{#MyAppVersion}
UninstallDisplayIcon={app}\{#MyAppExeName}
; Disable old version check (handled by AppId)
CloseApplications=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon";     Description: "Create a &desktop shortcut";            GroupDescription: "Additional icons:"; Flags: checked
Name: "quicklaunchicon"; Description: "Create a &Quick Launch shortcut";        GroupDescription: "Additional icons:"; Flags: unchecked; OnlyBelowVersion: 6.1; Check: not IsAdminInstallMode

[Files]
; Main application bundle (PyInstaller output)
Source: "..\dist\PortfolioTracker\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}";              Filename: "{app}\{#MyAppExeName}"; Comment: "Portfolio analytics for systematic futures traders"
Name: "{group}\Uninstall {#MyAppName}";   Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}";       Filename: "{app}\{#MyAppExeName}"; Comment: "Portfolio analytics for systematic futures traders"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName} now"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Remove any Streamlit cache / numba cache left behind
Type: filesandordirs; Name: "{localappdata}\portfolio-tracker"
Type: filesandordirs; Name: "{localappdata}\streamlit"

[Code]
// ─── Pre-install: warn if old version is running ────────────────────────────
function InitializeSetup(): Boolean;
var
  ResultCode: Integer;
begin
  Result := True;
  // Attempt to close any running instance gracefully
  Exec('taskkill', '/F /IM PortfolioTracker.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;

// ─── Finish page: open browser to localhost after launch ────────────────────
procedure CurStepChanged(CurStep: TSetupStep);
begin
  // Nothing extra needed — the app opens the browser automatically on launch
end;
