; Inno Setup script for the Email Notes native COM add-in for classic Outlook.
;
; Everything is per-user (HKCU) — no administrator rights, no Microsoft login,
; no internet. The COM class is registered in BOTH the 64-bit and 32-bit
; registry views so it loads regardless of whether Outlook is 32-bit or 64-bit.
;
; Build:  ISCC.exe EmailNotes.iss   (needs ..\native\bin\Release\EmailNotesAddin.dll)
; Output: Output\EmailNotesSetup.exe

#define AppName "Email Notes"
#define AppVersion "0.2.4"
#define Publisher "Pixelschmied"
#define AppUrl "https://github.com/Pixelschmied/Outlook-Notes"

; COM identifiers (leading brace doubled so Inno emits a literal "{").
#define ConnectClsid "{{E7A9C1F4-3B2D-4A6E-8C1B-9D0E2F3A4B51}"
#define PaneClsid "{{B1D9E7C2-6F1A-4C2E-9E7D-2A5B3C4D5E61}"
#define ControlCat "{{40FC6ED4-2438-11CF-A3DB-080036F12502}"
#define AsmFullName "EmailNotesAddin, Version=0.2.4.0, Culture=neutral, PublicKeyToken=null"
#define AsmVer "0.2.4.0"

[Setup]
AppId={{9EE94300-AFF0-4300-B6A7-51EBCEA1FBD7}}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#Publisher}
AppPublisherURL={#AppUrl}
AppSupportURL={#AppUrl}
VersionInfoDescription={#AppName} Setup
VersionInfoProductName={#AppName}
VersionInfoVersion={#AppVersion}
DefaultDirName={localappdata}\EmailNotes
DisableProgramGroupPage=yes
DisableDirPage=yes
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=yes
OutputDir=Output
OutputBaseFilename=EmailNotesSetup
SetupIconFile=..\assets\app.ico
UninstallDisplayIcon={app}\app.ico
UninstallDisplayName={#AppName}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
WizardImageFile=wizard-large.bmp
WizardSmallImageFile=wizard-small.bmp
WizardImageStretch=no
InfoAfterFile=after.txt

[Languages]
Name: "en"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "..\native\bin\Release\EmailNotesAddin.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\assets\app.ico"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; =========================================================================
;  Native (64-bit) registry view — used by 64-bit Outlook.
; =========================================================================
; --- COM add-in class (EmailNotes.Connect) ---
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}"; ValueType: string; ValueData: "EmailNotes.Connect"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "EmailNotes.Connect"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmFullName}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:CodeBase}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "Class"; ValueData: "EmailNotes.Connect"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmFullName}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:CodeBase}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}\ProgId"; ValueType: string; ValueData: "EmailNotes.Connect"
Root: HKCU; Subkey: "Software\Classes\EmailNotes.Connect"; ValueType: string; ValueData: "EmailNotes.Connect"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\EmailNotes.Connect\CLSID"; ValueType: string; ValueData: "{#ConnectClsid}"
; --- Task-pane control (EmailNotes.NotesPane) ---
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}"; ValueType: string; ValueData: "EmailNotes.NotesPane"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "EmailNotes.NotesPane"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmFullName}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:CodeBase}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "Class"; ValueData: "EmailNotes.NotesPane"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmFullName}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:CodeBase}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\ProgId"; ValueType: string; ValueData: "EmailNotes.NotesPane"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\Implemented Categories\{#ControlCat}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\Control"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\MiscStatus"; ValueType: string; ValueData: "0"
Root: HKCU; Subkey: "Software\Classes\EmailNotes.NotesPane"; ValueType: string; ValueData: "EmailNotes.NotesPane"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\EmailNotes.NotesPane\CLSID"; ValueType: string; ValueData: "{#PaneClsid}"

; =========================================================================
;  32-bit (Wow6432Node) view — used by 32-bit Outlook on 64-bit Windows.
; =========================================================================
; --- COM add-in class ---
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#ConnectClsid}"; ValueType: string; ValueData: "EmailNotes.Connect"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#ConnectClsid}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#ConnectClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#ConnectClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "EmailNotes.Connect"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#ConnectClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmFullName}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#ConnectClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#ConnectClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:CodeBase}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#ConnectClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "Class"; ValueData: "EmailNotes.Connect"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#ConnectClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmFullName}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#ConnectClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#ConnectClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:CodeBase}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#ConnectClsid}\ProgId"; ValueType: string; ValueData: "EmailNotes.Connect"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\EmailNotes.Connect"; ValueType: string; ValueData: "EmailNotes.Connect"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\EmailNotes.Connect\CLSID"; ValueType: string; ValueData: "{#ConnectClsid}"; Check: IsWin64
; --- Task-pane control ---
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}"; ValueType: string; ValueData: "EmailNotes.NotesPane"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "EmailNotes.NotesPane"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmFullName}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:CodeBase}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "Class"; ValueData: "EmailNotes.NotesPane"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmFullName}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:CodeBase}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\ProgId"; ValueType: string; ValueData: "EmailNotes.NotesPane"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\Implemented Categories\{#ControlCat}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\Control"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\MiscStatus"; ValueType: string; ValueData: "0"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\EmailNotes.NotesPane"; ValueType: string; ValueData: "EmailNotes.NotesPane"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\EmailNotes.NotesPane\CLSID"; ValueType: string; ValueData: "{#PaneClsid}"; Check: IsWin64

; =========================================================================
;  Outlook add-in load entry (not bitness-redirected — one entry serves both).
; =========================================================================
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\EmailNotes.Connect"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\EmailNotes.Connect"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: 0
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\EmailNotes.Connect"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Email Notes"
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\EmailNotes.Connect"; ValueType: string; ValueName: "Description"; ValueData: "Privater Notizblock neben der Mail"

[Messages]
WelcomeLabel1=Welcome to the [name] add-in
WelcomeLabel2=This installs a private notepad next to your mail in classic Outlook.%n%nEverything stays on your device — no account, no cloud, no login. You do not need administrator rights.%n%nPlease close Outlook before continuing.
FinishedHeadingLabel=Email Notes is installed
FinishedLabelNoIcons=Start Outlook and open a mail. The notepad docks on the right automatically; you can also toggle it with "Notes" in the "Email Notes" ribbon group.
FinishedLabel=Start Outlook and open a mail. The notepad docks on the right automatically; you can also toggle it with "Notes" in the "Email Notes" ribbon group.

[Code]
function CodeBase(Param: String): String;
var
  P: String;
begin
  P := ExpandConstant('{app}\EmailNotesAddin.dll');
  StringChangeEx(P, '\', '/', True);
  Result := 'file:///' + P;
end;
