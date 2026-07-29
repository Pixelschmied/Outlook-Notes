; Inno Setup script for the Outlook-Notes native COM add-in (NetOffice-based).
;
; Per-user (HKCU) — no admin, no Microsoft login, no internet. The managed COM
; classes are registered in BOTH the 64-bit and 32-bit views so they load
; whatever bitness Outlook is.
;
; Build:  ISCC.exe OutlookNotes.iss   (needs ..\native\bin\Release\*.dll)
; Output: Output\OutlookNotesSetup.exe

#define AppName "Outlook-Notes"
#define AppVersion "0.1.0"
#define Publisher "Pixelschmied"
#define AppUrl "https://github.com/Pixelschmied/Outlook-Notes"

#define AddinClsid "{{E7A9C1F4-3B2D-4A6E-8C1B-9D0E2F3A4B51}"
#define AddinClass "OutlookNotes.AddIn"
#define AddinProgId "OutlookNotes.AddIn"
#define PaneClsid "{{B1D9E7C2-6F1A-4C2E-9E7D-2A5B3C4D5E61}"
#define PaneClass "OutlookNotes.NotesPane"
#define PaneProgId "OutlookNotes.NotesPane"
#define ControlCat "{{40FC6ED4-2438-11CF-A3DB-080036F12502}"
#define DotNetCat "{{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"
#define AsmFullName "OutlookNotesAddin, Version=0.1.0.0, Culture=neutral, PublicKeyToken=null"
#define AsmVer "0.1.0.0"
#define CB "{code:CodeBase}"

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
DefaultDirName={localappdata}\OutlookNotes
DisableProgramGroupPage=yes
DisableDirPage=yes
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=yes
OutputDir=Output
OutputBaseFilename=OutlookNotesSetup
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
Source: "..\native\bin\Release\*.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\assets\app.ico"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; ================= Native (64-bit) view =================
; --- Add-in class ---
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AddinClsid}"; ValueType: string; ValueData: "{#AddinClass}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AddinClsid}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AddinClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AddinClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "{#AddinClass}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AddinClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmFullName}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AddinClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AddinClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{#CB}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AddinClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "Class"; ValueData: "{#AddinClass}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AddinClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmFullName}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AddinClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AddinClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "CodeBase"; ValueData: "{#CB}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AddinClsid}\ProgId"; ValueType: string; ValueData: "{#AddinProgId}"
Root: HKCU; Subkey: "Software\Classes\{#AddinProgId}"; ValueType: string; ValueData: "{#AppName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\{#AddinProgId}\CLSID"; ValueType: string; ValueData: "{#AddinClsid}"
; --- Task-pane control ---
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}"; ValueType: string; ValueData: "{#PaneClass}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Apartment"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "{#PaneClass}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmFullName}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{#CB}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "Class"; ValueData: "{#PaneClass}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmFullName}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "CodeBase"; ValueData: "{#CB}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\ProgId"; ValueType: string; ValueData: "{#PaneProgId}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\Implemented Categories\{#ControlCat}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\Control"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\MiscStatus"; ValueType: string; ValueData: "0"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\MiscStatus\1"; ValueType: string; ValueData: "131457"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#PaneClsid}\Implemented Categories\{#DotNetCat}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\{#PaneProgId}"; ValueType: string; ValueData: "{#PaneClass}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\{#PaneProgId}\CLSID"; ValueType: string; ValueData: "{#PaneClsid}"

; ================= 32-bit (Wow6432Node) view =================
; --- Add-in class ---
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#AddinClsid}"; ValueType: string; ValueData: "{#AddinClass}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#AddinClsid}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#AddinClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#AddinClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "{#AddinClass}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#AddinClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmFullName}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#AddinClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#AddinClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{#CB}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#AddinClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "Class"; ValueData: "{#AddinClass}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#AddinClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmFullName}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#AddinClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#AddinClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "CodeBase"; ValueData: "{#CB}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#AddinClsid}\ProgId"; ValueType: string; ValueData: "{#AddinProgId}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\{#AddinProgId}"; ValueType: string; ValueData: "{#AppName}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\{#AddinProgId}\CLSID"; ValueType: string; ValueData: "{#AddinClsid}"; Check: IsWin64
; --- Task-pane control ---
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}"; ValueType: string; ValueData: "{#PaneClass}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Apartment"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "{#PaneClass}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmFullName}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{#CB}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "Class"; ValueData: "{#PaneClass}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmFullName}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "CodeBase"; ValueData: "{#CB}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\ProgId"; ValueType: string; ValueData: "{#PaneProgId}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\Implemented Categories\{#ControlCat}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\Control"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\MiscStatus"; ValueType: string; ValueData: "0"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\MiscStatus\1"; ValueType: string; ValueData: "131457"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#PaneClsid}\Implemented Categories\{#DotNetCat}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\{#PaneProgId}"; ValueType: string; ValueData: "{#PaneClass}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\{#PaneProgId}\CLSID"; ValueType: string; ValueData: "{#PaneClsid}"; Check: IsWin64

; ================= Outlook add-in load entry (both bitnesses) =================
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AddinProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AddinProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: 0
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AddinProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AddinProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Privater Notizblock neben der Mail"

[Messages]
WelcomeLabel1=Welcome to the [name] add-in
WelcomeLabel2=This installs a private notepad next to your mail in classic Outlook.%n%nEverything stays on your device — no account, no cloud, no login. You do not need administrator rights.%n%nPlease close Outlook before continuing.
FinishedHeadingLabel=Outlook-Notes is installed
FinishedLabelNoIcons=Start Outlook, open a mail, and click "Notizen" in the "Outlook-Notes" ribbon group. The notepad docks on the right.
FinishedLabel=Start Outlook, open a mail, and click "Notizen" in the "Outlook-Notes" ribbon group. The notepad docks on the right.

[Code]
function CodeBase(Param: String): String;
var
  P: String;
begin
  P := ExpandConstant('{app}\OutlookNotesAddin.dll');
  StringChangeEx(P, '\', '/', True);
  Result := 'file:///' + P;
end;
