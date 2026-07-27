; Inno Setup script for the Outlook-Notes native COM add-in (NetOffice-based).
;
; Per-user (HKCU) — no admin, no Microsoft login, no internet. The managed COM
; class is registered in BOTH the 64-bit and 32-bit views so it loads whatever
; bitness Outlook is. Activation itself is standard mscoree; NetOffice provides
; the correct IDTExtensibility2 implementation.
;
; Build:  ISCC.exe EmailNotes.iss   (needs ..\native\bin\Release\*.dll)
; Output: Output\EmailNotesSetup.exe

#define AppName "Outlook-Notes"
#define AppVersion "0.3.0"
#define Publisher "Pixelschmied"
#define AppUrl "https://github.com/Pixelschmied/Outlook-Notes"

#define Clsid "{{E7A9C1F4-3B2D-4A6E-8C1B-9D0E2F3A4B51}"
#define ClassName "EmailNotes.AddIn"
#define ProgId "OutlookNotes.AddIn"
#define AsmFullName "EmailNotesAddin, Version=0.3.0.0, Culture=neutral, PublicKeyToken=null"
#define AsmVer "0.3.0.0"

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
; Ship the add-in DLL and all its NetOffice dependencies.
Source: "..\native\bin\Release\*.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\assets\app.ico"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; =========================================================================
;  Native (64-bit) view — used by 64-bit Outlook.
; =========================================================================
Root: HKCU; Subkey: "Software\Classes\CLSID\{#Clsid}"; ValueType: string; ValueData: "{#ClassName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{#Clsid}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#Clsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#Clsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "{#ClassName}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#Clsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmFullName}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#Clsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#Clsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:CodeBase}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#Clsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "Class"; ValueData: "{#ClassName}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#Clsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmFullName}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#Clsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#Clsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:CodeBase}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#Clsid}\ProgId"; ValueType: string; ValueData: "{#ProgId}"
Root: HKCU; Subkey: "Software\Classes\{#ProgId}"; ValueType: string; ValueData: "{#AppName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\{#ProgId}\CLSID"; ValueType: string; ValueData: "{#Clsid}"

; =========================================================================
;  32-bit (Wow6432Node) view — used by 32-bit Outlook on 64-bit Windows.
; =========================================================================
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#Clsid}"; ValueType: string; ValueData: "{#ClassName}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#Clsid}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#Clsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#Clsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "{#ClassName}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#Clsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmFullName}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#Clsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#Clsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:CodeBase}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#Clsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "Class"; ValueData: "{#ClassName}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#Clsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmFullName}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#Clsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#Clsid}\InprocServer32\{#AsmVer}"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:CodeBase}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\CLSID\{#Clsid}\ProgId"; ValueType: string; ValueData: "{#ProgId}"; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\{#ProgId}"; ValueType: string; ValueData: "{#AppName}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\Wow6432Node\{#ProgId}\CLSID"; ValueType: string; ValueData: "{#Clsid}"; Check: IsWin64

; =========================================================================
;  Outlook add-in load entry (not bitness-redirected — one entry serves both).
; =========================================================================
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#ProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#ProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: 0
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#ProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#ProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Privater Notizblock neben der Mail"

[Messages]
WelcomeLabel1=Welcome to the [name] add-in
WelcomeLabel2=This installs a private notepad next to your mail in classic Outlook.%n%nEverything stays on your device — no account, no cloud, no login. You do not need administrator rights.%n%nPlease close Outlook before continuing.
FinishedHeadingLabel=Outlook-Notes is installed
FinishedLabelNoIcons=Start Outlook. If the add-in loads you will briefly see a confirmation. (This build is a load test.)
FinishedLabel=Start Outlook. If the add-in loads you will briefly see a confirmation. (This build is a load test.)

[Code]
function CodeBase(Param: String): String;
var
  P: String;
begin
  P := ExpandConstant('{app}\EmailNotesAddin.dll');
  StringChangeEx(P, '\', '/', True);
  Result := 'file:///' + P;
end;
