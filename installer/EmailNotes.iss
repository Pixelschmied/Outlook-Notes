; Inno Setup script for the Email Notes Outlook add-in installer.
;
; What it does (per-user, no administrator rights needed):
;   * copies the add-in manifest to the user's local app data
;   * registers the manifest with classic Outlook so the add-in appears
;   * cleans everything up on uninstall
;
; The add-in's actual code (HTML/JS/CSS) is served over HTTPS from the project's
; GitHub Pages site, which the manifest points to. No local web server, no Node,
; no certificate handling on the user's machine.
;
; Build:  ISCC.exe EmailNotes.iss   (Inno Setup 6+)
; Output: Output\EmailNotesSetup.exe

#define AppName "Email Notes"
#define AppVersion "0.1.0"
#define Publisher "Pixelschmied"
#define AppUrl "https://github.com/Pixelschmied/Outlook-Notes"
#define AddinId "9ee94300-aff0-4300-b6a7-51ebcea1fbd7"

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
DisableReadyPage=no
PrivilegesRequired=lowest
OutputDir=Output
OutputBaseFilename=EmailNotesSetup
SetupIconFile=..\assets\app.ico
UninstallDisplayIcon={app}\app.ico
UninstallDisplayName={#AppName}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
WizardSizePercent=110
WizardImageFile=wizard-large.bmp
WizardSmallImageFile=wizard-small.bmp
WizardImageStretch=no
InfoAfterFile=after.txt

[Languages]
Name: "en"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "..\dist\manifest.xml"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\assets\app.ico"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; Sideload the add-in for the current user in classic Outlook (Office 16.0).
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\Developer"; Flags: uninsdeletekeyifempty
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\Developer"; ValueType: string; ValueName: "{#AddinId}"; ValueData: "{app}\manifest.xml"; Flags: uninsdeletevalue

[Messages]
WelcomeLabel1=Welcome to the [name] add-in
WelcomeLabel2=This will add a private notepad next to your mail in classic Outlook.%n%nEverything stays on your device — no account, no cloud. You do not need administrator rights.
FinishedHeadingLabel=Email Notes is installed
FinishedLabelNoIcons=Restart Outlook, open any mail, and click the "Notes" button on the ribbon. Pin the pane to keep it beside your inbox.
FinishedLabel=Restart Outlook, open any mail, and click the "Notes" button on the ribbon. Pin the pane to keep it beside your inbox.
