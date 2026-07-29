# Outlook-Notes

A private notepad that docks right next to your mail in **classic Outlook for
Windows**. Every mail gets its own note, linked by a stable, sequential ID
(Mail 1 = Note 1, Mail 2 = Note 2, …). Paste screenshots straight from the
clipboard and attach files — like a sticky note for your inbox.

It runs **completely locally**: no Microsoft account, no sign-in, no internet,
no cloud. Your notes never leave your computer. Installation is per-user and
needs **no administrator rights**.

---

## Install

1. **Close Outlook.**
2. Download the latest **`OutlookNotesSetup.exe`** from the
   [Releases page](https://github.com/Pixelschmied/Outlook-Notes/releases/latest).
3. Run it. Windows SmartScreen may warn about an unknown publisher (the
   installer is not code-signed) — choose **More info → Run anyway**.
4. Start Outlook and click a mail. The notepad docks on the right.

That's it. There is nothing else to configure, and no Node.js, developer tools
or Office sign-in are required.

## How it works

- **One note per mail.** The first time you select a mail, Outlook-Notes
  creates an empty note and gives it the next sequential number. Re-selecting
  the same mail always brings back the same note.
- **Auto-save.** Whatever you type is saved automatically while you work.
- **Screenshots.** Copy an image (e.g. `PrtScn` or the Snipping Tool), then
  click **📋 Screenshot einfügen** or press `Ctrl+V` in the note. The image is
  stored with the note and referenced inline.
- **Attachments.** Click **📎 Datei anhängen** to attach any file. Double-click
  an attachment to open it.

All data is stored locally under:

```
%APPDATA%\OutlookNotes\
├─ notes.json            (your notes, keyed by the mail's Outlook EntryID)
└─ attachments\<noteId>\ (screenshots and attached files)
```

Because everything is a plain file on your machine, you can back it up or move
it just by copying that folder.

## Automatic updates

On every Outlook start, Outlook-Notes quietly checks GitHub for a newer
release. If one exists, a banner appears at the top of the notepad:

> ⬆ Neue Version verfügbar – jetzt aktualisieren

One click downloads and runs the new installer for you. If you're offline,
nothing happens — no errors, no interruptions. After the first install you
never have to update by hand again.

## Privacy

Outlook-Notes is fully offline by design. The **only** network request it ever
makes is the version check against the public GitHub Releases API, and that
sends no personal data. Your mails, notes, screenshots and attachments stay on
your device.

## Uninstall

Uninstall **Outlook-Notes** from Windows Settings → **Apps**, then restart
Outlook. Your notes under `%APPDATA%\OutlookNotes` are left in place; delete
that folder too if you want to remove everything.

---

## For developers

Outlook-Notes is a native COM add-in for classic Outlook, written in C# against
.NET Framework 4.8 using [NetOffice](https://github.com/NetOfficeFw/NetOffice)
for the add-in plumbing. The docked notepad is a WinForms control hosted in an
Outlook Custom Task Pane. There is no web view and no Office.js.

```
native/                 The add-in
├─ AddIn.cs             NetOffice COMAddin: ribbon + task-pane wiring
├─ NotesPane.cs         WinForms notepad (UI, clipboard paste, attachments)
├─ Store.cs             Local JSON storage + sequential IDs
├─ Updater.cs           GitHub-release update check
├─ AssemblyInfo.cs      Version of the DLL
└─ OutlookNotesAddin.csproj
installer/
├─ OutlookNotes.iss     Inno Setup script (per-user HKCU COM registration)
├─ after.txt            Post-install notes
└─ wizard-*.bmp         Installer artwork
.github/workflows/
├─ installer.yml        Build the DLL + installer and publish a Release
└─ cleanup-releases.yml One-off: delete all releases + tags
```

### Build locally (Windows)

```powershell
dotnet build native/OutlookNotesAddin.csproj -c Release
# Inno Setup 6 must be installed:
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer\OutlookNotes.iss
# → installer\Output\OutlookNotesSetup.exe
```

### Cut a release

Releases are built on a Windows runner in GitHub Actions. To publish one:

1. Bump the version in **`native/AssemblyInfo.cs`** and
   **`installer/OutlookNotes.iss`** (`AppVersion`, `AsmFullName`, `AsmVer`).
2. Push, then run the **installer** workflow
   (Actions → *installer* → **Run workflow**) with the tag, e.g. `v0.2.0`, and
   *Create release* enabled.

The workflow builds `OutlookNotesSetup.exe`, attaches it to a new GitHub
Release, and marks it as *latest* — which is what the in-app auto-updater
picks up.

## Legal & compliance

Outlook-Notes is an unofficial, independent project. It is **not** affiliated
with, endorsed by, or sponsored by Microsoft. "Outlook", "Microsoft" and
related marks belong to Microsoft Corporation and are used only to describe
interoperability.

The add-in uses documented, supported extensibility points (a COM add-in with a
Custom Task Pane via the Outlook object model). It reads only the currently
selected mail's metadata to link a note to it, stores everything locally, and
transmits no user data.

## License

[MIT](LICENSE) © Pixelschmied
