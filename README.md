# Email Notes for Outlook

A private notepad that docks **right next to your mail** in classic Outlook.
Every mail is linked to its own note by a stable, sequential ID — the first mail
you take notes on is **Note 1**, the next distinct mail is **Note 2**, and so on.
Paste screenshots and attach files straight into the note. Nothing ever leaves
your device.

This is the "email notes" idea from the internal Pixel-App, rebuilt as a proper,
standalone **Office add-in** so it lives inside Outlook itself.

> **Independent, open-source project.** Not affiliated with, endorsed by, or
> sponsored by Microsoft. See [Legal & Microsoft compliance](#legal--microsoft-compliance).

---

## Install (one click, no Node, no dev tools)

1. **Download `EmailNotesSetup.exe`** from the
   [latest release](https://github.com/Pixelschmied/Outlook-Notes/releases).
2. **Run it.** No administrator rights are needed — it installs just for you.
   (Windows SmartScreen may warn about an unknown publisher because the
   installer isn't code-signed yet; choose **More info → Run anyway**. See
   [Code signing](#code-signing).)
3. **Restart Outlook.** Open any mail, click **Notes** on the ribbon, and click
   the **pin** icon so the pane stays docked while you browse your inbox.

That's it — no Node.js, no commands, no local server. The installer registers
the add-in with Outlook; the add-in's code is served over HTTPS from this
project's GitHub Pages site (your notes are **not** — they stay on your PC).

To remove it: uninstall **Email Notes** from *Windows Settings → Apps*, then
restart Outlook.

## What it does

- 📌 **Docks next to the mail.** A pinnable task pane on the message-read
  surface that automatically switches to the note for whichever mail you select.
- 🔗 **One note per mail, by a stable ID.** Each mail is matched to its note via
  the mail's `internetMessageId`, which never changes when a mail is moved
  between folders. See [How the linking works](#how-the-linking-works).
- 📋 **Paste screenshots.** `Ctrl+V` an image into the note (or use **Paste
  screenshot**) and it's stored as an attachment, with a `[📎 name]` reference
  dropped into the text — exactly like the Pixel-App.
- 📎 **Attach files.** Pick any file(s); they're kept with the note.
- ✏️ **Manage attachments.** Open, rename, remove, and copy images back to the
  clipboard.
- 💾 **Autosave.** Typing is saved automatically (and immediately when you
  switch mails or hide the pane).
- 🔒 **Fully local & offline.** All notes and attachments live in the add-in's
  local storage. There is no server, no account, no cloud.

## How the linking works

The rule is simple: *every mail is linked to a note by a unique id — Mail 1 =
Note 1, Mail 2 = Note 2, …*. Implementation:

1. When you select a mail, the add-in reads a **stable mail key** — preferring
   `internetMessageId` (globally unique, survives folder moves and restarts),
   falling back to the item id / conversation id.
2. A local registry maps `mailKey → noteId`. The **first** time a mail is seen
   it's assigned the **next sequential number**, permanently — so the mail
   always reopens the same note.
3. The note text and attachments are stored under that `noteId`.

"Mail 1 = Note 1" means *the first mail you actually take notes on* becomes
Note 1 — not a fixed position in your inbox.

## Privacy & data

- **Everything stays on your device.** Notes and attachments are stored in the
  add-in's local IndexedDB. There is no backend, and the add-in makes no network
  calls of its own.
- The only external resource loaded is **office.js from Microsoft's CDN**, which
  every Office add-in requires.
- The add-in requests the minimum permission, **`ReadItem`**: it reads the
  selected mail's id, subject, and sender to link and label the note. It does
  **not** modify, move, send, or delete your mail.
- Storage is local to the machine and Outlook profile, so notes are not synced
  across devices.

## Legal & Microsoft compliance

Built to follow Microsoft's rules for Office/Outlook add-ins:

- Uses the **official Office Add-ins platform** and a standard add-in manifest —
  the supported, documented way to extend Outlook.
- **office.js is loaded from the official Microsoft CDN**, never bundled or
  repackaged.
- **All add-in resource URLs are HTTPS.**
- Requests the **least privilege** it needs (`ReadItem`).
- **Installer footprint is minimal and per-user:** it writes only to your own
  user profile — one value under `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`
  (Microsoft's documented Windows sideloading location) plus a copy of the
  manifest in `%LOCALAPPDATA%`. No system files, no admin rights, fully removed
  on uninstall.
- **Trademarks.** "Microsoft", "Outlook", and "Office" are trademarks of
  Microsoft Corporation. This is an **independent, unofficial** add-in and is
  **not affiliated with, endorsed by, or sponsored by Microsoft**. The product
  name ("Email Notes") uses no Microsoft brand names, and the icons are original.

### Code signing

The installer is currently **unsigned**, so Windows SmartScreen shows an
"unknown publisher" prompt (**More info → Run anyway** to proceed). This is
expected for an open-source project and does not mean the software is unsafe —
you can review every line here and build it yourself. For wide distribution,
sign `EmailNotesSetup.exe` with an Authenticode certificate.

### Publishing to AppSource (optional)

Listing an add-in on Microsoft AppSource requires passing Microsoft's
[Commercial marketplace certification policies](https://learn.microsoft.com/legal/marketplace/certification-policies)
and the
[Office Store validation policies](https://learn.microsoft.com/office/dev/store/validation-policies).
Sideloading or admin-deploying inside your own organization does **not** require
certification. This repository targets self-hosting via the installer; review
those policies before publishing.

---

## For maintainers: one-time setup

The installer points the add-in at this project's GitHub Pages site, so the
static files must be published once:

1. In the repo, go to **Settings → Pages** and set **Source: GitHub Actions**.
2. The [`pages`](.github/workflows/pages.yml) workflow builds and deploys
   `dist/` to `https://pixelschmied.github.io/Outlook-Notes/` on every push to
   `main`.
3. The [`installer`](.github/workflows/installer.yml) workflow compiles
   `EmailNotesSetup.exe` on a Windows runner and uploads it as a build artifact;
   on a published GitHub Release it's attached as a downloadable asset.

If you host somewhere other than GitHub Pages, set `PUBLIC_URL` when building the
manifest (`PUBLIC_URL=https://your.host npm run manifest:prod`) and update the
same value in the two workflows.

## Build from source (contributors only)

You only need this if you want to **modify the add-in** — end users just run the
installer above.

```bash
npm install
npm run dev-certs   # trust a local HTTPS certificate (Office requires HTTPS)
npm start           # dev server on https://localhost:3000
```

Then sideload the dev `manifest.xml` (which points at `https://localhost:3000`)
into Outlook: *File → Get Add-ins → My add-ins → Add a custom add-in → Add from
file…*. See Microsoft's
[sideloading guide](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing).

| Script | What it does |
| --- | --- |
| `npm start` | Dev server with hot reload on `https://localhost:3000` |
| `npm run build` | Production build into `dist/` |
| `npm run typecheck` | TypeScript type-check |
| `npm run manifest:prod` | Write `dist/manifest.xml` for your `PUBLIC_URL` host |
| `npm run validate` | Validate `manifest.xml` (Microsoft's online validator) |
| `npm run make-icons` | Regenerate the PNG icons in `assets/` |
| `npm run make-installer-art` | Regenerate the installer artwork + `app.ico` |

Building the installer locally (on Windows) needs
[Inno Setup 6+](https://jrsoftware.org/isinfo.php):

```bash
npm run build
npm run make-installer-art
npm run manifest:prod
ISCC installer\EmailNotes.iss   # → installer\Output\EmailNotesSetup.exe
```

## Project structure

```
Outlook-Notes/
├─ manifest.xml               # Dev manifest (localhost); prod copy is stamped into dist/
├─ assets/                    # Generated PNG icons + app.ico
├─ installer/
│  ├─ EmailNotes.iss          # Inno Setup script (per-user sideload registration)
│  ├─ after.txt               # Post-install instructions shown in the wizard
│  └─ wizard-*.bmp            # Generated wizard artwork
├─ scripts/
│  ├─ make-icons.mjs          # Icon generator (no image libraries)
│  ├─ make-installer-art.mjs  # Installer artwork + .ico generator
│  └─ build-manifest.mjs      # Stamps the public host into dist/manifest.xml
├─ src/
│  ├─ core/                   # Storage + logic (db, mail key, attachments, types)
│  ├─ taskpane/               # Notepad UI (HTML/CSS/TS)
│  └─ commands/               # Ribbon function file
└─ .github/workflows/         # build, pages (host), installer (.exe)
```

## Requirements

- Classic Outlook for Windows (desktop), Microsoft 365 / Exchange account.
- Mailbox **1.5** requirement set (add-in commands + task-pane pinning), which
  classic Outlook has supported for years.

## Not included (yet)

To keep this a clean standalone add-in, the Pixel-App's integrations (GitLab
issues, the shared To-Do board) are intentionally left out — they depend on that
app's own backend. The core email-notes experience (per-mail notes, sequential
linking, screenshots, attachments) is complete.

## Contributing

Issues and pull requests welcome. Please keep documentation in English and run
`npm run typecheck` and `npm run build` before submitting.

## License

[MIT](./LICENSE) © Pixelschmied
