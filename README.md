# Email Notes for Outlook

A private notepad that docks **right next to your mail** in classic Outlook.
Every mail is linked to its own note by a stable, sequential ID — the first mail
you take notes on is **Note 1**, the next distinct mail is **Note 2**, and so on.
Paste screenshots and attach files straight into the note. Nothing ever leaves
your device.

This is the "email notes" idea from the internal Pixel-App, rebuilt as a proper,
standalone **Office Web Add-in** so it lives inside Outlook itself instead of a
separate window.

> **Independent, open-source project.** Not affiliated with, endorsed by, or
> sponsored by Microsoft. See [Legal & Microsoft compliance](#legal--microsoft-compliance).

---

## What it does

- 📌 **Docks next to the mail.** A pinnable task pane on the message-read
  surface. Pin it once and it stays open, automatically switching to the note
  for whichever mail you select.
- 🔗 **One note per mail, by a stable ID.** Each mail is matched to its note
  through the mail's own `internetMessageId` (the RFC 5322 Message-ID). That id
  never changes when a mail is moved between folders, so a mail always resolves
  back to the same note. Notes are numbered sequentially for a friendly label
  (`Note #1`, `Note #2`, …) — see [How the linking works](#how-the-linking-works).
- 📋 **Paste screenshots.** `Ctrl+V` an image into the note (or use **Paste
  screenshot**) and it's stored as an attachment, with a small `[📎 name]`
  reference dropped into the text — exactly like the Pixel-App.
- 📎 **Attach files.** Pick any file(s); they're kept with the note.
- ✏️ **Manage attachments.** Open, rename, remove, and copy images back to the
  clipboard.
- 💾 **Autosave.** Typing is saved automatically (and immediately when you
  switch mails or hide the pane).
- 🔒 **Fully local & offline.** All notes and attachments live in the browser's
  IndexedDB inside the add-in. There is no server, no account, no cloud.

## How the linking works

The requirement is simple: *every mail must be linked to a note by a unique id —
Mail 1 = Note 1, Mail 2 = Note 2, …*. Here's how that's implemented:

1. When you select a mail, the add-in reads a **stable mail key**. It prefers
   `internetMessageId` (globally unique, survives folder moves and restarts) and
   falls back to the item id / conversation id if that isn't available.
2. A local registry maps `mailKey → noteId`. The **first** time a mail is seen,
   it's assigned the **next sequential number**. That mapping is permanent, so
   the mail always reopens the same note.
3. The note (text + attachment metadata) is stored under that `noteId`, and
   attachment blobs are stored alongside it.

Because the number is assigned on first use, "Mail 1 = Note 1" means *the first
mail you actually take notes on* becomes Note 1 — not a fixed position in your
inbox.

## Architecture

```
┌───────────────────────────── Outlook (classic desktop) ─────────────────────────────┐
│                                                                                      │
│   Reading pane / message           ┌── Task pane (this add-in) ──────────────────┐   │
│   ┌──────────────────────┐         │  Email Notes            🟢 Mail selected    │   │
│   │  From: …             │         │  ┌─────────────────────────────────────┐    │   │
│   │  Subject: …          │  <────  │  │ Note #1   Subject   from Sender      │    │   │
│   │  Body …              │         │  │ [ Paste screenshot ] [ Attach file ] │    │   │
│   │                      │         │  │ ┌─────────────────────────────────┐  │    │   │
│   └──────────────────────┘         │  │ │ note text … [📎 1.png]          │  │    │   │
│                                    │  │ └─────────────────────────────────┘  │    │   │
│                                    │  │ 🖼 1.png  ✏ 📋 ✕                      │    │   │
│                                    │  └─────────────────────────────────────┘    │   │
│                                    └─────────────────────────────────────────────┘   │
└──────────────────────────────────────────────────────────────────────────────────────┘
        Office.js (from Microsoft CDN)  ──►  IndexedDB (local, on your machine)
```

- **`src/core/`** — storage and logic, framework-free:
  - `db.ts` — IndexedDB stores (`meta`, `mailMap`, `notes`, `blobs`) and the
    sequential-ID registry.
  - `mail.ts` — derives the stable mail key from the selected Office item.
  - `attachments.ts` — attachment naming, `[📎 name]` tokens, kinds.
  - `types.ts` — shared types.
- **`src/taskpane/`** — the notepad UI (vanilla TypeScript, no framework) that
  binds to the selected mail and reacts to Office's `ItemChanged` event.
- **`src/commands/`** — the ribbon function file.
- **`manifest.xml`** — the Office add-in manifest.

Office.js is always loaded from Microsoft's official CDN at runtime; it is never
bundled or redistributed by this project.

## Requirements

- Classic Outlook for Windows (desktop), Microsoft 365 / Exchange account.
  Web add-ins also load in new Outlook and Outlook on the web.
- The add-in needs the **Mailbox 1.5** requirement set (add-in commands + task
  pane pinning), which classic Outlook has supported for years.
- [Node.js](https://nodejs.org/) 18+ to build.

## Getting started (development)

```bash
# 1. Install dependencies
npm install

# 2. Trust a local HTTPS certificate (Office requires HTTPS)
npm run dev-certs

# 3. Start the dev server on https://localhost:3000
npm start
```

Then **sideload** `manifest.xml` into Outlook:

- **Classic Outlook (desktop):** File → Manage Add-ins / Get Add-ins → *My
  add-ins* → *Add a custom add-in* → *Add from file…* → pick `manifest.xml`.
- Or use Microsoft's
  [sideloading instructions](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing).

Open a mail, click **Notes** on the ribbon, and **pin** the pane (the pin icon
in the task pane header) so it stays open as you move between mails.

### Useful scripts

| Script | What it does |
| --- | --- |
| `npm start` | Dev server with hot reload on `https://localhost:3000` |
| `npm run build` | Production build into `dist/` |
| `npm run typecheck` | TypeScript type-check only |
| `npm run validate` | Validate `manifest.xml` (uses Microsoft's online validator) |
| `npm run make-icons` | Regenerate the PNG icons in `assets/` |

## Deploying for real use

The manifest points at `https://localhost:3000` for development. To use the
add-in day-to-day, host the built `dist/` folder on any HTTPS static host (for
example GitHub Pages) and replace every `https://localhost:3000` occurrence in
`manifest.xml` with your hosting URL. Then sideload the updated manifest, or have
a Microsoft 365 admin deploy it centrally via
[Integrated Apps](https://learn.microsoft.com/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps).

## Privacy & data

- **Everything stays on your device.** Notes and attachments are stored in the
  add-in's local IndexedDB. There is no backend and the add-in makes no network
  calls of its own.
- The only external resource loaded is **office.js from Microsoft's CDN**, which
  is required for any Office add-in to function.
- The add-in requests the minimum permission, **`ReadItem`**: it reads the
  selected mail's id, subject, and sender to link and label the note. It does
  **not** modify, move, send, or delete your mail.
- Because storage is local to the machine and Outlook profile, notes are not
  synced across devices. Clearing the add-in's data or the Outlook cache removes
  them.

## Legal & Microsoft compliance

This project is built to follow Microsoft's rules for Office/Outlook add-ins:

- It uses the **official Office Add-ins platform** and a standard add-in
  manifest — the supported, documented way to extend Outlook.
- **office.js is loaded from the official Microsoft CDN** and never bundled or
  repackaged, as Microsoft requires.
- **All add-in resource URLs are HTTPS**, per Microsoft's requirements.
- It requests the **least privilege** it needs (`ReadItem`).
- **Trademarks.** "Microsoft", "Outlook", "Office", and related names are
  trademarks of Microsoft Corporation. This project is an **independent,
  unofficial** add-in and is **not affiliated with, endorsed by, or sponsored
  by Microsoft**. The product name ("Email Notes") does not use Microsoft brand
  names, and the icons are original.
- **Publishing to AppSource** (optional): listing an add-in on Microsoft
  AppSource requires passing Microsoft's
  [Commercial marketplace certification policies](https://learn.microsoft.com/legal/marketplace/certification-policies)
  and the
  [Office Store validation policies](https://learn.microsoft.com/office/dev/store/validation-policies).
  Sideloading or admin-deploying the add-in inside your own organization does
  not require certification. This repository targets self-hosting/sideloading;
  anyone wishing to publish it should review those policies first.

If you believe anything here conflicts with Microsoft's current requirements,
please open an issue.

## Project structure

```
Outlook-Notes/
├─ manifest.xml            # Office add-in manifest (pinnable task pane, MessageRead)
├─ assets/                 # Generated PNG icons (see scripts/make-icons.mjs)
├─ scripts/make-icons.mjs  # Reproducible icon generator (no image libraries)
├─ src/
│  ├─ core/                # Storage + logic (db, mail key, attachments, types)
│  ├─ taskpane/            # Notepad UI (HTML/CSS/TS)
│  └─ commands/            # Ribbon function file
├─ webpack.config.js
├─ tsconfig.json
└─ package.json
```

## Not included (yet)

To keep this a clean, standalone open-source add-in, the Pixel-App's
integrations (GitLab issues, the shared To-Do board) are intentionally left out —
they depend on that app's own backend. The core email-notes experience
(per-mail notes, sequential linking, screenshots, attachments) is complete.

## Contributing

Issues and pull requests are welcome. Please keep documentation in English and
run `npm run typecheck` and `npm run build` before submitting.

## License

[MIT](./LICENSE) © Pixelschmied
