import type { Note, Selection } from '../core/types'
import { readSelection } from '../core/mail'
import { kindOf, nextImageName, stemOf, token } from '../core/attachments'
import {
  addAttachment,
  getAttachmentBlob,
  getOrCreateNote,
  listNotes,
  removeAttachment,
  renameAttachment,
  saveNoteText
} from '../core/db'

/**
 * Task pane controller. It keeps the notepad bound to the mail currently
 * selected in Outlook: when the selection changes (Office `ItemChanged`), it
 * flushes the open note and loads the one belonging to the new mail. All state
 * lives locally in IndexedDB — see core/db.ts.
 */

const SAVE_DEBOUNCE_MS = 400

// --- DOM handles ---
const el = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T
const statusEl = el<HTMLSpanElement>('status')
const emptyEl = el<HTMLElement>('empty')
const editorEl = el<HTMLElement>('editor')
const badgeEl = el<HTMLSpanElement>('note-badge')
const subjectEl = el<HTMLSpanElement>('subject')
const senderEl = el<HTMLSpanElement>('sender')
const noteEl = el<HTMLTextAreaElement>('note')
const attachmentsEl = el<HTMLDivElement>('attachments')
const savedEl = el<HTMLSpanElement>('saved')
const toastEl = el<HTMLDivElement>('toast')

// --- current state ---
let current: Note | null = null
let saveTimer: ReturnType<typeof setTimeout> | null = null
let toastTimer: ReturnType<typeof setTimeout> | null = null

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) {
    setStatus('This add-in runs inside Outlook.', 'warn')
    return
  }
  wireEvents()
  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => void refresh())
  void refresh()
})

function wireEvents(): void {
  noteEl.addEventListener('input', () => {
    if (!current) return
    current.text = noteEl.value
    scheduleSave()
  })
  noteEl.addEventListener('paste', onPaste)
  el('btn-screenshot').addEventListener('click', () => void pasteFromClipboard())
  el('btn-files').addEventListener('click', () => void pickFiles())
  el('btn-export').addEventListener('click', () => void copyNoteText())
  // Save promptly if the pane is about to be hidden.
  window.addEventListener('pagehide', () => flushSave())
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'hidden') flushSave()
  })
}

// --- selection handling ---

async function refresh(): Promise<void> {
  const sel: Selection = readSelection()

  if (sel.status !== 'ok') {
    flushSave()
    current = null
    editorEl.classList.add('hidden')
    emptyEl.classList.remove('hidden')
    if (sel.status === 'unsupported') {
      setStatus('This item has no stable id', 'warn')
    } else {
      setStatus('No mail selected', 'idle')
    }
    return
  }

  // A different mail was selected → persist the previous note first.
  if (current && current.mailKey !== sel.mailKey) flushSave()

  setStatus('Mail selected', 'live')
  const note = await getOrCreateNote(sel.mailKey, sel.subject, sel.sender)
  // Guard against a race where the selection changed again mid-load.
  const now = readSelection()
  if (now.status !== 'ok' || now.mailKey !== sel.mailKey) return

  current = note
  render(note)
}

function render(note: Note): void {
  emptyEl.classList.add('hidden')
  editorEl.classList.remove('hidden')
  badgeEl.textContent = `Note #${note.noteId}`
  subjectEl.textContent = note.subject || '(no subject)'
  senderEl.textContent = note.sender ? `from ${note.sender}` : ''
  noteEl.value = note.text
  savedEl.textContent = 'Saved automatically'
  renderAttachments(note)
}

function renderAttachments(note: Note): void {
  attachmentsEl.textContent = ''
  for (const att of note.attachments) {
    const chip = document.createElement('span')
    chip.className = 'chip'

    const open = document.createElement('button')
    open.className = 'chip-open'
    open.title = `${att.name} — open`
    open.textContent = `${att.kind === 'image' ? '🖼' : '📎'} ${att.name}`
    open.addEventListener('click', () => void openAttachment(att.name))
    chip.appendChild(open)

    const rename = iconButton('✏️', 'Rename', () => startRename(chip, att.name))
    rename.className = 'chip-act'
    chip.appendChild(rename)

    if (att.kind === 'image') {
      const copy = iconButton('📋', 'Copy image to clipboard', () => void copyImage(att.name, att.mime))
      copy.className = 'chip-act'
      chip.appendChild(copy)
    }

    const del = iconButton('✕', 'Remove attachment', () => void deleteAttachment(att.name))
    del.className = 'chip-del'
    chip.appendChild(del)

    attachmentsEl.appendChild(chip)
  }
}

function iconButton(label: string, title: string, onClick: () => void): HTMLButtonElement {
  const b = document.createElement('button')
  b.type = 'button'
  b.textContent = label
  b.title = title
  b.addEventListener('click', onClick)
  return b
}

// --- saving ---

function scheduleSave(): void {
  savedEl.textContent = 'Saving…'
  if (saveTimer) clearTimeout(saveTimer)
  saveTimer = setTimeout(() => {
    void doSave().then(() => {
      savedEl.textContent = `Saved ${new Date().toLocaleTimeString()}`
    })
  }, SAVE_DEBOUNCE_MS)
}

function flushSave(): void {
  if (saveTimer) {
    clearTimeout(saveTimer)
    saveTimer = null
  }
  void doSave()
}

async function doSave(): Promise<void> {
  if (!current) return
  await saveNoteText(current.noteId, current.text, current.subject, current.sender)
}

// --- attachments: clipboard, files ---

function onPaste(e: ClipboardEvent): void {
  const items = e.clipboardData?.items
  if (!items) return
  const image = Array.from(items).find((it) => it.type.startsWith('image/'))
  if (!image) return // let normal text paste through untouched
  e.preventDefault()
  const file = image.getAsFile()
  if (file) void attachImageBlob(file, file.type)
}

/**
 * Explicit "Paste screenshot" button. Uses the async Clipboard API where the
 * host grants it; otherwise it nudges the user to press Ctrl+V (which the paste
 * handler above catches reliably).
 */
async function pasteFromClipboard(): Promise<void> {
  if (!current) return
  try {
    const clip = navigator.clipboard as Clipboard & {
      read?: () => Promise<ClipboardItem[]>
    }
    if (!clip.read) throw new Error('no clipboard read')
    const items = await clip.read()
    for (const it of items) {
      const type = it.types.find((t) => t.startsWith('image/'))
      if (type) {
        const blob = await it.getType(type)
        await attachImageBlob(blob, type)
        return
      }
    }
    toast('No image on the clipboard.')
  } catch {
    noteEl.focus()
    toast('Press Ctrl+V to paste the screenshot here.')
  }
}

async function attachImageBlob(blob: Blob, mime: string): Promise<void> {
  if (!current) return
  const name = nextImageName(current.attachments)
  const { note, name: finalName } = await addAttachment(current.noteId, name, 'image', mime || 'image/png', blob)
  current = note
  renderAttachments(note)
  insertToken(finalName)
}

async function pickFiles(): Promise<void> {
  if (!current) return
  const input = document.createElement('input')
  input.type = 'file'
  input.multiple = true
  input.addEventListener('change', () => {
    const files = input.files
    if (files) void addFiles(Array.from(files))
  })
  input.click()
}

async function addFiles(files: File[]): Promise<void> {
  if (!current) return
  for (const file of files) {
    const kind = kindOf(file.name, file.type)
    const { note, name } = await addAttachment(
      current.noteId,
      file.name || 'file',
      kind,
      file.type || 'application/octet-stream',
      file
    )
    current = note
    insertToken(name)
  }
  renderAttachments(current)
}

// --- attachments: open, copy, rename, remove ---

async function openAttachment(name: string): Promise<void> {
  if (!current) return
  const blob = await getAttachmentBlob(current.noteId, name)
  if (!blob) return
  const url = URL.createObjectURL(blob)
  window.open(url, '_blank', 'noopener')
  setTimeout(() => URL.revokeObjectURL(url), 60_000)
}

async function copyImage(name: string, mime: string): Promise<void> {
  if (!current) return
  const blob = await getAttachmentBlob(current.noteId, name)
  if (!blob) return
  try {
    const writer = navigator.clipboard as Clipboard & {
      write?: (items: ClipboardItem[]) => Promise<void>
    }
    if (!writer.write || typeof ClipboardItem === 'undefined') throw new Error('no clipboard write')
    await writer.write([new ClipboardItem({ [mime || blob.type || 'image/png']: blob })])
    toast('Image copied to the clipboard.')
  } catch {
    toast('Could not copy — open the image and copy it there.')
  }
}

async function deleteAttachment(name: string): Promise<void> {
  if (!current) return
  const note = await removeAttachment(current.noteId, name)
  if (!note) return
  current = note
  // Drop the matching inline token from the text as well.
  const tk = token(name)
  if (noteEl.value.includes(tk)) {
    noteEl.value = noteEl.value.split(tk).join('').replace(/[ \t]{2,}/g, ' ')
    current.text = noteEl.value
    scheduleSave()
  }
  renderAttachments(note)
}

function startRename(chip: HTMLElement, name: string): void {
  const open = chip.querySelector('.chip-open')
  if (!open) return
  const input = document.createElement('input')
  input.className = 'chip-edit'
  input.value = stemOf(name)
  let cancelled = false
  const commit = async (): Promise<void> => {
    if (cancelled || !current) {
      renderAttachments(current!)
      return
    }
    const desired = input.value.trim()
    if (!desired) {
      renderAttachments(current)
      return
    }
    const res = await renameAttachment(current.noteId, name, desired)
    if (res) {
      current = res.note
      if (res.name !== name && noteEl.value.includes(token(name))) {
        noteEl.value = noteEl.value.split(token(name)).join(token(res.name))
        current.text = noteEl.value
        scheduleSave()
      }
    }
    renderAttachments(current)
  }
  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') input.blur()
    else if (e.key === 'Escape') {
      cancelled = true
      input.blur()
    }
  })
  input.addEventListener('blur', () => void commit())
  open.replaceWith(input)
  input.focus()
  input.select()
}

// --- text helpers ---

function insertToken(name: string): void {
  const tk = token(name)
  const start = noteEl.selectionStart ?? noteEl.value.length
  const end = noteEl.selectionEnd ?? start
  const before = noteEl.value.slice(0, start)
  const pad = before && !/\s$/.test(before) ? ' ' : ''
  const insert = `${pad}${tk} `
  noteEl.value = before + insert + noteEl.value.slice(end)
  const caret = start + insert.length
  noteEl.focus()
  noteEl.setSelectionRange(caret, caret)
  if (current) {
    current.text = noteEl.value
    scheduleSave()
  }
}

async function copyNoteText(): Promise<void> {
  if (!current) return
  const header = `Note #${current.noteId} — ${current.subject || '(no subject)'}`
  const body = `${header}\n\n${current.text}`.trim()
  try {
    await navigator.clipboard.writeText(body)
    toast('Note text copied.')
  } catch {
    toast('Could not access the clipboard.')
  }
}

// --- misc UI ---

function setStatus(text: string, tone: 'live' | 'idle' | 'warn'): void {
  statusEl.textContent = text
  statusEl.className = `status ${tone}`
}

function toast(message: string): void {
  toastEl.textContent = message
  toastEl.classList.remove('hidden')
  if (toastTimer) clearTimeout(toastTimer)
  toastTimer = setTimeout(() => toastEl.classList.add('hidden'), 3000)
}

// Exposed for debugging in the pane's dev tools (e.g. to inspect all notes).
;(window as unknown as { emailNotesDump: () => Promise<Note[]> }).emailNotesDump = listNotes
