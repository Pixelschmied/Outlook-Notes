import type { Note, NoteAttachment } from './types'

/**
 * Local, offline storage for notes and their attachments, built on IndexedDB.
 * Nothing here ever leaves the device — there is no server and no network call.
 *
 * Stores:
 *  - `meta`   : bookkeeping (the running note counter).
 *  - `mailMap`: mailKey -> noteId. This is what makes "Mail 1 = Note 1" stable:
 *               a mail always resolves to the same note number, for good.
 *  - `notes`  : noteId -> Note (text + attachment metadata).
 *  - `blobs`  : "<noteId>/<name>" -> Blob (the actual image/file bytes).
 */

const DB_NAME = 'email-notes'
const DB_VERSION = 1
const COUNTER_KEY = 'noteCounter'

let dbPromise: Promise<IDBDatabase> | null = null

function openDb(): Promise<IDBDatabase> {
  if (dbPromise) return dbPromise
  dbPromise = new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION)
    req.onupgradeneeded = () => {
      const db = req.result
      if (!db.objectStoreNames.contains('meta')) db.createObjectStore('meta')
      if (!db.objectStoreNames.contains('mailMap')) db.createObjectStore('mailMap')
      if (!db.objectStoreNames.contains('notes')) db.createObjectStore('notes', { keyPath: 'noteId' })
      if (!db.objectStoreNames.contains('blobs')) db.createObjectStore('blobs')
    }
    req.onsuccess = () => resolve(req.result)
    req.onerror = () => reject(req.error ?? new Error('IndexedDB open failed'))
  })
  return dbPromise
}

function tx<T>(
  stores: string | string[],
  mode: IDBTransactionMode,
  run: (t: IDBTransaction) => Promise<T> | T
): Promise<T> {
  return openDb().then(
    (db) =>
      new Promise<T>((resolve, reject) => {
        const t = db.transaction(stores, mode)
        let result: T
        Promise.resolve(run(t)).then(
          (r) => {
            result = r
          },
          reject
        )
        t.oncomplete = () => resolve(result)
        t.onerror = () => reject(t.error ?? new Error('IndexedDB transaction failed'))
        t.onabort = () => reject(t.error ?? new Error('IndexedDB transaction aborted'))
      })
  )
}

function reqAsync<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result)
    request.onerror = () => reject(request.error ?? new Error('IndexedDB request failed'))
  })
}

function blobKey(noteId: number, name: string): string {
  return `${noteId}/${name}`
}

function now(): string {
  return new Date().toISOString()
}

/**
 * Resolve a mail to its note, creating an empty note (with the next sequential
 * number) the first time a given mail is seen. Subject/sender are refreshed so
 * the header always shows the current mail's details.
 */
export async function getOrCreateNote(
  mailKey: string,
  subject: string,
  sender: string
): Promise<Note> {
  return tx(['meta', 'mailMap', 'notes'], 'readwrite', async (t) => {
    const mailMap = t.objectStore('mailMap')
    const notes = t.objectStore('notes')
    const meta = t.objectStore('meta')

    let noteId = (await reqAsync(mailMap.get(mailKey))) as number | undefined

    if (noteId === undefined) {
      const counter = ((await reqAsync(meta.get(COUNTER_KEY))) as number | undefined) ?? 0
      noteId = counter + 1
      meta.put(noteId, COUNTER_KEY)
      mailMap.put(noteId, mailKey)
    }

    const existing = (await reqAsync(notes.get(noteId))) as Note | undefined
    if (existing) {
      // Keep the display metadata fresh without touching the note body.
      const updated: Note = { ...existing, subject: subject || existing.subject, sender: sender || existing.sender }
      notes.put(updated)
      return updated
    }

    const created: Note = {
      noteId,
      mailKey,
      subject,
      sender,
      text: '',
      attachments: [],
      createdAt: now(),
      updatedAt: now()
    }
    notes.put(created)
    return created
  })
}

/** Save the note text (metadata kept for the header). */
export async function saveNoteText(
  noteId: number,
  text: string,
  subject: string,
  sender: string
): Promise<void> {
  await tx('notes', 'readwrite', async (t) => {
    const notes = t.objectStore('notes')
    const cur = (await reqAsync(notes.get(noteId))) as Note | undefined
    if (!cur) return
    notes.put({ ...cur, text, subject: subject || cur.subject, sender: sender || cur.sender, updatedAt: now() })
  })
}

/** Store an attachment blob and record its metadata on the note. */
export async function addAttachment(
  noteId: number,
  desiredName: string,
  kind: NoteAttachment['kind'],
  mime: string,
  blob: Blob
): Promise<{ note: Note; name: string }> {
  return tx(['notes', 'blobs'], 'readwrite', async (t) => {
    const notes = t.objectStore('notes')
    const blobs = t.objectStore('blobs')
    const cur = (await reqAsync(notes.get(noteId))) as Note | undefined
    if (!cur) throw new Error('Note not found')

    const name = uniqueName(cur.attachments, desiredName)
    const att: NoteAttachment = { name, kind, mime, size: blob.size }
    const note: Note = { ...cur, attachments: [...cur.attachments, att], updatedAt: now() }
    notes.put(note)
    blobs.put(blob, blobKey(noteId, name))
    return { note, name }
  })
}

export async function getAttachmentBlob(noteId: number, name: string): Promise<Blob | undefined> {
  return tx('blobs', 'readonly', (t) => reqAsync(t.objectStore('blobs').get(blobKey(noteId, name)))) as Promise<
    Blob | undefined
  >
}

export async function removeAttachment(noteId: number, name: string): Promise<Note | undefined> {
  return tx(['notes', 'blobs'], 'readwrite', async (t) => {
    const notes = t.objectStore('notes')
    const blobs = t.objectStore('blobs')
    const cur = (await reqAsync(notes.get(noteId))) as Note | undefined
    if (!cur) return undefined
    const note: Note = {
      ...cur,
      attachments: cur.attachments.filter((a) => a.name !== name),
      updatedAt: now()
    }
    notes.put(note)
    blobs.delete(blobKey(noteId, name))
    return note
  })
}

export async function renameAttachment(
  noteId: number,
  oldName: string,
  desiredName: string
): Promise<{ note: Note; name: string } | undefined> {
  return tx(['notes', 'blobs'], 'readwrite', async (t) => {
    const notes = t.objectStore('notes')
    const blobs = t.objectStore('blobs')
    const cur = (await reqAsync(notes.get(noteId))) as Note | undefined
    if (!cur) return undefined
    const att = cur.attachments.find((a) => a.name === oldName)
    if (!att) return undefined

    const others = cur.attachments.filter((a) => a.name !== oldName)
    const name = uniqueName(others, ensureExt(desiredName, oldName))
    if (name === oldName) return { note: cur, name: oldName }

    const blob = (await reqAsync(blobs.get(blobKey(noteId, oldName)))) as Blob | undefined
    if (blob) {
      blobs.put(blob, blobKey(noteId, name))
      blobs.delete(blobKey(noteId, oldName))
    }
    const note: Note = {
      ...cur,
      attachments: cur.attachments.map((a) => (a.name === oldName ? { ...a, name } : a)),
      updatedAt: now()
    }
    notes.put(note)
    return { note, name }
  })
}

/** All notes, newest first — used by the export/overview action. */
export async function listNotes(): Promise<Note[]> {
  const all = (await tx('notes', 'readonly', (t) => reqAsync(t.objectStore('notes').getAll()))) as Note[]
  return all.sort((a, b) => b.noteId - a.noteId)
}

// ---- helpers ----

function splitExt(name: string): [string, string] {
  const i = name.lastIndexOf('.')
  return i > 0 ? [name.slice(0, i), name.slice(i)] : [name, '']
}

function ensureExt(desired: string, fallbackFrom: string): string {
  const clean = desired.replace(/[\\/:*?"<>|]/g, '').trim() || 'file'
  const [, ext] = splitExt(clean)
  if (ext) return clean
  const [, oldExt] = splitExt(fallbackFrom)
  return clean + oldExt
}

function uniqueName(existing: NoteAttachment[], desired: string): string {
  const taken = new Set(existing.map((a) => a.name))
  if (!taken.has(desired)) return desired
  const [stem, ext] = splitExt(desired)
  let i = 1
  let name = `${stem} (${i})${ext}`
  while (taken.has(name)) {
    i++
    name = `${stem} (${i})${ext}`
  }
  return name
}
