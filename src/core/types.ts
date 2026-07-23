/** Shared types for the Email Notes add-in. */

export type AttachmentKind = 'image' | 'other'

/** One attachment of a note. The binary lives in its own IndexedDB store. */
export interface NoteAttachment {
  name: string
  kind: AttachmentKind
  /** MIME type, used to render/open the blob later. */
  mime: string
  /** Byte size, for display only. */
  size: number
}

/**
 * A note is bound to a mail through a stable `mailKey` and given a friendly
 * sequential number (`noteId`): the first mail you take notes on is Note 1,
 * the next distinct mail is Note 2, and so on.
 */
export interface Note {
  noteId: number
  mailKey: string
  subject: string
  sender: string
  text: string
  attachments: NoteAttachment[]
  createdAt: string
  updatedAt: string
}

/** What the task pane currently knows about the selected mail. */
export type Selection =
  | { status: 'ok'; mailKey: string; subject: string; sender: string }
  | { status: 'none' }
  | { status: 'unsupported' }
