import type { NoteAttachment } from './types'

const IMAGE_EXT = new Set(['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.tif', '.tiff', '.heic', '.svg'])

export function extOf(name: string): string {
  const i = name.lastIndexOf('.')
  return i > 0 ? name.slice(i).toLowerCase() : ''
}

export function stemOf(name: string): string {
  const i = name.lastIndexOf('.')
  return i > 0 ? name.slice(0, i) : name
}

export function kindOf(name: string, mime?: string): NoteAttachment['kind'] {
  if (mime?.startsWith('image/')) return 'image'
  return IMAGE_EXT.has(extOf(name)) ? 'image' : 'other'
}

/** The inline reference we drop into the note text for an attachment. */
export function token(name: string): string {
  return `[📎 ${name}]`
}

/**
 * Next sequential screenshot name for a note ("1.png", "2.png", …), mirroring
 * the numbering used by the Pixel-App so pasted screenshots stay tidy.
 */
export function nextImageName(existing: NoteAttachment[]): string {
  const taken = new Set(existing.map((a) => a.name))
  let n = 1
  while (taken.has(`${n}.png`)) n++
  return `${n}.png`
}
