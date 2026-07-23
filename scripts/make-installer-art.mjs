// Generates the installer artwork used by Inno Setup:
//   installer/wizard-large.bmp  (164x314) — the left banner of the wizard
//   installer/wizard-small.bmp  (55x58)   — the small header logo
//   assets/app.ico                        — installer + uninstall icon
//
// Self-contained: no image libraries. The BMPs are 24-bit uncompressed; the ICO
// simply embeds the existing PNG icons. Re-run with `npm run make-installer-art`.
import { writeFileSync, mkdirSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const ASSETS = join(root, 'assets')
const INSTALLER = join(root, 'installer')
mkdirSync(INSTALLER, { recursive: true })

// --- tiny drawing surface (top-down RGB) ---
function surface(w, h) {
  const buf = new Uint8Array(w * h * 3)
  return {
    w,
    h,
    buf,
    set(x, y, r, g, b) {
      if (x < 0 || y < 0 || x >= w || y >= h) return
      const i = (y * w + x) * 3
      buf[i] = r
      buf[i + 1] = g
      buf[i + 2] = b
    }
  }
}

function lerp(a, b, t) {
  return Math.round(a + (b - a) * t)
}

// Rounded-rectangle fill helper.
function roundRect(s, x0, y0, x1, y1, radius, color) {
  const [r, g, b] = color
  for (let y = y0; y < y1; y++) {
    for (let x = x0; x < x1; x++) {
      const nx = x < x0 + radius ? x0 + radius : x >= x1 - radius ? x1 - radius - 1 : x
      const ny = y < y0 + radius ? y0 + radius : y >= y1 - radius ? y1 - radius - 1 : y
      const dx = x - nx
      const dy = y - ny
      if (dx * dx + dy * dy <= radius * radius) s.set(x, y, r, g, b)
    }
  }
}

function disc(s, cx, cy, cr, color) {
  const [r, g, b] = color
  for (let y = -cr; y <= cr; y++) {
    for (let x = -cr; x <= cr; x++) {
      if (x * x + y * y <= cr * cr) s.set(cx + x, cy + y, r, g, b)
    }
  }
}

// Draw the indigo gradient background + a "note sheet" glyph.
function drawScene(w, h, opts) {
  const s = surface(w, h)
  // Diagonal indigo gradient.
  const topL = [141, 108, 247]
  const botR = [79, 63, 214]
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const t = (x / w + y / h) / 2
      s.set(x, y, lerp(topL[0], botR[0], t), lerp(topL[1], botR[1], t), lerp(topL[2], botR[2], t))
    }
  }
  // Soft accent dots for a modern feel.
  disc(s, Math.round(w * 0.85), Math.round(h * 0.12), Math.round(w * 0.06), [157, 132, 248])
  disc(s, Math.round(w * 0.15), Math.round(h * 0.9), Math.round(w * 0.05), [110, 92, 230])

  // Note sheet.
  const g = opts.glyph
  const gw = Math.round(w * g.width)
  const gh = Math.round(gw * 1.18)
  const gx = Math.round(w * g.cx - gw / 2)
  const gy = Math.round(h * g.cy - gh / 2)
  const rad = Math.max(3, Math.round(gw * 0.1))
  roundRect(s, gx, gy, gx + gw, gy + gh, rad, [246, 247, 251])
  // Text lines.
  const lines = 4
  const pad = Math.round(gw * 0.16)
  const lh = Math.max(1, Math.round(gh * 0.045))
  const step = Math.round((gh - pad * 2) / (lines + 1))
  for (let l = 1; l <= lines; l++) {
    const ly = gy + pad + step * l
    const lx1 = l === lines ? gx + Math.round(gw * 0.55) : gx + gw - pad
    for (let t = 0; t < lh; t++) {
      for (let x = gx + pad; x < lx1; x++) s.set(x, ly + t, 160, 174, 192)
    }
  }
  // Accent clip dot on the sheet corner.
  disc(s, gx + gw - Math.round(gw * 0.14), gy + Math.round(gw * 0.14), Math.round(gw * 0.1), [124, 92, 246])
  return s
}

// --- 24-bit BMP encoder (bottom-up, BGR, 4-byte row padding) ---
function encodeBmp(s) {
  const { w, h, buf } = s
  const rowSize = Math.floor((24 * w + 31) / 32) * 4
  const pixelBytes = rowSize * h
  const fileSize = 54 + pixelBytes
  const out = Buffer.alloc(fileSize)
  out.write('BM', 0, 'ascii')
  out.writeUInt32LE(fileSize, 2)
  out.writeUInt32LE(54, 10) // pixel data offset
  out.writeUInt32LE(40, 14) // DIB header size
  out.writeInt32LE(w, 18)
  out.writeInt32LE(h, 22) // positive = bottom-up
  out.writeUInt16LE(1, 26) // planes
  out.writeUInt16LE(24, 28) // bpp
  out.writeUInt32LE(2835, 38) // x ppm
  out.writeUInt32LE(2835, 42) // y ppm
  for (let y = 0; y < h; y++) {
    const srcRow = (h - 1 - y) * w * 3 // flip vertically
    let dst = 54 + y * rowSize
    for (let x = 0; x < w; x++) {
      const i = srcRow + x * 3
      out[dst++] = buf[i + 2] // B
      out[dst++] = buf[i + 1] // G
      out[dst++] = buf[i] // R
    }
  }
  return out
}

// Draw the note-tile icon at an arbitrary size as top-down RGBA (with a
// transparent rounded-corner background). Mirrors the PNG icons in look.
function drawTile(size) {
  const px = new Uint8Array(size * size * 4)
  const set = (x, y, r, g, b, a) => {
    if (x < 0 || y < 0 || x >= size || y >= size) return
    const i = (y * size + x) * 4
    px[i] = r
    px[i + 1] = g
    px[i + 2] = b
    px[i + 3] = a
  }
  const radius = Math.max(2, Math.round(size * 0.22))
  const inTile = (x, y) => {
    const cx = x < radius ? radius : x >= size - radius ? size - radius - 1 : x
    const cy = y < radius ? radius : y >= size - radius ? size - radius - 1 : y
    const dx = x - cx
    const dy = y - cy
    return dx * dx + dy * dy <= radius * radius
  }
  for (let y = 0; y < size; y++) for (let x = 0; x < size; x++) if (inTile(x, y)) set(x, y, 124, 92, 246, 255)
  const m = Math.round(size * 0.24)
  const sx0 = m
  const sy0 = Math.round(size * 0.2)
  const sx1 = size - m
  const sy1 = size - Math.round(size * 0.18)
  for (let y = sy0; y < sy1; y++) for (let x = sx0; x < sx1; x++) set(x, y, 246, 247, 251, 255)
  const lineCount = size >= 64 ? 3 : 2
  const gap = Math.max(2, Math.round((sy1 - sy0) / (lineCount + 1.5)))
  const th = Math.max(1, Math.round(size * 0.04))
  for (let l = 1; l <= lineCount; l++) {
    const ly = sy0 + gap * l
    const lx1 = l === lineCount ? sx0 + Math.round((sx1 - sx0) * 0.55) : sx1 - Math.round(size * 0.08)
    for (let t = 0; t < th; t++) for (let x = sx0 + Math.round(size * 0.08); x < lx1; x++) set(x, ly + t, 160, 174, 192, 255)
  }
  const cx = size - Math.round(size * 0.28)
  const cy = Math.round(size * 0.28)
  const cr = Math.max(1, Math.round(size * 0.07))
  for (let y = -cr; y <= cr; y++)
    for (let x = -cr; x <= cr; x++) if (x * x + y * y <= cr * cr) set(cx + x, cy + y, 124, 92, 246, 255)
  return px
}

// One classic (uncompressed, 32-bit BGRA) icon image with an all-zero AND mask.
// This format is the most broadly compatible — Inno Setup and every Windows
// version accept it, unlike PNG-compressed entries.
function icoImage(size) {
  const rgba = drawTile(size)
  const header = Buffer.alloc(40)
  header.writeUInt32LE(40, 0)
  header.writeInt32LE(size, 4)
  header.writeInt32LE(size * 2, 8) // height = XOR + AND
  header.writeUInt16LE(1, 12)
  header.writeUInt16LE(32, 14)
  const xor = Buffer.alloc(size * size * 4)
  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      const s = ((size - 1 - y) * size + x) * 4 // bottom-up
      const d = (y * size + x) * 4
      xor[d] = rgba[s + 2] // B
      xor[d + 1] = rgba[s + 1] // G
      xor[d + 2] = rgba[s] // R
      xor[d + 3] = rgba[s + 3] // A
    }
  }
  const maskRow = Math.floor((size + 31) / 32) * 4
  const and = Buffer.alloc(maskRow * size) // all zero = use alpha
  return Buffer.concat([header, xor, and])
}

function buildIco(sizes) {
  const imgs = sizes.map((sz) => ({ sz, data: icoImage(sz) }))
  const header = Buffer.alloc(6)
  header.writeUInt16LE(0, 0)
  header.writeUInt16LE(1, 2) // type: icon
  header.writeUInt16LE(imgs.length, 4)
  const dir = Buffer.alloc(16 * imgs.length)
  let offset = 6 + dir.length
  imgs.forEach((p, idx) => {
    const e = idx * 16
    dir[e] = p.sz >= 256 ? 0 : p.sz
    dir[e + 1] = p.sz >= 256 ? 0 : p.sz
    dir.writeUInt16LE(1, e + 4)
    dir.writeUInt16LE(32, e + 6)
    dir.writeUInt32LE(p.data.length, e + 8)
    dir.writeUInt32LE(offset, e + 12)
    offset += p.data.length
  })
  return Buffer.concat([header, dir, ...imgs.map((p) => p.data)])
}

const large = encodeBmp(drawScene(164, 314, { glyph: { width: 0.5, cx: 0.5, cy: 0.34 } }))
writeFileSync(join(INSTALLER, 'wizard-large.bmp'), large)
console.log(`installer/wizard-large.bmp (${large.length} bytes)`)

const small = encodeBmp(drawScene(55, 58, { glyph: { width: 0.62, cx: 0.5, cy: 0.46 } }))
writeFileSync(join(INSTALLER, 'wizard-small.bmp'), small)
console.log(`installer/wizard-small.bmp (${small.length} bytes)`)

const ico = buildIco([16, 32, 48, 64, 128])
writeFileSync(join(ASSETS, 'app.ico'), ico)
console.log(`assets/app.ico (${ico.length} bytes)`)
