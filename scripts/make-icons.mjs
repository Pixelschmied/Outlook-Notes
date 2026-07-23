// Generates the add-in icons (16/32/64/80/128 px) as PNGs into ./assets.
//
// Self-contained: no image libraries. It rasterises a small "note" glyph on a
// rounded indigo tile and writes valid RGBA PNGs. Re-run with `npm run make-icons`.
import { deflateSync } from 'node:zlib'
import { writeFileSync, mkdirSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'

const OUT = join(dirname(fileURLToPath(import.meta.url)), '..', 'assets')
mkdirSync(OUT, { recursive: true })

const BG = [124, 92, 246] // indigo tile
const SHEET = [246, 247, 251] // note sheet
const LINE = [160, 174, 192] // text lines
const ACCENT = [124, 92, 246] // clip dot (matches tile)

/** CRC32 for PNG chunks. */
function crc32(buf) {
  let c = ~0
  for (let i = 0; i < buf.length; i++) {
    c ^= buf[i]
    for (let k = 0; k < 8; k++) c = (c >>> 1) ^ (0xedb88320 & -(c & 1))
  }
  return ~c >>> 0
}

function chunk(type, data) {
  const len = Buffer.alloc(4)
  len.writeUInt32BE(data.length, 0)
  const typeBuf = Buffer.from(type, 'ascii')
  const body = Buffer.concat([typeBuf, data])
  const crc = Buffer.alloc(4)
  crc.writeUInt32BE(crc32(body), 0)
  return Buffer.concat([len, body, crc])
}

function encodePng(size, pixels) {
  const sig = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10])
  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(size, 0)
  ihdr.writeUInt32BE(size, 4)
  ihdr[8] = 8 // bit depth
  ihdr[9] = 6 // colour type RGBA
  // 10-12 already 0 (compression/filter/interlace)
  const raw = Buffer.alloc(size * (size * 4 + 1))
  for (let y = 0; y < size; y++) {
    raw[y * (size * 4 + 1)] = 0 // filter: none
    for (let x = 0; x < size; x++) {
      const src = (y * size + x) * 4
      const dst = y * (size * 4 + 1) + 1 + x * 4
      raw[dst] = pixels[src]
      raw[dst + 1] = pixels[src + 1]
      raw[dst + 2] = pixels[src + 2]
      raw[dst + 3] = pixels[src + 3]
    }
  }
  return Buffer.concat([
    sig,
    chunk('IHDR', ihdr),
    chunk('IDAT', deflateSync(raw, { level: 9 })),
    chunk('IEND', Buffer.alloc(0))
  ])
}

function draw(size) {
  const px = Buffer.alloc(size * size * 4) // transparent
  const set = (x, y, [r, g, b], a = 255) => {
    if (x < 0 || y < 0 || x >= size || y >= size) return
    const i = (y * size + x) * 4
    px[i] = r
    px[i + 1] = g
    px[i + 2] = b
    px[i + 3] = a
  }
  const radius = Math.max(2, Math.round(size * 0.22))
  const inCorner = (x, y) => {
    // rounded-rectangle mask
    const cx = x < radius ? radius : x >= size - radius ? size - radius - 1 : x
    const cy = y < radius ? radius : y >= size - radius ? size - radius - 1 : y
    const dx = x - cx
    const dy = y - cy
    return dx * dx + dy * dy <= radius * radius
  }
  // Tile background
  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      if (inCorner(x, y)) set(x, y, BG)
    }
  }
  // Note sheet inset
  const m = Math.round(size * 0.24)
  const sx0 = m
  const sy0 = Math.round(size * 0.2)
  const sx1 = size - m
  const sy1 = size - Math.round(size * 0.18)
  for (let y = sy0; y < sy1; y++) {
    for (let x = sx0; x < sx1; x++) set(x, y, SHEET)
  }
  // Text lines on the sheet
  const lineCount = size >= 64 ? 3 : 2
  const gap = Math.max(2, Math.round((sy1 - sy0) / (lineCount + 1.5)))
  const th = Math.max(1, Math.round(size * 0.04))
  for (let l = 1; l <= lineCount; l++) {
    const ly = sy0 + gap * l
    const lx1 = l === lineCount ? sx0 + Math.round((sx1 - sx0) * 0.55) : sx1 - Math.round(size * 0.08)
    for (let t = 0; t < th; t++) {
      for (let x = sx0 + Math.round(size * 0.08); x < lx1; x++) set(x, ly + t, LINE)
    }
  }
  // Little accent clip dot in the top-right of the tile
  const cx = size - Math.round(size * 0.28)
  const cy = Math.round(size * 0.28)
  const cr = Math.max(1, Math.round(size * 0.07))
  for (let y = -cr; y <= cr; y++) {
    for (let x = -cr; x <= cr; x++) {
      if (x * x + y * y <= cr * cr) set(cx + x, cy + y, ACCENT)
    }
  }
  return px
}

for (const size of [16, 32, 64, 80, 128]) {
  const png = encodePng(size, draw(size))
  writeFileSync(join(OUT, `icon-${size}.png`), png)
  console.log(`assets/icon-${size}.png (${png.length} bytes)`)
}
