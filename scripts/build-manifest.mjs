// Produces dist/manifest.xml pointing at the public HTTPS host instead of the
// localhost dev URL. The host comes from PUBLIC_URL (falls back to the project's
// GitHub Pages URL). Run after `npm run build`.
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const DEV_URL = 'https://localhost:3000'
const base = (process.env.PUBLIC_URL || 'https://pixelschmied.github.io/Outlook-Notes').replace(/\/+$/, '')

const src = readFileSync(join(root, 'manifest.xml'), 'utf8')
const out = src.split(DEV_URL).join(base)

mkdirSync(join(root, 'dist'), { recursive: true })
writeFileSync(join(root, 'dist', 'manifest.xml'), out)
console.log(`Wrote dist/manifest.xml (host: ${base})`)
