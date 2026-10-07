import { extname } from 'node:path'
import { PDF_VIEWER_ORIGIN } from '@shared/constants.js'

/** Without `electron`, to be testable; `pdf-scheme.ts` serves what this allows. */

export type ViewerRoute =
  | { readonly kind: 'document'; readonly token: string }
  | { readonly kind: 'pdfjs'; readonly name: string }
  | { readonly kind: 'page'; readonly name: string }

export interface ByteRange {
  readonly start: number
  readonly end: number
}

const VIEWER = new URL(PDF_VIEWER_ORIGIN)
const DOCUMENT_PREFIX = '/document/'
const PDFJS_PREFIX = '/pdfjs/'
const TOKEN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/

/** What pdf.js fetches besides its code: CMaps, the 14 standard fonts, image decoders, ICC profiles. */
const PDFJS_FOLDERS: readonly string[] = ['cmaps', 'standard_fonts', 'wasm', 'iccs']

/** The viewer page and the bundle it shares with the app. */
const PAGE_FOLDERS: readonly string[] = ['pdf-viewer', 'assets']

const CONTENT_TYPES: Readonly<Record<string, string>> = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.wasm': 'application/wasm',
  '.ttf': 'font/ttf',
  '.pdf': 'application/pdf',
}

/** Compares scheme and host: Node gives `origin` "null" to a scheme it does not know. */
export function viewerRoute(rawUrl: string): ViewerRoute | null {
  let url: URL
  let path: string
  try {
    url = new URL(rawUrl)
    path = decodeURIComponent(url.pathname)
  } catch {
    return null
  }
  if (url.protocol !== VIEWER.protocol || url.host !== VIEWER.host) return null

  if (path.startsWith(DOCUMENT_PREFIX)) {
    const token = path.slice(DOCUMENT_PREFIX.length)
    return TOKEN.test(token) ? { kind: 'document', token } : null
  }
  if (path.startsWith(PDFJS_PREFIX)) {
    const name = path.slice(PDFJS_PREFIX.length)
    return PDFJS_FOLDERS.includes(firstSegment(name)) ? { kind: 'pdfjs', name } : null
  }

  const name = path.replace(/^\/+/, '')
  return PAGE_FOLDERS.includes(firstSegment(name)) ? { kind: 'page', name } : null
}

export function contentTypeOf(name: string): string {
  return CONTENT_TYPES[extname(name).toLowerCase()] ?? 'application/octet-stream'
}

/**
 * The single range pdf.js asks for (`bytes=start-end` or `bytes=start-`). `null` means the whole
 * file; `'unsatisfiable'` a range the file does not have.
 */
export function parseByteRange(header: string | null, size: number): ByteRange | null | 'unsatisfiable' {
  if (header === null) return null
  const match = /^bytes=(\d+)-(\d*)$/.exec(header.trim())
  if (match === null) return null

  const start = Number(match[1])
  const end = match[2] === '' ? size - 1 : Math.min(Number(match[2]), size - 1)
  if (start >= size || end < start) return 'unsatisfiable'
  return { start, end }
}

function firstSegment(name: string): string {
  return name.split('/', 1)[0] ?? ''
}
