import { randomUUID } from 'node:crypto'
import { createReadStream } from 'node:fs'
import { readFile, stat } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { Readable } from 'node:stream'
import { fileURLToPath } from 'node:url'
import { net, protocol } from 'electron'
import { PDF_VIEWER_ORIGIN, PDF_VIEWER_SCHEME } from '@shared/constants.js'
import { fontsRoot } from './fonts.js'
import { containedPath } from './fs/contained.js'
import { VIEWER_CONTENT_SECURITY_POLICY } from './security-policy.js'
import { contentTypeOf, parseByteRange, viewerRoute, type ViewerRoute } from './viewer-routes.js'
import { devServerUrl } from './window.js'

/**
 * The PDF viewer runs in a frame of its own origin, served from here: the code that parses the
 * PDF stays in another process, out of reach of `window.api`.
 */

/** Token → path. The frame never sees a path, and reads only the files the app opened for it. */
const sharedDocuments = new Map<string, string>()

/** The URL the viewer fetches the PDF from. The path must already be allowed for viewing. */
export function shareDocument(path: string): string {
  const token = randomUUID()
  sharedDocuments.set(token, path)
  return `${PDF_VIEWER_ORIGIN}/document/${token}`
}

/** After `app.whenReady()`. */
export function serveViewer(): void {
  protocol.handle(PDF_VIEWER_SCHEME, async (request) => {
    const route = routeOf(request.url)
    try {
      const response = route === null ? null : await respond(route, request)
      return withPolicy(response ?? status(404))
    } catch {
      return withPolicy(status(404))
    }
  })
}

/** Vite serves modules from paths of its own (`/@vite/client`, `/node_modules/.vite/…`). */
function routeOf(url: string): ViewerRoute | null {
  const route = viewerRoute(url)
  if (route !== null || devServerUrl() === null) return route
  const { pathname, search } = new URL(url)
  return { kind: 'page', name: pathname.replace(/^\/+/, '') + search }
}

async function respond(route: ViewerRoute, request: Request): Promise<Response | null> {
  switch (route.kind) {
    case 'document':
      return serveDocument(route.token, request.headers.get('range'))
    case 'pdfjs':
      return serveFile(pdfjsFile(route.name))
    case 'page': {
      const devServer = devServerUrl()
      // In development the page comes from Vite, still under this origin, so the frame stays isolated
      // (and without hot reload, whose socket the frame's policy refuses).
      if (devServer !== null) return net.fetch(new URL(route.name, devServer).href)
      return serveFile(containedPath(rendererRoot(), route.name))
    }
  }
}

async function serveFile(path: string | null): Promise<Response | null> {
  if (path === null) return null
  return new Response(await readFile(path), { headers: { 'Content-Type': contentTypeOf(path) } })
}

/** By ranges: a PDF of hundreds of megabytes opens without reaching memory whole. */
async function serveDocument(token: string, rangeHeader: string | null): Promise<Response | null> {
  const path = sharedDocuments.get(token)
  if (path === undefined) return null

  const { size } = await stat(path)
  const headers = {
    'Content-Type': 'application/pdf',
    'Accept-Ranges': 'bytes',
    'Cache-Control': 'no-store',
  }
  const range = parseByteRange(rangeHeader, size)
  if (range === 'unsatisfiable') {
    return new Response(null, { status: 416, headers: { ...headers, 'Content-Range': `bytes */${size}` } })
  }
  if (range === null) {
    return new Response(streamOf(path), { headers: { ...headers, 'Content-Length': String(size) } })
  }
  return new Response(streamOf(path, range), {
    status: 206,
    headers: {
      ...headers,
      'Content-Length': String(range.end - range.start + 1),
      'Content-Range': `bytes ${range.start}-${range.end}/${size}`,
    },
  })
}

function streamOf(path: string, range?: { start: number; end: number }): ReadableStream<Uint8Array> {
  return Readable.toWeb(createReadStream(path, range)) as ReadableStream<Uint8Array>
}

function status(code: number): Response {
  return new Response(null, { status: code })
}

/** Every response carries the frame's policy, including the page Vite serves in development. */
function withPolicy(response: Response): Response {
  const headers = new Headers(response.headers)
  headers.set('Content-Security-Policy', VIEWER_CONTENT_SECURITY_POLICY)
  headers.set('X-Content-Type-Options', 'nosniff')
  return new Response(response.body, { status: response.status, headers })
}

/** pdf.js's Liberation Sans is version 1, under the GPL; the app ships version 2, under the OFL. */
const LIBERATION_SANS = /^standard_fonts\/(LiberationSans-[A-Za-z]+\.ttf)$/

function pdfjsFile(name: string): string | null {
  const liberation = LIBERATION_SANS.exec(name)?.[1]
  if (liberation !== undefined) return containedPath(fontsRoot(), liberation)
  return containedPath(pdfjsRoot(), name)
}

/** `out/main/index.js` → `out/renderer`, in the asar when packaged. */
function rendererRoot(): string {
  return join(dirname(fileURLToPath(import.meta.url)), '..', 'renderer')
}

/** The build copies them to `out/renderer/pdfjs`; in development they come from the package. */
function pdfjsRoot(): string {
  if (devServerUrl() === null) return join(rendererRoot(), 'pdfjs')
  return join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'node_modules', 'pdfjs-dist')
}
