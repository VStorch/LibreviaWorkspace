import { readFile } from 'node:fs/promises'
import { dirname, extname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { app, protocol } from 'electron'

/**
 * A scheme of our own, not `file:` or `data:`: the window that renders the PDF loads HTML from a
 * temp folder with `webSecurity`, so `file:` would be cross-origin; `data:` would add 9.4 MB of
 * base64 to the renderer and to every print.
 */
export const FONT_SCHEME = 'librevia-font'

/** Must be declared **before** `app.whenReady()`. */
export function registerFontScheme(): void {
  protocol.registerSchemesAsPrivileged([
    {
      scheme: FONT_SCHEME,
      privileges: {
        // Without a real origin Chromium would refuse the font through CORS.
        standard: true,
        secure: true,
        supportFetchAPI: true,
        corsEnabled: true,
      },
    },
  ])
}

/** After `app.whenReady()`. */
export function serveFonts(): void {
  protocol.handle(FONT_SCHEME, async (request) => {
    const file = fontFileFor(request.url)
    if (file === null) return new Response('', { status: 404 })

    try {
      return new Response(await readFile(file), {
        headers: {
          'Content-Type': 'font/ttf',
          'Cache-Control': 'public, max-age=31536000, immutable',
        },
      })
    } catch {
      return new Response('', { status: 404 })
    }
  })
}

/** Checked against the resolved path, not the text: the request comes from any document. */
export function fontFileFor(url: string, root: string = fontsRoot()): string | null {
  let name: string
  try {
    name = decodeURIComponent(new URL(url).pathname).replace(/^\/+/, '')
  } catch {
    return null
  }

  if (name === '' || extname(name).toLowerCase() !== '.ttf') return null

  const candidate = resolve(root, name)
  const inside = resolve(root)
  if (candidate !== join(inside, name) || !candidate.startsWith(inside)) return null

  return candidate
}

/**
 * As in the sidecar: packaged it is `process.resourcesPath`; otherwise it derives from the bundle.
 */
function fontsRoot(): string {
  const root = app.isPackaged
    ? process.resourcesPath
    : join(dirname(fileURLToPath(import.meta.url)), '..', '..')
  return join(root, 'resources', 'fonts')
}
