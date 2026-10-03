import { readFile } from 'node:fs/promises'
import { dirname, extname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { app, protocol } from 'electron'

/**
 * Um esquema próprio, e não `file:` nem `data:`: a janela que gera o PDF carrega
 * um HTML de pasta temporária com `webSecurity`, e `file:` seria requisição
 * entre origens; `data:` somaria 9,4 MB de base64 ao renderer e a cada impressão.
 */
export const FONT_SCHEME = 'librevia-font'

/** Precisa ser declarado **antes** de `app.whenReady()`. */
export function registerFontScheme(): void {
  protocol.registerSchemesAsPrivileged([
    {
      scheme: FONT_SCHEME,
      privileges: {
        // Sem origem de verdade o Chromium recusaria a fonte por CORS.
        standard: true,
        secure: true,
        supportFetchAPI: true,
        corsEnabled: true,
      },
    },
  ])
}

/** Depois de `app.whenReady()`. */
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

/** Conferido contra o caminho resolvido, e não contra o texto: o pedido vem de um documento qualquer. */
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

/** Como no sidecar: empacotado é `process.resourcesPath`; fora, deriva do próprio bundle. */
function fontsRoot(): string {
  const root = app.isPackaged
    ? process.resourcesPath
    : join(dirname(fileURLToPath(import.meta.url)), '..', '..')
  return join(root, 'resources', 'fonts')
}
