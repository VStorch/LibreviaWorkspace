import { readFile } from 'node:fs/promises'
import { dirname, extname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { app, protocol } from 'electron'

/**
 * As fontes empacotadas, servidas por um esquema próprio.
 *
 * O documento corporativo pede Calibri, Cambria, Arial e Times New Roman, e
 * nenhuma delas existe num Linux limpo. Sem elas o Chromium substitui por conta
 * própria, as métricas mudam e a quebra de linha cai noutro lugar — o documento
 * de três páginas vira quatro. Com paginação ao vivo isso deixou de ser um
 * detalhe do PDF: é o número que a pessoa lê na barra de status.
 *
 * As substitutas são **metricamente compatíveis**: cada glifo ocupa a mesma
 * largura da fonte original, então a linha quebra no mesmo ponto. É o mesmo
 * conjunto que o LibreOffice usa para abrir documentos do Word, e é por isso que
 * a comparação com ele passa a fazer sentido.
 *
 * ## Por que um esquema próprio, e não `data:` nem `file:`
 *
 * As fontes precisam valer nos **dois** lugares que desenham documento: o editor
 * e a janela oculta que gera o PDF. São contextos diferentes — a janela de
 * impressão carrega um HTML de pasta temporária, com `webSecurity` ligado, e
 * dali um `file:` para `resources/` é requisição entre origens, que o Chromium
 * recusa.
 *
 * Embutir como `data:` funcionaria nos dois e o CSP já permite, mas são 7,2 MB
 * de fonte: viraria 9,4 MB de base64 no pacote do renderer **e** em cada arquivo
 * temporário escrito a cada impressão.
 *
 * Um esquema registrado resolve os dois casos com um caminho só, e o custo é uma
 * entrada no `font-src` do CSP.
 */
export const FONT_SCHEME = 'librevia-font'

/** Precisa ser declarado **antes** de `app.whenReady()`. */
export function registerFontScheme(): void {
  protocol.registerSchemesAsPrivileged([
    {
      scheme: FONT_SCHEME,
      privileges: {
        // `standard` dá ao esquema uma origem de verdade, sem a qual o
        // Chromium o trata como opaco e recusa a fonte por CORS.
        standard: true,
        secure: true,
        supportFetchAPI: true,
        corsEnabled: true,
      },
    },
  ])
}

/** Liga o esquema à pasta empacotada. Depois de `app.whenReady()`. */
export function serveFonts(): void {
  protocol.handle(FONT_SCHEME, async (request) => {
    const file = fontFileFor(request.url)
    if (file === null) return new Response('', { status: 404 })

    try {
      return new Response(await readFile(file), {
        headers: {
          'Content-Type': 'font/ttf',
          // O conjunto é fixo e viaja dentro do instalador: reler a cada
          // parágrafo redesenhado seria trabalho puro.
          'Cache-Control': 'public, max-age=31536000, immutable',
        },
      })
    } catch {
      return new Response('', { status: 404 })
    }
  })
}

/**
 * URL do esquema → caminho em disco, ou `null` se não for uma fonte nossa.
 *
 * A pasta é fixa e só `.ttf` passa. O nome é conferido contra o caminho
 * resolvido, e não contra o texto pedido: `../` normalizado é a forma clássica
 * de sair de uma pasta que se acreditava fechada, e aqui o pedido vem de um
 * documento que pode ter sido escrito por qualquer um.
 */
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
 * Onde a pasta mora.
 *
 * Mesma regra do sidecar (`sidecar/index.ts`): empacotado é
 * `process.resourcesPath`, fora do pacote deriva da localização do próprio
 * bundle. `app.getAppPath()` muda conforme o Electron é chamado, e essa variação
 * já custou um diagnóstico errado uma vez.
 */
function fontsRoot(): string {
  const root = app.isPackaged
    ? process.resourcesPath
    : join(dirname(fileURLToPath(import.meta.url)), '..', '..')
  return join(root, 'resources', 'fonts')
}
