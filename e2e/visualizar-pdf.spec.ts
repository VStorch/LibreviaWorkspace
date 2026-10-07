import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test, type Frame } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'

/**
 * One Letter page with a 200 × 200 JPEG 2000 image, red on the left and blue on the right, at
 * x 156–456 and y 300–600 (in points). Generated with Pillow; Node has no JPEG 2000 encoder.
 */
const JPEG2000_PDF = [
  'JVBERi0xLjUKMSAwIG9iago8PCAvVHlwZSAvQ2F0YWxvZyAvUGFnZXMgMiAwIFIgPj4KZW5kb2JqCjIgMCBvYmoKPDwgL1R5',
  'cGUgL1BhZ2VzIC9LaWRzIFszIDAgUl0gL0NvdW50IDEgPj4KZW5kb2JqCjMgMCBvYmoKPDwgL1R5cGUgL1BhZ2UgL1BhcmVu',
  'dCAyIDAgUiAvTWVkaWFCb3ggWzAgMCA2MTIgNzkyXSAvQ29udGVudHMgNCAwIFIgL1Jlc291cmNlcyA8PCAvWE9iamVjdCA8',
  'PCAvSW0xIDUgMCBSID4+ID4+ID4+CmVuZG9iago0IDAgb2JqCjw8IC9MZW5ndGggMzQgPj4Kc3RyZWFtCnEgMzAwIDAgMCAz',
  'MDAgMTU2IDMwMCBjbSAvSW0xIERvIFEKZW5kc3RyZWFtCmVuZG9iago1IDAgb2JqCjw8IC9UeXBlIC9YT2JqZWN0IC9TdWJ0',
  'eXBlIC9JbWFnZSAvV2lkdGggMjAwIC9IZWlnaHQgMjAwIC9GaWx0ZXIgL0pQWERlY29kZSAvTGVuZ3RoIDYzNCA+PgpzdHJl',
  'YW0KAAAADGpQICANCocKAAAAFGZ0eXBqcDIgAAAAAGpwMiAAAAAtanAyaAAAABZpaGRyAAAAyAAAAMgAAwcHAAAAAAAPY29s',
  'cgEAAAAAABAAAAItanAyY/9P/1EALwAAAAAAyAAAAMgAAAAAAAAAAAAAAMgAAADIAAAAAAAAAAAAAwcBAQcBAQcBAf9SAAwA',
  'AAABAAUEBAAB/1wAE0BASEhQSEhQSEhQSEhQSEhQ/2QAJQABQ3JlYXRlZCBieSBPcGVuSlBFRyB2ZXJzaW9uIDIuNS4w/5AA',
  'CgAAAAABpgAB/5PPtJwUACrQ8rYiGOAFpT2oPMAbz/gfBQN2RSfwUDdkT4p1gJwvMBVlBH/PtEARUFSvF1twAAAAYSC2nGEg',
  'z7SkEVAi0PK1g4gbn1HUqeAGHDMKe0oeUcWErZKQeCQ4JtWq7tn7/4B/BwfH2hoAIW1dCTA6J019qB4Ch4DH2hwAIhlfCTA6',
  'J019pyftMxfH2h4AMNHFO2irIKZsC/7HFwpngMfaLAAxfcc7VID1OQdk7pQSBmv3X2HgL3W/w+oeAE+XxT/8aj64gIiL1VYw',
  'bgOLW6+b5uZZ2CO8NIjpf4DD6hoAUEPHP/xqPriAiIvVbLlyETp9e2SYYtbr5vrH2mgAg+2vZn/9zjQ3BHwAAGEhEn/c8Ij5',
  'DJhi1uvm+b5vm+b5vm+Z6Zfm+b5vm+b5vm+b5vm+Z4DH2mgAg8Ku5n/9zjQ3BHwAAGEhEn/c8Ij5DJhi1uvm+b5vm+b5vm+Z',
  '6Zfm+b5vm+b5vm+b5vm+Z+P2hD9oOJ6KMc//f3XPnooxz/9Zr4Dj9oQ/aDieXzFP/391z55fMU//Wa//2QplbmRzdHJlYW0K',
  'ZW5kb2JqCnhyZWYKMCA2CjAwMDAwMDAwMDAgNjU1MzUgZiAKMDAwMDAwMDAwOSAwMDAwMCBuIAowMDAwMDAwMDU4IDAwMDAw',
  'IG4gCjAwMDAwMDAxMTUgMDAwMDAgbiAKMDAwMDAwMDI0NSAwMDAwMCBuIAowMDAwMDAwMzI5IDAwMDAwIG4gCnRyYWlsZXIK',
  'PDwgL1NpemUgNiAvUm9vdCAxIDAgUiA+PgpzdGFydHhyZWYKMTA4NwolJUVPRgo=',
].join('')

let session: Session
let folder: string
let file: string

test.beforeEach(async () => {
  folder = await mkdtemp(join(tmpdir(), 'librevia-viewer-'))
  file = join(folder, 'relatorio.pdf')
  await writeFile(file, pdfOf(['Relatorio anual']))
  session = await launch()
})

test.afterEach(async () => {
  await session.close()
  await rm(folder, { recursive: true, force: true })
})

test('abre o PDF na mesma janela, desenhado pelo visualizador do aplicativo', async () => {
  await openPdf(file)

  await expect(pageText()).toContainText('Relatorio anual')
  expect(session.app.windows()).toHaveLength(1)
  await expect(session.window).toHaveTitle('relatorio.pdf — Librevia')
  await expect(session.window.locator('.statusbar')).toContainText(file)
  await expect(session.window.getByTestId('pdf-page')).toHaveText('Página 1 de 1')
})

test('o visualizador roda isolado: sem acesso ao aplicativo, sem janelas e em outro processo', async () => {
  await openPdf(file)
  const frame = await viewerFrame()

  expect(await frame.evaluate(() => 'api' in window)).toBe(false)
  expect(await frame.evaluate(() => attempt(() => (parent as unknown as { api: unknown }).api))).toBe(
    'SecurityError',
  )
  expect(await frame.evaluate(() => attempt(() => parent.document.title))).toBe('SecurityError')
  expect(await frame.evaluate(() => window.open('https://example.com') === null)).toBe(true)
  expect(await frame.evaluate(() => attempt(() => (top!.location.href = 'https://example.com')))).toBe(
    'SecurityError',
  )
  expect(session.app.windows()).toHaveLength(1)

  const processes = await session.app.evaluate(({ BrowserWindow }) => {
    const main = BrowserWindow.getAllWindows()[0]!.webContents.mainFrame
    return [main.osProcessId, main.frames[0]?.osProcessId]
  })
  expect(processes[1]).toBeDefined()
  expect(processes[1]).not.toBe(processes[0])
})

test('abrir um PDF com alterações não salvas pergunta antes', async () => {
  await menu(session, 'new-document')
  await session.window.locator('.ProseMirror').fill('Trabalho não salvo')

  await stubDialogs(session.app, { open: file, messageBox: 2 })
  await menu(session, 'open')
  await expect(session.window.locator('.ProseMirror')).toHaveText('Trabalho não salvo')
  await expect(session.window.locator('.pdf-view__frame')).toHaveCount(0)

  await stubDialogs(session.app, { messageBox: 1 })
  await menu(session, 'open')
  await expect(pageText()).toContainText('Relatorio anual')
})

test('um diálogo do aplicativo aparece por cima do PDF', async () => {
  await openPdf(file)

  await menu(session, 'new-from-template')
  const dialog = session.window.locator('[aria-modal="true"]')
  await expect(dialog).toBeVisible()
  const box = (await dialog.boundingBox())!
  const onTop = await session.window.evaluate(
    ({ x, y }) => document.elementFromPoint(x, y)?.closest('[aria-modal="true"]') !== null,
    { x: box.x + box.width / 2, y: box.y + box.height / 2 },
  )
  expect(onTop).toBe(true)
})

test('salvar com um PDF aberto explica que ele é só para leitura', async () => {
  await openPdf(file)

  await menu(session, 'save')
  await expect(session.window.locator('.banner')).toContainText('só para leitura')
})

test('fechar o arquivo tira o PDF da tela', async () => {
  await openPdf(file)

  await menu(session, 'close-file')
  await expect(session.window.locator('.home')).toBeVisible()
  await expect(session.window.locator('.pdf-view__frame')).toHaveCount(0)
})

test('abre o PDF passado ao iniciar o aplicativo', async () => {
  await session.close()
  session = await launch({ file })

  await expect(pageText()).toContainText('Relatorio anual')
  await expect(session.window).toHaveTitle('relatorio.pdf — Librevia')
})

test('um PDF danificado mostra uma explicação no lugar das páginas', async () => {
  const damaged = join(folder, 'danificado.pdf')
  await writeFile(damaged, '%PDF-1.4\nisto não é um PDF de verdade\n')
  await openPdf(damaged)

  await expect(session.window.locator('.pdf-view--failed')).toContainText('danificado')
})

test('uma imagem JPEG 2000 aparece: o decodificador em WebAssembly roda sob a política do visualizador', async () => {
  const image = join(folder, 'imagem.pdf')
  await writeFile(image, Buffer.from(JPEG2000_PDF, 'base64'))
  await openPdf(image)
  const frame = await viewerFrame()

  const pixelAt = (x: number, y: number) =>
    frame.evaluate(
      ([pointX, pointY]) => {
        const canvas = document.querySelector<HTMLCanvasElement>('.page canvas')!
        const at = [(pointX! / 612) * canvas.width, ((792 - pointY!) / 792) * canvas.height] as const
        return Array.from(canvas.getContext('2d')!.getImageData(at[0], at[1], 1, 1).data.slice(0, 3))
      },
      [x, y],
    )
  await expect.poll(() => pixelAt(230, 450)).toEqual([220, 30, 30])
  expect(await pixelAt(400, 450)).toEqual([30, 30, 220])
})

test('um PDF grande abre lendo só os trechos de que a página precisa', async () => {
  const large = join(folder, 'grande.pdf')
  const size = 5 * 1024 * 1024
  await writeFile(large, pdfOf(['Arquivo grande'], undefined, size))
  await session.app.evaluate(({ session: electronSession }) => {
    const ranges: string[] = []
    Object.assign(globalThis, { documentRanges: ranges })
    electronSession.defaultSession.webRequest.onSendHeaders({ urls: ['<all_urls>'] }, (details) => {
      if (details.url.includes('/document/'))
        ranges.push(String(details.requestHeaders['Range'] ?? 'inteiro'))
    })
  })

  await openPdf(large)
  await expect(pageText()).toContainText('Arquivo grande')

  const ranges = await session.app.evaluate(
    () => (globalThis as { documentRanges?: string[] }).documentRanges ?? [],
  )
  const read = ranges.reduce((total, range) => {
    const [start, end] = range.replace('bytes=', '').split('-').map(Number)
    return total + end! - start! + 1
  }, 0)
  expect(ranges).not.toContain('inteiro')
  expect(read).toBeLessThan(size / 10)
})

test('a barra de status acompanha a página e o zoom do PDF', async () => {
  const report = join(folder, 'tres-paginas.pdf')
  await writeFile(report, pdfOf(['Primeira', 'Segunda', 'Terceira']))
  await openPdf(report)
  const status = session.window.getByTestId('pdf-page')
  await expect(status).toHaveText('Página 1 de 3')

  const frame = await viewerFrame()
  await frame.evaluate(() => {
    const container = document.getElementById('viewerContainer')!
    container.scrollTop = container.scrollHeight
  })
  await expect(status).toHaveText('Página 3 de 3')

  await menu(session, 'zoom-reset')
  await expect(session.window.locator('.statusbar__zoom-level')).toHaveText('100%')
  // 8.5 inches at 96 pixels per inch: 100% is the paper's real size, as in the editor.
  await expect.poll(() => pageWidth(frame)).toBe(816)

  await menu(session, 'zoom-in')
  await expect(session.window.locator('.statusbar__zoom-level')).toHaveText('110%')
  // pdf.js rounds the page to whole pixels.
  await expect.poll(async () => Math.abs((await pageWidth(frame)) - 816 * 1.1)).toBeLessThanOrEqual(1)
})

test('imprimir um PDF abre o arquivo original na pré-visualização do Chromium', async () => {
  await openPdf(file)

  await menu(session, 'print')

  await expect.poll(() => session.app.windows().length).toBe(2)
  await expect.poll(() => chromiumViewerRendered(session)).toBe(true)
})

test('um link do PDF passa pela confirmação antes de abrir o navegador', async () => {
  const linked = join(folder, 'com-link.pdf')
  await writeFile(linked, pdfOf(['Veja o site'], 'https://example.com/ajuda'))
  await session.app.evaluate(({ dialog, shell }) => {
    const opened: string[] = []
    Object.assign(globalThis, { openedLinks: opened })
    shell.openExternal = async (url) => {
      opened.push(url)
    }
    dialog.showMessageBox = async () => ({ response: 0, checkboxChecked: false })
  })
  await openPdf(linked)

  await session.window.frameLocator('.pdf-view__frame').locator('.annotationLayer a').click()

  await expect
    .poll(() => session.app.evaluate(() => (globalThis as { openedLinks?: string[] }).openedLinks))
    .toEqual(['https://example.com/ajuda'])
})

test('a visualização de impressão de um documento mostra as páginas', async () => {
  await menu(session, 'new-document')
  await session.window.locator('.ProseMirror').fill('Ata')
  await menu(session, 'print-preview')

  await expect.poll(() => chromiumViewerRendered(session), { timeout: 30_000 }).toBe(true)
})

async function openPdf(path: string): Promise<void> {
  await stubDialogs(session.app, { open: path })
  await menu(session, 'open')
}

function pageText() {
  return session.window.frameLocator('.pdf-view__frame').locator('.textLayer').first()
}

async function viewerFrame(): Promise<Frame> {
  await expect(pageText()).toBeVisible()
  const frame = session.window.frames().find((candidate) => candidate.url().startsWith('librevia-viewer:'))
  if (frame === undefined) throw new Error('o visualizador não está na página')
  await frame.evaluate(() => {
    Object.assign(window, {
      attempt: (action: () => unknown): string => {
        try {
          action()
          return 'allowed'
        } catch (error) {
          return error instanceof Error ? error.name : String(error)
        }
      },
    })
  })
  return frame
}

function pageWidth(frame: Frame): Promise<number> {
  return frame.evaluate(() => document.querySelector('.page')?.clientWidth ?? 0)
}

declare function attempt(action: () => unknown): string

/** Chromium's own viewer only draws if its scripts ran, which the app's CSP would block. */
async function chromiumViewerRendered(current: Session): Promise<boolean> {
  const frames = current.app.windows().flatMap((page) => page.frames())
  const viewer = frames.find((frame) => frame.url().startsWith('chrome-extension://'))
  if (viewer === undefined) return false
  return viewer
    .evaluate(() => Boolean(document.querySelector('pdf-viewer')?.shadowRoot?.querySelector('#toolbar')))
    .catch(() => false)
}

/**
 * A real PDF, small enough to read: one Letter page per text, one line of Helvetica each. With
 * `link`, the first line is a link to that address; `paddingBytes` adds an unused stream.
 */
function pdfOf(texts: readonly string[], link?: string, paddingBytes = 0): Buffer {
  const pageCount = texts.length
  const firstPage = 4
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    `<< /Type /Pages /Kids [${texts.map((_, index) => `${firstPage + index * 2} 0 R`).join(' ')}] /Count ${pageCount} >>`,
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
  ]
  const linkObject = firstPage + pageCount * 2
  texts.forEach((text, index) => {
    const stream = `BT /F1 24 Tf 72 720 Td (${text}) Tj ET`
    const annotations = link !== undefined && index === 0 ? ` /Annots [${linkObject} 0 R]` : ''
    objects.push(
      `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents ${firstPage + index * 2 + 1} 0 R /Resources << /Font << /F1 3 0 R >> >>${annotations} >>`,
      `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`,
    )
  })
  if (paddingBytes > 0) {
    objects.push(`<< /Length ${paddingBytes} >>\nstream\n${' '.repeat(paddingBytes)}\nendstream`)
  }
  if (link !== undefined) {
    objects.push(
      `<< /Type /Annot /Subtype /Link /Rect [72 710 300 750] /Border [0 0 0] /A << /S /URI /URI (${link}) >> >>`,
    )
  }

  let body = '%PDF-1.4\n'
  const offsets = objects.map((object, index) => {
    const offset = body.length
    body += `${index + 1} 0 obj\n${object}\nendobj\n`
    return offset
  })
  const xref = body.length
  body += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`
  body += offsets.map((offset) => `${String(offset).padStart(10, '0')} 00000 n \n`).join('')
  body += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`
  return Buffer.from(body, 'latin1')
}
