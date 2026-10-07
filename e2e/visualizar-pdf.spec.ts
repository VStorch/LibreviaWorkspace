import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'

let session: Session
let folder: string
let file: string

test.beforeEach(async () => {
  folder = await mkdtemp(join(tmpdir(), 'librevia-viewer-'))
  file = join(folder, 'relatorio.pdf')
  await writeFile(file, onePagePdf('Relatorio anual'))
  session = await launch()
})

test.afterEach(async () => {
  await session.close()
  await rm(folder, { recursive: true, force: true })
})

test('abre o PDF na mesma janela, como os outros arquivos', async () => {
  await stubDialogs(session.app, { open: file })
  await menu(session, 'open')

  await expect.poll(() => viewerRendered(session)).toBe(true)
  expect(await pdfView(session)).toMatchObject({ windows: 1, visible: true })
  await expect(session.window).toHaveTitle('relatorio.pdf — Librevia')
  await expect(session.window.locator('.statusbar')).toContainText(file)
})

test('o PDF ocupa a área reservada, abaixo da barra do aplicativo', async () => {
  await stubDialogs(session.app, { open: file })
  await menu(session, 'open')
  await expect.poll(() => viewerRendered(session)).toBe(true)

  const area = await session.window.locator('.pdf-view').boundingBox()
  await expect
    .poll(async () => (await pdfView(session)).bounds)
    .toEqual({
      x: Math.round(area!.x),
      y: Math.round(area!.y),
      width: Math.round(area!.width),
      height: Math.round(area!.height),
    })
})

test('abrir um PDF com alterações não salvas pergunta antes', async () => {
  await menu(session, 'new-document')
  await session.window.locator('.ProseMirror').fill('Trabalho não salvo')

  await stubDialogs(session.app, { open: file, messageBox: 2 })
  await menu(session, 'open')
  await expect(session.window.locator('.ProseMirror')).toHaveText('Trabalho não salvo')
  expect((await pdfView(session)).attached).toBe(false)

  await stubDialogs(session.app, { messageBox: 1 })
  await menu(session, 'open')
  await expect.poll(() => viewerRendered(session)).toBe(true)
})

test('um diálogo do aplicativo aparece por cima do PDF', async () => {
  await stubDialogs(session.app, { open: file })
  await menu(session, 'open')
  await expect.poll(() => viewerRendered(session)).toBe(true)

  await menu(session, 'new-from-template')
  await expect(session.window.locator('[aria-modal="true"]')).toBeVisible()
  await expect.poll(async () => (await pdfView(session)).visible).toBe(false)

  await session.window.getByRole('button', { name: 'Cancelar' }).click()
  await expect.poll(async () => (await pdfView(session)).visible).toBe(true)
})

test('salvar com um PDF aberto explica que ele é só para leitura', async () => {
  await stubDialogs(session.app, { open: file })
  await menu(session, 'open')
  await expect.poll(() => viewerRendered(session)).toBe(true)

  await menu(session, 'save')
  await expect(session.window.locator('.banner')).toContainText('só para leitura')
})

test('fechar o arquivo tira o PDF da tela', async () => {
  await stubDialogs(session.app, { open: file })
  await menu(session, 'open')
  await expect.poll(() => viewerRendered(session)).toBe(true)

  await menu(session, 'close-file')
  await expect(session.window.locator('.home')).toBeVisible()
  await expect.poll(async () => (await pdfView(session)).attached).toBe(false)
})

test('abre o PDF passado ao iniciar o aplicativo', async () => {
  await session.close()
  session = await launch({ file })

  await expect.poll(() => viewerRendered(session)).toBe(true)
  await expect(session.window).toHaveTitle('relatorio.pdf — Librevia')
})

test('a visualização de impressão mostra as páginas', async () => {
  await menu(session, 'new-document')
  await session.window.locator('.ProseMirror').fill('Ata')
  await menu(session, 'print-preview')

  await expect.poll(() => viewerRendered(session), { timeout: 30_000 }).toBe(true)
})

interface PdfViewState {
  readonly windows: number
  readonly attached: boolean
  readonly visible: boolean
  readonly bounds: { x: number; y: number; width: number; height: number } | null
}

/** Main's view, from main: the renderer only sees the empty area it reserved. */
async function pdfView(current: Session): Promise<PdfViewState> {
  return current.app.evaluate(({ BrowserWindow }) => {
    const windows = BrowserWindow.getAllWindows()
    const view = windows[0]?.contentView.children[0]
    return {
      windows: windows.length,
      attached: view !== undefined,
      visible: view?.getVisible() ?? false,
      bounds: view?.getBounds() ?? null,
    }
  })
}

/** The viewer is Chromium's own page: it only draws if its scripts ran, which the app's CSP blocked. */
async function viewerRendered(current: Session): Promise<boolean> {
  const frames = current.app.windows().flatMap((page) => page.frames())
  const viewer = frames.find((frame) => frame.url().startsWith('chrome-extension://'))
  if (viewer === undefined) return false
  return viewer
    .evaluate(() => Boolean(document.querySelector('pdf-viewer')?.shadowRoot?.querySelector('#toolbar')))
    .catch(() => false)
}

/** A real PDF, small enough to read: one page with one line of Helvetica. */
function onePagePdf(text: string): Buffer {
  const stream = `BT /F1 24 Tf 72 720 Td (${text}) Tj ET`
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>',
    `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`,
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
  ]
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
