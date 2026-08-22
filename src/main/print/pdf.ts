import { randomUUID } from 'node:crypto'
import { unlink, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { BrowserWindow, app, type WebContents } from 'electron'
import { AppError, ErrorCode } from '@shared/errors.js'
import type { PageSetup } from '@services/document/model.js'
import { buildNativePrintOptions, buildPrintOptions } from '@services/pdf/page-setup.js'

/**
 * Renderização para impressão.
 *
 * O documento é carregado numa janela oculta e o próprio Chromium produz o
 * PDF. É o mesmo motor que desenha o editor, então o resultado sai igual ao
 * que estava na tela — e sem nenhuma biblioteca de PDF envolvida.
 *
 * A janela roda com **JavaScript desligado**: o HTML vem do documento do
 * usuário, que pode ter vindo de qualquer lugar, e para virar PDF não é
 * preciso executar nada.
 */
async function withRenderWindow<T>(html: string, run: (contents: WebContents) => Promise<T>): Promise<T> {
  const temporaryPath = join(app.getPath('temp'), `librevia-print-${randomUUID()}.html`)
  await writeFile(temporaryPath, html, 'utf8')

  const window = new BrowserWindow({
    show: false,
    webPreferences: {
      javascript: false,
      sandbox: true,
      contextIsolation: true,
      nodeIntegration: false,
      webSecurity: true,
      webviewTag: false,
    },
  })

  try {
    const loaded = new Promise<void>((resolve, reject) => {
      window.webContents.once('did-finish-load', () => resolve())
      window.webContents.once('did-fail-load', (_event, _code, description) =>
        reject(new AppError(ErrorCode.Internal, `Não foi possível preparar o documento: ${description}`)),
      )
    })

    await window.loadFile(temporaryPath)
    await loaded

    return await run(window.webContents)
  } finally {
    if (!window.isDestroyed()) window.destroy()
    await unlink(temporaryPath).catch(() => undefined)
  }
}

export async function renderPdf(html: string, page: PageSetup, paged = false): Promise<Buffer> {
  return withRenderWindow(html, async (contents) => contents.printToPDF(buildPrintOptions(page, paged)))
}

/**
 * Abre o diálogo de impressão do sistema.
 *
 * Resolve com `false` quando o usuário cancela — cancelar não é erro.
 */
export async function printDocument(html: string, page: PageSetup): Promise<boolean> {
  return withRenderWindow(
    html,
    (contents) =>
      new Promise<boolean>((resolve, reject) => {
        contents.print({ ...buildNativePrintOptions(page), silent: false }, (success, reason) => {
          if (success) {
            resolve(true)
            return
          }
          // O Chromium usa a mesma via para "cancelado" e para falha real.
          if (reason === 'cancelled' || reason === 'canceled') {
            resolve(false)
            return
          }
          reject(new AppError(ErrorCode.Internal, `Não foi possível imprimir: ${reason}`))
        })
      }),
  )
}

/**
 * Visualização de impressão.
 *
 * Mostra o **PDF já gerado**, não uma segunda renderização em HTML: é a única
 * forma de a prévia responder à pergunta que importa — onde as páginas
 * quebram de verdade. Compensa o editor não paginar ao vivo (§6.3 do plano).
 */
export async function openPdfPreview(parent: BrowserWindow, pdf: Buffer, title: string): Promise<void> {
  const temporaryPath = join(app.getPath('temp'), `librevia-preview-${randomUUID()}.pdf`)
  await writeFile(temporaryPath, pdf)

  const window = new BrowserWindow({
    parent,
    width: 900,
    height: 1000,
    title: `Visualizar impressão — ${title}`,
    autoHideMenuBar: true,
    webPreferences: {
      // O visualizador de PDF do Chromium precisa de plugins habilitados.
      plugins: true,
      javascript: false,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  })

  window.on('closed', () => {
    void unlink(temporaryPath).catch(() => undefined)
  })

  await window.loadFile(temporaryPath)
}
