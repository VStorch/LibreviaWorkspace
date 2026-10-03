import { randomUUID } from 'node:crypto'
import { unlink, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { BrowserWindow, app, type WebContents } from 'electron'
import { AppError, ErrorCode } from '@shared/errors.js'
import type { PageSetup } from '@services/document/model.js'
import { buildNativePrintOptions, buildPrintOptions } from '@services/pdf/page-setup.js'
import { t } from '../i18n.js'

/**
 * O próprio Chromium gera o PDF numa janela oculta, com **JavaScript
 * desligado**: o HTML vem do documento, e virar PDF não exige executar nada.
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
        reject(new AppError(ErrorCode.Internal, t('errors.print.prepareFailed', { description }))),
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

/** `false` quando a pessoa cancela: cancelar não é erro. */
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
          reject(new AppError(ErrorCode.Internal, t('errors.print.printFailed', { reason })))
        })
      }),
  )
}

/** Mostra o **PDF já gerado**, e não outra renderização: é onde as páginas quebram de verdade. */
export async function openPdfPreview(parent: BrowserWindow, pdf: Buffer, title: string): Promise<void> {
  const temporaryPath = join(app.getPath('temp'), `librevia-preview-${randomUUID()}.pdf`)
  await writeFile(temporaryPath, pdf)

  const window = new BrowserWindow({
    parent,
    width: 900,
    height: 1000,
    title: t('errors.print.previewTitle', { title }),
    autoHideMenuBar: true,
    webPreferences: {
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
