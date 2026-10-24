import { randomUUID } from 'node:crypto'
import { unlink, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { BrowserWindow, app, type WebContents } from 'electron'
import { AppError, ErrorCode } from '@shared/errors.js'
import type { PageSetup } from '@services/document/model.js'
import { buildNativePrintOptions, buildPrintOptions } from '@services/pdf/page-setup.js'
import { t } from '../i18n.js'
import { openPdfWindow } from '../pdf-viewer.js'

/**
 * Chromium itself renders the PDF in a hidden window with **JavaScript off**: the HTML comes from
 * the document, and turning it into a PDF needs nothing executed. With `over`, the window shows above
 * that one before `run`.
 */
async function withRenderWindow<T>(
  html: string,
  run: (contents: WebContents) => Promise<T>,
  over?: BrowserWindow,
): Promise<T> {
  const temporaryPath = join(app.getPath('temp'), `librevia-print-${randomUUID()}.html`)
  await writeFile(temporaryPath, html, 'utf8')

  const window = new BrowserWindow({
    show: false,
    ...(over === undefined ? {} : { parent: over, modal: true }),
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
    if (over !== undefined) window.show()

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
 * On Wayland, the print dialog of a window never shown takes the app down: Chromium hands the
 * compositor a surface without a role. On Linux the page to print shows above the editor meanwhile.
 */
const PRINTS_FROM_SHOWN_WINDOW = process.platform === 'linux'

/** `false` when canceled: canceling is not an error. */
export async function printDocument(html: string, page: PageSetup, owner: BrowserWindow): Promise<boolean> {
  return withRenderWindow(
    html,
    (contents) =>
      new Promise<boolean>((resolve, reject) => {
        contents.print({ ...buildNativePrintOptions(page), silent: false }, (success, reason) => {
          if (success) {
            resolve(true)
            return
          }
          // Chromium uses the same path for "canceled" and for a real failure.
          if (reason === 'cancelled' || reason === 'canceled') {
            resolve(false)
            return
          }
          reject(new AppError(ErrorCode.Internal, t('errors.print.printFailed', { reason })))
        })
      }),
    PRINTS_FROM_SHOWN_WINDOW ? owner : undefined,
  )
}

/** Shows the **already rendered PDF**, not another rendering: that is where pages really break. */
export async function openPdfPreview(parent: BrowserWindow, pdf: Buffer, title: string): Promise<void> {
  const temporaryPath = join(app.getPath('temp'), `librevia-preview-${randomUUID()}.pdf`)
  await writeFile(temporaryPath, pdf)

  const window = openPdfWindow({
    path: temporaryPath,
    title: t('errors.print.previewTitle', { title }),
    parent,
  })
  window.on('closed', () => {
    void unlink(temporaryPath).catch(() => undefined)
  })
}
