import { BrowserWindow, type IpcMainInvokeEvent } from 'electron'
import { AppError, ErrorCode } from '@shared/errors.js'
import { IpcChannel } from '@shared/ipc-channels.js'
import { fileNameFromPath } from '@services/file/formats.js'
import { showPdfSaveDialog } from '../dialogs.js'
import { writeFileAtomic } from '../fs/atomic-write.js'
import { authorizePath } from '../fs/paths.js'
import { openPdfPreview, printDocument, renderPdf } from '../print/pdf.js'
import { handle } from './registry.js'

function windowOf(event: IpcMainInvokeEvent): BrowserWindow {
  const window = BrowserWindow.fromWebContents(event.sender)
  if (window === null) {
    throw new AppError(ErrorCode.Internal, 'A janela do aplicativo não está disponível.')
  }
  return window
}

/** Troca a extensão do documento pela do PDF, preservando o nome. */
function toPdfName(suggestedName: string): string {
  const dot = suggestedName.lastIndexOf('.')
  return `${dot > 0 ? suggestedName.slice(0, dot) : suggestedName}.pdf`
}

export function registerPrintHandlers(): void {
  handle(IpcChannel.PrintExportPdf, async (payload, event) => {
    // O destino é escolhido antes de gerar: se o usuário desistir, não gastamos
    // tempo renderizando, e nada é escrito.
    const chosen = await showPdfSaveDialog(windowOf(event), toPdfName(payload.suggestedName))
    if (chosen === null) return { canceled: true as const }

    const pdf = await renderPdf(payload.html, payload.page, payload.paged)
    const path = authorizePath(chosen)
    await writeFileAtomic(path, pdf)

    return { canceled: false as const, path, name: fileNameFromPath(path) }
  })

  handle(IpcChannel.PrintDialog, async (payload) => ({
    printed: await printDocument(payload.html, payload.page),
  }))

  handle(IpcChannel.PrintPreview, async (payload, event) => {
    const pdf = await renderPdf(payload.html, payload.page, payload.paged)
    await openPdfPreview(windowOf(event), pdf, payload.title)
    return { opened: true as const }
  })
}
