import { IpcChannel } from '@shared/ipc-channels.js'
import { fileNameFromPath } from '@services/file/formats.js'
import { showPdfSaveDialog } from '../dialogs.js'
import { writeFileAtomic } from '../fs/atomic-write.js'
import { authorizePath } from '../fs/paths.js'
import { withPdfProperties } from '../pdf-metadata.js'
import { sidecar } from '../sidecar/index.js'
import { openPdfPreview, printDocument, renderPdf } from '../print/pdf.js'
import { handle } from './registry.js'
import { windowOf } from './sender-window.js'

/** What Chromium does not write: it takes only the title, from `<title>`. */
const COMPLETED_FIELDS = ['creator', 'subject', 'keywords'] as const

type CompletedFields = Partial<Record<(typeof COMPLETED_FIELDS)[number], string | undefined>>

/** Only when there is something to add: without it, the export does not need the sidecar. */
async function completedMetadata(
  pdf: Uint8Array,
  properties: CompletedFields | undefined,
): Promise<Uint8Array> {
  const missing: Record<string, string> = {}
  for (const key of COMPLETED_FIELDS) {
    const text = properties?.[key]?.trim() ?? ''
    if (text !== '') missing[key] = text
  }
  if (Object.keys(missing).length === 0) return pdf
  return withPdfProperties(sidecar(), pdf, missing)
}

function toPdfName(suggestedName: string): string {
  const dot = suggestedName.lastIndexOf('.')
  return `${dot > 0 ? suggestedName.slice(0, dot) : suggestedName}.pdf`
}

export function registerPrintHandlers(): void {
  handle(IpcChannel.PrintExportPdf, async (payload, event) => {
    // The destination before rendering: if the user gives up, nothing is rendered or written.
    const chosen = await showPdfSaveDialog(windowOf(event), toPdfName(payload.suggestedName))
    if (chosen === null) return { canceled: true as const }

    const pdf = await renderPdf(payload.html, payload.page, payload.paged)
    const path = authorizePath(chosen)
    await writeFileAtomic(path, await completedMetadata(pdf, payload.properties))

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
