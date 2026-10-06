import { mkdir } from 'node:fs/promises'
import { basename, dirname, extname, join } from 'node:path'
import { deflateRawSync } from 'node:zlib'
import { Language } from '@shared/i18n/language.js'
import { IpcChannel } from '@shared/ipc-channels.js'
import { exportHtml } from '@services/document/export-html.js'
import { exportMarkdown } from '@services/document/export-markdown.js'
import { exportOdt } from '@services/document/export-odt.js'
import { parseDocument } from '@services/document/serialize.js'
import { fileNameFromPath } from '@services/file/formats.js'
import { showExportSaveDialog } from '../dialogs.js'
import { writeFileAtomic } from '../fs/atomic-write.js'
import { authorizePath } from '../fs/paths.js'
import { t } from '../i18n.js'
import { editorPreferences } from '../preferences.js'
import { handle } from './registry.js'
import { windowOf } from './sender-window.js'

/** Writes a **new** file: does not touch the open document's path or "modified" state. */

export function exportName(suggestedName: string, extension: string): string {
  const dot = suggestedName.lastIndexOf('.')
  return `${dot > 0 ? suggestedName.slice(0, dot) : suggestedName}.${extension}`
}

/** `relatorio.md` → `relatorio_arquivos`. */
export function assetFolderOf(path: string): string {
  return `${basename(path, extname(path))}_arquivos`
}

export function registerExportHandlers(): void {
  handle(IpcChannel.FileExport, async (payload, event) => {
    const html = payload.format === 'html'
    const extension = { html: 'html', markdown: 'md', odt: 'odt' }[payload.format]
    const chosen = await showExportSaveDialog(
      windowOf(event),
      exportName(payload.suggestedName, extension),
      payload.format,
    )
    if (chosen === null) return { canceled: true as const }

    const language = editorPreferences().language
    const model = parseDocument(payload.content, language)
    const path = authorizePath(chosen)

    if (payload.format === 'odt') {
      const bytes = exportOdt(model, {}, (data) => deflateRawSync(data))
      await writeFileAtomic(path, Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength))
    } else if (html) {
      const text = exportHtml(model, {
        fileName: payload.suggestedName,
        lang: language === Language.Portuguese ? 'pt-BR' : 'en',
        labels: { notes: t('dialog.export.notes'), backToText: t('dialog.export.backToText') },
      })
      await writeFileAtomic(path, text)
    } else {
      const folder = assetFolderOf(path)
      const { markdown, assets } = exportMarkdown(model, { assetFolder: folder })
      // The folder only appears when there are images.
      if (assets.length > 0) {
        const directory = join(dirname(path), folder)
        await mkdir(directory, { recursive: true })
        for (const asset of assets) {
          await writeFileAtomic(join(directory, asset.name), Buffer.from(asset.base64, 'base64'))
        }
      }
      await writeFileAtomic(path, markdown)
    }

    return { canceled: false as const, path, name: fileNameFromPath(path) }
  })
}
