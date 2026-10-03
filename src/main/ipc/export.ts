import { mkdir } from 'node:fs/promises'
import { basename, dirname, extname, join } from 'node:path'
import { deflateRawSync } from 'node:zlib'
import { BrowserWindow, type IpcMainInvokeEvent } from 'electron'
import { AppError, ErrorCode } from '@shared/errors.js'
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

/**
 * Exportação para HTML, Markdown e ODT.
 *
 * O renderer manda o documento serializado — o mesmo texto do salvar — e o main
 * monta o arquivo com as funções puras de `@services/document`. Escrever é um
 * arquivo **novo**: nada aqui toca o caminho do documento aberto nem o estado de
 * "alterado", que é o que distingue exportar de salvar como.
 */

function windowOf(event: IpcMainInvokeEvent): BrowserWindow {
  const window = BrowserWindow.fromWebContents(event.sender)
  if (window === null) {
    throw new AppError(ErrorCode.Internal, t('errors.ipc.windowNotAvailable'))
  }
  return window
}

/** Troca a extensão do documento pela do formato, preservando o nome. */
export function exportName(suggestedName: string, extension: string): string {
  const dot = suggestedName.lastIndexOf('.')
  return `${dot > 0 ? suggestedName.slice(0, dot) : suggestedName}.${extension}`
}

/** A pasta das imagens do Markdown, ao lado dele: `relatorio.md` → `relatorio_arquivos`. */
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
      // O pacote sai pronto da função pura; a compressão é a do Node.
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
      // A pasta só nasce quando há imagem: um texto sem figuras não deixa
      // diretório vazio ao lado.
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
