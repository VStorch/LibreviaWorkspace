import { AppError, ErrorCode } from '@shared/errors.js'
import { IpcChannel } from '@shared/ipc-channels.js'
import type { LoadedFile } from '@shared/types.js'
import {
  ensureSupportedExtension,
  fileNameFromPath,
  isExcelPath,
  isWordPackagePath,
  isWordTemplatePath,
  kindFromPath,
  WORD_EXTENSION,
} from '@services/file/formats.js'
import { showOpenFileDialog, showSaveFileDialog } from '../dialogs.js'
import { followDocxOriginal, forgetOpenedDocx, openDocx, saveDocx } from '../docx/index.js'
import { forgetOpenedXlsx, openXlsx, saveXlsx } from '../xlsx/index.js'
import { sidecar } from '../sidecar/index.js'
import { writeFileAtomic } from '../fs/atomic-write.js'
import { assertPathAuthorized, assertReadableFile, authorizePath, normalizePath } from '../fs/paths.js'
import { clearRecentFiles, isRemembered, listRecentFiles, rememberRecentFile } from '../fs/recent.js'
import { readTextFile } from '../fs/read-text.js'
import { refreshMenu } from '../menu.js'
import { t } from '../i18n.js'
import { handle } from './registry.js'
import { windowOf } from './sender-window.js'
import { forgetExternalFileRequest, isExternalFileRequested } from '../external-files.js'

/**
 * A Word template opens as a new document, and its path is **not** authorized: nothing writes over
 * the template. `templateName` is the new document's name: the builtin template's, translated; when
 * absent, the file's.
 */
export async function loadFile(path: string, templateName?: string): Promise<LoadedFile> {
  await assertReadableFile(path)

  // DOCX and XLSX become the internal format here: the renderer never sees OOXML.
  const loaded = isWordPackagePath(path)
    ? await openDocx(sidecar(), path)
    : isExcelPath(path)
      ? await openXlsx(sidecar(), path)
      : { content: await readTextFile(path), inventory: undefined }

  // The original bytes apply to **one** file: otherwise the next save would write into the wrong
  // package.
  if (!isWordPackagePath(path)) forgetOpenedDocx()
  if (!isExcelPath(path)) forgetOpenedXlsx()

  if (isWordTemplatePath(path)) {
    const source = normalizePath(path)
    rememberRecentFile(source)
    void refreshMenu()
    const base = templateName ?? fileNameFromPath(source).replace(/\.dot[xm]$/i, '')
    const file: LoadedFile = {
      path: source,
      name: `${base}${WORD_EXTENSION}`,
      kind: kindFromPath(source),
      content: loaded.content,
      template: true,
    }
    return loaded.inventory === undefined ? file : { ...file, inventory: loaded.inventory }
  }

  const authorized = authorizePath(path)
  rememberRecentFile(authorized)
  void refreshMenu()

  const file: LoadedFile = {
    path: authorized,
    name: fileNameFromPath(authorized),
    kind: kindFromPath(authorized),
    content: loaded.content,
  }

  return loaded.inventory === undefined ? file : { ...file, inventory: loaded.inventory }
}

export function registerFileHandlers(): void {
  handle(IpcChannel.FileOpen, async (_payload, event) => {
    const path = await showOpenFileDialog(windowOf(event))
    if (path === null) return { canceled: true as const }
    return { canceled: false as const, file: await loadFile(path) }
  })

  handle(IpcChannel.FileOpenRecent, async (payload) => {
    // The renderer does not choose paths: only recent files or file manager requests received by
    // main.
    if (!isRemembered(payload.path) && !isExternalFileRequested(payload.path)) {
      throw new AppError(ErrorCode.PathNotAuthorized, t('errors.ipc.notInRecents'))
    }
    try {
      return { file: await loadFile(payload.path) }
    } finally {
      forgetExternalFileRequest(payload.path)
    }
  })

  handle(IpcChannel.FileSave, async (payload) => {
    const path = assertPathAuthorized(payload.path)

    // `.docx` and `.xlsx` go to the sidecar, which rewrites only what was touched.
    const word = isWordPackagePath(path)
      ? await saveDocx(sidecar(), payload.content, {
          origin: payload.origin,
          destination: path,
          template: isWordTemplatePath(path),
        })
      : null
    const saved = word ?? (isExcelPath(path) ? await saveXlsx(sidecar(), payload.content) : null)
    await writeFileAtomic(path, saved?.bytes ?? payload.content)
    followDocxOriginal(payload.origin, path, word)

    rememberRecentFile(path)
    void refreshMenu()

    const result = { path, name: fileNameFromPath(path) }
    return saved === null ? result : { ...result, inventory: saved.inventory }
  })

  handle(IpcChannel.FileChooseSavePath, async (payload, event) => {
    const chosen = await showSaveFileDialog(windowOf(event), payload.suggestedName, payload.kind)
    if (chosen === null) return { canceled: true as const }

    // Only authorizes: writing is another call, so the loss warning comes before writing.
    const path = authorizePath(ensureSupportedExtension(chosen, payload.kind))
    return { canceled: false as const, path, name: fileNameFromPath(path) }
  })

  handle(IpcChannel.RecentList, async () => ({ files: [...(await listRecentFiles())] }))

  handle(IpcChannel.RecentClear, async () => {
    clearRecentFiles()
    await refreshMenu()
    return { files: [] }
  })
}
