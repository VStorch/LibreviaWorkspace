import { AppError, ErrorCode } from '@shared/errors.js'
import { IpcChannel } from '@shared/ipc-channels.js'
import type { LoadedFile, LossInventory } from '@shared/types.js'
import {
  ensureSupportedExtension,
  extensionOf,
  fileNameFromPath,
  isExcelPath,
  isOdtPackagePath,
  isOdtTemplatePath,
  isPdfPath,
  isWordPackagePath,
  isWordTemplatePath,
  kindFromPath,
  ODT_EXTENSION,
  WORD_EXTENSION,
} from '@services/file/formats.js'
import { showOpenFileDialog, showSaveFileDialog } from '../dialogs.js'
import { forgetOpenedDocx, openDocx, saveDocx, type SavedDocx } from '../docx/index.js'
import { openOdt, saveOdt } from '../odt/index.js'
import { odtGenerator } from '../odt/generator.js'
import { followOriginal, OriginalFormat, type SavedOriginal } from '../opened-original.js'
import { forgetOpenedXlsx, openXlsx, saveXlsx } from '../xlsx/index.js'
import { sidecar } from '../sidecar/index.js'
import { writeFileAtomic } from '../fs/atomic-write.js'
import {
  allowViewing,
  assertPathAuthorized,
  assertReadableFile,
  assertViewablePdf,
  authorizePath,
  normalizePath,
} from '../fs/paths.js'
import { clearRecentFiles, isRemembered, listRecentFiles, rememberRecentFile } from '../fs/recent.js'
import { readTextFile } from '../fs/read-text.js'
import { refreshMenu } from '../menu.js'
import { t } from '../i18n.js'
import { exportName } from './export.js'
import { handle } from './registry.js'
import { windowOf } from './sender-window.js'
import { forgetExternalFileRequest, isExternalFileRequested } from '../external-files.js'

interface LoadedContent {
  readonly content: string
  readonly inventory: LossInventory | undefined
}

/**
 * A Word or ODT template opens as a new document, and its path is **not** authorized: nothing writes over
 * it. `templateName` is the new document's name: the builtin template's, translated; when absent,
 * the file's.
 */
export async function loadFile(path: string, templateName?: string): Promise<LoadedFile> {
  if (isPdfPath(path)) return loadPdf(path)
  await assertReadableFile(path)

  const loaded = await loadContent(path)

  // The original bytes apply to **one** file: otherwise the next save would write into the wrong
  // package.
  if (!isWordPackagePath(path) && !isOdtPackagePath(path)) forgetOpenedDocx()
  if (!isExcelPath(path)) forgetOpenedXlsx()

  if (isWordTemplatePath(path)) return loadAsNewDocument(path, loaded, WORD_EXTENSION, templateName)
  if (isOdtTemplatePath(path)) return loadAsNewDocument(path, loaded, ODT_EXTENSION, templateName)

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

/** DOCX, ODT and XLSX become the internal format here: the renderer never sees the package. */
async function loadContent(path: string): Promise<LoadedContent> {
  if (isWordPackagePath(path)) return openDocx(sidecar(), path)
  if (isOdtPackagePath(path)) return openOdt(sidecar(), path)
  if (isExcelPath(path)) return openXlsx(sidecar(), path)
  return { content: await readTextFile(path), inventory: undefined }
}

/**
 * Named after the template's document format: "Save" asks where, through "Save as". The template's
 * package stays as the original, so saving into its format rewrites only what was edited.
 */
function loadAsNewDocument(
  path: string,
  loaded: LoadedContent,
  extension: string,
  templateName?: string,
): LoadedFile {
  const source = normalizePath(path)
  rememberRecentFile(source)
  void refreshMenu()
  const name = fileNameFromPath(source)
  const base = templateName ?? name.slice(0, name.length - extensionOf(source).length)
  const file: LoadedFile = {
    path: source,
    name: `${base}${extension}`,
    kind: kindFromPath(source),
    content: loaded.content,
    template: true,
  }
  return loaded.inventory === undefined ? file : { ...file, inventory: loaded.inventory }
}

/**
 * No content: Chromium reads the PDF from disk when the viewer attaches. The path is only allowed for
 * viewing, never authorized: nothing is ever written to it.
 */
async function loadPdf(path: string): Promise<LoadedFile> {
  await assertViewablePdf(path)
  forgetOpenedDocx()
  forgetOpenedXlsx()

  const source = allowViewing(path)
  rememberRecentFile(source)
  void refreshMenu()
  return { path: source, name: fileNameFromPath(source), kind: kindFromPath(source), content: '' }
}

/** A name typed without `.docx` still goes out as Word, which is what was asked. */
function asWordPath(path: string): string {
  return extensionOf(path) === WORD_EXTENSION ? path : `${path}${WORD_EXTENSION}`
}

interface Encoded {
  readonly saved: { readonly bytes: Uint8Array; readonly inventory: LossInventory } | null
  /** The package the next save starts from, when the destination has one. */
  readonly original: SavedOriginal | null
}

/** Word, ODT and `.xlsx` packages go to the sidecar, which rewrites only what was touched. */
async function encodeForDestination(path: string, content: string, origin: string | null): Promise<Encoded> {
  if (isWordPackagePath(path)) {
    const word = await saveDocx(sidecar(), content, {
      origin,
      destination: path,
      template: isWordTemplatePath(path),
    })
    return { saved: word, original: startingPackage(word, OriginalFormat.Docx) }
  }
  if (isOdtPackagePath(path)) {
    const odt = await saveOdt(sidecar(), content, {
      origin,
      generator: odtGenerator(),
      template: isOdtTemplatePath(path),
    })
    return { saved: odt, original: { ...startingPackage(odt, OriginalFormat.Odt), created: odt.created } }
  }
  return { saved: isExcelPath(path) ? await saveXlsx(sidecar(), content) : null, original: null }
}

function startingPackage(saved: SavedDocx, format: OriginalFormat): SavedOriginal {
  return { original: saved.original, format }
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
    const { saved, original } = await encodeForDestination(path, payload.content, payload.origin)
    await writeFileAtomic(path, saved?.bytes ?? payload.content)
    followOriginal(payload.origin, path, original)

    rememberRecentFile(path)
    void refreshMenu()

    const result = { path, name: fileNameFromPath(path) }
    return saved === null ? result : { ...result, inventory: saved.inventory }
  })

  handle(IpcChannel.FileChooseSavePath, async (payload, event) => {
    const { wordOnly } = payload
    const chosen = await showSaveFileDialog(
      windowOf(event),
      wordOnly ? exportName(payload.suggestedName, WORD_EXTENSION.slice(1)) : payload.suggestedName,
      payload.kind,
      { wordOnly },
    )
    if (chosen === null) return { canceled: true as const }

    const destination = wordOnly ? asWordPath(chosen) : ensureSupportedExtension(chosen, payload.kind)
    // Only authorizes: writing is another call, so the loss warning comes before writing.
    const path = authorizePath(destination)
    return { canceled: false as const, path, name: fileNameFromPath(path) }
  })

  handle(IpcChannel.RecentList, async () => ({ files: [...(await listRecentFiles())] }))

  handle(IpcChannel.RecentClear, async () => {
    clearRecentFiles()
    await refreshMenu()
    return { files: [] }
  })
}
