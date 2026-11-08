import { DocumentKind, PlainTextChoice, type LossInventory } from '@shared/types.js'
import { DEFAULT_PAGE_SETUP, createEmptyDocument, type DocumentModel } from '@services/document/model.js'
import { documentToPlainText, hasRichFormatting, plainTextToDocument } from '@services/document/plain-text.js'
import { parseDocument, serializeDocument } from '@services/document/serialize.js'
import { stampProperties, w3cdtf, type DocumentStatistics } from '@services/file/properties.js'
import { BUILTIN_STYLES } from '@services/document/styles.js'
import { createEmptyWorkbook } from '@services/spreadsheet/model.js'
import { parseWorkbook, serializeWorkbook } from '@services/spreadsheet/serialize.js'
import { recalculate } from '@services/spreadsheet/formula/recalc.js'
import { defaultFileName, isPlainTextPath, isWordPackagePath, kindFromPath } from '@services/file/formats.js'
import { documentStatistics } from '@services/document/word-count.js'
import { hasReportableLoss, locksEditing, lostOnSave } from '@services/file/inventory.js'
import { currentPreferences } from './preferences.js'
import { t } from '../i18n.js'
import { refuseViewOnly } from './view-only.js'
import { refreshRecents } from './recents.js'
import { isOpenInAnotherTab, isSamePath } from './path-holder.js'
import {
  toSerialized,
  workbookOf,
  type GetWorkspace,
  type SetWorkspace,
  type WorkspaceContext,
  releaseFile,
} from './context.js'
import type { LoadedFile, OpenedFile, OpenFile, WorkspaceState } from './types.js'

type PlainTextAnswer = 'proceed' | 'cancel' | 'chooseAnother'

type FileActions = Pick<
  WorkspaceState,
  'loadFile' | 'startNew' | 'save' | 'saveAs' | 'confirmClose' | 'release' | 'suspend'
>

export function createFileActions(set: SetWorkspace, get: GetWorkspace, ctx: WorkspaceContext): FileActions {
  return {
    ...createOpenActions(set, ctx),
    ...createSaveActions(set, get, ctx),
    ...createTabLifecycle(set, get, ctx),
  }
}

function createOpenActions(
  set: SetWorkspace,
  ctx: WorkspaceContext,
): Pick<FileActions, 'loadFile' | 'startNew'> {
  return {
    loadFile: (opened) => {
      try {
        ctx.show(interpret(opened))
        set(openedNotice(opened.inventory))
        return true
      } catch (cause) {
        set({ error: toSerialized(cause) })
        return false
      }
    },

    startNew: (kind) => {
      ctx.show({
        file: { path: null, name: untitledName(kind), kind },
        model: createEmptyDocument(),
        workbook:
          kind === DocumentKind.Spreadsheet ? createEmptyWorkbook(currentPreferences().language) : null,
      })
    },
  }
}

function createTabLifecycle(
  set: SetWorkspace,
  get: GetWorkspace,
  ctx: WorkspaceContext,
): Pick<FileActions, 'confirmClose' | 'release' | 'suspend'> {
  return {
    confirmClose: () => ctx.ensureChangesHandled(),

    release: async () => {
      const kept = keptOriginalOf(get().file)
      if (kept !== null) releaseFile(kept, get())
      await ctx.forgetDraft()
    },

    suspend: () => {
      const source = ctx.source()
      if (source !== null) set({ initialDoc: source.readDoc() })
    },
  }
}

/** A document created from a template saves from the template's package. */
function keptOriginalOf(file: OpenFile | null): string | null {
  return file?.origin ?? file?.path ?? null
}

/** Two tabs on one file would each save over the other. */
function refuseHeldPath(set: SetWorkspace, name: string): false {
  set({ error: { code: 'INTERNAL', message: t('shell.saveAs.openInAnotherTab', { name }) } })
  return false
}

/** A spreadsheet has no Word format to go to. */
function refuseWordExport(set: SetWorkspace): false {
  set({ error: { code: 'INTERNAL', message: t('shell.export.documentOnly') } })
  return false
}

function createSaveActions(
  set: SetWorkspace,
  get: GetWorkspace,
  ctx: WorkspaceContext,
): Pick<FileActions, 'save' | 'saveAs'> {
  /** A spreadsheet does not take the text path: it has no formatting to lose. */
  function encodeFor(path: string): string {
    if (get().workbook !== null || !isPlainTextPath(path)) stampForSave(set, get)
    const { workbook } = get()
    if (workbook !== null) return serializeWorkbook(workbookOf(workbook, get()))

    const model = ctx.currentModel()
    if (isPlainTextPath(path)) return documentToPlainText(model.doc)
    return serializeDocument(model, isWordPackagePath(path) ? statisticsToSave(model, get()) : undefined)
  }

  /** Nothing is lost silently. */
  async function confirmPlainTextLoss(path: string, fileName: string): Promise<PlainTextAnswer> {
    const { workbook } = get()
    if (workbook !== null || !isPlainTextPath(path)) return 'proceed'
    if (!hasRichFormatting(ctx.currentModel().doc)) return 'proceed'

    const answer = await ctx.call(() => window.api.dialog.confirmPlainText({ fileName }))
    if (answer === null || answer.choice === PlainTextChoice.Cancel) return 'cancel'
    return answer.choice === PlainTextChoice.SaveAsDocument ? 'chooseAnother' : 'proceed'
  }

  /**
   * The draft no longer counts; what saving lost goes to the banner, and a save that lost nothing
   * clears the previous notice.
   */
  async function afterSave(file: OpenFile, inventory: LossInventory | undefined): Promise<void> {
    const lost = lostOnSave(inventory)
    set({ file, isDirty: false, textEdited: false, savedLoss: lost.length > 0 ? lost : null })
    await ctx.forgetDraft()
    await refreshRecents()
  }

  return {
    save: async () => {
      const file = get().file
      if (file === null) return false
      if (file.kind === DocumentKind.Pdf) {
        if (!(await savePdfProperties(set, get, ctx, file))) return false
        await afterSave(file, undefined)
        return true
      }

      const { path } = file
      if (path === null) return get().saveAs()

      const answer = await confirmPlainTextLoss(path, file.name)
      if (answer === 'cancel') return false
      if (answer === 'chooseAnother') return get().saveAs()

      const data = await ctx.call(() =>
        window.api.file.save({ path, content: encodeFor(path), origin: path }),
      )
      if (data === null) return false

      await afterSave({ ...file, name: data.name }, data.inventory)
      return true
    },

    saveAs: async (options) => {
      const file = get().file
      if (file === null) return false
      if (file.kind === DocumentKind.Pdf) return refuseViewOnly(set)
      const wordOnly = options?.wordOnly === true
      if (wordOnly && file.kind !== DocumentKind.Document) return refuseWordExport(set)

      const chosen = await ctx.call(() => chooseSavePath(file, wordOnly))
      if (chosen === null || chosen.canceled) return false
      if (isOpenInAnotherTab(chosen.path, get())) return refuseHeldPath(set, chosen.name)

      // Before writing: giving up here writes no byte.
      const answer = await confirmPlainTextLoss(chosen.path, chosen.name)
      if (answer === 'cancel') return false
      if (answer === 'chooseAnother') return get().saveAs()

      const origin = keptOriginalOf(file)
      const data = await ctx.call(() =>
        window.api.file.save({ path: chosen.path, content: encodeFor(chosen.path), origin }),
      )
      if (data === null) return false

      // Main keeps the origin's package too: another tab made from the same template may need it.
      if (origin !== null && !isSamePath(origin, chosen.path)) releaseFile(origin, get())
      await afterSave({ path: chosen.path, name: data.name, kind: file.kind }, data.inventory)
      return true
    },
  }
}

function untitledName(kind: DocumentKind): string {
  const { language, saveFormat } = currentPreferences()
  return defaultFileName(kind, language, saveFormat)
}

function chooseSavePath(
  file: OpenFile,
  wordOnly: boolean,
): ReturnType<typeof window.api.file.chooseSavePath> {
  return window.api.file.chooseSavePath({
    suggestedName: file.name,
    kind: file.kind,
    wordOnly,
    fresh: file.path === null,
  })
}

/**
 * Only when the document changed or was never saved: opened and saved without editing, `docProps/`
 * goes back byte for byte. It stays in the state even if saving fails.
 */
/** A PDF only has its properties to save; the rest of it is read only. */
async function savePdfProperties(
  set: SetWorkspace,
  get: GetWorkspace,
  ctx: WorkspaceContext,
  file: OpenFile,
): Promise<boolean> {
  const { path } = file
  const { isDirty, properties } = get()
  if (path === null || !isDirty || properties === undefined) return refuseViewOnly(set)

  const stamped = { ...properties, modified: w3cdtf(new Date()) }
  const data = await ctx.call(() => window.api.pdf.saveProperties({ path, properties: stamped }))
  if (data === null) return false

  set({ properties: stamped })
  return true
}

function stampForSave(set: SetWorkspace, get: GetWorkspace): void {
  const state = get()
  const stamped = stampProperties(state.properties, {
    author: currentPreferences().authorName,
    now: new Date(),
    fresh: state.file?.path === null,
    edited: state.isDirty,
  })
  if (stamped !== state.properties) set({ properties: stamped })
}

/**
 * Only for Word and only after the text changed: with only the properties edited, `app.xml` goes
 * back byte for byte.
 */
function statisticsToSave(model: DocumentModel, state: WorkspaceState): DocumentStatistics | undefined {
  if (!state.textEdited && state.file?.path !== null) return undefined
  return documentStatistics(model.doc, {
    pages: state.pageCount,
    words: state.stats.words,
    charactersWithSpaces: state.stats.characters,
  })
}

/** What opening a file says about what it lost or hides. */
export function openedNotice(
  inventory: LossInventory | undefined,
): Pick<WorkspaceState, 'notice' | 'readOnly'> {
  return {
    notice: hasReportableLoss(inventory) ? (inventory ?? null) : null,
    readOnly: locksEditing(inventory),
  }
}

/** `.xlsx` arrives converted in the `.ssheet` envelope: the extension picks the editor. */
export function interpret(opened: OpenedFile): LoadedFile {
  const kind = kindFromPath(opened.path)
  // A template opens without a path, so "save" never writes over it.
  const file: OpenFile =
    opened.template === true
      ? { path: null, name: opened.name, kind, origin: opened.path }
      : { path: opened.path, name: opened.name, kind }

  // The viewer reads the PDF from disk: there is no model to build.
  if (kind === DocumentKind.Pdf) return { file, model: createEmptyDocument(), workbook: null }

  if (kind === DocumentKind.Spreadsheet) {
    // Recalculates on open: `HOJE()` and a hand-edited formula would be stale.
    return { file, model: createEmptyDocument(), workbook: recalculate(parseWorkbook(opened.content)) }
  }

  return {
    file,
    model: opened.template === true ? parseDocument(opened.content) : decode(opened.path, opened.content),
    workbook: null,
  }
}

function decode(path: string, content: string): DocumentModel {
  if (isPlainTextPath(path)) {
    // Plain text gets the new document's styles.
    return { page: DEFAULT_PAGE_SETUP, doc: plainTextToDocument(content), styles: BUILTIN_STYLES }
  }
  return parseDocument(content)
}
