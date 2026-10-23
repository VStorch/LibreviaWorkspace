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
import { refuseViewOnly } from './view-only.js'
import {
  loadedState,
  toSerialized,
  workbookOf,
  type GetWorkspace,
  type SetWorkspace,
  type WorkspaceContext,
} from './context.js'
import type { LoadedFile, OpenFile, WorkspaceState } from './types.js'

interface OpenedFile {
  readonly path: string
  readonly name: string
  readonly content: string
  // `| undefined`: the IPC contract declares the property as possibly undefined.
  readonly inventory?: LossInventory | undefined
  readonly template?: boolean | undefined
}

type PlainTextAnswer = 'proceed' | 'cancel' | 'chooseAnother'

type FileActions = Pick<
  WorkspaceState,
  | 'refreshRecents'
  | 'clearRecents'
  | 'newDocument'
  | 'newSpreadsheet'
  | 'openViaDialog'
  | 'openRecent'
  | 'newFromTemplate'
  | 'save'
  | 'saveAs'
  | 'closeFile'
>

export function createFileActions(set: SetWorkspace, get: GetWorkspace, ctx: WorkspaceContext): FileActions {
  return { ...createOpenActions(set, get, ctx), ...createSaveActions(set, get, ctx) }
}

type OpenActions = Pick<
  FileActions,
  | 'refreshRecents'
  | 'clearRecents'
  | 'newDocument'
  | 'newSpreadsheet'
  | 'openViaDialog'
  | 'openRecent'
  | 'newFromTemplate'
>

function createOpenActions(set: SetWorkspace, get: GetWorkspace, ctx: WorkspaceContext): OpenActions {
  async function openFile(load: () => Promise<OpenedFile | null>): Promise<boolean> {
    const opened = await load()
    if (opened === null) return false

    try {
      ctx.show(interpret(opened))
      set(openedNotice(opened.inventory))
    } catch (cause) {
      set({ error: toSerialized(cause) })
      return false
    }

    await get().refreshRecents()
    return true
  }

  return {
    refreshRecents: async () => {
      const data = await ctx.call(() => window.api.recent.list({}))
      if (data !== null) set({ recents: data.files })
    },

    clearRecents: async () => {
      const data = await ctx.call(() => window.api.recent.clear({}))
      if (data !== null) set({ recents: data.files })
    },

    newDocument: async () => {
      if (!(await ctx.ensureChangesHandled())) return
      ctx.show({
        file: { path: null, name: defaultFileName(DocumentKind.Document), kind: DocumentKind.Document },
        model: createEmptyDocument(),
        workbook: null,
      })
    },

    newSpreadsheet: async () => {
      if (!(await ctx.ensureChangesHandled())) return
      ctx.show({
        file: {
          path: null,
          name: defaultFileName(DocumentKind.Spreadsheet),
          kind: DocumentKind.Spreadsheet,
        },
        model: createEmptyDocument(),
        workbook: createEmptyWorkbook(currentPreferences().language),
      })
    },

    openViaDialog: async () => {
      if (!(await ctx.ensureChangesHandled())) return
      await openFile(async () => {
        const data = await ctx.call(() => window.api.file.open({}))
        return data === null || data.canceled ? null : data.file
      })
    },

    newFromTemplate: async (template) => {
      if (!(await ctx.ensureChangesHandled())) return false
      return openFile(async () => {
        if (template !== null) {
          const data = await ctx.call(() => window.api.template.open(template))
          return data === null ? null : data.file
        }
        const data = await ctx.call(() => window.api.template.browse({}))
        return data === null || data.canceled ? null : data.file
      })
    },

    openRecent: async (path) => {
      if (!(await ctx.ensureChangesHandled())) return
      await openFile(async () => {
        const data = await ctx.call(() => window.api.file.openRecent({ path }))
        if (data === null) {
          await get().refreshRecents()
          return null
        }
        return data.file
      })
    },
  }
}

function createSaveActions(
  set: SetWorkspace,
  get: GetWorkspace,
  ctx: WorkspaceContext,
): Pick<FileActions, 'save' | 'saveAs' | 'closeFile'> {
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
    await get().refreshRecents()
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

    saveAs: async () => {
      const file = get().file
      if (file === null) return false
      if (file.kind === DocumentKind.Pdf) return refuseViewOnly(set)

      const chosen = await ctx.call(() =>
        window.api.file.chooseSavePath({ suggestedName: file.name, kind: file.kind }),
      )
      if (chosen === null || chosen.canceled) return false

      // Before writing: giving up here writes no byte.
      const answer = await confirmPlainTextLoss(chosen.path, chosen.name)
      if (answer === 'cancel') return false
      if (answer === 'chooseAnother') return get().saveAs()

      const data = await ctx.call(() =>
        window.api.file.save({
          path: chosen.path,
          content: encodeFor(chosen.path),
          // A document created from a template saves from its package.
          origin: file.origin ?? file.path,
        }),
      )
      if (data === null) return false

      await afterSave({ path: chosen.path, name: data.name, kind: file.kind }, data.inventory)
      return true
    },

    closeFile: async () => {
      if (!(await ctx.ensureChangesHandled())) return

      set((state) => loadedState(null, createEmptyDocument(), null, state.generation))
      await ctx.forgetDraft()
      await get().refreshRecents()
    },
  }
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
  // A Word template opens without a path, so "save" never writes over it.
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
