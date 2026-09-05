import { DocumentKind, PlainTextChoice, type LossInventory } from '@shared/types.js'
import { DEFAULT_PAGE_SETUP, createEmptyDocument, type DocumentModel } from '@services/document/model.js'
import { documentToPlainText, hasRichFormatting, plainTextToDocument } from '@services/document/plain-text.js'
import { parseDocument, serializeDocument } from '@services/document/serialize.js'
import { createEmptyWorkbook } from '@services/spreadsheet/model.js'
import { parseWorkbook, serializeWorkbook } from '@services/spreadsheet/serialize.js'
import { recalculate } from '@services/spreadsheet/formula/recalc.js'
import { defaultFileName, isPlainTextPath, kindFromPath } from '@services/file/formats.js'
import { hasReportableLoss, locksEditing } from '@services/file/inventory.js'
import { toSerialized, type GetWorkspace, type SetWorkspace, type WorkspaceContext } from './context.js'
import type { LoadedFile, OpenFile, WorkspaceState } from './types.js'

/** O que o processo main devolve quando um arquivo abre. */
interface OpenedFile {
  readonly path: string
  readonly name: string
  readonly content: string
  // `| undefined` explícito por causa de `exactOptionalPropertyTypes`: o
  // contrato de IPC declara a propriedade como podendo vir indefinida.
  readonly inventory?: LossInventory | undefined
}

/** O que o usuário respondeu ao aviso de que `.txt` não guarda formatação. */
type PlainTextAnswer = 'proceed' | 'cancel' | 'chooseAnother'

type FileActions = Pick<
  WorkspaceState,
  | 'refreshRecents'
  | 'clearRecents'
  | 'newDocument'
  | 'newSpreadsheet'
  | 'openViaDialog'
  | 'openRecent'
  | 'save'
  | 'saveAs'
  | 'closeFile'
>

export function createFileActions(set: SetWorkspace, get: GetWorkspace, ctx: WorkspaceContext): FileActions {
  /**
   * O conteúdo a gravar num destino, no formato que a extensão dele pede.
   *
   * Planilha não passa pelo caminho de texto: não tem formatação a perder para
   * `.txt`, e o conteúdo é outro.
   */
  function encodeFor(path: string): string {
    const { workbook } = get()
    if (workbook !== null) return serializeWorkbook(workbook)

    const model = ctx.currentModel()
    return isPlainTextPath(path) ? documentToPlainText(model.doc) : serializeDocument(model)
  }

  /**
   * Salvar em `.txt` descartaria formatação. Perguntar antes é a regra do
   * projeto: nada se perde em silêncio.
   */
  async function confirmPlainTextLoss(path: string, fileName: string): Promise<PlainTextAnswer> {
    const { workbook } = get()
    if (workbook !== null || !isPlainTextPath(path)) return 'proceed'
    if (!hasRichFormatting(ctx.currentModel().doc)) return 'proceed'

    const answer = await ctx.call(() => window.api.dialog.confirmPlainText({ fileName }))
    if (answer === null || answer.choice === PlainTextChoice.Cancel) return 'cancel'
    return answer.choice === PlainTextChoice.SaveAsDocument ? 'chooseAnother' : 'proceed'
  }

  /** O disco passou a ter a versão boa: o rascunho não vale mais. */
  async function afterSave(file: OpenFile): Promise<void> {
    set({ file, isDirty: false })
    await ctx.forgetDraft()
    await get().refreshRecents()
  }

  async function openFile(fetch: () => Promise<OpenedFile | null>): Promise<void> {
    const opened = await fetch()
    if (opened === null) return

    try {
      ctx.show(interpret(opened))
      // O aviso só aparece quando há o que avisar: um alerta que abre em todo
      // arquivo é um alerta que o usuário fecha sem ler.
      set({
        notice: hasReportableLoss(opened.inventory) ? (opened.inventory ?? null) : null,
        readOnly: locksEditing(opened.inventory),
      })
    } catch (cause) {
      set({ error: toSerialized(cause) })
    }

    await get().refreshRecents()
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
        workbook: createEmptyWorkbook(),
      })
    },

    openViaDialog: async () => {
      if (!(await ctx.ensureChangesHandled())) return
      await openFile(async () => {
        const data = await ctx.call(() => window.api.file.open({}))
        return data === null || data.canceled ? null : data.file
      })
    },

    openRecent: async (path) => {
      if (!(await ctx.ensureChangesHandled())) return
      await openFile(async () => {
        const data = await ctx.call(() => window.api.file.openRecent({ path }))
        if (data === null) {
          // O arquivo pode ter sumido; a lista precisa refletir isso.
          await get().refreshRecents()
          return null
        }
        return data.file
      })
    },

    save: async () => {
      const file = get().file
      if (file === null) return false

      // Arquivo que nunca foi gravado não tem destino: vira "salvar como".
      const { path } = file
      if (path === null) return get().saveAs()

      const answer = await confirmPlainTextLoss(path, file.name)
      if (answer === 'cancel') return false
      if (answer === 'chooseAnother') return get().saveAs()

      const data = await ctx.call(() => window.api.file.save({ path, content: encodeFor(path) }))
      if (data === null) return false

      await afterSave({ ...file, name: data.name })
      return true
    },

    saveAs: async () => {
      const file = get().file
      if (file === null) return false

      const chosen = await ctx.call(() =>
        window.api.file.chooseSavePath({ suggestedName: file.name, kind: file.kind }),
      )
      if (chosen === null || chosen.canceled) return false

      // O aviso vem antes da gravação: se o usuário desistir aqui, nenhum byte
      // foi escrito e o destino continua como estava.
      const answer = await confirmPlainTextLoss(chosen.path, chosen.name)
      if (answer === 'cancel') return false
      // Quis preservar a formatação: escolhe outro destino.
      if (answer === 'chooseAnother') return get().saveAs()

      const data = await ctx.call(() =>
        window.api.file.save({ path: chosen.path, content: encodeFor(chosen.path) }),
      )
      if (data === null) return false

      await afterSave({ path: chosen.path, name: data.name, kind: file.kind })
      return true
    },

    closeFile: async () => {
      if (!(await ctx.ensureChangesHandled())) return

      const empty = createEmptyDocument()
      set((state) => ({
        file: null,
        workbook: null,
        page: empty.page,
        initialDoc: empty.doc,
        generation: state.generation + 1,
        isDirty: false,
        error: null,
        notice: null,
        readOnly: false,
      }))

      await ctx.forgetDraft()
      await get().refreshRecents()
    },
  }
}

/**
 * O que veio do disco, já no formato do editor.
 *
 * `.xlsx` chega aqui convertido pelo processo main, no mesmo envelope do
 * `.ssheet` — por isso a extensão decide o editor, e não o conteúdo.
 */
function interpret(opened: OpenedFile): LoadedFile {
  const kind = kindFromPath(opened.path)
  const file: OpenFile = { path: opened.path, name: opened.name, kind }

  if (kind === DocumentKind.Spreadsheet) {
    // Recalcula ao abrir: o arquivo guarda o valor de quando foi salvo, e uma
    // fórmula com HOJE() ou editada à mão estaria desatualizada.
    return { file, model: createEmptyDocument(), workbook: recalculate(parseWorkbook(opened.content)) }
  }

  return { file, model: decode(opened.path, opened.content), workbook: null }
}

/** Interpreta o conteúdo lido do disco conforme a extensão do arquivo. */
function decode(path: string, content: string): DocumentModel {
  if (isPlainTextPath(path)) {
    return { page: DEFAULT_PAGE_SETUP, doc: plainTextToDocument(content) }
  }
  return parseDocument(content)
}
