import { DocumentKind, PlainTextChoice, type LossInventory } from '@shared/types.js'
import { DEFAULT_PAGE_SETUP, createEmptyDocument, type DocumentModel } from '@services/document/model.js'
import { documentToPlainText, hasRichFormatting, plainTextToDocument } from '@services/document/plain-text.js'
import { parseDocument, serializeDocument } from '@services/document/serialize.js'
import { stampProperties } from '@services/document/properties.js'
import { BUILTIN_STYLES } from '@services/document/styles.js'
import { createEmptyWorkbook } from '@services/spreadsheet/model.js'
import { parseWorkbook, serializeWorkbook } from '@services/spreadsheet/serialize.js'
import { recalculate } from '@services/spreadsheet/formula/recalc.js'
import { defaultFileName, isPlainTextPath, kindFromPath } from '@services/file/formats.js'
import { hasReportableLoss, locksEditing, lostOnSave } from '@services/file/inventory.js'
import { currentPreferences } from './preferences.js'
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
  readonly template?: boolean | undefined
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
  | 'newFromTemplate'
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
    if (!isPlainTextPath(path)) stampForSave()

    const model = ctx.currentModel()
    return isPlainTextPath(path) ? documentToPlainText(model.doc) : serializeDocument(model)
  }

  /**
   * Quem modificou, quando e a revisão (M11), no estado e portanto no que vai ao
   * disco — ver `stampProperties`. Só quando o documento mudou, ou quando nunca
   * foi gravado: o arquivo aberto e salvo sem edição volta com `docProps/` byte a
   * byte. Fica no estado mesmo que a gravação falhe: é só a data da tentativa.
   */
  function stampForSave(): void {
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

  /**
   * O disco passou a ter a versão boa: o rascunho não vale mais.
   *
   * E o que a gravação não conseguiu levar ao disco vai para a faixa de aviso —
   * nada se perde em silêncio, nem na hora de salvar. Cada gravação diz só o
   * que perdeu **ela**: a que não perdeu nada apaga o aviso da anterior.
   */
  async function afterSave(file: OpenFile, inventory: LossInventory | undefined): Promise<void> {
    const lost = lostOnSave(inventory)
    set({ file, isDirty: false, savedLoss: lost.length > 0 ? lost : null })
    await ctx.forgetDraft()
    await get().refreshRecents()
  }

  async function openFile(fetch: () => Promise<OpenedFile | null>): Promise<boolean> {
    const opened = await fetch()
    if (opened === null) return false

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
        window.api.file.save({
          path: chosen.path,
          content: encodeFor(chosen.path),
          // O documento criado a partir de um modelo (M11) ainda não tem caminho:
          // a origem é o modelo, e é sobre o pacote dele que a gravação parte.
          origin: file.origin ?? file.path,
        }),
      )
      if (data === null) return false

      await afterSave({ path: chosen.path, name: data.name, kind: file.kind }, data.inventory)
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
        styles: empty.styles,
        flattened: false,
        beforeReferences: false,
        sections: [],
        beforeSections: false,
        outsideBookmarks: [],
        comments: [],
        commentsOutside: [],
        commentDraft: null,
        beforeComments: false,
        trackChanges: undefined,
        beforeRevisions: false,
        notes: undefined,
        beforeNotes: false,
        properties: undefined,
        generation: state.generation + 1,
        isDirty: false,
        error: null,
        notice: null,
        savedLoss: null,
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
  // O modelo do Word (M11) abre como documento novo: sem caminho, para que
  // "salvar" pergunte o destino e nunca grave por cima do modelo.
  const file: OpenFile =
    opened.template === true
      ? { path: null, name: opened.name, kind, origin: opened.path }
      : { path: opened.path, name: opened.name, kind }

  if (kind === DocumentKind.Spreadsheet) {
    // Recalcula ao abrir: o arquivo guarda o valor de quando foi salvo, e uma
    // fórmula com HOJE() ou editada à mão estaria desatualizada.
    return { file, model: createEmptyDocument(), workbook: recalculate(parseWorkbook(opened.content)) }
  }

  return {
    file,
    model: opened.template === true ? parseDocument(opened.content) : decode(opened.path, opened.content),
    workbook: null,
  }
}

/** Interpreta o conteúdo lido do disco conforme a extensão do arquivo. */
function decode(path: string, content: string): DocumentModel {
  if (isPlainTextPath(path)) {
    // Texto simples não tem estilo nenhum a trazer: recebe os do documento novo,
    // que são a aparência com que o editor já o desenhava.
    return { page: DEFAULT_PAGE_SETUP, doc: plainTextToDocument(content), styles: BUILTIN_STYLES }
  }
  return parseDocument(content)
}
