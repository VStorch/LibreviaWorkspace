import { create } from 'zustand'
import { AppError, type SerializedError } from '@shared/errors.js'
import type { IpcResult } from '@shared/ipc.js'
import {
  DiscardChoice,
  DocumentKind,
  PlainTextChoice,
  type DraftSummary,
  type LossInventory,
  type RecentFile,
} from '@shared/types.js'
import {
  DEFAULT_PAGE_SETUP,
  createEmptyDocument,
  type DocumentModel,
  type DocumentNode,
  type PageSetup,
} from '@services/document/model.js'
import { documentToPlainText, hasRichFormatting, plainTextToDocument } from '@services/document/plain-text.js'
import { buildPrintHtml } from '@services/document/print-html.js'
import { parseDocument, serializeDocument } from '@services/document/serialize.js'
import {
  createEmptyWorkbook,
  createSheet,
  type Sheet,
  type WorkbookModel,
} from '@services/spreadsheet/model.js'
import { parseWorkbook, serializeWorkbook } from '@services/spreadsheet/serialize.js'
import { recalculate } from '@services/spreadsheet/formula/recalc.js'
import {
  applyStructuralChange,
  renameSheet as renameSheetIn,
  type StructuralChange,
} from '@services/spreadsheet/structure.js'
import { defaultFileName, isPlainTextPath, kindFromPath } from '@services/file/formats.js'

interface OpenFile {
  /** `null` enquanto o arquivo nunca foi gravado. */
  readonly path: string | null
  readonly name: string
  readonly kind: DocumentKind
}

interface WorkspaceState {
  file: OpenFile | null
  page: PageSetup
  /**
   * Documento com que o editor é inicializado. Não acompanha a digitação — o
   * conteúdo ao vivo mora no editor, e só é lido na hora de salvar.
   */
  initialDoc: DocumentNode
  /**
   * A planilha aberta, quando o que está em edição é uma planilha.
   *
   * Fica ao lado do documento em vez de substituí-lo porque o aplicativo edita
   * um arquivo por vez: qual dos dois está preenchido é o que diz o tipo.
   */
  workbook: WorkbookModel | null
  /** Muda a cada novo/abrir; o editor é remontado, evitando estado residual. */
  generation: number
  isDirty: boolean
  /** Contagens exibidas na barra de status; alimentadas pelo editor. */
  stats: { characters: number; words: number }
  /**
   * Páginas estimadas, medidas na tela pelo editor. É **estimativa**: a
   * paginação de verdade só acontece na exportação, com regras de viúvas e
   * órfãs que o editor não aplica. Por isso a barra mostra "≈".
   */
  estimatedPages: number
  recents: readonly RecentFile[]
  error: SerializedError | null
  /**
   * O que o documento aberto tem e o editor não dá conta. Fica separado do
   * `error` porque não é falha: o arquivo abriu, e isto é informação sobre o
   * que você vai e o que você não vai ver.
   */
  notice: LossInventory | null
  /**
   * Rascunho de uma sessão que não terminou bem, esperando decisão.
   *
   * Enquanto ele está aqui o autosave não escreve: gravar por cima do rascunho
   * antes de o usuário decidir apagaria justamente o trabalho que ele existe
   * para devolver.
   */
  pendingDraft: DraftSummary | null
  /**
   * O autosave falhou e parou de tentar.
   *
   * Insistir a cada oito segundos numa gravação que não vai dar certo encheria a
   * tela de avisos; ficar tentando em silêncio deixaria o usuário confiando numa
   * rede de proteção que não existe mais.
   */
  autosaveBroken: boolean
  busy: boolean

  /**
   * O editor se registra aqui para que salvar e imprimir consigam ler o
   * conteúdo atual. O HTML vem do próprio editor, e não de uma segunda
   * renderização a partir do modelo — é o que garante que o PDF saia igual
   * ao que está na tela.
   */
  registerDocumentSource: (source: DocumentSource | null) => void
  markDirty: () => void
  setStats: (stats: { characters: number; words: number }) => void
  setEstimatedPages: (pages: number) => void
  setPage: (page: PageSetup) => void
  dismissError: () => void
  dismissNotice: () => void
  showError: (error: SerializedError) => void
  refreshRecents: () => Promise<void>

  newDocument: () => Promise<void>
  newSpreadsheet: () => Promise<void>
  updateSheet: (sheet: Sheet) => void
  changeStructure: (change: StructuralChange) => void
  selectSheet: (index: number) => void
  addSheet: () => void
  renameSheet: (index: number, name: string) => void
  removeSheet: (index: number) => void
  openViaDialog: () => Promise<void>
  openRecent: (path: string) => Promise<void>
  save: () => Promise<boolean>
  saveAs: () => Promise<boolean>
  closeFile: () => Promise<void>
  clearRecents: () => Promise<void>

  /** Guarda o que está na tela como rascunho. Chamado por relógio, não por tecla. */
  autosave: () => Promise<void>
  checkRecovery: () => Promise<void>
  recoverDraft: () => Promise<void>
  dismissDraft: () => Promise<void>

  exportPdf: () => Promise<boolean>
  print: () => Promise<boolean>
  printPreview: () => Promise<void>
}

export interface DocumentSource {
  readonly readDoc: () => DocumentNode
  readonly readHtml: () => string
}

function toSerialized(cause: unknown): SerializedError {
  if (cause instanceof AppError) return cause.toSerialized()
  return { code: 'INTERNAL', message: 'Ocorreu um erro inesperado. A operação não foi concluída.' }
}

export const useWorkspace = create<WorkspaceState>((set, get) => {
  let documentSource: DocumentSource | null = null

  /** Executa uma chamada de IPC e transforma a falha em erro exibível. */
  async function call<T>(operation: () => Promise<IpcResult<T>>): Promise<T | null> {
    set({ busy: true })
    try {
      const result = await operation()
      if (result.ok) return result.data
      set({ error: result.error })
      return null
    } catch {
      set({
        error: { code: 'INTERNAL', message: 'A comunicação com o aplicativo falhou. Tente novamente.' },
      })
      return null
    } finally {
      set({ busy: false })
    }
  }

  /** Modelo com o conteúdo que está na tela neste instante. */
  function currentModel(): DocumentModel {
    const state = get()
    return { page: state.page, doc: documentSource?.readDoc() ?? state.initialDoc }
  }

  /** O que está na tela, no formato interno — o mesmo que o rascunho guarda. */
  function currentContent(): string {
    const { workbook } = get()
    return workbook === null ? serializeDocument(currentModel()) : serializeWorkbook(workbook)
  }

  /**
   * Apaga o rascunho porque ele deixou de valer.
   *
   * Vale depois de gravar (o disco passou a ter a versão boa) e ao trocar de
   * arquivo (o rascunho é de outro trabalho). Falha aqui é engolida de
   * propósito: sobrar um rascunho velho custa um aviso a mais na próxima
   * abertura, e derrubar a gravação por causa disso custaria o arquivo.
   */
  async function forgetDraft(): Promise<void> {
    set({ autosaveBroken: false })
    await window.api.recovery.discard({}).catch(() => undefined)
  }

  /**
   * Portão de proteção contra perda de trabalho.
   *
   * Devolve `false` quando o usuário desistiu da ação — e nesse caso quem
   * chamou precisa parar sem alterar nada.
   */
  async function ensureChangesHandled(): Promise<boolean> {
    const state = get()
    if (!state.isDirty) return true

    const answer = await call(() =>
      window.api.dialog.confirmDiscard({ fileName: state.file?.name ?? 'Sem título' }),
    )
    if (answer === null) return false

    if (answer.choice === DiscardChoice.Cancel) return false
    if (answer.choice === DiscardChoice.Discard) return true
    return get().save()
  }

  function load(file: OpenFile, model: DocumentModel): void {
    // O rascunho é de outro trabalho a partir de agora. Vale inclusive para a
    // recuperação: o conteúdo já está na tela e marcado como não salvo, então o
    // relógio do autosave o grava de novo em segundos.
    void forgetDraft()

    set((state) => ({
      file,
      page: model.page,
      initialDoc: model.doc,
      generation: state.generation + 1,
      isDirty: false,
      error: null,
    }))
  }

  /** Interpreta o conteúdo lido do disco conforme a extensão do arquivo. */
  function decode(path: string, content: string): DocumentModel {
    if (isPlainTextPath(path)) {
      return { page: DEFAULT_PAGE_SETUP, doc: plainTextToDocument(content) }
    }
    return parseDocument(content)
  }

  async function openPath(
    load1: () => Promise<{
      path: string
      name: string
      content: string
      // `| undefined` explícito por causa de `exactOptionalPropertyTypes`: o
      // contrato de IPC declara a propriedade como podendo vir indefinida.
      inventory?: LossInventory | undefined
    } | null>,
  ) {
    const file = await load1()
    if (file === null) return

    try {
      // `.xlsx` chega aqui já convertido pelo processo main, no mesmo envelope
      // do `.ssheet` — por isso a extensão decide o editor, e não o conteúdo.
      if (kindFromPath(file.path) === DocumentKind.Spreadsheet) {
        // Recalcula ao abrir: o arquivo guarda o valor de quando foi salvo, e
        // uma fórmula com HOJE() ou editada à mão estaria desatualizada.
        const workbook = recalculate(parseWorkbook(file.content))
        load({ path: file.path, name: file.name, kind: DocumentKind.Spreadsheet }, createEmptyDocument())
        set({ workbook, notice: hasSomethingToSay(file.inventory) ? file.inventory! : null })
        await get().refreshRecents()
        return
      }

      load({ path: file.path, name: file.name, kind: DocumentKind.Document }, decode(file.path, file.content))
      set({ workbook: null })
      // O aviso só aparece quando há o que avisar: um alerta que abre em todo
      // arquivo é um alerta que o usuário fecha sem ler.
      set({ notice: hasSomethingToSay(file.inventory) ? file.inventory! : null })
    } catch (cause) {
      set({ error: toSerialized(cause) })
    }
    await get().refreshRecents()
  }

  /** "Planilha2", "Planilha3"… pulando os nomes já usados. */
  const nextSheetName = (workbook: WorkbookModel): string => {
    const used = new Set(workbook.sheets.map((sheet) => sheet.name))
    let index = workbook.sheets.length + 1
    while (used.has(`Planilha${index}`)) index += 1
    return `Planilha${index}`
  }

  const hasSomethingToSay = (inventory: LossInventory | undefined): boolean =>
    inventory !== undefined && (inventory.invisible.length > 0 || inventory.lost.length > 0)

  return {
    file: null,
    page: DEFAULT_PAGE_SETUP,
    initialDoc: createEmptyDocument().doc,
    workbook: null,
    generation: 0,
    isDirty: false,
    stats: { characters: 0, words: 0 },
    estimatedPages: 1,
    recents: [],
    error: null,
    notice: null,
    pendingDraft: null,
    autosaveBroken: false,
    busy: false,

    registerDocumentSource: (source) => {
      documentSource = source
    },
    markDirty: () => {
      if (!get().isDirty) set({ isDirty: true })
    },
    setStats: (stats) => {
      const current = get().stats
      if (current.characters !== stats.characters || current.words !== stats.words) set({ stats })
    },
    setEstimatedPages: (pages) => {
      if (get().estimatedPages !== pages) set({ estimatedPages: pages })
    },
    setPage: (page) => set({ page, isDirty: true }),
    dismissError: () => set({ error: null }),
    dismissNotice: () => set({ notice: null }),
    showError: (error) => set({ error }),

    refreshRecents: async () => {
      const data = await call(() => window.api.recent.list({}))
      if (data !== null) set({ recents: data.files })
    },

    newDocument: async () => {
      if (!(await ensureChangesHandled())) return
      const empty = createEmptyDocument()
      load({ path: null, name: defaultFileName(DocumentKind.Document), kind: DocumentKind.Document }, empty)
      set({ workbook: null })
    },

    newSpreadsheet: async () => {
      if (!(await ensureChangesHandled())) return
      load(
        {
          path: null,
          name: defaultFileName(DocumentKind.Spreadsheet),
          kind: DocumentKind.Spreadsheet,
        },
        createEmptyDocument(),
      )
      set({ workbook: createEmptyWorkbook() })
    },

    /**
     * Toda mudança de célula passa por aqui, e por isso o recálculo mora aqui.
     *
     * Um só ponto de recálculo é o que garante que nenhum caminho de edição
     * deixe um valor velho na tela: quem esquecer de recalcular teria uma
     * planilha que só mostra o resultado certo na próxima vez que for tocada.
     */
    updateSheet: (sheet) => {
      const { workbook } = get()
      if (workbook === null) return

      const sheets = [...workbook.sheets]
      sheets[workbook.activeSheet] = sheet
      set({ workbook: recalculate({ ...workbook, sheets }), isDirty: true })
    },

    changeStructure: (change: StructuralChange) => {
      const { workbook } = get()
      if (workbook === null) return

      const changed = applyStructuralChange(workbook, workbook.activeSheet, change)
      if (changed === workbook) return

      set({ workbook: recalculate(changed), isDirty: true })
    },

    selectSheet: (index) => {
      const { workbook } = get()
      if (workbook === null || index < 0 || index >= workbook.sheets.length) return
      // Trocar de aba não suja o arquivo: é navegação, não edição.
      set({ workbook: { ...workbook, activeSheet: index } })
    },

    addSheet: () => {
      const { workbook } = get()
      if (workbook === null) return

      const sheets = [...workbook.sheets, createSheet(nextSheetName(workbook))]
      set({ workbook: { sheets, activeSheet: sheets.length - 1 }, isDirty: true })
    },

    renameSheet: (index, name) => {
      const { workbook } = get()
      if (workbook === null) return

      const trimmed = name.trim()
      const sheet = workbook.sheets[index]
      if (sheet === undefined || trimmed.length === 0) return

      // Nome repetido quebraria a referência entre abas que o motor de fórmulas
      // da Fase 6 vai precisar — melhor recusar agora, em silêncio, do que
      // aceitar e falhar depois.
      const taken = workbook.sheets.some((other, at) => at !== index && other.name === trimmed)
      if (taken) return

      // Renomear reescreve as fórmulas que citam a aba: sem isso o gesto, que o
      // usuário considera cosmético, viraria #REF! em toda planilha que a usa.
      set({ workbook: recalculate(renameSheetIn(workbook, index, trimmed)), isDirty: true })
    },

    removeSheet: (index) => {
      const { workbook } = get()
      // Uma pasta sem planilha nenhuma não é estado válido do modelo.
      if (workbook === null || workbook.sheets.length <= 1) return

      const sheets = workbook.sheets.filter((_, at) => at !== index)
      const activeSheet = Math.min(workbook.activeSheet, sheets.length - 1)
      set({ workbook: { sheets, activeSheet }, isDirty: true })
    },

    openViaDialog: async () => {
      if (!(await ensureChangesHandled())) return
      await openPath(async () => {
        const data = await call(() => window.api.file.open({}))
        return data === null || data.canceled ? null : data.file
      })
    },

    openRecent: async (path) => {
      if (!(await ensureChangesHandled())) return
      await openPath(async () => {
        const data = await call(() => window.api.file.openRecent({ path }))
        if (data === null) {
          // O arquivo pode ter sumido; a lista precisa refletir isso.
          await get().refreshRecents()
          return null
        }
        return data.file
      })
    },

    save: async () => {
      const state = get()
      if (state.file === null) return false
      // Arquivo que nunca foi gravado não tem destino: vira "salvar como".
      if (state.file.path === null) return get().saveAs()

      // Planilha não passa pelo caminho de texto: não tem formatação a
      // perder para `.txt`, e o conteúdo é outro.
      const { workbook } = state
      if (workbook !== null) {
        const saved = await call(() =>
          window.api.file.save({
            path: state.file?.path ?? '',
            content: serializeWorkbook(workbook),
          }),
        )
        if (saved === null) return false

        set({ isDirty: false, file: { ...state.file, name: saved.name } })
        await forgetDraft()
        await get().refreshRecents()
        return true
      }

      const model = currentModel()
      let content: string

      if (isPlainTextPath(state.file.path)) {
        if (hasRichFormatting(model.doc)) {
          // Salvar em .txt descartaria formatação. Perguntar antes é a regra
          // do projeto: nada se perde em silêncio.
          const answer = await call(() =>
            window.api.dialog.confirmPlainText({ fileName: state.file?.name ?? 'documento' }),
          )
          if (answer === null || answer.choice === PlainTextChoice.Cancel) return false
          if (answer.choice === PlainTextChoice.SaveAsDocument) return get().saveAs()
        }
        content = documentToPlainText(model.doc)
      } else {
        content = serializeDocument(model)
      }

      const data = await call(() => window.api.file.save({ path: state.file?.path ?? '', content }))
      if (data === null) return false

      set({ isDirty: false, file: { ...state.file, name: data.name } })
      await forgetDraft()
      await get().refreshRecents()
      return true
    },

    saveAs: async () => {
      const state = get()
      if (state.file === null) return false

      const model = currentModel()
      const chosen = await call(() =>
        window.api.file.chooseSavePath({
          suggestedName: state.file?.name ?? defaultFileName(state.file?.kind ?? DocumentKind.Document),
          kind: state.file?.kind ?? DocumentKind.Document,
        }),
      )
      if (chosen === null || chosen.canceled) return false

      // O aviso vem antes da gravação: se o usuário desistir aqui, nenhum byte
      // foi escrito e o destino continua como estava.
      if (state.workbook === null && isPlainTextPath(chosen.path) && hasRichFormatting(model.doc)) {
        const answer = await call(() => window.api.dialog.confirmPlainText({ fileName: chosen.name }))
        if (answer === null || answer.choice === PlainTextChoice.Cancel) return false
        // Quis preservar a formatação: escolhe outro destino.
        if (answer.choice === PlainTextChoice.SaveAsDocument) return get().saveAs()
      }

      const content =
        state.workbook !== null
          ? serializeWorkbook(state.workbook)
          : isPlainTextPath(chosen.path)
            ? documentToPlainText(model.doc)
            : serializeDocument(model)

      const data = await call(() => window.api.file.save({ path: chosen.path, content }))
      if (data === null) return false

      set({ file: { path: chosen.path, name: data.name, kind: state.file.kind }, isDirty: false })
      await forgetDraft()
      await get().refreshRecents()
      return true
    },

    closeFile: async () => {
      if (!(await ensureChangesHandled())) return
      const empty = createEmptyDocument()
      set((state) => ({
        file: null,
        page: empty.page,
        initialDoc: empty.doc,
        generation: state.generation + 1,
        isDirty: false,
        error: null,
      }))
      await forgetDraft()
      await get().refreshRecents()
    },

    clearRecents: async () => {
      const data = await call(() => window.api.recent.clear({}))
      if (data !== null) set({ recents: data.files })
    },

    autosave: async () => {
      const state = get()
      // Sem arquivo aberto não há o que guardar; sem alteração, o rascunho já
      // está em dia. E enquanto houver rascunho esperando decisão, escrever
      // apagaria o trabalho que ele existe para devolver.
      if (state.file === null || !state.isDirty) return
      if (state.pendingDraft !== null || state.autosaveBroken) return

      // Fora do `call`: o autosave não pode piscar o indicador de ocupado nem
      // roubar a vez de uma operação que o usuário pediu.
      const result = await window.api.file.autosave({
        path: state.file.path,
        name: state.file.name,
        kind: state.file.kind,
        content: currentContent(),
      })

      if (!result.ok) {
        // Uma vez só: insistir a cada oito segundos encheria a tela, e falhar em
        // silêncio deixaria o usuário confiando numa proteção que não existe.
        set({ autosaveBroken: true, error: result.error })
      }
    },

    checkRecovery: async () => {
      const result = await window.api.recovery.peek({})
      if (result.ok && result.data.draft !== null) set({ pendingDraft: result.data.draft })
    },

    recoverDraft: async () => {
      const data = await call(() => window.api.recovery.restore({}))
      set({ pendingDraft: null })
      if (data === null || data.draft === null) return

      const { draft } = data
      try {
        if (draft.kind === DocumentKind.Spreadsheet) {
          const workbook = recalculate(parseWorkbook(draft.content))
          load({ path: draft.path, name: draft.name, kind: DocumentKind.Spreadsheet }, createEmptyDocument())
          set({ workbook, notice: null })
        } else {
          load(
            { path: draft.path, name: draft.name, kind: DocumentKind.Document },
            parseDocument(draft.content),
          )
          set({ workbook: null, notice: null })
        }

        // Recuperado é, por definição, diferente do que está no disco: marcar
        // como salvo faria o usuário fechar a janela e perder tudo de novo.
        set({ isDirty: true })
      } catch (cause) {
        set({ error: toSerialized(cause) })
      }
    },

    dismissDraft: async () => {
      set({ pendingDraft: null })
      await window.api.recovery.discard({})
    },

    exportPdf: async () => {
      const request = buildPrintRequest()
      if (request === null) return false

      const data = await call(() =>
        window.api.print.exportPdf({ ...request, suggestedName: get().file?.name ?? 'documento' }),
      )
      return data !== null && !data.canceled
    },

    print: async () => {
      const request = buildPrintRequest()
      if (request === null) return false

      const data = await call(() => window.api.print.dialog(request))
      return data !== null && data.printed
    },

    printPreview: async () => {
      const request = buildPrintRequest()
      if (request === null) return

      await call(() => window.api.print.preview({ ...request, title: get().file?.name ?? 'Documento' }))
    },
  }

  /** Reúne o que o processo main precisa para renderizar: HTML e página. */
  function buildPrintRequest(): { html: string; page: PageSetup } | null {
    const state = get()
    if (documentSource === null) return null

    return {
      html: buildPrintHtml(documentSource.readHtml(), state.file?.name ?? 'Documento'),
      page: state.page,
    }
  }
})
