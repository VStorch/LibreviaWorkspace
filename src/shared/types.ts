import { Language } from './i18n/language.js'

/** O que está sendo editado. Guia ícone, filtros de diálogo e editor. */
export const DocumentKind = {
  Document: 'document',
  Spreadsheet: 'spreadsheet',
} as const

export type DocumentKind = (typeof DocumentKind)[keyof typeof DocumentKind]

/**
 * O que o documento tem e o aplicativo não dá conta — em duas categorias que
 * **não** são o mesmo problema (ver docs/02-docx-cirurgico.md).
 *
 * `invisible`: continua no arquivo depois de salvar, mas não aparece na tela.
 * `lost`: some de verdade ao salvar.
 *
 * Misturar os dois produz um aviso genérico que o usuário aprende a ignorar, e
 * aí ele deixa de proteger de qualquer coisa.
 */
export interface LossInventory {
  // Coleções mutáveis, como em `DocumentNode`: este tipo precisa ser atribuível
  // ao que o zod infere no contrato de IPC, e um array `readonly` não é
  // atribuível a um comum. As propriedades continuam `readonly`.
  readonly invisible: string[]
  readonly lost: string[]
  /**
   * Subconjunto de `invisible`: o que some se o bloco que o ancora for editado.
   *
   * Comentário, revisão, nota e campo calculado entram aqui; posicionamento de
   * imagem e decoração, não. A diferença decide se o arquivo abre em somente
   * leitura — travar a edição por perda de aparência travaria o uso do dia a
   * dia, e o usuário aprenderia a liberar sem ler.
   */
  readonly structural: string[]
}

/**
 * Um arquivo carregado do disco.
 *
 * `content` é sempre texto: `.txt` vem cru e `.sdoc` vem como JSON. Um `.docx`
 * chega aqui **já convertido** para o formato interno pelo processo main, de
 * modo que o renderer segue com um caminho só. Os bytes originais ficam no
 * main, que é quem precisa deles para gravar cirurgicamente.
 */
export interface LoadedFile {
  readonly path: string
  readonly name: string
  readonly kind: DocumentKind
  readonly content: string
  /** Presente só quando o arquivo veio de um formato do Office. */
  readonly inventory?: LossInventory
}

/**
 * O rascunho de recuperação, sem o conteúdo.
 *
 * O aviso precisa dizer de que arquivo veio e de quando é; carregar junto o
 * conteúdo — que pode ter dezenas de megabytes com imagens embutidas — só para
 * decidir se mostra um aviso seria desperdício.
 */
export interface DraftSummary {
  readonly path: string | null
  readonly name: string
  readonly kind: DocumentKind
  readonly savedAt: number
}

export interface RecentFile {
  readonly path: string
  readonly name: string
  readonly kind: DocumentKind
  readonly openedAt: number
}

/** Comandos que o menu nativo despacha para o renderer. */
export const MenuCommand = {
  NewDocument: 'new-document',
  NewSpreadsheet: 'new-spreadsheet',
  Open: 'open',
  OpenRecent: 'open-recent',
  ClearRecent: 'clear-recent',
  Save: 'save',
  SaveAs: 'save-as',
  CloseFile: 'close-file',
  FindReplace: 'find-replace',
  ExportPdf: 'export-pdf',
  Print: 'print',
  PrintPreview: 'print-preview',
  PageSetup: 'page-setup',
  /** Abre o diálogo de parágrafo — espaçamento, entrelinha, recuo, alinhamento. */
  ParagraphSetup: 'paragraph-setup',
  InsertPageBreak: 'insert-page-break',
  /** Zoom da folha: o do editor, e não o do Chromium, que aumentaria a interface. */
  ZoomIn: 'zoom-in',
  ZoomOut: 'zoom-out',
  ZoomReset: 'zoom-reset',
  ZoomFitWidth: 'zoom-fit-width',
  /** Insere a área de transferência como texto, sem trazer formatação. */
  PasteWithoutFormat: 'paste-without-format',
  /** Abre o diálogo de contagem de palavras. */
  WordCount: 'word-count',
  /** Abre o seletor de caracteres especiais. */
  SpecialCharacter: 'special-character',
  /** Propriedades da imagem selecionada: texto alternativo e alinhamento. */
  ImageProperties: 'image-properties',
  /** Marcadores: adicionar, ir para e excluir. */
  InsertBookmark: 'insert-bookmark',
  /** Sumário dos títulos, com número de página e link. */
  InsertTableOfContents: 'insert-table-of-contents',
  /** Refaz as entradas do sumário a partir dos títulos de agora. */
  UpdateTableOfContents: 'update-table-of-contents',
  /** F9: recalcula os campos da seleção, ou do documento inteiro. */
  UpdateFields: 'update-fields',
  // O menu "Tabela". Os valores são os mesmos de `TableAction` (ver
  // `table-actions.ts`), e é por eles que o `App` os repassa ao editor.
  TableInsert: 'table-insert',
  TableRowBefore: 'table-row-before',
  TableRowAfter: 'table-row-after',
  TableDeleteRow: 'table-delete-row',
  TableColumnBefore: 'table-column-before',
  TableColumnAfter: 'table-column-after',
  TableDeleteColumn: 'table-delete-column',
  TableMergeCells: 'table-merge-cells',
  TableSplitCell: 'table-split-cell',
  TableHeaderRow: 'table-header-row',
  TableDelete: 'table-delete',
  TableProperties: 'table-properties',
  /** Emitido quando o usuário escolhe "Salvar" no aviso de saída. */
  SaveAndExit: 'save-and-exit',
} as const

export type MenuCommand = (typeof MenuCommand)[keyof typeof MenuCommand]

/** Resposta do aviso de alterações não salvas. */
export const DiscardChoice = {
  Save: 'save',
  Discard: 'discard',
  Cancel: 'cancel',
} as const

export type DiscardChoice = (typeof DiscardChoice)[keyof typeof DiscardChoice]

/** Resposta ao aviso de que `.txt` não guarda formatação. */
export const PlainTextChoice = {
  /** Salvar assim mesmo, aceitando a perda de formatação. */
  KeepPlain: 'keep-plain',
  /** Salvar como documento, preservando tudo. */
  SaveAsDocument: 'save-as-document',
  Cancel: 'cancel',
} as const

export type PlainTextChoice = (typeof PlainTextChoice)[keyof typeof PlainTextChoice]

/**
 * O tema da interface, como a pessoa o escolheu.
 *
 * Três valores e não dois: `system` é o padrão, e é o único que sabe a resposta
 * certa para quem troca de claro para escuro ao anoitecer. `light` e `dark`
 * são a escolha explícita, que o sistema não desfaz.
 *
 * O que a tela desenha é o tema **resolvido** — ver `ResolvedTheme`. Quem
 * resolve é o main, porque é ele que enxerga o `nativeTheme` do Chromium.
 */
export const Theme = {
  System: 'system',
  Light: 'light',
  Dark: 'dark',
} as const

export type Theme = (typeof Theme)[keyof typeof Theme]

/** O tema depois de `system` virar um dos dois de verdade. */
export type ResolvedTheme = 'light' | 'dark'

/**
 * Preferências de edição e de aparência.
 *
 * Moram no processo main porque três delas não podem morar em outro lugar: a
 * ortografia é configuração de `session`, o idioma monta a barra de menus
 * nativa, e o tema precisa do `nativeTheme` para resolver `system`. Guardar
 * duas metades da mesma preferência em dois lugares faria o menu marcar o que
 * o editor não estava fazendo.
 *
 * Um único conjunto, e não um "editor" e um "aparência" separados: são o mesmo
 * arquivo, o mesmo canal de IPC e o mesmo aviso de mudança. Dois canais seriam
 * duas chances de a tela e o menu discordarem.
 */
export interface EditorPreferences {
  /** Verificação ortográfica em português, no corpo e nas faixas. */
  readonly spellcheck: boolean
  /** Marcas de formatação: ¶, espaço, tabulação e quebra de linha. */
  readonly invisibleCharacters: boolean
  /** Autocorreção tipográfica: aspas curvas, travessão, reticências. */
  readonly typography: boolean
  /**
   * O idioma da interface.
   *
   * Não muda as fórmulas, que continuam aceitando os dois idiomas como sempre,
   * nem o dicionário do corretor. Quem escreve em português numa interface em
   * inglês é caso comum, e amarrar as três coisas obrigaria a escolher qual
   * delas sacrificar.
   */
  readonly language: Language
  /** A escolha da pessoa; `system` deixa o sistema operacional decidir. */
  readonly theme: Theme
  /**
   * Modo de leitura: sem barras, em rolagem contínua e sem edição.
   *
   * Guardado como preferência, e não como estado da sessão, porque o item do
   * menu nativo precisa mostrar a marca — e o menu mora no main, que só sabe o
   * que está aqui.
   */
  readonly readingMode: boolean
  readonly showToolbar: boolean
  readonly showStatusBar: boolean
  /** Zoom da folha na tela, em porcento (50–200). Não muda a paginação. */
  readonly zoom: number
  /** Ajustar à largura: o zoom acompanha a janela, e `zoom` fica como estava. */
  readonly zoomFit: boolean
  /**
   * O painel de navegação: os títulos do documento, à esquerda da folha.
   *
   * Preferência, e não estado do documento, pelo mesmo motivo do modo de leitura:
   * o item do menu "Exibir" mostra a marca, e o menu só sabe o que está aqui. E é
   * assim no Word — quem abre o painel num documento o encontra aberto no próximo.
   */
  readonly navigationPane: boolean
}

/**
 * Remendo de preferências: o que o canal `prefs:set` aceita.
 *
 * Um tipo próprio porque `Partial` não basta com `exactOptionalPropertyTypes`: o
 * que o zod infere de um schema parcial admite a chave presente com `undefined`,
 * e é exatamente esse valor que atravessa o IPC.
 */
export type EditorPreferencesPatch = {
  readonly [K in keyof EditorPreferences]?: EditorPreferences[K] | undefined
}

/**
 * O estado em que o aplicativo abre.
 *
 * Ortografia e tipografia ligadas, porque é o que um editor de texto em
 * português faz de útil sem ninguém pedir. Marcas de formatação desligadas: elas
 * são ferramenta de conferência, e deixá-las ligadas sujaria a tela de quem só
 * quer escrever. Modo de leitura desligado: é para ler o que já existe, e o
 * aplicativo abre para escrever.
 *
 * O `language` aqui é só o valor de última instância. Na primeira execução quem
 * decide é o sistema operacional — ver `load()` em `src/main/preferences.ts`,
 * que é o único lugar que enxerga `app.getLocale()`. Este módulo é `shared` e
 * não pode importar `electron`.
 */
export const DEFAULT_EDITOR_PREFERENCES: EditorPreferences = {
  spellcheck: true,
  invisibleCharacters: false,
  typography: true,
  language: Language.Portuguese,
  theme: Theme.System,
  readingMode: false,
  showToolbar: true,
  showStatusBar: true,
  zoom: 100,
  zoomFit: false,
  navigationPane: false,
}

/** Operações de área de transferência que só o `webContents` sabe fazer. */
export const EditCommand = {
  Cut: 'cut',
  Copy: 'copy',
  Paste: 'paste',
} as const

export type EditCommand = (typeof EditCommand)[keyof typeof EditCommand]

/**
 * O que estava debaixo do botão direito, como o Chromium o descreve.
 *
 * Vem do evento `context-menu` do `webContents`: é o **único** lugar onde o
 * corretor ortográfico do Chromium conta o que ele achou errado e o que sugere
 * no lugar. O renderer desenha o menu; o main é quem tem esses dados.
 */
export interface ContextMenuTarget {
  /** Onde clicou, em pixels da janela. */
  readonly x: number
  readonly y: number
  /** O clique caiu em algo editável — fora disso, colar não faz sentido. */
  readonly editable: boolean
  /** Vazio quando o clique não caiu sobre palavra marcada como errada. */
  readonly misspelledWord: string
  /** As sugestões do corretor, na ordem em que ele as deu. */
  readonly dictionarySuggestions: string[]
  readonly canCut: boolean
  readonly canCopy: boolean
  readonly canPaste: boolean
}

/**
 * Até onde vai o "não marque mais esta palavra".
 *
 * `permanent` grava no dicionário do usuário, que sobrevive a fechar o
 * aplicativo. `session` é o "ignorar": vale enquanto a janela estiver aberta e
 * é desfeito na saída — o Chromium não tem lista de ignorados, então ela é
 * imitada com uma entrada temporária no dicionário.
 */
export const DictionaryScope = {
  Permanent: 'permanent',
  Session: 'session',
} as const

export type DictionaryScope = (typeof DictionaryScope)[keyof typeof DictionaryScope]
