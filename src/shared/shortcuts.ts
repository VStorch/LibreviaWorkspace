/**
 * A tabela única dos atalhos de teclado.
 *
 * Antes daqui havia duas listas que não se conheciam: o mapa de teclado do
 * editor, no renderer, e os aceleradores do menu nativo, declarados um a um no
 * main. Como acelerador de menu é registrado no processo main e **intercepta a
 * tecla antes de o renderer vê-la**, toda vez que as duas listas se cruzavam o
 * atalho do editor simplesmente parava de existir — sem erro, sem aviso. Foram
 * três colisões descobertas assim, uma de cada vez, e a última ("Ampliar" comendo
 * o `Ctrl+Shift+=` do sobrescrito) foi a que custou a troca de tecla.
 *
 * Agora cada tecla é declarada uma vez, dizendo **quem a atende** e **o que ela
 * faz**; o menu e o editor leem daqui. Não é documentação: é a origem das duas
 * formas — o acelerador do Electron e a chave do `prosemirror-keymap` — e o teste
 * ao lado falha se duas entradas pedirem a mesma tecla para donos diferentes.
 *
 * ## Entradas reservadas
 *
 * Algumas teclas não são registradas por nós: vêm de extensões do Tiptap
 * (`Ctrl+B`, os títulos em `Ctrl+Alt+1`…`6`, o recuo…). Elas estão aqui de
 * propósito, marcadas em `registeredBy`, para **ocupar** a tecla: quem amanhã
 * quiser um item de menu em `Ctrl+B` descobre o choque no teste, e não meses
 * depois, pelo relato de que o negrito "às vezes não pega".
 *
 * ## O que não está aqui
 *
 * Os papéis prontos do Electron que mantêm o acelerador padrão — desfazer,
 * copiar, colar, selecionar tudo, sair, tela cheia. São
 * atendidos pelo próprio Chromium dentro do campo de edição, então declarar
 * `Ctrl+Z` como "do menu" inventaria uma colisão que não existe. Entram apenas
 * os papéis cujo acelerador nós trocamos (`reload`) e o zoom, que é nosso.
 *
 * Fora também o alinhamento em `Ctrl+Shift+L/E/R/J`, que a extensão `TextAlign`
 * dá por padrão: em desenvolvimento o item "Recarregar" cobre o `Ctrl+Shift+R`
 * dessa lista. É sobreposição de verdade, e inofensiva — o atalho que a interface
 * anuncia para alinhar à direita é o `Ctrl+R` do Word, atendido aqui —, mas
 * declará-la faria o teste de colisão acusar algo que não se quer mudar agora.
 * Mudar a tecla de um dos dois é decisão de produto, não de refatoração.
 */

/** Quem atende a tecla — e, portanto, quem a tira do outro. */
export const ShortcutOwner = {
  /** O mapa de teclado do editor, no renderer. */
  Editor: 'editor',
  /** O acelerador de item de menu, registrado no processo main. */
  Menu: 'menu',
} as const

export type ShortcutOwner = (typeof ShortcutOwner)[keyof typeof ShortcutOwner]

/**
 * A combinação em forma neutra, de onde saem as duas escritas.
 *
 * Todo atalho nosso usa o modificador de comando, então ele é obrigatório: o que
 * varia é o resto. A tecla vai pelo nome que o Electron usa (`B`, `=`, `[`,
 * `F10`, `Enter`, `numadd`), e nunca por um apelido: escrever `Plus` seria dizer
 * `Shift+=` com outro nome, e a colisão que se quer impossível voltaria a passar
 * em silêncio porque as duas grafias não se parecem.
 */
export interface ShortcutKey {
  /** `Ctrl` no Windows e no Linux, `Cmd` no macOS. */
  readonly mod: true
  readonly shift?: true
  readonly alt?: true
  readonly key: string
}

export interface Shortcut {
  readonly owner: ShortcutOwner
  readonly key: ShortcutKey
  /** O que a tecla faz, na língua da interface. */
  readonly does: string
  /**
   * Quem registra a tecla, quando não somos nós — a extensão do Tiptap ou o papel
   * do Electron. Entrada com este campo está aqui só para reservar a combinação.
   */
  readonly registeredBy?: string
}

/**
 * Cada atalho do aplicativo, uma vez.
 *
 * A chave do registro é o identificador com que o main e o editor pedem o atalho;
 * mudar uma tecla é mudar uma linha daqui, e as duas pontas acompanham.
 */
export const SHORTCUTS = {
  // ## Menu → Arquivo
  newDocument: { owner: ShortcutOwner.Menu, key: { mod: true, key: 'N' }, does: 'Novo documento' },
  /**
   * No Word `Ctrl+Shift+N` volta o parágrafo para corpo de texto. Aqui é do menu,
   * e o menu ganha porque ganharia de todo jeito: o acelerador chega primeiro. O
   * corpo de texto segue no `Ctrl+Alt+0` da extensão `Paragraph`, vizinho natural
   * do `Ctrl+Alt+1`…`6` dos títulos.
   */
  newSpreadsheet: {
    owner: ShortcutOwner.Menu,
    key: { mod: true, shift: true, key: 'N' },
    does: 'Nova planilha',
  },
  open: { owner: ShortcutOwner.Menu, key: { mod: true, key: 'O' }, does: 'Abrir…' },
  save: { owner: ShortcutOwner.Menu, key: { mod: true, key: 'S' }, does: 'Salvar' },
  saveAs: { owner: ShortcutOwner.Menu, key: { mod: true, shift: true, key: 'S' }, does: 'Salvar como…' },
  print: { owner: ShortcutOwner.Menu, key: { mod: true, key: 'P' }, does: 'Imprimir…' },
  closeFile: { owner: ShortcutOwner.Menu, key: { mod: true, key: 'W' }, does: 'Fechar arquivo' },

  // ## Menu → Editar
  /**
   * O Chromium já responde a `Ctrl+Shift+V` dentro de um campo editável, e o que
   * ele faz não é o que o Word faz. O acelerador daqui chega primeiro, então passa
   * a valer o nosso.
   */
  pasteWithoutFormat: {
    owner: ShortcutOwner.Menu,
    key: { mod: true, shift: true, key: 'V' },
    does: 'Colar sem formatação',
  },
  findReplace: { owner: ShortcutOwner.Menu, key: { mod: true, key: 'F' }, does: 'Localizar e substituir…' },

  // ## Menu → Inserir
  insertPageBreak: { owner: ShortcutOwner.Menu, key: { mod: true, key: 'Enter' }, does: 'Quebra de página' },

  // ## Menu → Tabela
  /**
   * `Ctrl+F12` é o do LibreOffice para inserir tabela. O Word não tem tecla para
   * isto, e as letras livres com `Ctrl` já acabaram nesta tabela.
   */
  insertTable: { owner: ShortcutOwner.Menu, key: { mod: true, key: 'F12' }, does: 'Inserir tabela…' },

  // ## Menu → Exibir
  /**
   * `Ctrl+F11`, vizinho do `Ctrl+F10` das marcas de formatação.
   *
   * As duas são chaves do mesmo tipo — ligam e desligam um jeito de ver o
   * documento — e ficar uma ao lado da outra é o que faz a segunda ser
   * lembrada por quem já sabe a primeira. `F11` sozinho é a tela cheia do
   * sistema, e não se mexe nele.
   */
  readingMode: {
    owner: ShortcutOwner.Menu,
    key: { mod: true, key: 'F11' },
    does: 'Modo de leitura',
  },
  /**
   * `Ctrl+F10`, e não o `Ctrl+*` do Word: `Ctrl+Shift+8` **é** o `Ctrl+*`, e é
   * também a lista com marcadores logo abaixo nesta tabela. Quem se muda é o item
   * novo, e `Ctrl+F10` é o que o LibreOffice usa para isto.
   */
  formattingMarks: {
    owner: ShortcutOwner.Menu,
    key: { mod: true, key: 'F10' },
    does: 'Marcas de formatação',
  },
  /**
   * Ampliar sai do `Ctrl+Shift+=`, e não por capricho.
   *
   * O acelerador padrão do papel `zoomIn` é `CommandOrControl+Plus`, e no Electron
   * "Plus" é a tecla do `=` **com Shift** — a mesma que liga o sobrescrito. Num
   * editor de texto a formatação vem antes do zoom, então quem se muda é o zoom,
   * para o `+` do teclado numérico, que não disputa com ninguém. Reduzir fica no
   * padrão: `Ctrl+-` não colide com nada.
   */
  zoomIn: { owner: ShortcutOwner.Menu, key: { mod: true, key: 'numadd' }, does: 'Ampliar' },
  /** Os dois deixaram de ser papéis do Electron: o zoom agora é da folha. */
  zoomOut: { owner: ShortcutOwner.Menu, key: { mod: true, key: '-' }, does: 'Reduzir' },
  zoomReset: { owner: ShortcutOwner.Menu, key: { mod: true, key: '0' }, does: 'Zoom 100 %' },
  /**
   * `Ctrl+Shift+R` e não `Ctrl+R`: o padrão do papel `reload` engoliria o `Ctrl+R`
   * de "alinhar à direita", e o atalho pareceria quebrado só na máquina de quem
   * programa — o item só existe em desenvolvimento.
   */
  reload: { owner: ShortcutOwner.Menu, key: { mod: true, shift: true, key: 'R' }, does: 'Recarregar' },

  // ## Menu → Ferramentas
  /** O mesmo atalho do Word. */
  wordCount: {
    owner: ShortcutOwner.Menu,
    key: { mod: true, shift: true, key: 'G' },
    does: 'Contar palavras…',
  },

  // ## Editor: alinhamento, como no Word e no Writer
  alignLeft: { owner: ShortcutOwner.Editor, key: { mod: true, key: 'L' }, does: 'Alinhar à esquerda' },
  /**
   * `Ctrl+E` é do Word e da marca de código do Tiptap, que esta barra nem oferece.
   * Centralizar ganha pela `priority` da extensão `WordShortcuts`; a marca de
   * código continua alcançável pela regra de entrada de crase.
   */
  alignCenter: { owner: ShortcutOwner.Editor, key: { mod: true, key: 'E' }, does: 'Centralizar' },
  alignRight: { owner: ShortcutOwner.Editor, key: { mod: true, key: 'R' }, does: 'Alinhar à direita' },
  alignJustify: { owner: ShortcutOwner.Editor, key: { mod: true, key: 'J' }, does: 'Justificar' },

  /**
   * Entrelinha: `Ctrl+1` simples, `Ctrl+5` um e meio, `Ctrl+2` duplo.
   *
   * São os do Word, e é por isso que os títulos ficam no `Ctrl+Alt+1`…`6`: no Word
   * `Ctrl+1` nunca foi "Título 1". A medida é dita em **linhas**, e quem traduz
   * para a do CSS — que depende da altura natural da fonte — é `paragraph-format`.
   */
  lineHeightSingle: { owner: ShortcutOwner.Editor, key: { mod: true, key: '1' }, does: 'Entrelinha simples' },
  lineHeightOneAndHalf: {
    owner: ShortcutOwner.Editor,
    key: { mod: true, key: '5' },
    does: 'Entrelinha de um e meio',
  },
  lineHeightDouble: { owner: ShortcutOwner.Editor, key: { mod: true, key: '2' }, does: 'Entrelinha dupla' },

  /**
   * Sobrescrito e subscrito, como no Word.
   *
   * O `=` sai pelo código da tecla, e não pelo caractere: com Shift o navegador
   * informa `+`, e é o `prosemirror-keymap` que desfaz isso ao tentar o nome
   * derivado do `keyCode`.
   */
  superscript: {
    owner: ShortcutOwner.Editor,
    key: { mod: true, shift: true, key: '=' },
    does: 'Sobrescrito',
  },
  subscript: { owner: ShortcutOwner.Editor, key: { mod: true, key: '=' }, does: 'Subscrito' },

  // ## Editor: reservados — quem registra é uma extensão, não nós
  bold: {
    owner: ShortcutOwner.Editor,
    key: { mod: true, key: 'B' },
    does: 'Negrito',
    registeredBy: 'StarterKit',
  },
  italic: {
    owner: ShortcutOwner.Editor,
    key: { mod: true, key: 'I' },
    does: 'Itálico',
    registeredBy: 'StarterKit',
  },
  underline: {
    owner: ShortcutOwner.Editor,
    key: { mod: true, key: 'U' },
    does: 'Sublinhado',
    registeredBy: 'StarterKit',
  },
  bulletList: {
    owner: ShortcutOwner.Editor,
    key: { mod: true, shift: true, key: '8' },
    does: 'Lista com marcadores',
    registeredBy: 'StarterKit',
  },
  orderedList: {
    owner: ShortcutOwner.Editor,
    key: { mod: true, shift: true, key: '7' },
    does: 'Lista numerada',
    registeredBy: 'StarterKit',
  },
  bodyText: {
    owner: ShortcutOwner.Editor,
    key: { mod: true, alt: true, key: '0' },
    does: 'Corpo de texto',
    registeredBy: 'Paragraph',
  },
  heading1: {
    owner: ShortcutOwner.Editor,
    key: { mod: true, alt: true, key: '1' },
    does: 'Título 1',
    registeredBy: 'Heading',
  },
  heading2: {
    owner: ShortcutOwner.Editor,
    key: { mod: true, alt: true, key: '2' },
    does: 'Título 2',
    registeredBy: 'Heading',
  },
  heading3: {
    owner: ShortcutOwner.Editor,
    key: { mod: true, alt: true, key: '3' },
    does: 'Título 3',
    registeredBy: 'Heading',
  },
  heading4: {
    owner: ShortcutOwner.Editor,
    key: { mod: true, alt: true, key: '4' },
    does: 'Título 4',
    registeredBy: 'Heading',
  },
  heading5: {
    owner: ShortcutOwner.Editor,
    key: { mod: true, alt: true, key: '5' },
    does: 'Título 5',
    registeredBy: 'Heading',
  },
  heading6: {
    owner: ShortcutOwner.Editor,
    key: { mod: true, alt: true, key: '6' },
    does: 'Título 6',
    registeredBy: 'Heading',
  },
  outdent: {
    owner: ShortcutOwner.Editor,
    key: { mod: true, key: '[' },
    does: 'Diminuir recuo',
    registeredBy: 'Indent',
  },
  indent: {
    owner: ShortcutOwner.Editor,
    key: { mod: true, key: ']' },
    does: 'Aumentar recuo',
    registeredBy: 'Indent',
  },
  /** A saída para o teclado em que o `=` não é uma tecla só. */
  superscriptAlternate: {
    owner: ShortcutOwner.Editor,
    key: { mod: true, key: '.' },
    does: 'Sobrescrito',
    registeredBy: 'Superscript',
  },
  subscriptAlternate: {
    owner: ShortcutOwner.Editor,
    key: { mod: true, key: ',' },
    does: 'Subscrito',
    registeredBy: 'Subscript',
  },
} as const satisfies Record<string, Shortcut>

export type ShortcutId = keyof typeof SHORTCUTS

/**
 * Os atalhos que o **nosso** mapa de teclado registra: do editor e sem
 * `registeredBy`. O tipo existe para que o mapa em `word-shortcuts` tenha de
 * tratar todos — acrescentar uma entrada aqui sem dar-lhe comando não compila.
 */
export type EditorShortcutId = {
  [Id in ShortcutId]: (typeof SHORTCUTS)[Id] extends { owner: 'editor'; registeredBy?: undefined }
    ? Id
    : never
}[ShortcutId]

function parts(key: ShortcutKey): { readonly modifiers: readonly string[]; readonly key: string } {
  const modifiers: string[] = []
  if (key.shift === true) modifiers.push('Shift')
  if (key.alt === true) modifiers.push('Alt')
  return { modifiers, key: key.key }
}

/** A tecla como o Electron a quer, para `accelerator` de item de menu. */
export function acceleratorOf(shortcut: Shortcut): string {
  const { modifiers, key } = parts(shortcut.key)
  return ['CmdOrCtrl', ...modifiers, key].join('+')
}

/**
 * A tecla como o `prosemirror-keymap` a quer. Letra minúscula de propósito: para
 * ele `L` é a tecla que só sai com Shift, e o atalho nunca dispararia.
 */
export function editorKeyOf(shortcut: Shortcut): string {
  const { modifiers, key } = parts(shortcut.key)
  return ['Mod', ...modifiers, key.length === 1 ? key.toLowerCase() : key].join('-')
}

/**
 * A tecla como a barra de ferramentas a anuncia na dica do botão. Diz `Ctrl` em
 * todo sistema, como sempre disse: é o nome que o usuário deste aplicativo lê.
 */
export function shortcutHintOf(shortcut: Shortcut): string {
  const { modifiers, key } = parts(shortcut.key)
  return ['Ctrl', ...modifiers, key].join('+')
}

/**
 * A identidade da combinação, para comparar declarações entre si. É por ela que o
 * teste de colisão descobre que duas entradas pedem a mesma tecla.
 */
export function canonicalKeyOf(key: ShortcutKey): string {
  const { modifiers, key: name } = parts(key)
  return ['Mod', ...modifiers, name.toLowerCase()].join('+')
}
