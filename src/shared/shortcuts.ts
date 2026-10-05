/**
 * The single table of keyboard shortcuts.
 *
 * Menu accelerators are registered in main and **intercept the key before the renderer sees it**.
 * With two lists (the editor keymap and the menu accelerators), every overlap would silently kill
 * the editor shortcut.
 *
 * So each key is declared once, saying **who handles it** and **what it does**; the menu and the
 * editor read from here, and the test next door fails if two entries claim the same key for
 * different owners.
 *
 * Some keys are not registered by us but by Tiptap extensions (`Ctrl+B`, headings on
 * `Ctrl+Alt+1`…`6`, indent…). They are here on purpose, marked with `registeredBy`, to **reserve**
 * the key: a future menu item on `Ctrl+B` hits the test, not a report months later that bold
 * "sometimes doesn't work".
 *
 * Left out: Electron roles that keep their default accelerator (undo, copy, paste, select all,
 * quit, full screen). Chromium handles them inside the editable field, so declaring `Ctrl+Z` as
 * "menu" would invent a collision. Only roles whose accelerator we change (`reload`) and zoom are
 * listed.
 *
 * Also left out: alignment on `Ctrl+Shift+L/E/R/J`, which `TextAlign` adds by default. In
 * development "Reload" covers its `Ctrl+Shift+R`; the overlap is real and harmless, since the UI
 * announces Word's `Ctrl+R` for right alignment, but declaring it would make the collision test
 * flag something we do not want to change now.
 */

/** Whoever handles the key takes it from the other. */
export const ShortcutOwner = {
  Editor: 'editor',
  Menu: 'menu',
} as const

export type ShortcutOwner = (typeof ShortcutOwner)[keyof typeof ShortcutOwner]

/**
 * The key combination in a neutral form, from which both spellings derive.
 *
 * Almost every shortcut uses the command modifier; the exceptions are function keys Word uses
 * alone, like `F9` to update fields, and only for them may `mod` be missing. The key goes by
 * Electron's name (`B`, `=`, `[`, `F10`, `Enter`, `numadd`), never an alias: writing `Plus` would
 * mean `Shift+=` under another name, and the collision would slip through because the spellings do
 * not match.
 */
export interface ShortcutKey {
  /**
   * `Ctrl` on Windows and Linux, `Cmd` on macOS. Missing only on function keys and Word's `Alt+=`.
   */
  readonly mod?: true
  readonly shift?: true
  readonly alt?: true
  readonly key: string
}

export interface Shortcut {
  readonly owner: ShortcutOwner
  readonly key: ShortcutKey
  readonly does: string
  /**
   * Who registers the key when it is not us: the Tiptap extension or the Electron role. An entry
   * with this field only reserves the combination.
   */
  readonly registeredBy?: string
}

/**
 * The registry key is the id main and the editor ask for; changing a key means changing one line
 * here, and both ends follow.
 */
export const SHORTCUTS = {
  // ## Menu → File
  newDocument: { owner: ShortcutOwner.Menu, key: { mod: true, key: 'N' }, does: 'Novo documento' },
  /**
   * In Word `Ctrl+Shift+N` resets the paragraph to body text. Here the menu owns it, and the menu
   * would win anyway: the accelerator arrives first. Body text stays on `Ctrl+Alt+0` from the
   * `Paragraph` extension, next to the headings on `Ctrl+Alt+1`…`6`.
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

  // ## Menu → Edit
  /**
   * Chromium already handles `Ctrl+Shift+V` in an editable field, and not the way Word does. This
   * accelerator arrives first, so ours wins.
   */
  pasteWithoutFormat: {
    owner: ShortcutOwner.Menu,
    key: { mod: true, shift: true, key: 'V' },
    does: 'Colar sem formatação',
  },
  findReplace: { owner: ShortcutOwner.Menu, key: { mod: true, key: 'F' }, does: 'Localizar e substituir…' },

  // ## Menu → Insert
  insertPageBreak: { owner: ShortcutOwner.Menu, key: { mod: true, key: 'Enter' }, does: 'Quebra de página' },

  // ## Menu → Table
  /** Word's `Ctrl+Shift+F5` for the bookmark dialog. */
  insertBookmark: {
    owner: ShortcutOwner.Menu,
    key: { mod: true, shift: true, key: 'F5' },
    does: 'Marcador…',
  },
  /**
   * `Ctrl+F12` is LibreOffice's insert table. Word has no key for it, and the free `Ctrl` letters
   * ran out in this table.
   */
  insertTable: { owner: ShortcutOwner.Menu, key: { mod: true, key: 'F12' }, does: 'Inserir tabela…' },
  /** Word's keys for footnotes and endnotes. */
  insertFootnote: {
    owner: ShortcutOwner.Menu,
    key: { mod: true, alt: true, key: 'F' },
    does: 'Nota de rodapé',
  },
  insertEndnote: {
    owner: ShortcutOwner.Menu,
    key: { mod: true, alt: true, key: 'D' },
    does: 'Nota de fim',
  },
  /**
   * Word's `Alt+=` for inserting an equation. No `Ctrl`, like `F9`: it is the key Word users
   * already know, and no other entry uses it.
   */
  insertEquation: {
    owner: ShortcutOwner.Menu,
    key: { alt: true, key: '=' },
    does: 'Equação',
  },
  /** Word's and LibreOffice's key for inserting a comment. */
  insertComment: {
    owner: ShortcutOwner.Menu,
    key: { mod: true, alt: true, key: 'M' },
    does: 'Comentário',
  },

  // ## Menu → View
  /**
   * `Ctrl+F11`, next to `Ctrl+F10` for formatting marks: both toggle a way of viewing the document,
   * and sitting side by side makes the second easy to remember. `F11` alone is the system's full
   * screen.
   */
  readingMode: {
    owner: ShortcutOwner.Menu,
    key: { mod: true, key: 'F11' },
    does: 'Modo de leitura',
  },
  /**
   * `Ctrl+F10`, not Word's `Ctrl+*`: `Ctrl+Shift+8` **is** `Ctrl+*`, and it is also the bullet list
   * below. The new item moves, and `Ctrl+F10` is what LibreOffice uses for this.
   */
  formattingMarks: {
    owner: ShortcutOwner.Menu,
    key: { mod: true, key: 'F10' },
    does: 'Marcas de formatação',
  },
  /**
   * `F9` alone, as in Word: updates the fields in the selection, or in the whole document when the
   * cursor is collapsed.
   */
  updateFields: {
    owner: ShortcutOwner.Menu,
    key: { key: 'F9' },
    does: 'Atualizar campos',
  },
  /**
   * `Ctrl+F5`: `F5` is LibreOffice's Navigator, and every key of ours carries `Ctrl`. Word's
   * `Ctrl+F` opens the pane through search, and here it is already find and replace.
   */
  navigationPane: {
    owner: ShortcutOwner.Menu,
    key: { mod: true, key: 'F5' },
    does: 'Painel de navegação',
  },
  /**
   * Zoom in leaves `Ctrl+Shift+=` for a reason. The default accelerator of the `zoomIn` role is
   * `CommandOrControl+Plus`, which in Electron is the `=` key **with Shift**, the same that toggles
   * superscript. In a text editor formatting beats zoom, so zoom moves to the keypad `+`, which
   * competes with nothing. Zoom out keeps the default: `Ctrl+-` collides with nothing.
   */
  zoomIn: { owner: ShortcutOwner.Menu, key: { mod: true, key: 'numadd' }, does: 'Ampliar' },
  /** Not Electron roles: the zoom belongs to the sheet. */
  zoomOut: { owner: ShortcutOwner.Menu, key: { mod: true, key: '-' }, does: 'Reduzir' },
  zoomReset: { owner: ShortcutOwner.Menu, key: { mod: true, key: '0' }, does: 'Zoom 100 %' },
  /**
   * `Ctrl+Shift+R`, not `Ctrl+R`: the `reload` role default would swallow `Ctrl+R` for right
   * alignment, and the shortcut would look broken only on developers' machines, since the item
   * exists only in development.
   */
  reload: { owner: ShortcutOwner.Menu, key: { mod: true, shift: true, key: 'R' }, does: 'Recarregar' },

  // ## Menu → Review
  /**
   * Word's shortcut. `TextAlign` gives `Ctrl+Shift+E` to centering, but the accelerator arrives
   * first, and centering stays on Word's `Ctrl+E`.
   */
  trackChanges: {
    owner: ShortcutOwner.Menu,
    key: { mod: true, shift: true, key: 'E' },
    does: 'Controlar alterações',
  },

  // ## Menu → Tools
  wordCount: {
    owner: ShortcutOwner.Menu,
    key: { mod: true, shift: true, key: 'G' },
    does: 'Contar palavras…',
  },

  // ## Editor: alignment, as in Word and Writer
  alignLeft: { owner: ShortcutOwner.Editor, key: { mod: true, key: 'L' }, does: 'Alinhar à esquerda' },
  /**
   * `Ctrl+E` is Word's and also Tiptap's code mark, which this toolbar does not offer. Centering
   * wins through the `WordShortcuts` extension `priority`; the code mark stays reachable through
   * the backtick input rule.
   */
  alignCenter: { owner: ShortcutOwner.Editor, key: { mod: true, key: 'E' }, does: 'Centralizar' },
  alignRight: { owner: ShortcutOwner.Editor, key: { mod: true, key: 'R' }, does: 'Alinhar à direita' },
  alignJustify: { owner: ShortcutOwner.Editor, key: { mod: true, key: 'J' }, does: 'Justificar' },

  /**
   * Line spacing: `Ctrl+1` single, `Ctrl+5` one and a half, `Ctrl+2` double. They are Word's, which
   * is why headings sit on `Ctrl+Alt+1`…`6`. The measure is in **lines**, and `paragraph-format`
   * translates it to CSS, which depends on the font's natural height.
   */
  lineHeightSingle: { owner: ShortcutOwner.Editor, key: { mod: true, key: '1' }, does: 'Entrelinha simples' },
  lineHeightOneAndHalf: {
    owner: ShortcutOwner.Editor,
    key: { mod: true, key: '5' },
    does: 'Entrelinha de um e meio',
  },
  lineHeightDouble: { owner: ShortcutOwner.Editor, key: { mod: true, key: '2' }, does: 'Entrelinha dupla' },

  /**
   * Superscript and subscript, as in Word. The `=` goes by key code, not character: with Shift the
   * browser reports `+`, and `prosemirror-keymap` undoes that by trying the name derived from
   * `keyCode`.
   */
  superscript: {
    owner: ShortcutOwner.Editor,
    key: { mod: true, shift: true, key: '=' },
    does: 'Sobrescrito',
  },
  subscript: { owner: ShortcutOwner.Editor, key: { mod: true, key: '=' }, does: 'Subscrito' },

  // ## Editor: reserved, registered by an extension, not by us
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
  /** For keyboards where `=` is not a single key. */
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
 * The shortcuts **our** keymap registers: editor-owned and without `registeredBy`. The type forces
 * the map in `word-shortcuts` to handle all of them: adding an entry here without a command does
 * not compile.
 */
export type EditorShortcutId = {
  [Id in ShortcutId]: (typeof SHORTCUTS)[Id] extends { owner: 'editor'; registeredBy?: undefined }
    ? Id
    : never
}[ShortcutId]

function parts(key: ShortcutKey): {
  readonly mod: boolean
  readonly modifiers: readonly string[]
  readonly key: string
} {
  const modifiers: string[] = []
  if (key.shift === true) modifiers.push('Shift')
  if (key.alt === true) modifiers.push('Alt')
  return { mod: key.mod === true, modifiers, key: key.key }
}

export function acceleratorOf(shortcut: Shortcut): string {
  const { mod, modifiers, key } = parts(shortcut.key)
  return [...(mod ? ['CmdOrCtrl'] : []), ...modifiers, key].join('+')
}

/**
 * Lowercase on purpose: to the keymap `L` is the key that needs Shift, and the shortcut would never
 * fire.
 */
export function editorKeyOf(shortcut: Shortcut): string {
  const { mod, modifiers, key } = parts(shortcut.key)
  return [...(mod ? ['Mod'] : []), ...modifiers, key.length === 1 ? key.toLowerCase() : key].join('-')
}

/** Says `Ctrl` on every system, as it always has: it is the name users of this app read. */
export function shortcutHintOf(shortcut: Shortcut): string {
  const { mod, modifiers, key } = parts(shortcut.key)
  return [...(mod ? ['Ctrl'] : []), ...modifiers, key].join('+')
}

/** The collision test uses it to find two entries claiming the same key. */
export function canonicalKeyOf(key: ShortcutKey): string {
  const { mod, modifiers, key: name } = parts(key)
  return [...(mod ? ['Mod'] : []), ...modifiers, name.toLowerCase()].join('+')
}
