import { Extension, Mark, mergeAttributes } from '@tiptap/core'
import type { Mark as ProseMirrorMark, MarkType, Node as ProseMirrorNode } from '@tiptap/pm/model'
import { closeHistory } from '@tiptap/pm/history'
import type { Transaction } from '@tiptap/pm/state'
import { blockStyleOfNode } from '@services/document/style-cascade.js'
import { headingLevelOfStyle, nextStyleIdOf } from '@services/document/style-editing.js'
import type { StyleCharacterFormat, StyleSheet } from '@services/document/styles.js'

/**
 * Aplicar estilos, limpar a formatação direta e o "desligado" que vence o estilo.
 *
 * Tudo em transação do editor, e por isso com desfazer. Modificar e criar estilo
 * não passa por aqui: mexe na folha de estilos do store (`style-editing.ts`), que
 * está fora do documento.
 *
 * As funções de transação ficam soltas, e não dentro da extensão, para que o teste
 * as rode sobre um `EditorState` sem montar editor nenhum — o mesmo motivo de
 * `paragraph-format.ts`.
 */

/** A formatação direta de parágrafo que o bloco carrega — o que "limpar" apaga. */
export const DIRECT_BLOCK_ATTRS = [
  'textAlign',
  'indentMm',
  'indentRightMm',
  'firstLineMm',
  'spaceBefore',
  'spaceAfter',
  'lineHeight',
  'background',
  'keepNext',
  'fontFamily',
  'fontSize',
] as const

/**
 * As marcas que um estilo pode ligar, e o campo do estilo que liga cada uma.
 *
 * São as que ganham o "desligado" (`off`): com o bloco desenhado pelo estilo, o
 * trecho sem marca mostra o que o estilo diz, e tirar o negrito de uma palavra num
 * título precisa de uma marca que diga o contrário. Na leitura ela vem de
 * `w:b w:val="0"` sobre estilo que liga; na gravação volta a ser isso.
 */
export const INHERITABLE_MARKS: Readonly<Record<string, keyof StyleCharacterFormat>> = {
  bold: 'bold',
  italic: 'italic',
  underline: 'underline',
  strike: 'strike',
}

/** Marcas que "limpar" deixa: o link é conteúdo, e o estilo de caractere é estilo. */
const KEPT_MARKS = new Set(['link', 'charStyle'])

const isStyledBlock = (node: ProseMirrorNode): boolean =>
  node.type.name === 'paragraph' || node.type.name === 'heading'

/**
 * O id que o bloco grava para o estilo: nenhum para o padrão de parágrafo — é o
 * que o leitor produz para o parágrafo sem `w:pStyle` —, o próprio para o resto.
 */
function storedIdOf(sheet: StyleSheet, styleId: string): string | null {
  return styleId === sheet.defaults.paragraphStyleId ? null : styleId
}

/** O estilo de parágrafo do bloco, por id — o título sem id, pelo nome `heading N`. */
export function styleIdOfBlock(sheet: StyleSheet, node: ProseMirrorNode): string | null {
  const declared = node.attrs['styleId']
  if (typeof declared === 'string' && declared !== '') return declared
  if (node.type.name === 'heading') {
    const name = `heading ${String(node.attrs['level'])}`
    return Object.values(sheet.styles).find((style) => style.name.toLowerCase() === name)?.id ?? null
  }
  return null
}

/**
 * Aplica um estilo de parágrafo aos blocos da seleção.
 *
 * Título ↔ parágrafo pelo nome do estilo (`heading N`), como o leitor e o escritor
 * reconhecem um título. A formatação direta do parágrafo sai, como no Word: aplicar
 * um estilo é pedir a aparência dele. A de caractere (as marcas) fica.
 */
export function applyParagraphStyle(tr: Transaction, sheet: StyleSheet, styleId: string): boolean {
  const style = sheet.styles[styleId]
  if (style === undefined) return false

  const level = headingLevelOfStyle(style)
  const schema = tr.doc.type.schema
  const type = level === null ? schema.nodes['paragraph'] : schema.nodes['heading']
  if (type === undefined) return false

  const { from, to } = tr.selection
  const targets: Array<{ pos: number; node: ProseMirrorNode }> = []
  tr.doc.nodesBetween(from, to, (node, pos) => {
    if (isStyledBlock(node)) targets.push({ pos, node })
    return !node.isTextblock
  })

  for (const { pos, node } of targets) {
    const attrs: Record<string, unknown> = { ...node.attrs, styleId: storedIdOf(sheet, styleId), indent: 0 }
    for (const name of DIRECT_BLOCK_ATTRS) attrs[name] = null
    if (level !== null) attrs['level'] = level
    tr.setNodeMarkup(pos, type, attrs)
  }

  return targets.length > 0
}

/** Aplica (ou tira, com `null`) um estilo de caractere ao trecho selecionado. */
export function applyCharacterStyle(tr: Transaction, styleId: string | null): boolean {
  const type = tr.doc.type.schema.marks['charStyle']
  if (type === undefined) return false

  const { from, to, empty } = tr.selection
  if (empty) {
    const marks = type.removeFromSet(tr.storedMarks ?? tr.selection.$from.marks())
    tr.setStoredMarks(styleId === null ? marks : type.create({ styleId }).addToSet(marks))
    return true
  }

  if (styleId === null) tr.removeMark(from, to, type)
  else tr.addMark(from, to, type.create({ styleId }))
  return true
}

/**
 * Limpa a formatação direta: os atributos de parágrafo do bloco e as marcas dos
 * trechos — o estilo fica, e é ele que passa a desenhar.
 *
 * Sem seleção, o bloco do cursor inteiro, como no Word.
 */
export function clearDirectFormatting(tr: Transaction): boolean {
  const { $from, empty } = tr.selection
  const from = empty ? $from.start() : tr.selection.from
  const to = empty ? $from.end() : tr.selection.to
  let changed = false

  tr.doc.nodesBetween(from, to, (node, pos) => {
    if (!isStyledBlock(node)) return true
    for (const name of [...DIRECT_BLOCK_ATTRS, 'indent'] as const) {
      const cleared = name === 'indent' ? 0 : null
      if (node.attrs[name] === undefined || node.attrs[name] === cleared) continue
      tr.setNodeAttribute(pos, name, cleared)
      changed = true
    }
    return false
  })

  for (const type of Object.values(tr.doc.type.schema.marks)) {
    if (KEPT_MARKS.has(type.name) || !tr.doc.rangeHasMark(from, to, type)) continue
    tr.removeMark(from, to, type)
    changed = true
  }

  tr.setStoredMarks([])
  return changed
}

/** O estilo do bloco do cursor liga esta marca? */
function inheritedOn(tr: Transaction, sheet: StyleSheet | null, name: string): boolean {
  const field = INHERITABLE_MARKS[name]
  if (field === undefined) return false
  return blockStyleOfNode(tr.selection.$from.parent, sheet)?.character[field] === true
}

/** As marcas da seleção de que se fala: as guardadas, as do cursor, ou as do trecho. */
function marksHere(tr: Transaction, type: MarkType): ProseMirrorMark | null {
  const { $from, empty, from, to } = tr.selection
  if (empty) return type.isInSet(tr.storedMarks ?? $from.marks()) ?? null

  let found: ProseMirrorMark | null = null
  let all = true
  tr.doc.nodesBetween(from, to, (node) => {
    if (!node.isText) return true
    const mark = type.isInSet(node.marks)
    if (mark === undefined) all = false
    else found ??= mark
    return false
  })
  return all ? found : null
}

/**
 * A marca aparece? Com o estilo por baixo: sem marca vale o estilo, a marca com
 * `off` desliga, e a marca comum liga.
 */
export function markVisiblyOn(tr: Transaction, sheet: StyleSheet | null, name: string): boolean {
  const type = tr.doc.type.schema.marks[name]
  if (type === undefined) return false
  const mark = marksHere(tr, type)
  if (mark !== null) return mark.attrs['off'] !== true
  return inheritedOn(tr, sheet, name)
}

/**
 * Liga ou desliga uma marca que o estilo pode dar.
 *
 * Onde o estilo não a liga, é a alternância de sempre (`false`: quem chama usa o
 * comando do próprio Tiptap). Onde liga, o trecho que aparece ligado ganha a marca
 * com `off`, e o que aparece desligado perde a marca — e volta ao estilo.
 */
export function toggleInheritedMark(tr: Transaction, sheet: StyleSheet | null, name: string): boolean {
  if (!inheritedOn(tr, sheet, name)) return false
  const type = tr.doc.type.schema.marks[name]
  if (type === undefined) return false

  const on = markVisiblyOn(tr, sheet, name)
  const { from, to, empty } = tr.selection
  if (empty) {
    const stored = type.removeFromSet(tr.storedMarks ?? tr.selection.$from.marks())
    tr.setStoredMarks(on ? type.create({ off: true }).addToSet(stored) : stored)
    return true
  }

  if (on) tr.addMark(from, to, type.create({ off: true }))
  else tr.removeMark(from, to, type)
  return true
}

/**
 * Enter no fim de um parágrafo cujo estilo tem `next`: o bloco novo nasce no
 * estilo seguinte — o título dá lugar ao Normal. Devolve falso quando não é o
 * caso, e o Enter de sempre segue.
 *
 * Só no bloco solto do corpo: o Enter de lista e de célula é deles.
 */
export function splitWithNextStyle(tr: Transaction, sheet: StyleSheet | null): boolean {
  const { $from, empty } = tr.selection
  if (sheet === null || !empty || $from.depth !== 1) return false

  const parent = $from.parent
  if (!isStyledBlock(parent) || $from.parentOffset !== parent.content.size) return false

  const own = styleIdOfBlock(sheet, parent)
  const next = nextStyleIdOf(sheet, own)
  if (next === null || next === (own ?? sheet.defaults.paragraphStyleId)) return false

  const schema = tr.doc.type.schema
  const level = headingLevelOfStyle(sheet.styles[next])
  const type = level === null ? schema.nodes['paragraph'] : schema.nodes['heading']
  if (type === undefined) return false

  const attrs = { styleId: storedIdOf(sheet, next), ...(level === null ? {} : { level }) }
  tr.split($from.pos, 1, [{ type, attrs }])
  tr.setStoredMarks([])
  tr.scrollIntoView()
  return true
}

/** Os estilos que o editor conhece — os mesmos de `paragraphCommands`. */
function stylesOf(storage: Record<string, unknown>): StyleSheet | null {
  const paragraph = storage['paragraphCommands'] as { styles?: StyleSheet | null } | undefined
  return paragraph?.styles ?? null
}

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    styleCommands: {
      applyParagraphStyle: (styleId: string) => ReturnType
      applyCharacterStyle: (styleId: string | null) => ReturnType
      clearDirectFormatting: () => ReturnType
      /** Negrito, itálico, sublinhado e tachado, com o estilo por baixo. */
      toggleInheritedMark: (name: string) => ReturnType
    }
  }
}

export const StyleCommands = Extension.create({
  name: 'styleCommands',

  // Acima do Tiptap: `Mod-b` e o Enter são dele também, e aqui decidem primeiro.
  priority: 200,

  addGlobalAttributes() {
    // O "desligado" de cada marca: como o CSS desfaz o que a regra do estilo dá.
    // Sublinhado e tachado do bloco não se desfazem num filho — a decoração do
    // pai atravessa —, e só um `inline-block` a interrompe.
    const off = (css: string) => ({
      off: {
        default: null,
        parseHTML: (element: HTMLElement) => (element.hasAttribute('data-off') ? true : null),
        renderHTML: (attributes: Record<string, unknown>) =>
          attributes['off'] === true ? { 'data-off': '', style: css } : {},
      },
    })
    return [
      { types: ['bold'], attributes: off('font-weight: 400') },
      { types: ['italic'], attributes: off('font-style: normal') },
      { types: ['underline', 'strike'], attributes: off('text-decoration: none; display: inline-block') },
    ]
  },

  addCommands() {
    const sheet = () => stylesOf(this.editor.storage as unknown as Record<string, unknown>)
    return {
      applyParagraphStyle:
        (styleId) =>
        ({ tr, dispatch }) => {
          const current = sheet()
          if (current === null) return false
          if (dispatch === undefined) return true
          // Um passo de desfazer só dele, e não emendado ao texto digitado antes.
          closeHistory(tr)
          return applyParagraphStyle(tr, current, styleId)
        },
      applyCharacterStyle:
        (styleId) =>
        ({ tr, dispatch }) =>
          dispatch === undefined || applyCharacterStyle(tr, styleId),
      clearDirectFormatting:
        () =>
        ({ tr, dispatch }) => {
          if (dispatch !== undefined) clearDirectFormatting(closeHistory(tr))
          return true
        },
      toggleInheritedMark:
        (name) =>
        ({ tr, dispatch, commands }) => {
          if (!inheritedOn(tr, sheet(), name)) return commands.toggleMark(name)
          return dispatch === undefined || toggleInheritedMark(tr, sheet(), name)
        },
    }
  },

  addKeyboardShortcuts() {
    const toggle = (name: string) => () => this.editor.commands.toggleInheritedMark(name)
    return {
      'Mod-b': toggle('bold'),
      'Mod-i': toggle('italic'),
      'Mod-u': toggle('underline'),
      Enter: () =>
        this.editor.commands.command(({ tr, dispatch }) => {
          const current = stylesOf(this.editor.storage as unknown as Record<string, unknown>)
          if (dispatch === undefined) return false
          return splitWithNextStyle(tr, current)
        }),
    }
  },
})

/**
 * O estilo de caractere do trecho (`w:rStyle`). Não desenha nada sozinho: quem
 * desenha é a regra de `style-css.ts` para `[data-char-style]`.
 */
export const CharacterStyle = Mark.create({
  name: 'charStyle',

  addAttributes() {
    return {
      styleId: {
        default: null,
        parseHTML: (element) => element.getAttribute('data-char-style'),
        renderHTML: (attributes) => {
          const id = attributes['styleId']
          return typeof id === 'string' && id !== '' ? { 'data-char-style': id } : {}
        },
      },
    }
  },

  parseHTML() {
    return [{ tag: 'span[data-char-style]' }]
  },

  renderHTML({ HTMLAttributes }) {
    return ['span', mergeAttributes(HTMLAttributes), 0]
  },
})
