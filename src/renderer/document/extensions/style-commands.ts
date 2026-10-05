import { Extension, Mark, mergeAttributes } from '@tiptap/core'
import type { Mark as ProseMirrorMark, MarkType, Node as ProseMirrorNode } from '@tiptap/pm/model'
import { closeHistory } from '@tiptap/pm/history'
import type { Transaction } from '@tiptap/pm/state'
import { blockStyleOfNode } from '@services/document/style-cascade.js'
import { headingLevelOfStyle, nextStyleIdOf } from '@services/document/style-editing.js'
import type { StyleCharacterFormat, StyleSheet } from '@services/document/styles.js'

/**
 * Everything in editor transactions, with undo; modifying and creating styles changes the store's
 * sheet (`style-editing.ts`). The functions stand alone so tests run them on an `EditorState`.
 */

/** What "clear" removes. */
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
  'keepLines',
  'widowControl',
  'fontFamily',
  'fontSize',
] as const

/**
 * The ones that get "off": removing bold from a word in a heading needs a mark that says the
 * opposite. It comes from `w:b w:val="0"` and goes back to that.
 */
export const INHERITABLE_MARKS: Readonly<Record<string, keyof StyleCharacterFormat>> = {
  bold: 'bold',
  italic: 'italic',
  underline: 'underline',
  strike: 'strike',
}

/** A link is content, and a character style is style. */
const KEPT_MARKS = new Set(['link', 'charStyle'])

const isStyledBlock = (node: ProseMirrorNode): boolean =>
  node.type.name === 'paragraph' || node.type.name === 'heading'

/**
 * None for the default paragraph style, as the reader produces for a paragraph without `w:pStyle`.
 */
function storedIdOf(sheet: StyleSheet, styleId: string): string | null {
  return styleId === sheet.defaults.paragraphStyleId ? null : styleId
}

/** A heading without an id, by the name `heading N`. */
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
 * Heading ↔ paragraph by style name. The paragraph's direct formatting goes, as in Word; marks
 * stay.
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

/** `null` removes the style. */
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

/** The style stays and starts drawing. Without a selection, the whole block, as in Word. */
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

function inheritedOn(tr: Transaction, sheet: StyleSheet | null, name: string): boolean {
  const field = INHERITABLE_MARKS[name]
  if (field === undefined) return false
  return blockStyleOfNode(tr.selection.$from.parent, sheet)?.character[field] === true
}

/** Stored marks, the cursor's, or the range's. */
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
 * Without a mark the style applies; a mark with `off` turns it off, and a plain one turns it on.
 */
export function markVisiblyOn(tr: Transaction, sheet: StyleSheet | null, name: string): boolean {
  const type = tr.doc.type.schema.marks[name]
  if (type === undefined) return false
  const mark = marksHere(tr, type)
  if (mark !== null) return mark.attrs['off'] !== true
  return inheritedOn(tr, sheet, name)
}

/**
 * Where the style does not turn the mark on, `false`: the caller uses Tiptap's command. Where it
 * does, the range that is on gets the mark with `off`, and the range that is off goes back to the
 * style.
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
 * A heading gives way to its `next` style. Only on a loose block: lists and cells own their Enter.
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

/** The same as `paragraphCommands`. */
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
      toggleInheritedMark: (name: string) => ReturnType
    }
  }
}

export const StyleCommands = Extension.create({
  name: 'styleCommands',

  // Above Tiptap: `Mod-b` and Enter are decided here first.
  priority: 200,

  addGlobalAttributes() {
    // Block underline and strikethrough cross into children: only an `inline-block` interrupts
    // them.
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
          // An undo step of its own.
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

/** `w:rStyle`: the `style-css.ts` rule for `[data-char-style]` draws it. */
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
