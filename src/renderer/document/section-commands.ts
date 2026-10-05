import type { Editor } from '@tiptap/react'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import type { SectionStart } from '@services/document/model.js'
import {
  blockSections,
  planSectionBreak,
  planSectionDelete,
  resolveSections,
  sectionBreakIn,
  storeSections,
  type ResolvedSections,
  type SectionBlock,
  type SectionList,
} from '@services/document/sections.js'
import { t } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'

/**
 * The id lives in the paragraph closing the section, like `w:sectPr`, and the setup in the store's
 * library. The text says which sections count (`resolveSections`): so structure changes in an
 * editor transaction, and undo brings the section back with the text.
 */

export function marksOfDoc(doc: ProseMirrorNode): (string | null)[] {
  const marks: (string | null)[] = []
  doc.forEach((block) => marks.push(sectionBreakIn(block as unknown as SectionBlock)))
  return marks
}

export function resolvedOf(doc: ProseMirrorNode): ResolvedSections {
  const store = useWorkspace.getState()
  return resolveSections(marksOfDoc(doc), doc.attrs['bodySection'], store.page, store.sections)
}

export function commitSections(next: SectionList, bodyId: string | null): void {
  const store = useWorkspace.getState()
  const stored = storeSections(next, bodyId, store.page, store.sections)
  if (stored.page !== store.page) store.setPage(stored.page)
  store.setSections(stored.library)
}

/** As an index into `allSections`. */
export function sectionAtCursor(
  editor: Editor,
  resolved: ResolvedSections = resolvedOf(editor.state.doc),
): number {
  const marks = marksOfDoc(editor.state.doc)
  const index = Math.min(editor.state.selection.$from.index(0), marks.length - 1)
  return blockSections(marks, resolved.sections)[index] ?? resolved.sections.length
}

/**
 * A draft older than sections (`.sdoc` < 6) cannot store breaks, columns or links: the command
 * refuses.
 */
export function sectionEditsAllowed(): boolean {
  const store = useWorkspace.getState()
  if (!store.beforeSections) return true
  store.showError({ code: 'INTERNAL', message: t('document.sections.legacyDraft') })
  return false
}

/** The mark cannot live in a cell. */
function insideTable(editor: Editor): boolean {
  const { $from } = editor.state.selection
  for (let depth = $from.depth; depth > 0; depth--) {
    if ($from.node(depth).type.spec.tableRole !== undefined) return true
  }
  return false
}

/** The break goes after the list. */
function insideList(editor: Editor): boolean {
  const { $from } = editor.state.selection
  for (let depth = $from.depth; depth > 0; depth--) {
    if ($from.node(depth).type.name === 'listItem') return true
  }
  return false
}

function renameMark(tr: Transaction, from: string, to: string): void {
  tr.doc.descendants((node, pos) => {
    if (node.attrs['sectionBreak'] === from) tr.setNodeAttribute(pos, 'sectionBreak', to)
    return node.isBlock && !node.isTextblock
  })
}

/**
 * As in Word: the paragraph splits, and the upper half closes the new section, a copy of the split
 * one. In a list item the list splits; in a table, the break comes after it, in its own paragraph.
 */
export function insertSectionBreak(editor: Editor, start: SectionStart): void {
  if (!sectionEditsAllowed()) return
  const store = useWorkspace.getState()
  const resolved = resolvedOf(editor.state.doc)
  const plan = planSectionBreak(resolved, store.sections, sectionAtCursor(editor, resolved), start)
  // The library before the text: the new mark must find its section when pagination measures.
  store.setSections([...store.sections, ...plan.additions])

  const structure = (tr: Transaction): void => {
    if (plan.rename !== null) renameMark(tr, plan.rename.from, plan.rename.to)
    if (plan.bodyId !== null) tr.setDocAttribute('bodySection', plan.bodyId)
  }

  // In an item the break stays in the item, as in Word: the list splits after the outer item, and
  // the second part takes the same attributes.
  const $cursor = editor.state.selection.$from
  if (
    !insideTable(editor) &&
    insideList(editor) &&
    $cursor.parent.isTextblock &&
    $cursor.parent.attrs['sectionBreak'] === null
  ) {
    editor
      .chain()
      .focus()
      .command(({ tr }) => {
        structure(tr)
        const $from = tr.selection.$from
        tr.setNodeAttribute($from.before($from.depth), 'sectionBreak', plan.upperId)
        const list = $from.node(1)
        if ($from.index(1) < list.childCount - 1) tr.split($from.after(2), 1)
        return true
      })
      .run()
    return
  }

  if (insideTable(editor) || insideList(editor)) {
    editor
      .chain()
      .focus()
      .command(({ tr, state }) => {
        structure(tr)
        const after = tr.mapping.map(tr.selection.$from.after(1))
        tr.insert(
          after,
          state.schema.nodes['paragraph']!.create({ sectionBreak: plan.upperId, sectionMark: true }),
        )
        return true
      })
      .run()
    return
  }

  editor
    .chain()
    .focus()
    .splitBlock()
    .command(({ tr }) => {
      // The mark the paragraph already had stays with the lower half, where it ends.
      const $cursor = tr.selection.$from
      const lower = $cursor.before($cursor.depth)
      const upperNode = tr.doc.resolve(lower).nodeBefore
      if (upperNode === null) return false
      const upper = lower - upperNode.nodeSize
      const previous = upperNode.attrs['sectionBreak'] as string | null
      tr.setNodeAttribute(upper, 'sectionBreak', plan.upperId)
      // An empty paragraph with a mark is a mark, as the reader delivers it.
      if (upperNode.content.size === 0) tr.setNodeAttribute(upper, 'sectionMark', true)
      if (previous !== null) tr.setNodeAttribute(lower, 'sectionBreak', previous)
      structure(tr)
      return true
    })
    .run()
}

/**
 * `w:br w:type="column"`: a block property (`columnBreakAfter`), like the page break Word stores in
 * the paragraph.
 */
export function insertColumnBreak(editor: Editor): void {
  if (!sectionEditsAllowed() || insideTable(editor)) return
  editor
    .chain()
    .focus()
    .splitBlock()
    .command(({ tr }) => {
      const $cursor = tr.selection.$from
      const lower = $cursor.before($cursor.depth)
      const upperNode = tr.doc.resolve(lower).nodeBefore
      if (upperNode === null) return false
      tr.setNodeAttribute(lower - upperNode.nodeSize, 'columnBreakAfter', true)
      return true
    })
    .run()
}

/**
 * Or, in the last section, the one opening it. As in Word, the text above takes the format of the
 * section below, which gets the bands it inherited (`planSectionDelete`).
 */
export function deleteSectionBreak(editor: Editor): boolean {
  if (!sectionEditsAllowed()) return false
  const resolved = resolvedOf(editor.state.doc)
  const plan = planSectionDelete(resolved, sectionAtCursor(editor, resolved))
  if (plan === null) return false

  let found: { pos: number; node: ProseMirrorNode } | null = null
  editor.state.doc.descendants((node, pos) => {
    if (found !== null) return false
    if (node.attrs['sectionBreak'] === plan.removeId) found = { pos, node }
    return found === null
  })
  const target = found as { pos: number; node: ProseMirrorNode } | null
  if (target === null) return false

  commitSections(plan.next, resolved.bodyId)
  editor
    .chain()
    .focus()
    .command(({ tr }) => {
      if (
        target.node.attrs['sectionMark'] === true &&
        target.node.content.size === 0 &&
        tr.doc.childCount > 1
      ) {
        tr.delete(target.pos, target.pos + target.node.nodeSize)
      } else {
        tr.setNodeAttribute(target.pos, 'sectionBreak', null)
      }
      return true
    })
    .run()
  return true
}
