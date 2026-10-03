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
 * O id mora no parágrafo que fecha a seção, como o `w:sectPr`, e a configuração
 * na biblioteca da loja. Quem diz quais seções valem é o texto
 * (`resolveSections`): por isso a estrutura muda numa transação do editor, e o
 * desfazer volta a seção com o texto.
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

/** Como índice em `allSections`. */
export function sectionAtCursor(
  editor: Editor,
  resolved: ResolvedSections = resolvedOf(editor.state.doc),
): number {
  const marks = marksOfDoc(editor.state.doc)
  const index = Math.min(editor.state.selection.$from.index(0), marks.length - 1)
  return blockSections(marks, resolved.sections)[index] ?? resolved.sections.length
}

/** O rascunho anterior às seções (`.sdoc` < 6) não grava quebra, coluna nem vínculo: o comando recusa. */
export function sectionEditsAllowed(): boolean {
  const store = useWorkspace.getState()
  if (!store.beforeSections) return true
  store.showError({ code: 'INTERNAL', message: t('document.sections.legacyDraft') })
  return false
}

/** A marca não pode morar numa célula. */
function insideTable(editor: Editor): boolean {
  const { $from } = editor.state.selection
  for (let depth = $from.depth; depth > 0; depth--) {
    if ($from.node(depth).type.spec.tableRole !== undefined) return true
  }
  return false
}

/** A quebra vai depois da lista. */
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
 * Como o Word: o parágrafo se parte, e a metade de cima fecha a seção nova,
 * cópia da partida. Num item de lista a lista se parte; numa tabela, a quebra vem
 * depois dela, num parágrafo próprio.
 */
export function insertSectionBreak(editor: Editor, start: SectionStart): void {
  if (!sectionEditsAllowed()) return
  const store = useWorkspace.getState()
  const resolved = resolvedOf(editor.state.doc)
  const plan = planSectionBreak(resolved, store.sections, sectionAtCursor(editor, resolved), start)
  // A biblioteca antes do texto: a marca nova precisa achar a seção quando a paginação medir.
  store.setSections([...store.sections, ...plan.additions])

  const structure = (tr: Transaction): void => {
    if (plan.rename !== null) renameMark(tr, plan.rename.from, plan.rename.to)
    if (plan.bodyId !== null) tr.setDocAttribute('bodySection', plan.bodyId)
  }

  // Num item a quebra fica no item, como no Word: a lista se parte depois do item
  // de fora, e a segunda parte leva os mesmos atributos.
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
      // A marca que o parágrafo já tinha fica com a metade de baixo, onde ele termina.
      const $cursor = tr.selection.$from
      const lower = $cursor.before($cursor.depth)
      const upperNode = tr.doc.resolve(lower).nodeBefore
      if (upperNode === null) return false
      const upper = lower - upperNode.nodeSize
      const previous = upperNode.attrs['sectionBreak'] as string | null
      tr.setNodeAttribute(upper, 'sectionBreak', plan.upperId)
      // Parágrafo vazio com marca é marca, como o leitor o entrega.
      if (upperNode.content.size === 0) tr.setNodeAttribute(upper, 'sectionMark', true)
      if (previous !== null) tr.setNodeAttribute(lower, 'sectionBreak', previous)
      structure(tr)
      return true
    })
    .run()
}

/** `w:br w:type="column"`: propriedade do bloco (`columnBreakAfter`), como a quebra de página que o Word grava no parágrafo. */
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
 * Ou, na última seção, a que a abre. Como no Word, o texto de cima passa ao
 * formato da seção de baixo, que recebe as faixas que herdava (`planSectionDelete`).
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
