import type { Editor } from '@tiptap/react'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import type { SectionSetup, SectionStart } from '@services/document/model.js'
import {
  blockSections,
  freshSectionId,
  sectionBreakIn,
  withSectionBreak,
  withoutSection,
  type SectionBlock,
} from '@services/document/sections.js'
import { useWorkspace } from '../state/workspace.js'

/**
 * Inserir e excluir quebra de seção (M9).
 *
 * A quebra mora em dois lugares: o id no parágrafo que fecha a seção (a marca,
 * como o `w:sectPr` do OOXML) e a configuração na lista de seções da loja. Os
 * dois mudam juntos aqui, e só aqui.
 */

/** A seção do cursor, como índice em `allSections`. */
export function sectionAtCursor(editor: Editor, sections: readonly SectionSetup[]): number {
  const marks: (string | null)[] = []
  editor.state.doc.forEach((block) => marks.push(sectionBreakIn(block as unknown as SectionBlock)))
  const index = Math.min(editor.state.selection.$from.index(0), marks.length - 1)
  return blockSections(marks, sections)[index] ?? sections.length
}

/** O cursor está dentro de uma tabela: a marca não pode morar numa célula. */
function insideTable(editor: Editor): boolean {
  const { $from } = editor.state.selection
  for (let depth = $from.depth; depth > 0; depth--) {
    if ($from.node(depth).type.spec.tableRole !== undefined) return true
  }
  return false
}

/**
 * Quebra de seção no cursor, como o Word: o parágrafo se parte, e a metade de
 * cima fecha a seção nova — que é cópia da seção partida. A de baixo continua a
 * seção de antes, começando do jeito pedido. Numa tabela, a quebra vem logo
 * depois dela, num parágrafo próprio.
 */
export function insertSectionBreak(editor: Editor, start: SectionStart): void {
  const store = useWorkspace.getState()
  const current = sectionAtCursor(editor, store.sections)
  const id = freshSectionId(store.sections)

  const inserted = insideTable(editor)
    ? editor
        .chain()
        .focus()
        .command(({ tr, state }) => {
          const after = tr.selection.$from.after(1)
          tr.insert(after, state.schema.nodes['paragraph']!.create({ sectionBreak: id, sectionMark: true }))
          return true
        })
        .run()
    : editor
        .chain()
        .focus()
        .splitBlock()
        .command(({ tr }) => {
          // A metade de baixo é o bloco do cursor; a de cima, o irmão antes dele.
          // A marca que o parágrafo já tinha fica com a de baixo, que é onde o
          // parágrafo termina — o Enter não a leva adiante sozinho.
          const $cursor = tr.selection.$from
          const lower = $cursor.before($cursor.depth)
          const upperNode = tr.doc.resolve(lower).nodeBefore
          if (upperNode === null) return false
          const upper = lower - upperNode.nodeSize
          const previous = upperNode.attrs['sectionBreak'] as string | null
          tr.setNodeAttribute(upper, 'sectionBreak', id)
          if (previous !== null) tr.setNodeAttribute(lower, 'sectionBreak', previous)
          return true
        })
        .run()
  if (!inserted) return

  const next = withSectionBreak({ page: store.page, sections: store.sections }, current, id, start)
  if (next.page !== store.page) store.setPage(next.page)
  store.setSections(next.sections)
}

/**
 * Quebra de coluna no cursor (`w:br w:type="column"`): o parágrafo se parte, e a
 * metade de cima termina a coluna. Como a de página que o Word grava dentro do
 * parágrafo, ela é propriedade do bloco (`columnBreakAfter`).
 */
export function insertColumnBreak(editor: Editor): void {
  if (insideTable(editor)) return
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
 * Exclui a quebra que fecha a seção do cursor — ou, na última seção, a que a
 * abre. Como no Word, o texto de cima passa a ter o formato da seção de baixo:
 * sem a marca, os blocos são da seção da próxima marca.
 */
export function deleteSectionBreak(editor: Editor): boolean {
  const store = useWorkspace.getState()
  if (store.sections.length === 0) return false
  const current = sectionAtCursor(editor, store.sections)
  const id = (store.sections[current] ?? store.sections.at(-1))!.id

  let target: { pos: number; node: ProseMirrorNode } | null = null
  editor.state.doc.descendants((node, pos) => {
    if (target !== null) return false
    if (node.attrs['sectionBreak'] === id) target = { pos, node }
    return target === null
  })

  const found = target as { pos: number; node: ProseMirrorNode } | null
  if (found !== null) {
    editor
      .chain()
      .focus()
      .command(({ tr }) => {
        // O parágrafo que era só a marca não tem o que sobrar: vai junto.
        if (
          found.node.attrs['sectionMark'] === true &&
          found.node.content.size === 0 &&
          tr.doc.childCount > 1
        ) {
          tr.delete(found.pos, found.pos + found.node.nodeSize)
        } else {
          tr.setNodeAttribute(found.pos, 'sectionBreak', null)
        }
        return true
      })
      .run()
  }

  store.setSections(withoutSection({ page: store.page, sections: store.sections }, id).sections)
  return true
}
