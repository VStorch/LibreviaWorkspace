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
 * Inserir e excluir quebra de seção (M9).
 *
 * A quebra mora em dois lugares: o id no parágrafo que fecha a seção (a marca,
 * como o `w:sectPr` do OOXML) e a configuração na biblioteca de seções da loja.
 * Quem diz quais seções valem, e em que ordem, é o texto (`resolveSections`):
 * por isso as mudanças de estrutura vão todas numa transação do editor, e a
 * biblioteca só ganha entradas — o desfazer volta o texto, e a seção volta com
 * ele.
 */

/** As marcas de seção dos blocos de primeiro nível, na ordem do corpo. */
export function marksOfDoc(doc: ProseMirrorNode): (string | null)[] {
  const marks: (string | null)[] = []
  doc.forEach((block) => marks.push(sectionBreakIn(block as unknown as SectionBlock)))
  return marks
}

/** As seções que o texto do editor usa agora, com a configuração da loja. */
export function resolvedOf(doc: ProseMirrorNode): ResolvedSections {
  const store = useWorkspace.getState()
  return resolveSections(marksOfDoc(doc), doc.attrs['bodySection'], store.page, store.sections)
}

/** Grava na loja as seções mudadas por um painel, pela biblioteca. */
export function commitSections(next: SectionList, bodyId: string | null): void {
  const store = useWorkspace.getState()
  const stored = storeSections(next, bodyId, store.page, store.sections)
  if (stored.page !== store.page) store.setPage(stored.page)
  store.setSections(stored.library)
}

/** A seção do cursor, como índice em `allSections` das seções resolvidas. */
export function sectionAtCursor(
  editor: Editor,
  resolved: ResolvedSections = resolvedOf(editor.state.doc),
): number {
  const marks = marksOfDoc(editor.state.doc)
  const index = Math.min(editor.state.selection.$from.index(0), marks.length - 1)
  return blockSections(marks, resolved.sections)[index] ?? resolved.sections.length
}

/**
 * Seções e colunas só se editam em documento que as grava.
 *
 * O rascunho de antes das seções (`.sdoc` < 6) é gravado pelo caminho de então,
 * que não conhece quebra, coluna nem vínculo de faixa: a mudança apareceria na
 * tela e sumiria no arquivo. O comando recusa e diz por quê.
 */
export function sectionEditsAllowed(): boolean {
  const store = useWorkspace.getState()
  if (!store.beforeSections) return true
  store.showError({ code: 'INTERNAL', message: t('document.sections.legacyDraft') })
  return false
}

/** O cursor está dentro de uma tabela: a marca não pode morar numa célula. */
function insideTable(editor: Editor): boolean {
  const { $from } = editor.state.selection
  for (let depth = $from.depth; depth > 0; depth--) {
    if ($from.node(depth).type.spec.tableRole !== undefined) return true
  }
  return false
}

/** O cursor está num item de lista: a quebra vai depois da lista. */
function insideList(editor: Editor): boolean {
  const { $from } = editor.state.selection
  for (let depth = $from.depth; depth > 0; depth--) {
    if ($from.node(depth).type.name === 'listItem') return true
  }
  return false
}

/** A marca `from` passa a ser `to`, onde quer que esteja no texto. */
function renameMark(tr: Transaction, from: string, to: string): void {
  tr.doc.descendants((node, pos) => {
    if (node.attrs['sectionBreak'] === from) tr.setNodeAttribute(pos, 'sectionBreak', to)
    return node.isBlock && !node.isTextblock
  })
}

/**
 * Quebra de seção no cursor, como o Word: o parágrafo se parte, e a metade de
 * cima fecha a seção nova — que é cópia da seção partida. A de baixo continua a
 * seção de antes, começando do jeito pedido. Numa tabela ou numa lista, a
 * quebra vem logo depois dela, num parágrafo próprio.
 */
export function insertSectionBreak(editor: Editor, start: SectionStart): void {
  if (!sectionEditsAllowed()) return
  const store = useWorkspace.getState()
  const resolved = resolvedOf(editor.state.doc)
  const plan = planSectionBreak(resolved, store.sections, sectionAtCursor(editor, resolved), start)
  // A biblioteca antes do texto: a marca nova precisa achar a seção dela quando
  // a paginação medir. Entradas a mais não mudam nada até o texto apontá-las.
  store.setSections([...store.sections, ...plan.additions])

  const structure = (tr: Transaction): void => {
    if (plan.rename !== null) renameMark(tr, plan.rename.from, plan.rename.to)
    if (plan.bodyId !== null) tr.setDocAttribute('bodySection', plan.bodyId)
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
      // A metade de baixo é o bloco do cursor; a de cima, o irmão antes dele.
      // A marca que o parágrafo já tinha fica com a de baixo, que é onde o
      // parágrafo termina — o Enter não a leva adiante sozinho.
      const $cursor = tr.selection.$from
      const lower = $cursor.before($cursor.depth)
      const upperNode = tr.doc.resolve(lower).nodeBefore
      if (upperNode === null) return false
      const upper = lower - upperNode.nodeSize
      const previous = upperNode.attrs['sectionBreak'] as string | null
      tr.setNodeAttribute(upper, 'sectionBreak', plan.upperId)
      // Parágrafo vazio que só carrega a marca é marca, como o leitor o entrega:
      // sem altura, e excluído junto com a quebra.
      if (upperNode.content.size === 0) tr.setNodeAttribute(upper, 'sectionMark', true)
      if (previous !== null) tr.setNodeAttribute(lower, 'sectionBreak', previous)
      structure(tr)
      return true
    })
    .run()
}

/**
 * Quebra de coluna no cursor (`w:br w:type="column"`): o parágrafo se parte, e a
 * metade de cima termina a coluna. Como a de página que o Word grava dentro do
 * parágrafo, ela é propriedade do bloco (`columnBreakAfter`).
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
 * Exclui a quebra que fecha a seção do cursor — ou, na última seção, a que a
 * abre. Como no Word, o texto de cima passa a ter o formato da seção de baixo:
 * sem a marca, os blocos são da seção da próxima marca. A seção de baixo recebe
 * as faixas que herdava da excluída (`planSectionDelete`).
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
      // O parágrafo que era só a marca não tem o que sobrar: vai junto.
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
