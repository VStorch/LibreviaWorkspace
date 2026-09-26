/**
 * A estrutura do documento: os títulos, cada um no seu nível.
 *
 * É o que o painel de navegação lista e, no M8, o que o sumário numera. O nível é
 * o **efetivo**, como o Word o decide: o do nó `heading` (que o leitor reconhece
 * pelo nome `heading N` do estilo) ou o `outlineLevel` que a cascata de estilos
 * dá ao parágrafo. Um estilo "Capítulo" criado pelo autor com nível 1 na
 * estrutura é título para o Word — e tem de ser para o painel também, senão o
 * documento que o usa aparece sem estrutura nenhuma.
 *
 * Função pura sobre uma lista plana de blocos, e não sobre o documento do
 * ProseMirror: quem percorre a árvore é o renderer, e aqui fica só a regra, que se
 * testa sem editor.
 */

import { resolveStyle } from './style-cascade.js'
import type { StyleSheet } from './styles.js'

/** Um bloco de texto do documento, com a posição em que começa. */
export interface OutlineBlock {
  readonly type: string
  readonly attrs: Readonly<Record<string, unknown>>
  readonly text: string
  readonly pos: number
}

export interface OutlineEntry {
  /** De 1 a 9, como o Word mostra: o `outlineLevel` do arquivo mais um. */
  readonly level: number
  readonly text: string
  readonly pos: number
}

/** O último nível que é título. `outlineLvl` 9 no arquivo é corpo de texto. */
const DEEPEST_LEVEL = 9

/**
 * O nível do bloco na estrutura, de 1 a 9 — ou `null` quando ele é texto.
 *
 * O `heading` vale pelo `level` que traz: foi pelo nome do estilo que o leitor o
 * fez título, e o estilo pode nem declarar nível. O parágrafo vale pelo que a
 * cascata diz, e sem folha de estilos (o rascunho antigo) não há cascata a
 * consultar.
 */
export function outlineLevelOf(
  block: Pick<OutlineBlock, 'type' | 'attrs'>,
  sheet: StyleSheet | null,
): number | null {
  if (block.type === 'heading') {
    const level = block.attrs['level']
    return typeof level === 'number' && level >= 1 && level <= DEEPEST_LEVEL ? level : null
  }

  if (block.type !== 'paragraph' || sheet === null) return null

  const styleId = block.attrs['styleId']
  const { outlineLevel } = resolveStyle(sheet, typeof styleId === 'string' ? styleId : null).paragraph
  return outlineLevel !== undefined && outlineLevel >= 0 && outlineLevel < DEEPEST_LEVEL
    ? outlineLevel + 1
    : null
}

/**
 * Os títulos, na ordem do documento.
 *
 * O título vazio fica de fora, como no Word: uma linha em branco no estilo
 * "Título 1" — resto comum de documento editado — não é capítulo nenhum, e na
 * lista seria uma entrada sem nome em que não há o que clicar.
 */
export function outlineOf(blocks: readonly OutlineBlock[], sheet: StyleSheet | null): OutlineEntry[] {
  const entries: OutlineEntry[] = []

  for (const block of blocks) {
    const level = outlineLevelOf(block, sheet)
    const text = block.text.replace(/\s+/g, ' ').trim()
    if (level === null || text === '') continue
    entries.push({ level, text, pos: block.pos })
  }

  return entries
}

/**
 * O título da seção em que a posição está: o último que começa antes dela.
 *
 * `-1` antes do primeiro título — a capa, o resumo —, onde não há o que destacar.
 */
export function currentEntryIndex(entries: readonly OutlineEntry[], pos: number): number {
  let current = -1
  for (const [index, entry] of entries.entries()) {
    if (entry.pos > pos) break
    current = index
  }
  return current
}
