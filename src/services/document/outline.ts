/**
 * O nível é o **efetivo**, como o Word o decide: o do nó `heading` ou o
 * `outlineLevel` que a cascata dá ao parágrafo. Um estilo "Capítulo" com nível 1
 * na estrutura é título para o Word, e para o painel também.
 */

import { resolveStyle } from './style-cascade.js'
import type { StyleSheet } from './styles.js'

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

/** O `heading` vale pelo `level` que traz, porque o estilo pode nem declarar nível. */
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

/** O título vazio fica de fora, como no Word. */
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

/** O último título que começa antes da posição; `-1` antes do primeiro. */
export function currentEntryIndex(entries: readonly OutlineEntry[], pos: number): number {
  let current = -1
  for (const [index, entry] of entries.entries()) {
    if (entry.pos > pos) break
    current = index
  }
  return current
}
