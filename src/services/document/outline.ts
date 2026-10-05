/**
 * The level is the **effective** one, as Word decides it: the `heading` node's, or the
 * `outlineLevel` the cascade gives the paragraph. A "Chapter" style with outline level 1 is a
 * heading to Word, and to the pane as well.
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
  /** 1 to 9, as Word shows it: the file's `outlineLevel` plus one. */
  readonly level: number
  readonly text: string
  readonly pos: number
}

/** The deepest heading level. `outlineLvl` 9 in the file is body text. */
const DEEPEST_LEVEL = 9

/** A `heading` counts by its `level`, because the style may not declare a level at all. */
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

/** An empty heading is left out, as in Word. */
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

/** The last heading starting before the position; `-1` before the first. */
export function currentEntryIndex(entries: readonly OutlineEntry[], pos: number): number {
  let current = -1
  for (const [index, entry] of entries.entries()) {
    if (entry.pos > pos) break
    current = index
  }
  return current
}
