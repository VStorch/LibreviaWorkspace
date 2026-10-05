import type { DocumentNode } from './model.js'

/**
 * Words and characters come from `CharacterCount`, as in the status bar, so there are not two
 * numbers.
 */

/** Every Unicode space, like the `\u00a0` Word puts in dates and numbers. */
export function charactersWithoutSpaces(text: string): number {
  return [...text.replace(/\s|\u00a0/gu, '')].length
}

/** Lists, quotes and cells **contain** paragraphs: counting them would count each line twice. */
const PARAGRAPH_TYPES = new Set(['paragraph', 'heading', 'codeBlock'])

/** As in Word, only lines with text; the section mark, an empty paragraph, is left out too. */
export function countParagraphs(node: DocumentNode): number {
  let total = 0

  const visit = (current: DocumentNode): void => {
    if (PARAGRAPH_TYPES.has(current.type)) {
      if (inlineTextOf(current).trim() !== '') total += 1
      return
    }

    for (const child of current.content ?? []) visit(child)
  }

  visit(node)
  return total
}

function inlineTextOf(node: DocumentNode): string {
  if (typeof node.text === 'string') return node.text
  return (node.content ?? []).map(inlineTextOf).join('')
}
