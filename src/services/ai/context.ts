export interface DocumentContext {
  readonly before: string
  readonly after: string
  /** Part of the document was left out: the model is told it sees an excerpt. */
  readonly excerpt: boolean
}

export const SELECTION_OPEN = '<selection>'
export const SELECTION_CLOSE = '</selection>'

const SELECTION_TAGS = /<\/?selection>/gi
const WORD_BOUNDARY = /\s/

export const NO_CONTEXT: DocumentContext = { before: '', after: '', excerpt: false }

/** A document that happens to contain the tags would confuse the model about where the selection is. */
export function stripSelectionTags(text: string): string {
  return text.replace(SELECTION_TAGS, '')
}

/** The end of `text`, starting at a word: a cut word would read as a typo to fix. */
function tail(text: string, length: number): string {
  if (length >= text.length) return text
  const cut = text.slice(text.length - length)
  if (WORD_BOUNDARY.test(text[text.length - length - 1]!)) return cut
  const start = cut.search(WORD_BOUNDARY)
  return start < 0 ? cut : cut.slice(start + 1)
}

function head(text: string, length: number): string {
  if (length >= text.length) return text
  const cut = text.slice(0, length)
  if (WORD_BOUNDARY.test(text[length]!)) return cut
  let end = cut.length
  while (end > 0 && !WORD_BOUNDARY.test(cut[end - 1]!)) end--
  return end === 0 ? cut : cut.slice(0, end - 1)
}

/**
 * Half the budget to each side of the selection; what one side does not use goes to the other, so
 * a selection near the start still gets as much context as one in the middle.
 */
export function contextWindow(before: string, after: string, budget: number): DocumentContext {
  const cleanBefore = stripSelectionTags(before)
  const cleanAfter = stripSelectionTags(after)
  if (budget <= 0) return NO_CONTEXT
  const half = Math.floor(budget / 2)
  const beforeLength = Math.min(cleanBefore.length, Math.max(half, budget - cleanAfter.length))
  const afterLength = Math.min(cleanAfter.length, budget - beforeLength)
  return {
    before: tail(cleanBefore, beforeLength),
    after: head(cleanAfter, afterLength),
    excerpt: beforeLength < cleanBefore.length || afterLength < cleanAfter.length,
  }
}
