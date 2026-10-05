export interface Occurrence {
  readonly start: number
  readonly end: number
}

/** Non-overlapping: "aa" in "aaaa" is two, and "replace all" ends. */
export function findOccurrences(haystack: string, needle: string, caseSensitive = false): Occurrence[] {
  if (needle.length === 0) return []

  const subject = caseSensitive ? haystack : foldCase(haystack)
  const target = caseSensitive ? needle : foldCase(needle)

  const found: Occurrence[] = []
  let index = subject.indexOf(target)

  while (index !== -1) {
    found.push({ start: index, end: index + needle.length })
    index = subject.indexOf(target, index + target.length)
  }

  return found
}

/**
 * `toLowerCase()` changes the length of Turkish `İ`, and positions are the document's. Each code
 * point is only converted when it fits in the same space.
 */
function foldCase(text: string): string {
  // By code point, so a surrogate pair is not split.
  let folded = ''
  for (const character of text) {
    const lower = character.toLowerCase()
    folded += lower.length === character.length ? lower : character
  }

  return folded
}

export function stepIndex(current: number, total: number, delta: number): number {
  if (total === 0) return -1
  if (current < 0) return delta > 0 ? 0 : total - 1
  return (current + delta + total) % total
}
