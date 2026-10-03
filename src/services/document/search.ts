export interface Occurrence {
  readonly start: number
  readonly end: number
}

/** Não sobrepostas: "aa" em "aaaa" são duas, e "substituir tudo" termina. */
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
 * `toLowerCase()` muda o comprimento do `İ` turco, e as posições são as do
 * documento. Cada ponto de código só é convertido quando cabe no mesmo espaço.
 */
function foldCase(text: string): string {
  // Por ponto de código, para não partir um par substituto.
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
