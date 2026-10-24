export const DiffOp = { Same: 'same', Removed: 'removed', Added: 'added' } as const

export type DiffOp = (typeof DiffOp)[keyof typeof DiffOp]

export interface DiffPiece {
  readonly op: DiffOp
  readonly text: string
}

/** Words, runs of spaces and each other character: a fixed comma does not take its word along. */
const TOKEN = /[\p{L}\p{N}_]+|\s+|[^\p{L}\p{N}_\s]/gu

/** The comparison table grows with the product of the lengths; past it, the whole text changes. */
const MAX_TABLE_CELLS = 250_000

const tokensOf = (text: string): string[] => text.match(TOKEN) ?? []

function push(pieces: DiffPiece[], op: DiffOp, text: string): void {
  const last = pieces[pieces.length - 1]
  if (last?.op === op) pieces[pieces.length - 1] = { op, text: last.text + text }
  else if (text !== '') pieces.push({ op, text })
}

/** How many tokens the two tails from each pair of indexes have in common. */
function commonTails(before: readonly string[], after: readonly string[]): Uint32Array {
  const width = after.length + 1
  const table = new Uint32Array((before.length + 1) * width)
  for (let i = before.length - 1; i >= 0; i--) {
    for (let j = after.length - 1; j >= 0; j--) {
      table[i * width + j] =
        before[i] === after[j]
          ? table[(i + 1) * width + j + 1]! + 1
          : Math.max(table[(i + 1) * width + j]!, table[i * width + j + 1]!)
    }
  }
  return table
}

/** Removals come before additions at the same place, as a reader expects a correction. */
export function wordDiff(before: string, after: string): DiffPiece[] {
  const a = tokensOf(before)
  const b = tokensOf(after)
  const pieces: DiffPiece[] = []
  if ((a.length + 1) * (b.length + 1) > MAX_TABLE_CELLS) {
    push(pieces, DiffOp.Removed, before)
    push(pieces, DiffOp.Added, after)
    return pieces
  }
  const table = commonTails(a, b)
  const width = b.length + 1
  let i = 0
  let j = 0
  let added = ''
  const flush = (): void => {
    push(pieces, DiffOp.Added, added)
    added = ''
  }
  while (i < a.length || j < b.length) {
    if (i < a.length && j < b.length && a[i] === b[j]) {
      flush()
      push(pieces, DiffOp.Same, a[i++]!)
      j++
    } else if (j < b.length && (i === a.length || table[i * width + j + 1]! >= table[(i + 1) * width + j]!)) {
      added += b[j++]!
    } else {
      push(pieces, DiffOp.Removed, a[i++]!)
    }
  }
  flush()
  return pieces
}
