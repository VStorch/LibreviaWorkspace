export interface SplitLines {
  readonly lines: string[]
  readonly rest: string
}

/** A network chunk may end mid-line: the unfinished tail waits for the next one. */
export function takeLines(buffer: string): SplitLines {
  const parts = buffer.split(/\r?\n/u)
  const rest = parts.pop() ?? ''
  return { lines: parts.filter((line) => line.trim() !== ''), rest }
}
