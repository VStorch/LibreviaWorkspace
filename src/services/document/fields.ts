/**
 * `PAGEREF _Toc123 \h`, `SEQ Figura \* ARABIC`, `TOC \o "1-3" \h \z \u`. Only what references use;
 * the instruction goes back to the file as it came.
 */

/** `PAGEREF`, `REF`, `SEQ`, `TOC`… Uppercase, as Word writes it. */
export function fieldKind(instr: string): string {
  return /^\s*([A-Za-z]+)/.exec(instr)?.[1]?.toUpperCase() ?? ''
}

/** Unquoted: `TOC \o "1-3"` → `TOC`, `\o`, `1-3`. */
export function fieldTokens(instr: string): string[] {
  const tokens: string[] = []
  const pattern = /"([^"]*)"|(\S+)/g
  for (const match of instr.matchAll(pattern)) tokens.push(match[1] ?? match[2] ?? '')
  return tokens
}

/** The bookmark of `REF` and `PAGEREF`, the identifier of `SEQ`. */
export function fieldArgument(instr: string): string | null {
  const argument = fieldTokens(instr)[1]
  return argument === undefined || argument.startsWith('\\') ? null : argument
}

/** `\o "1-3"` → `1-3`; `''` for a switch without a value, `null` when absent. */
export function fieldSwitch(instr: string, name: string): string | null {
  const tokens = fieldTokens(instr)
  const index = tokens.findIndex((token) => token.toLowerCase() === `\\${name.toLowerCase()}`)
  if (index < 0) return null
  const value = tokens[index + 1]
  return value === undefined || value.startsWith('\\') ? '' : value
}

/**
 * Without `\o`, or with it empty, all nine, as in Word; a table of contents with only `\t` falls
 * back to the first three.
 */
export function tocLevels(instr: string): { readonly from: number; readonly to: number } {
  const range = fieldSwitch(instr, 'o')
  if (range === null) return { from: 1, to: 3 }
  const match = /^(\d)\s*-\s*(\d)$/.exec(range)
  if (match === null) return { from: 1, to: 9 }
  const from = Number(match[1])
  const to = Number(match[2])
  return from <= to ? { from, to } : { from: to, to: from }
}

/** `\h`. */
export function tocLinks(instr: string): boolean {
  return fieldSwitch(instr, 'h') !== null
}

/** `\n` without a range. */
export function tocOmitsPages(instr: string): boolean {
  return fieldSwitch(instr, 'n') === ''
}

const ROMAN: ReadonlyArray<readonly [number, string]> = [
  [1000, 'M'],
  [900, 'CM'],
  [500, 'D'],
  [400, 'CD'],
  [100, 'C'],
  [90, 'XC'],
  [50, 'L'],
  [40, 'XL'],
  [10, 'X'],
  [9, 'IX'],
  [5, 'V'],
  [4, 'IV'],
  [1, 'I'],
]

function roman(value: number): string {
  let rest = value
  let text = ''
  for (const [amount, letters] of ROMAN) {
    while (rest >= amount) {
      text += letters
      rest -= amount
    }
  }
  return text
}

function alphabetic(value: number): string {
  // After Z, Word counts AA, BB…
  const letter = String.fromCharCode(65 + ((value - 1) % 26))
  return letter.repeat(Math.floor((value - 1) / 26) + 1)
}

/** `\* ARABIC`, `ROMAN`, `roman`, `ALPHABETIC` or `alphabetic`. */
export function formatFieldNumber(value: number, instr: string): string {
  const format = fieldSwitch(instr, '*') ?? 'ARABIC'
  if (value < 1) return String(value)
  switch (format) {
    case 'ROMAN':
      return roman(value)
    case 'roman':
      return roman(value).toLowerCase()
    case 'ALPHABETIC':
      return alphabetic(value)
    case 'alphabetic':
      return alphabetic(value).toLowerCase()
    default:
      return String(value)
  }
}

/**
 * Each identifier counts apart, case-insensitively; `\r n` restarts at `n`, `\c` repeats the last.
 */
export function sequenceNumbers(instructions: readonly string[]): string[] {
  const counters = new Map<string, number>()
  return instructions.map((instr) => {
    const name = (fieldArgument(instr) ?? '').toLowerCase()
    const current = counters.get(name) ?? 0
    const reset = fieldSwitch(instr, 'r')
    const next =
      reset !== null && reset !== '' && Number.isInteger(Number(reset))
        ? Number(reset)
        : fieldSwitch(instr, 'c') !== null
          ? current
          : current + 1
    counters.set(name, next)
    return formatFieldNumber(next, instr)
  })
}
