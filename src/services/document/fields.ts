/**
 * A gramática dos campos do Word que o editor sabe recalcular, sem editor.
 *
 * Uma instrução é a palavra do campo seguida de argumentos e de chaves:
 * `PAGEREF _Toc123 \h`, `SEQ Figura \* ARABIC`, `TOC \o "1-3" \h \z \u`. O
 * argumento pode vir entre aspas, e a chave `\*` leva o formato do número. Só o
 * que o M8 usa está aqui; o resto da instrução é preservado como veio — é ela
 * que volta ao arquivo, e não o que se entendeu dela.
 */

/** A palavra do campo: `PAGEREF`, `REF`, `SEQ`, `TOC`… Maiúscula, como o Word a escreve. */
export function fieldKind(instr: string): string {
  return /^\s*([A-Za-z]+)/.exec(instr)?.[1]?.toUpperCase() ?? ''
}

/** As palavras da instrução, com as aspas desfeitas: `TOC \o "1-3"` → `TOC`, `\o`, `1-3`. */
export function fieldTokens(instr: string): string[] {
  const tokens: string[] = []
  const pattern = /"([^"]*)"|(\S+)/g
  for (const match of instr.matchAll(pattern)) tokens.push(match[1] ?? match[2] ?? '')
  return tokens
}

/** O primeiro argumento: o marcador de `REF` e `PAGEREF`, o identificador de `SEQ`. */
export function fieldArgument(instr: string): string | null {
  const argument = fieldTokens(instr)[1]
  return argument === undefined || argument.startsWith('\\') ? null : argument
}

/** O valor de uma chave (`\o "1-3"` → `1-3`), `''` para chave sem valor, `null` se ausente. */
export function fieldSwitch(instr: string, name: string): string | null {
  const tokens = fieldTokens(instr)
  const index = tokens.findIndex((token) => token.toLowerCase() === `\\${name.toLowerCase()}`)
  if (index < 0) return null
  const value = tokens[index + 1]
  return value === undefined || value.startsWith('\\') ? '' : value
}

/**
 * Os níveis de título que o sumário lista: o `\o "1-3"` do Word.
 *
 * Sem a chave, ou com ela vazia, vão os nove, que é o que o Word faz com `\o`
 * sozinho. O sumário sem `\o` nenhum (só `\t`, estilos escolhidos à mão) não é
 * reproduzido aqui, e cai nos três primeiros, que é o sumário padrão.
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

/** As entradas do sumário são links (`\h`)? */
export function tocLinks(instr: string): boolean {
  return fieldSwitch(instr, 'h') !== null
}

/** O sumário omite os números de página (`\n` sem intervalo)? */
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
  // A, B… Z, AA, BB…: é assim que o Word conta depois do Z.
  const letter = String.fromCharCode(65 + ((value - 1) % 26))
  return letter.repeat(Math.floor((value - 1) / 26) + 1)
}

/** O número no formato de `\* ARABIC`, `ROMAN`, `roman`, `ALPHABETIC` ou `alphabetic`. */
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
 * Os números de uma sequência de campos `SEQ`, na ordem do documento.
 *
 * Cada identificador conta à parte (`SEQ Figura` e `SEQ Tabela`), sem diferença
 * entre maiúscula e minúscula, como no Word. `\r n` recomeça em `n`, `\c` repete
 * o último sem avançar.
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
