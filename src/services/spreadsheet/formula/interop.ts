/**
 * O XLSX guarda nomes de função em inglês, vírgula entre argumentos e ponto
 * decimal; o aplicativo mostra o dialeto do Excel em português. A tradução é
 * **caractere a caractere**, e não pela árvore: a gravação cirúrgica só reescreve
 * a célula cuja fórmula mudou, e reconstruir normalizaria espaços e maiúsculas.
 */

import { localizedName } from './functions/index.js'

/** Erros escritos dentro da fórmula, como em `=SEERRO(A1;#N/D)`. */
const ERROR_PAIRS: readonly (readonly [string, string])[] = [
  ['#DIV/0!', '#DIV/0!'],
  ['#VALOR!', '#VALUE!'],
  ['#REF!', '#REF!'],
  ['#NOME?', '#NAME?'],
  ['#NÚM!', '#NUM!'],
  ['#N/D', '#N/A'],
]

interface Dialect {
  readonly language: 'pt' | 'en'
  /** Separadores decimais aceitos na origem. */
  readonly decimalsIn: string
  /** Separador decimal a escrever. */
  readonly decimalOut: string
  /** Separador de argumentos: o da origem, depois o do destino. */
  readonly separator: readonly [string, string]
  readonly errors: ReadonlyMap<string, string>
}

const TO_APP: Dialect = {
  language: 'pt',
  // No arquivo a vírgula sempre separa argumentos, nunca é decimal.
  decimalsIn: '.',
  decimalOut: ',',
  separator: [',', ';'],
  errors: new Map(ERROR_PAIRS.map(([portuguese, english]) => [english, portuguese])),
}

const TO_FILE: Dialect = {
  language: 'en',
  // Os dois, porque o ponto é o que sai ao colar de planilha estrangeira.
  decimalsIn: '.,',
  decimalOut: '.',
  separator: [';', ','],
  errors: new Map(ERROR_PAIRS),
}

/** Fórmula lida do XLSX, no dialeto do aplicativo. */
export function fromXlsxFormula(formula: string): string {
  return convert(formula, TO_APP)
}

/** Fórmula do aplicativo, no dialeto do arquivo. */
export function toXlsxFormula(formula: string): string {
  return convert(formula, TO_FILE)
}

function convert(formula: string, dialect: Dialect): string {
  let out = ''
  let at = 0

  while (at < formula.length) {
    const char = formula[at]!

    // Texto entre aspas e nome de planilha entre apóstrofos passam intactos:
    // uma vírgula ali dentro é conteúdo, não separador.
    if (char === '"' || char === "'") {
      const end = closingAt(formula, at, char)
      out += formula.slice(at, end)
      at = end
      continue
    }

    // Matriz literal `{1,2;3,4}`: o aplicativo não a entende, e ela passa intacta.
    if (char === '{') {
      const end = formula.indexOf('}', at)
      const stop = end === -1 ? formula.length : end + 1
      out += formula.slice(at, stop)
      at = stop
      continue
    }

    if (char === '#') {
      const literal = longestMatch(formula, at, dialect.errors)
      if (literal !== undefined) {
        out += dialect.errors.get(literal)!
        at += literal.length
        continue
      }
    }

    if (isDigit(char)) {
      const number = readNumber(formula, at, dialect.decimalsIn)
      out += rewriteDecimal(number, dialect)
      at += number.length
      continue
    }

    if (char === dialect.separator[0]) {
      out += dialect.separator[1]
      at++
      continue
    }

    const word = readWord(formula, at)
    if (word.length > 0) {
      // Só é função com o parêntese colado, como no analisador: uma célula chamada `SOMA` não é.
      out += formula[at + word.length] === '(' ? localizedName(word, dialect.language) : word
      at += word.length
      continue
    }

    out += char
    at++
  }

  return out
}

/** Já depois do fechamento. A aspa duplicada (`""`) não fecha, e o apóstrofo segue a mesma convenção. */
function closingAt(formula: string, start: number, quote: string): number {
  let at = start + 1
  while (at < formula.length) {
    if (formula[at] === quote) {
      if (formula[at + 1] === quote) {
        at += 2
        continue
      }
      return at + 1
    }
    at++
  }
  // Aspa sem fechar: copia o resto em vez de inventar um fechamento.
  return formula.length
}

function longestMatch(formula: string, at: number, table: ReadonlyMap<string, string>): string | undefined {
  let best: string | undefined
  for (const candidate of table.keys()) {
    if (!formula.startsWith(candidate, at)) continue
    if (best === undefined || candidate.length > best.length) best = candidate
  }
  return best
}

function isDigit(char: string | undefined): boolean {
  return char !== undefined && char >= '0' && char <= '9'
}

/** Número inteiro ou decimal, com expoente. Espelha o do analisador. */
function readNumber(formula: string, start: number, decimals: string): string {
  let at = start
  let seenSeparator = false

  while (at < formula.length) {
    const char = formula[at]!
    if (isDigit(char)) {
      at++
      continue
    }
    if (decimals.includes(char) && !seenSeparator && isDigit(formula[at + 1])) {
      seenSeparator = true
      at += 2
      continue
    }
    if ((char === 'e' || char === 'E') && at > start) {
      if (isDigit(formula[at + 1])) {
        at += 2
        continue
      }
      if ((formula[at + 1] === '+' || formula[at + 1] === '-') && isDigit(formula[at + 2])) {
        at += 3
        continue
      }
    }
    break
  }

  return formula.slice(start, at)
}

function rewriteDecimal(number: string, dialect: Dialect): string {
  for (const separator of dialect.decimalsIn) {
    const dot = number.indexOf(separator)
    if (dot !== -1) {
      return number.slice(0, dot) + dialect.decimalOut + number.slice(dot + 1)
    }
  }
  return number
}

/** O mesmo conjunto do analisador: `$`, `!` e `.` fazem parte de ``, `Plan1!A1` e `CONT.NÚM`. */
function readWord(formula: string, start: number): string {
  let at = start
  while (at < formula.length && /[\p{L}\p{N}_$.!]/u.test(formula[at]!)) at++
  return formula.slice(start, at)
}
