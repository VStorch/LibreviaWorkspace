/**
 * Valores de fórmula e as conversões entre eles.
 *
 * O tipo mais importante aqui é o **vazio**, representado por `null`. Célula
 * vazia não é zero nem texto vazio: em `MÉDIA(A1:A3)` com duas células
 * preenchidas o divisor é dois, e tratá-la como zero daria uma média errada
 * sem nenhum aviso. Ela vira zero só quando entra numa conta, que é o que o
 * Excel faz.
 */

import { FormulaError, isFormulaError } from './errors.js'

/** Um valor de fórmula. `null` é célula vazia. */
export type Scalar = number | string | boolean | FormulaError | null

/**
 * Texto → número, para uso em contas.
 *
 * Aceita o formato brasileiro, porque é o que o usuário digitou na célula ao
 * lado. Texto que não é número vira `#VALOR!` — nunca zero: somar uma coluna
 * com "n/a" no meio e receber um total silenciosamente menor é pior que receber
 * um erro.
 */
export function toNumber(value: Scalar): number | FormulaError {
  if (isFormulaError(value)) return value
  if (value === null) return 0
  if (typeof value === 'number') return value
  if (typeof value === 'boolean') return value ? 1 : 0

  const text = value.trim()
  if (text.length === 0) return 0

  const normalized = text.includes(',') ? text.replaceAll('.', '').replace(',', '.') : text
  const number = Number(normalized)
  return Number.isFinite(number) ? number : FormulaError.Value
}

export function toText(value: Scalar): string | FormulaError {
  if (isFormulaError(value)) return value
  if (value === null) return ''
  if (typeof value === 'boolean') return value ? 'VERDADEIRO' : 'FALSO'
  if (typeof value === 'number') return formatNumber(value)
  return value
}

/**
 * Número → texto dentro da fórmula.
 *
 * Usa ponto decimal, e não vírgula: este texto entra em concatenação e em
 * comparação, onde precisa ser o mesmo em qualquer máquina. A vírgula aparece
 * na exibição da célula, que é outra camada.
 */
function formatNumber(value: number): string {
  return Number.isInteger(value) ? String(value) : String(Number(value.toPrecision(15)))
}

export function toBoolean(value: Scalar): boolean | FormulaError {
  if (isFormulaError(value)) return value
  if (value === null) return false
  if (typeof value === 'boolean') return value
  if (typeof value === 'number') return value !== 0

  const upper = value.trim().toUpperCase()
  if (upper === 'VERDADEIRO' || upper === 'TRUE') return true
  if (upper === 'FALSO' || upper === 'FALSE') return false
  return FormulaError.Value
}

/**
 * Ordem entre dois valores, para os operadores de comparação.
 *
 * Duas regras do Excel que surpreendem quem não as conhece:
 *
 * - **texto compara sem diferenciar maiúsculas**: `="a"="A"` é verdadeiro;
 * - **tipos diferentes nunca são iguais**: `=1="1"` é falso. Comparar por
 *   conversão faria `=A1=0` ser verdadeiro para uma célula de texto, e a conta
 *   toda passaria por um caminho errado.
 *
 * Entre tipos, a ordem é número < texto < booleano.
 */
export function compare(left: Scalar, right: Scalar): number {
  const a = left ?? blankLike(right)
  const b = right ?? blankLike(left)

  if (typeof a === 'number' && typeof b === 'number') return a === b ? 0 : a < b ? -1 : 1
  if (typeof a === 'string' && typeof b === 'string') {
    const x = a.toUpperCase()
    const y = b.toUpperCase()
    return x === y ? 0 : x < y ? -1 : 1
  }
  if (typeof a === 'boolean' && typeof b === 'boolean') return a === b ? 0 : a ? 1 : -1

  return rank(a) - rank(b)
}

/**
 * O vazio assume o tipo do outro lado.
 *
 * É o que faz `=A1=""` ser verdadeiro para célula vazia e `=A1=0` também — as
 * duas coisas que o usuário espera, e que o Excel entrega.
 */
function blankLike(other: Scalar): Scalar {
  if (typeof other === 'string') return ''
  if (typeof other === 'boolean') return false
  return 0
}

function rank(value: Scalar): number {
  if (typeof value === 'number') return 0
  if (typeof value === 'string') return 1
  return 2
}

/**
 * Critério de `SOMASE` e `CONT.SE`.
 *
 * Aceita `">10"`, `"<>0"`, `"São Paulo"` e curingas `*` e `?`. Sem os
 * comparadores, o critério só serviria para igualdade exata e as duas funções
 * perderiam a maior parte do uso real.
 */
export function matchesCriteria(value: Scalar, criteria: Scalar): boolean {
  if (typeof criteria !== 'string') return compare(value, criteria) === 0

  const match = /^(<>|>=|<=|=|<|>)(.*)$/.exec(criteria.trim())
  if (match === null) return matchesPattern(value, criteria)

  const operator = match[1]!
  const rest = match[2]!.trim()
  const target: Scalar = rest.length === 0 ? null : (numberOrText(rest) as Scalar)

  // Igualdade com curinga é o único caso em que o critério não é ordem.
  if ((operator === '=' || operator === '<>') && typeof target === 'string' && hasWildcard(target)) {
    const hit = matchesPattern(value, target)
    return operator === '=' ? hit : !hit
  }

  const order = compare(value, target)
  switch (operator) {
    case '=':
      return order === 0
    case '<>':
      return order !== 0
    case '>':
      return order > 0
    case '>=':
      return order >= 0
    case '<':
      return order < 0
    default:
      return order <= 0
  }
}

function numberOrText(text: string): number | string {
  const number = toNumber(text)
  return typeof number === 'number' && text.trim().length > 0 ? number : text
}

function hasWildcard(text: string): boolean {
  return /(^|[^~])[*?]/.test(text)
}

/** Curingas do Excel: `*` é qualquer coisa, `?` é um caractere, `~` escapa. */
function matchesPattern(value: Scalar, pattern: string): boolean {
  if (!hasWildcard(pattern)) return compare(value, numberOrText(pattern) as Scalar) === 0

  const text = toText(value)
  if (isFormulaError(text)) return false

  let regex = ''
  for (let at = 0; at < pattern.length; at++) {
    const char = pattern[at]!
    if (char === '~' && (pattern[at + 1] === '*' || pattern[at + 1] === '?')) {
      regex += escapeRegex(pattern[++at]!)
    } else if (char === '*') regex += '.*'
    else if (char === '?') regex += '.'
    else regex += escapeRegex(char)
  }

  return new RegExp(`^${regex}$`, 'iu').test(text)
}

function escapeRegex(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}
