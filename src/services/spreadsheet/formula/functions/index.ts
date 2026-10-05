/** Each function answers in Portuguese and English; XLSX stores the English name (`interop.ts`). */

import { DATE } from './date.js'
import { LOGICAL } from './logical.js'
import { LOOKUP } from './lookup.js'
import { MATH } from './math.js'
import { STATS } from './stats.js'
import { TEXT } from './text.js'
import type { FunctionDefinition } from './kit.js'

export type { FunctionDefinition } from './kit.js'

const ALL: readonly FunctionDefinition[] = [...MATH, ...STATS, ...LOGICAL, ...TEXT, ...LOOKUP, ...DATE]

const BY_NAME = new Map<string, FunctionDefinition>()
for (const definition of ALL) {
  for (const name of definition.names) {
    if (BY_NAME.has(name)) throw new Error(`Função duplicada no catálogo: ${name}`)
    BY_NAME.set(name, definition)
  }
}

/** Lazy functions live in the evaluator, but the XLSX translation looks them up here. */
const LAZY: readonly (readonly string[])[] = [
  ['SE', 'IF'],
  ['SEERRO', 'IFERROR'],
  ['SENÃODISP', 'SEND', 'IFNA'],
]

export function findFunction(name: string): FunctionDefinition | undefined {
  return BY_NAME.get(name.toUpperCase())
}

export function isKnownFunction(name: string): boolean {
  return findFunction(name) !== undefined || lazyNames(name) !== undefined
}

function lazyNames(name: string): readonly string[] | undefined {
  const upper = name.toUpperCase()
  return LAZY.find((group) => group.includes(upper))
}

/** An unknown name comes back as it came, to cross the file without being mangled. */
export function localizedName(name: string, language: 'pt' | 'en'): string {
  const names = findFunction(name)?.names ?? lazyNames(name)
  if (names === undefined) return name
  return (language === 'pt' ? names[0] : names.at(-1))!
}

/** Portuguese names, to complete what the user types. */
export function functionNames(): string[] {
  const names = [...ALL.map((definition) => definition.names[0]!), ...LAZY.map((group) => group[0]!)]
  return [...new Set(names)].sort((a, b) => a.localeCompare(b, 'pt-BR'))
}
