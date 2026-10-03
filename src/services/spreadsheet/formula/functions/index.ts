/** Cada função responde em português e em inglês; o XLSX guarda o nome em inglês (`interop.ts`). */

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

/** As preguiçosas moram no avaliador, mas a tradução para XLSX as consulta aqui. */
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

/** Nome desconhecido volta como veio, para atravessar o arquivo sem ser desfigurado. */
export function localizedName(name: string, language: 'pt' | 'en'): string {
  const names = findFunction(name)?.names ?? lazyNames(name)
  if (names === undefined) return name
  return (language === 'pt' ? names[0] : names.at(-1))!
}

/** Nomes em português, para completar o que o usuário digita. */
export function functionNames(): string[] {
  const names = [...ALL.map((definition) => definition.names[0]!), ...LAZY.map((group) => group[0]!)]
  return [...new Set(names)].sort((a, b) => a.localeCompare(b, 'pt-BR'))
}
