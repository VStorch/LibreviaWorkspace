/**
 * Catálogo de funções.
 *
 * Cada função responde pelo nome em **português e em inglês**. O usuário digita
 * `SOMA`, quem colou de uma planilha estrangeira digita `SUM`, e as duas
 * funcionam. O XLSX guarda sempre o nome canônico em inglês, independente do
 * idioma de quem escreveu — quem traduz para lá e de lá é `interop.ts`, usando
 * a ordem desta lista.
 */

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
    // Nome repetido é engano de quem escreveu o catálogo, e silenciar isso
    // faria uma das duas funções sumir sem aviso.
    if (BY_NAME.has(name)) throw new Error(`Função duplicada no catálogo: ${name}`)
    BY_NAME.set(name, definition)
  }
}

/**
 * As preguiçosas moram no avaliador, mas precisam constar do catálogo.
 *
 * Guardadas com a mesma forma das outras — português primeiro, inglês por
 * último — porque a tradução para XLSX consulta as duas listas.
 */
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

/**
 * O mesmo nome de função na outra língua.
 *
 * Nome desconhecido volta como veio: uma função que o motor não calcula ainda
 * precisa atravessar a ida e volta pelo arquivo sem ser desfigurada.
 */
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
