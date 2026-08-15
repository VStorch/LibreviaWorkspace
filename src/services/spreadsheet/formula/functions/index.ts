/**
 * Catálogo de funções.
 *
 * Cada função responde pelo nome em **português e em inglês**. O usuário digita
 * `SOMA`, quem colou de uma planilha estrangeira digita `SUM`, e as duas
 * funcionam. Traduzir na hora de gravar XLSX é problema da Fase 7 — o formato
 * guarda sempre o nome canônico em inglês, independente do idioma da interface.
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

/** As preguiçosas moram no avaliador, mas precisam constar do catálogo. */
const LAZY_NAMES: readonly string[] = ['SE', 'IF', 'SEERRO', 'IFERROR', 'SENÃODISP', 'SEND', 'IFNA']

export function findFunction(name: string): FunctionDefinition | undefined {
  return BY_NAME.get(name.toUpperCase())
}

export function isKnownFunction(name: string): boolean {
  const upper = name.toUpperCase()
  return BY_NAME.has(upper) || LAZY_NAMES.includes(upper)
}

/** Nomes em português, para completar o que o usuário digita. */
export function functionNames(): string[] {
  return [...new Set([...ALL.map((definition) => definition.names[0]!), 'SE', 'SEERRO'])].sort((a, b) =>
    a.localeCompare(b, 'pt-BR'),
  )
}
