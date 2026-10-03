import type { Language } from './language.js'

export type Message = string | PluralMessage

export interface PluralMessage {
  readonly one: string
  readonly other: string
}

export interface Entry {
  readonly pt: Message
  readonly en: Message
}

/** O formato de um catálogo. Usado com `satisfies`, para as chaves ficarem literais. */
export type Catalog = Readonly<Record<string, Entry>>

/** `count` também escolhe entre singular e plural. */
export type Vars = Readonly<Record<string, string | number>>

export function format(entry: Entry, language: Language, vars?: Vars): string {
  const message = entry[language]
  const count = vars?.['count']

  const text =
    typeof message === 'string'
      ? message
      : // Em português e em inglês só o 1 é singular; o zero é plural.
        count === 1
        ? message.one
        : message.other

  if (vars === undefined) return text
  return interpolate(text, vars)
}

/**
 * Uma varredura só: um valor que contenha `{outra}`, como um nome de arquivo,
 * não é interpolado de novo. Buraco sem valor fica visível, para a falta
 * aparecer na tela.
 */
function interpolate(text: string, vars: Vars): string {
  return text.replace(/\{(\w+)\}/g, (whole, name: string) => {
    const value = vars[name]
    return value === undefined ? whole : String(value)
  })
}
