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

/** Used with `satisfies`, so the keys stay literal. */
export type Catalog = Readonly<Record<string, Entry>>

/** `count` also picks singular or plural. */
export type Vars = Readonly<Record<string, string | number>>

export function format(entry: Entry, language: Language, vars?: Vars): string {
  const message = entry[language]
  const count = vars?.['count']

  const text =
    typeof message === 'string'
      ? message
      : // In Portuguese and English only 1 is singular; zero is plural.
        count === 1
        ? message.one
        : message.other

  if (vars === undefined) return text
  return interpolate(text, vars)
}

/**
 * A single pass: a value that contains `{other}`, such as a file name, is not interpolated again. A
 * placeholder without a value stays visible, so the gap shows on screen.
 */
function interpolate(text: string, vars: Vars): string {
  return text.replace(/\{(\w+)\}/g, (whole, name: string) => {
    const value = vars[name]
    return value === undefined ? whole : String(value)
  })
}
