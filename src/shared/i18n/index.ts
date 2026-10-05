import { MESSAGES, type MessageKey } from './catalog/index.js'
import { format, type Vars } from './message.js'
import type { Language } from './language.js'

export { Language, LANGUAGES, LANGUAGE_NAMES, languageFromLocale } from './language.js'
export type { Catalog, Entry, Message, PluralMessage, Vars } from './message.js'
export { MESSAGES } from './catalog/index.js'
export type { MessageKey } from './catalog/index.js'

/** The language is an argument, not state, so main and renderer share the same function. */
export function translate(language: Language, key: MessageKey, vars?: Vars): string {
  return format(MESSAGES[key], language, vars)
}
