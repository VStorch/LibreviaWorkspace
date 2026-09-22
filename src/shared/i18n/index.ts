import { MESSAGES, type MessageKey } from './catalog/index.js'
import { format, type Vars } from './message.js'
import type { Language } from './language.js'

export { Language, LANGUAGES, LANGUAGE_NAMES, languageFromLocale } from './language.js'
export type { Catalog, Entry, Message, PluralMessage, Vars } from './message.js'
export { MESSAGES } from './catalog/index.js'
export type { MessageKey } from './catalog/index.js'

/**
 * A frase, no idioma pedido.
 *
 * Pura, sem estado e sem idioma embutido — o idioma é **argumento**. É o que
 * permite ao main e ao renderer compartilharem esta função: os dois sabem o
 * idioma atual por caminhos diferentes (a loja de preferências num, a cópia
 * desenhável no outro), e nenhum dos dois precisa que o outro saiba.
 *
 * A chave é verificada pelo compilador. Não existe caminho de "chave que faltou"
 * em tempo de execução, e por isso não existe o retorno feio que todo sistema de
 * tradução acaba mostrando ao usuário.
 */
export function translate(language: Language, key: MessageKey, vars?: Vars): string {
  return format(MESSAGES[key], language, vars)
}
