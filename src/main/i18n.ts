import { APP_NAME } from '@shared/constants.js'
import { translate, type MessageKey, type Vars } from '@shared/i18n/index.js'
import { editorPreferences } from './preferences.js'

/**
 * No main o idioma vem da loja de preferências, e no renderer da cópia que a
 * tela desenha; a tradução é a mesma função de `shared`. `{app}` já vem preenchido.
 */
export function t(key: MessageKey, vars?: Vars): string {
  return translate(editorPreferences().language, key, { app: APP_NAME, ...vars })
}
