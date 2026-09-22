import { APP_NAME } from '@shared/constants.js'
import { translate, type MessageKey, type Vars } from '@shared/i18n/index.js'
import { editorPreferences } from './preferences.js'

/**
 * A frase, no idioma que está guardado.
 *
 * O main tem um `t` próprio — e não o do renderer — porque o idioma chega por
 * caminhos diferentes nos dois lados. Aqui ele vem da loja de preferências, que
 * é a dona; lá vem da cópia que a tela desenha. A tradução em si é a mesma
 * função pura de `shared`, então as duas pontas não podem divergir no texto.
 *
 * `{app}` já vem preenchido: o nome do aplicativo aparece em meia dúzia de
 * rótulos do menu, e passá-lo em cada chamada seria seis oportunidades de
 * escrevê-lo errado.
 */
export function t(key: MessageKey, vars?: Vars): string {
  return translate(editorPreferences().language, key, { app: APP_NAME, ...vars })
}
