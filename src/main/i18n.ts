import { APP_NAME } from '@shared/constants.js'
import { translate, type MessageKey, type Vars } from '@shared/i18n/index.js'
import { editorPreferences } from './preferences.js'

/**
 * In main the language comes from the preferences store, in the renderer from the copy the screen
 * draws; the translation is the same function from `shared`. `{app}` is already filled in.
 */
export function t(key: MessageKey, vars?: Vars): string {
  return translate(editorPreferences().language, key, { app: APP_NAME, ...vars })
}
