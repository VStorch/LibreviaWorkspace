import { useCallback } from 'react'
import { APP_NAME } from '@shared/constants.js'
import { translate, type Language, type MessageKey, type Vars } from '@shared/i18n/index.js'
import { currentPreferences, usePreferences } from './state/preferences.js'

/** Redraws its users when it changes. */
export function useLanguage(): Language {
  return usePreferences((state) => state.preferences.language)
}

/**
 * A hook: switching language must redraw the screen, and React only knows that if the component
 * read the language from the store. `useCallback` bound to the language, so `t` does not change
 * every frame.
 */
export function useT(): (key: MessageKey, vars?: Vars) => string {
  const language = useLanguage()
  return useCallback(
    (key: MessageKey, vars?: Vars) => translate(language, key, { app: APP_NAME, ...vars }),
    [language],
  )
}

/** For extensions and dispatchers, which are recreated when the preference changes. */
export function t(key: MessageKey, vars?: Vars): string {
  return translate(currentPreferences().language, key, { app: APP_NAME, ...vars })
}
