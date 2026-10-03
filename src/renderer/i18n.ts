import { useCallback } from 'react'
import { APP_NAME } from '@shared/constants.js'
import { translate, type Language, type MessageKey, type Vars } from '@shared/i18n/index.js'
import { currentPreferences, usePreferences } from './state/preferences.js'

/** Redesenha quem o usa quando ele muda. */
export function useLanguage(): Language {
  return usePreferences((state) => state.preferences.language)
}

/**
 * Um hook: trocar de idioma tem de redesenhar a tela, e o React só sabe disso se
 * o componente leu o idioma da loja. `useCallback` preso ao idioma, para `t` não
 * mudar a cada quadro.
 */
export function useT(): (key: MessageKey, vars?: Vars) => string {
  const language = useLanguage()
  return useCallback(
    (key: MessageKey, vars?: Vars) => translate(language, key, { app: APP_NAME, ...vars }),
    [language],
  )
}

/** Para extensões e despachantes, que são recriados quando a preferência muda. */
export function t(key: MessageKey, vars?: Vars): string {
  return translate(currentPreferences().language, key, { app: APP_NAME, ...vars })
}
