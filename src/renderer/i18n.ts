import { useCallback } from 'react'
import { APP_NAME } from '@shared/constants.js'
import { translate, type Language, type MessageKey, type Vars } from '@shared/i18n/index.js'
import { currentPreferences, usePreferences } from './state/preferences.js'

/** O idioma que a tela está desenhando. Redesenha quem o usa quando ele muda. */
export function useLanguage(): Language {
  return usePreferences((state) => state.preferences.language)
}

/**
 * O `t` da interface.
 *
 * Um hook, e não uma função solta, porque trocar de idioma tem de redesenhar a
 * tela — e a única forma de o React saber disso é o componente ter lido o
 * idioma da loja. Uma função global leria o valor certo e ninguém pediria o
 * redesenho: a barra de ferramentas ficaria na língua antiga até o próximo
 * clique, que é o defeito clássico deste tipo de sistema.
 *
 * `useCallback` preso ao idioma: sem ele, todo componente memoizado que receba
 * `t` como dependência se redesenharia a cada quadro.
 */
export function useT(): (key: MessageKey, vars?: Vars) => string {
  const language = useLanguage()
  return useCallback(
    (key: MessageKey, vars?: Vars) => translate(language, key, { app: APP_NAME, ...vars }),
    [language],
  )
}

/**
 * A frase, para quem não é componente.
 *
 * As extensões do editor e os despachantes de comando não têm como chamar um
 * hook, e leem o idioma da mesma forma que já leem as outras preferências. Elas
 * são recriadas quando a preferência muda, então a frase não envelhece — o
 * mesmo caminho que `currentPreferences` já servia para a ortografia.
 */
export function t(key: MessageKey, vars?: Vars): string {
  return translate(currentPreferences().language, key, { app: APP_NAME, ...vars })
}
