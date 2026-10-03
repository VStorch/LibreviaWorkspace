import { useEffect } from 'react'
import { Theme, type ResolvedTheme } from '@shared/types.js'
import { usePreferences } from './preferences.js'

const DARK_QUERY = '(prefers-color-scheme: dark)'

/**
 * A escolha explícita decide sozinha; **só** `system` consulta a mídia. Deixar a
 * consulta decidir sempre, contando com o `themeSource` do main, não funciona: o
 * teste de ponta a ponta pede o escuro e o `data-theme` continua `light`.
 */
function resolve(theme: Theme, systemPrefersDark: boolean): ResolvedTheme {
  if (theme === Theme.Light) return 'light'
  if (theme === Theme.Dark) return 'dark'
  return systemPrefersDark ? 'dark' : 'light'
}

/** Um atributo, e não classe: o CSS usa `:root[data-theme='dark']` nos dois casos. */
export function useTheme(): ResolvedTheme {
  const theme = usePreferences((state) => state.preferences.theme)

  useEffect(() => {
    const media = window.matchMedia(DARK_QUERY)

    const apply = (): void => {
      const resolved = resolve(theme, media.matches)
      document.documentElement.setAttribute('data-theme', resolved)
      // `color-scheme` põe barras de rolagem, campos e menus nativos na cor certa.
      document.documentElement.style.colorScheme = resolved
    }

    apply()
    media.addEventListener('change', apply)
    return () => media.removeEventListener('change', apply)
    // Sem a dependência, trocar o tema reaplicaria a escolha antiga.
  }, [theme])

  return resolve(theme, window.matchMedia(DARK_QUERY).matches)
}
